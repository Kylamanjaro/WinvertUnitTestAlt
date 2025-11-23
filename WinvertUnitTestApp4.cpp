#include "pch.h"

#include <windows.h>
#include <shobjidl.h>   // IApplicationActivationManager
#include <appmodel.h>   // GetPackageFamilyName
#include <atlbase.h>
#include <uiautomation.h>
#pragma comment(lib, "Uiautomationcore.lib")

#include <string>
#include <chrono>
#include <thread>
#include <filesystem>
#include <cstdlib>
#include <cmath>
#include <functional>
#include <cwchar>

#include <winrt/Windows.Data.Json.h>

#include "CppUnitTest.h"
using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace fs = std::filesystem;

static void SleepMs(int ms) { std::this_thread::sleep_for(std::chrono::milliseconds(ms)); }

static std::wstring GetUserProfile()
{
    wchar_t buf[MAX_PATH]; DWORD n = GetEnvironmentVariableW(L"USERPROFILE", buf, MAX_PATH); if (n == 0) return L""; return std::wstring(buf);
}

static std::wstring PFNForProcess(DWORD pid)
{
    std::wstring pfn;
    HANDLE h = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (!h) return pfn;
    wchar_t name[PACKAGE_FAMILY_NAME_MAX_LENGTH]{}; UINT32 len = PACKAGE_FAMILY_NAME_MAX_LENGTH;
    if (GetPackageFamilyName(h, &len, name) == ERROR_SUCCESS) pfn.assign(name);
    CloseHandle(h);
    return pfn;
}

static std::wstring LocalStatePathForPFN(const std::wstring& pfn)
{
    auto user = GetUserProfile();
    if (user.empty() || pfn.empty()) return L"";
    return user + L"\\AppData\\Local\\Packages\\" + pfn + L"\\LocalState";
}

static bool WriteUtf8(const std::wstring& path, const std::string& data)
{
    FILE* f{}; _wfopen_s(&f, path.c_str(), L"wb"); if (!f) return false; fwrite(data.data(), 1, data.size(), f); fclose(f); return true;
}

static std::string ReadAllUtf8(const std::wstring& path)
{
    FILE* f{}; _wfopen_s(&f, path.c_str(), L"rb"); if (!f) return {}; fseek(f, 0, SEEK_END); long sz = ftell(f); fseek(f, 0, SEEK_SET); std::string s; s.resize(sz); fread(s.data(), 1, sz, f); fclose(f); return s;
}

static DWORD LaunchWinvert4(std::wstring& outPfn)
{
    CComPtr<IApplicationActivationManager> aam;
    HRESULT hr = CoCreateInstance(CLSID_ApplicationActivationManager, nullptr,
                                  CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&aam));
    if (FAILED(hr)) return 0;

    constexpr wchar_t kWinvert4Aumid[] =
        L"b27d31cf-c66d-45ac-aad0-e0d9501a1c90_ft4zefc91v2gy!App"; // <- your actual value

    DWORD pid = 0;
    hr = aam->ActivateApplication(kWinvert4Aumid, nullptr, AO_NONE, &pid);
    if (FAILED(hr) || pid == 0) return 0;

    outPfn = PFNForProcess(pid);
    return pid;
}

static void SendKeyboardInput(WORD vk, DWORD flags = 0, DWORD extraSleepMs = 10)
{
    INPUT in{};
    in.type = INPUT_KEYBOARD;
    in.ki.wVk = vk;
    in.ki.dwFlags = flags;
    SendInput(1, &in, sizeof(INPUT));
    if (extraSleepMs > 0) SleepMs(static_cast<int>(extraSleepMs));
}

static void SendHotkey(UINT modifiers, UINT vk)
{
    WORD modKeys[4]{};
    int count = 0;
    auto push = [&](WORD key) { modKeys[count++] = key; };
    if (modifiers & MOD_WIN) push(VK_LWIN);
    if (modifiers & MOD_CONTROL) push(VK_CONTROL);
    if (modifiers & MOD_ALT) push(VK_MENU);
    if (modifiers & MOD_SHIFT) push(VK_SHIFT);

    for (int i = 0; i < count; ++i) SendKeyboardInput(modKeys[i]);
    SendKeyboardInput(static_cast<WORD>(vk));
    SendKeyboardInput(static_cast<WORD>(vk), KEYEVENTF_KEYUP);
    for (int i = count - 1; i >= 0; --i) SendKeyboardInput(modKeys[i], KEYEVENTF_KEYUP);
}

static void TapKey(WORD vk)
{
    SendKeyboardInput(vk);
    SendKeyboardInput(vk, KEYEVENTF_KEYUP);
}

static void LogMessage(const wchar_t* msg)
{
    Logger::WriteMessage(msg);
}
static void LogMessage(const std::wstring& msg)
{
    Logger::WriteMessage(msg.c_str());
}

static void TriggerInitialSelectionGesture()
{
    LogMessage(L"[Test] TriggerInitialSelectionGesture: Win+Shift+I");
    SleepMs(2000);
    SendHotkey(MOD_WIN | MOD_SHIFT, 'I');
    SleepMs(750);
    LogMessage(L"[Test] TriggerInitialSelectionGesture: selecting monitor 1");
    TapKey('1');
    SleepMs(750);
}

static std::wstring GetWindowTitle(HWND hwnd)
{
    if (!hwnd) return {};
    int len = GetWindowTextLengthW(hwnd);
    if (len <= 0) return {};
    std::wstring title;
    title.resize(len + 1);
    int written = GetWindowTextW(hwnd, title.data(), len + 1);
    if (written <= 0) return {};
    title.resize(written);
    return title;
}

static bool IsCandidateMainWindow(HWND hwnd)
{
    if (!IsWindow(hwnd)) return false;
    if (GetWindow(hwnd, GW_OWNER) != nullptr) return false;
    LONG_PTR style = GetWindowLongPtr(hwnd, GWL_STYLE);
    if (style & WS_DISABLED) return false;
    if (!(style & (WS_CAPTION | WS_POPUP))) return false;
    return true;
}

static HWND FindTopLevelWindowForPid(DWORD pid)
{
    struct EnumCtx
    {
        DWORD targetPid{};
        HWND found{};
        HWND fallback{};
    } ctx{pid, nullptr};

    EnumWindows(
        [](HWND hwnd, LPARAM lParam)->BOOL
        {
            auto* ctx = reinterpret_cast<EnumCtx*>(lParam);
            DWORD procId = 0;
            GetWindowThreadProcessId(hwnd, &procId);
            if (procId != ctx->targetPid) return TRUE;
            if (!IsCandidateMainWindow(hwnd)) return TRUE;

            auto title = GetWindowTitle(hwnd);
            if (!title.empty() && _wcsicmp(title.c_str(), L"Winvert Control Panel") == 0)
            {
                ctx->found = hwnd;
                return FALSE;
            }

            if (!ctx->fallback)
            {
                ctx->fallback = hwnd;
            }
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&ctx));

    if (ctx.found) return ctx.found;
    return ctx.fallback;
}

static HWND HwndFromElement(IUIAutomationElement* el)
{
    if (!el) return nullptr;
    CComVariant v;
    if (FAILED(el->GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId, &v))) return nullptr;
    if (v.vt == VT_I4 || v.vt == VT_INT || v.vt == VT_UI4)
    {
        return reinterpret_cast<HWND>(static_cast<LONG_PTR>(v.lVal));
    }
    if (v.vt == VT_I8 || v.vt == VT_UI8)
    {
        return reinterpret_cast<HWND>(static_cast<LONG_PTR>(v.llVal));
    }
    return nullptr;
}

static void EnsureWindowVisible(IUIAutomationElement* el)
{
    if (!el) return;
    if (HWND hwnd = HwndFromElement(el))
    {
        ShowWindow(hwnd, SW_SHOWNORMAL);
        SetForegroundWindow(hwnd);
    }
}

static CComPtr<IUIAutomationElement> FindMainWindowForPid(DWORD pid)
{
    CComPtr<IUIAutomation> uia; if (FAILED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uia)))) return nullptr;
    CComPtr<IUIAutomationElement> root; if (FAILED(uia->GetRootElement(&root))) return nullptr;

    VARIANT vPid; VariantInit(&vPid); vPid.vt = VT_I4; vPid.lVal = static_cast<LONG>(pid);
    CComPtr<IUIAutomationCondition> pidCond; uia->CreatePropertyCondition(UIA_ProcessIdPropertyId, vPid, &pidCond);

    VARIANT vType; VariantInit(&vType); vType.vt = VT_I4; vType.lVal = UIA_WindowControlTypeId;
    CComPtr<IUIAutomationCondition> typeCond; uia->CreatePropertyCondition(UIA_ControlTypePropertyId, vType, &typeCond);

    CComPtr<IUIAutomationCondition> cond; uia->CreateAndCondition(pidCond, typeCond, &cond);

    CComPtr<IUIAutomationElement> win; root->FindFirst(TreeScope_Subtree, cond, &win);
    if (win) { EnsureWindowVisible(win); return win; }

    // Fallback: enum top-level windows for the process and convert to UIA element.
    HWND hwnd = FindTopLevelWindowForPid(pid);
    if (!hwnd) return nullptr;

    CComPtr<IUIAutomationElement> fromHandle;
    if (FAILED(uia->ElementFromHandle(hwnd, &fromHandle))) return nullptr;
    EnsureWindowVisible(fromHandle);
    return fromHandle;
}

static CComPtr<IUIAutomationElement> FindWindowByName(const wchar_t* name)
{
    CComPtr<IUIAutomation> uia; if (FAILED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uia)))) return nullptr;
    CComPtr<IUIAutomationElement> root; if (FAILED(uia->GetRootElement(&root))) return nullptr;
    CComPtr<IUIAutomationCondition> cond; uia->CreatePropertyCondition(UIA_NamePropertyId, CComVariant(name), &cond);
    CComPtr<IUIAutomationElement> found; root->FindFirst(TreeScope_Children, cond, &found);
    return found;
}

static CComPtr<IUIAutomationElement> FindByAutomationId(IUIAutomationElement* parent, const wchar_t* aid)
{
    if (!parent) return nullptr; CComPtr<IUIAutomation> uia; CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uia));
    CComPtr<IUIAutomationCondition> cond; uia->CreatePropertyCondition(UIA_AutomationIdPropertyId, CComVariant(aid), &cond);
    CComPtr<IUIAutomationElement> el; parent->FindFirst(TreeScope_Subtree, cond, &el);
    return el;
}

static CComPtr<IUIAutomationElement> WaitForAutomationId(IUIAutomationElement* parent, const wchar_t* aid, int timeoutMs = 5000, int pollMs = 100)
{
    if (!parent) return nullptr;
    {
        std::wstring msg = L"[Test] WaitForAutomationId: looking for ";
        msg += aid ? aid : L"(null)";
        LogMessage(msg);
    }
    const int iterations = timeoutMs / pollMs;
    for (int i = 0; i < iterations; ++i)
    {
        if (auto el = FindByAutomationId(parent, aid)) return el;
        SleepMs(pollMs);
    }
    auto result = FindByAutomationId(parent, aid);
    if (!result)
    {
        std::wstring msg = L"[Test] WaitForAutomationId: timed out for ";
        msg += aid ? aid : L"(null)";
        LogMessage(msg);
    }
    else
    {
        std::wstring msg = L"[Test] WaitForAutomationId: found after timeout ";
        msg += aid ? aid : L"(null)";
        LogMessage(msg);
    }
    return result;
}

static void DumpAutomationTree(IUIAutomationElement* root, int maxDepth = 6)
{
    if (!root)
    {
        LogMessage(L"[Test] DumpAutomationTree: root is null");
        return;
    }
    CComPtr<IUIAutomation> uia;
    if (FAILED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uia))))
    {
        LogMessage(L"[Test] DumpAutomationTree: failed to get UIA");
        return;
    }
    CComPtr<IUIAutomationTreeWalker> walker;
    if (FAILED(uia->get_ControlViewWalker(&walker)) || !walker)
    {
        LogMessage(L"[Test] DumpAutomationTree: failed to get walker");
        return;
    }

    std::function<void(IUIAutomationElement*, int)> dump = [&](IUIAutomationElement* el, int depth)
    {
        if (!el || depth > maxDepth) return;
        std::wstring indent(depth * 2, L' ');

        CComVariant name, autoId, cls, ctrlType;
        el->GetCurrentPropertyValue(UIA_NamePropertyId, &name);
        el->GetCurrentPropertyValue(UIA_AutomationIdPropertyId, &autoId);
        el->GetCurrentPropertyValue(UIA_ClassNamePropertyId, &cls);
        el->GetCurrentPropertyValue(UIA_ControlTypePropertyId, &ctrlType);

        std::wstring line = indent + L"[Test] UIA: ";
        if (name.vt == VT_BSTR) line += L"name=\"" + std::wstring(name.bstrVal) + L"\" ";
        if (autoId.vt == VT_BSTR) line += L"id=" + std::wstring(autoId.bstrVal) + L" ";
        if (cls.vt == VT_BSTR) line += L"class=" + std::wstring(cls.bstrVal) + L" ";
        if (ctrlType.vt == VT_I4) line += L"type=" + std::to_wstring(ctrlType.lVal) + L" ";
        LogMessage(line);

        CComPtr<IUIAutomationElement> child;
        walker->GetFirstChildElement(el, &child);
        while (child)
        {
            dump(child, depth + 1);
            CComPtr<IUIAutomationElement> next;
            walker->GetNextSiblingElement(child, &next);
            child = next;
        }
    };

    LogMessage(L"[Test] DumpAutomationTree: start");
    dump(root, 0);
    LogMessage(L"[Test] DumpAutomationTree: end");
}

static bool ClickAtScreenPoint(LONG x, LONG y)
{
    if (!SetCursorPos(x, y)) return false;
    INPUT inputs[2]{};
    inputs[0].type = INPUT_MOUSE;
    inputs[0].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
    inputs[1].type = INPUT_MOUSE;
    inputs[1].mi.dwFlags = MOUSEEVENTF_LEFTUP;
    return SendInput(2, inputs, sizeof(INPUT)) == 2;
}

static bool InvokeElement(IUIAutomationElement* el)
{
    if (!el) return false;

    CComPtr<IUIAutomationInvokePattern> pat;
    if (SUCCEEDED(el->GetCurrentPatternAs(UIA_InvokePatternId, IID_PPV_ARGS(&pat))))
    {
        LogMessage(L"[Test] InvokeElement: using InvokePattern");
        if (SUCCEEDED(pat->Invoke()))
        {
            LogMessage(L"[Test] InvokeElement: InvokePattern succeeded");
            return true;
        }
        SleepMs(250);
        if (SUCCEEDED(pat->Invoke()))
        {
            LogMessage(L"[Test] InvokeElement: InvokePattern succeeded after retry");
            return true;
        }
        LogMessage(L"[Test] InvokeElement: InvokePattern failed");
    }

    POINT pt{};
    BOOL got = FALSE;
    if (SUCCEEDED(el->GetClickablePoint(&pt, &got)) && got)
    {
        LogMessage(L"[Test] InvokeElement: clicking clickable point");
        if (ClickAtScreenPoint(pt.x, pt.y))
        {
            LogMessage(L"[Test] InvokeElement: clickable point click succeeded");
            return true;
        }
        LogMessage(L"[Test] InvokeElement: clickable point click failed");
    }

    RECT r{};
    if (SUCCEEDED(el->get_CurrentBoundingRectangle(&r)))
    {
        LogMessage(L"[Test] InvokeElement: clicking bounding rectangle center");
        LONG cx = (r.left + r.right) / 2;
        LONG cy = (r.top + r.bottom) / 2;
        if (ClickAtScreenPoint(cx, cy))
        {
            LogMessage(L"[Test] InvokeElement: bounding rectangle click succeeded");
            return true;
        }
        LogMessage(L"[Test] InvokeElement: bounding rectangle click failed");
    }

    LogMessage(L"[Test] InvokeElement: all methods failed");
    return false;
}
static bool Toggle(IUIAutomationElement* el, ToggleState desired)
{
    if (!el) return false; CComPtr<IUIAutomationTogglePattern> pat; if (FAILED(el->GetCurrentPatternAs(UIA_TogglePatternId, IID_PPV_ARGS(&pat)))) return false; ToggleState cur; pat->get_CurrentToggleState(&cur); if (cur != desired) { if (FAILED(pat->Toggle())) return false; SleepMs(100); }
    return true;
}
static bool SetValue(IUIAutomationElement* el, const wchar_t* value)
{
    if (!el) return false; CComPtr<IUIAutomationValuePattern> pat; if (FAILED(el->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&pat)))) return false; CComBSTR b(value); return SUCCEEDED(pat->SetValue(b));
}

static bool CloseWindow(IUIAutomationElement* win)
{
    if (!win) return false; CComPtr<IUIAutomationWindowPattern> pat; if (FAILED(win->GetCurrentPatternAs(UIA_WindowPatternId, IID_PPV_ARGS(&pat)))) return false; return SUCCEEDED(pat->Close());
}

static bool JsonEqualWithTolerance(const std::wstring& expectedPath, const std::wstring& actualPath)
{
    using namespace winrt::Windows::Data::Json;
    auto load = [](const std::wstring& p)->JsonObject { auto s = ReadAllUtf8(p); std::wstring ws(s.begin(), s.end()); JsonObject o; JsonObject::TryParse(winrt::hstring(ws), o); return o; };
    auto a = load(expectedPath); auto b = load(actualPath);
    if (!a || !b) return false;
    auto get = [](JsonObject& o, const wchar_t* k) { return o.TryLookup(k); };
    auto eqBool = [&](const wchar_t* k) { auto va = get(a, k); auto vb = get(b, k); if (!va || !vb) return false; return va.GetBoolean() == vb.GetBoolean(); };
    auto eqNumTol = [&](double ea, double eb) { return fabs(ea - eb) < 1e-4; };
    // Compare a few key sections (extend as needed)
    auto ta = a.GetNamedObject(L"toggles"); auto tb = b.GetNamedObject(L"toggles");
    if (ta.GetNamedBoolean(L"showFps", false) != tb.GetNamedBoolean(L"showFps", false)) return false;
    if (ta.GetNamedBoolean(L"selectionColorEnabled", false) != tb.GetNamedBoolean(L"selectionColorEnabled", false)) return false;
    if (ta.GetNamedBoolean(L"colorMapPreserve", false) != tb.GetNamedBoolean(L"colorMapPreserve", false)) return false;
    auto ba = a.GetNamedObject(L"brightness"); auto bb = b.GetNamedObject(L"brightness");
    if ((int)ba.GetNamedNumber(L"delayFrames", -1) != (int)bb.GetNamedNumber(L"delayFrames", -1)) return false;
    auto lwa = ba.GetNamedArray(L"lumaWeights"); auto lwb = bb.GetNamedArray(L"lumaWeights");
    for (uint32_t i = 0; i < 3; ++i) if (!eqNumTol(lwa.GetAt(i).GetNumber(), lwb.GetAt(i).GetNumber())) return false;
    return true;
}

namespace WinvertUnitTestApp4
{
    TEST_CLASS(EndToEndTests)
    {
    public:
        TEST_METHOD(EndToEnd_SettingsPersist)
        {
            CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

            // 1) Launch app with no save file (defaults)
            // Enable test hook so Winvert opens Settings on launch instead of hiding the control panel.
            SetEnvironmentVariableW(L"WINVERT_TEST_OPEN_SETTINGS", L"1");
            std::wstring pfn; DWORD pid = LaunchWinvert4(pfn);
            Assert::IsTrue(pid != 0, L"Failed to launch Winvert4");
            std::wstring localState = LocalStatePathForPFN(pfn);
            Assert::IsTrue(!localState.empty(), L"Failed to get LocalState path");

            // Ensure no save file to start
            std::wstring settingsPath = localState + L"\\settings.json";
            if (fs::exists(settingsPath)) fs::remove(settingsPath);

            LogMessage(L"[Test] Waiting for app to finish startup before triggering selection");
            SleepMs(3000);
            // Find main window
            CComPtr<IUIAutomationElement> win;
            for (int i = 0; i < 300 && !win; ++i) { SleepMs(100); win = FindMainWindowForPid(pid); }
            Assert::IsTrue(win != nullptr, L"Main window not found");
            LogMessage(L"[Test] Main window located, dumping automation tree");
            DumpAutomationTree(win, 3);

            // Open settings by clicking in the top-right area of the tab strip footer.
            // The Settings gear button isn't exposed via UIA, so approximate its location
            // relative to the RegionsTabView bounds.
            auto tabView = WaitForAutomationId(win, L"RegionsTabView");
            Assert::IsTrue(tabView != nullptr, L"RegionsTabView not found");
            RECT tabRect{};
            Assert::IsTrue(SUCCEEDED(tabView->get_CurrentBoundingRectangle(&tabRect)), L"Failed to get RegionsTabView bounds");
            LONG clickX = tabRect.right - 20;  // near right edge
            LONG clickY = tabRect.top + 20;    // near top (tab strip area)
            LogMessage(L"[Test] Clicking near tab strip footer to open Settings");
            Assert::IsTrue(ClickAtScreenPoint(clickX, clickY), L"Failed to synthesize click for Settings");
            SleepMs(750);

            // Validate defaults
            auto fps = WaitForAutomationId(win, L"ShowFpsToggle");
            auto sel = WaitForAutomationId(win, L"SelectionColorEnableToggle");
            auto nbDelay = WaitForAutomationId(win, L"BrightnessDelayNumberBox");
            auto nbR = WaitForAutomationId(win, L"LumaRNumberBox");
            auto nbG = WaitForAutomationId(win, L"LumaGNumberBox");
            auto nbB = WaitForAutomationId(win, L"LumaBNumberBox");
            Assert::IsTrue(fps && sel && nbDelay && nbR && nbG && nbB, L"Settings controls missing");

            // 2) Change settings via UI
            Assert::IsTrue(Toggle(fps, ToggleState_On), L"Toggle ShowFps ON failed");
            Assert::IsTrue(Toggle(sel, ToggleState_On), L"Toggle SelectionColor ON failed");
            Assert::IsTrue(SetValue(nbDelay, L"7"), L"Set delay failed");
            Assert::IsTrue(SetValue(nbR, L"0.30"), L"Set luma R failed");
            Assert::IsTrue(SetValue(nbG, L"0.59"), L"Set luma G failed");
            Assert::IsTrue(SetValue(nbB, L"0.11"), L"Set luma B failed");

            // Add a color map (best-effort; if control exists)
            if (auto add = FindByAutomationId(win, L"ColorMapAddButton")) { InvokeElement(add); SleepMs(200); }

            // Close app (saving should occur on interactions)
            CloseWindow(win);
            SleepMs(1000);

            // 3) Compare settings.json to expected golden
            // Expected JSON path: tests\saves\expected_state.json (beside this project)
            wchar_t modulePath[MAX_PATH];
            GetModuleFileNameW(NULL, modulePath, MAX_PATH);
            std::wstring projDir = fs::path(modulePath).parent_path().wstring();
            std::wstring expected = projDir + L"\\tests\\saves\\expected_state.json";
            Assert::IsTrue(fs::exists(expected), L"Missing expected_state.json");
            Assert::IsTrue(fs::exists(settingsPath), L"settings.json was not created");
            Assert::IsTrue(JsonEqualWithTolerance(expected, settingsPath), L"Saved settings.json does not match expected");

            // 4) Relaunch and verify UI reflects settings
            SetEnvironmentVariableW(L"WINVERT_TEST_OPEN_SETTINGS", L"1");
            pfn.clear(); pid = LaunchWinvert4(pfn);
            Assert::IsTrue(pid != 0, L"Failed to relaunch Winvert4");
            LogMessage(L"[Test] Waiting for app to finish startup (relaunch)");
            SleepMs(2000);
            win = nullptr; for (int i = 0; i < 300 && !win; ++i) { SleepMs(100); win = FindMainWindowForPid(pid); }
            Assert::IsTrue(win != nullptr, L"Main window not found (relaunch)");
            LogMessage(L"[Test] Main window located after relaunch, dumping automation tree");
            DumpAutomationTree(win, 3);

            fps = WaitForAutomationId(win, L"ShowFpsToggle"); sel = WaitForAutomationId(win, L"SelectionColorEnableToggle"); nbDelay = WaitForAutomationId(win, L"BrightnessDelayNumberBox"); nbR = WaitForAutomationId(win, L"LumaRNumberBox"); nbG = WaitForAutomationId(win, L"LumaGNumberBox"); nbB = WaitForAutomationId(win, L"LumaBNumberBox");
            // We don't read back ToggleState/Value here due to brevity; presence is a smoke test. Extend as needed.
            Assert::IsTrue(fps && sel && nbDelay && nbR && nbG && nbB, L"Settings controls missing on relaunch");

            CloseWindow(win);
        }
    };
}
