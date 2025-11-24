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
#include <cmath>
#include <sstream>
#include <unordered_map>
#include <cstdint>

#include <winrt/base.h>
#include <winrt/Windows.Data.Json.h>
#pragma comment(lib, "windowsapp.lib")

#include "CppUnitTest.h"
using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace fs = std::filesystem;

static void SleepMs(int ms) { std::this_thread::sleep_for(std::chrono::milliseconds(ms)); }

// Returns the directory of the current test module (DLL), not the IDE/test host EXE.
static fs::path GetCurrentModuleDirectory()
{
    HMODULE hMod = nullptr;
    if (!GetModuleHandleExW(
            GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            reinterpret_cast<LPCWSTR>(&GetCurrentModuleDirectory),
            &hMod))
    {
        return {};
    }

    wchar_t path[MAX_PATH]{};
    DWORD len = GetModuleFileNameW(hMod, path, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) return {};
    return fs::path(path).parent_path();
}

static std::wstring GetUserProfile()
{
    wchar_t buf[MAX_PATH]; DWORD n = GetEnvironmentVariableW(L"USERPROFILE", buf, MAX_PATH);
    if (n == 0) return L"";
    return std::wstring(buf);
}

static std::wstring LocalStatePathForPFN(const std::wstring& pfn)
{
    auto user = GetUserProfile();
    if (user.empty() || pfn.empty()) return L"";
    return user + L"\\AppData\\Local\\Packages\\" + pfn + L"\\LocalState";
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

static DWORD LaunchWinvert4(std::wstring& outPfn)
{
    CComPtr<IApplicationActivationManager> aam;
    HRESULT hr = CoCreateInstance(CLSID_ApplicationActivationManager, nullptr,
                                  CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&aam));
    if (FAILED(hr)) return 0;

    constexpr wchar_t kWinvert4Aumid[] =
        L"b27d31cf-c66d-45ac-aad0-e0d9501a1c90_ft4zefc91v2gy!App";

    DWORD pid = 0;
    hr = aam->ActivateApplication(kWinvert4Aumid, nullptr, AO_NONE, &pid);
    if (FAILED(hr) || pid == 0) return 0;

    outPfn = PFNForProcess(pid);
    return pid;
}

static void LogMessage(const std::wstring& msg)
{
    Logger::WriteMessage(msg.c_str());
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
    } ctx{ pid, nullptr, nullptr };

    EnumWindows(
        [](HWND hwnd, LPARAM lParam)->BOOL
        {
            auto* ctx = reinterpret_cast<EnumCtx*>(lParam);
            DWORD procId = 0;
            GetWindowThreadProcessId(hwnd, &procId);
            if (procId != ctx->targetPid) return TRUE;
            if (!IsCandidateMainWindow(hwnd)) return TRUE;

            ctx->found = hwnd;
            return FALSE;
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
        return reinterpret_cast<HWND>(static_cast<LONG_PTR>(v.lVal));
    if (v.vt == VT_I8 || v.vt == VT_UI8)
        return reinterpret_cast<HWND>(static_cast<LONG_PTR>(v.llVal));
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
    CComPtr<IUIAutomation> uia;
    if (FAILED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uia))))
        return nullptr;

    CComPtr<IUIAutomationElement> root;
    if (FAILED(uia->GetRootElement(&root)) || !root) return nullptr;

    VARIANT vPid; VariantInit(&vPid); vPid.vt = VT_I4; vPid.lVal = static_cast<LONG>(pid);
    CComPtr<IUIAutomationCondition> pidCond;
    uia->CreatePropertyCondition(UIA_ProcessIdPropertyId, vPid, &pidCond);

    VARIANT vType; VariantInit(&vType); vType.vt = VT_I4; vType.lVal = UIA_WindowControlTypeId;
    CComPtr<IUIAutomationCondition> typeCond;
    uia->CreatePropertyCondition(UIA_ControlTypePropertyId, vType, &typeCond);

    CComPtr<IUIAutomationCondition> cond;
    uia->CreateAndCondition(pidCond, typeCond, &cond);

    CComPtr<IUIAutomationElement> win;
    root->FindFirst(TreeScope_Subtree, cond, &win);
    if (win)
    {
        EnsureWindowVisible(win);
        return win;
    }

    HWND hwnd = FindTopLevelWindowForPid(pid);
    if (!hwnd) return nullptr;

    CComPtr<IUIAutomationElement> fromHandle;
    if (FAILED(uia->ElementFromHandle(hwnd, &fromHandle))) return nullptr;
    EnsureWindowVisible(fromHandle);
    return fromHandle;
}

static CComPtr<IUIAutomationElement> FindByAutomationId(IUIAutomationElement* parent, const wchar_t* aid)
{
    if (!parent) return nullptr;

    CComPtr<IUIAutomation> uia;
    if (FAILED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uia))))
        return nullptr;

    CComPtr<IUIAutomationCondition> cond;
    uia->CreatePropertyCondition(UIA_AutomationIdPropertyId, CComVariant(aid), &cond);

    CComPtr<IUIAutomationElement> el;
    parent->FindFirst(TreeScope_Subtree, cond, &el);
    return el;
}

static CComPtr<IUIAutomationElement> FindByAutomationIdNameClass(
    IUIAutomationElement* parent,
    const wchar_t* aid,
    const wchar_t* name,
    const wchar_t* cls)
{
    if (!parent) return nullptr;

    CComPtr<IUIAutomation> uia;
    if (FAILED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uia))))
        return nullptr;

    CComPtr<IUIAutomationCondition> conds[3];
    int count = 0;

    if (aid && *aid)
    {
        uia->CreatePropertyCondition(UIA_AutomationIdPropertyId, CComVariant(aid), &conds[count++]);
    }
    if (name && *name)
    {
        uia->CreatePropertyCondition(UIA_NamePropertyId, CComVariant(name), &conds[count++]);
    }
    if (cls && *cls)
    {
        uia->CreatePropertyCondition(UIA_ClassNamePropertyId, CComVariant(cls), &conds[count++]);
    }

    if (count == 0) return nullptr;

    CComPtr<IUIAutomationCondition> combined = conds[0];
    for (int i = 1; i < count; ++i)
    {
        CComPtr<IUIAutomationCondition> tmp;
        uia->CreateAndCondition(combined, conds[i], &tmp);
        combined = tmp;
    }

    CComPtr<IUIAutomationElement> el;
    parent->FindFirst(TreeScope_Subtree, combined, &el);
    return el;
}

static CComPtr<IUIAutomationElement> WaitForAutomationId(IUIAutomationElement* parent, const wchar_t* aid, int timeoutMs = 5000, int pollMs = 100)
{
    if (!parent) return nullptr;
    int iterations = timeoutMs / pollMs;
    for (int i = 0; i < iterations; ++i)
    {
        if (auto el = FindByAutomationId(parent, aid)) return el;
        SleepMs(pollMs);
    }
    return FindByAutomationId(parent, aid);
}

static CComPtr<IUIAutomationElement> WaitForAutomationIdNameClass(
    IUIAutomationElement* parent,
    const wchar_t* aid,
    const wchar_t* name,
    const wchar_t* cls,
    int timeoutMs = 5000,
    int pollMs = 100)
{
    if (!parent) return nullptr;
    int iterations = timeoutMs / pollMs;
    for (int i = 0; i < iterations; ++i)
    {
        if (auto el = FindByAutomationIdNameClass(parent, aid, name, cls)) return el;
        SleepMs(pollMs);
    }
    return FindByAutomationIdNameClass(parent, aid, name, cls);
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
    if (SUCCEEDED(el->GetCurrentPatternAs(UIA_InvokePatternId, IID_PPV_ARGS(&pat))) && pat)
    {
        if (SUCCEEDED(pat->Invoke())) return true;
        SleepMs(250);
        if (SUCCEEDED(pat->Invoke())) return true;
    }

    POINT pt{}; BOOL got = FALSE;
    if (SUCCEEDED(el->GetClickablePoint(&pt, &got)) && got)
    {
        if (ClickAtScreenPoint(pt.x, pt.y)) return true;
    }

    RECT r{};
    if (SUCCEEDED(el->get_CurrentBoundingRectangle(&r)))
    {
        LONG cx = (r.left + r.right) / 2;
        LONG cy = (r.top + r.bottom) / 2;
        if (ClickAtScreenPoint(cx, cy)) return true;
    }

    return false;
}

static bool Toggle(IUIAutomationElement* el, ToggleState desired)
{
    if (!el) return false;
    CComPtr<IUIAutomationTogglePattern> pat;
    if (FAILED(el->GetCurrentPatternAs(UIA_TogglePatternId, IID_PPV_ARGS(&pat))) || !pat) return false;
    ToggleState cur{};
    pat->get_CurrentToggleState(&cur);
    if (cur != desired)
    {
        if (FAILED(pat->Toggle())) return false;
        SleepMs(100);
    }
    return true;
}

static bool SetValue(IUIAutomationElement* el, const wchar_t* value)
{
    if (!el) return false;
    CComPtr<IUIAutomationValuePattern> pat;
    if (FAILED(el->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&pat))) || !pat) return false;
    CComBSTR b(value);
    return SUCCEEDED(pat->SetValue(b));
}

static bool CloseWindow(IUIAutomationElement* win)
{
    if (!win) return false;
    CComPtr<IUIAutomationWindowPattern> pat;
    if (FAILED(win->GetCurrentPatternAs(UIA_WindowPatternId, IID_PPV_ARGS(&pat))) || !pat) return false;
    return SUCCEEDED(pat->Close());
}

static void DumpAllAutomationElements(IUIAutomationElement* root)
{
    if (!root) return;

    CComPtr<IUIAutomation> uia;
    if (FAILED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uia))))
        return;

    CComPtr<IUIAutomationCondition> cond;
    if (FAILED(uia->CreateTrueCondition(&cond)) || !cond)
        return;

    CComPtr<IUIAutomationElementArray> arr;
    if (FAILED(root->FindAll(TreeScope_Subtree, cond, &arr)) || !arr)
        return;

    int length = 0;
    arr->get_Length(&length);

    LogMessage(L"[Test] DumpAllAutomationElements: count=" + std::to_wstring(length));

    for (int i = 0; i < length; ++i)
    {
        CComPtr<IUIAutomationElement> el;
        if (FAILED(arr->GetElement(i, &el)) || !el) continue;

        CComVariant name, autoId, cls, ctrlType;
        el->GetCurrentPropertyValue(UIA_NamePropertyId, &name);
        el->GetCurrentPropertyValue(UIA_AutomationIdPropertyId, &autoId);
        el->GetCurrentPropertyValue(UIA_ClassNamePropertyId, &cls);
        el->GetCurrentPropertyValue(UIA_ControlTypePropertyId, &ctrlType);

        std::wstring line = L"[Test] ALL: #" + std::to_wstring(i) + L" ";
        if (name.vt == VT_BSTR && name.bstrVal)
            line += L"name=\"" + std::wstring(name.bstrVal) + L"\" ";
        if (autoId.vt == VT_BSTR && autoId.bstrVal)
            line += L"id=" + std::wstring(autoId.bstrVal) + L" ";
        if (cls.vt == VT_BSTR && cls.bstrVal)
            line += L"class=" + std::wstring(cls.bstrVal) + L" ";
        if (ctrlType.vt == VT_I4)
            line += L"type=" + std::to_wstring(ctrlType.lVal) + L" ";
        LogMessage(line);
    }
}

static void ExpandAllSettingsExpanders(IUIAutomationElement* root)
{
    if (!root) return;
    const wchar_t* kExpanders[] =
    {
        L"SelectionColorExpander",
        L"BrightnessExpander",
        L"HotkeysExpander",
        L"CustomFiltersExpander",
        L"ColorMappingExpander"
    };

    for (auto id : kExpanders)
    {
        auto expander = FindByAutomationId(root, id);
        if (!expander) continue;

        CComPtr<IUIAutomationExpandCollapsePattern> pat;
        if (SUCCEEDED(expander->GetCurrentPatternAs(UIA_ExpandCollapsePatternId, IID_PPV_ARGS(&pat))) && pat)
        {
            ExpandCollapseState state{};
            if (SUCCEEDED(pat->get_CurrentExpandCollapseState(&state)) &&
                (state == ExpandCollapseState_Collapsed || state == ExpandCollapseState_PartiallyExpanded))
            {
                pat->Expand();
                SleepMs(200);
            }
        }
        else
        {
            InvokeElement(expander);
            SleepMs(200);
        }
    }
}

static std::string ReadAllUtf8(const std::wstring& path)
{
    FILE* f{}; _wfopen_s(&f, path.c_str(), L"rb"); if (!f) return {};
    fseek(f, 0, SEEK_END); long sz = ftell(f); fseek(f, 0, SEEK_SET);
    std::string s; s.resize(sz);
    fread(s.data(), 1, sz, f); fclose(f);
    return s;
}

// Simple CRC32 for quick equality check
static uint32_t Crc32(const uint8_t* data, size_t len)
{
    static uint32_t table[256];
    static bool init = false;
    if (!init)
    {
        for (uint32_t i = 0; i < 256; ++i)
        {
            uint32_t c = i;
            for (int j = 0; j < 8; ++j)
                c = (c & 1) ? (0xEDB88320u ^ (c >> 1)) : (c >> 1);
            table[i] = c;
        }
        init = true;
    }
    uint32_t crc = 0xFFFFFFFFu;
    for (size_t i = 0; i < len; ++i)
        crc = table[(crc ^ data[i]) & 0xFF] ^ (crc >> 8);
    return crc ^ 0xFFFFFFFFu;
}

static void FlattenJsonValueWinRT(
    winrt::Windows::Data::Json::IJsonValue const& value,
    const std::wstring& path,
    std::unordered_map<std::wstring, std::wstring>& out)
{
    using namespace winrt::Windows::Data::Json;

    switch (value.ValueType())
    {
    case JsonValueType::Object:
    {
        JsonObject obj = value.GetObject();
        for (auto const& kvp : obj)
        {
            std::wstring key = kvp.Key().c_str();
            std::wstring childPath = path.empty() ? key : (path + L"." + key);
            FlattenJsonValueWinRT(kvp.Value(), childPath, out);
        }
        break;
    }
    case JsonValueType::Array:
    {
        JsonArray arr = value.GetArray();
        for (uint32_t i = 0; i < arr.Size(); ++i)
        {
            std::wstring childPath = path + L"[" + std::to_wstring(i) + L"]";
            FlattenJsonValueWinRT(arr.GetAt(i), childPath, out);
        }
        break;
    }
    case JsonValueType::Null:
    case JsonValueType::Boolean:
    case JsonValueType::Number:
    case JsonValueType::String:
    default:
    {
        if (!path.empty())
        {
            std::wstring v = value.Stringify().c_str();
            out[path] = v;
        }
        break;
    }
    }
}

static bool CompareFilesAndLog(const std::wstring& expectedPath, const std::wstring& actualPath)
{
    auto e = ReadAllUtf8(expectedPath);
    auto a = ReadAllUtf8(actualPath);

    // Fast path: CRC match -> treat as equal.
    uint32_t crcE = Crc32(reinterpret_cast<const uint8_t*>(e.data()), e.size());
    uint32_t crcA = Crc32(reinterpret_cast<const uint8_t*>(a.data()), a.size());
    if (crcE == crcA)
        return true;

    LogMessage(L"[Test] settings.json mismatch detected (CRC differs)");

    // Parse via WinRT JSON
    winrt::init_apartment(winrt::apartment_type::single_threaded);

    winrt::Windows::Data::Json::JsonObject je;
    winrt::Windows::Data::Json::JsonObject ja;
    try
    {
        je = winrt::Windows::Data::Json::JsonObject::Parse(winrt::to_hstring(e));
        ja = winrt::Windows::Data::Json::JsonObject::Parse(winrt::to_hstring(a));
    }
    catch (winrt::hresult_error const& ex)
    {
        std::wstring msg = L"[Test] JSON parse threw hresult_error: ";
        msg += ex.message().c_str();
        LogMessage(msg);
        return false;
    }

    std::unordered_map<std::wstring, std::wstring> fieldsExpected;
    std::unordered_map<std::wstring, std::wstring> fieldsActual;

    FlattenJsonValueWinRT(je, L"", fieldsExpected);
    FlattenJsonValueWinRT(ja, L"", fieldsActual);

    bool ok = true;

    // Fields missing or different vs default
    for (const auto& kv : fieldsExpected)
    {
        auto it = fieldsActual.find(kv.first);
        if (it == fieldsActual.end())
        {
            std::wstring msg = L"[Test] settings.json missing field: ";
            msg += kv.first;
            LogMessage(msg);
            ok = false;
        }
        else if (it->second != kv.second)
        {
            std::wstring msg = L"[Test] settings.json field mismatch at ";
            msg += kv.first;
            msg += L": expected=";
            msg += kv.second;
            msg += L", actual=";
            msg += it->second;
            LogMessage(msg);
            ok = false;
        }
    }

    // Extra fields present in actual
    for (const auto& kv : fieldsActual)
    {
        if (fieldsExpected.find(kv.first) == fieldsExpected.end())
        {
            std::wstring msg = L"[Test] settings.json has extra field: ";
            msg += kv.first;
            msg += L" = ";
            msg += kv.second;
            LogMessage(msg);
            ok = false;
        }
    }

    return ok;
}

struct AppCloser
{
    CComPtr<IUIAutomationElement> win;
    bool closed{ false };
    ~AppCloser()
    {
        if (win && !closed)
        {
            CloseWindow(win);
        }
    }
};

namespace WinvertUnitTestApp4
{
    TEST_CLASS(EndToEndTests)
    {
    public:
        TEST_METHOD(Winvert_Launch)
        {
            LogMessage(L"[Test] Running Winvert_Launch");
            CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
            AppCloser closer;

            std::wstring pfn; DWORD pid = LaunchWinvert4(pfn);
            Assert::IsTrue(pid != 0, L"Failed to launch Winvert4");

            LogMessage(L"[Test] Winvert_Launch: waiting for main window");
            CComPtr<IUIAutomationElement> win;
            for (int i = 0; i < 300 && !win; ++i)
            {
                SleepMs(100);
                win = FindMainWindowForPid(pid);
            }
            Assert::IsTrue(win != nullptr, L"Main window not found");
            closer.win = win;

            LogMessage(L"[Test] Winvert_Launch: dumping main window elements");
            DumpAllAutomationElements(win);

            const wchar_t* mainIds[] =
            {
                L"RegionsTabView",
                L"AddButton",
                L"SettingsButton",
                L"InvertEffectButton",
                L"BrightnessProtectionButton",
                L"FiltersDropDownButton",
                L"HideAllWindowsButton",
                L"ColorMappingToggleButton"
            };
            for (auto id : mainIds)
            {
                auto el = WaitForAutomationId(win, id, 3000);
                Assert::IsTrue(el != nullptr, (std::wstring(L"Main window element missing: ") + id).c_str());
            }

            CloseWindow(win);
            closer.closed = true;
        }

        TEST_METHOD(Save_File_Default)
        {
            LogMessage(L"[Test] Running Save_File_Default");
            CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
            AppCloser closer;

            // 1) Ensure no save file to start (compute path from known package family)
            std::wstring packageFamily = L"b27d31cf-c66d-45ac-aad0-e0d9501a1c90_ft4zefc91v2gy";
            std::wstring localState = LocalStatePathForPFN(packageFamily);
            Assert::IsTrue(!localState.empty(), L"Failed to get LocalState path");
            std::wstring settingsPath = localState + L"\\settings.json";
            if (fs::exists(settingsPath)) fs::remove(settingsPath);
            LogMessage(std::wstring(L"[Test] settingsPath=") + settingsPath);

            // 2) Launch app with no save file (defaults)
            std::wstring pfn; DWORD pid = LaunchWinvert4(pfn);
            Assert::IsTrue(pid != 0, L"Failed to launch Winvert4");

            LogMessage(L"[Test] Waiting for app to finish startup");
            SleepMs(3000);

            // Find main window
            CComPtr<IUIAutomationElement> win;
            for (int i = 0; i < 300 && !win; ++i) { SleepMs(100); win = FindMainWindowForPid(pid); }
            Assert::IsTrue(win != nullptr, L"Main window not found");
            closer.win = win;

            // Open settings via the visible SettingsButton in the tab strip footer.
            auto settingsBtn = WaitForAutomationId(win, L"SettingsButton");
            Assert::IsTrue(settingsBtn != nullptr, L"SettingsButton not found");
            Assert::IsTrue(InvokeElement(settingsBtn), L"Failed to click SettingsButton");
            SleepMs(750);

            // Expand all settings expanders so nested controls are realized in the tree.
            LogMessage(L"[Test] Expanding all settings expanders");
            ExpandAllSettingsExpanders(win);

            LogMessage(L"[Test] Dumping elements after opening Settings");
            DumpAllAutomationElements(win);

            // Verify that expected settings-page elements exist (but do not modify them)
            struct ExpectedElement { const wchar_t* id; const wchar_t* name; const wchar_t* cls; };
            const ExpectedElement settingsExpected[] =
            // ID, NAME, CLASS
            {
                { L"BackButton",              L"Back",                      L"Button" },
                { L"SelectionColorExpander",  nullptr,                      L"Microsoft.UI.Xaml.Controls.Expander" },
                { L"SelectionColorEnableToggle", nullptr,                   L"ToggleSwitch" },
                { L"ColorSpectrum",           L"Color picker",              L"Microsoft.UI.Xaml.Controls.Primitives.ColorSpectrum" },
                { L"BrightnessExpander",      L"Brightness protection",     L"Microsoft.UI.Xaml.Controls.Expander" },
                { L"BrightnessDelayNumberBox",nullptr,                      L"Microsoft.UI.Xaml.Controls.NumberBox" },
                { L"BrightnessResetButton",   L"Reset defaults",            L"Button" },
                { L"LumaRNumberBox",          L"R",                         L"Microsoft.UI.Xaml.Controls.NumberBox" },
                { L"LumaGNumberBox",          L"G",                         L"Microsoft.UI.Xaml.Controls.NumberBox" },
                { L"LumaBNumberBox",          L"B",                         L"Microsoft.UI.Xaml.Controls.NumberBox" },
                { L"ShowFpsToggle",           nullptr,                      L"ToggleSwitch" },
                { L"HotkeysExpander",         nullptr,                      L"Microsoft.UI.Xaml.Controls.Expander" },
                { L"RunAtStartupToggle",      nullptr,                      L"ToggleSwitch" },
                { L"InvertHotkeyTextBox",     nullptr,                      L"TextBox" },
                { L"RebindInvertHotkeyButton",L"Rebind Invert/Add",         L"Button" },
                { L"FilterHotkeyTextBox",     nullptr,                      L"TextBox" },
                { L"RebindFilterHotkeyButton",L"Rebind Filter/Add",         L"Button" },
                { L"RemoveHotkeyTextBox",     nullptr,                      L"TextBox" },
                { L"RebindRemoveHotkeyButton",L"Rebind Remove Last",        L"Button" },
                { L"CustomFiltersExpander",   L"Custom filters",            L"Microsoft.UI.Xaml.Controls.Expander" },
                { L"FavoriteFilterComboBox",  nullptr,                      L"ComboBox" },
                { L"SavedFiltersComboBox",    nullptr,                      L"ComboBox" },
                { L"SimpleResetButton",       L"Reset",                     L"Button" },
                { L"BrightnessSlider",        nullptr,                      L"Slider" },
                { L"ContrastSlider",          nullptr,                      L"Slider" },
                { L"SaturationSlider",        nullptr,                      L"Slider" },
                { L"HueSlider",               nullptr,                      L"Slider" },
                { L"TemperatureSlider",       nullptr,                      L"Slider" },
                { L"TintSlider",              nullptr,                      L"Slider" },
                { L"ColorMappingExpander",    L"Color mapping",             L"Microsoft.UI.Xaml.Controls.Expander" },
                { L"ColorMapAddButton",       L"Add",                       L"Button" },
                { L"ColorMapSampleButton",    L"Sample",                    L"Button" },
                { L"PreviewColorMapToggle",   L"Preview",                   L"ToggleButton" },
                // Color picker within Color mapping section shares AutomationId "ColorSpectrum"
                // and we already validated one ColorSpectrum above.
                { L"ColorMapPreserveToggle",  L"Preserve brightness",       L"ToggleSwitch" }
            };
            for (const auto& e : settingsExpected)
            {
                auto el = WaitForAutomationIdNameClass(win, e.id, e.name, e.cls, 5000);
                std::wstring desc = L"id=" + std::wstring(e.id ? e.id : L"<null>");
                if (e.name && *e.name)
                {
                    desc += L", name=" + std::wstring(e.name);
                }
                if (e.cls && *e.cls)
                {
                    desc += L", class=" + std::wstring(e.cls);
                }
                Assert::IsTrue(el != nullptr, (std::wstring(L"Settings element missing: ") + desc).c_str());
            }

            // Navigate back to main panel
            auto backBtn = WaitForAutomationId(win, L"BackButton", 5000);
            Assert::IsTrue(backBtn != nullptr, L"BackButton not found");
            Assert::IsTrue(InvokeElement(backBtn), L"Failed to click BackButton");
            SleepMs(500);

            // Close app (saving should occur on interactions)
            CloseWindow(win);
            closer.closed = true;
            SleepMs(1000);

            // 3) Compare settings.json to expected golden
            fs::path moduleDir = GetCurrentModuleDirectory();
            // Walk up until we find the repos root (folder that contains Winvert4 and WinvertUnitTestApp4)
            fs::path cur = moduleDir;
            fs::path reposRoot;
            for (int i = 0; i < 6 && !cur.empty(); ++i)
            {
                auto name = cur.filename().wstring();
                if (name == L"Winvert4" || name == L"WinvertUnitTestApp4")
                {
                    reposRoot = cur.parent_path();
                    break;
                }
                cur = cur.parent_path();
            }
            if (reposRoot.empty())
            {
                Assert::Fail(L"Failed to locate repos root for default.json");
            }

            fs::path expectedPath = reposRoot / L"WinvertUnitTestApp4" / L"saves" / L"default.json";
            std::wstring expected = expectedPath.wstring();
            LogMessage(L"[Test] expected default.json path=" + expected);
            Assert::IsTrue(fs::exists(expected), L"Missing expected_state.json");
            Assert::IsTrue(fs::exists(settingsPath), L"settings.json was not created");
            Assert::IsTrue(CompareFilesAndLog(expected, settingsPath), L"Saved settings.json does not match expected");
        }

        TEST_METHOD(Save_File_Persistance_1)
        {
            LogMessage(L"[Test] Running Save_File_Persistance_1");
            CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
            AppCloser closer;

            // Ensure clean settings.json
            std::wstring packageFamily = L"b27d31cf-c66d-45ac-aad0-e0d9501a1c90_ft4zefc91v2gy";
            std::wstring localState = LocalStatePathForPFN(packageFamily);
            Assert::IsTrue(!localState.empty(), L"Failed to get LocalState path");
            std::wstring settingsPath = localState + L"\\settings.json";
            if (fs::exists(settingsPath)) fs::remove(settingsPath);

            // First launch: modify settings
            std::wstring pfn; DWORD pid = LaunchWinvert4(pfn);
            Assert::IsTrue(pid != 0, L"Failed to launch Winvert4");

            LogMessage(L"[Test] Save_File_Persistance_1: waiting for main window");
            CComPtr<IUIAutomationElement> win;
            for (int i = 0; i < 300 && !win; ++i) { SleepMs(100); win = FindMainWindowForPid(pid); }
            Assert::IsTrue(win != nullptr, L"Main window not found");
            closer.win = win;

            auto settingsBtn = WaitForAutomationId(win, L"SettingsButton");
            Assert::IsTrue(settingsBtn != nullptr, L"SettingsButton not found");
            Assert::IsTrue(InvokeElement(settingsBtn), L"Failed to click SettingsButton");
            SleepMs(750);

            ExpandAllSettingsExpanders(win);

            auto fps = WaitForAutomationId(win, L"ShowFpsToggle");
            auto sel = WaitForAutomationId(win, L"SelectionColorEnableToggle");
            auto nbDelay = WaitForAutomationId(win, L"BrightnessDelayNumberBox");
            auto nbR = WaitForAutomationId(win, L"LumaRNumberBox");
            auto nbG = WaitForAutomationId(win, L"LumaGNumberBox");
            auto nbB = WaitForAutomationId(win, L"LumaBNumberBox");
            Assert::IsTrue(fps && sel && nbDelay && nbR && nbG && nbB, L"Settings controls missing");

            Assert::IsTrue(Toggle(fps, ToggleState_On), L"Toggle ShowFps ON failed");
            Assert::IsTrue(Toggle(sel, ToggleState_On), L"Toggle SelectionColor ON failed");
            Assert::IsTrue(SetValue(nbDelay, L"7"), L"Set delay failed");
            Assert::IsTrue(SetValue(nbR, L"0.30"), L"Set luma R failed");
            Assert::IsTrue(SetValue(nbG, L"0.59"), L"Set luma G failed");
            Assert::IsTrue(SetValue(nbB, L"0.11"), L"Set luma B failed");

            CloseWindow(win);
            closer.closed = true;
            SleepMs(1000);

            // Relaunch: verify persistence via UI
            pfn.clear(); pid = LaunchWinvert4(pfn);
            Assert::IsTrue(pid != 0, L"Failed to relaunch Winvert4");
            win = nullptr; for (int i = 0; i < 300 && !win; ++i) { SleepMs(100); win = FindMainWindowForPid(pid); }
            Assert::IsTrue(win != nullptr, L"Main window not found (relaunch)");
            closer.win = win;

            settingsBtn = WaitForAutomationId(win, L"SettingsButton");
            Assert::IsTrue(settingsBtn != nullptr, L"SettingsButton not found (relaunch)");
            Assert::IsTrue(InvokeElement(settingsBtn), L"Failed to click SettingsButton (relaunch)");
            SleepMs(750);

            ExpandAllSettingsExpanders(win);

            fps = WaitForAutomationId(win, L"ShowFpsToggle");
            sel = WaitForAutomationId(win, L"SelectionColorEnableToggle");
            nbDelay = WaitForAutomationId(win, L"BrightnessDelayNumberBox");
            nbR = WaitForAutomationId(win, L"LumaRNumberBox");
            nbG = WaitForAutomationId(win, L"LumaGNumberBox");
            nbB = WaitForAutomationId(win, L"LumaBNumberBox");
            Assert::IsTrue(fps && sel && nbDelay && nbR && nbG && nbB, L"Settings controls missing on relaunch");

            // Verify toggles and values persisted
            {
                CComPtr<IUIAutomationTogglePattern> tpat;
                ToggleState state{};

                if (fps && SUCCEEDED(fps->GetCurrentPatternAs(UIA_TogglePatternId, IID_PPV_ARGS(&tpat))) && tpat)
                {
                    tpat->get_CurrentToggleState(&state);
                    Assert::IsTrue(state == ToggleState_On, L"ShowFpsToggle did not persist ON");
                }
                tpat.Release();
                if (sel && SUCCEEDED(sel->GetCurrentPatternAs(UIA_TogglePatternId, IID_PPV_ARGS(&tpat))) && tpat)
                {
                    tpat->get_CurrentToggleState(&state);
                    Assert::IsTrue(state == ToggleState_On, L"SelectionColorEnableToggle did not persist ON");
                }
            }
            {
                CComPtr<IUIAutomationValuePattern> vpat;
                CComBSTR val;

                if (nbDelay && SUCCEEDED(nbDelay->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&vpat))) && vpat)
                {
                    vpat->get_CurrentValue(&val);
                    std::wstring s(val);
                    wchar_t* end = nullptr;
                    double d = wcstod(s.c_str(), &end);
                    Assert::IsTrue(std::fabs(d - 7.0) < 0.001, L"BrightnessDelayNumberBox did not persist value ~7");
                }
                vpat.Release(); val.Empty();
                if (nbR && SUCCEEDED(nbR->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&vpat))) && vpat)
                {
                    vpat->get_CurrentValue(&val);
                    std::wstring s(val);
                    wchar_t* end = nullptr;
                    double d = wcstod(s.c_str(), &end);
                    Assert::IsTrue(std::fabs(d - 0.30) < 0.001, L"LumaRNumberBox did not persist value ~0.30");
                }
                vpat.Release(); val.Empty();
                if (nbG && SUCCEEDED(nbG->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&vpat))) && vpat)
                {
                    vpat->get_CurrentValue(&val);
                    std::wstring s(val);
                    wchar_t* end = nullptr;
                    double d = wcstod(s.c_str(), &end);
                    Assert::IsTrue(std::fabs(d - 0.59) < 0.001, L"LumaGNumberBox did not persist value ~0.59");
                }
                vpat.Release(); val.Empty();
                if (nbB && SUCCEEDED(nbB->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&vpat))) && vpat)
                {
                    vpat->get_CurrentValue(&val);
                    std::wstring s(val);
                    wchar_t* end = nullptr;
                    double d = wcstod(s.c_str(), &end);
                    Assert::IsTrue(std::fabs(d - 0.11) < 0.001, L"LumaBNumberBox did not persist value ~0.11");
                }
            }

            CloseWindow(win);
            closer.closed = true;
        }
    };
}
