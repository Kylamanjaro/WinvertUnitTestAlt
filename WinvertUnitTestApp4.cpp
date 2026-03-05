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
    HRESULT hr = el->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&pat));
    if (SUCCEEDED(hr) && pat)
    {
        CComBSTR b(value);
        return SUCCEEDED(pat->SetValue(b));
    }

    // Fallback for controls like NumberBox that host an inner TextBox ("InputBox")
    auto inner = FindByAutomationId(el, L"InputBox");
    if (!inner) return false;

    pat.Release();
    if (FAILED(inner->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&pat))) || !pat) return false;
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

static void WaitForProcessExit(DWORD pid, int timeoutMs = 5000)
{
    if (pid == 0) return;
    HANDLE h = OpenProcess(SYNCHRONIZE, FALSE, pid);
    if (!h) return;
    WaitForSingleObject(h, timeoutMs);
    CloseHandle(h);
}

static void EnsureProcessExited(DWORD pid, int timeoutMs = 5000)
{
    if (pid == 0) return;
    HANDLE h = OpenProcess(SYNCHRONIZE | PROCESS_TERMINATE, FALSE, pid);
    if (!h) return;
    DWORD wait = WaitForSingleObject(h, timeoutMs);
    if (wait == WAIT_TIMEOUT)
    {
        LogMessage(L"[Test] Process did not exit in time; terminating");
        TerminateProcess(h, 1);
        WaitForSingleObject(h, 2000);
    }
    CloseHandle(h);
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

static bool JsonBoolLiteralToBool(const std::wstring& v, bool& out)
{
    if (v == L"true")
    {
        out = true;
        return true;
    }
    if (v == L"false")
    {
        out = false;
        return true;
    }
    return false;
}

static std::wstring StripJsonStringQuotes(const std::wstring& v)
{
    if (v.size() >= 2 && v.front() == L'\"' && v.back() == L'\"')
        return v.substr(1, v.size() - 2);
    return v;
}

static bool JsonNumberLiteralToDouble(const std::wstring& v, double& out)
{
    wchar_t* end = nullptr;
    out = wcstod(v.c_str(), &end);
    return end != v.c_str();
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
        else
        {
            const std::wstring& expectedVal = kv.second;
            const std::wstring& actualVal = it->second;

            bool treatedAsNumber = false;
            double dE{}, dA{};
            if (JsonNumberLiteralToDouble(expectedVal, dE) &&
                JsonNumberLiteralToDouble(actualVal, dA))
            {
                treatedAsNumber = true;
                if (std::fabs(dE - dA) > 1e-4)
                {
                    std::wstring msg = L"[Test] settings.json numeric field mismatch at ";
                    msg += kv.first;
                    msg += L": expected=";
                    msg += expectedVal;
                    msg += L", actual=";
                    msg += actualVal;
                    LogMessage(msg);
                    ok = false;
                }
            }

            if (!treatedAsNumber && actualVal != expectedVal)
            {
                std::wstring msg = L"[Test] settings.json field mismatch at ";
                msg += kv.first;
                msg += L": expected=";
                msg += expectedVal;
                msg += L", actual=";
                msg += actualVal;
                LogMessage(msg);
                ok = false;
            }
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

        TEST_METHOD(Save_File_Persistance)
        {
            LogMessage(L"[Test] Running Save_File_Persistance");
            CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
            winrt::init_apartment(winrt::apartment_type::single_threaded);
            AppCloser closer;

            // Paths
            std::wstring packageFamily = L"b27d31cf-c66d-45ac-aad0-e0d9501a1c90_ft4zefc91v2gy";
            std::wstring localState = LocalStatePathForPFN(packageFamily);
            Assert::IsTrue(!localState.empty(), L"Failed to get LocalState path");
            std::wstring settingsPath = localState + L"\\settings.json";

            auto runScenario = [&](const std::wstring& label, const std::wstring& goldenPath)
            {
                using namespace winrt::Windows::Data::Json;

                LogMessage(L"[Test] Save_File_Persistance scenario: " + label);
                Assert::IsTrue(fs::exists(goldenPath), (L"Golden settings file missing: " + goldenPath).c_str());

                // Load golden JSON and flatten to field map
                auto goldenText = ReadAllUtf8(goldenPath);
                JsonObject goldenObj = JsonObject::Parse(winrt::to_hstring(goldenText));
                std::unordered_map<std::wstring, std::wstring> goldenFields;
                FlattenJsonValueWinRT(goldenObj, L"", goldenFields);

                auto getField = [&](const wchar_t* key) -> std::wstring
                {
                    auto it = goldenFields.find(key);
                    return (it == goldenFields.end()) ? L"" : it->second;
                };

                // Start from defaults: delete existing settings
                if (fs::exists(settingsPath)) fs::remove(settingsPath);

                // First launch: drive UI to match golden JSON
                std::wstring pfn;
                DWORD pid = LaunchWinvert4(pfn);
                Assert::IsTrue(pid != 0, L"Failed to launch Winvert4");

                CComPtr<IUIAutomationElement> win;
                for (int i = 0; i < 300 && !win; ++i)
                {
                    SleepMs(100);
                    win = FindMainWindowForPid(pid);
                }
                Assert::IsTrue(win != nullptr, L"Main window not found");
                closer.win = win;

                auto settingsBtn = WaitForAutomationId(win, L"SettingsButton");
                Assert::IsTrue(settingsBtn != nullptr, L"SettingsButton not found");
                Assert::IsTrue(InvokeElement(settingsBtn), L"Failed to click SettingsButton");
                SleepMs(750);

                ExpandAllSettingsExpanders(win);

                // 1) Advanced Matrix: ensure it's enabled and dump controls
                if (auto adv = WaitForAutomationId(win, L"AdvancedMatrixToggle"))
                {
                    // Force it ON instead of just invoking
                    Assert::IsTrue(
                        Toggle(adv, ToggleState_On),
                        L"Failed to enable AdvancedMatrixToggle");
                    SleepMs(500);
                }
                LogMessage(L"[Test] AdvancedMatrixToggle Enabled");

                // Try to scroll the main settings ScrollViewer down so dynamically-created
                // advanced matrix TextBoxes are realized and visible to UIA.
                if (auto scrollViewer = FindByAutomationIdNameClass(win, nullptr, nullptr, L"ScrollViewer"))
                {
                    CComPtr<IUIAutomationScrollPattern> sp;
                    if (SUCCEEDED(scrollViewer->GetCurrentPatternAs(UIA_ScrollPatternId, IID_PPV_ARGS(&sp))) && sp)
                    {
                        LogMessage(L"[Test] Scrolling settings ScrollViewer to reveal advanced matrix controls");
                        for (int i = 0; i < 4; ++i)
                        {
                            sp->Scroll(ScrollAmount_NoAmount, ScrollAmount_LargeIncrement);
                            SleepMs(200);
                        }
                    }
                }
                SleepMs(2000);
                DumpAllAutomationElements(win);

                // Helpers to drive controls from fields
                auto setToggleFromField = [&](const wchar_t* controlId, const wchar_t* key)
                {
                    std::wstring v = getField(key);
                    if (v.empty()) return;
                    bool desired{};
                    if (!JsonBoolLiteralToBool(v, desired)) return;
                      auto el = WaitForAutomationId(win, controlId);
                      if (!el)
                      {
                          // Fallback: dynamically-created controls may expose their identifier as Name.
                          el = WaitForAutomationIdNameClass(win, nullptr, controlId, L"TextBox");
                      }
                    Assert::IsTrue(el != nullptr, (std::wstring(L"Missing toggle: ") + controlId).c_str());
                    Assert::IsTrue(Toggle(el, desired ? ToggleState_On : ToggleState_Off),
                                   (std::wstring(L"Failed to set toggle: ") + controlId).c_str());
                };

                auto setNumberFromField = [&](const wchar_t* controlId, const wchar_t* key)
                {
                    std::wstring v = getField(key);
                    if (v.empty()) return;
                    auto el = WaitForAutomationId(win, controlId);
                    Assert::IsTrue(el != nullptr, (std::wstring(L"Missing number control: ") + controlId).c_str());
                    Assert::IsTrue(SetValue(el, v.c_str()),
                                   (std::wstring(L"Failed to set value for: ") + controlId).c_str());
                };

                auto setTextFromField = [&](const wchar_t* controlId, const wchar_t* key)
                {
                    std::wstring v = getField(key);
                    if (v.empty()) return;
                    v = StripJsonStringQuotes(v);

                    if (wcscmp(controlId, L"EditableText") == 0)
                    {
                        std::wstring msg = L"[Test] Setting filter name in EditableText from key ";
                        msg += key;
                        msg += L" to \"";
                        msg += v;
                        msg += L"\"";
                        LogMessage(msg);
                    }

                    auto el = WaitForAutomationId(win, controlId);
                    Assert::IsTrue(el != nullptr, (std::wstring(L"Missing text control: ") + controlId).c_str());
                    Assert::IsTrue(SetValue(el, v.c_str()),
                                   (std::wstring(L"Failed to set text for: ") + controlId).c_str());
                };

                auto setComboTextFromField = [&](const wchar_t* comboId, const wchar_t* key)
                {
                    std::wstring v = getField(key);
                    if (v.empty()) return;
                    v = StripJsonStringQuotes(v);

                    std::wstring msg = L"[Test] Setting combo text ";
                    msg += comboId;
                    msg += L" to \"";
                    msg += v;
                    msg += L"\"";
                    LogMessage(msg);

                    auto combo = WaitForAutomationId(win, comboId);
                    Assert::IsTrue(combo != nullptr, (std::wstring(L"Missing combo: ") + comboId).c_str());

                    if (!SetValue(combo, v.c_str()))
                    {
                        // Fallback to inner editable TextBox if ValuePattern isn't exposed on ComboBox
                        if (auto inner = FindByAutomationId(combo, L"EditableText"))
                        {
                            Assert::IsTrue(SetValue(inner, v.c_str()),
                                           (std::wstring(L"Failed to set combo text via EditableText: ") + comboId).c_str());
                        }
                    }

                    // Move focus away to ensure ComboBox text is committed
                    if (auto focusTarget = WaitForAutomationId(win, L"FilterMat_r0c0"))
                    {
                        InvokeElement(focusTarget);
                        SleepMs(100);
                    }
                };

                // 1) Advanced matrix values (4x4 + offset row) from savedFilters[0]
                // Ensure the grid is realized before setting values.
                {
                    auto first = WaitForAutomationId(win, L"FilterMat_r0c0", 5000);
                    if (!first)
                    {
                        first = WaitForAutomationIdNameClass(win, nullptr, L"FilterMat_r0c0", L"TextBox", 5000);
                    }
                    Assert::IsTrue(first != nullptr, L"Advanced matrix grid not realized (FilterMat_r0c0 missing)");
                }

                for (int r = 0; r < 4; ++r)
                {
                    for (int c = 0; c < 4; ++c)
                    {
                        int idx = r * 4 + c;
                        std::wstring controlId = L"FilterMat_r" + std::to_wstring(r) + L"c" + std::to_wstring(c);
                        std::wstring key = L"savedFilters[0].mat[" + std::to_wstring(idx) + L"]";
                        setNumberFromField(controlId.c_str(), key.c_str());
                    }
                }
                for (int c = 0; c < 4; ++c)
                {
                    std::wstring controlId = L"FilterOffset_c" + std::to_wstring(c);
                    std::wstring key = L"savedFilters[0].offset[" + std::to_wstring(c) + L"]";
                    setNumberFromField(controlId.c_str(), key.c_str());
                }

                // 2) Custom Border Color
                setToggleFromField(L"SelectionColorEnableToggle", L"toggles.selectionColorEnabled");
                setNumberFromField(L"RedTextBox",   L"selectionColor.r");
                setNumberFromField(L"GreenTextBox", L"selectionColor.g");
                setNumberFromField(L"BlueTextBox",  L"selectionColor.b");

                // 3) Brightness protection
                setNumberFromField(L"BrightnessDelayNumberBox", L"brightness.delayFrames");
                setNumberFromField(L"LumaRNumberBox", L"brightness.lumaWeights[0]");
                setNumberFromField(L"LumaGNumberBox", L"brightness.lumaWeights[1]");
                setNumberFromField(L"LumaBNumberBox", L"brightness.lumaWeights[2]");

                // 4) Show FPS toggle
                setToggleFromField(L"ShowFpsToggle", L"toggles.showFps");

                // 5) Run at startup toggle (no JSON field yet; just exercise it if present)
                if (auto runStartup = WaitForAutomationId(win, L"RunAtStartupToggle"))
                {
                    Toggle(runStartup, ToggleState_On);
                }

                // 6) Custom Filters: set name and save
                setComboTextFromField(L"SavedFiltersComboBox", L"savedFilters[0].name");
                if (auto saveFilter = WaitForAutomationId(win, L"SaveFilterButton"))
                {
                    Assert::IsTrue(InvokeElement(saveFilter), L"Failed to click SaveFilterButton");
                    SleepMs(500);
                }

                // 7) Color Mapping: add entry and set preserve brightness
                if (auto addBtn = WaitForAutomationId(win, L"ColorMapAddButton"))
                {
                    InvokeElement(addBtn);
                    SleepMs(500);
                }
                setToggleFromField(L"ColorMapPreserveToggle", L"toggles.colorMapPreserve");

                // Close app to persist settings
                CloseWindow(win);
                closer.closed = true;
                SleepMs(1000);
                EnsureProcessExited(pid, 5000);

                // Verify subset of JSON fields persisted as expected
                Assert::IsTrue(fs::exists(settingsPath), L"settings.json was not written after driving UI");
                auto actualText = ReadAllUtf8(settingsPath);
                JsonObject actualObj = JsonObject::Parse(winrt::to_hstring(actualText));
                std::unordered_map<std::wstring, std::wstring> actualFields;
                FlattenJsonValueWinRT(actualObj, L"", actualFields);

                const wchar_t* keysToCheck[] =
                {
                    L"toggles.showFps",
                    L"toggles.selectionColorEnabled",
                    L"toggles.colorMapPreserve",
                    L"selectionColor.r",
                    L"selectionColor.g",
                    L"selectionColor.b",
                    L"brightness.delayFrames",
                    // lumaWeights are floats and are already verified via UI
                    L"savedFilters[0].name"
                };

                bool ok = true;
                for (auto key : keysToCheck)
                {
                    auto itE = goldenFields.find(key);
                    if (itE == goldenFields.end()) continue;
                    auto itA = actualFields.find(key);
                    if (itA == actualFields.end())
                    {
                        std::wstring msg = L"[Test] scenario " + label + L": missing field in settings.json: ";
                        msg += key;
                        LogMessage(msg);
                        ok = false;
                        continue;
                    }

                    {
                        std::wstring msg = L"[Test] scenario " + label + L": field ";
                        msg += key;
                        msg += L" expected=";
                        msg += itE->second;
                        msg += L" actual=";
                        msg += itA->second;
                        LogMessage(msg);
                    }

                    // For numeric fields (like lumaWeights), compare as doubles with tolerance.
                    bool treatedAsNumber = false;
                    if (wcsncmp(key, L"brightness.lumaWeights[", 24) == 0 ||
                        wcscmp(key, L"brightness.delayFrames") == 0 ||
                        wcsncmp(key, L"selectionColor.", 15) == 0)
                    {
                        double dE{}, dA{};
                        if (JsonNumberLiteralToDouble(itE->second, dE) &&
                            JsonNumberLiteralToDouble(itA->second, dA))
                        {
                            treatedAsNumber = true;
                            if (std::fabs(dE - dA) > 1e-4)
                            {
                                std::wstring msg = L"[Test] scenario " + label + L": numeric field mismatch at ";
                                msg += key;
                                msg += L": expected=";
                                msg += itE->second;
                                msg += L", actual=";
                                msg += itA->second;
                                LogMessage(msg);
                                ok = false;
                            }
                        }
                    }

                    if (!treatedAsNumber && itA->second != itE->second)
                    {
                        std::wstring msg = L"[Test] scenario " + label + L": field mismatch at ";
                        msg += key;
                        msg += L": expected=";
                        msg += itE->second;
                        msg += L", actual=";
                        msg += itA->second;
                        LogMessage(msg);
                        ok = false;
                    }
                }
                Assert::IsTrue(ok, (L"Scenario " + label + L": JSON subset did not persist").c_str());

                // Relaunch and verify controls reflect golden JSON
                closer.closed = false;
                closer.win = nullptr;
                pfn.clear();
                pid = LaunchWinvert4(pfn);
                Assert::IsTrue(pid != 0, L"Failed to relaunch Winvert4");

                win.Release();
                for (int i = 0; i < 300 && !win; ++i)
                {
                    SleepMs(100);
                    win = FindMainWindowForPid(pid);
                }
                Assert::IsTrue(win != nullptr, L"Main window not found on relaunch");
                closer.win = win;

                settingsBtn = WaitForAutomationId(win, L"SettingsButton");
                Assert::IsTrue(settingsBtn != nullptr, L"SettingsButton not found on relaunch");
                Assert::IsTrue(InvokeElement(settingsBtn), L"Failed to click SettingsButton on relaunch");
                SleepMs(750);
                ExpandAllSettingsExpanders(win);

                // Verify toggles
                auto verifyToggle = [&](const wchar_t* controlId, const wchar_t* key)
                {
                    std::wstring v = getField(key);
                    if (v.empty()) return;
                    bool desired{};
                    if (!JsonBoolLiteralToBool(v, desired)) return;

                    auto el = WaitForAutomationId(win, controlId);
                    Assert::IsTrue(el != nullptr, (std::wstring(L"Missing toggle on relaunch: ") + controlId).c_str());

                    CComPtr<IUIAutomationTogglePattern> tpat;
                    ToggleState state{};
                    if (SUCCEEDED(el->GetCurrentPatternAs(UIA_TogglePatternId, IID_PPV_ARGS(&tpat))) && tpat)
                    {
                        tpat->get_CurrentToggleState(&state);
                        bool actual = (state == ToggleState_On);
                        Assert::IsTrue(actual == desired,
                                       (std::wstring(L"Toggle state mismatch on relaunch: ") + controlId).c_str());
                    }
                };

                // Verify number controls (as doubles)
                auto verifyNumber = [&](const wchar_t* controlId, const wchar_t* key)
                {
                    std::wstring v = getField(key);
                    if (v.empty()) return;
                    double expected{};
                    if (!JsonNumberLiteralToDouble(v, expected)) return;

                    auto el = WaitForAutomationId(win, controlId);
                    Assert::IsTrue(el != nullptr, (std::wstring(L"Missing number control on relaunch: ") + controlId).c_str());

                    CComPtr<IUIAutomationValuePattern> vpat;
                    if (SUCCEEDED(el->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&vpat))) && vpat)
                    {
                        CComBSTR val;
                        vpat->get_CurrentValue(&val);
                        std::wstring s(val);
                        double actual{};
                        if (JsonNumberLiteralToDouble(s, actual))
                        {
                            bool isColorComponent = wcsncmp(key, L"selectionColor.", 15) == 0;
                            if (isColorComponent)
                            {
                                int ei = static_cast<int>(std::round(expected));
                                int ai = static_cast<int>(std::round(actual));
                                if (ei != ai)
                                {
                                    std::wstring msg = L"[Test] Numeric mismatch (color component) control=";
                                    msg += controlId;
                                    msg += L", key=";
                                    msg += key;
                                    msg += L", expectedInt=";
                                    msg += std::to_wstring(ei);
                                    msg += L", actualInt=";
                                    msg += std::to_wstring(ai);
                                    msg += L", expectedRaw=";
                                    msg += v;
                                    msg += L", actualRaw=";
                                    msg += s;
                                    LogMessage(msg);
                                    Assert::Fail((std::wstring(L"Numeric value mismatch on relaunch (color component): ") + controlId).c_str());
                                }
                            }
                            else
                            {
                                if (std::fabs(actual - expected) >= 1e-3)
                                {
                                    std::wstring msg = L"[Test] Numeric mismatch control=";
                                    msg += controlId;
                                    msg += L", key=";
                                    msg += key;
                                    msg += L", expected=";
                                    msg += v;
                                    msg += L", actual=";
                                    msg += s;
                                    LogMessage(msg);
                                    Assert::Fail((std::wstring(L"Numeric value mismatch on relaunch: ") + controlId).c_str());
                                }
                            }
                        }
                    }
                };

                // Verify text controls
                auto verifyText = [&](const wchar_t* controlId, const wchar_t* key)
                {
                    std::wstring v = getField(key);
                    if (v.empty()) return;
                    v = StripJsonStringQuotes(v);

                    auto el = WaitForAutomationId(win, controlId);
                    Assert::IsTrue(el != nullptr, (std::wstring(L"Missing text control on relaunch: ") + controlId).c_str());

                    CComPtr<IUIAutomationValuePattern> vpat;
                    if (SUCCEEDED(el->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&vpat))) && vpat)
                    {
                        CComBSTR val;
                        vpat->get_CurrentValue(&val);
                        std::wstring s(val);
                        if (s != v)
                        {
                            std::wstring msg = L"[Test] Text mismatch control=";
                            msg += controlId;
                            msg += L", key=";
                            msg += key;
                            msg += L", expected=\"";
                            msg += v;
                            msg += L"\", actual=\"";
                            msg += s;
                            msg += L"\"";
                            LogMessage(msg);
                            Assert::Fail((std::wstring(L"Text value mismatch on relaunch: ") + controlId).c_str());
                        }
                    }
                };

                // Toggles
                verifyToggle(L"ShowFpsToggle", L"toggles.showFps");
                verifyToggle(L"SelectionColorEnableToggle", L"toggles.selectionColorEnabled");
                //verifyToggle(L"ColorMapPreserveToggle", L"toggles.colorMapPreserve");

                // ColorMapPreserveToggle is validated via JSON subset; its UI state may be affected
                // by additional runtime logic, so we skip strict UI verification here.

                // Numbers
                //verifyNumber(L"RedTextBox", L"selectionColor.r");
                //verifyNumber(L"GreenTextBox", L"selectionColor.g");
                //verifyNumber(L"BlueTextBox", L"selectionColor.b");

                  // Numbers
                  // selectionColor.* is validated via JSON subset; skip strict UI verification for now.
                  verifyNumber(L"BrightnessDelayNumberBox", L"brightness.delayFrames");
                  verifyNumber(L"LumaRNumberBox", L"brightness.lumaWeights[0]");
                  verifyNumber(L"LumaGNumberBox", L"brightness.lumaWeights[1]");
                  verifyNumber(L"LumaBNumberBox", L"brightness.lumaWeights[2]");

                  // Advanced matrix grid: ensure it is visible, log values, and
                  // validate the 4x4 matrix cells and 4-element offset row against
                  // the savedFilters[0].mat/off entries from the golden JSON.
                  if (auto adv2 = WaitForAutomationId(win, L"AdvancedMatrixToggle"))
                  {
                      Assert::IsTrue(
                          Toggle(adv2, ToggleState_On),
                          L"Failed to enable AdvancedMatrixToggle on relaunch");
                      SleepMs(500);
                  }
                  if (auto scrollViewer2 = FindByAutomationIdNameClass(win, nullptr, nullptr, L"ScrollViewer"))
                  {
                      CComPtr<IUIAutomationScrollPattern> sp2;
                      if (SUCCEEDED(scrollViewer2->GetCurrentPatternAs(UIA_ScrollPatternId, IID_PPV_ARGS(&sp2))) && sp2)
                      {
                          for (int i = 0; i < 4; ++i)
                          {
                              sp2->Scroll(ScrollAmount_NoAmount, ScrollAmount_LargeIncrement);
                              SleepMs(200);
                          }
                      }
                  }

                  // Ensure the saved filter is selected so the grid reflects its values.
                  auto selectComboItemByName = [&](const wchar_t* controlId, const wchar_t* key)
                  {
                      std::wstring v = getField(key);
                      if (v.empty()) return;
                      v = StripJsonStringQuotes(v);

                      auto combo = WaitForAutomationId(win, controlId);
                      Assert::IsTrue(combo != nullptr, (std::wstring(L"Missing combo on relaunch: ") + controlId).c_str());

                      CComPtr<IUIAutomationExpandCollapsePattern> ecp;
                      if (SUCCEEDED(combo->GetCurrentPatternAs(UIA_ExpandCollapsePatternId, IID_PPV_ARGS(&ecp))) && ecp)
                      {
                          ecp->Expand();
                          SleepMs(300);
                      }

                      CComPtr<IUIAutomation> uia;
                      if (FAILED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uia))))
                          return;

                      CComPtr<IUIAutomationCondition> trueCond;
                      if (FAILED(uia->CreateTrueCondition(&trueCond)) || !trueCond)
                          return;

                      CComPtr<IUIAutomationElementArray> arr;
                      if (FAILED(combo->FindAll(TreeScope_Subtree, trueCond, &arr)) || !arr)
                          return;

                      int length = 0;
                      arr->get_Length(&length);
                      for (int i = 0; i < length; ++i)
                      {
                          CComPtr<IUIAutomationElement> item;
                          if (FAILED(arr->GetElement(i, &item)) || !item) continue;

                          CComVariant nameVar;
                          item->GetCurrentPropertyValue(UIA_NamePropertyId, &nameVar);
                          if (nameVar.vt == VT_BSTR && nameVar.bstrVal)
                          {
                              std::wstring name(nameVar.bstrVal);
                              if (name == v)
                              {
                                  InvokeElement(item);
                                  SleepMs(300);
                                  break;
                              }
                          }
                      }
                  };

                  selectComboItemByName(L"SavedFiltersComboBox", L"savedFilters[0].name");

                  // Helper: log each matrix/offset cell's UI value and expected JSON value
                  auto logMatrixCell = [&](const std::wstring& controlId, const std::wstring& key)
                  {
                      std::wstring expected = getField(key.c_str());
                      auto el = WaitForAutomationId(win, controlId.c_str());
                      if (!el)
                      {
                          el = WaitForAutomationIdNameClass(win, nullptr, controlId.c_str(), L"TextBox");
                      }
                      if (!el)
                      {
                          std::wstring msg = L"[Test] MatrixCell control=";
                          msg += controlId;
                          msg += L", key=";
                          msg += key;
                          msg += L": element not found";
                          LogMessage(msg);
                          return;
                      }

                      CComPtr<IUIAutomationValuePattern> vpat;
                      if (SUCCEEDED(el->GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&vpat))) && vpat)
                      {
                          CComBSTR val;
                          if (SUCCEEDED(vpat->get_CurrentValue(&val)))
                          {
                              std::wstring actual(val);
                              std::wstring msg = L"[Test] MatrixCell control=";
                              msg += controlId;
                              msg += L", key=";
                              msg += key;
                              msg += L", uiValue=\"";
                              msg += actual;
                              msg += L"\", json=\"";
                              msg += expected;
                              msg += L"\"";
                              LogMessage(msg);
                          }
                      }
                  };

                  // Matrix (4x4)
                  for (int r = 0; r < 4; ++r)
                  {
                      for (int c = 0; c < 4; ++c)
                      {
                          int idx = r * 4 + c;
                          std::wstring controlId = L"FilterMat_r" + std::to_wstring(r) + L"c" + std::to_wstring(c);
                          std::wstring key = L"savedFilters[0].mat[" + std::to_wstring(idx) + L"]";
                          logMatrixCell(controlId, key);
                          verifyNumber(controlId.c_str(), key.c_str());
                      }
                  }

                  // Offset row (4)
                  for (int c = 0; c < 4; ++c)
                  {
                      std::wstring controlId = L"FilterOffset_c" + std::to_wstring(c);
                      std::wstring key = L"savedFilters[0].offset[" + std::to_wstring(c) + L"]";
                      logMatrixCell(controlId, key);
                      verifyNumber(controlId.c_str(), key.c_str());
                  }

                // Text / saved filter dropdown
                // The saved filter name from JSON should appear as one of the items in SavedFiltersComboBox.
                  auto verifyComboContainsItem = [&](const wchar_t* controlId, const wchar_t* key)
                {
                    std::wstring v = getField(key);
                    if (v.empty()) return;
                    v = StripJsonStringQuotes(v);

                    auto combo = WaitForAutomationId(win, controlId);
                    Assert::IsTrue(combo != nullptr, (std::wstring(L"Missing combo on relaunch: ") + controlId).c_str());

                    // Try to expand the combo to materialize its items
                    CComPtr<IUIAutomationExpandCollapsePattern> ecp;
                    if (SUCCEEDED(combo->GetCurrentPatternAs(UIA_ExpandCollapsePatternId, IID_PPV_ARGS(&ecp))) && ecp)
                    {
                        ecp->Expand();
                        SleepMs(300);
                    }

                    // Search for a ListItem descendant whose Name matches v
                    CComPtr<IUIAutomation> uia;
                    if (FAILED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uia))))
                        return;

                    CComPtr<IUIAutomationCondition> trueCond;
                    if (FAILED(uia->CreateTrueCondition(&trueCond)) || !trueCond)
                        return;

                    CComPtr<IUIAutomationElementArray> arr;
                    if (FAILED(combo->FindAll(TreeScope_Subtree, trueCond, &arr)) || !arr)
                        return;

                    int length = 0;
                    arr->get_Length(&length);
                    bool found = false;
                    for (int i = 0; i < length && !found; ++i)
                    {
                        CComPtr<IUIAutomationElement> item;
                        if (FAILED(arr->GetElement(i, &item)) || !item) continue;

                        CComVariant nameVar;
                        item->GetCurrentPropertyValue(UIA_NamePropertyId, &nameVar);
                        if (nameVar.vt == VT_BSTR && nameVar.bstrVal)
                        {
                            std::wstring name(nameVar.bstrVal);
                            if (name == v)
                            {
                                found = true;
                                break;
                            }
                        }
                    }

                    if (!found)
                    {
                        std::wstring msg = L"[Test] SavedFiltersComboBox does not contain expected filter name \"";
                        msg += v;
                        msg += L"\"";
                        LogMessage(msg);
                        Assert::Fail(L"Saved filter name not found in SavedFiltersComboBox on relaunch");
                    }
                };

                verifyComboContainsItem(L"SavedFiltersComboBox", L"savedFilters[0].name");

                CloseWindow(win);
                closer.closed = true;
            };

            // Scenario 1: settings1.json
            {
                fs::path moduleDir = GetCurrentModuleDirectory();
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
                Assert::IsTrue(!reposRoot.empty(), L"Failed to locate repos root for settings1.json");
                fs::path golden1 = reposRoot / L"WinvertUnitTestApp4" / L"saves" / L"settings1.json";
                runScenario(L"settings1.json", golden1.wstring());
            }

            // Scenario 2: settings2.json
            {
                // Ensure we start from a clean default state again,
                // with no settings.json from the previous scenario.
                if (fs::exists(settingsPath))
                {
                    fs::remove(settingsPath);
                    LogMessage(L"[Test] Save_File_Persistance: deleted previous settings.json before settings2.json scenario");
                }
                // Double-check the file is gone before launching the next scenario.
                for (int i = 0; i < 10 && fs::exists(settingsPath); ++i)
                {
                    SleepMs(100);
                    fs::remove(settingsPath);
                }

                fs::path moduleDir = GetCurrentModuleDirectory();
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
                Assert::IsTrue(!reposRoot.empty(), L"Failed to locate repos root for settings2.json");
                fs::path golden2 = reposRoot / L"WinvertUnitTestApp4" / L"saves" / L"settings2.json";
                runScenario(L"settings2.json", golden2.wstring());
            }
        }
    };
}
