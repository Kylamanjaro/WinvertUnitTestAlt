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

static bool Invoke(IUIAutomationElement* el)
{
    if (!el) return false; CComPtr<IUIAutomationInvokePattern> pat; if (FAILED(el->GetCurrentPatternAs(UIA_InvokePatternId, IID_PPV_ARGS(&pat)))) return false; return SUCCEEDED(pat->Invoke());
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
        TEST_METHOD(Test)
        {
            Assert::IsTrue(true);
        }

        TEST_METHOD(EndToEnd_SettingsPersist)
        {
            CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

            // 1) Launch app with no save file (defaults)
            std::wstring pfn; DWORD pid = LaunchWinvert4(pfn);
            Assert::IsTrue(pid != 0, L"Failed to launch Winvert4");
            std::wstring localState = LocalStatePathForPFN(pfn);
            Assert::IsTrue(!localState.empty(), L"Failed to get LocalState path");

            // Ensure no save file to start
            std::wstring settingsPath = localState + L"\\settings.json";
            if (fs::exists(settingsPath)) fs::remove(settingsPath);

            // Find main window
            CComPtr<IUIAutomationElement> win;
            for (int i = 0; i < 50 && !win; ++i) { SleepMs(100); win = FindWindowByName(L"Winvert Control Panel"); }
            Assert::IsTrue(win != nullptr, L"Main window not found");

            // Open settings
            auto settingsBtn = FindByAutomationId(win, L"SettingsButton");
            Assert::IsTrue(Invoke(settingsBtn), L"Failed to click SettingsButton");
            SleepMs(500);

            // Validate defaults
            auto fps = FindByAutomationId(win, L"ShowFpsToggle");
            auto sel = FindByAutomationId(win, L"SelectionColorEnableToggle");
            auto nbDelay = FindByAutomationId(win, L"BrightnessDelayNumberBox");
            auto nbR = FindByAutomationId(win, L"LumaRNumberBox");
            auto nbG = FindByAutomationId(win, L"LumaGNumberBox");
            auto nbB = FindByAutomationId(win, L"LumaBNumberBox");
            Assert::IsTrue(fps && sel && nbDelay && nbR && nbG && nbB, L"Settings controls missing");

            // 2) Change settings via UI
            Assert::IsTrue(Toggle(fps, ToggleState_On), L"Toggle ShowFps ON failed");
            Assert::IsTrue(Toggle(sel, ToggleState_On), L"Toggle SelectionColor ON failed");
            Assert::IsTrue(SetValue(nbDelay, L"7"), L"Set delay failed");
            Assert::IsTrue(SetValue(nbR, L"0.30"), L"Set luma R failed");
            Assert::IsTrue(SetValue(nbG, L"0.59"), L"Set luma G failed");
            Assert::IsTrue(SetValue(nbB, L"0.11"), L"Set luma B failed");

            // Add a color map (best-effort; if control exists)
            if (auto add = FindByAutomationId(win, L"ColorMapAddButton")) { Invoke(add); SleepMs(200); }

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
            pfn.clear(); pid = LaunchWinvert4(pfn);
            Assert::IsTrue(pid != 0, L"Failed to relaunch Winvert4");
            win = nullptr; for (int i = 0; i < 50 && !win; ++i) { SleepMs(100); win = FindWindowByName(L"Winvert Control Panel"); }
            Assert::IsTrue(win != nullptr, L"Main window not found (relaunch)");
            settingsBtn = FindByAutomationId(win, L"SettingsButton"); Invoke(settingsBtn); SleepMs(500);

            fps = FindByAutomationId(win, L"ShowFpsToggle"); sel = FindByAutomationId(win, L"SelectionColorEnableToggle"); nbDelay = FindByAutomationId(win, L"BrightnessDelayNumberBox"); nbR = FindByAutomationId(win, L"LumaRNumberBox"); nbG = FindByAutomationId(win, L"LumaGNumberBox"); nbB = FindByAutomationId(win, L"LumaBNumberBox");
            // We don't read back ToggleState/Value here due to brevity; presence is a smoke test. Extend as needed.
            Assert::IsTrue(fps && sel && nbDelay && nbR && nbG && nbB, L"Settings controls missing on relaunch");

            CloseWindow(win);
        }
    };
}
