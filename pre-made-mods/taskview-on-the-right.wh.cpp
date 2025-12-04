// ==WindhawkMod==
// @id              taskbar-taskview-button-position-right
// @name            Task View on the right
// @description     Moves the Task View button to the right of the open/pinned apps
// @version         1.1
// @author          Gemini 3.0, Lockframe
// @github          https://github.com/Lockframe
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -ldwmapi -lole32 -loleaut32 -lruntimeobject -lshcore
// ==/WindhawkMod==

// ==WindhawkModSettings==
/*
- customMargin: -48,0
  $name: Custom Margin (Left,Top,Right,Bottom)
  $description: Sets the margin for the Task View button. Negative margins help collapse the space in the original layout. Default = -48,0

- customWidth: 48
  $name: Button Width
  $description: width in pixels.

- customHeight: 48
  $name: Button Height
  $description: Height in pixels (set to 0 to use default taskbar height).

- customPadding: 4
  $name: Button Padding (Left,Top,Right,Bottom)
  $description: Internal padding for the icon. Default = 4
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <atomic>
#include <functional>
#include <cwctype>
#include <string>
#include <sstream>

#include <dwmapi.h>
#include <roapi.h>
#include <winstring.h>

#undef GetCurrentTime

#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.UI.Core.h>
#include <winrt/Windows.UI.Xaml.Automation.h>
#include <winrt/Windows.UI.Xaml.Controls.h>
#include <winrt/Windows.UI.Xaml.Media.h>
#include <winrt/Windows.UI.Xaml.Shapes.h>
#include <winrt/Windows.UI.Xaml.h>
#include <winrt/base.h>

using namespace winrt::Windows::UI::Xaml;

struct ModSettings {
    Thickness margin;
    double width;
    double height;
    Thickness padding;
} g_settings;

std::atomic<bool> g_taskbarViewDllLoaded;
std::atomic<bool> g_unloading;

thread_local bool g_TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride;

void* CTaskBand_ITaskListWndSite_vftable;
void* CSecondaryTaskBand_ITaskListWndSite_vftable;

using CTaskBand_GetTaskbarHost_t = void*(WINAPI*)(void* pThis, void** result);
CTaskBand_GetTaskbarHost_t CTaskBand_GetTaskbarHost_Original;

void* TaskbarHost_FrameHeight_Original;

using CSecondaryTaskBand_GetTaskbarHost_t = void*(WINAPI*)(void* pThis, void** result);
CSecondaryTaskBand_GetTaskbarHost_t CSecondaryTaskBand_GetTaskbarHost_Original;

using std__Ref_count_base__Decref_t = void(WINAPI*)(void* pThis);
std__Ref_count_base__Decref_t std__Ref_count_base__Decref_Original;

// --- Helper Functions ---

Thickness ParseThickness(PCWSTR str, Thickness defaultVal) {
    if (!str || !*str) return defaultVal;
    
    double values[4] = {0};
    int count = 0;
    
    std::wstring s(str);
    std::wstringstream ss(s);
    std::wstring item;
    
    while (std::getline(ss, item, L',')) {
        if (count >= 4) break;
        try {
            values[count++] = std::stod(item);
        } catch (...) {}
    }

    if (count == 1) return {values[0], values[0], values[0], values[0]};
    if (count == 4) return {values[0], values[1], values[2], values[3]};
    
    return defaultVal;
}

void LoadSettings() {
    // Parse Margin
    PCWSTR marginStr = Wh_GetStringSetting(L"customMargin");
    g_settings.margin = ParseThickness(marginStr, {-48, 0, -48, 0});
    Wh_FreeStringSetting(marginStr);

    // Parse Padding
    PCWSTR paddingStr = Wh_GetStringSetting(L"customPadding");
    g_settings.padding = ParseThickness(paddingStr, {12, 4, 2, 4});
    Wh_FreeStringSetting(paddingStr);

    // Parse Dimensions
    g_settings.width = (double)Wh_GetIntSetting(L"customWidth");
    g_settings.height = (double)Wh_GetIntSetting(L"customHeight");

    if (g_settings.width <= 0) g_settings.width = 48.0;
}

HWND FindCurrentProcessTaskbarWnd() {
    HWND hTaskbarWnd = nullptr;
    EnumWindows([](HWND hWnd, LPARAM lParam) -> BOOL {
        DWORD dwProcessId;
        WCHAR className[32];
        if (GetWindowThreadProcessId(hWnd, &dwProcessId) &&
            dwProcessId == GetCurrentProcessId() &&
            GetClassName(hWnd, className, ARRAYSIZE(className)) &&
            _wcsicmp(className, L"Shell_TrayWnd") == 0) {
            *reinterpret_cast<HWND*>(lParam) = hWnd;
            return FALSE;
        }
        return TRUE;
    }, reinterpret_cast<LPARAM>(&hTaskbarWnd));
    return hTaskbarWnd;
}

FrameworkElement EnumChildElements(FrameworkElement element, std::function<bool(FrameworkElement)> enumCallback) {
    int childrenCount = Media::VisualTreeHelper::GetChildrenCount(element);
    for (int i = 0; i < childrenCount; i++) {
        auto child = Media::VisualTreeHelper::GetChild(element, i).try_as<FrameworkElement>();
        if (!child) continue;
        if (enumCallback(child)) return child;
    }
    return nullptr;
}

FrameworkElement FindChildByName(FrameworkElement element, PCWSTR name) {
    return EnumChildElements(element, [name](FrameworkElement child) {
        return child.Name() == name;
    });
}

FrameworkElement FindChildByClassName(FrameworkElement element, PCWSTR className) {
    return EnumChildElements(element, [className](FrameworkElement child) {
        return winrt::get_class_name(child) == className;
    });
}

// --- Layout Logic ---

bool ApplyStyle(XamlRoot xamlRoot) {
    FrameworkElement xamlRootContent = xamlRoot.Content().try_as<FrameworkElement>();
    FrameworkElement taskbarFrameRepeater = nullptr;

    FrameworkElement child = xamlRootContent;
    if (child &&
        (child = FindChildByClassName(child, L"Taskbar.TaskbarFrame")) &&
        (child = FindChildByName(child, L"RootGrid")) &&
        (child = FindChildByName(child, L"TaskbarFrameRepeater"))) {
        taskbarFrameRepeater = child;
    }

    if (!taskbarFrameRepeater) return false;

    auto startButton = EnumChildElements(taskbarFrameRepeater, [](FrameworkElement child) {
        auto childClassName = winrt::get_class_name(child);
        if (childClassName != L"Taskbar.ExperienceToggleButton") return false;
        auto automationId = Automation::AutomationProperties::GetAutomationId(child);
        return automationId == L"TaskViewButton";
    });

    if (startButton) {
        // Enforce dimensions
        if (!g_unloading) {
            if (startButton.Width() != g_settings.width) startButton.Width(g_settings.width);
            if (startButton.MinWidth() != g_settings.width) startButton.MinWidth(g_settings.width);
        }

        Thickness currentMargin = startButton.Margin();
        Thickness targetMargin;

        if (g_unloading) {
            targetMargin = {0, 0, 0, 0};
        } else {
            targetMargin = g_settings.margin;
        }

        // Apply if changed
        if (currentMargin.Left != targetMargin.Left || 
            currentMargin.Top != targetMargin.Top || 
            currentMargin.Right != targetMargin.Right || 
            currentMargin.Bottom != targetMargin.Bottom) {
            startButton.Margin(targetMargin);
        }
    }
    return true;
}

XamlRoot XamlRootFromTaskbarHostSharedPtr(void* taskbarHostSharedPtr[2]) {
    if (!taskbarHostSharedPtr[0] && !taskbarHostSharedPtr[1]) return nullptr;
    size_t taskbarElementIUnknownOffset = 0x48;

#if defined(_M_X64)
    const BYTE* b = (const BYTE*)TaskbarHost_FrameHeight_Original;
    if (b[0] == 0x48 && b[1] == 0x83 && b[2] == 0xEC && b[4] == 0x48 &&
        b[5] == 0x83 && b[6] == 0xC1 && b[7] <= 0x7F) {
        taskbarElementIUnknownOffset = b[7];
    }
#endif

    auto* taskbarElementIUnknown = *(IUnknown**)((BYTE*)taskbarHostSharedPtr[0] + taskbarElementIUnknownOffset);
    FrameworkElement taskbarElement = nullptr;
    taskbarElementIUnknown->QueryInterface(winrt::guid_of<FrameworkElement>(), winrt::put_abi(taskbarElement));
    auto result = taskbarElement ? taskbarElement.XamlRoot() : nullptr;
    std__Ref_count_base__Decref_Original(taskbarHostSharedPtr[1]);
    return result;
}

XamlRoot GetTaskbarXamlRoot(HWND hTaskbarWnd) {
    HWND hTaskSwWnd = (HWND)GetProp(hTaskbarWnd, L"TaskbandHWND");
    if (!hTaskSwWnd) return nullptr;

    void* taskBand = (void*)GetWindowLongPtr(hTaskSwWnd, 0);
    void* taskBandForTaskListWndSite = taskBand;
    for (int i = 0; *(void**)taskBandForTaskListWndSite != CTaskBand_ITaskListWndSite_vftable; i++) {
        if (i == 20) return nullptr;
        taskBandForTaskListWndSite = (void**)taskBandForTaskListWndSite + 1;
    }
    void* taskbarHostSharedPtr[2]{};
    CTaskBand_GetTaskbarHost_Original(taskBandForTaskListWndSite, taskbarHostSharedPtr);
    return XamlRootFromTaskbarHostSharedPtr(taskbarHostSharedPtr);
}

XamlRoot GetSecondaryTaskbarXamlRoot(HWND hSecondaryTaskbarWnd) {
    HWND hTaskSwWnd = (HWND)FindWindowEx(hSecondaryTaskbarWnd, nullptr, L"WorkerW", nullptr);
    if (!hTaskSwWnd) return nullptr;

    void* taskBand = (void*)GetWindowLongPtr(hTaskSwWnd, 0);
    void* taskBandForTaskListWndSite = taskBand;
    for (int i = 0; *(void**)taskBandForTaskListWndSite != CSecondaryTaskBand_ITaskListWndSite_vftable; i++) {
        if (i == 20) return nullptr;
        taskBandForTaskListWndSite = (void**)taskBandForTaskListWndSite + 1;
    }
    void* taskbarHostSharedPtr[2]{};
    CSecondaryTaskBand_GetTaskbarHost_Original(taskBandForTaskListWndSite, taskbarHostSharedPtr);
    return XamlRootFromTaskbarHostSharedPtr(taskbarHostSharedPtr);
}

// --- Hooks ---

using IUIElement_Arrange_t = HRESULT(WINAPI*)(void* pThis, winrt::Windows::Foundation::Rect rect);
IUIElement_Arrange_t IUIElement_Arrange_Original;

HRESULT WINAPI IUIElement_Arrange_Hook(void* pThis, winrt::Windows::Foundation::Rect rect) {
    auto original = [=] { return IUIElement_Arrange_Original(pThis, rect); };

    if (!g_TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride || g_unloading) {
        return original();
    }

    FrameworkElement element = nullptr;
    ((IUnknown*)pThis)->QueryInterface(winrt::guid_of<FrameworkElement>(), winrt::put_abi(element));
    if (!element) return original();

    auto className = winrt::get_class_name(element);
    if (className != L"Taskbar.ExperienceToggleButton") return original();

    auto automationId = Automation::AutomationProperties::GetAutomationId(element);
    if (automationId != L"TaskViewButton") return original();

    // 1. Enforce Width/Height Properties
    if (element.Width() != g_settings.width) element.Width(g_settings.width);
    if (element.MinWidth() != g_settings.width) element.MinWidth(g_settings.width);
    if (g_settings.height > 0 && element.Height() != g_settings.height) element.Height(g_settings.height);

    // 2. Enforce Layout Rect Width
    rect.Width = (float)g_settings.width;
    if (g_settings.height > 0) rect.Height = (float)g_settings.height;

    // 3. Call Original Arrange
    HRESULT hr = IUIElement_Arrange_Original(pThis, rect);

    // 4. Async Position Correction
    element.Dispatcher().TryRunAsync(winrt::Windows::UI::Core::CoreDispatcherPriority::Low, [element]() {
        try {
            auto taskbarFrameRepeater = Media::VisualTreeHelper::GetParent(element).try_as<FrameworkElement>();
            if (!taskbarFrameRepeater) return;

            double maxX = 0.0;
            int childrenCount = Media::VisualTreeHelper::GetChildrenCount(taskbarFrameRepeater);

            // Find rightmost visible app
            for (int i = 0; i < childrenCount; i++) {
                auto child = Media::VisualTreeHelper::GetChild(taskbarFrameRepeater, i).try_as<FrameworkElement>();
                if (!child || child == element) continue;
                if (child.Visibility() != Visibility::Visible) continue;

                auto offset = child.ActualOffset();
                double rightEdge = offset.x + child.ActualWidth();
                if (rightEdge > maxX) {
                    maxX = rightEdge;
                }
            }

            double currentX = element.ActualOffset().x;
            double translationNeeded = maxX - currentX;

            auto transform = element.RenderTransform().try_as<Media::TranslateTransform>();
            if (!transform) {
                transform = Media::TranslateTransform();
                element.RenderTransform(transform);
            }

            if (abs(transform.X() - translationNeeded) > 0.1) {
                transform.X(translationNeeded);
            }
        } catch (...) {}
    });

    return hr;
}

using TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride_t = HRESULT(WINAPI*)(void* pThis, void* context, winrt::Windows::Foundation::Size size, winrt::Windows::Foundation::Size* resultSize);
TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride_t TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride_Original;

HRESULT WINAPI TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride_Hook(
    void* pThis,
    void* context,
    winrt::Windows::Foundation::Size size,
    winrt::Windows::Foundation::Size* resultSize) {
    
    [[maybe_unused]] static bool hooked = [] {
        Shapes::Rectangle rectangle;
        IUIElement element = rectangle;
        void** vtable = *(void***)winrt::get_abi(element);
        auto arrange = (IUIElement_Arrange_t)vtable[92]; 
        WindhawkUtils::SetFunctionHook(arrange, IUIElement_Arrange_Hook, &IUIElement_Arrange_Original);
        Wh_ApplyHookOperations();
        return true;
    }();

    g_TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride = true;
    HRESULT ret = TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride_Original(pThis, context, size, resultSize);
    g_TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride = false;
    return ret;
}

using ExperienceToggleButton_UpdateButtonPadding_t = void(WINAPI*)(void* pThis);
ExperienceToggleButton_UpdateButtonPadding_t ExperienceToggleButton_UpdateButtonPadding_Original;
void WINAPI ExperienceToggleButton_UpdateButtonPadding_Hook(void* pThis) {
    ExperienceToggleButton_UpdateButtonPadding_Original(pThis);
    if (g_unloading) return;

    FrameworkElement toggleButtonElement = nullptr;
    ((IUnknown**)pThis)[1]->QueryInterface(winrt::guid_of<FrameworkElement>(), winrt::put_abi(toggleButtonElement));
    if (!toggleButtonElement) return;

    auto panelElement = FindChildByName(toggleButtonElement, L"ExperienceToggleButtonRootPanel").try_as<Controls::Grid>();
    if (!panelElement) return;

    auto automationId = Automation::AutomationProperties::GetAutomationId(toggleButtonElement);
    if (automationId == L"TaskViewButton") {
        if (panelElement.Width() != g_settings.width) {
            panelElement.Width(g_settings.width);
        }
        
        Thickness currentPadding = panelElement.Padding();
        if (currentPadding.Left != g_settings.padding.Left || 
            currentPadding.Top != g_settings.padding.Top ||
            currentPadding.Right != g_settings.padding.Right ||
            currentPadding.Bottom != g_settings.padding.Bottom) {
             panelElement.Padding(g_settings.padding);
        }
    }
}

// --- Threading Helper ---

using RunFromWindowThreadProc_t = void(WINAPI*)(void* parameter);
bool RunFromWindowThread(HWND hWnd, RunFromWindowThreadProc_t proc, void* procParam) {
    static const UINT runFromWindowThreadRegisteredMsg = RegisterWindowMessage(L"Windhawk_RunFromWindowThread_" WH_MOD_ID);
    struct RUN_FROM_WINDOW_THREAD_PARAM {
        RunFromWindowThreadProc_t proc;
        void* procParam;
    };

    DWORD dwThreadId = GetWindowThreadProcessId(hWnd, nullptr);
    if (dwThreadId == 0) return false;
    if (dwThreadId == GetCurrentThreadId()) {
        proc(procParam);
        return true;
    }

    HHOOK hook = SetWindowsHookEx(WH_CALLWNDPROC, [](int nCode, WPARAM wParam, LPARAM lParam) -> LRESULT {
        if (nCode == HC_ACTION) {
            const CWPSTRUCT* cwp = (const CWPSTRUCT*)lParam;
            if (cwp->message == runFromWindowThreadRegisteredMsg) {
                RUN_FROM_WINDOW_THREAD_PARAM* param = (RUN_FROM_WINDOW_THREAD_PARAM*)cwp->lParam;
                param->proc(param->procParam);
            }
        }
        return CallNextHookEx(nullptr, nCode, wParam, lParam);
    }, nullptr, dwThreadId);

    if (!hook) return false;

    RUN_FROM_WINDOW_THREAD_PARAM param{ proc, procParam };
    SendMessage(hWnd, runFromWindowThreadRegisteredMsg, 0, (LPARAM)&param);
    UnhookWindowsHookEx(hook);
    return true;
}

void ApplySettingsFromTaskbarThread() {
    EnumThreadWindows(GetCurrentThreadId(), [](HWND hWnd, LPARAM lParam) -> BOOL {
        WCHAR szClassName[32];
        if (GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName)) == 0) return TRUE;
        XamlRoot xamlRoot = nullptr;
        if (_wcsicmp(szClassName, L"Shell_TrayWnd") == 0) xamlRoot = GetTaskbarXamlRoot(hWnd);
        else if (_wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") == 0) xamlRoot = GetSecondaryTaskbarXamlRoot(hWnd);
        else return TRUE;

        if (xamlRoot) ApplyStyle(xamlRoot);
        return TRUE;
    }, 0);
}

void ApplySettings(HWND hTaskbarWnd) {
    RunFromWindowThread(hTaskbarWnd, [](void* pParam) { ApplySettingsFromTaskbarThread(); }, 0);
}

// --- Module Loading & Hooks ---

bool HookTaskbarDllSymbols() {
    HMODULE module = LoadLibraryEx(L"taskbar.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!module) return false;

    WindhawkUtils::SYMBOL_HOOK taskbarDllHooks[] = {
        {{LR"(const CTaskBand::`vftable'{for `ITaskListWndSite'})"}, &CTaskBand_ITaskListWndSite_vftable},
        {{LR"(const CSecondaryTaskBand::`vftable'{for `ITaskListWndSite'})"}, &CSecondaryTaskBand_ITaskListWndSite_vftable},
        {{LR"(public: virtual class std::shared_ptr<class TaskbarHost> __cdecl CTaskBand::GetTaskbarHost(void)const )"}, &CTaskBand_GetTaskbarHost_Original},
        {{LR"(public: int __cdecl TaskbarHost::FrameHeight(void)const )"}, &TaskbarHost_FrameHeight_Original},
        {{LR"(public: virtual class std::shared_ptr<class TaskbarHost> __cdecl CSecondaryTaskBand::GetTaskbarHost(void)const )"}, &CSecondaryTaskBand_GetTaskbarHost_Original},
        {{LR"(public: void __cdecl std::_Ref_count_base::_Decref(void))"}, &std__Ref_count_base__Decref_Original},
    };
    return HookSymbols(module, taskbarDllHooks, ARRAYSIZE(taskbarDllHooks));
}

bool HookTaskbarViewDllSymbols(HMODULE module) {
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(public: virtual int __cdecl winrt::impl::produce<struct winrt::Taskbar::implementation::TaskbarCollapsibleLayout,struct winrt::Microsoft::UI::Xaml::Controls::IVirtualizingLayoutOverrides>::ArrangeOverride(void *,struct winrt::Windows::Foundation::Size,struct winrt::Windows::Foundation::Size *))"},
            &TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride_Original,
            TaskbarCollapsibleLayoutXamlTraits_ArrangeOverride_Hook,
        },
        {
            {LR"(protected: virtual void __cdecl winrt::Taskbar::implementation::ExperienceToggleButton::UpdateButtonPadding(void))"},
            &ExperienceToggleButton_UpdateButtonPadding_Original,
            ExperienceToggleButton_UpdateButtonPadding_Hook,
        },
    };
    return HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks));
}

HMODULE GetTaskbarViewModuleHandle() {
    HMODULE module = GetModuleHandle(L"Taskbar.View.dll");
    if (!module) module = GetModuleHandle(L"ExplorerExtensions.dll");
    return module;
}

void HandleLoadedModuleIfTaskbarView(HMODULE module, LPCWSTR lpLibFileName) {
    if (!g_taskbarViewDllLoaded && GetTaskbarViewModuleHandle() == module &&
        !g_taskbarViewDllLoaded.exchange(true)) {
        if (HookTaskbarViewDllSymbols(module)) {
            Wh_ApplyHookOperations();
        }
    }
}

using LoadLibraryExW_t = decltype(&LoadLibraryExW);
LoadLibraryExW_t LoadLibraryExW_Original;
HMODULE WINAPI LoadLibraryExW_Hook(LPCWSTR lpLibFileName, HANDLE hFile, DWORD dwFlags) {
    HMODULE module = LoadLibraryExW_Original(lpLibFileName, hFile, dwFlags);
    if (module) HandleLoadedModuleIfTaskbarView(module, lpLibFileName);
    return module;
}

// --- Main Windhawk Functions ---

BOOL Wh_ModInit() {
    Wh_Log(L">");
    
    LoadSettings();

    if (!HookTaskbarDllSymbols()) return FALSE;

    if (HMODULE taskbarViewModule = GetTaskbarViewModuleHandle()) {
        g_taskbarViewDllLoaded = true;
        if (!HookTaskbarViewDllSymbols(taskbarViewModule)) return FALSE;
    } else {
        HMODULE kernelBaseModule = GetModuleHandle(L"kernelbase.dll");
        auto pKernelBaseLoadLibraryExW = (decltype(&LoadLibraryExW))GetProcAddress(kernelBaseModule, "LoadLibraryExW");
        WindhawkUtils::SetFunctionHook(pKernelBaseLoadLibraryExW, LoadLibraryExW_Hook, &LoadLibraryExW_Original);
    }

    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");
    if (!g_taskbarViewDllLoaded) {
        if (HMODULE taskbarViewModule = GetTaskbarViewModuleHandle()) {
            if (!g_taskbarViewDllLoaded.exchange(true)) {
                if (HookTaskbarViewDllSymbols(taskbarViewModule)) {
                    Wh_ApplyHookOperations();
                }
            }
        }
    }
    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (hTaskbarWnd) ApplySettings(hTaskbarWnd);
}

void Wh_ModBeforeUninit() {
    Wh_Log(L">");
    g_unloading = true;
    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (hTaskbarWnd) ApplySettings(hTaskbarWnd);
}

void Wh_ModUninit() {
    Wh_Log(L">");
}

BOOL Wh_ModSettingsChanged(BOOL* bReload) {
    *bReload = TRUE;
    return TRUE;
}