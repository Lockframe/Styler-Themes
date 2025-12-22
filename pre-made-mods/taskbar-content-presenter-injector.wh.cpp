// ==WindhawkMod==
// @id              taskbar-content-presenter-injector
// @name            Taskbar ContentPresenter Injector
// @description     Injects a ContentPresenter into Taskbar.TaskListLabeledButtonPanel and Taskbar.TaskListButtonPanel
// @version         1.0
// @author          Gemini 3.0, Lockframe
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -lole32 -loleaut32 -lruntimeobject
// ==/WindhawkMod==

#include <windhawk_utils.h>

// Fix for conflict between Windows macro and WinRT method names
#undef GetCurrentTime

#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.UI.Xaml.Controls.h>
#include <winrt/Windows.UI.Xaml.Media.h>
#include <winrt/base.h>

#include <atomic>
#include <string>

using namespace winrt::Windows::UI::Xaml;

// Global state tracking
std::atomic<bool> g_taskbarViewDllLoaded = false;

const std::wstring c_TargetPanelLabeled = L"Taskbar.TaskListLabeledButtonPanel";
const std::wstring c_TargetPanelButton = L"Taskbar.TaskListButtonPanel";
const std::wstring c_InjectedControlName = L"CustomInjectedPresenter";

// -------------------------------------------------------------------------
// Original Function Pointers
// -------------------------------------------------------------------------
using TaskListButton_UpdateVisualStates_t = void(WINAPI*)(void* pThis);
TaskListButton_UpdateVisualStates_t TaskListButton_UpdateVisualStates_Original;

using TaskListButton_UpdateButtonPadding_t = void(WINAPI*)(void* pThis);
TaskListButton_UpdateButtonPadding_t TaskListButton_UpdateButtonPadding_Original;

using ExperienceToggleButton_UpdateVisualStates_t = void(WINAPI*)(void* pThis);
ExperienceToggleButton_UpdateVisualStates_t ExperienceToggleButton_UpdateVisualStates_Original;

using TaskbarFrame_OnTaskbarLayoutChildBoundsChanged_t = void(WINAPI*)(void* pThis);
TaskbarFrame_OnTaskbarLayoutChildBoundsChanged_t TaskbarFrame_OnTaskbarLayoutChildBoundsChanged_Original;

// -------------------------------------------------------------------------
// Helpers
// -------------------------------------------------------------------------

// Helper: Check if the element is already injected
bool IsAlreadyInjected(Controls::Panel panel) {
    for (auto child : panel.Children()) {
        if (auto elem = child.try_as<FrameworkElement>()) {
            if (elem.Name() == c_InjectedControlName) {
                return true;
            }
        }
    }
    return false;
}

// Logic to inject the ContentPresenter
void InjectContentPresenterIntoPanel(FrameworkElement targetPanel) {
    if (!targetPanel) return;

    auto panel = targetPanel.try_as<Controls::Panel>();
    if (!panel) return;

    if (IsAlreadyInjected(panel)) return;

    Controls::ContentPresenter presenter;
    presenter.Name(c_InjectedControlName);
    presenter.HorizontalAlignment(HorizontalAlignment::Stretch);
    presenter.VerticalAlignment(VerticalAlignment::Stretch);

    // Wh_Log(L"Injecting into %s", winrt::get_class_name(targetPanel).c_str());
    panel.Children().Append(presenter);
}

// Universal scanner: Checks current element and recurses
// Returns true if injection happened to avoid unnecessary deep scanning of that branch (optional optimization)
void ScanAndInjectRecursive(FrameworkElement element) {
    if (!element) return;

    std::wstring className = winrt::get_class_name(element).c_str();

    // Check if THIS element is one of our targets
    if (className == c_TargetPanelLabeled || className == c_TargetPanelButton) {
        InjectContentPresenterIntoPanel(element);
        // We generally don't need to look inside the panel itself for another panel
        return; 
    }

    // Recurse into children
    int childrenCount = Media::VisualTreeHelper::GetChildrenCount(element);
    for (int i = 0; i < childrenCount; i++) {
        auto childDependencyObject = Media::VisualTreeHelper::GetChild(element, i);
        auto child = childDependencyObject.try_as<FrameworkElement>();
        if (child) {
            ScanAndInjectRecursive(child);
        }
    }
}

// Helper to get FrameworkElement from native implementation pointer
FrameworkElement GetFrameworkElementFromNative(void* pThis) {
    try {
        void* iUnknownPtr = (void**)pThis + 3;
        winrt::Windows::Foundation::IUnknown iUnknown;
        winrt::copy_from_abi(iUnknown, iUnknownPtr);
        return iUnknown.try_as<FrameworkElement>();
    } catch (...) {
        return nullptr;
    }
}

// -------------------------------------------------------------------------
// Hooks
// -------------------------------------------------------------------------

// Hook for Standard App Buttons
void WINAPI TaskListButton_UpdateVisualStates_Hook(void* pThis) {
    TaskListButton_UpdateVisualStates_Original(pThis);
    if (auto elem = GetFrameworkElementFromNative(pThis)) {
        ScanAndInjectRecursive(elem);
    }
}

void WINAPI TaskListButton_UpdateButtonPadding_Hook(void* pThis) {
    TaskListButton_UpdateButtonPadding_Original(pThis);
    if (auto elem = GetFrameworkElementFromNative(pThis)) {
        ScanAndInjectRecursive(elem);
    }
}

// Hook for System Buttons (Search, Widgets, etc.)
void WINAPI ExperienceToggleButton_UpdateVisualStates_Hook(void* pThis) {
    ExperienceToggleButton_UpdateVisualStates_Original(pThis);
    if (auto elem = GetFrameworkElementFromNative(pThis)) {
        ScanAndInjectRecursive(elem);
    }
}

// Global Layout Hook: Catches everything inside TaskbarFrame
void WINAPI TaskbarFrame_OnTaskbarLayoutChildBoundsChanged_Hook(void* pThis) {
    TaskbarFrame_OnTaskbarLayoutChildBoundsChanged_Original(pThis);

    auto taskbarFrame = GetFrameworkElementFromNative(pThis);
    if (!taskbarFrame) return;

    // Perform a full scan of the TaskbarFrame visual tree.
    // This finds ANY TaskListButtonPanel or TaskListLabeledButtonPanel
    // regardless of whether they are in an ExperienceToggleButton, TaskListButton, or other container.
    ScanAndInjectRecursive(taskbarFrame);
}

// -------------------------------------------------------------------------
// Initialization Logic
// -------------------------------------------------------------------------

bool HookTaskbarViewDllSymbols(HMODULE module) {
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(private: void __cdecl winrt::Taskbar::implementation::TaskListButton::UpdateVisualStates(void))"},
            &TaskListButton_UpdateVisualStates_Original,
            TaskListButton_UpdateVisualStates_Hook,
        },
        {
            {LR"(private: void __cdecl winrt::Taskbar::implementation::TaskListButton::UpdateButtonPadding(void))"},
            &TaskListButton_UpdateButtonPadding_Original,
            TaskListButton_UpdateButtonPadding_Hook,
        },
        {
            // Note: If this symbol is missing/renamed in your specific Windows version, 
            // the hook will be skipped, but the Layout hook (TaskbarFrame) will still cover it.
            {LR"(private: void __cdecl winrt::Taskbar::implementation::ExperienceToggleButton::UpdateVisualStates(void))"},
            &ExperienceToggleButton_UpdateVisualStates_Original,
            ExperienceToggleButton_UpdateVisualStates_Hook,
            true // Optional, don't fail init if missing
        },
        {
            {LR"(private: void __cdecl winrt::Taskbar::implementation::TaskbarFrame::OnTaskbarLayoutChildBoundsChanged(void))"},
            &TaskbarFrame_OnTaskbarLayoutChildBoundsChanged_Original,
            TaskbarFrame_OnTaskbarLayoutChildBoundsChanged_Hook,
        }
    };

    if (!HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks))) {
        Wh_Log(L"Failed to hook Taskbar.View.dll symbols");
        return false;
    }

    return true;
}

HMODULE GetTaskbarViewModuleHandle() {
    HMODULE module = GetModuleHandle(L"Taskbar.View.dll");
    if (!module) {
        module = GetModuleHandle(L"ExplorerExtensions.dll");
    }
    return module;
}

void HandleLoadedModuleIfTaskbarView(HMODULE module, LPCWSTR lpLibFileName) {
    if (!g_taskbarViewDllLoaded && GetTaskbarViewModuleHandle() == module &&
        !g_taskbarViewDllLoaded.exchange(true)) {
        
        Wh_Log(L"Taskbar View DLL loaded: %s", lpLibFileName);
        
        if (HookTaskbarViewDllSymbols(module)) {
            Wh_ApplyHookOperations();
        }
    }
}

using LoadLibraryExW_t = decltype(&LoadLibraryExW);
LoadLibraryExW_t LoadLibraryExW_Original;

HMODULE WINAPI LoadLibraryExW_Hook(LPCWSTR lpLibFileName, HANDLE hFile, DWORD dwFlags) {
    HMODULE module = LoadLibraryExW_Original(lpLibFileName, hFile, dwFlags);
    if (module) {
        HandleLoadedModuleIfTaskbarView(module, lpLibFileName);
    }
    return module;
}

BOOL Wh_ModInit() {
    Wh_Log(L"Initializing Taskbar Injector Mod v1.3");

    if (HMODULE taskbarViewModule = GetTaskbarViewModuleHandle()) {
        g_taskbarViewDllLoaded = true;
        if (!HookTaskbarViewDllSymbols(taskbarViewModule)) {
            return FALSE;
        }
    } else {
        HMODULE kernelBaseModule = GetModuleHandle(L"kernelbase.dll");
        auto pKernelBaseLoadLibraryExW = (decltype(&LoadLibraryExW))GetProcAddress(kernelBaseModule, "LoadLibraryExW");
        WindhawkUtils::Wh_SetFunctionHookT(pKernelBaseLoadLibraryExW, LoadLibraryExW_Hook, &LoadLibraryExW_Original);
    }

    return TRUE;
}

void Wh_ModUninit() {
    Wh_Log(L"Uninitializing Taskbar Injector Mod");
}