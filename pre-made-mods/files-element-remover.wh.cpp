// ==WindhawkMod==
// @id              files-element-remover
// @name            Files Element Remover
// @description     Permanently remove specific UI elements from the Files app visual tree
// @include         Files.exe
// @architecture    x86-64
// @compilerOptions -lcomctl32 -lole32 -loleaut32 -lruntimeobject -Wl,--export-all-symbols
// ==/WindhawkMod==

// ==WindhawkModSettings==
/*
- targets:
  - ""
  $name: Elements to remove (Target Selectors)
  $description: "Examples:\nAppBarButton#refreshButton\nTabViewItem\nGrid#TabContainerGrid > Border"
*/
// ==/WindhawkModSettings==

#include <xamlom.h>
#include <atomic>
#include <vector>
#include <string>
#include <string_view>
#include <list>
#include <optional>
#include <unordered_map>
#include <unordered_set>
#include <variant>
#include <sstream>

#undef GetCurrentTime
#include <winrt/Microsoft.UI.Xaml.h>
#include <winrt/Microsoft.UI.Xaml.Controls.h>
#include <winrt/Microsoft.UI.Xaml.Markup.h>
#include <winrt/Microsoft.UI.Xaml.Media.h>
#include <winrt/Microsoft.UI.Dispatching.h>
#include <winrt/Windows.Foundation.Collections.h>

#include <windhawk_utils.h>

using namespace winrt::Microsoft::UI::Xaml;
namespace mux = winrt::Microsoft::UI::Xaml;
using namespace winrt::Microsoft::UI::Xaml::Controls;

// ----------------------------------------------------------------------------
// Helpers
// ----------------------------------------------------------------------------

// Helper for settings management
template <auto fn>
struct deleter_from_fn {
    template <typename T>
    constexpr void operator()(T* arg) const {
        fn(arg);
    }
};
using string_setting_unique_ptr =
    std::unique_ptr<const WCHAR[], deleter_from_fn<Wh_FreeStringSetting>>;

HMODULE GetCurrentModuleHandle() {
    HMODULE module;
    if (!GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                               GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                           L"", &module)) {
        return nullptr;
    }
    return module;
}

// Global state
std::atomic<bool> g_initialized;
thread_local bool g_initializedForThread;

// Forward declarations
void ApplyCustomizations(InstanceHandle handle,
                         winrt::Microsoft::UI::Xaml::FrameworkElement element,
                         PCWSTR fallbackClassName);
void CleanupCustomizations(InstanceHandle handle);

// ----------------------------------------------------------------------------
// WinRT / VisualTreeWatcher Boilerplate
// ----------------------------------------------------------------------------
#pragma region visualtreewatcher

#include <Unknwn.h>
#include <winrt/base.h>

class VisualTreeWatcher : public winrt::implements<VisualTreeWatcher, IVisualTreeServiceCallback2, winrt::non_agile>
{
public:
    VisualTreeWatcher(winrt::com_ptr<IUnknown> site);
    ~VisualTreeWatcher();
    void UnadviseVisualTreeChange();

private:
    HRESULT STDMETHODCALLTYPE OnVisualTreeChange(ParentChildRelation relation, VisualElement element, VisualMutationType mutationType) override;
    HRESULT STDMETHODCALLTYPE OnElementStateChanged(InstanceHandle element, VisualElementState elementState, LPCWSTR context) noexcept override;

    winrt::Windows::Foundation::IInspectable FromHandle(InstanceHandle handle)
    {
        winrt::Windows::Foundation::IInspectable obj;
        winrt::check_hresult(m_XamlDiagnostics->GetIInspectableFromHandle(handle, reinterpret_cast<::IInspectable**>(winrt::put_abi(obj))));
        return obj;
    }

    winrt::com_ptr<IXamlDiagnostics> m_XamlDiagnostics = nullptr;
};

VisualTreeWatcher::VisualTreeWatcher(winrt::com_ptr<IUnknown> site) :
    m_XamlDiagnostics(site.as<IXamlDiagnostics>())
{
    // Create a thread to avoid hanging on the UI thread during initialization
    HANDLE thread = CreateThread(
        nullptr, 0,
        [](LPVOID lpParam) -> DWORD {
            auto watcher = reinterpret_cast<VisualTreeWatcher*>(lpParam);
            HRESULT hr = watcher->m_XamlDiagnostics.as<IVisualTreeService3>()->AdviseVisualTreeChange(watcher);
            watcher->Release();
            return 0;
        },
        this, 0, nullptr);
    if (thread) {
        AddRef();
        CloseHandle(thread);
    }
}

VisualTreeWatcher::~VisualTreeWatcher() {}

void VisualTreeWatcher::UnadviseVisualTreeChange()
{
    m_XamlDiagnostics.as<IVisualTreeService3>()->UnadviseVisualTreeChange(this);
}

HRESULT VisualTreeWatcher::OnVisualTreeChange(ParentChildRelation, VisualElement element, VisualMutationType mutationType) try
{
    if (mutationType == Add)
    {
        const auto inspectable = FromHandle(element.Handle);
        auto frameworkElement = inspectable.try_as<mux::FrameworkElement>();
        if (frameworkElement)
        {
            ApplyCustomizations(element.Handle, frameworkElement, element.Type);
        }
    }
    else if (mutationType == Remove)
    {
        CleanupCustomizations(element.Handle);
    }
    return S_OK;
}
catch (...)
{
    return S_OK; 
}

HRESULT VisualTreeWatcher::OnElementStateChanged(InstanceHandle, VisualElementState, LPCWSTR) noexcept
{
    return S_OK;
}

winrt::com_ptr<VisualTreeWatcher> g_visualTreeWatcher;

// {C85D8CC7-5463-40E8-A432-F5916B6427E5}
static constexpr CLSID CLSID_WindhawkTAP = { 0xc85d8cc7, 0x5463, 0x40e8, { 0xa4, 0x32, 0xf5, 0x91, 0x6b, 0x64, 0x27, 0xe5 } };

class WindhawkTAP : public winrt::implements<WindhawkTAP, IObjectWithSite, winrt::non_agile>
{
public:
    HRESULT STDMETHODCALLTYPE SetSite(IUnknown *pUnkSite) override;
    HRESULT STDMETHODCALLTYPE GetSite(REFIID riid, void **ppvSite) noexcept override;
private:
    winrt::com_ptr<IUnknown> site;
};

HRESULT WindhawkTAP::SetSite(IUnknown *pUnkSite) try
{
    if (g_visualTreeWatcher)
    {
        g_visualTreeWatcher->UnadviseVisualTreeChange();
        g_visualTreeWatcher = nullptr;
    }
    site.copy_from(pUnkSite);
    if (site)
    {
        FreeLibrary(GetCurrentModuleHandle());
        g_visualTreeWatcher = winrt::make_self<VisualTreeWatcher>(site);
    }
    return S_OK;
}
catch (...) { return winrt::to_hresult(); }

HRESULT WindhawkTAP::GetSite(REFIID riid, void **ppvSite) noexcept
{
    return site.as(riid, ppvSite);
}

template<class T>
struct SimpleFactory : winrt::implements<SimpleFactory<T>, IClassFactory, winrt::non_agile>
{
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override try
    {
        if (!pUnkOuter) { *ppvObject = nullptr; return winrt::make<T>().as(riid, ppvObject); }
        else { return CLASS_E_NOAGGREGATION; }
    }
    catch (...) { return winrt::to_hresult(); }

    HRESULT STDMETHODCALLTYPE LockServer(BOOL) noexcept override { return S_OK; }
};

#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wdll-attribute-on-redeclaration"

__declspec(dllexport)
_Use_decl_annotations_ STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) try
{
    if (rclsid == CLSID_WindhawkTAP) { *ppv = nullptr; return winrt::make<SimpleFactory<WindhawkTAP>>().as(riid, ppv); }
    else { return CLASS_E_CLASSNOTAVAILABLE; }
}
catch (...) { return winrt::to_hresult(); }

__declspec(dllexport)
_Use_decl_annotations_ STDAPI DllCanUnloadNow(void)
{
    return winrt::get_module_lock() ? S_FALSE : S_OK;
}
#pragma clang diagnostic pop

using PFN_INITIALIZE_XAML_DIAGNOSTICS_EX = decltype(&InitializeXamlDiagnosticsEx);

HRESULT InjectWindhawkTAP() noexcept
{
    HMODULE module = GetCurrentModuleHandle();
    if (!module) return HRESULT_FROM_WIN32(GetLastError());

    WCHAR location[MAX_PATH];
    if (!GetModuleFileName(module, location, ARRAYSIZE(location))) return HRESULT_FROM_WIN32(GetLastError());

    const HMODULE wux(GetModuleHandle(L"Microsoft.Internal.FrameworkUdk.dll"));
    if (!wux) return HRESULT_FROM_WIN32(GetLastError());

    const auto ixde = reinterpret_cast<PFN_INITIALIZE_XAML_DIAGNOSTICS_EX>(GetProcAddress(wux, "InitializeXamlDiagnosticsEx"));
    if (!ixde) return HRESULT_FROM_WIN32(GetLastError());

    HRESULT hr;
    for (int i = 0; i < 10000; i++)
    {
        WCHAR connectionName[256];
        wsprintf(connectionName, L"WinUIVisualDiagConnection%d", i + 1);
        hr = ixde(connectionName, GetCurrentProcessId(), L"", location, CLSID_WindhawkTAP, nullptr);
        if (hr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND)) break;
    }
    return hr;
}

#pragma endregion

// ----------------------------------------------------------------------------
// Matching Logic
// ----------------------------------------------------------------------------

std::wstring EscapeXmlAttribute(std::wstring_view data) {
    std::wstring buffer;
    buffer.reserve(data.size());
    for (const auto c : data) {
        switch (c) {
            case '&': buffer.append(L"&amp;"); break;
            case '\"': buffer.append(L"&quot;"); break;
            case '<': buffer.append(L"&lt;"); break;
            case '>': buffer.append(L"&gt;"); break;
            default: buffer.push_back(c); break;
        }
    }
    return buffer;
}

std::wstring_view TrimStringView(std::wstring_view s) {
    s.remove_prefix(std::min(s.find_first_not_of(L" \t\r\v\n"), s.size()));
    s.remove_suffix(std::min(s.size() - s.find_last_not_of(L" \t\r\v\n") - 1, s.size()));
    return s;
}

using PropertyKeyValue = std::pair<DependencyProperty, winrt::Windows::Foundation::IInspectable>;
using PropertyValuesUnresolved = std::vector<std::pair<std::wstring, std::wstring>>;
using PropertyValues = std::vector<PropertyKeyValue>;
using PropertyValuesMaybeUnresolved = std::variant<PropertyValuesUnresolved, PropertyValues>;

struct ElementMatcher {
    std::wstring type;
    std::wstring name;
    int oneBasedIndex = 0;
    PropertyValuesMaybeUnresolved propertyValues;
};

struct ElementRemovalRule {
    ElementMatcher elementMatcher;
    std::vector<ElementMatcher> parentElementMatchers;
};

thread_local std::vector<ElementRemovalRule> g_removalRules;

enum class ParentType { None, Panel, Border, ContentControl, UserControl, Viewbox };

// Structure for runtime state to restore elements on unload
// NOTE: C++/WinRT projected types like FrameworkElement act as smart pointers.
// We use FrameworkElement directly (not com_ptr<FrameworkElement>) to hold a strong reference.
struct ElementCleanupState {
    FrameworkElement keptElement = nullptr; 
    winrt::weak_ref<DependencyObject> parent;
    ParentType parentType = ParentType::None;
    int panelIndex = -1;
    bool removedByMod = false;
};

thread_local std::unordered_map<InstanceHandle, ElementCleanupState> g_elementsCleanupState;

Style GetStyleFromXamlSetters(const std::wstring_view type, const std::wstring_view xamlStyleSetters) {
    std::wstring xaml = LR"(<ResourceDictionary
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:muxc="using:Microsoft.UI.Xaml.Controls")";

    if (auto pos = type.rfind('.'); pos != type.npos) {
        auto typeNamespace = std::wstring_view(type).substr(0, pos);
        auto typeName = std::wstring_view(type).substr(pos + 1);
        xaml += L"\n    xmlns:target=\"using:";
        xaml += EscapeXmlAttribute(typeNamespace);
        xaml += L"\">\n    <Style TargetType=\"target:";
        xaml += EscapeXmlAttribute(typeName);
        xaml += L"\">\n";
    } else {
        xaml += L">\n    <Style TargetType=\"";
        xaml += EscapeXmlAttribute(type);
        xaml += L"\">\n";
    }

    xaml += xamlStyleSetters;
    xaml += L"    </Style>\n</ResourceDictionary>";

    auto resourceDictionary = Markup::XamlReader::Load(xaml).as<ResourceDictionary>();
    auto [styleKey, styleInspectable] = resourceDictionary.First().Current();
    return styleInspectable.as<Style>();
}

const PropertyValues& GetResolvedPropertyValues(const std::wstring_view type, PropertyValuesMaybeUnresolved* propertyValuesMaybeUnresolved) {
    if (const auto* resolved = std::get_if<PropertyValues>(propertyValuesMaybeUnresolved)) {
        return *resolved;
    }

    PropertyValues propertyValues;
    try {
        const auto& propertyValuesStr = std::get<PropertyValuesUnresolved>(*propertyValuesMaybeUnresolved);
        if (!propertyValuesStr.empty()) {
            std::wstring xaml;
            for (const auto& [property, value] : propertyValuesStr) {
                xaml += L"        <Setter Property=\"";
                xaml += EscapeXmlAttribute(property);
                xaml += L"\" Value=\"";
                xaml += EscapeXmlAttribute(value);
                xaml += L"\" />\n";
            }
            auto style = GetStyleFromXamlSetters(type, xaml);
            for (size_t i = 0; i < propertyValuesStr.size(); i++) {
                const auto setter = style.Setters().GetAt(i).as<Setter>();
                propertyValues.push_back({ setter.Property(), setter.Value() });
            }
        }
    } catch (...) {}

    *propertyValuesMaybeUnresolved = std::move(propertyValues);
    return std::get<PropertyValues>(*propertyValuesMaybeUnresolved);
}

bool TestElementMatcher(FrameworkElement element, ElementMatcher& matcher, PCWSTR fallbackClassName) {
    if (!matcher.type.empty() && matcher.type != winrt::get_class_name(element) && (!fallbackClassName || matcher.type != fallbackClassName)) {
        return false;
    }

    if (!matcher.name.empty() && matcher.name != element.Name()) {
        return false;
    }

    if (matcher.oneBasedIndex) {
        auto parent = Media::VisualTreeHelper::GetParent(element);
        if (!parent) return false;

        int index = matcher.oneBasedIndex - 1;
        if (index < 0 || index >= Media::VisualTreeHelper::GetChildrenCount(parent) || Media::VisualTreeHelper::GetChild(parent, index) != element) {
            return false;
        }
    }

    auto elementDo = element.as<DependencyObject>();
    for (const auto& propertyValue : GetResolvedPropertyValues(matcher.type, &matcher.propertyValues)) {
        auto value = elementDo.ReadLocalValue(propertyValue.first);
        if (value == DependencyProperty::UnsetValue()) return false;
        
        if (winrt::get_class_name(value) != winrt::get_class_name(propertyValue.second)) return false;
        
        try {
           if (winrt::unbox_value<winrt::hstring>(propertyValue.second) == winrt::unbox_value<winrt::hstring>(value)) continue;
        } catch (...) {}
        return false; 
    }

    return true;
}

// ----------------------------------------------------------------------------
// Core Logic: Remove Elements from Tree
// ----------------------------------------------------------------------------

void ApplyCustomizations(InstanceHandle handle, FrameworkElement element, PCWSTR fallbackClassName) {
    bool shouldRemove = false;

    for (auto& rule : g_removalRules) {
        if (!TestElementMatcher(element, rule.elementMatcher, fallbackClassName)) {
            continue;
        }

        auto parentElementIter = element;
        bool parentElementMatchFailed = false;

        for (auto& matcher : rule.parentElementMatchers) {
            parentElementIter = Media::VisualTreeHelper::GetParent(parentElementIter).try_as<FrameworkElement>();
            if (!parentElementIter) {
                parentElementMatchFailed = true;
                break;
            }
            if (!TestElementMatcher(parentElementIter, matcher, nullptr)) {
                parentElementMatchFailed = true;
                break;
            }
        }

        if (!parentElementMatchFailed) {
            shouldRemove = true;
            break;
        }
    }

    if (!shouldRemove) return;

    // We must ensure the element is valid before using DispatcherQueue
    if (!element) return;

    auto dispatcher = element.DispatcherQueue();
    if (dispatcher) {
        winrt::weak_ref<FrameworkElement> weakElem = element;
        dispatcher.TryEnqueue([weakElem, handle]() {
            auto elem = weakElem.get();
            if (!elem) return;

            auto parentDO = Media::VisualTreeHelper::GetParent(elem);
            if (!parentDO) return; // Already removed or root

            Wh_Log(L"Removing element: %s", elem.Name().c_str());

            if (g_elementsCleanupState.contains(handle)) return;

            ElementCleanupState state;
            state.removedByMod = true;
            state.keptElement = elem; // Strong reference
            state.parent = parentDO;

            bool removed = false;

            if (auto panel = parentDO.try_as<Panel>()) {
                uint32_t index = 0;
                if (panel.Children().IndexOf(elem, index)) {
                    state.parentType = ParentType::Panel;
                    state.panelIndex = index;
                    panel.Children().RemoveAt(index);
                    removed = true;
                }
            }
            else if (auto border = parentDO.try_as<Border>()) {
                state.parentType = ParentType::Border;
                border.Child(nullptr);
                removed = true;
            }
            else if (auto cc = parentDO.try_as<ContentControl>()) {
                state.parentType = ParentType::ContentControl;
                cc.Content(nullptr);
                removed = true;
            }
            else if (auto uc = parentDO.try_as<UserControl>()) {
                state.parentType = ParentType::UserControl;
                uc.Content(nullptr);
                removed = true;
            }
            else if (auto vb = parentDO.try_as<Viewbox>()) {
                state.parentType = ParentType::Viewbox;
                vb.Child(nullptr);
                removed = true;
            }

            if (removed) {
                // FrameworkElement is movable/copyable (it's a smart pointer)
                g_elementsCleanupState[handle] = std::move(state);
            }
        });
    }
}

void CleanupCustomizations(InstanceHandle handle) {
    if (auto it = g_elementsCleanupState.find(handle); it != g_elementsCleanupState.end()) {
        if (it->second.removedByMod) {
             return;
        }
        g_elementsCleanupState.erase(it);
    }
}

// ----------------------------------------------------------------------------
// Settings Parsing
// ----------------------------------------------------------------------------

ElementMatcher ElementMatcherFromString(std::wstring_view str) {
    ElementMatcher result;
    PropertyValuesUnresolved propertyValuesUnresolved;

    auto i = str.find_first_of(L"#@[");
    result.type = TrimStringView(str.substr(0, i));
    
    if (result.type.find_first_of(L".:") == result.type.npos) {
        if (result.type == L"Rectangle") result.type = L"Microsoft.UI.Xaml.Shapes.Rectangle";
        else result.type = L"Microsoft.UI.Xaml.Controls." + std::wstring{result.type};
    }

    while (i != str.npos) {
        auto iNext = str.find_first_of(L"#@[", i + 1);
        auto nextPart = str.substr(i + 1, iNext == str.npos ? str.npos : iNext - (i + 1));

        switch (str[i]) {
            case L'#':
                result.name = TrimStringView(nextPart);
                break;
            case L'[': {
                auto rule = TrimStringView(nextPart);
                if (rule.back() == L']') {
                    rule = TrimStringView(rule.substr(0, rule.length() - 1));
                    if (rule.find_first_not_of(L"0123456789") == rule.npos) {
                        result.oneBasedIndex = std::stoi(std::wstring(rule));
                    } else {
                        auto ruleEqPos = rule.find(L'=');
                        if (ruleEqPos != rule.npos) {
                            auto ruleKey = TrimStringView(rule.substr(0, ruleEqPos));
                            auto ruleVal = TrimStringView(rule.substr(ruleEqPos + 1));
                            propertyValuesUnresolved.push_back({std::wstring(ruleKey), std::wstring(ruleVal)});
                        }
                    }
                }
                break;
            }
        }
        i = iNext;
    }
    result.propertyValues = std::move(propertyValuesUnresolved);
    return result;
}

void ParseSettings() {
    g_removalRules.clear();

    for (int i = 0;; i++) {
        string_setting_unique_ptr targetStringSetting(Wh_GetStringSetting(L"targets[%d]", i));
        if (!*targetStringSetting.get()) break;

        std::wstring_view target = targetStringSetting.get();
        if (target.empty() || (target.size() >= 2 && target[0] == '/' && target[1] == '/')) continue;

        try {
            size_t pos_start = 0, pos_end;
            std::vector<std::wstring_view> parts;
            while ((pos_end = target.find(L" > ", pos_start)) != std::wstring_view::npos) {
                parts.push_back(target.substr(pos_start, pos_end - pos_start));
                pos_start = pos_end + 3;
            }
            parts.push_back(target.substr(pos_start));

            ElementRemovalRule rule;
            bool first = true;
            for (auto it = parts.rbegin(); it != parts.rend(); ++it) {
                auto matcher = ElementMatcherFromString(*it);
                if (first) {
                    rule.elementMatcher = std::move(matcher);
                    first = false;
                } else {
                    rule.parentElementMatchers.push_back(std::move(matcher));
                }
            }
            g_removalRules.push_back(std::move(rule));

        } catch (...) {
            Wh_Log(L"Failed to parse target: %s", target.data());
        }
    }
}

// ----------------------------------------------------------------------------
// Initialization & Restoration
// ----------------------------------------------------------------------------

void InitializeForCurrentThread() {
    if (g_initializedForThread) return;
    ParseSettings();
    g_initializedForThread = true;
}

void UninitializeForCurrentThread() {
    // Restore elements
    for (auto& [handle, state] : g_elementsCleanupState) {
        if (!state.removedByMod || !state.keptElement) continue;

        auto parent = state.parent.get();
        auto element = state.keptElement;
        
        if (parent) {
             try {
                 switch (state.parentType) {
                     case ParentType::Panel: {
                         auto panel = parent.as<Panel>();
                         if (state.panelIndex >= 0 && state.panelIndex <= panel.Children().Size()) {
                             panel.Children().InsertAt(state.panelIndex, element);
                         } else {
                             panel.Children().Append(element);
                         }
                         break;
                     }
                     case ParentType::Border:
                         parent.as<Border>().Child(element);
                         break;
                     case ParentType::ContentControl:
                         parent.as<ContentControl>().Content(element);
                         break;
                     case ParentType::UserControl:
                         parent.as<UserControl>().Content(element);
                         break;
                     case ParentType::Viewbox:
                         parent.as<Viewbox>().Child(element);
                         break;
                 }
             } catch (...) {
                 Wh_Log(L"Failed to restore element %s", element.Name().c_str());
             }
        }
    }
    
    g_elementsCleanupState.clear();
    g_removalRules.clear();
    g_initializedForThread = false;
}

void InitializeSettingsAndTap() {
    if (g_initialized.exchange(true)) return;
    HRESULT hr = InjectWindhawkTAP();
    if (FAILED(hr)) Wh_Log(L"Error %08X", hr);
}

void UninitializeSettingsAndTap() {
    if (g_visualTreeWatcher) {
        g_visualTreeWatcher->UnadviseVisualTreeChange();
        g_visualTreeWatcher = nullptr;
    }
    g_initialized = false;
}

bool IsTargetWindow(HWND hWnd) {
    WCHAR className[64];
    if (!GetClassName(hWnd, className, ARRAYSIZE(className))) return false;
    return _wcsicmp(className, L"Microsoft.UI.Content.DesktopChildSiteBridge") == 0 ||
           _wcsicmp(className, L"XamlExplorerHostIslandWindow_WASDK") == 0;
}

void OnWindowCreated(HWND hWnd) {
    if (IsTargetWindow(hWnd)) {
        InitializeForCurrentThread();
        InitializeSettingsAndTap();
    }
}

using CreateWindowExW_t = decltype(&CreateWindowExW);
CreateWindowExW_t CreateWindowExW_Original;
HWND WINAPI CreateWindowExW_Hook(DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName, DWORD dwStyle, int X, int Y, int nWidth, int nHeight, HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, PVOID lpParam) {
    HWND hWnd = CreateWindowExW_Original(dwExStyle, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);
    if (hWnd) OnWindowCreated(hWnd);
    return hWnd;
}

using RunFromWindowThreadProc_t = void(WINAPI*)(PVOID parameter);

bool RunFromWindowThread(HWND hWnd, RunFromWindowThreadProc_t proc, PVOID procParam) {
    static const UINT runFromWindowThreadRegisteredMsg =
        RegisterWindowMessage(L"Windhawk_RunFromWindowThread_" WH_MOD_ID);

    struct RUN_FROM_WINDOW_THREAD_PARAM {
        RunFromWindowThreadProc_t proc;
        PVOID procParam;
    };

    DWORD dwThreadId = GetWindowThreadProcessId(hWnd, nullptr);
    if (dwThreadId == 0) return false;
    if (dwThreadId == GetCurrentThreadId()) {
        proc(procParam);
        return true;
    }

    HHOOK hook = SetWindowsHookEx(
        WH_CALLWNDPROC,
        [](int nCode, WPARAM wParam, LPARAM lParam) -> LRESULT {
            if (nCode == HC_ACTION) {
                const CWPSTRUCT* cwp = (const CWPSTRUCT*)lParam;
                if (cwp->message == runFromWindowThreadRegisteredMsg) {
                    RUN_FROM_WINDOW_THREAD_PARAM* param =
                        (RUN_FROM_WINDOW_THREAD_PARAM*)cwp->lParam;
                    param->proc(param->procParam);
                }
            }
            return CallNextHookEx(nullptr, nCode, wParam, lParam);
        },
        nullptr, dwThreadId);
    if (!hook) return false;

    RUN_FROM_WINDOW_THREAD_PARAM param;
    param.proc = proc;
    param.procParam = procParam;
    SendMessage(hWnd, runFromWindowThreadRegisteredMsg, 0, (LPARAM)&param);
    UnhookWindowsHookEx(hook);
    return true;
}

std::vector<HWND> GetTargetWnds() {
    struct ENUM_WINDOWS_PARAM {
        std::vector<HWND>* hWnds;
    };

    std::vector<HWND> hWnds;
    ENUM_WINDOWS_PARAM param = {&hWnds};
    EnumWindows(
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            ENUM_WINDOWS_PARAM& param = *(ENUM_WINDOWS_PARAM*)lParam;
            DWORD dwProcessId = 0;
            if (!GetWindowThreadProcessId(hWnd, &dwProcessId) ||
                dwProcessId != GetCurrentProcessId()) {
                return TRUE;
            }
            if (IsTargetWindow(hWnd)) {
                param.hWnds->push_back(hWnd);
            }
            return TRUE;
        },
        (LPARAM)&param);

    return hWnds;
}

BOOL Wh_ModInit() {
    Wh_SetFunctionHook((void*)CreateWindowExW, (void*)CreateWindowExW_Hook, (void**)&CreateWindowExW_Original);
    return TRUE;
}

void Wh_ModAfterInit() {
    auto hTargetWnds = GetTargetWnds();
    for (auto hTargetWnd : hTargetWnds) {
        RunFromWindowThread(
            hTargetWnd, [](PVOID) { InitializeForCurrentThread(); }, nullptr);
    }
    if (hTargetWnds.size() > 0) {
        InitializeSettingsAndTap();
    }
}

void Wh_ModUninit() {
    UninitializeSettingsAndTap();
    auto hTargetWnds = GetTargetWnds();
    for (auto hTargetWnd : hTargetWnds) {
        RunFromWindowThread(
            hTargetWnd, [](PVOID) { UninitializeForCurrentThread(); }, nullptr);
    }
}

void Wh_ModSettingsChanged() {
    Wh_ModUninit();
    Wh_ModAfterInit();
}