// ==WindhawkMod==
// @id              files-styler
// @name            Files Styler
// @description     Customize the Files app
// @version         1.5.1
// @author          Lockframe
// @include         Files.exe
// @architecture    x86-64
// @compilerOptions -lcomctl32 -lgdi32 -ldwmapi -lole32 -loleaut32 -lruntimeobject -Wl,--export-all-symbols
// ==/WindhawkMod==

// ==WindhawkModSettings==
/*
- styleConstants: [""]
  $name: Style constants
  $description: >-
    Some themes support style constants for customization, such as colors. Refer
    to the theme page for available constants. For technical details, refer to
    the mod description.
- controlStyles:
  - - target: ""
      $name: Target
    - styles: [""]
      $name: Styles
  $name: Control styles
- themeResourceVariables: [""]
  $name: Resource variables
  $description: >-
    Use "Key=Value" to override an existing resource with a new value.

    Use "Key@Dark=Value" or "Key@Light=Value" to define theme-aware resources
    that can be referenced with {ThemeResource Key} in styles.

    The ":=" syntax can be used to set a XAML value. For details, refer to the
    mod description.
- xamlDiagnosticsHandling: alert
  $name: XAML diagnostics consumer handling
  $description: >-
    How to handle other programs (e.g. ExplorerBlurMica) that try to use XAML
    diagnostics. There can only be one consumer at a time. Block will prevent
    other programs from using it, which might break them. Allow will let them
    use it, which might break this mod.
  $options:
  - alert: Alert (prompt before blocking)
  - block: Block other consumers
  - allow: Allow other consumers
*/
// ==/WindhawkModSettings==

#include <xamlom.h>

#include <atomic>
#include <optional>
#include <vector>

#undef GetCurrentTime

#include <winrt/Microsoft.UI.Xaml.h>

enum class XamlDiagnosticsHandling {
    kAlert,
    kBlock,
    kAllow,
};

struct {
    XamlDiagnosticsHandling xamlDiagnosticsHandling;
} g_settings;

std::atomic<bool> g_initialized;
thread_local bool g_initializedForThread;

void ApplyCustomizations(InstanceHandle handle,
                         winrt::Microsoft::UI::Xaml::FrameworkElement element,
                         PCWSTR fallbackClassName);
void CleanupCustomizations(InstanceHandle handle);
bool HasCustomizationRulesForType(PCWSTR type);

HMODULE GetCurrentModuleHandle() {
    HMODULE module;
    if (!GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                               GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                           L"", &module)) {
        return nullptr;
    }

    return module;
}

////////////////////////////////////////////////////////////////////////////////
// clang-format off

#pragma region winrt_hpp

#include <Unknwn.h>
#include <winrt/base.h>

namespace winrt {
    namespace Windows {
        namespace Foundation {}
    }
    namespace Microsoft {
        namespace UI::Xaml {}
    }
}

namespace wf = winrt::Windows::Foundation;
namespace mux = winrt::Microsoft::UI::Xaml;

#pragma endregion  // winrt_hpp

#pragma region visualtreewatcher_hpp

#include <winrt/Microsoft.UI.Xaml.h>

class VisualTreeWatcher : public winrt::implements<VisualTreeWatcher, IVisualTreeServiceCallback2, winrt::non_agile>
{
public:
    VisualTreeWatcher(winrt::com_ptr<IUnknown> site);

    VisualTreeWatcher(const VisualTreeWatcher&) = delete;
    VisualTreeWatcher& operator=(const VisualTreeWatcher&) = delete;

    VisualTreeWatcher(VisualTreeWatcher&&) = delete;
    VisualTreeWatcher& operator=(VisualTreeWatcher&&) = delete;

    ~VisualTreeWatcher();

    void UnadviseVisualTreeChange();

private:
    HRESULT STDMETHODCALLTYPE OnVisualTreeChange(ParentChildRelation relation, VisualElement element, VisualMutationType mutationType) override;
    HRESULT STDMETHODCALLTYPE OnElementStateChanged(InstanceHandle element, VisualElementState elementState, LPCWSTR context) noexcept override;

    wf::IInspectable FromHandle(InstanceHandle handle)
    {
        wf::IInspectable obj;
        winrt::check_hresult(m_XamlDiagnostics->GetIInspectableFromHandle(handle, reinterpret_cast<::IInspectable**>(winrt::put_abi(obj))));
        return obj;
    }

    winrt::com_ptr<IXamlDiagnostics> m_XamlDiagnostics = nullptr;
};

#pragma endregion  // visualtreewatcher_hpp

#pragma region visualtreewatcher_cpp

VisualTreeWatcher::VisualTreeWatcher(winrt::com_ptr<IUnknown> site) :
    m_XamlDiagnostics(site.as<IXamlDiagnostics>())
{
    Wh_Log(L"Constructing VisualTreeWatcher");

    HANDLE thread = CreateThread(
        nullptr, 0,
        [](LPVOID lpParam) -> DWORD {
            auto watcher = reinterpret_cast<VisualTreeWatcher*>(lpParam);
            HRESULT hr = watcher->m_XamlDiagnostics.as<IVisualTreeService3>()->AdviseVisualTreeChange(watcher);
            watcher->Release();
            if (FAILED(hr)) {
                Wh_Log(L"Error %08X", hr);
            }
            return 0;
        },
        this, 0, nullptr);
    if (thread) {
        AddRef();
        CloseHandle(thread);
    }
}

VisualTreeWatcher::~VisualTreeWatcher()
{
    Wh_Log(L"Destructing VisualTreeWatcher");
}

void VisualTreeWatcher::UnadviseVisualTreeChange()
{
    Wh_Log(L"UnadviseVisualTreeChange VisualTreeWatcher");
    HRESULT hr = m_XamlDiagnostics.as<IVisualTreeService3>()->UnadviseVisualTreeChange(this);
    if (FAILED(hr)) {
        Wh_Log(L"UnadviseVisualTreeChange failed with error %08X", hr);
    }
}

HRESULT VisualTreeWatcher::OnVisualTreeChange(ParentChildRelation, VisualElement element, VisualMutationType mutationType) try
{
    if (!g_initializedForThread)
    {
        return S_OK;
    }

    if (mutationType == Add)
    {
        if (!HasCustomizationRulesForType(element.Type)) {
            return S_OK;
        }

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
    HRESULT hr = winrt::to_hresult();
    Wh_Log(L"Error %08X", hr);
    return S_OK;
}

HRESULT VisualTreeWatcher::OnElementStateChanged(InstanceHandle, VisualElementState, LPCWSTR) noexcept
{
    return S_OK;
}

#pragma endregion  // visualtreewatcher_cpp

#pragma region tap_hpp

#include <ocidl.h>

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

#pragma endregion  // tap_hpp

#pragma region tap_cpp

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
catch (...)
{
    HRESULT hr = winrt::to_hresult();
    Wh_Log(L"Error %08X", hr);
    return hr;
}

HRESULT WindhawkTAP::GetSite(REFIID riid, void **ppvSite) noexcept
{
    return site.as(riid, ppvSite);
}

#pragma endregion  // tap_cpp

#pragma region simplefactory_hpp

#include <Unknwn.h>

template<class T>
struct SimpleFactory : winrt::implements<SimpleFactory<T>, IClassFactory, winrt::non_agile>
{
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override try
    {
        if (!pUnkOuter)
        {
            *ppvObject = nullptr;
            return winrt::make<T>().as(riid, ppvObject);
        }
        else
        {
            return CLASS_E_NOAGGREGATION;
        }
    }
    catch (...)
    {
        HRESULT hr = winrt::to_hresult();
        Wh_Log(L"Error %08X", hr);
        return hr;
    }

    HRESULT STDMETHODCALLTYPE LockServer(BOOL) noexcept override
    {
        return S_OK;
    }
};

#pragma endregion  // simplefactory_hpp

#pragma region module_cpp

#include <combaseapi.h>

#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wdll-attribute-on-redeclaration"

__declspec(dllexport)
_Use_decl_annotations_ STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) try
{
    if (rclsid == CLSID_WindhawkTAP)
    {
        *ppv = nullptr;
        return winrt::make<SimpleFactory<WindhawkTAP>>().as(riid, ppv);
    }
    else
    {
        return CLASS_E_CLASSNOTAVAILABLE;
    }
}
catch (...)
{
    HRESULT hr = winrt::to_hresult();
    Wh_Log(L"Error %08X", hr);
    return hr;
}

__declspec(dllexport)
_Use_decl_annotations_ STDAPI DllCanUnloadNow()
{
    if (winrt::get_module_lock())
    {
        return S_FALSE;
    }
    else
    {
        return S_OK;
    }
}

#pragma clang diagnostic pop

#pragma endregion  // module_cpp

#pragma region api_cpp

bool g_inInjectWindhawkTAP = false;

using PFN_INITIALIZE_XAML_DIAGNOSTICS_EX = decltype(&InitializeXamlDiagnosticsEx);

HRESULT InjectWindhawkTAP() noexcept
{
    HMODULE module = GetCurrentModuleHandle();
    if (!module)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    WCHAR location[MAX_PATH];
    switch (GetModuleFileName(module, location, ARRAYSIZE(location)))
    {
    case 0:
    case ARRAYSIZE(location):
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const HMODULE wux(GetModuleHandle(L"Microsoft.Internal.FrameworkUdk.dll"));
    if (!wux) [[unlikely]]
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const auto ixde = reinterpret_cast<PFN_INITIALIZE_XAML_DIAGNOSTICS_EX>(GetProcAddress(wux, "InitializeXamlDiagnosticsEx"));
    if (!ixde) [[unlikely]]
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    g_inInjectWindhawkTAP = true;

    HRESULT hr;
    for (int i = 0; i < 10000; i++)
    {
        WCHAR connectionName[256];
        wsprintf(connectionName, L"WinUIVisualDiagConnection%d", i + 1);

        hr = ixde(connectionName, GetCurrentProcessId(), L"", location, CLSID_WindhawkTAP, nullptr);
        if (hr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
        {
            break;
        }
    }

    g_inInjectWindhawkTAP = false;

    return hr;
}

#pragma endregion  // api_cpp

// clang-format on
////////////////////////////////////////////////////////////////////////////////

#include <windhawk_utils.h>

#include <algorithm>
#include <cmath>
#include <limits>
#include <list>
#include <memory>
#include <mutex>
#include <optional>
#include <random>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <variant>
#include <vector>

using namespace std::string_view_literals;

#include <initguid.h>

#include <commctrl.h>
#include <d2d1_1.h>
#include <dwmapi.h>
#include <roapi.h>
#include <windows.graphics.effects.h>
#include <winstring.h>

#include <winrt/Microsoft.UI.Composition.h>
#include <winrt/Microsoft.UI.Dispatching.h>
#include <winrt/Microsoft.UI.Text.h>
#include <winrt/Microsoft.UI.Xaml.Controls.h>
#include <winrt/Microsoft.UI.Xaml.Hosting.h>
#include <winrt/Microsoft.UI.Xaml.Markup.h>
#include <winrt/Microsoft.UI.Xaml.Media.Imaging.h>
#include <winrt/Microsoft.UI.Xaml.Media.h>
#include <winrt/Microsoft.UI.Xaml.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Graphics.Effects.h>
#include <winrt/Windows.Networking.Connectivity.h>
#include <winrt/Windows.Storage.Streams.h>
#include <winrt/Windows.System.Power.h>
#include <winrt/Windows.UI.ViewManagement.h>

using namespace winrt::Microsoft::UI::Xaml;

namespace wge = winrt::Windows::Graphics::Effects;
namespace muc = winrt::Microsoft::UI::Composition;
namespace muxh = mux::Hosting;
namespace awge = ABI::Windows::Graphics::Effects;

template <auto fn>
struct deleter_from_fn {
    template <typename T>
    constexpr void operator()(T* arg) const {
        fn(arg);
    }
};
using string_setting_unique_ptr =
    std::unique_ptr<const WCHAR[], deleter_from_fn<Wh_FreeStringSetting>>;

using PropertyKeyValue =
    std::pair<DependencyProperty, winrt::Windows::Foundation::IInspectable>;

using PropertyValuesUnresolved =
    std::vector<std::pair<std::wstring, std::wstring>>;
using PropertyValues = std::vector<PropertyKeyValue>;
using PropertyValuesMaybeUnresolved =
    std::variant<PropertyValuesUnresolved, PropertyValues>;

struct ElementMatcher {
    enum class Kind {
        Element,   // Normal element matcher.
        Wildcard,  // '*': matches zero or more intermediate ancestors.
        Root,      // ':root': asserts the next element has no parent.
    };
    Kind kind = Kind::Element;
    std::wstring type;
    std::wstring name;
    std::optional<std::wstring> visualStateGroupName;
    int oneBasedIndex = 0;
    PropertyValuesMaybeUnresolved propertyValues;
};

struct ValueRule {
    std::wstring propertyName;
    std::wstring visualState;
    std::wstring value;
    bool isXamlValue = false;

    bool isDynamic() const { return value.find(L"{{") != std::wstring::npos; }
};

struct CaptureRule {
    std::wstring propertyName;
    std::wstring varName;
};

struct UnresolvedRules {
    std::vector<ValueRule> valueRules;
    std::vector<CaptureRule> captureRules;
};

struct XamlBlurBrushParams {
    float blurAmount;
    winrt::Windows::UI::Color tint;
    std::optional<uint8_t> tintOpacity;
    std::wstring tintThemeResourceKey;
    std::optional<float> tintLuminosityOpacity;
    std::optional<float> tintSaturation;
    std::optional<float> noiseOpacity;
    std::optional<float> noiseDensity;
    std::optional<winrt::Windows::UI::Color> fallbackColor;
    std::wstring fallbackThemeResourceKey;
};

struct DynamicStyleTemplate {
    std::wstring propertyName;
    std::wstring rawValue;
    bool isXamlValue = false;
};

using PropertyOverrideValue =
    std::variant<winrt::Windows::Foundation::IInspectable,
                 XamlBlurBrushParams,
                 DynamicStyleTemplate>;

using PropertyOverrides =
    std::unordered_map<DependencyProperty,
                       std::unordered_map<std::wstring, PropertyOverrideValue>>;

struct CaptureSpec {
    DependencyProperty property{nullptr};
    std::wstring varName;
};

struct ResolvedRules {
    PropertyOverrides propertyOverrides;
    std::vector<CaptureSpec> captures;
    bool hasDynamicValues = false;
};

using PropertyOverridesMaybeUnresolved =
    std::variant<UnresolvedRules, ResolvedRules>;

struct StyleVariableDependency {
    std::wstring name;
    InstanceHandle owner = 0;
};

struct ElementTreeNode {
    winrt::weak_ref<DependencyObject> ref;
    std::shared_ptr<ElementTreeNode> parent;
    uint32_t depth = 0;
};

thread_local std::unordered_map<void*, std::weak_ptr<ElementTreeNode>>
    g_elementTreeNodes;

thread_local size_t g_elementTreeNodesReapThreshold = 64;

void* ElementIdentityKey(DependencyObject const& object) {
    return winrt::get_abi(object.as<winrt::Windows::Foundation::IUnknown>());
}

std::shared_ptr<ElementTreeNode> GetOrCreateElementTreeNode(
    DependencyObject object) {
    if (!object) {
        return nullptr;
    }

    std::shared_ptr<ElementTreeNode> node;
    std::vector<DependencyObject> missing;

    try {
        for (auto iter = object; iter;
             iter = Media::VisualTreeHelper::GetParent(iter)) {
            auto key = ElementIdentityKey(iter);

            if (auto it = g_elementTreeNodes.find(key);
                it != g_elementTreeNodes.end()) {
                auto existing = it->second.lock();
                if (!existing || !existing->ref.get()) {
                    g_elementTreeNodes.erase(it);
                } else if (existing->depth > 0 ||
                           !Media::VisualTreeHelper::GetParent(iter)) {
                    node = std::move(existing);
                    break;
                } else {
                    g_elementTreeNodes.erase(it);
                }
            }

            missing.push_back(iter);
        }

        for (auto it = missing.rbegin(); it != missing.rend(); ++it) {
            auto fresh = std::make_shared<ElementTreeNode>();
            fresh->ref = *it;
            fresh->depth = node ? node->depth + 1 : 0;
            fresh->parent = std::move(node);
            g_elementTreeNodes[ElementIdentityKey(*it)] = fresh;
            node = std::move(fresh);
        }
    } catch (...) {
        return nullptr;
    }

    return node;
}

void ReapElementTreeNodesIfNeeded() {
    if (g_elementTreeNodes.size() < g_elementTreeNodesReapThreshold) {
        return;
    }

    std::erase_if(g_elementTreeNodes,
                  [](const auto& item) { return item.second.expired(); });
    g_elementTreeNodesReapThreshold =
        std::max<size_t>(64, g_elementTreeNodes.size() * 2);
}

int ElementTreeLcaDepth(ElementTreeNode const* a, ElementTreeNode const* b) {
    if (!a || !b) {
        return -1;
    }

    while (a->depth > b->depth) {
        a = a->parent.get();
    }
    while (b->depth > a->depth) {
        b = b->parent.get();
    }

    while (a != b) {
        a = a->parent.get();
        b = b->parent.get();
        if (!a || !b) {
            return -1;
        }
    }

    return static_cast<int>(a->depth);
}

struct ElementCustomizationRules {
    ElementMatcher elementMatcher;
    std::vector<ElementMatcher> parentElementMatchers;
    PropertyOverridesMaybeUnresolved propertyOverrides;
};

thread_local std::vector<ElementCustomizationRules>
    g_elementsCustomizationRules;

struct StringHash {
    using is_transparent = void;
    size_t operator()(std::wstring_view txt) const { return std::hash<std::wstring_view>{}(txt); }
};

thread_local std::unordered_map<std::wstring, bool, StringHash, std::equal_to<>> g_typeMatchCache;
thread_local std::unordered_map<std::wstring, std::vector<const ElementCustomizationRules*>, StringHash, std::equal_to<>> g_rulesByTypeMap;
thread_local std::vector<const ElementCustomizationRules*> g_genericRules;
thread_local std::unordered_set<std::wstring, StringHash, std::equal_to<>> g_candidateTypeSet;

void BuildRuleIndex() {
    g_rulesByTypeMap.clear();
    g_genericRules.clear();
    g_candidateTypeSet.clear();
    g_typeMatchCache.clear();

    for (const auto& rules : g_elementsCustomizationRules) {
        std::wstring_view target = rules.elementMatcher.type;
        if (target.empty() || target == L"*" || target == L":root" ||
            target.find(L"IUIElementOverrides") != target.npos ||
            target.find(L"FrameworkElement") != target.npos ||
            target.find(L"Control") != target.npos ||
            target.find(L"UserControl") != target.npos) {
            g_genericRules.push_back(&rules);
            g_candidateTypeSet.insert(L"Microsoft.UI.Xaml.IUIElementOverrides");
            g_candidateTypeSet.insert(L"Microsoft.UI.Xaml.FrameworkElement");
            g_candidateTypeSet.insert(L"Microsoft.UI.Xaml.Controls.Control");
            g_candidateTypeSet.insert(L"Microsoft.UI.Xaml.Controls.ContentControl");
            g_candidateTypeSet.insert(L"Microsoft.UI.Xaml.Controls.UserControl");
            g_candidateTypeSet.insert(L"IUIElementOverrides");
            g_candidateTypeSet.insert(L"FrameworkElement");
            g_candidateTypeSet.insert(L"Control");
            g_candidateTypeSet.insert(L"ContentControl");
            g_candidateTypeSet.insert(L"UserControl");
        } else {
            g_rulesByTypeMap[std::wstring(target)].push_back(&rules);
            g_candidateTypeSet.insert(std::wstring(target));

            if (auto pos = target.rfind(L'.'); pos != target.npos) {
                std::wstring shortName(target.substr(pos + 1));
                g_rulesByTypeMap[shortName].push_back(&rules);
                g_candidateTypeSet.insert(std::move(shortName));
            }
        }
    }
}

bool HasCustomizationRulesForType(PCWSTR type) {
    if (!type || g_elementsCustomizationRules.empty()) return false;
    if (!g_genericRules.empty()) return true;
    std::wstring_view typeView(type);

    auto it = g_typeMatchCache.find(typeView);
    if (it != g_typeMatchCache.end()) return it->second;

    bool matched = false;
    if (typeView.find(L"IUIElementOverrides") != typeView.npos ||
        typeView.find(L"FrameworkElement") != typeView.npos ||
        typeView.find(L"Control") != typeView.npos ||
        typeView.find(L"UserControl") != typeView.npos) {
        matched = true;
    } else if (g_candidateTypeSet.contains(typeView)) {
        matched = true;
    } else if (auto pos = typeView.rfind(L'.'); pos != typeView.npos) {
        if (g_candidateTypeSet.contains(typeView.substr(pos + 1))) {
            matched = true;
        }
    }

    g_typeMatchCache[std::wstring(typeView)] = matched;
    return matched;
}

struct ElementPropertyCustomizationState {
    std::optional<winrt::Windows::Foundation::IInspectable> originalValue;
    std::optional<PropertyOverrideValue> customValue;
    winrt::Windows::Foundation::IInspectable lastAppliedValue{nullptr};
    int64_t propertyChangedToken = 0;
    std::optional<DynamicStyleTemplate> dynamicTemplate;
    std::vector<StyleVariableDependency> variableDependencies;
    bool lastResolveFailed = false;
};

struct CapturePropertyCustomizationState {
    std::wstring varName;
    int64_t propertyChangedToken = 0;
};

struct ElementCustomizationStateForVisualStateGroup {
    std::unordered_map<DependencyProperty, ElementPropertyCustomizationState>
        propertyCustomizationStates;
    winrt::event_token visualStateGroupCurrentStateChangedToken;
};

struct ElementCustomizationState {
    winrt::weak_ref<FrameworkElement> element;
    std::shared_ptr<ElementTreeNode> treeNode;
    std::unordered_map<DependencyProperty, CapturePropertyCustomizationState>
        captureCustomizationStates;
    winrt::event_token captureSizeChangedToken;

    std::list<std::pair<std::optional<winrt::weak_ref<VisualStateGroup>>,
                        ElementCustomizationStateForVisualStateGroup>>
        perVisualStateGroup;
};

thread_local std::unordered_map<InstanceHandle, ElementCustomizationState>
    g_elementsCustomizationState;

ElementTreeNode* EnsureElementTreeNode(
    ElementCustomizationState& elementCustomizationState) {
    if (!elementCustomizationState.treeNode) {
        if (auto element = elementCustomizationState.element.get()) {
            elementCustomizationState.treeNode =
                GetOrCreateElementTreeNode(element);
        }
    }

    return elementCustomizationState.treeNode.get();
}

struct StyleVariableValue {
    std::wstring stringForm;
    std::optional<double> numeric;
    bool substitutable = false;
};

struct StyleVariableCapture {
    InstanceHandle elementHandle;
    StyleVariableValue value;
};

struct StyleVariableConsumer {
    InstanceHandle elementHandle;
    DependencyProperty property{nullptr};
    std::wstring fallbackClassName;
};

struct StyleVariableState {
    std::unordered_map<std::wstring, std::vector<StyleVariableCapture>>
        variables;
    std::unordered_map<std::wstring, std::vector<StyleVariableConsumer>>
        consumers;
};

thread_local StyleVariableState g_styleVariableState;

thread_local int g_styleVariablePropagationDepth;

struct PendingStyleVariablePropagation {
    StyleVariableState* state;
    std::wstring varName;
    std::optional<InstanceHandle> changedOwner;

    bool operator==(const PendingStyleVariablePropagation&) const = default;
};

thread_local std::vector<PendingStyleVariablePropagation>
    g_pendingStyleVariablePropagations;

StyleVariableState* GetStyleVariableState() {
    return &g_styleVariableState;
}

thread_local bool g_elementPropertyModifying;

struct TrackedImageBrush {
    winrt::weak_ref<Media::ImageBrush> brush;
    winrt::Windows::Foundation::Uri uri{nullptr};

    int32_t decodePixelWidth = 0;
    int32_t decodePixelHeight = 0;
    Media::Imaging::DecodePixelType decodePixelType =
        Media::Imaging::DecodePixelType::Physical;
    Media::Imaging::BitmapCreateOptions createOptions =
        Media::Imaging::BitmapCreateOptions::None;
    bool autoPlay = true;

    Media::ImageBrush::ImageFailed_revoker imageFailedRevoker;
    Media::ImageBrush::ImageOpened_revoker imageOpenedRevoker;

    bool loaded = false;

    ULONGLONG lastRetryTick = 0;
    int retryCount = 0;
};

struct TrackedImageBrushesForThread {
    std::list<std::shared_ptr<TrackedImageBrush>> brushes;
    winrt::Microsoft::UI::Dispatching::DispatcherQueue dispatcher{nullptr};
    winrt::Microsoft::UI::Dispatching::DispatcherQueueTimer retryDebounceTimer{
        nullptr};
    winrt::Microsoft::UI::Dispatching::DispatcherQueueTimer::Tick_revoker
        retryDebounceTimerTickRevoker;
};

thread_local TrackedImageBrushesForThread g_trackedImageBrushesForThread;

constexpr DWORD kNetworkChangeDebounceMs = 2000;
constexpr ULONGLONG kImageRetryBaseDelayMs = 5000;
constexpr int kImageRetryMaxBackoffShift = 6;
constexpr int kImageRetryMaxCount = 20;

std::mutex g_imageRetryMutex;
bool g_imageRetryActive;
std::vector<winrt::weak_ref<winrt::Microsoft::UI::Dispatching::DispatcherQueue>>
    g_imageRetryDispatchers;
winrt::event_token g_networkStatusChangedToken;

enum class ResourceVariableTheme {
    None,
    Dark,
    Light,
};

enum class ResourceVariableType {
    String,
    Xaml,
    ThemeResourceReference,
};

struct ResourceVariableEntry {
    std::wstring key;
    std::wstring value;
    ResourceVariableTheme theme;
    ResourceVariableType type;
};

thread_local std::vector<ResourceVariableEntry> g_resourceVariables;

thread_local std::unordered_map<std::wstring,
                                winrt::Windows::Foundation::IInspectable>
    g_originalResourceValues;

thread_local ResourceDictionary g_resourceVariablesThemeDict{nullptr};

thread_local winrt::Windows::UI::ViewManagement::UISettings g_uiSettings{
    nullptr};
thread_local winrt::event_token g_colorValuesChangedToken;

winrt::Windows::Foundation::IInspectable ReadLocalValueWithWorkaround(
    DependencyObject elementDo,
    DependencyProperty property) {
    auto value = elementDo.ReadLocalValue(property);
    if (value) {
        if (value == DependencyProperty::UnsetValue()) {
            auto grid = elementDo.try_as<Controls::Grid>();
            if (grid && grid.Name() == L"NavigationBarControlGrid") {
                auto value2 = elementDo.GetValue(property);
                if (value2 && winrt::get_class_name(value2) ==
                                  L"Microsoft.UI.Xaml.Controls."
                                  L"ColumnDefinitionCollection") {
                    value = std::move(value2);
                }
            }
        }
    }

    return value;
}

////////////////////////////////////////////////////////////////////////////////
// Noise generation
winrt::Windows::Storage::Streams::IRandomAccessStream CreateNoiseStream(
    float density) {
    thread_local float cachedDensity = std::numeric_limits<float>::quiet_NaN();
    thread_local winrt::Windows::Storage::Streams::InMemoryRandomAccessStream
        cachedStream{nullptr};

    if (density == cachedDensity && cachedStream) {
        return cachedStream.CloneStream();
    }

    constexpr int kSize = 256;
    constexpr DWORD kBpp = 32;
    constexpr DWORD rowSize = kSize * (kBpp / 8);
    constexpr DWORD dataSize = rowSize * kSize;

    BITMAPFILEHEADER fileHeader{
        .bfType = 0x4D42,
        .bfSize =
            sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER) + dataSize,
        .bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER),
    };

    BITMAPINFOHEADER infoHeader{
        .biSize = sizeof(BITMAPINFOHEADER),
        .biWidth = kSize,
        .biHeight = kSize,
        .biPlanes = 1,
        .biBitCount = kBpp,
        .biSizeImage = dataSize,
    };

    std::vector<uint8_t> pixels(dataSize);

    float safeDensity = std::clamp(density, 0.001f, 1.0f);
    float exponent = 1.0f / safeDensity;

    uint8_t lut[256];
    for (int i = 0; i < 256; i++) {
        lut[i] = static_cast<uint8_t>(std::pow(i / 255.0f, exponent) * 255.0f);
    }

    std::mt19937 rng(0);
    std::uniform_int_distribution<int> dist(0, 255);

    for (size_t i = 0; i < pixels.size(); i += 4) {
        uint8_t gray = lut[dist(rng)];
        pixels[i] = gray;
        pixels[i + 1] = gray;
        pixels[i + 2] = gray;
        pixels[i + 3] = 255;
    }

    winrt::Windows::Storage::Streams::InMemoryRandomAccessStream stream;
    winrt::Windows::Storage::Streams::DataWriter writer(stream);
    writer.WriteBytes(winrt::array_view<const uint8_t>(
        reinterpret_cast<const uint8_t*>(&fileHeader), sizeof(fileHeader)));
    writer.WriteBytes(winrt::array_view<const uint8_t>(
        reinterpret_cast<const uint8_t*>(&infoHeader), sizeof(infoHeader)));
    writer.WriteBytes(pixels);
    writer.StoreAsync().get();
    writer.DetachStream();

    cachedStream = std::move(stream);
    cachedDensity = density;

    return cachedStream.CloneStream();
}

////////////////////////////////////////////////////////////////////////////////
// D2D1 Effect Interop wrappers
template <> inline constexpr winrt::guid winrt::impl::guid_v<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>{
    winrt::impl::guid_v<winrt::Windows::Foundation::IPropertyValue>
};

typedef enum MY_D2D1_GAUSSIANBLUR_OPTIMIZATION
{
    MY_D2D1_GAUSSIANBLUR_OPTIMIZATION_SPEED = 0,
    MY_D2D1_GAUSSIANBLUR_OPTIMIZATION_BALANCED = 1,
    MY_D2D1_GAUSSIANBLUR_OPTIMIZATION_QUALITY = 2,
    MY_D2D1_GAUSSIANBLUR_OPTIMIZATION_FORCE_DWORD = 0xffffffff

} MY_D2D1_GAUSSIANBLUR_OPTIMIZATION;

#ifndef BUILD_WINDOWS
namespace ABI {
#endif
namespace Windows {
namespace Graphics {
namespace Effects {

typedef interface IGraphicsEffectSource                         IGraphicsEffectSource;
typedef interface IGraphicsEffectD2D1Interop                    IGraphicsEffectD2D1Interop;

typedef enum GRAPHICS_EFFECT_PROPERTY_MAPPING
{
    GRAPHICS_EFFECT_PROPERTY_MAPPING_UNKNOWN,
    GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT,
    GRAPHICS_EFFECT_PROPERTY_MAPPING_VECTORX,
    GRAPHICS_EFFECT_PROPERTY_MAPPING_VECTORY,
    GRAPHICS_EFFECT_PROPERTY_MAPPING_VECTORZ,
    GRAPHICS_EFFECT_PROPERTY_MAPPING_VECTORW,
    GRAPHICS_EFFECT_PROPERTY_MAPPING_RECT_TO_VECTOR4,
    GRAPHICS_EFFECT_PROPERTY_MAPPING_RADIANS_TO_DEGREES,
    GRAPHICS_EFFECT_PROPERTY_MAPPING_COLORMATRIX_ALPHA_MODE,
    GRAPHICS_EFFECT_PROPERTY_MAPPING_COLOR_TO_VECTOR3,
    GRAPHICS_EFFECT_PROPERTY_MAPPING_COLOR_TO_VECTOR4
} GRAPHICS_EFFECT_PROPERTY_MAPPING;

#undef INTERFACE
#define INTERFACE IGraphicsEffectD2D1Interop
DECLARE_INTERFACE_IID_(IGraphicsEffectD2D1Interop, IUnknown, "2FC57384-A068-44D7-A331-30982FCF7177")
{
    STDMETHOD(GetEffectId)(_Out_ GUID * id) PURE;
    STDMETHOD(GetNamedPropertyMapping)(LPCWSTR name, _Out_ UINT * index, _Out_ GRAPHICS_EFFECT_PROPERTY_MAPPING * mapping) PURE;
    STDMETHOD(GetPropertyCount)(_Out_ UINT * count) PURE;
    STDMETHOD(GetProperty)(UINT index, _Outptr_ winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue> ** value) PURE;
    STDMETHOD(GetSource)(UINT index, _Outptr_ IGraphicsEffectSource ** source) PURE;
    STDMETHOD(GetSourceCount)(_Out_ UINT * count) PURE;
};

} // namespace Effects
} // namespace Graphics
} // namespace Windows
#ifndef BUILD_WINDOWS
} // namespace ABI
#endif

template <> inline constexpr winrt::guid winrt::impl::guid_v<ABI::Windows::Graphics::Effects::IGraphicsEffectD2D1Interop>{
    0x2FC57384, 0xA068, 0x44D7, { 0xA3, 0x31, 0x30, 0x98, 0x2F, 0xCF, 0x71, 0x77 }
};

struct CompositeEffect : winrt::implements<CompositeEffect, wge::IGraphicsEffect, wge::IGraphicsEffectSource, awge::IGraphicsEffectD2D1Interop>
{
public:
    HRESULT STDMETHODCALLTYPE GetEffectId(GUID* id) noexcept override {
        if (!id) return E_INVALIDARG;
        *id = CLSID_D2D1Composite;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetNamedPropertyMapping(LPCWSTR name, UINT* index, awge::GRAPHICS_EFFECT_PROPERTY_MAPPING* mapping) noexcept override {
        if (!index || !mapping) return E_INVALIDARG;
        if (std::wstring_view(name) == L"Mode") {
            *index = D2D1_COMPOSITE_PROP_MODE;
            *mapping = awge::GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT;
            return S_OK;
        }
        return E_INVALIDARG;
    }
    HRESULT STDMETHODCALLTYPE GetPropertyCount(UINT* count) noexcept override {
        if (!count) return E_INVALIDARG;
        *count = 1;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetProperty(UINT index, winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>** value) noexcept override try {
        if (!value) return E_INVALIDARG;
        if (index == D2D1_COMPOSITE_PROP_MODE) {
            *value = wf::PropertyValue::CreateUInt32((UINT32)Mode).as<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>().detach();
            return S_OK;
        }
        return E_BOUNDS;
    } catch (...) { return winrt::to_hresult(); }
    HRESULT STDMETHODCALLTYPE GetSource(UINT index, awge::IGraphicsEffectSource** source) noexcept override try {
        if (!source) return E_INVALIDARG;
        winrt::copy_to_abi(Sources.at(index), *reinterpret_cast<void**>(source));
        return S_OK;
    } catch (...) { return winrt::to_hresult(); }
    HRESULT STDMETHODCALLTYPE GetSourceCount(UINT* count) noexcept override {
        if (!count) return E_INVALIDARG;
        *count = static_cast<UINT>(Sources.size());
        return S_OK;
    }
    winrt::hstring Name() { return m_name; }
    void Name(winrt::hstring name) { m_name = name; }

    std::vector<wge::IGraphicsEffectSource> Sources;
    D2D1_COMPOSITE_MODE Mode = D2D1_COMPOSITE_MODE_SOURCE_OVER;
private:
    winrt::hstring m_name = L"CompositeEffect";
};

struct FloodEffect : winrt::implements<FloodEffect, wge::IGraphicsEffect, wge::IGraphicsEffectSource, awge::IGraphicsEffectD2D1Interop>
{
public:
    HRESULT STDMETHODCALLTYPE GetEffectId(GUID* id) noexcept override {
        if (!id) return E_INVALIDARG;
        *id = CLSID_D2D1Flood;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetNamedPropertyMapping(LPCWSTR name, UINT* index, awge::GRAPHICS_EFFECT_PROPERTY_MAPPING* mapping) noexcept override {
        if (!index || !mapping) return E_INVALIDARG;
        if (std::wstring_view(name) == L"Color") {
            *index = D2D1_FLOOD_PROP_COLOR;
            *mapping = awge::GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT;
            return S_OK;
        }
        return E_INVALIDARG;
    }
    HRESULT STDMETHODCALLTYPE GetPropertyCount(UINT* count) noexcept override {
        if (!count) return E_INVALIDARG;
        *count = 1;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetProperty(UINT index, winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>** value) noexcept override try {
        if (!value) return E_INVALIDARG;
        if (index == D2D1_FLOOD_PROP_COLOR) {
            *value = wf::PropertyValue::CreateSingleArray({
                Color.R / 255.0f, Color.G / 255.0f, Color.B / 255.0f, Color.A / 255.0f
            }).as<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>().detach();
            return S_OK;
        }
        return E_BOUNDS;
    } catch (...) { return winrt::to_hresult(); }
    HRESULT STDMETHODCALLTYPE GetSource(UINT, awge::IGraphicsEffectSource** source) noexcept override {
        if (!source) return E_INVALIDARG;
        return E_BOUNDS;
    }
    HRESULT STDMETHODCALLTYPE GetSourceCount(UINT* count) noexcept override {
        if (!count) return E_INVALIDARG;
        *count = 0;
        return S_OK;
    }
    winrt::hstring Name() { return m_name; }
    void Name(winrt::hstring name) { m_name = name; }

    winrt::Windows::UI::Color Color{};
private:
    winrt::hstring m_name = L"FloodEffect";
};

struct BorderEffect : winrt::implements<BorderEffect, wge::IGraphicsEffect, wge::IGraphicsEffectSource, awge::IGraphicsEffectD2D1Interop>
{
public:
    HRESULT STDMETHODCALLTYPE GetEffectId(GUID* id) noexcept override {
        if (!id) return E_INVALIDARG;
        *id = CLSID_D2D1Border;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetNamedPropertyMapping(LPCWSTR name, UINT* index, awge::GRAPHICS_EFFECT_PROPERTY_MAPPING* mapping) noexcept override {
        if (!index || !mapping) return E_INVALIDARG;
        std::wstring_view sv(name);
        if (sv == L"ExtendX") { *index = D2D1_BORDER_PROP_EDGE_MODE_X; *mapping = awge::GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT; return S_OK; }
        if (sv == L"ExtendY") { *index = D2D1_BORDER_PROP_EDGE_MODE_Y; *mapping = awge::GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT; return S_OK; }
        return E_INVALIDARG;
    }
    HRESULT STDMETHODCALLTYPE GetPropertyCount(UINT* count) noexcept override { if (!count) return E_INVALIDARG; *count = 2; return S_OK; }
    HRESULT STDMETHODCALLTYPE GetProperty(UINT index, winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>** value) noexcept override try {
        if (!value) return E_INVALIDARG;
        if (index == D2D1_BORDER_PROP_EDGE_MODE_X) { *value = wf::PropertyValue::CreateUInt32((UINT32)ExtendX).as<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>().detach(); return S_OK; }
        if (index == D2D1_BORDER_PROP_EDGE_MODE_Y) { *value = wf::PropertyValue::CreateUInt32((UINT32)ExtendY).as<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>().detach(); return S_OK; }
        return E_BOUNDS;
    } catch (...) { return winrt::to_hresult(); }
    HRESULT STDMETHODCALLTYPE GetSource(UINT index, awge::IGraphicsEffectSource** source) noexcept override {
        if (!source) return E_INVALIDARG;
        if (index == 0 && Source) { winrt::copy_to_abi(Source, *reinterpret_cast<void**>(source)); return S_OK; }
        return E_BOUNDS;
    }
    HRESULT STDMETHODCALLTYPE GetSourceCount(UINT* count) noexcept override { if (!count) return E_INVALIDARG; *count = 1; return S_OK; }
    winrt::hstring Name() { return m_name; }
    void Name(winrt::hstring name) { m_name = name; }

    wge::IGraphicsEffectSource Source{nullptr};
    D2D1_BORDER_EDGE_MODE ExtendX = D2D1_BORDER_EDGE_MODE_WRAP;
    D2D1_BORDER_EDGE_MODE ExtendY = D2D1_BORDER_EDGE_MODE_WRAP;
private:
    winrt::hstring m_name = L"BorderEffect";
};

struct GaussianBlurEffect : winrt::implements<GaussianBlurEffect, wge::IGraphicsEffect, wge::IGraphicsEffectSource, awge::IGraphicsEffectD2D1Interop>
{
public:
    HRESULT STDMETHODCALLTYPE GetEffectId(GUID* id) noexcept override { if (!id) return E_INVALIDARG; *id = CLSID_D2D1GaussianBlur; return S_OK; }
    HRESULT STDMETHODCALLTYPE GetNamedPropertyMapping(LPCWSTR name, UINT* index, awge::GRAPHICS_EFFECT_PROPERTY_MAPPING* mapping) noexcept override {
        if (!index || !mapping) return E_INVALIDARG;
        std::wstring_view sv(name);
        if (sv == L"BlurAmount") { *index = D2D1_GAUSSIANBLUR_PROP_STANDARD_DEVIATION; *mapping = awge::GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT; return S_OK; }
        if (sv == L"Optimization") { *index = D2D1_GAUSSIANBLUR_PROP_OPTIMIZATION; *mapping = awge::GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT; return S_OK; }
        if (sv == L"BorderMode") { *index = D2D1_GAUSSIANBLUR_PROP_BORDER_MODE; *mapping = awge::GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT; return S_OK; }
        return E_INVALIDARG;
    }
    HRESULT STDMETHODCALLTYPE GetPropertyCount(UINT* count) noexcept override { if (!count) return E_INVALIDARG; *count = 3; return S_OK; }
    HRESULT STDMETHODCALLTYPE GetProperty(UINT index, winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>** value) noexcept override try {
        if (!value) return E_INVALIDARG;
        if (index == D2D1_GAUSSIANBLUR_PROP_STANDARD_DEVIATION) { *value = wf::PropertyValue::CreateSingle(BlurAmount).as<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>().detach(); return S_OK; }
        if (index == D2D1_GAUSSIANBLUR_PROP_OPTIMIZATION) { *value = wf::PropertyValue::CreateUInt32((UINT32)Optimization).as<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>().detach(); return S_OK; }
        if (index == D2D1_GAUSSIANBLUR_PROP_BORDER_MODE) { *value = wf::PropertyValue::CreateUInt32((UINT32)BorderMode).as<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>().detach(); return S_OK; }
        return E_BOUNDS;
    } catch (...) { return winrt::to_hresult(); }
    HRESULT STDMETHODCALLTYPE GetSource(UINT index, awge::IGraphicsEffectSource** source) noexcept override {
        if (!source) return E_INVALIDARG;
        if (index == 0) { winrt::copy_to_abi(Source, *reinterpret_cast<void**>(source)); return S_OK; }
        return E_BOUNDS;
    }
    HRESULT STDMETHODCALLTYPE GetSourceCount(UINT* count) noexcept override { if (!count) return E_INVALIDARG; *count = 1; return S_OK; }
    winrt::hstring Name() { return m_name; }
    void Name(winrt::hstring name) { m_name = name; }

    wge::IGraphicsEffectSource Source;
    float BlurAmount = 3.0f;
    MY_D2D1_GAUSSIANBLUR_OPTIMIZATION Optimization = MY_D2D1_GAUSSIANBLUR_OPTIMIZATION_BALANCED;
    D2D1_BORDER_MODE BorderMode = D2D1_BORDER_MODE_SOFT;
private:
    winrt::hstring m_name = L"GaussianBlurEffect";
};

struct ColorMatrixEffect : winrt::implements<ColorMatrixEffect, wge::IGraphicsEffect, wge::IGraphicsEffectSource, awge::IGraphicsEffectD2D1Interop>
{
public:
    HRESULT STDMETHODCALLTYPE GetEffectId(GUID* id) noexcept override { if (!id) return E_INVALIDARG; *id = CLSID_D2D1ColorMatrix; return S_OK; }
    HRESULT STDMETHODCALLTYPE GetNamedPropertyMapping(LPCWSTR name, UINT* index, awge::GRAPHICS_EFFECT_PROPERTY_MAPPING* mapping) noexcept override {
        if (!index || !mapping) return E_INVALIDARG;
        std::wstring_view sv(name);
        if (sv == L"ColorMatrix") { *index = D2D1_COLORMATRIX_PROP_COLOR_MATRIX; *mapping = awge::GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT; return S_OK; }
        if (sv == L"AlphaMode") { *index = D2D1_COLORMATRIX_PROP_ALPHA_MODE; *mapping = awge::GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT; return S_OK; }
        if (sv == L"ClampOutput") { *index = D2D1_COLORMATRIX_PROP_CLAMP_OUTPUT; *mapping = awge::GRAPHICS_EFFECT_PROPERTY_MAPPING_DIRECT; return S_OK; }
        return E_INVALIDARG;
    }
    HRESULT STDMETHODCALLTYPE GetPropertyCount(UINT* count) noexcept override { if (!count) return E_INVALIDARG; *count = 3; return S_OK; }
    HRESULT STDMETHODCALLTYPE GetProperty(UINT index, winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>** value) noexcept override try {
        if (!value) return E_INVALIDARG;
        if (index == D2D1_COLORMATRIX_PROP_COLOR_MATRIX) {
            *value = wf::PropertyValue::CreateSingleArray(winrt::array_view<const float>(Matrix, Matrix + 20)).as<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>().detach();
            return S_OK;
        }
        if (index == D2D1_COLORMATRIX_PROP_ALPHA_MODE) { *value = wf::PropertyValue::CreateUInt32(AlphaMode).as<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>().detach(); return S_OK; }
        if (index == D2D1_COLORMATRIX_PROP_CLAMP_OUTPUT) { *value = wf::PropertyValue::CreateBoolean(ClampOutput).as<winrt::impl::abi_t<winrt::Windows::Foundation::IPropertyValue>>().detach(); return S_OK; }
        return E_BOUNDS;
    } catch (...) { return winrt::to_hresult(); }
    HRESULT STDMETHODCALLTYPE GetSource(UINT index, awge::IGraphicsEffectSource** source) noexcept override {
        if (!source) return E_INVALIDARG;
        if (index == 0 && Source) { winrt::copy_to_abi(Source, *reinterpret_cast<void**>(source)); return S_OK; }
        return E_BOUNDS;
    }
    HRESULT STDMETHODCALLTYPE GetSourceCount(UINT* count) noexcept override { if (!count) return E_INVALIDARG; *count = 1; return S_OK; }
    winrt::hstring Name() { return m_name; }
    void Name(winrt::hstring name) { m_name = name; }

    wge::IGraphicsEffectSource Source{nullptr};
    float Matrix[20] = { 1,0,0,0, 0,1,0,0, 0,0,1,0, 0,0,0,1, 0,0,0,0 };
    uint32_t AlphaMode = D2D1_COLORMATRIX_ALPHA_MODE_PREMULTIPLIED;
    bool ClampOutput = false;
private:
    winrt::hstring m_name = L"ColorMatrixEffect";
};

class XamlBlurBrush : public Media::XamlCompositionBrushBaseT<XamlBlurBrush>
{
public:
    XamlBlurBrush(UIElement element,
                  float blurAmount,
                  winrt::Windows::UI::Color tint,
                  std::optional<uint8_t> tintOpacity,
                  winrt::hstring tintThemeResourceKey,
                  std::optional<float> tintLuminosityOpacity,
                  std::optional<float> tintSaturation,
                  std::optional<float> noiseOpacity,
                  std::optional<float> noiseDensity,
                  std::optional<winrt::Windows::UI::Color> fallbackColor,
                  winrt::hstring fallbackThemeResourceKey);
    ~XamlBlurBrush();

    void OnConnected();
    void OnDisconnected();

private:
    void RefreshThemeTint();
    void RefreshFallbackColor();
    bool ShouldUseFallback() const;
    void RefreshBrush();
    muc::CompositionBrush CreateEffectBrush();
    muc::CompositionBrush CreateFallbackBrush();

    muc::Compositor m_compositor;
    float m_blurAmount;
    winrt::Windows::UI::Color m_tint;
    std::optional<uint8_t> m_tintOpacity;
    winrt::hstring m_tintThemeResourceKey;
    std::optional<float> m_tintLuminosityOpacity;
    std::optional<float> m_tintSaturation;
    std::optional<float> m_noiseOpacity;
    std::optional<float> m_noiseDensity;
    std::optional<winrt::Windows::UI::Color> m_fallbackColor;
    winrt::hstring m_fallbackThemeResourceKey;
    Media::SolidColorBrush m_proxyBrush{nullptr};
    Media::SolidColorBrush m_fallbackProxyBrush{nullptr};
    winrt::weak_ref<FrameworkElement> m_weakProxyElement;
    winrt::hstring m_proxyKey;
    winrt::hstring m_fallbackProxyKey;
    winrt::Windows::UI::ViewManagement::UISettings m_uiSettings{nullptr};
    winrt::event_token m_advancedEffectsEnabledChangedToken{};
    winrt::event_token m_energySaverStatusChangedToken{};
    winrt::Microsoft::UI::Dispatching::DispatcherQueue m_dispatcher{nullptr};
    HKEY m_powerKey{nullptr};
    HANDLE m_regNotifyEvent{nullptr};
    HANDLE m_regWaitHandle{nullptr};

    static void CALLBACK OnEnergySaverRegistryChanged(PVOID context, BOOLEAN timerOrWaitFired);
};

XamlBlurBrush::XamlBlurBrush(UIElement element,
                             float blurAmount,
                             winrt::Windows::UI::Color tint,
                             std::optional<uint8_t> tintOpacity,
                             winrt::hstring tintThemeResourceKey,
                             std::optional<float> tintLuminosityOpacity,
                             std::optional<float> tintSaturation,
                             std::optional<float> noiseOpacity,
                             std::optional<float> noiseDensity,
                             std::optional<winrt::Windows::UI::Color> fallbackColor,
                             winrt::hstring fallbackThemeResourceKey) :
    m_compositor(muxh::ElementCompositionPreview::GetElementVisual(element).Compositor()),
    m_blurAmount(blurAmount),
    m_tint(tint),
    m_tintOpacity(tintOpacity),
    m_tintThemeResourceKey(std::move(tintThemeResourceKey)),
    m_tintLuminosityOpacity(tintLuminosityOpacity),
    m_tintSaturation(tintSaturation),
    m_noiseOpacity(noiseOpacity),
    m_noiseDensity(noiseDensity),
    m_fallbackColor(fallbackColor),
    m_fallbackThemeResourceKey(std::move(fallbackThemeResourceKey))
{
    auto fe = element.try_as<FrameworkElement>();

    auto createProxy = [&](winrt::hstring const& themeResourceKey) -> Media::SolidColorBrush {
        if (!fe) return nullptr;
        std::wstring xaml =
            L"<SolidColorBrush xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\" Color=\"{ThemeResource " +
            std::wstring(themeResourceKey) + L"}\"/>";
        try {
            return Markup::XamlReader::Load(winrt::hstring(xaml)).try_as<Media::SolidColorBrush>();
        } catch (...) {
            return nullptr;
        }
    };

    static std::atomic<uint64_t> s_proxyCounter{0};

    if (!m_tintThemeResourceKey.empty()) {
        if (auto proxyBrush = createProxy(m_tintThemeResourceKey)) {
            auto proxyKey = winrt::hstring(L"__WhBlurProxy_" + std::to_wstring(++s_proxyCounter));
            fe.Resources().Insert(winrt::box_value(proxyKey), proxyBrush);
            m_proxyBrush = proxyBrush;
            m_weakProxyElement = winrt::make_weak(fe);
            m_proxyKey = proxyKey;
        }

        if (m_proxyBrush) {
            m_proxyBrush.RegisterPropertyChangedCallback(
                Media::SolidColorBrush::ColorProperty(),
                [weakThis = get_weak()](auto&&, auto&&) {
                    if (auto self = weakThis.get()) self->RefreshBrush();
                });
        }
    }

    if (!m_fallbackThemeResourceKey.empty()) {
        if (auto proxyBrush = createProxy(m_fallbackThemeResourceKey)) {
            auto proxyKey = winrt::hstring(L"__WhBlurFallbackProxy_" + std::to_wstring(++s_proxyCounter));
            fe.Resources().Insert(winrt::box_value(proxyKey), proxyBrush);
            m_fallbackProxyBrush = proxyBrush;
            if (!m_weakProxyElement.get()) m_weakProxyElement = winrt::make_weak(fe);
            m_fallbackProxyKey = proxyKey;
        }

        if (m_fallbackProxyBrush) {
            m_fallbackProxyBrush.RegisterPropertyChangedCallback(
                Media::SolidColorBrush::ColorProperty(),
                [weakThis = get_weak()](auto&&, auto&&) {
                    if (auto self = weakThis.get()) self->RefreshBrush();
                });
        }
    }

    if (m_fallbackColor || !m_fallbackThemeResourceKey.empty()) {
        m_dispatcher = winrt::Microsoft::UI::Dispatching::DispatcherQueue::GetForCurrentThread();
        try {
            m_uiSettings = winrt::Windows::UI::ViewManagement::UISettings();
            auto dispatcher = m_dispatcher;
            m_advancedEffectsEnabledChangedToken = m_uiSettings.AdvancedEffectsEnabledChanged(
                [weakThis = get_weak(), dispatcher](auto&&, auto&&) {
                    dispatcher.TryEnqueue([weakThis] {
                        if (auto self = weakThis.get()) self->RefreshBrush();
                    });
                });
            m_energySaverStatusChangedToken = winrt::Windows::System::Power::PowerManager::EnergySaverStatusChanged(
                [weakThis = get_weak(), dispatcher](auto&&, auto&&) {
                    dispatcher.TryEnqueue([weakThis] {
                        if (auto self = weakThis.get()) self->RefreshBrush();
                    });
                });
        } catch (...) {}
    }
}

void CALLBACK XamlBlurBrush::OnEnergySaverRegistryChanged(PVOID context, BOOLEAN) {
    auto* self = static_cast<XamlBlurBrush*>(context);
    if (self->m_powerKey && self->m_regNotifyEvent) {
        RegNotifyChangeKeyValue(self->m_powerKey, FALSE, REG_NOTIFY_CHANGE_LAST_SET, self->m_regNotifyEvent, TRUE);
    }
    if (self->m_dispatcher) {
        auto weakThis = self->get_weak();
        self->m_dispatcher.TryEnqueue([weakThis] {
            if (auto strongThis = weakThis.get()) strongThis->RefreshBrush();
        });
    }
}

XamlBlurBrush::~XamlBlurBrush() {
    if (m_regWaitHandle) { UnregisterWaitEx(m_regWaitHandle, INVALID_HANDLE_VALUE); m_regWaitHandle = nullptr; }
    if (m_regNotifyEvent) { CloseHandle(m_regNotifyEvent); m_regNotifyEvent = nullptr; }
    if (m_powerKey) { RegCloseKey(m_powerKey); m_powerKey = nullptr; }
    if (m_uiSettings && m_advancedEffectsEnabledChangedToken.value) {
        try { m_uiSettings.AdvancedEffectsEnabledChanged(m_advancedEffectsEnabledChangedToken); } catch (...) {}
    }
    if (m_energySaverStatusChangedToken.value) {
        try { winrt::Windows::System::Power::PowerManager::EnergySaverStatusChanged(m_energySaverStatusChangedToken); } catch (...) {}
    }
    if (auto element = m_weakProxyElement.get()) {
        try {
            if (!m_proxyKey.empty()) element.Resources().Remove(winrt::box_value(m_proxyKey));
            if (!m_fallbackProxyKey.empty()) element.Resources().Remove(winrt::box_value(m_fallbackProxyKey));
        } catch (...) {}
    }
}

void XamlBlurBrush::OnConnected() {
    if (!CompositionBrush()) {
        RefreshThemeTint();
        RefreshFallbackColor();
        CompositionBrush(ShouldUseFallback() ? CreateFallbackBrush() : CreateEffectBrush());
    }
}

muc::CompositionBrush XamlBlurBrush::CreateFallbackBrush() {
    return m_compositor.CreateColorBrush(m_fallbackColor.value_or(m_tint));
}

muc::CompositionBrush XamlBlurBrush::CreateEffectBrush() {
    auto backdropBrush = m_compositor.CreateBackdropBrush();
    constexpr float kLumaR = 0.2126f;
    constexpr float kLumaG = 0.7152f;
    constexpr float kLumaB = 0.0722f;

    auto blurEffect = winrt::make_self<GaussianBlurEffect>();
    blurEffect->Source = muc::CompositionEffectSourceParameter(L"backdrop");
    blurEffect->BlurAmount = m_blurAmount;
    blurEffect->Name(L"BlurEffect");

    wge::IGraphicsEffectSource topOfStack = *blurEffect;

    if (m_tintSaturation && *m_tintSaturation != 1.0f) {
        float s = std::max(*m_tintSaturation, 0.0f);
        float invS = 1.0f - s;
        auto satMatrix = winrt::make_self<ColorMatrixEffect>();
        satMatrix->Source = topOfStack;
        auto& m = satMatrix->Matrix;
        m[0]  = invS * kLumaR + s; m[1]  = invS * kLumaR;     m[2]  = invS * kLumaR;     m[3]  = 0.0f;
        m[4]  = invS * kLumaG;     m[5]  = invS * kLumaG + s; m[6]  = invS * kLumaG;     m[7]  = 0.0f;
        m[8]  = invS * kLumaB;     m[9]  = invS * kLumaB;     m[10] = invS * kLumaB + s; m[11] = 0.0f;
        m[12] = 0.0f;              m[13] = 0.0f;              m[14] = 0.0f;              m[15] = 1.0f;
        satMatrix->Name(L"SaturationEffect");
        topOfStack = *satMatrix;
    }

    if (m_tintLuminosityOpacity && *m_tintLuminosityOpacity > 0.0f) {
        float op = std::clamp(*m_tintLuminosityOpacity, 0.0f, 1.0f);
        float tintLum = (m_tint.R / 255.0f) * kLumaR + (m_tint.G / 255.0f) * kLumaG + (m_tint.B / 255.0f) * kLumaB;
        auto lumMatrix = winrt::make_self<ColorMatrixEffect>();
        lumMatrix->Source = topOfStack;
        auto& m = lumMatrix->Matrix;
        m[0]  = 1.0f - (kLumaR * op); m[1]  = -(kLumaR * op);       m[2]  = -(kLumaR * op);       m[3]  = 0.0f;
        m[4]  = -(kLumaG * op);       m[5]  = 1.0f - (kLumaG * op); m[6]  = -(kLumaG * op);       m[7]  = 0.0f;
        m[8]  = -(kLumaB * op);       m[9]  = -(kLumaB * op);       m[10] = 1.0f - (kLumaB * op); m[11] = 0.0f;
        m[12] = 0.0f;                 m[13] = 0.0f;                 m[14] = 0.0f;                 m[15] = 1.0f;
        m[16] = tintLum * op;         m[17] = tintLum * op;         m[18] = tintLum * op;         m[19] = 0.0f;
        lumMatrix->Name(L"LuminosityBlend");
        topOfStack = *lumMatrix;
    }

    muc::CompositionSurfaceBrush noiseBrush{nullptr};
    if (m_noiseOpacity && *m_noiseOpacity > 0.0f) {
        float density = m_noiseDensity.value_or(1.0f);
        auto stream = CreateNoiseStream(density);
        auto surface = Media::LoadedImageSurface::StartLoadFromStream(stream);
        noiseBrush = m_compositor.CreateSurfaceBrush(surface);
        noiseBrush.Stretch(muc::CompositionStretch::None);

        auto borderEffect = winrt::make_self<BorderEffect>();
        borderEffect->Source = muc::CompositionEffectSourceParameter(L"NoiseSource");
        float nOp = std::clamp(*m_noiseOpacity, 0.0f, 1.0f);
        auto opacityEffect = winrt::make_self<ColorMatrixEffect>();
        opacityEffect->Source = *borderEffect;
        opacityEffect->Matrix[0] = nOp; opacityEffect->Matrix[5] = nOp; opacityEffect->Matrix[10] = nOp; opacityEffect->Matrix[15] = nOp;
        opacityEffect->Name(L"NoiseOpacityEffect");

        auto noiseComposite = winrt::make_self<CompositeEffect>();
        noiseComposite->Mode = D2D1_COMPOSITE_MODE_SOURCE_OVER;
        noiseComposite->Sources.push_back(topOfStack);
        noiseComposite->Sources.push_back(*opacityEffect);
        noiseComposite->Name(L"NoiseComposite");
        topOfStack = *noiseComposite;
    }

    auto floodEffect = winrt::make_self<FloodEffect>();
    floodEffect->Color = m_tint;
    floodEffect->Name(L"FloodEffect");

    auto compositeEffect = winrt::make_self<CompositeEffect>();
    compositeEffect->Mode = D2D1_COMPOSITE_MODE_SOURCE_OVER;
    compositeEffect->Sources.push_back(topOfStack);
    compositeEffect->Sources.push_back(*floodEffect);

    auto factory = m_compositor.CreateEffectFactory(*compositeEffect);
    auto brush = factory.CreateBrush();
    brush.SetSourceParameter(L"backdrop", backdropBrush);
    if (noiseBrush) brush.SetSourceParameter(L"NoiseSource", noiseBrush);

    return brush;
}

void XamlBlurBrush::OnDisconnected() {
    if (const auto brush = CompositionBrush()) {
        brush.Close();
        CompositionBrush(nullptr);
    }
}

void XamlBlurBrush::RefreshThemeTint() {
    if (!m_proxyBrush) return;
    m_tint = m_proxyBrush.Color();
    if (m_tintOpacity) m_tint.A = *m_tintOpacity;
}

void XamlBlurBrush::RefreshFallbackColor() {
    if (!m_fallbackProxyBrush) return;
    m_fallbackColor = m_fallbackProxyBrush.Color();
}

bool XamlBlurBrush::ShouldUseFallback() const {
    if (!m_fallbackColor && m_fallbackThemeResourceKey.empty()) return false;
    bool energySaverActive = false;
    HKEY key{};
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SYSTEM\\CurrentControlSet\\Control\\Power", 0, KEY_QUERY_VALUE, &key) == ERROR_SUCCESS) {
        DWORD value = 0, type = 0, size = sizeof(value);
        if (RegQueryValueExW(key, L"EnergySaverState", nullptr, &type, reinterpret_cast<LPBYTE>(&value), &size) == ERROR_SUCCESS && type == REG_DWORD) {
            energySaverActive = (value == 1);
        }
        RegCloseKey(key);
    }
    if (!energySaverActive) {
        SYSTEM_POWER_STATUS powerStatus{};
        if (GetSystemPowerStatus(&powerStatus) && powerStatus.SystemStatusFlag != 0) energySaverActive = true;
    }
    bool advancedEffectsOff = false;
    if (m_uiSettings) {
        try { advancedEffectsOff = !m_uiSettings.AdvancedEffectsEnabled(); } catch (...) {}
    }
    return energySaverActive || advancedEffectsOff;
}

void XamlBlurBrush::RefreshBrush() {
    if (const auto brush = CompositionBrush()) {
        brush.Close();
        CompositionBrush(nullptr);
        OnConnected();
    }
}

////////////////////////////////////////////////////////////////////////////////
// Helper functions for tracking and retrying failed ImageBrush loads.
bool HasInternetAccess() {
    try {
        auto profile = winrt::Windows::Networking::Connectivity::NetworkInformation::GetInternetConnectionProfile();
        return profile && profile.GetNetworkConnectivityLevel() == winrt::Windows::Networking::Connectivity::NetworkConnectivityLevel::InternetAccess;
    } catch (...) { return true; }
}

void StartImageBrushRetry(const std::shared_ptr<TrackedImageBrush>& tracked) {
    auto brush = tracked->brush.get();
    if (!brush) return;
    tracked->lastRetryTick = GetTickCount64();
    tracked->retryCount++;
    try {
        Media::Imaging::BitmapImage retryImage;
        retryImage.CreateOptions(tracked->createOptions | Media::Imaging::BitmapCreateOptions::IgnoreImageCache);
        retryImage.DecodePixelType(tracked->decodePixelType);
        retryImage.DecodePixelWidth(tracked->decodePixelWidth);
        retryImage.DecodePixelHeight(tracked->decodePixelHeight);
        retryImage.AutoPlay(tracked->autoPlay);
        retryImage.UriSource(tracked->uri);
        brush.ImageSource(retryImage);
    } catch (...) {}
}

void RetryFailedImageLoadsOnCurrentThread() {
    if (!g_initializedForThread) return;
    auto& brushes = g_trackedImageBrushesForThread.brushes;
    std::erase_if(brushes, [](const auto& tracked) { return !tracked->brush.get(); });
    std::vector<std::shared_ptr<TrackedImageBrush>> snapshot(brushes.begin(), brushes.end());
    ULONGLONG tick = GetTickCount64();

    for (const auto& tracked : snapshot) {
        if (tracked->loaded || tracked->retryCount >= kImageRetryMaxCount) continue;
        if (tracked->lastRetryTick) {
            ULONGLONG delay = kImageRetryBaseDelayMs << std::clamp(tracked->retryCount - 1, 0, kImageRetryMaxBackoffShift);
            if (tick - tracked->lastRetryTick < delay) continue;
        }
        StartImageBrushRetry(tracked);
    }
}

void ScheduleImageLoadRetryOnCurrentThread() {
    if (!g_initializedForThread) return;
    auto& timer = g_trackedImageBrushesForThread.retryDebounceTimer;
    try {
        if (!timer) {
            auto dispatcher = g_trackedImageBrushesForThread.dispatcher;
            if (!dispatcher) return;
            timer = dispatcher.CreateTimer();
            timer.Interval(std::chrono::milliseconds{kNetworkChangeDebounceMs});
            timer.IsRepeating(false);
            g_trackedImageBrushesForThread.retryDebounceTimerTickRevoker = timer.Tick(
                winrt::auto_revoke,
                [](auto&&, auto&&) { RetryFailedImageLoadsOnCurrentThread(); });
        }
        timer.Stop();
        timer.Start();
    } catch (...) {}
}

void ScheduleImageLoadRetryOnAllUiThreads() {
    if (!HasInternetAccess()) return;
    std::vector<winrt::Microsoft::UI::Dispatching::DispatcherQueue> dispatchers;
    {
        std::lock_guard<std::mutex> lock(g_imageRetryMutex);
        if (!g_imageRetryActive) return;
        for (auto& weakDispatcher : g_imageRetryDispatchers) {
            if (auto dispatcher = weakDispatcher.get()) dispatchers.push_back(dispatcher);
        }
        std::erase_if(g_imageRetryDispatchers, [](const auto& weakDispatcher) { return !weakDispatcher.get(); });
    }
    for (auto& dispatcher : dispatchers) {
        try { dispatcher.TryEnqueue([]() { ScheduleImageLoadRetryOnCurrentThread(); }); } catch (...) {}
    }
}

void OnNetworkStatusChanged(winrt::Windows::Foundation::IInspectable const&) {
    ScheduleImageLoadRetryOnAllUiThreads();
}

void StopImageLoadRetries() {
    std::lock_guard<std::mutex> lock(g_imageRetryMutex);
    g_imageRetryActive = false;
    if (g_networkStatusChangedToken) {
        try {
            winrt::Windows::Networking::Connectivity::NetworkInformation::NetworkStatusChanged(g_networkStatusChangedToken);
        } catch (...) {}
        g_networkStatusChangedToken = {};
    }
    g_imageRetryDispatchers.clear();
}

void TrackImageBrushIfRemoteSource(
    Media::ImageBrush const& brush,
    winrt::Windows::Foundation::IInspectable const& imageSource) {
    auto bitmapImage = imageSource.try_as<Media::Imaging::BitmapImage>();
    if (!bitmapImage) return;
    auto uri = bitmapImage.UriSource();
    if (!uri) return;
    auto scheme = uri.SchemeName();
    if (scheme != L"http" && scheme != L"https") return;

    auto& brushes = g_trackedImageBrushesForThread.brushes;
    std::erase_if(brushes, [](const auto& tracked) { return !tracked->brush.get(); });
    auto it = std::find_if(brushes.begin(), brushes.end(), [&brush](const auto& tracked) {
        if (auto trackedBrush = tracked->brush.get()) return trackedBrush == brush;
        return false;
    });

    if (it != brushes.end()) {
        if ((*it)->uri.Equals(uri)) return;
        brushes.erase(it);
    }

    auto tracked = std::make_shared<TrackedImageBrush>();
    tracked->brush = winrt::make_weak(brush);
    tracked->uri = uri;

    try {
        tracked->decodePixelWidth = bitmapImage.DecodePixelWidth();
        tracked->decodePixelHeight = bitmapImage.DecodePixelHeight();
        tracked->decodePixelType = bitmapImage.DecodePixelType();
        tracked->createOptions = bitmapImage.CreateOptions();
        tracked->autoPlay = bitmapImage.AutoPlay();
        tracked->loaded = bitmapImage.PixelWidth() != 0;
    } catch (...) {}

    std::weak_ptr<TrackedImageBrush> trackedWeak = tracked;
    tracked->imageFailedRevoker = brush.ImageFailed(
        winrt::auto_revoke, [trackedWeak](auto&&, auto&&) {
            if (auto tracked = trackedWeak.lock()) tracked->loaded = false;
        });
    tracked->imageOpenedRevoker = brush.ImageOpened(
        winrt::auto_revoke, [trackedWeak](auto&&, auto&&) {
            if (auto tracked = trackedWeak.lock()) { tracked->loaded = true; tracked->retryCount = 0; tracked->lastRetryTick = 0; }
        });

    brushes.push_back(std::move(tracked));

    std::lock_guard<std::mutex> lock(g_imageRetryMutex);
    g_imageRetryActive = true;

    if (!g_trackedImageBrushesForThread.dispatcher) {
        try {
            if (auto dispatcher = winrt::Microsoft::UI::Dispatching::DispatcherQueue::GetForCurrentThread()) {
                g_trackedImageBrushesForThread.dispatcher = dispatcher;
                g_imageRetryDispatchers.push_back(winrt::make_weak(dispatcher));
            }
        } catch (...) {}
    }

    if (!g_networkStatusChangedToken) {
        try {
            g_networkStatusChangedToken = winrt::Windows::Networking::Connectivity::NetworkInformation::NetworkStatusChanged(OnNetworkStatusChanged);
        } catch (...) {}
    }
}

void SetOrClearValue(DependencyObject elementDo,
                     DependencyProperty property,
                     const PropertyOverrideValue& overrideValue,
                     bool initialApply = false) {
    winrt::Windows::Foundation::IInspectable value;
    if (auto* inspectable = std::get_if<winrt::Windows::Foundation::IInspectable>(&overrideValue)) {
        value = *inspectable;
    } else if (auto* blurBrushParams = std::get_if<XamlBlurBrushParams>(&overrideValue)) {
        if (auto uiElement = elementDo.try_as<UIElement>()) {
            value = winrt::make<XamlBlurBrush>(
                uiElement, blurBrushParams->blurAmount, blurBrushParams->tint,
                blurBrushParams->tintOpacity, winrt::hstring(blurBrushParams->tintThemeResourceKey),
                blurBrushParams->tintLuminosityOpacity, blurBrushParams->tintSaturation,
                blurBrushParams->noiseOpacity, blurBrushParams->noiseDensity,
                blurBrushParams->fallbackColor, winrt::hstring(blurBrushParams->fallbackThemeResourceKey));
        } else return;
    } else return;

    if (value == DependencyProperty::UnsetValue()) {
        try { elementDo.ClearValue(property); } catch (...) {}
        return;
    }

    if (auto imageBrush = value.try_as<Media::ImageBrush>()) {
        TrackImageBrushIfRemoteSource(imageBrush, imageBrush.ImageSource());
    } else if (auto imageBrush = elementDo.try_as<Media::ImageBrush>()) {
        if (property == Media::ImageBrush::ImageSourceProperty()) {
            TrackImageBrushIfRemoteSource(imageBrush, value);
        }
    }

    try {
        if (property == Controls::TextBlock::FontWeightProperty() ||
            property == Controls::Control::FontWeightProperty() ||
            property == Controls::RichTextBlock::FontWeightProperty() ||
            property == Controls::FontIcon::FontWeightProperty() ||
            property == Controls::FontIconSource::FontWeightProperty() ||
            property == Controls::ContentPresenter::FontWeightProperty()) {
            auto valueInt = value.try_as<int>();
            if (valueInt && *valueInt >= std::numeric_limits<uint16_t>::min() &&
                *valueInt <= std::numeric_limits<uint16_t>::max()) {
                value = winrt::box_value(winrt::Windows::UI::Text::FontWeight{static_cast<uint16_t>(*valueInt)});
            }
        }

        Controls::Grid definitionsCloneOwner{nullptr};
        if (auto sourceColumns = value.try_as<Controls::ColumnDefinitionCollection>()) {
            definitionsCloneOwner = Controls::Grid{};
            auto clonedColumns = definitionsCloneOwner.ColumnDefinitions();
            for (auto const& column : sourceColumns) {
                Controls::ColumnDefinition clonedColumn;
                clonedColumn.Width(column.Width());
                clonedColumn.MinWidth(column.MinWidth());
                clonedColumn.MaxWidth(column.MaxWidth());
                clonedColumns.Append(clonedColumn);
            }
            value = clonedColumns;
        } else if (auto sourceRows = value.try_as<Controls::RowDefinitionCollection>()) {
            definitionsCloneOwner = Controls::Grid{};
            auto clonedRows = definitionsCloneOwner.RowDefinitions();
            for (auto const& row : sourceRows) {
                Controls::RowDefinition clonedRow;
                clonedRow.Height(row.Height());
                clonedRow.MinHeight(row.MinHeight());
                clonedRow.MaxHeight(row.MaxHeight());
                clonedRows.Append(clonedRow);
            }
            value = clonedRows;
        }

        elementDo.SetValue(property, value);
    } catch (...) {}
}

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

std::vector<std::wstring_view> SplitStringView(std::wstring_view s, std::wstring_view delimiter) {
    size_t pos_start = 0, pos_end, delim_len = delimiter.length();
    std::wstring_view token;
    std::vector<std::wstring_view> res;

    while ((pos_end = s.find(delimiter, pos_start)) != std::wstring_view::npos) {
        token = s.substr(pos_start, pos_end - pos_start);
        pos_start = pos_end + delim_len;
        res.push_back(token);
    }

    res.push_back(s.substr(pos_start));
    return res;
}

std::optional<PropertyOverrideValue> ParseNonXamlPropertyOverrideValue(std::wstring_view stringValue) {
    auto substr = TrimStringView(stringValue);
    constexpr auto kWindhawkBlurPrefix = L"<WindhawkBlur "sv;
    if (!substr.starts_with(kWindhawkBlurPrefix)) return std::nullopt;
    substr = substr.substr(std::size(kWindhawkBlurPrefix));

    constexpr auto kWindhawkBlurSuffix = L"/>"sv;
    if (!substr.ends_with(kWindhawkBlurSuffix)) throw std::runtime_error("WindhawkBlur: Bad suffix");
    substr = substr.substr(0, substr.size() - std::size(kWindhawkBlurSuffix));

    bool pendingTintColorThemeResource = false;
    bool pendingFallbackColorThemeResource = false;
    std::wstring tintThemeResourceKey;
    std::wstring fallbackThemeResourceKey;
    winrt::Windows::UI::Color tint{};
    std::optional<winrt::Windows::UI::Color> fallbackColor;
    float tintOpacity = std::numeric_limits<float>::quiet_NaN();
    float tintLuminosityOpacity = std::numeric_limits<float>::quiet_NaN();
    float tintSaturation = std::numeric_limits<float>::quiet_NaN();
    float noiseOpacity = std::numeric_limits<float>::quiet_NaN();
    float noiseDensity = std::numeric_limits<float>::quiet_NaN();
    float blurAmount = 0;

    constexpr auto kTintColorThemeResourcePrefix = L"TintColor=\"{ThemeResource"sv;
    constexpr auto kTintColorThemeResourceSuffix = L"}\""sv;
    constexpr auto kTintColorPrefix = L"TintColor=\"#"sv;
    constexpr auto kTintOpacityPrefix = L"TintOpacity=\""sv;
    constexpr auto kTintLuminosityOpacityPrefix = L"TintLuminosityOpacity=\""sv;
    constexpr auto kTintSaturationPrefix = L"TintSaturation=\""sv;
    constexpr auto kNoiseOpacityPrefix = L"NoiseOpacity=\""sv;
    constexpr auto kNoiseDensityPrefix = L"NoiseDensity=\""sv;
    constexpr auto kBlurAmountPrefix = L"BlurAmount=\""sv;
    constexpr auto kFallbackColorThemeResourcePrefix = L"FallbackColor=\"{ThemeResource"sv;
    constexpr auto kFallbackColorThemeResourceSuffix = L"}\""sv;
    constexpr auto kFallbackColorPrefix = L"FallbackColor=\"#"sv;

    for (const auto prop : SplitStringView(substr, L" ")) {
        const auto propSubstr = TrimStringView(prop);
        if (propSubstr.empty()) continue;

        if (pendingTintColorThemeResource) {
            if (!propSubstr.ends_with(kTintColorThemeResourceSuffix)) throw std::runtime_error("WindhawkBlur: Invalid TintColor theme resource syntax");
            pendingTintColorThemeResource = false;
            tintThemeResourceKey = propSubstr.substr(0, propSubstr.size() - std::size(kTintColorThemeResourceSuffix));
            continue;
        }

        if (pendingFallbackColorThemeResource) {
            if (!propSubstr.ends_with(kFallbackColorThemeResourceSuffix)) throw std::runtime_error("WindhawkBlur: Invalid FallbackColor theme resource syntax");
            pendingFallbackColorThemeResource = false;
            fallbackThemeResourceKey = propSubstr.substr(0, propSubstr.size() - std::size(kFallbackColorThemeResourceSuffix));
            continue;
        }

        if (propSubstr == kTintColorThemeResourcePrefix) { pendingTintColorThemeResource = true; continue; }
        if (propSubstr == kFallbackColorThemeResourcePrefix) { pendingFallbackColorThemeResource = true; continue; }

        if (propSubstr.starts_with(kTintColorPrefix) && propSubstr.back() == L'\"') {
            auto valStr = propSubstr.substr(std::size(kTintColorPrefix), propSubstr.size() - std::size(kTintColorPrefix) - 1);
            bool hasAlpha = (valStr.size() == 8);
            auto valNum = std::stoul(std::wstring(valStr), nullptr, 16);
            uint8_t a = hasAlpha ? HIBYTE(HIWORD(valNum)) : 255;
            uint8_t r = LOBYTE(HIWORD(valNum));
            uint8_t g = HIBYTE(LOWORD(valNum));
            uint8_t b = LOBYTE(LOWORD(valNum));
            tint = {a, r, g, b};
            continue;
        }

        if (propSubstr.starts_with(kFallbackColorPrefix) && propSubstr.back() == L'\"') {
            auto valStr = propSubstr.substr(std::size(kFallbackColorPrefix), propSubstr.size() - std::size(kFallbackColorPrefix) - 1);
            bool hasAlpha = (valStr.size() == 8);
            auto valNum = std::stoul(std::wstring(valStr), nullptr, 16);
            uint8_t a = hasAlpha ? HIBYTE(HIWORD(valNum)) : 255;
            uint8_t r = LOBYTE(HIWORD(valNum));
            uint8_t g = HIBYTE(LOWORD(valNum));
            uint8_t b = LOBYTE(LOWORD(valNum));
            fallbackColor = winrt::Windows::UI::Color{a, r, g, b};
            continue;
        }

        if (propSubstr.starts_with(kTintOpacityPrefix) && propSubstr.back() == L'\"') {
            tintOpacity = std::stof(std::wstring(propSubstr.substr(std::size(kTintOpacityPrefix), propSubstr.size() - std::size(kTintOpacityPrefix) - 1)));
            continue;
        }

        if (propSubstr.starts_with(kTintLuminosityOpacityPrefix) && propSubstr.back() == L'\"') {
            tintLuminosityOpacity = std::stof(std::wstring(propSubstr.substr(std::size(kTintLuminosityOpacityPrefix), propSubstr.size() - std::size(kTintLuminosityOpacityPrefix) - 1)));
            continue;
        }

        if (propSubstr.starts_with(kTintSaturationPrefix) && propSubstr.back() == L'\"') {
            tintSaturation = std::stof(std::wstring(propSubstr.substr(std::size(kTintSaturationPrefix), propSubstr.size() - std::size(kTintSaturationPrefix) - 1)));
            continue;
        }

        if (propSubstr.starts_with(kNoiseOpacityPrefix) && propSubstr.back() == L'\"') {
            noiseOpacity = std::stof(std::wstring(propSubstr.substr(std::size(kNoiseOpacityPrefix), propSubstr.size() - std::size(kNoiseOpacityPrefix) - 1)));
            continue;
        }

        if (propSubstr.starts_with(kNoiseDensityPrefix) && propSubstr.back() == L'\"') {
            noiseDensity = std::stof(std::wstring(propSubstr.substr(std::size(kNoiseDensityPrefix), propSubstr.size() - std::size(kNoiseDensityPrefix) - 1)));
            continue;
        }

        if (propSubstr.starts_with(kBlurAmountPrefix) && propSubstr.back() == L'\"') {
            blurAmount = std::stof(std::wstring(propSubstr.substr(std::size(kBlurAmountPrefix), propSubstr.size() - std::size(kBlurAmountPrefix) - 1)));
            continue;
        }

        throw std::runtime_error("WindhawkBlur: Bad property");
    }

    if (!std::isnan(tintOpacity)) {
        tintOpacity = std::clamp(tintOpacity, 0.0f, 1.0f);
        tint.A = static_cast<uint8_t>(tintOpacity * 255.0f);
    }

    return XamlBlurBrushParams{
        .blurAmount = blurAmount,
        .tint = tint,
        .tintOpacity = !std::isnan(tintOpacity) ? std::optional(tint.A) : std::nullopt,
        .tintThemeResourceKey = std::move(tintThemeResourceKey),
        .tintLuminosityOpacity = !std::isnan(tintLuminosityOpacity) ? std::optional(tintLuminosityOpacity) : std::nullopt,
        .tintSaturation = !std::isnan(tintSaturation) ? std::optional(tintSaturation) : std::nullopt,
        .noiseOpacity = !std::isnan(noiseOpacity) ? std::optional(noiseOpacity) : std::nullopt,
        .noiseDensity = !std::isnan(noiseDensity) ? std::optional(noiseDensity) : std::nullopt,
        .fallbackColor = fallbackColor,
        .fallbackThemeResourceKey = std::move(fallbackThemeResourceKey),
    };
}

Style GetStyleFromXamlSetters(const std::wstring_view type,
                              const std::wstring_view xamlStyleSetters) {
    std::wstring xaml =
        LR"(<ResourceDictionary
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:muxc="using:Microsoft.UI.Xaml.Controls")";

    if (auto pos = type.rfind('.'); pos != type.npos) {
        auto typeNamespace = std::wstring_view(type).substr(0, pos);
        auto typeName = std::wstring_view(type).substr(pos + 1);

        xaml += L"\n    xmlns:windhawkstyler=\"using:";
        xaml += EscapeXmlAttribute(typeNamespace);
        xaml += L"\">\n    <Style TargetType=\"windhawkstyler:";
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

const ResolvedRules& GetResolvedPropertyOverrides(
    const std::wstring_view type,
    PropertyOverridesMaybeUnresolved* propertyOverridesMaybeUnresolved) {
    if (const auto* resolved = std::get_if<ResolvedRules>(propertyOverridesMaybeUnresolved)) {
        return *resolved;
    }

    ResolvedRules resolved;
    try {
        const auto& unresolved = std::get<UnresolvedRules>(*propertyOverridesMaybeUnresolved);
        const auto& valueRules = unresolved.valueRules;
        const auto& captureRules = unresolved.captureRules;

        if (!valueRules.empty() || !captureRules.empty()) {
            std::wstring xaml;
            std::vector<std::optional<PropertyOverrideValue>> propertyOverrideValues;
            propertyOverrideValues.reserve(valueRules.size());

            for (const auto& rule : valueRules) {
                const bool isDynamic = rule.isDynamic();
                propertyOverrideValues.push_back(!isDynamic && rule.isXamlValue ? ParseNonXamlPropertyOverrideValue(rule.value) : std::nullopt);

                xaml += L"        <Setter Property=\"";
                xaml += EscapeXmlAttribute(rule.propertyName);
                xaml += L"\"";
                if (isDynamic || propertyOverrideValues.back() || (rule.isXamlValue && rule.value.empty())) {
                    xaml += L" Value=\"{x:Null}\" />\n";
                } else if (!rule.isXamlValue) {
                    xaml += L" Value=\"";
                    xaml += EscapeXmlAttribute(rule.value);
                    xaml += L"\" />\n";
                } else {
                    xaml += L">\n            <Setter.Value>\n";
                    xaml += rule.value;
                    xaml += L"\n            </Setter.Value>\n        </Setter>\n";
                }
            }

            for (const auto& rule : captureRules) {
                xaml += L"        <Setter Property=\"";
                xaml += EscapeXmlAttribute(rule.propertyName);
                xaml += L"\" Value=\"{x:Null}\" />\n";
            }

            auto style = GetStyleFromXamlSetters(type, xaml);

            uint32_t setterIndex = 0;
            for (size_t i = 0; i < valueRules.size(); i++, setterIndex++) {
                const auto& rule = valueRules[i];
                const auto setter = style.Setters().GetAt(setterIndex).as<Setter>();
                auto property = setter.Property();
                if (rule.isDynamic()) {
                    resolved.propertyOverrides[property][rule.visualState] = DynamicStyleTemplate{rule.propertyName, rule.value, rule.isXamlValue};
                    resolved.hasDynamicValues = true;
                } else {
                    resolved.propertyOverrides[property][rule.visualState] = propertyOverrideValues[i].value_or(
                        rule.isXamlValue && rule.value.empty() ? DependencyProperty::UnsetValue() : setter.Value());
                }
            }

            for (const auto& rule : captureRules) {
                const auto setter = style.Setters().GetAt(setterIndex++).as<Setter>();
                resolved.captures.push_back({setter.Property(), rule.varName});
            }
        }
    } catch (...) {}

    *propertyOverridesMaybeUnresolved = std::move(resolved);
    return std::get<ResolvedRules>(*propertyOverridesMaybeUnresolved);
}

const PropertyValues& GetResolvedPropertyValues(
    const std::wstring_view type,
    PropertyValuesMaybeUnresolved* propertyValuesMaybeUnresolved) {
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

            auto style = GetStyleFromXamlSetters(type.empty() ? L"FrameworkElement" : type, xaml);
            for (size_t i = 0; i < propertyValuesStr.size(); i++) {
                const auto setter = style.Setters().GetAt(i).as<Setter>();
                propertyValues.push_back({setter.Property(), setter.Value()});
            }
        }
    } catch (winrt::hresult_error const& ex) {
        Wh_Log(L"Error %08X: %s", ex.code(), ex.message().c_str());
        if (ex.code() == (HRESULT)0x802B000A) {
            static const PropertyValues emptyValues;
            return emptyValues;
        }
    } catch (std::exception const& ex) {
        Wh_Log(L"Error: %S", ex.what());
    } catch (...) {}

    *propertyValuesMaybeUnresolved = std::move(propertyValues);
    return std::get<PropertyValues>(*propertyValuesMaybeUnresolved);
}

VisualStateGroup GetVisualStateGroup(FrameworkElement element,
                                     std::wstring_view visualStateGroupName) {
    auto list = VisualStateManager::GetVisualStateGroups(element);
    for (const auto& v : list) {
        if (v.Name() == visualStateGroupName) return v;
    }
    return nullptr;
}

using StyleConstant = std::pair<std::wstring, std::wstring>;
using StyleConstants = std::vector<StyleConstant>;

std::wstring ApplyStyleConstants(std::wstring_view style,
                                 const StyleConstants& styleConstants) {
    std::wstring result;
    size_t lastPos = 0;
    size_t findPos;

    while ((findPos = style.find('$', lastPos)) != style.npos) {
        result.append(style, lastPos, findPos - lastPos);

        const StyleConstant* constant = nullptr;
        for (const auto& s : styleConstants) {
            if (s.first == style.substr(findPos + 1, s.first.size())) {
                constant = &s;
                break;
            }
        }

        if (constant) {
            result += constant->second;
            lastPos = findPos + 1 + constant->first.size();
        } else {
            result += L'$';
            lastPos = findPos + 1;
        }
    }

    result.append(style, lastPos);
    return result;
}

std::optional<StyleConstant> ParseStyleConstant(
    std::wstring_view entry,
    const StyleConstants& existingConstants) {
    if (entry.starts_with(L"//")) return std::nullopt;

    auto eqPos = entry.find(L'=');
    if (eqPos == entry.npos) return std::nullopt;

    auto keyPart = TrimStringView(entry.substr(0, eqPos));
    auto valueRaw = TrimStringView(entry.substr(eqPos + 1));
    auto value = ApplyStyleConstants(valueRaw, existingConstants);

    if (keyPart.empty()) return std::nullopt;
    return StyleConstant{std::wstring(keyPart), std::move(value)};
}

StyleConstants LoadStyleConstants(const std::vector<PCWSTR>& themeStyleConstants) {
    StyleConstants result;
    auto addToResult = [&result](StyleConstant constant) {
        for (auto& s : result) {
            if (s.first == constant.first) {
                s.second = std::move(constant.second);
                return;
            }
        }
        result.push_back(std::move(constant));
    };

    for (const auto themeStyleConstant : themeStyleConstants) {
        if (auto parsed = ParseStyleConstant(themeStyleConstant, result)) {
            addToResult(std::move(*parsed));
        }
    }

    for (int i = 0;; i++) {
        string_setting_unique_ptr constantSetting(
            Wh_GetStringSetting(L"styleConstants[%d]", i));
        if (!*constantSetting.get()) break;

        if (auto parsed = ParseStyleConstant(constantSetting.get(), result)) {
            addToResult(std::move(*parsed));
        }
    }

    return result;
}

ElementMatcher ElementMatcherFromString(std::wstring_view str) {
    ElementMatcher result;
    PropertyValuesUnresolved propertyValuesUnresolved;

    auto trimmed = TrimStringView(str);
    if (trimmed == L"*") {
        result.kind = ElementMatcher::Kind::Wildcard;
        return result;
    }
    if (trimmed == L":root") {
        result.kind = ElementMatcher::Kind::Root;
        return result;
    }

    auto i = str.find_first_of(L"#@[");
    result.type = TrimStringView(str.substr(0, i));
    if (result.type.empty()) {
        throw std::runtime_error("Bad target syntax, empty type");
    }

    while (i != str.npos) {
        auto iNext = str.find_first_of(L"#@[", i + 1);
        auto nextPart = str.substr(i + 1, iNext == str.npos ? str.npos : iNext - (i + 1));

        switch (str[i]) {
            case L'#':
                result.name = TrimStringView(nextPart);
                break;
            case L'@':
                result.visualStateGroupName = TrimStringView(nextPart);
                break;
            case L'[': {
                auto rule = TrimStringView(nextPart);
                if (rule.length() > 0 && rule.back() == L']') {
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

std::variant<ValueRule, CaptureRule> ParseRule(std::wstring_view str) {
    auto eqPos = str.find(L'=');
    if (eqPos == str.npos) {
        throw std::runtime_error("Bad style syntax, '=' is missing");
    }

    auto name = str.substr(0, eqPos);
    auto value = str.substr(eqPos + 1);

    if (!value.empty() && value.front() == L'>') {
        value = value.substr(1);
        auto trimmedPropertyName = TrimStringView(name);
        auto trimmedVarName = TrimStringView(value);
        return CaptureRule{std::wstring(trimmedPropertyName), std::wstring(trimmedVarName)};
    }

    ValueRule result;
    result.value = TrimStringView(value);

    if (!name.empty() && name.back() == L':') {
        result.isXamlValue = true;
        name = name.substr(0, name.size() - 1);
    }

    auto atPos = name.find(L'@');
    if (atPos != name.npos) {
        result.visualState = TrimStringView(name.substr(atPos + 1));
        name = name.substr(0, atPos);
    }

    result.propertyName = TrimStringView(name);
    return result;
}

std::wstring AdjustTypeName(std::wstring_view type) {
    if (type.find_first_of(L".:") == type.npos) {
        if (type == L"Rectangle") return L"Microsoft.UI.Xaml.Shapes.Rectangle";
        return L"Microsoft.UI.Xaml.Controls." + std::wstring{type};
    }
    if (type.starts_with(L"muxc:")) {
        return L"Microsoft.UI.Xaml.Controls." + std::wstring{type.substr(5)};
    }
    return std::wstring{type};
}

void AddElementCustomizationRules(std::wstring_view target,
                                  std::vector<std::wstring> styles) {
    ElementCustomizationRules elementCustomizationRules;
    auto targetParts = SplitStringView(target, L" > ");

    bool first = true;
    for (auto i = targetParts.rbegin(); i != targetParts.rend(); ++i) {
        const auto& targetPart = *i;
        auto matcher = ElementMatcherFromString(targetPart);

        if (matcher.kind == ElementMatcher::Kind::Element) {
            matcher.type = AdjustTypeName(matcher.type);
        }

        if (first) {
            UnresolvedRules unresolvedRules;
            for (const auto& style : styles) {
                auto parsed = ParseRule(style);
                if (auto* valueRule = std::get_if<ValueRule>(&parsed)) {
                    unresolvedRules.valueRules.push_back(std::move(*valueRule));
                } else {
                    unresolvedRules.captureRules.push_back(std::move(std::get<CaptureRule>(parsed)));
                }
            }
            elementCustomizationRules.elementMatcher = std::move(matcher);
            elementCustomizationRules.propertyOverrides = std::move(unresolvedRules);
        } else {
            elementCustomizationRules.parentElementMatchers.push_back(std::move(matcher));
        }

        first = false;
    }

    g_elementsCustomizationRules.push_back(std::move(elementCustomizationRules));
}

bool ProcessSingleTargetStylesFromSettings(
    int index,
    const StyleConstants& styleConstants) {
    string_setting_unique_ptr targetStringSetting(
        Wh_GetStringSetting(L"controlStyles[%d].target", index));
    if (!*targetStringSetting.get()) {
        return false;
    }

    if (targetStringSetting[0] == L'/' && targetStringSetting[1] == L'/') {
        return true;
    }

    Wh_Log(L"Processing styles for %s", targetStringSetting.get());
    std::vector<std::wstring> styles;

    for (int styleIndex = 0;; styleIndex++) {
        string_setting_unique_ptr styleSetting(Wh_GetStringSetting(
            L"controlStyles[%d].styles[%d]", index, styleIndex));
        if (!*styleSetting.get()) {
            break;
        }

        if (styleSetting[0] == L'/' && styleSetting[1] == L'/') {
            continue;
        }

        styles.push_back(ApplyStyleConstants(styleSetting.get(), styleConstants));
    }

    if (styles.size() > 0) {
        AddElementCustomizationRules(targetStringSetting.get(), std::move(styles));
    }

    return true;
}

winrt::Windows::Foundation::IInspectable ParseXamlValue(std::wstring_view xamlValue) {
    std::wstring xaml = L"        <Setter Property=\"Tag\">\n            <Setter.Value>\n" +
                        std::wstring(xamlValue) +
                        L"\n            </Setter.Value>\n        </Setter>\n";
    auto style = GetStyleFromXamlSetters(L"FrameworkElement", xaml);
    return style.Setters().GetAt(0).as<Setter>().Value();
}

std::optional<ResourceVariableEntry> ParseResourceVariable(
    std::wstring_view entry,
    const StyleConstants& styleConstants) {
    if (entry.starts_with(L"//")) return std::nullopt;

    auto eqPos = entry.find(L'=');
    if (eqPos == entry.npos) return std::nullopt;

    auto keyPart = TrimStringView(entry.substr(0, eqPos));
    auto valueRaw = TrimStringView(entry.substr(eqPos + 1));
    auto value = ApplyStyleConstants(valueRaw, styleConstants);

    constexpr std::wstring_view kThemeResourcePrefix = L"{ThemeResource ";

    ResourceVariableType type = ResourceVariableType::String;
    if (keyPart.size() > 0 && keyPart.back() == L':') {
        type = ResourceVariableType::Xaml;
        keyPart = TrimStringView(keyPart.substr(0, keyPart.size() - 1));
    } else if (value.starts_with(kThemeResourcePrefix) && value.ends_with(L"}")) {
        type = ResourceVariableType::ThemeResourceReference;
        value = TrimStringView(value.substr(kThemeResourcePrefix.size(), value.size() - kThemeResourcePrefix.size() - 1));
    }

    ResourceVariableTheme theme = ResourceVariableTheme::None;
    std::wstring key;

    auto atPos = keyPart.find(L'@');
    if (atPos != keyPart.npos) {
        key = TrimStringView(keyPart.substr(0, atPos));
        auto themePart = TrimStringView(keyPart.substr(atPos + 1));
        if (themePart == L"Dark") theme = ResourceVariableTheme::Dark;
        else if (themePart == L"Light") theme = ResourceVariableTheme::Light;
    } else {
        key = std::wstring(keyPart);
    }

    return ResourceVariableEntry{std::move(key), std::move(value), theme, type};
}

std::vector<ResourceVariableEntry> ProcessResourceVariablesFromSettings(
    const StyleConstants& styleConstants,
    const std::vector<PCWSTR>& themeResourceVariables) {
    std::vector<ResourceVariableEntry> resourceVariables;

    for (int i = 0;; i++) {
        string_setting_unique_ptr setting(
            Wh_GetStringSetting(L"themeResourceVariables[%d]", i));
        if (!*setting.get()) break;

        auto parsed = ParseResourceVariable(setting.get(), styleConstants);
        if (parsed) resourceVariables.push_back(std::move(*parsed));
    }

    return resourceVariables;
}

bool ProcessResourceVariable(ResourceDictionary resources,
                             ResourceDictionary darkDict,
                             ResourceDictionary lightDict,
                             const ResourceVariableEntry& entry) {
    auto boxedKey = winrt::box_value(entry.key);

    if (entry.theme != ResourceVariableTheme::None) {
        ResourceDictionary& targetDict = entry.theme == ResourceVariableTheme::Dark ? darkDict : lightDict;
        if (targetDict.HasKey(boxedKey)) return false;

        winrt::Windows::Foundation::IInspectable value;
        switch (entry.type) {
            case ResourceVariableType::String: value = winrt::box_value(entry.value); break;
            case ResourceVariableType::Xaml: value = entry.value.empty() ? nullptr : ParseXamlValue(entry.value); break;
            case ResourceVariableType::ThemeResourceReference: value = resources.Lookup(winrt::box_value(entry.value)); break;
        }
        targetDict.Insert(boxedKey, value);
        return true;
    }

    auto existingResource = resources.TryLookup(boxedKey);
    if (!existingResource) return false;

    auto [it, inserted] = g_originalResourceValues.try_emplace(entry.key, existingResource);
    if (!inserted) return false;

    winrt::Windows::Foundation::IInspectable value;
    switch (entry.type) {
        case ResourceVariableType::String: {
            auto resourceClassName = winrt::get_class_name(existingResource);
            if (resourceClassName.starts_with(L"Windows.Foundation.IReference`1<") && resourceClassName.ends_with(L'>')) {
                size_t prefixSize = sizeof("Windows.Foundation.IReference`1<") - 1;
                resourceClassName = winrt::hstring(resourceClassName.data() + prefixSize, resourceClassName.size() - prefixSize - 1);
            }
            value = Markup::XamlBindingHelper::ConvertValue(
                winrt::Windows::UI::Xaml::Interop::TypeName{resourceClassName}, winrt::box_value(entry.value));
            break;
        }
        case ResourceVariableType::Xaml: value = entry.value.empty() ? nullptr : ParseXamlValue(entry.value); break;
        case ResourceVariableType::ThemeResourceReference: value = resources.Lookup(winrt::box_value(entry.value)); break;
    }

    resources.Insert(boxedKey, value);
    return true;
}

void RefreshThemeResourceEntries() {
    if (g_resourceVariables.empty()) return;
    Wh_Log(L"Refreshing theme resource entries");

    auto resources = Application::Current().Resources();
    auto darkDict = g_resourceVariablesThemeDict ? g_resourceVariablesThemeDict.ThemeDictionaries().TryLookup(winrt::box_value(L"Dark")).try_as<ResourceDictionary>() : nullptr;
    auto lightDict = g_resourceVariablesThemeDict ? g_resourceVariablesThemeDict.ThemeDictionaries().TryLookup(winrt::box_value(L"Light")).try_as<ResourceDictionary>() : nullptr;

    for (const auto& entry : g_resourceVariables) {
        if (entry.type != ResourceVariableType::ThemeResourceReference) continue;
        try {
            auto boxedKey = winrt::box_value(entry.key);
            auto value = resources.Lookup(winrt::box_value(entry.value));
            if (entry.theme == ResourceVariableTheme::Dark && darkDict) {
                darkDict.Insert(boxedKey, value);
            } else if (entry.theme == ResourceVariableTheme::Light && lightDict) {
                lightDict.Insert(boxedKey, value);
            } else {
                resources.Insert(boxedKey, value);
            }
        } catch (...) {}
    }
}

void MergeResourceVariables() {
    auto resources = Application::Current().Resources();
    g_resourceVariablesThemeDict = ResourceDictionary();
    ResourceDictionary darkDict;
    ResourceDictionary lightDict;
    bool hasThemeResources = false;
    bool hasThemeResourceReferences = false;

    for (auto it = g_resourceVariables.rbegin(); it != g_resourceVariables.rend(); ++it) {
        try {
            if (ProcessResourceVariable(resources, darkDict, lightDict, *it)) {
                if (it->theme != ResourceVariableTheme::None) hasThemeResources = true;
                if (it->type == ResourceVariableType::ThemeResourceReference) hasThemeResourceReferences = true;
            }
        } catch (...) {}
    }

    if (hasThemeResources) {
        g_resourceVariablesThemeDict.ThemeDictionaries().Insert(winrt::box_value(L"Dark"), darkDict);
        g_resourceVariablesThemeDict.ThemeDictionaries().Insert(winrt::box_value(L"Light"), lightDict);
        resources.MergedDictionaries().Append(g_resourceVariablesThemeDict);
    }

    if (hasThemeResourceReferences) {
        g_uiSettings = winrt::Windows::UI::ViewManagement::UISettings();
        auto dispatcherQueue = winrt::Microsoft::UI::Dispatching::DispatcherQueue::GetForCurrentThread();
        if (dispatcherQueue && g_uiSettings) {
            g_colorValuesChangedToken = g_uiSettings.ColorValuesChanged([dispatcherQueue](auto&&, auto&&) {
                dispatcherQueue.TryEnqueue(RefreshThemeResourceEntries);
            });
        }
    }
}

void ProcessAllStylesFromSettings() {
    StyleConstants styleConstants = LoadStyleConstants({});
    for (int i = 0;; i++) {
        try {
            if (!ProcessSingleTargetStylesFromSettings(i, styleConstants)) break;
        } catch (...) {}
    }
    g_resourceVariables = ProcessResourceVariablesFromSettings(styleConstants, {});
    BuildRuleIndex();
}

void UninitializeResourceVariables() {
    if (g_colorValuesChangedToken) {
        g_uiSettings.ColorValuesChanged(g_colorValuesChangedToken);
        g_colorValuesChangedToken = {};
    }
    g_uiSettings = nullptr;
    g_resourceVariables.clear();

    if (g_originalResourceValues.empty() && !g_resourceVariablesThemeDict) return;

    auto resources = Application::Current().Resources();
    for (const auto& [key, originalValue] : g_originalResourceValues) {
        try { resources.Insert(winrt::box_value(key), originalValue); } catch (...) {}
    }
    g_originalResourceValues.clear();

    if (g_resourceVariablesThemeDict) {
        auto merged = resources.MergedDictionaries();
        uint32_t index;
        if (merged.IndexOf(g_resourceVariablesThemeDict, index)) merged.RemoveAt(index);
        g_resourceVariablesThemeDict = nullptr;
    }
}

bool TestElementMatcher(FrameworkElement element,
                        ElementMatcher& matcher,
                        VisualStateGroup* visualStateGroup,
                        PCWSTR fallbackClassName) {
    if (!matcher.type.empty()) {
        bool typeMatches = false;
        if (fallbackClassName && matcher.type == fallbackClassName) {
            typeMatches = true;
        } else {
            winrt::hstring runtimeClassName = winrt::get_class_name(element);
            if (matcher.type == runtimeClassName) {
                typeMatches = true;
            }
        }
        if (!typeMatches) {
            return false;
        }
    }

    if (!matcher.name.empty() && matcher.name != element.Name()) {
        return false;
    }

    if (matcher.oneBasedIndex) {
        auto parent = Media::VisualTreeHelper::GetParent(element);
        if (!parent) return false;
        int index = matcher.oneBasedIndex - 1;
        if (index < 0 || index >= Media::VisualTreeHelper::GetChildrenCount(parent) ||
            Media::VisualTreeHelper::GetChild(parent, index) != element) {
            return false;
        }
    }

    const auto& propertyValues =
        GetResolvedPropertyValues(matcher.type, &matcher.propertyValues);
    if (!propertyValues.empty()) {
        auto elementDo = element.as<DependencyObject>();

        for (const auto& propertyValue : propertyValues) {
            const auto value =
                ReadLocalValueWithWorkaround(elementDo, propertyValue.first);
            if (!value) {
                Wh_Log(L"Null property value");
                return false;
            }

            const auto className = winrt::get_class_name(value);
            const auto expectedClassName =
                winrt::get_class_name(propertyValue.second);
            if (className != expectedClassName) {
                Wh_Log(L"Different property class: %s vs. %s", className.c_str(),
                       expectedClassName.c_str());
                return false;
            }

            if (className == L"Windows.Foundation.IReference`1<String>") {
                if (winrt::unbox_value<winrt::hstring>(propertyValue.second) ==
                    winrt::unbox_value<winrt::hstring>(value)) {
                    continue;
                }
                return false;
            }

            if (className == L"Windows.Foundation.IReference`1<Double>") {
                if (winrt::unbox_value<double>(propertyValue.second) ==
                    winrt::unbox_value<double>(value)) {
                    continue;
                }
                return false;
            }

            if (className == L"Windows.Foundation.IReference`1<Boolean>") {
                if (winrt::unbox_value<bool>(propertyValue.second) ==
                    winrt::unbox_value<bool>(value)) {
                    continue;
                }
                return false;
            }

            if (className == L"Windows.Foundation.IReference`1<Int32>") {
                if (winrt::unbox_value<int32_t>(propertyValue.second) ==
                    winrt::unbox_value<int32_t>(value)) {
                    continue;
                }
                return false;
            }

            Wh_Log(L"Unsupported property class: %s", className.c_str());
            return false;
        }
    }

    if (matcher.visualStateGroupName && visualStateGroup) {
        *visualStateGroup = GetVisualStateGroup(element, *matcher.visualStateGroupName);
    }

    return true;
}

struct ElementResolvedRules {
    std::unordered_map<VisualStateGroup, PropertyOverrides> overridesPerVSG;
    std::vector<CaptureSpec> captures;
    bool hasDynamicValues = false;
};

ElementResolvedRules FindElementPropertyOverrides(FrameworkElement element,
                                                  PCWSTR fallbackClassName) {
    ElementResolvedRules result;
    std::unordered_set<DependencyProperty> propertiesAdded;
    std::unordered_set<std::wstring> capturesAdded;

    std::vector<const ElementCustomizationRules*> candidateRules;
    candidateRules.reserve(g_genericRules.size() + 16);
    candidateRules.insert(candidateRules.end(), g_genericRules.begin(), g_genericRules.end());

    auto addRulesForType = [&](std::wstring_view typeName) {
        if (auto it = g_rulesByTypeMap.find(typeName); it != g_rulesByTypeMap.end()) {
            candidateRules.insert(candidateRules.end(), it->second.begin(), it->second.end());
        }
        if (auto pos = typeName.rfind(L'.'); pos != typeName.npos) {
            std::wstring_view shortName = typeName.substr(pos + 1);
            if (auto it = g_rulesByTypeMap.find(shortName); it != g_rulesByTypeMap.end()) {
                candidateRules.insert(candidateRules.end(), it->second.begin(), it->second.end());
            }
        }
    };

    if (fallbackClassName) {
        addRulesForType(fallbackClassName);
    }

    winrt::hstring runtimeClassName;
    if (!fallbackClassName || std::wstring_view(fallbackClassName).find(L"IUIElementOverrides") != std::wstring_view::npos) {
        runtimeClassName = winrt::get_class_name(element);
        addRulesForType(runtimeClassName);
    }

    std::unordered_set<const ElementCustomizationRules*> seenRules;

    for (auto it = candidateRules.rbegin(); it != candidateRules.rend(); ++it) {
        const auto* overridePtr = *it;
        if (!seenRules.insert(overridePtr).second) continue;
        auto& override = *const_cast<ElementCustomizationRules*>(overridePtr);
        VisualStateGroup visualStateGroup = nullptr;

        if (!TestElementMatcher(element, override.elementMatcher, &visualStateGroup, fallbackClassName)) {
            continue;
        }

        auto& parentMatchers = override.parentElementMatchers;
        auto matchParents = [&](auto& self, FrameworkElement iter, size_t mi) -> bool {
            if (mi >= parentMatchers.size()) return true;
            auto& matcher = parentMatchers[mi];

            if (matcher.kind == ElementMatcher::Kind::Root) {
                if (Media::VisualTreeHelper::GetParent(iter)) return false;
                return self(self, iter, mi + 1);
            }

            if (matcher.kind == ElementMatcher::Kind::Wildcard) {
                auto& nextMatcher = parentMatchers[mi + 1];
                auto cur = iter;
                while (true) {
                    auto parent = Media::VisualTreeHelper::GetParent(cur).try_as<FrameworkElement>();
                    if (!parent) return false;
                    cur = parent;
                    if (TestElementMatcher(cur, nextMatcher, &visualStateGroup, nullptr) &&
                        self(self, cur, mi + 2)) {
                        return true;
                    }
                }
            }

            auto parent = Media::VisualTreeHelper::GetParent(iter).try_as<FrameworkElement>();
            if (!parent) return false;

            if (!TestElementMatcher(parent, matcher, &visualStateGroup, nullptr)) {
                return false;
            }

            return self(self, parent, mi + 1);
        };

        if (!matchParents(matchParents, element, 0)) continue;

        const auto& resolvedRules = GetResolvedPropertyOverrides(
            override.elementMatcher.type, &override.propertyOverrides);

        result.hasDynamicValues |= resolvedRules.hasDynamicValues;
        auto& propertyOverridesForVSG = result.overridesPerVSG[visualStateGroup];

        for (const auto& [property, valuesPerVisualState] : resolvedRules.propertyOverrides) {
            if (!propertiesAdded.insert(property).second) continue;
            auto& propertyOverrides = propertyOverridesForVSG[property];
            for (const auto& [visualState, value] : valuesPerVisualState) {
                propertyOverrides.insert({visualState, value});
            }
        }

        for (const auto& capture : resolvedRules.captures) {
            if (!capturesAdded.insert(capture.varName).second) continue;
            result.captures.push_back(capture);
        }
    }

    std::erase_if(result.overridesPerVSG, [](const auto& item) { return item.second.empty(); });
    return result;
}

void RestoreCapturesForElement(FrameworkElement element,
                               const ElementCustomizationState& elementState) {
    if (!element) return;
    for (const auto& [property, captureState] : elementState.captureCustomizationStates) {
        if (!captureState.propertyChangedToken) continue;
        try { element.UnregisterPropertyChangedCallback(property, captureState.propertyChangedToken); } catch (...) {}
    }
    if (elementState.captureSizeChangedToken) {
        try { element.SizeChanged(elementState.captureSizeChangedToken); } catch (...) {}
    }
}

void ApplyCustomizationsForVisualStateGroup(
    StyleVariableState* state,
    InstanceHandle handle,
    FrameworkElement element,
    VisualStateGroup visualStateGroup,
    PCWSTR fallbackClassName,
    PropertyOverrides propertyOverrides,
    ElementCustomizationStateForVisualStateGroup* elementCustomizationStateForVisualStateGroup) {
    auto elementDo = element.as<DependencyObject>();
    VisualState currentVisualState(visualStateGroup ? visualStateGroup.CurrentState() : nullptr);
    std::wstring currentVisualStateName(currentVisualState ? currentVisualState.Name() : L"");

    for (const auto& [property, valuesPerVisualState] : propertyOverrides) {
        const auto [propertyCustomizationStatesIt, inserted] =
            elementCustomizationStateForVisualStateGroup->propertyCustomizationStates.insert({property, {}});
        if (!inserted) continue;

        auto& propertyCustomizationState = propertyCustomizationStatesIt->second;
        auto it = valuesPerVisualState.find(currentVisualStateName);
        if (it == valuesPerVisualState.end() && !currentVisualStateName.empty()) {
            it = valuesPerVisualState.find(L"");
        }

        if (it != valuesPerVisualState.end()) {
            std::optional<PropertyOverrideValue> resolved = it->second;
            if (resolved) {
                propertyCustomizationState.originalValue = ReadLocalValueWithWorkaround(element, property);
                propertyCustomizationState.customValue = *resolved;
                SetOrClearValue(element, property, *resolved, true);
                propertyCustomizationState.lastAppliedValue = ReadLocalValueWithWorkaround(element, property);
            }
        }

        propertyCustomizationState.propertyChangedToken = elementDo.RegisterPropertyChangedCallback(
            property,
            [&propertyCustomizationState](DependencyObject sender, DependencyProperty property) {
                if (g_elementPropertyModifying) return;
                auto element = sender.try_as<FrameworkElement>();
                if (!element || !propertyCustomizationState.customValue) return;

                auto localValue = ReadLocalValueWithWorkaround(element, property);
                if (localValue != propertyCustomizationState.lastAppliedValue) {
                    propertyCustomizationState.originalValue = localValue;
                }

                g_elementPropertyModifying = true;
                SetOrClearValue(element, property, *propertyCustomizationState.customValue);
                propertyCustomizationState.lastAppliedValue = ReadLocalValueWithWorkaround(element, property);
                g_elementPropertyModifying = false;
            });
    }

    bool hasVisualStateDependentRules = false;
    for (const auto& [property, valuesPerVisualState] : propertyOverrides) {
        for (const auto& [visualStateName, val] : valuesPerVisualState) {
            if (!visualStateName.empty()) {
                hasVisualStateDependentRules = true;
                break;
            }
        }
        if (hasVisualStateDependentRules) break;
    }

    if (visualStateGroup && hasVisualStateDependentRules) {
        winrt::weak_ref<FrameworkElement> elementWeakRef = element;
        std::wstring fallbackClassNameStr = fallbackClassName ? fallbackClassName : L"";
        elementCustomizationStateForVisualStateGroup->visualStateGroupCurrentStateChangedToken =
            visualStateGroup.CurrentStateChanged(
                [elementWeakRef, propertyOverrides, fallbackClassNameStr,
                 elementCustomizationStateForVisualStateGroup](
                    winrt::Windows::Foundation::IInspectable const& sender,
                    VisualStateChangedEventArgs const& e) {
                    auto element = elementWeakRef.get();
                    if (!element) return;

                    Wh_Log(L"Re-applying all styles for %s on visual state change", winrt::get_class_name(element).c_str());
                    g_elementPropertyModifying = true;

                    auto& propertyCustomizationStates = elementCustomizationStateForVisualStateGroup->propertyCustomizationStates;
                    for (const auto& [property, valuesPerVisualState] : propertyOverrides) {
                        auto propertyIt = propertyCustomizationStates.find(property);
                        if (propertyIt == propertyCustomizationStates.end()) continue;
                        auto& propertyCustomizationState = propertyIt->second;

                        auto newState = e.NewState();
                        auto newStateName = std::wstring{newState ? newState.Name() : L""};
                        auto it = valuesPerVisualState.find(newStateName);
                        if (it == valuesPerVisualState.end()) {
                            it = valuesPerVisualState.find(L"");
                            if (it != valuesPerVisualState.end()) {
                                auto oldState = e.OldState();
                                auto oldStateName = std::wstring{oldState ? oldState.Name() : L""};
                                if (!valuesPerVisualState.contains(oldStateName)) continue;
                            }
                        }

                        if (it != valuesPerVisualState.end()) {
                            std::optional<PropertyOverrideValue> resolved = it->second;
                            if (resolved) {
                                propertyCustomizationState.customValue = *resolved;
                                SetOrClearValue(element, property, *resolved);
                                propertyCustomizationState.lastAppliedValue = ReadLocalValueWithWorkaround(element, property);
                            }
                        }
                    }

                    g_elementPropertyModifying = false;
                });
    }
}

void RestoreCustomizationsForVisualStateGroup(
    StyleVariableState* state,
    InstanceHandle handle,
    FrameworkElement element,
    std::optional<winrt::weak_ref<VisualStateGroup>> visualStateGroupOptionalWeakPtr,
    const ElementCustomizationStateForVisualStateGroup& elementCustomizationStateForVisualStateGroup) {
    if (elementCustomizationStateForVisualStateGroup.visualStateGroupCurrentStateChangedToken && visualStateGroupOptionalWeakPtr) {
        if (auto visualStateGroup = visualStateGroupOptionalWeakPtr->get()) {
            try { visualStateGroup.CurrentStateChanged(elementCustomizationStateForVisualStateGroup.visualStateGroupCurrentStateChangedToken); } catch (...) {}
        }
    }
    if (element) {
        for (const auto& [property, propState] : elementCustomizationStateForVisualStateGroup.propertyCustomizationStates) {
            try { element.UnregisterPropertyChangedCallback(property, propState.propertyChangedToken); } catch (...) {}
            if (propState.originalValue) {
                SetOrClearValue(element, property, *propState.originalValue);
            }
        }
    }
}

void ApplyCustomizations(InstanceHandle handle,
                         FrameworkElement element,
                         PCWSTR fallbackClassName) {
    if (!element) return;

    if (!g_resourceVariablesThemeDict) {
        MergeResourceVariables();
    }

    auto resolved = FindElementPropertyOverrides(element, fallbackClassName);
    if (resolved.overridesPerVSG.empty() && resolved.captures.empty()) {
        return;
    }

    Wh_Log(L"Applying styles to %s", winrt::get_class_name(element).c_str());

    auto& elementCustomizationState = g_elementsCustomizationState[handle];
    elementCustomizationState.element = element;
    elementCustomizationState.perVisualStateGroup.clear();

    for (auto& [visualStateGroup, overridesForVisualStateGroup] : resolved.overridesPerVSG) {
        std::optional<winrt::weak_ref<VisualStateGroup>> visualStateGroupOptionalWeakPtr;
        if (visualStateGroup) visualStateGroupOptionalWeakPtr = visualStateGroup;

        elementCustomizationState.perVisualStateGroup.push_back({visualStateGroupOptionalWeakPtr, {}});
        auto* elementCustomizationStateForVisualStateGroup = &elementCustomizationState.perVisualStateGroup.back().second;

        ApplyCustomizationsForVisualStateGroup(
            GetStyleVariableState(), handle, element, visualStateGroup, fallbackClassName,
            std::move(overridesForVisualStateGroup), elementCustomizationStateForVisualStateGroup);
    }
}

void CleanupCustomizations(InstanceHandle handle) {
    auto it = g_elementsCustomizationState.find(handle);
    if (it == g_elementsCustomizationState.end()) return;

    auto& elementCustomizationState = it->second;
    auto element = elementCustomizationState.element.get();
    auto* state = GetStyleVariableState();

    RestoreCapturesForElement(element, elementCustomizationState);

    for (const auto& [visualStateGroupOptionalWeakPtrIter, stateIter] : elementCustomizationState.perVisualStateGroup) {
        RestoreCustomizationsForVisualStateGroup(state, handle, element, visualStateGroupOptionalWeakPtrIter, stateIter);
    }

    g_elementsCustomizationState.erase(handle);
}

void UninitializeForCurrentThread() {
    if (auto& timer = g_trackedImageBrushesForThread.retryDebounceTimer) {
        try { timer.Stop(); } catch (...) {}
    }
    g_trackedImageBrushesForThread.retryDebounceTimerTickRevoker.revoke();
    g_trackedImageBrushesForThread.retryDebounceTimer = nullptr;
    g_trackedImageBrushesForThread.brushes.clear();
    g_trackedImageBrushesForThread.dispatcher = nullptr;

    g_elementsCustomizationState.clear();
    g_elementTreeNodes.clear();
    g_elementTreeNodesReapThreshold = 64;
    g_pendingStyleVariablePropagations.clear();
    g_styleVariableState = {};
    g_elementsCustomizationRules.clear();
    g_rulesByTypeMap.clear();
    g_genericRules.clear();
    g_candidateTypeSet.clear();
    g_typeMatchCache.clear();

    UninitializeResourceVariables();
    g_initializedForThread = false;
}

void UninitializeSettingsAndTap() {
    if (g_visualTreeWatcher) {
        g_visualTreeWatcher->UnadviseVisualTreeChange();
        g_visualTreeWatcher = nullptr;
    }
    g_initialized = false;
}

void InitializeForCurrentThread() {
    if (g_initializedForThread) return;
    ProcessAllStylesFromSettings();
    g_initializedForThread = true;
}

void InitializeSettingsAndTap() {
    if (g_initialized.exchange(true)) return;
    HRESULT hr = InjectWindhawkTAP();
    if (FAILED(hr)) Wh_Log(L"Error %08X", hr);
}

enum class TargetWindowType {
    None,
    FileExplorer,
    XamlExplorerHost,
};

TargetWindowType GetTargetWindowType(HWND hWnd) {
    WCHAR className[64];
    if (!GetClassName(hWnd, className, ARRAYSIZE(className))) return TargetWindowType::None;

    if (_wcsicmp(className, L"Microsoft.UI.Content.DesktopChildSiteBridge") == 0) {
        return TargetWindowType::FileExplorer;
    }
    if (_wcsicmp(className, L"XamlExplorerHostIslandWindow_WASDK") == 0) {
        return TargetWindowType::XamlExplorerHost;
    }
    return TargetWindowType::None;
}

void OnWindowCreated(HWND hWnd, PCSTR funcName) {
    TargetWindowType windowType = GetTargetWindowType(hWnd);
    if (windowType != TargetWindowType::None) {
        Wh_Log(L"Initializing - Created window %08X via %S", (DWORD)(ULONG_PTR)hWnd, funcName);
        InitializeForCurrentThread();
        InitializeSettingsAndTap();
    }
}

using CreateWindowExW_t = decltype(&CreateWindowExW);
CreateWindowExW_t CreateWindowExW_Original;
HWND WINAPI CreateWindowExW_Hook(DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName, DWORD dwStyle, int X, int Y, int nWidth, int nHeight, HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, PVOID lpParam) {
    HWND hWnd = CreateWindowExW_Original(dwExStyle, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);
    if (hWnd) OnWindowCreated(hWnd, __FUNCTION__);
    return hWnd;
}

using CreateWindowInBand_t = HWND(WINAPI*)(DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName, DWORD dwStyle, int X, int Y, int nWidth, int nHeight, HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, PVOID lpParam, DWORD dwBand);
CreateWindowInBand_t CreateWindowInBand_Original;
HWND WINAPI CreateWindowInBand_Hook(DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName, DWORD dwStyle, int X, int Y, int nWidth, int nHeight, HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, PVOID lpParam, DWORD dwBand) {
    HWND hWnd = CreateWindowInBand_Original(dwExStyle, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam, dwBand);
    if (hWnd) OnWindowCreated(hWnd, __FUNCTION__);
    return hWnd;
}

using CreateWindowInBandEx_t = HWND(WINAPI*)(DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName, DWORD dwStyle, int X, int Y, int nWidth, int nHeight, HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, PVOID lpParam, DWORD dwBand, DWORD dwTypeFlags);
CreateWindowInBandEx_t CreateWindowInBandEx_Original;
HWND WINAPI CreateWindowInBandEx_Hook(DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName, DWORD dwStyle, int X, int Y, int nWidth, int nHeight, HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, PVOID lpParam, DWORD dwBand, DWORD dwTypeFlags) {
    HWND hWnd = CreateWindowInBandEx_Original(dwExStyle, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam, dwBand, dwTypeFlags);
    if (hWnd) OnWindowCreated(hWnd, __FUNCTION__);
    return hWnd;
}

PFN_INITIALIZE_XAML_DIAGNOSTICS_EX InitializeXamlDiagnosticsEx_Original;
HRESULT WINAPI InitializeXamlDiagnosticsEx_Hook(_In_ PCWSTR endPointName, _In_ DWORD pid, _In_ PCWSTR wszDllXamlDiagnostics, _In_ PCWSTR wszTAPDllName, _In_ CLSID tapClsid, _In_opt_ PCWSTR wszInitializationData) {
    if (g_inInjectWindhawkTAP) {
        return InitializeXamlDiagnosticsEx_Original(endPointName, pid, wszDllXamlDiagnostics, wszTAPDllName, tapClsid, wszInitializationData);
    }
    bool blockCall = false;
    switch (g_settings.xamlDiagnosticsHandling) {
        case XamlDiagnosticsHandling::kAlert: {
            void* retAddress = __builtin_return_address(0);
            WCHAR modulePath[MAX_PATH];
            PCWSTR modulePathStr = L"<unknown>";
            HMODULE module;
            if (GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, reinterpret_cast<LPCWSTR>(retAddress), &module)) {
                if (GetModuleFileName(module, modulePath, ARRAYSIZE(modulePath))) modulePathStr = modulePath;
            }
            WCHAR message[1024];
            _snwprintf_s(message, _TRUNCATE,
                L"The following module is trying to use XAML diagnostics:\n\n%s\n\nThere can only be one consumer at a time. Blocking it might break that module, but allowing it might break this mod.\n\nDo you want to block it?\n\nNote: You can change this behavior in the mod settings.",
                modulePathStr);
            int result = MessageBox(nullptr, message, L"Files Styler - Windhawk", MB_YESNO | MB_ICONQUESTION | MB_TOPMOST);
            blockCall = (result == IDYES);
            break;
        }
        case XamlDiagnosticsHandling::kBlock: blockCall = true; break;
        case XamlDiagnosticsHandling::kAllow: blockCall = false; break;
    }
    if (blockCall) { Wh_Log(L"Blocking InitializeXamlDiagnosticsEx call"); return S_OK; }
    return InitializeXamlDiagnosticsEx_Original(endPointName, pid, wszDllXamlDiagnostics, wszTAPDllName, tapClsid, wszInitializationData);
}

bool HookInitializeXamlDiagnosticsExIfNeeded() {
    if (InitializeXamlDiagnosticsEx_Original) return false;
    const HMODULE wux = GetModuleHandle(L"Microsoft.Internal.FrameworkUdk.dll");
    if (!wux) return false;
    const auto ixde = reinterpret_cast<PFN_INITIALIZE_XAML_DIAGNOSTICS_EX>(GetProcAddress(wux, "InitializeXamlDiagnosticsEx"));
    if (!ixde) return false;
    return WindhawkUtils::SetFunctionHook(ixde, InitializeXamlDiagnosticsEx_Hook, &InitializeXamlDiagnosticsEx_Original);
}

using LoadLibraryExW_t = decltype(&LoadLibraryExW);
LoadLibraryExW_t LoadLibraryExW_Original;
HMODULE WINAPI LoadLibraryExW_Hook(LPCWSTR lpLibFileName, HANDLE hFile, DWORD dwFlags) {
    HMODULE module = LoadLibraryExW_Original(lpLibFileName, hFile, dwFlags);
    if (module && !InitializeXamlDiagnosticsEx_Original && lpLibFileName) {
        PCWSTR fileName = wcsrchr(lpLibFileName, L'\\');
        fileName = fileName ? fileName + 1 : lpLibFileName;
        if (_wcsicmp(fileName, L"CoreMessagingXP.dll") == 0 && HookInitializeXamlDiagnosticsExIfNeeded()) {
            Wh_ApplyHookOperations();
        }
    }
    return module;
}

using RunFromWindowThreadProc_t = void(WINAPI*)(PVOID parameter);

bool RunFromWindowThread(HWND hWnd, RunFromWindowThreadProc_t proc, PVOID procParam) {
    static const UINT runFromWindowThreadRegisteredMsg = RegisterWindowMessage(L"Windhawk_RunFromWindowThread_" WH_MOD_ID);
    struct RUN_FROM_WINDOW_THREAD_PARAM { RunFromWindowThreadProc_t proc; PVOID procParam; };
    DWORD dwThreadId = GetWindowThreadProcessId(hWnd, nullptr);
    if (dwThreadId == 0) return false;
    if (dwThreadId == GetCurrentThreadId()) { proc(procParam); return true; }

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
    RUN_FROM_WINDOW_THREAD_PARAM param{proc, procParam};
    SendMessage(hWnd, runFromWindowThreadRegisteredMsg, 0, (LPARAM)&param);
    UnhookWindowsHookEx(hook);
    return true;
}

std::vector<HWND> GetTargetWnds() {
    struct ENUM_WINDOWS_PARAM { std::vector<HWND>* hWnds; };
    std::vector<HWND> hWnds;
    ENUM_WINDOWS_PARAM param = {&hWnds};
    EnumWindows([](HWND hWnd, LPARAM lParam) -> BOOL {
        ENUM_WINDOWS_PARAM& param = *(ENUM_WINDOWS_PARAM*)lParam;
        DWORD dwProcessId = 0;
        if (!GetWindowThreadProcessId(hWnd, &dwProcessId) || dwProcessId != GetCurrentProcessId()) return TRUE;
        if (GetTargetWindowType(hWnd) != TargetWindowType::None) param.hWnds->push_back(hWnd);
        return TRUE;
    }, (LPARAM)&param);
    return hWnds;
}

void LoadSettings() {
    PCWSTR xamlDiagnosticsHandling = Wh_GetStringSetting(L"xamlDiagnosticsHandling");
    g_settings.xamlDiagnosticsHandling = XamlDiagnosticsHandling::kAlert;
    if (wcscmp(xamlDiagnosticsHandling, L"block") == 0) g_settings.xamlDiagnosticsHandling = XamlDiagnosticsHandling::kBlock;
    else if (wcscmp(xamlDiagnosticsHandling, L"allow") == 0) g_settings.xamlDiagnosticsHandling = XamlDiagnosticsHandling::kAllow;
    Wh_FreeStringSetting(xamlDiagnosticsHandling);
}

BOOL Wh_ModInit() {
    Wh_Log(L">");
    LoadSettings();

    WindhawkUtils::SetFunctionHook(CreateWindowExW, CreateWindowExW_Hook, &CreateWindowExW_Original);

    HMODULE user32Module = LoadLibraryEx(L"user32.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (user32Module) {
        auto pCreateWindowInBand = (CreateWindowInBand_t)GetProcAddress(user32Module, "CreateWindowInBand");
        if (pCreateWindowInBand) WindhawkUtils::SetFunctionHook(pCreateWindowInBand, CreateWindowInBand_Hook, &CreateWindowInBand_Original);
        auto pCreateWindowInBandEx = (CreateWindowInBandEx_t)GetProcAddress(user32Module, "CreateWindowInBandEx");
        if (pCreateWindowInBandEx) WindhawkUtils::SetFunctionHook(pCreateWindowInBandEx, CreateWindowInBandEx_Hook, &CreateWindowInBandEx_Original);
    }

    HMODULE kernelBaseModule = GetModuleHandle(L"kernelbase.dll");
    auto pKernelBaseLoadLibraryExW = (decltype(&LoadLibraryExW))GetProcAddress(kernelBaseModule, "LoadLibraryExW");
    WindhawkUtils::SetFunctionHook(pKernelBaseLoadLibraryExW, LoadLibraryExW_Hook, &LoadLibraryExW_Original);

    HookInitializeXamlDiagnosticsExIfNeeded();
    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");
    auto hTargetWnds = GetTargetWnds();
    for (auto hTargetWnd : hTargetWnds) {
        RunFromWindowThread(hTargetWnd, [](PVOID) { InitializeForCurrentThread(); }, (PVOID)hTargetWnd);
    }
    if (hTargetWnds.size() > 0) {
        InitializeSettingsAndTap();
    }
}

void Wh_ModUninit() {
    Wh_Log(L">");
    StopImageLoadRetries();
    UninitializeSettingsAndTap();

    auto hTargetWnds = GetTargetWnds();
    for (auto hTargetWnd : hTargetWnds) {
        RunFromWindowThread(hTargetWnd, [](PVOID) { UninitializeForCurrentThread(); }, (PVOID)hTargetWnd);
    }
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");
    UninitializeSettingsAndTap();
    LoadSettings();

    auto hTargetWnds = GetTargetWnds();
    for (auto hTargetWnd : hTargetWnds) {
        RunFromWindowThread(hTargetWnd, [](PVOID) {
            UninitializeForCurrentThread();
            InitializeForCurrentThread();
        }, (PVOID)hTargetWnd);
    }
    if (hTargetWnds.size() > 0) {
        InitializeSettingsAndTap();
    }
}
