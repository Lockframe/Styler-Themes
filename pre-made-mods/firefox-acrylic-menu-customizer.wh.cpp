// ==WindhawkMod==
// @id              firefox-acrylic-menu-customizer
// @name            Firefox Acrylic Menu Customizer
// @description     Customize Firefox's implementation of acrylic context menus.
// @version         1.1
// @author          Lockframe
// @github          https://www.github.com/Lockframe
// @include         firefox.exe
// @compilerOptions -ldwmapi
// ==/WindhawkMod==

// ==WindhawkModSettings==
/*
- hideBorderLight: false
  $name: Hide Border (Light Mode)
  $description: Check to remove the DWM border when in Light Mode.

- borderColorLight: "#FFDDDDDD"
  $name: Border Color (Light Mode)
  $description: Format RRGGBB. Ignored if "Hide Border" is checked.

- hideBorderDark: false
  $name: Hide Border (Dark Mode)
  $description: Check to remove the DWM border when in Dark Mode.

- borderColorDark: "#FF1D1D1D"
  $name: Border Color (Dark Mode)
  $description: Format RRGGBB. Ignored if "Hide Border" is checked.

- backdropOverride: "mica_alt"
  $name: Backdrop Material
  $description: Force a specific material type.
  $options:
    - "default": "Default (Firefox Choice)"
    - "none": "None (Solid)"
    - "mica": "Mica (Opaque Theme Aware)"
    - "acrylic": "Acrylic (Transient / Washed Out)"
    - "mica_alt": "Mica Alt (Deep Theme Aware)"
    - "glass": "Legacy Glass / Accent Blur (Undocumented)"

- cornerPreference: "default"
  $name: Corner Rounding
  $options:
    - "default": "System Default (Rounded)"
    - "square": "Square (No Rounding)"
    - "round": "Round (Standard)"
    - "small": "Small Round"

- forceDarkMode: false
  $name: Force Dark Mode Context
  $description: Forces the OS to treat the popup as a Dark Mode window.

- glassState: 4
  $name: (Glass Only) Blur Method
  $description: 3 = Standard Blur (Clear). 4 = Acrylic Blur (Supports Tint).
  $options:
    - 3: "Standard Blur (Clear)"
    - 4: "Acrylic Blur (Supports Tint)"

- glassColorLight: "#90FFFFFF"
  $name: (Glass Only) Light Mode Tint
  $description: Format AARRGGBB.

- glassColorDark: "#90444444"
  $name: (Glass Only) Dark Mode Tint
  $description: Format AARRGGBB.
*/
// ==/WindhawkModSettings==

#include <windows.h>
#include <dwmapi.h>
#include <cwchar>

// --- Constants ---
#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif
#ifndef DWMWA_WINDOW_CORNER_PREFERENCE
#define DWMWA_WINDOW_CORNER_PREFERENCE 33
#endif
#ifndef DWMWA_BORDER_COLOR
#define DWMWA_BORDER_COLOR 34
#endif
#ifndef DWMWA_SYSTEMBACKDROP_TYPE
#define DWMWA_SYSTEMBACKDROP_TYPE 38
#endif

// --- Undocumented API ---
typedef enum _WINDOWCOMPOSITIONATTRIB { WCA_ACCENT_POLICY = 19 } WINDOWCOMPOSITIONATTRIB;
typedef enum _ACCENT_STATE {
    ACCENT_DISABLED = 0, ACCENT_ENABLE_GRADIENT = 1, ACCENT_ENABLE_TRANSPARENTGRADIENT = 2,
    ACCENT_ENABLE_BLURBEHIND = 3, ACCENT_ENABLE_ACRYLICBLURBEHIND = 4, ACCENT_INVALID_STATE = 5
} ACCENT_STATE;
typedef struct _ACCENT_POLICY {
    ACCENT_STATE AccentState; DWORD AccentFlags; DWORD GradientColor; DWORD AnimationId;
} ACCENT_POLICY;
typedef struct _WINDOWCOMPOSITIONATTRIBDATA {
    WINDOWCOMPOSITIONATTRIB Attribute; PVOID Data; SIZE_T SizeOfData;
} WINDOWCOMPOSITIONATTRIBDATA;
typedef BOOL (WINAPI *pSetWindowCompositionAttribute)(HWND, WINDOWCOMPOSITIONATTRIBDATA*);
pSetWindowCompositionAttribute SetWindowCompositionAttribute = NULL;

// --- Helper Functions ---
BOOL IsSystemDarkMode() {
    DWORD value = 0; DWORD size = sizeof(value);
    LSTATUS status = RegGetValueW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", L"AppsUseLightTheme", RRF_RT_REG_DWORD, NULL, &value, &size);
    return (status == ERROR_SUCCESS && value == 0);
}

COLORREF ParseColorHexForDwm(PCWSTR hexStr) {
    if (!hexStr || !*hexStr) return 0;
    if (hexStr[0] == L'#') hexStr++;
    unsigned long val = wcstoul(hexStr, NULL, 16);
    return RGB((val >> 16) & 0xFF, (val >> 8) & 0xFF, val & 0xFF);
}

DWORD ParseColorHexForGlass(PCWSTR hexStr) {
    if (!hexStr || !*hexStr) return 0;
    if (hexStr[0] == L'#') hexStr++;
    unsigned long val = wcstoul(hexStr, NULL, 16);
    DWORD a = (val >> 24) & 0xFF; DWORD r = (val >> 16) & 0xFF; DWORD g = (val >> 8) & 0xFF; DWORD b = val & 0xFF;
    if (a == 0) a = 1;
    return (a << 24) | (b << 16) | (g << 8) | r;
}

int GetBackdropType() {
    PCWSTR typeStr = Wh_GetStringSetting(L"backdropOverride");
    int type = 0; 
    if (wcscmp(typeStr, L"none") == 0) type = 1;
    else if (wcscmp(typeStr, L"mica") == 0) type = 2;
    else if (wcscmp(typeStr, L"acrylic") == 0) type = 3;
    else if (wcscmp(typeStr, L"mica_alt") == 0) type = 4;
    else if (wcscmp(typeStr, L"glass") == 0) type = 99; 
    Wh_FreeStringSetting(typeStr);
    return type;
}

int GetCornerPreference() {
    PCWSTR typeStr = Wh_GetStringSetting(L"cornerPreference");
    int type = 0;
    if (wcscmp(typeStr, L"square") == 0) type = 1;
    else if (wcscmp(typeStr, L"round") == 0) type = 2;
    else if (wcscmp(typeStr, L"small") == 0) type = 3;
    Wh_FreeStringSetting(typeStr);
    return type;
}

BOOL IsTargetPopup(HWND hWnd) {
    if (!IsWindow(hWnd)) return FALSE;
    LONG_PTR style = GetWindowLongPtrW(hWnd, GWL_STYLE);
    if (!(style & WS_POPUP)) return FALSE;
    wchar_t className[256];
    if (GetClassNameW(hWnd, className, 256)) {
        if (wcscmp(className, L"MozillaDropShadowWindowClass") == 0) return TRUE;
    }
    return FALSE;
}

void EnableBlurBehind(HWND hWnd) {
    if (!SetWindowCompositionAttribute) {
        HMODULE hUser32 = GetModuleHandle(L"user32.dll");
        if (hUser32) SetWindowCompositionAttribute = (pSetWindowCompositionAttribute)GetProcAddress(hUser32, "SetWindowCompositionAttribute");
    }

    if (SetWindowCompositionAttribute) {
        PCWSTR colorStr = IsSystemDarkMode() ? Wh_GetStringSetting(L"glassColorDark") : Wh_GetStringSetting(L"glassColorLight");
        DWORD glassColor = ParseColorHexForGlass(colorStr); 
        Wh_FreeStringSetting(colorStr);
        
        int glassState = Wh_GetIntSetting(L"glassState");
        if (glassState == 0) glassState = 3; 

        ACCENT_POLICY policy;
        policy.AccentState = (ACCENT_STATE)glassState;
        policy.AccentFlags = 0; 
        policy.GradientColor = glassColor; 
        policy.AnimationId = 0;

        WINDOWCOMPOSITIONATTRIBDATA data;
        data.Attribute = WCA_ACCENT_POLICY;
        data.Data = &policy;
        data.SizeOfData = sizeof(policy);

        SetWindowCompositionAttribute(hWnd, &data);
    }
}

void ApplyDwmSettings(HWND hWnd) {
    if (!IsTargetPopup(hWnd)) return;

    // 1. Theme-Aware Border Logic
    BOOL isDark = IsSystemDarkMode();
    BOOL shouldHide = isDark ? Wh_GetIntSetting(L"hideBorderDark") : Wh_GetIntSetting(L"hideBorderLight");
    
    COLORREF color;
    if (shouldHide) {
        color = 0xFFFFFFFE; 
    } else {
        PCWSTR colorStr = isDark ? Wh_GetStringSetting(L"borderColorDark") : Wh_GetStringSetting(L"borderColorLight");
        color = ParseColorHexForDwm(colorStr); 
        Wh_FreeStringSetting(colorStr);
    }
    DwmSetWindowAttribute(hWnd, DWMWA_BORDER_COLOR, &color, sizeof(color));

    // 2. Backdrop
    int backdropType = GetBackdropType();
    if (backdropType == 99) { 
        int none = 1; 
        DwmSetWindowAttribute(hWnd, DWMWA_SYSTEMBACKDROP_TYPE, &none, sizeof(none));
        EnableBlurBehind(hWnd);
    } else if (backdropType > 0) { 
        DwmSetWindowAttribute(hWnd, DWMWA_SYSTEMBACKDROP_TYPE, &backdropType, sizeof(backdropType));
    }

    // 3. Corners
    int cornerPref = GetCornerPreference();
    if (cornerPref > 0) DwmSetWindowAttribute(hWnd, DWMWA_WINDOW_CORNER_PREFERENCE, &cornerPref, sizeof(cornerPref));

    // 4. Force Dark Mode
    BOOL forceDark = Wh_GetIntSetting(L"forceDarkMode");
    if (forceDark) {
        BOOL useDark = TRUE;
        DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDark, sizeof(useDark));
    }

    SetWindowPos(hWnd, NULL, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
}

// --- Hooks ---
typedef HRESULT (WINAPI *DwmSetWindowAttribute_t)(HWND, DWORD, LPCVOID, DWORD);
DwmSetWindowAttribute_t DwmSetWindowAttribute_Original;
HRESULT WINAPI DwmSetWindowAttribute_Hook(HWND hwnd, DWORD dwAttribute, LPCVOID pvAttribute, DWORD cbAttribute) {
    if (dwAttribute == DWMWA_SYSTEMBACKDROP_TYPE) {
        if (IsTargetPopup(hwnd)) {
            int overrideType = GetBackdropType();
            if (overrideType > 0 && overrideType != 99) {
                return DwmSetWindowAttribute_Original(hwnd, dwAttribute, &overrideType, sizeof(overrideType));
            }
        }
    }
    return DwmSetWindowAttribute_Original(hwnd, dwAttribute, pvAttribute, cbAttribute);
}

typedef BOOL (WINAPI *ShowWindow_t)(HWND, int);
ShowWindow_t ShowWindow_Original;
BOOL WINAPI ShowWindow_Hook(HWND hWnd, int nCmdShow) {
    if (nCmdShow != SW_HIDE && nCmdShow != SW_MINIMIZE) ApplyDwmSettings(hWnd);
    return ShowWindow_Original(hWnd, nCmdShow);
}

typedef BOOL (WINAPI *SetWindowPos_t)(HWND, HWND, int, int, int, int, UINT);
SetWindowPos_t SetWindowPos_Original;
BOOL WINAPI SetWindowPos_Hook(HWND hWnd, HWND hWndInsertAfter, int X, int Y, int cx, int cy, UINT uFlags) {
    if (uFlags & SWP_SHOWWINDOW) ApplyDwmSettings(hWnd);
    return SetWindowPos_Original(hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);
}

typedef HWND (WINAPI *CreateWindowExW_t)(DWORD, LPCWSTR, LPCWSTR, DWORD, int, int, int, int, HWND, HMENU, HINSTANCE, LPVOID);
CreateWindowExW_t CreateWindowExW_Original;
HWND WINAPI CreateWindowExW_Hook(DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName, DWORD dwStyle, int X, int Y, int nWidth, int nHeight, HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam) {
    HWND hWnd = CreateWindowExW_Original(dwExStyle, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);
    if (hWnd) ApplyDwmSettings(hWnd);
    return hWnd;
}

// --- Init/Uninit ---
BOOL Wh_ModInit() {
    HMODULE hUser32 = GetModuleHandle(L"user32.dll");
    HMODULE hDwmapi = GetModuleHandle(L"dwmapi.dll");
    Wh_SetFunctionHook((void*)GetProcAddress(hUser32, "CreateWindowExW"), (void*)CreateWindowExW_Hook, (void**)&CreateWindowExW_Original);
    Wh_SetFunctionHook((void*)GetProcAddress(hUser32, "ShowWindow"), (void*)ShowWindow_Hook, (void**)&ShowWindow_Original);
    Wh_SetFunctionHook((void*)GetProcAddress(hUser32, "SetWindowPos"), (void*)SetWindowPos_Hook, (void**)&SetWindowPos_Original);
    if (hDwmapi) Wh_SetFunctionHook((void*)GetProcAddress(hDwmapi, "DwmSetWindowAttribute"), (void*)DwmSetWindowAttribute_Hook, (void**)&DwmSetWindowAttribute_Original);
    return TRUE;
}
void Wh_ModUninit() {
    HMODULE hUser32 = GetModuleHandle(L"user32.dll");
    HMODULE hDwmapi = GetModuleHandle(L"dwmapi.dll");
    Wh_RemoveFunctionHook((void*)GetProcAddress(hUser32, "CreateWindowExW"));
    Wh_RemoveFunctionHook((void*)GetProcAddress(hUser32, "ShowWindow"));
    Wh_RemoveFunctionHook((void*)GetProcAddress(hUser32, "SetWindowPos"));
    if (hDwmapi) Wh_RemoveFunctionHook((void*)GetProcAddress(hDwmapi, "DwmSetWindowAttribute"));
}