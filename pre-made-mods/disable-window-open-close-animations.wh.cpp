// ==WindhawkMod==
// @id              disable-window-anim-v6
// @name            Disable Window Open/Close Animations (UWP Close Fix)
// @description     Disables Open/Close animations. Adds message hooks to catch UWP close commands (SC_CLOSE) before animation starts.
// @version         6.0
// @author          Gemini
// @include         *
// @compilerOptions -ldwmapi -luser32
// ==/WindhawkMod==

#include <windhawk_utils.h>
#include <dwmapi.h>

// ---------------------------------------------------------------------------
// CONFIGURATION
// ---------------------------------------------------------------------------
#define RESTORE_ANIM_DELAY_MS 500 
#define TIMER_ID_RESTORE 2999

#ifndef DWMWA_TRANSITIONS_FORCEDISABLED
#define DWMWA_TRANSITIONS_FORCEDISABLED 3
#endif

#ifndef DWMWA_CLOAK
#define DWMWA_CLOAK 13
#endif

// ---------------------------------------------------------------------------
// HELPERS
// ---------------------------------------------------------------------------

void SetTransitions(HWND hWnd, BOOL enable) {
    BOOL value = !enable; // TRUE = Disable, FALSE = Enable
    DwmSetWindowAttribute(hWnd, DWMWA_TRANSITIONS_FORCEDISABLED, &value, sizeof(value));
}

bool IsTargetWindow(HWND hWnd) {
    if (GetAncestor(hWnd, GA_ROOT) != hWnd) return false;

    char className[256];
    if (GetClassNameA(hWnd, className, sizeof(className))) {
        if (strcmp(className, "ApplicationFrameWindow") == 0) return true;
        if (strcmp(className, "Windows.UI.Core.CoreWindow") == 0) return true;
    }

    LONG_PTR style = GetWindowLongPtr(hWnd, GWL_STYLE);
    if (style & WS_CHILD) return false;
    return ((style & WS_CAPTION) || (style & WS_THICKFRAME));
}

// ---------------------------------------------------------------------------
// TIMER CALLBACK
// ---------------------------------------------------------------------------
VOID CALLBACK RestoreAnimTimerProc(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime) {
    KillTimer(hWnd, idEvent);
    if (IsWindow(hWnd)) {
        SetTransitions(hWnd, TRUE); 
    }
}

// ---------------------------------------------------------------------------
// HOOKS
// ---------------------------------------------------------------------------

// 1. Hook DefWindowProc
// This catches the system processing the "Close" command (SC_CLOSE).
// UWP apps often rely on the default handler for the title bar 'X'.
typedef LRESULT (WINAPI *DefWindowProcW_t)(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);
DefWindowProcW_t DefWindowProcW_Original;

LRESULT WINAPI DefWindowProcW_Hook(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam) {
    // Check for System Command -> Close
    if (Msg == WM_SYSCOMMAND && (wParam & 0xFFF0) == SC_CLOSE) {
        if (IsTargetWindow(hWnd)) {
            // User clicked X or pressed Alt+F4.
            // Disable animations IMMEDIATELY before the system processes the close.
            SetTransitions(hWnd, FALSE);
        }
    }
    return DefWindowProcW_Original(hWnd, Msg, wParam, lParam);
}

// 2. Hook DwmSetWindowAttribute (UWP Cloaking Fix)
typedef HRESULT (WINAPI *DwmSetWindowAttribute_t)(HWND hwnd, DWORD dwAttribute, LPCVOID pvAttribute, DWORD cbAttribute);
DwmSetWindowAttribute_t DwmSetWindowAttribute_Original;

HRESULT WINAPI DwmSetWindowAttribute_Hook(HWND hwnd, DWORD dwAttribute, LPCVOID pvAttribute, DWORD cbAttribute) {
    if (dwAttribute == DWMWA_CLOAK && IsTargetWindow(hwnd)) {
        int cloakVal = *(int*)pvAttribute;
        if (cloakVal == 0) { // UNCLOAK (Show)
            SetTransitions(hwnd, FALSE);
            SetTimer(hwnd, TIMER_ID_RESTORE, RESTORE_ANIM_DELAY_MS, RestoreAnimTimerProc);
        } else { // CLOAK (Hide)
            SetTransitions(hwnd, FALSE);
        }
    }
    return DwmSetWindowAttribute_Original(hwnd, dwAttribute, pvAttribute, cbAttribute);
}

// 3. Hook SetWindowPos
typedef BOOL (WINAPI *SetWindowPos_t)(HWND hWnd, HWND hWndInsertAfter, int X, int Y, int cx, int cy, UINT uFlags);
SetWindowPos_t SetWindowPos_Original;

BOOL WINAPI SetWindowPos_Hook(HWND hWnd, HWND hWndInsertAfter, int X, int Y, int cx, int cy, UINT uFlags) {
    bool showing = (uFlags & SWP_SHOWWINDOW);
    bool hiding  = (uFlags & SWP_HIDEWINDOW);

    if ((!showing && !hiding) || !IsTargetWindow(hWnd)) {
        return SetWindowPos_Original(hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);
    }

    if (hiding) {
        SetTransitions(hWnd, FALSE);
        return SetWindowPos_Original(hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);
    }

    if (showing) {
        SetTransitions(hWnd, FALSE);
        BOOL result = SetWindowPos_Original(hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);
        SetTimer(hWnd, TIMER_ID_RESTORE, RESTORE_ANIM_DELAY_MS, RestoreAnimTimerProc);
        return result;
    }
    return SetWindowPos_Original(hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);
}

// 4. Hook ShowWindow
typedef BOOL (WINAPI *ShowWindow_t)(HWND hWnd, int nCmdShow);
ShowWindow_t ShowWindow_Original;

BOOL WINAPI ShowWindow_Hook(HWND hWnd, int nCmdShow) {
    bool isOpenCmd  = (nCmdShow == SW_SHOWNORMAL || nCmdShow == SW_SHOW || nCmdShow == SW_RESTORE || nCmdShow == SW_SHOWDEFAULT);
    bool isCloseCmd = (nCmdShow == SW_HIDE);

    if (!IsTargetWindow(hWnd)) {
        return ShowWindow_Original(hWnd, nCmdShow);
    }

    if (IsWindowVisible(hWnd) == FALSE && isOpenCmd) {
        SetTransitions(hWnd, FALSE);
        BOOL res = ShowWindow_Original(hWnd, nCmdShow);
        SetTimer(hWnd, TIMER_ID_RESTORE, RESTORE_ANIM_DELAY_MS, RestoreAnimTimerProc);
        return res;
    }
    if (IsWindowVisible(hWnd) && isCloseCmd) {
        SetTransitions(hWnd, FALSE);
    }
    return ShowWindow_Original(hWnd, nCmdShow);
}

// 5. Hook DestroyWindow
typedef BOOL (WINAPI *DestroyWindow_t)(HWND hWnd);
DestroyWindow_t DestroyWindow_Original;

BOOL WINAPI DestroyWindow_Hook(HWND hWnd) {
    if (IsWindowVisible(hWnd) && IsTargetWindow(hWnd)) {
        SetTransitions(hWnd, FALSE);
    }
    return DestroyWindow_Original(hWnd);
}

BOOL Wh_ModInit() {
    Wh_SetFunctionHook((void*)DefWindowProcW, (void*)DefWindowProcW_Hook, (void**)&DefWindowProcW_Original);
    Wh_SetFunctionHook((void*)SetWindowPos, (void*)SetWindowPos_Hook, (void**)&SetWindowPos_Original);
    Wh_SetFunctionHook((void*)ShowWindow, (void*)ShowWindow_Hook, (void**)&ShowWindow_Original);
    Wh_SetFunctionHook((void*)DestroyWindow, (void*)DestroyWindow_Hook, (void**)&DestroyWindow_Original);
    Wh_SetFunctionHook((void*)DwmSetWindowAttribute, (void*)DwmSetWindowAttribute_Hook, (void**)&DwmSetWindowAttribute_Original);
    return TRUE;
}