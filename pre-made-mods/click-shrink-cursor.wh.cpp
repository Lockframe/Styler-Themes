// ==WindhawkMod==
// @id              cursor-shrink-click-stable
// @name            Cursor Shrink (Stable)
// @description     Optimized click animation with safe unloading
// @version         1.4
// @author          Gemini
// @include         explorer.exe
// ==/WindhawkMod==

#include <windows.h>
#include <atomic>

// --- CONSTANTS ---
#ifndef OCR_NORMAL
#define OCR_NORMAL 32512
#endif

#ifndef LR_COPYRETURNORG
#define LR_COPYRETURNORG 0x0004
#endif

// --- STATE ---
HHOOK g_mouseHook = NULL;
std::atomic<bool> g_isAnimating(false);
HANDLE g_hThread = NULL;
DWORD g_dwThreadId = 0; // NEW: Store ID to signal thread later

void restore_default_cursor() {
    SystemParametersInfo(SPI_SETCURSORS, 0, NULL, 0);
}

DWORD WINAPI AnimationThread(LPVOID lpParam) {
    g_isAnimating = true;

    int baseWidth = GetSystemMetrics(SM_CXCURSOR);
    int baseHeight = GetSystemMetrics(SM_CYCURSOR);

    HCURSOR hBase = (HCURSOR)LoadImage(NULL, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_SHARED | LR_DEFAULTSIZE);

    if (!hBase) {
        g_isAnimating = false;
        return 0;
    }

    // Shrink Phase
    for (float scale = 0.9f; scale >= 0.7f; scale -= 0.1f) {
        int w = (int)(baseWidth * scale);
        int h = (int)(baseHeight * scale);

        HCURSOR hNew = (HCURSOR)CopyImage(hBase, IMAGE_CURSOR, w, h, LR_COPYRETURNORG);
        
        if (hNew) SetSystemCursor(hNew, OCR_NORMAL); 
        Sleep(20); 
    }

    // Snap back Phase
    Sleep(30);
    restore_default_cursor();

    g_isAnimating = false;
    return 0;
}

LRESULT CALLBACK MouseHookProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0) {
        if (wParam == WM_LBUTTONDOWN) {
            if (!g_isAnimating) {
                HANDLE hAnim = CreateThread(NULL, 0, AnimationThread, NULL, 0, NULL);
                if (hAnim) {
                    SetThreadPriority(hAnim, THREAD_PRIORITY_ABOVE_NORMAL);
                    CloseHandle(hAnim);
                }
            }
        }
    }
    return CallNextHookEx(g_mouseHook, nCode, wParam, lParam);
}

DWORD WINAPI HookThread(LPVOID lpParam) {
    // Force the message queue to be created
    PeekMessage(NULL, NULL, WM_USER, WM_USER, PM_NOREMOVE);
    
    g_mouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseHookProc, NULL, 0);
    
    MSG msg;
    // GetMessage blocks here until a message (like WM_QUIT) arrives
    while (GetMessage(&msg, NULL, 0, 0)) {
        if (msg.message == WM_QUIT) break; // Explicitly break on quit
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    
    if (g_mouseHook) {
        UnhookWindowsHookEx(g_mouseHook);
        g_mouseHook = NULL;
    }
    return 0;
}

BOOL Wh_ModInit() {
    // Create thread and store the ID so we can message it later
    g_hThread = CreateThread(NULL, 0, HookThread, NULL, 0, &g_dwThreadId);
    return TRUE;
}

void Wh_ModUninit() {
    // 1. Tell the thread to stop waiting and exit
    if (g_dwThreadId != 0) {
        PostThreadMessage(g_dwThreadId, WM_QUIT, 0, 0);
    }

    // 2. Wait for it to actually finish (prevent unloading code while thread is running)
    if (g_hThread) {
        WaitForSingleObject(g_hThread, 3000); // Wait up to 3 seconds
        CloseHandle(g_hThread);
        g_hThread = NULL;
    }

    // 3. Final cleanup
    if (g_mouseHook) {
        UnhookWindowsHookEx(g_mouseHook);
        g_mouseHook = NULL;
    }
    
    restore_default_cursor();
}