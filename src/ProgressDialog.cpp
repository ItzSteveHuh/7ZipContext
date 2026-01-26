// ProgressDialog.cpp - Progress dialog implementation
#include "ProgressDialog.h"
#include <CommCtrl.h>

#pragma comment(lib, "comctl32.lib")

// Window class name
static const wchar_t* PROGRESS_WND_CLASS = L"7ZipContextProgress";
static bool g_classRegistered = false;

CProgressDialog::CProgressDialog() {
}

CProgressDialog::~CProgressDialog() {
    Close();
}

void CProgressDialog::Show(const std::wstring& title) {
    m_title = title;
    m_cancelled.store(false);
    m_closed.store(false);

    // Create dialog in separate thread to avoid blocking
    m_hThread = CreateThread(NULL, 0, DialogThread, this, 0, &m_threadId);
}

DWORD WINAPI CProgressDialog::DialogThread(LPVOID lpParam) {
    CProgressDialog* pThis = (CProgressDialog*)lpParam;
    pThis->CreateDialogWindow();

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0) && !pThis->m_closed.load()) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}

void CProgressDialog::CreateDialogWindow() {
    // Initialize common controls
    INITCOMMONCONTROLSEX icc;
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_PROGRESS_CLASS;
    InitCommonControlsEx(&icc);

    // Register window class
    if (!g_classRegistered) {
        WNDCLASSEXW wc = {};
        wc.cbSize = sizeof(wc);
        wc.style = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc = WndProc;
        wc.hInstance = GetModuleHandle(NULL);
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = PROGRESS_WND_CLASS;
        RegisterClassExW(&wc);
        g_classRegistered = true;
    }

    // Create main window
    int width = 400;
    int height = 120;
    int x = (GetSystemMetrics(SM_CXSCREEN) - width) / 2;
    int y = (GetSystemMetrics(SM_CYSCREEN) - height) / 2;

    m_hWnd = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
        PROGRESS_WND_CLASS,
        m_title.c_str(),
        WS_POPUP | WS_CAPTION | WS_SYSMENU,
        x, y, width, height,
        NULL, NULL,
        GetModuleHandle(NULL),
        this
    );

    if (!m_hWnd) return;

    // Create status label
    m_hStatus = CreateWindowExW(
        0, L"STATIC", L"",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        10, 10, width - 20, 20,
        m_hWnd, (HMENU)IDC_STATUS,
        GetModuleHandle(NULL), NULL
    );

    // Create progress bar
    m_hProgress = CreateWindowExW(
        0, PROGRESS_CLASSW, NULL,
        WS_CHILD | WS_VISIBLE | PBS_SMOOTH,
        10, 35, width - 20, 20,
        m_hWnd, (HMENU)IDC_PROGRESS,
        GetModuleHandle(NULL), NULL
    );
    SendMessage(m_hProgress, PBM_SETRANGE, 0, MAKELPARAM(0, 1000));

    // Create cancel button
    m_hCancelBtn = CreateWindowExW(
        0, L"BUTTON", L"Cancel",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        (width - 80) / 2, 65, 80, 25,
        m_hWnd, (HMENU)IDC_CANCEL,
        GetModuleHandle(NULL), NULL
    );

    // Set font
    HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    SendMessage(m_hStatus, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(m_hCancelBtn, WM_SETFONT, (WPARAM)hFont, TRUE);

    // Store this pointer
    SetWindowLongPtr(m_hWnd, GWLP_USERDATA, (LONG_PTR)this);

    ShowWindow(m_hWnd, SW_SHOW);
    UpdateWindow(m_hWnd);
}

LRESULT CALLBACK CProgressDialog::WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    CProgressDialog* pThis = (CProgressDialog*)GetWindowLongPtr(hWnd, GWLP_USERDATA);

    switch (msg) {
        case WM_COMMAND:
            if (LOWORD(wParam) == IDC_CANCEL) {
                if (pThis) {
                    pThis->m_cancelled.store(true);
                }
            }
            break;

        case WM_CLOSE:
            if (pThis) {
                pThis->m_cancelled.store(true);
            }
            return 0;

        case WM_USER + 1:  // WM_UPDATE_PROGRESS
            if (pThis) {
                pThis->UpdateUI();
            }
            break;

        case WM_DESTROY:
            PostQuitMessage(0);
            break;
    }

    return DefWindowProc(hWnd, msg, wParam, lParam);
}

void CProgressDialog::SetProgress(uint64_t completed, uint64_t total) {
    m_completed = completed;
    m_total = total;

    if (m_hWnd) {
        PostMessage(m_hWnd, WM_USER + 1, 0, 0);
    }
}

void CProgressDialog::SetStatus(const std::wstring& status) {
    m_status = status;

    if (m_hWnd) {
        PostMessage(m_hWnd, WM_USER + 1, 0, 0);
    }
}

void CProgressDialog::UpdateUI() {
    if (!m_hWnd) return;

    // Update status
    if (m_hStatus) {
        SetWindowTextW(m_hStatus, m_status.c_str());
    }

    // Update progress bar
    if (m_hProgress && m_total > 0) {
        int pos = (int)((m_completed * 1000) / m_total);
        SendMessage(m_hProgress, PBM_SETPOS, pos, 0);
    }
}

void CProgressDialog::Close() {
    m_closed.store(true);

    if (m_hWnd) {
        PostMessage(m_hWnd, WM_CLOSE, 0, 0);
        DestroyWindow(m_hWnd);
        m_hWnd = NULL;
    }

    if (m_hThread) {
        // Wait for thread to finish
        WaitForSingleObject(m_hThread, 1000);
        CloseHandle(m_hThread);
        m_hThread = NULL;
    }
}
