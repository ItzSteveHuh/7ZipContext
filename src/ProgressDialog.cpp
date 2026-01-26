// ProgressDialog.cpp - Windows 11 style progress dialog implementation
#include "ProgressDialog.h"
#include <CommCtrl.h>
#include <algorithm>
#include <sstream>
#include <iomanip>
#include <cmath>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "uxtheme.lib")

// Window class names
static const wchar_t* PROGRESS_WND_CLASS = L"7ZipProgressDialog";
static const wchar_t* GRAPH_WND_CLASS = L"7ZipSpeedGraph";
static bool g_classRegistered = false;
static bool g_graphClassRegistered = false;

// Colors - Windows 11 green theme (exact match)
static const COLORREF CLR_WINDOW_BG = RGB(255, 255, 255);
static const COLORREF CLR_GRAPH_BG = RGB(237, 247, 237);        // Lighter green background
static const COLORREF CLR_GRAPH_BORDER = RGB(180, 180, 180);    // Gray border
static const COLORREF CLR_GRAPH_GRID = RGB(220, 240, 220);      // Subtle grid
static const COLORREF CLR_GRAPH_LINE = RGB(56, 142, 60);        // Darker green line (curve)
static const COLORREF CLR_GRAPH_FILL = RGB(165, 214, 167);      // Soft green fill
static const COLORREF CLR_SPEED_LINE = RGB(0, 0, 0);            // Black speed line and text
static const COLORREF CLR_TEXT_PRIMARY = RGB(32, 32, 32);
static const COLORREF CLR_TEXT_SECONDARY = RGB(102, 102, 102);
static const COLORREF CLR_BTN_BORDER = RGB(204, 204, 204);
static const COLORREF CLR_BTN_HOVER = RGB(243, 243, 243);

CProgressDialog::CProgressDialog() {
    m_speedHistory.reserve(SPEED_HISTORY_SIZE);
}

CProgressDialog::~CProgressDialog() {
    Close();
    if (m_hSmallFont) DeleteObject(m_hSmallFont);
    if (m_hNormalFont) DeleteObject(m_hNormalFont);
    if (m_hLargeFont) DeleteObject(m_hLargeFont);
    if (m_hIconFont) DeleteObject(m_hIconFont);
}

void CProgressDialog::Show(const std::wstring& title, const std::wstring& operation) {
    m_title = title;
    m_operation = operation;
    m_cancelled.store(false);
    m_paused.store(false);
    m_closed.store(false);
    m_startTime = std::chrono::steady_clock::now();

    {
        std::lock_guard<std::mutex> lock(m_dataMutex);
        m_samples.clear();
        m_speedHistory.clear();
        m_maxSpeedInHistory = 0;
    }

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
    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icc);

    HINSTANCE hInst = GetModuleHandle(NULL);

    // Register main window class
    if (!g_classRegistered) {
        WNDCLASSEXW wc = {};
        wc.cbSize = sizeof(wc);
        wc.style = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc = WndProc;
        wc.hInstance = hInst;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = CreateSolidBrush(CLR_WINDOW_BG);
        wc.lpszClassName = PROGRESS_WND_CLASS;
        RegisterClassExW(&wc);
        g_classRegistered = true;
    }

    // Register graph panel class
    if (!g_graphClassRegistered) {
        WNDCLASSEXW wc = {};
        wc.cbSize = sizeof(wc);
        wc.style = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc = WndProc;
        wc.hInstance = hInst;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = CreateSolidBrush(CLR_GRAPH_BG);
        wc.lpszClassName = GRAPH_WND_CLASS;
        RegisterClassExW(&wc);
        g_graphClassRegistered = true;
    }

    // Center on screen
    int x = (GetSystemMetrics(SM_CXSCREEN) - m_windowWidth) / 2;
    int y = (GetSystemMetrics(SM_CYSCREEN) - m_windowHeight) / 2;

    // Create main window
    m_hWnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        PROGRESS_WND_CLASS,
        m_title.c_str(),
        WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_CLIPCHILDREN,
        x, y, m_windowWidth, m_windowHeight,
        NULL, NULL, hInst, this
    );

    if (!m_hWnd) return;

    // Create fonts - use system font metrics
    HDC hdc = GetDC(m_hWnd);
    int dpi = GetDeviceCaps(hdc, LOGPIXELSY);
    ReleaseDC(m_hWnd, hdc);

    int smallSize = -MulDiv(9, dpi, 72);
    int normalSize = -MulDiv(10, dpi, 72);
    int largeSize = -MulDiv(14, dpi, 72);

    m_hSmallFont = CreateFontW(smallSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    m_hNormalFont = CreateFontW(normalSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    m_hLargeFont = CreateFontW(largeSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    m_hIconFont = CreateFontW(normalSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI Symbol");

    int margin = 20;
    int yPos = 16;
    int contentWidth = m_windowWidth - margin * 2;

    // Operation label (small gray text at top)
    m_hOperationLabel = CreateWindowExW(0, L"STATIC",
        m_operation.empty() ? L"正在处理..." : m_operation.c_str(),
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        margin, yPos, contentWidth, 18,
        m_hWnd, (HMENU)IDC_OPERATION, hInst, NULL);
    SendMessage(m_hOperationLabel, WM_SETFONT, (WPARAM)m_hSmallFont, TRUE);

    yPos += 26;

    // Large percent label (left side)
    m_hPercentLabel = CreateWindowExW(0, L"STATIC", L"已完成 0%",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        margin, yPos, contentWidth - 70, 28,
        m_hWnd, (HMENU)IDC_PERCENT, hInst, NULL);
    SendMessage(m_hPercentLabel, WM_SETFONT, (WPARAM)m_hLargeFont, TRUE);

    // Pause button (inline with percent, right side)
    m_hPauseBtn = CreateWindowExW(0, L"BUTTON", L"| |",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        m_windowWidth - margin - 58, yPos + 2, 26, 24,
        m_hWnd, (HMENU)IDC_PAUSE, hInst, NULL);

    // Cancel button (inline with percent, right side)
    m_hCancelBtn = CreateWindowExW(0, L"BUTTON", L"\u00D7",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        m_windowWidth - margin - 28, yPos + 2, 26, 24,
        m_hWnd, (HMENU)IDC_CANCEL, hInst, NULL);

    yPos += 36;

    // Graph panel (larger height for Windows 11 look)
    m_hGraphPanel = CreateWindowExW(0, GRAPH_WND_CLASS, NULL,
        WS_CHILD | WS_VISIBLE,
        margin, yPos, contentWidth, 120,
        m_hWnd, (HMENU)IDC_GRAPH, hInst, NULL);
    SetWindowLongPtr(m_hGraphPanel, GWLP_USERDATA, (LONG_PTR)this);
    yPos += 130;

    // Name label
    m_hNameLabel = CreateWindowExW(0, L"STATIC", L"名称:",
        WS_CHILD | WS_VISIBLE | SS_LEFT | SS_PATHELLIPSIS,
        margin, yPos, contentWidth, 18,
        m_hWnd, (HMENU)IDC_NAME, hInst, NULL);
    SendMessage(m_hNameLabel, WM_SETFONT, (WPARAM)m_hSmallFont, TRUE);
    yPos += 22;

    // Time remaining label
    m_hTimeLabel = CreateWindowExW(0, L"STATIC", L"剩余时间: 正在计算...",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        margin, yPos, contentWidth, 18,
        m_hWnd, (HMENU)IDC_TIME, hInst, NULL);
    SendMessage(m_hTimeLabel, WM_SETFONT, (WPARAM)m_hSmallFont, TRUE);
    yPos += 22;

    // Items remaining label
    m_hItemsLabel = CreateWindowExW(0, L"STATIC", L"剩余项目: 0 (0 字节)",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        margin, yPos, contentWidth, 18,
        m_hWnd, (HMENU)IDC_ITEMS, hInst, NULL);
    SendMessage(m_hItemsLabel, WM_SETFONT, (WPARAM)m_hSmallFont, TRUE);
    yPos += 28;

    // Separator line above details button - full width, 1 pixel
    m_hSeparator = CreateWindowExW(0, L"STATIC", NULL,
        WS_CHILD | WS_VISIBLE | SS_OWNERDRAW,
        0, yPos, m_windowWidth, 1,
        m_hWnd, (HMENU)IDC_SEPARATOR, hInst, NULL);
    yPos += 12;

    // Details toggle button (left aligned with arrow)
    m_hDetailsBtn = CreateWindowExW(0, L"BUTTON", L"\u2227  简略信息",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        margin, yPos, 100, 22,
        m_hWnd, (HMENU)IDC_DETAILS, hInst, NULL);

    SetWindowLongPtr(m_hWnd, GWLP_USERDATA, (LONG_PTR)this);

    // Fast update timer - 50ms for smooth animation
    SetTimer(m_hWnd, TIMER_UPDATE, 50, NULL);

    ShowWindow(m_hWnd, SW_SHOW);
    UpdateWindow(m_hWnd);
}

LRESULT CALLBACK CProgressDialog::WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    CProgressDialog* pThis = (CProgressDialog*)GetWindowLongPtr(hWnd, GWLP_USERDATA);

    switch (msg) {
        case WM_COMMAND:
            if (pThis) {
                switch (LOWORD(wParam)) {
                    case IDC_CANCEL:
                        pThis->m_cancelled.store(true);
                        EnableWindow(pThis->m_hCancelBtn, FALSE);
                        EnableWindow(pThis->m_hPauseBtn, FALSE);
                        SetWindowTextW(pThis->m_hOperationLabel, L"正在取消...");
                        break;
                    case IDC_PAUSE:
                        {
                            bool paused = !pThis->m_paused.load();
                            pThis->m_paused.store(paused);
                            InvalidateRect(pThis->m_hPauseBtn, NULL, TRUE);
                        }
                        break;
                    case IDC_DETAILS:
                        pThis->m_detailsExpanded = !pThis->m_detailsExpanded;
                        pThis->ToggleDetails();
                        break;
                }
            }
            break;

        case WM_DRAWITEM:
            if (pThis) {
                LPDRAWITEMSTRUCT pDIS = (LPDRAWITEMSTRUCT)lParam;
                if (pDIS->CtlID == IDC_SEPARATOR) {
                    // Draw 1-pixel separator line
                    HBRUSH hBrush = CreateSolidBrush(RGB(229, 229, 229));
                    FillRect(pDIS->hDC, &pDIS->rcItem, hBrush);
                    DeleteObject(hBrush);
                } else {
                    pThis->DrawButton(pDIS);
                }
                return TRUE;
            }
            break;

        case WM_TIMER:
            if (wParam == TIMER_UPDATE && pThis) {
                pThis->CalculateSpeed();
                pThis->UpdateUI();
                InvalidateRect(pThis->m_hGraphPanel, NULL, FALSE);
            }
            break;

        case WM_PAINT:
            if (pThis && hWnd == pThis->m_hGraphPanel) {
                PAINTSTRUCT ps;
                HDC hdc = BeginPaint(hWnd, &ps);
                RECT rect;
                GetClientRect(hWnd, &rect);
                pThis->PaintGraph(hdc, rect);
                EndPaint(hWnd, &ps);
                return 0;
            }
            break;

        case WM_ERASEBKGND:
            if (pThis && hWnd == pThis->m_hGraphPanel) {
                return 1;
            }
            break;

        case WM_CLOSE:
            if (pThis) {
                pThis->m_cancelled.store(true);
            }
            return 0;

        case WM_CTLCOLORSTATIC:
            {
                HDC hdcStatic = (HDC)wParam;
                HWND hStatic = (HWND)lParam;
                SetBkColor(hdcStatic, CLR_WINDOW_BG);
                // Operation label uses secondary (gray) color
                if (pThis && hStatic == pThis->m_hOperationLabel) {
                    SetTextColor(hdcStatic, CLR_TEXT_SECONDARY);
                } else {
                    SetTextColor(hdcStatic, CLR_TEXT_PRIMARY);
                }
                static HBRUSH hBrush = CreateSolidBrush(CLR_WINDOW_BG);
                return (LRESULT)hBrush;
            }

        case WM_CTLCOLORBTN:
            {
                static HBRUSH hBrush = CreateSolidBrush(CLR_WINDOW_BG);
                return (LRESULT)hBrush;
            }

        case WM_DESTROY:
            KillTimer(hWnd, TIMER_UPDATE);
            PostQuitMessage(0);
            break;
    }

    return DefWindowProc(hWnd, msg, wParam, lParam);
}

void CProgressDialog::ToggleDetails() {
    if (!m_hWnd) return;

    // Update button text
    InvalidateRect(m_hDetailsBtn, NULL, TRUE);

    // Show/hide detail controls and resize window
    BOOL show = m_detailsExpanded ? SW_SHOW : SW_HIDE;
    ShowWindow(m_hNameLabel, show);
    ShowWindow(m_hTimeLabel, show);
    ShowWindow(m_hItemsLabel, show);
    ShowWindow(m_hSeparator, show);

    // Resize window
    RECT rc;
    GetWindowRect(m_hWnd, &rc);
    int newHeight = m_detailsExpanded ? m_windowHeight : m_collapsedHeight;
    SetWindowPos(m_hWnd, NULL, 0, 0, rc.right - rc.left, newHeight,
                 SWP_NOMOVE | SWP_NOZORDER);
}

void CProgressDialog::DrawButton(LPDRAWITEMSTRUCT pDIS) {
    HDC hdc = pDIS->hDC;
    RECT rc = pDIS->rcItem;
    bool isPressed = (pDIS->itemState & ODS_SELECTED) != 0;
    bool isDisabled = (pDIS->itemState & ODS_DISABLED) != 0;

    // Get mouse position to check hover state
    POINT pt;
    GetCursorPos(&pt);
    ScreenToClient(pDIS->hwndItem, &pt);
    bool isHover = PtInRect(&rc, pt);

    // Fill background
    COLORREF bgColor = CLR_WINDOW_BG;
    if (isPressed) {
        bgColor = RGB(230, 230, 230);
    } else if (isHover) {
        bgColor = CLR_BTN_HOVER;
    }
    HBRUSH hBgBrush = CreateSolidBrush(bgColor);
    FillRect(hdc, &rc, hBgBrush);
    DeleteObject(hBgBrush);

    // Draw border for Pause/Cancel buttons ONLY on hover (Windows 11 style)
    if ((pDIS->CtlID == IDC_PAUSE || pDIS->CtlID == IDC_CANCEL) && (isHover || isPressed)) {
        HPEN hPen = CreatePen(PS_SOLID, 1, CLR_BTN_BORDER);
        HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
        HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));

        RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, 4, 4);

        SelectObject(hdc, hOldBrush);
        SelectObject(hdc, hOldPen);
        DeleteObject(hPen);
    }

    // Draw text
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, isDisabled ? RGB(180, 180, 180) : CLR_TEXT_PRIMARY);

    wchar_t text[64] = {};
    GetWindowTextW(pDIS->hwndItem, text, 64);

    HFONT hFont = m_hSmallFont;
    if (pDIS->CtlID == IDC_PAUSE) {
        hFont = m_hNormalFont;
        if (m_paused.load()) {
            wcscpy_s(text, L"\u25B6");  // Play triangle
        } else {
            wcscpy_s(text, L"||");  // Simple pause bars
        }
    } else if (pDIS->CtlID == IDC_CANCEL) {
        hFont = m_hNormalFont;
        wcscpy_s(text, L"\u00D7");  // X mark
    } else if (pDIS->CtlID == IDC_DETAILS) {
        SetTextColor(hdc, RGB(0, 102, 204));  // Blue link color
        wcscpy_s(text, m_detailsExpanded ? L"\u2227  简略信息" : L"\u2228  详细信息");
    }

    HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);
    DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, hOldFont);
}

void CProgressDialog::PaintGraph(HDC hdc, const RECT& rect) {
    int width = rect.right - rect.left;
    int height = rect.bottom - rect.top;

    // Double buffer
    HDC hdcMem = CreateCompatibleDC(hdc);
    HBITMAP hbmMem = CreateCompatibleBitmap(hdc, width, height);
    HBITMAP hbmOld = (HBITMAP)SelectObject(hdcMem, hbmMem);

    // Fill background
    HBRUSH hBgBrush = CreateSolidBrush(CLR_GRAPH_BG);
    FillRect(hdcMem, &rect, hBgBrush);
    DeleteObject(hBgBrush);

    // Draw border around the graph (Windows 11 style)
    HPEN hBorderPen = CreatePen(PS_SOLID, 1, CLR_GRAPH_BORDER);
    HPEN hOldPen = (HPEN)SelectObject(hdcMem, hBorderPen);
    HBRUSH hOldBrush = (HBRUSH)SelectObject(hdcMem, GetStockObject(NULL_BRUSH));
    Rectangle(hdcMem, 0, 0, width, height);
    SelectObject(hdcMem, hOldBrush);
    SelectObject(hdcMem, hOldPen);
    DeleteObject(hBorderPen);

    // Draw fine grid - smaller spacing for Windows 11 look
    HPEN hGridPen = CreatePen(PS_SOLID, 1, CLR_GRAPH_GRID);
    hOldPen = (HPEN)SelectObject(hdcMem, hGridPen);

    int gridSpacing = 12;  // Finer grid
    for (int y = gridSpacing; y < height; y += gridSpacing) {
        MoveToEx(hdcMem, 1, y, NULL);
        LineTo(hdcMem, width - 1, y);
    }
    for (int x = gridSpacing; x < width; x += gridSpacing) {
        MoveToEx(hdcMem, x, 1, NULL);
        LineTo(hdcMem, x, height - 1);
    }

    SelectObject(hdcMem, hOldPen);
    DeleteObject(hGridPen);

    // Draw speed curve
    std::vector<double> history;
    double maxSpeed;
    double currentSpeed;
    {
        std::lock_guard<std::mutex> lock(m_dataMutex);
        history = m_speedHistory;
        maxSpeed = m_maxSpeedInHistory;
        currentSpeed = m_currentSpeed;
    }

    // Ensure maxSpeed has a reasonable value for drawing
    if (maxSpeed < 1.0) maxSpeed = 1.0;
    if (currentSpeed > maxSpeed) maxSpeed = currentSpeed * 1.1;

    int graphHeight = height - 4;  // Leave space for border
    int currentSpeedY = -1;  // Y position for current speed line

    // Calculate current speed Y position
    if (currentSpeed > 0) {
        currentSpeedY = height - 2 - (int)(currentSpeed / maxSpeed * graphHeight);
        currentSpeedY = (std::max)(2, (std::min)(height - 2, currentSpeedY));
    }

    if (!history.empty()) {
        // Create smoothed points
        std::vector<POINT> points;
        std::vector<POINT> fillPoints;

        fillPoints.push_back({ 1, height - 1 });

        for (size_t i = 0; i < history.size(); i++) {
            int px = 1 + (int)(i * (width - 2) / SPEED_HISTORY_SIZE);
            double smoothedSpeed = history[i];

            // Simple smoothing with neighbors
            if (i > 0 && i < history.size() - 1) {
                smoothedSpeed = (history[i-1] + history[i] * 2 + history[i+1]) / 4.0;
            }

            int py = height - 2 - (int)(smoothedSpeed / maxSpeed * graphHeight);
            py = (std::max)(2, (std::min)(height - 2, py));
            points.push_back({ px, py });
            fillPoints.push_back({ px, py });
        }

        if (!points.empty()) {
            fillPoints.push_back({ points.back().x, height - 1 });
        }

        // Fill area under curve with gradient-like effect
        if (fillPoints.size() >= 3) {
            HBRUSH hFillBrush = CreateSolidBrush(CLR_GRAPH_FILL);
            HPEN hNullPen = CreatePen(PS_NULL, 0, 0);
            hOldBrush = (HBRUSH)SelectObject(hdcMem, hFillBrush);
            SelectObject(hdcMem, hNullPen);
            Polygon(hdcMem, fillPoints.data(), (int)fillPoints.size());
            SelectObject(hdcMem, hOldBrush);
            DeleteObject(hFillBrush);
            DeleteObject(hNullPen);
        }

        // Draw the curve line
        if (points.size() >= 2) {
            HPEN hLinePen = CreatePen(PS_SOLID, 2, CLR_GRAPH_LINE);
            SelectObject(hdcMem, hLinePen);
            Polyline(hdcMem, points.data(), (int)points.size());
            DeleteObject(hLinePen);
        }
    }

    // Draw current speed horizontal line and text
    if (currentSpeed > 0 && currentSpeedY >= 2) {
        std::wstring speedText = L"速度: " + FormatSpeed(currentSpeed);

        HFONT hOldFont = (HFONT)SelectObject(hdcMem, m_hSmallFont);
        SetBkMode(hdcMem, TRANSPARENT);
        SetTextColor(hdcMem, CLR_SPEED_LINE);

        // Measure text size
        SIZE textSize;
        GetTextExtentPoint32W(hdcMem, speedText.c_str(), (int)speedText.length(), &textSize);

        // Draw horizontal line across full width first
        HPEN hSpeedLinePen = CreatePen(PS_SOLID, 1, CLR_SPEED_LINE);
        hOldPen = (HPEN)SelectObject(hdcMem, hSpeedLinePen);
        MoveToEx(hdcMem, 1, currentSpeedY, NULL);
        LineTo(hdcMem, width - 1, currentSpeedY);
        SelectObject(hdcMem, hOldPen);
        DeleteObject(hSpeedLinePen);

        // Calculate text position - right side, ABOVE the line
        int textX = width - textSize.cx - 6;
        int textY = currentSpeedY - textSize.cy - 2;  // Above the line
        // Ensure text stays within bounds
        if (textY < 2) {
            textY = currentSpeedY + 2;  // If no room above, put below
        }

        // Draw speed text
        TextOutW(hdcMem, textX, textY, speedText.c_str(), (int)speedText.length());

        SelectObject(hdcMem, hOldFont);
    }

    // Copy to screen
    BitBlt(hdc, 0, 0, width, height, hdcMem, 0, 0, SRCCOPY);

    SelectObject(hdcMem, hbmOld);
    DeleteObject(hbmMem);
    DeleteDC(hdcMem);
}

void CProgressDialog::CalculateSpeed() {
    std::lock_guard<std::mutex> lock(m_dataMutex);

    auto now = std::chrono::steady_clock::now();

    SpeedSample sample = { m_completed, now };
    m_samples.push_back(sample);

    // Keep samples from last 2 seconds for smoother calculation
    auto cutoff = now - std::chrono::milliseconds(2000);
    while (m_samples.size() > 1 && m_samples.front().time < cutoff) {
        m_samples.erase(m_samples.begin());
    }

    // Calculate speed
    if (m_samples.size() >= 2) {
        auto& first = m_samples.front();
        auto& last = m_samples.back();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(last.time - first.time);
        if (duration.count() > 0) {
            uint64_t diff = last.bytes > first.bytes ? last.bytes - first.bytes : 0;
            m_currentSpeed = diff * 1000.0 / duration.count();
        }
    }

    // Update history more frequently
    m_speedHistory.push_back(m_currentSpeed);
    while (m_speedHistory.size() > SPEED_HISTORY_SIZE) {
        m_speedHistory.erase(m_speedHistory.begin());
    }

    // Update max with decay
    double newMax = 1.0;
    for (double s : m_speedHistory) {
        if (s > newMax) newMax = s;
    }
    // Smooth max transition
    if (newMax > m_maxSpeedInHistory) {
        m_maxSpeedInHistory = newMax * 1.1;
    } else {
        m_maxSpeedInHistory = m_maxSpeedInHistory * 0.99 + newMax * 0.01 * 1.1;
    }
    if (m_maxSpeedInHistory < 1.0) m_maxSpeedInHistory = 1.0;

    // Estimate time
    if (m_currentSpeed > 100 && m_total > m_completed) {
        m_estimatedSeconds = (int)((m_total - m_completed) / m_currentSpeed);
    } else {
        m_estimatedSeconds = -1;
    }
}

void CProgressDialog::SetProgress(uint64_t completed, uint64_t total) {
    m_completed = completed;
    m_total = total;
}

void CProgressDialog::SetCurrentItem(const std::wstring& itemName) {
    m_currentItem = itemName;
}

void CProgressDialog::UpdateUI() {
    if (!m_hWnd) return;

    // Update window title and percent label
    int percent = m_total > 0 ? (int)(m_completed * 100 / m_total) : 0;

    std::wstring titleText = L"已完成 " + std::to_wstring(percent) + L"%";
    SetWindowTextW(m_hWnd, titleText.c_str());
    SetWindowTextW(m_hPercentLabel, titleText.c_str());

    // Update name label
    if (m_hNameLabel && !m_currentItem.empty()) {
        std::wstring nameText = L"名称: " + m_currentItem;
        SetWindowTextW(m_hNameLabel, nameText.c_str());
    }

    // Update time label
    if (m_hTimeLabel) {
        std::wstring timeText;
        if (m_estimatedSeconds > 0) {
            timeText = L"剩余时间: " + FormatTime(m_estimatedSeconds);
        } else if (m_completed >= m_total && m_total > 0) {
            timeText = L"剩余时间: 完成";
        } else {
            timeText = L"剩余时间: 正在计算...";
        }
        SetWindowTextW(m_hTimeLabel, timeText.c_str());
    }

    // Update items label
    if (m_hItemsLabel) {
        uint64_t remaining = m_total > m_completed ? m_total - m_completed : 0;
        std::wstring itemsText = L"剩余项目: " + FormatSize(remaining);
        SetWindowTextW(m_hItemsLabel, itemsText.c_str());
    }
}

std::wstring CProgressDialog::FormatSize(uint64_t bytes) {
    std::wstringstream ss;
    ss << std::fixed << std::setprecision(1);

    if (bytes >= 1099511627776ULL) {
        ss << (bytes / 1099511627776.0) << L" TB";
    } else if (bytes >= 1073741824ULL) {
        ss << (bytes / 1073741824.0) << L" GB";
    } else if (bytes >= 1048576ULL) {
        ss << (bytes / 1048576.0) << L" MB";
    } else if (bytes >= 1024ULL) {
        ss << (bytes / 1024.0) << L" KB";
    } else {
        ss << bytes << L" 字节";
    }
    return ss.str();
}

std::wstring CProgressDialog::FormatTime(int seconds) {
    if (seconds < 0) return L"正在计算...";
    if (seconds < 60) {
        return std::to_wstring(seconds) + L" 秒";
    } else if (seconds < 3600) {
        int mins = seconds / 60;
        int secs = seconds % 60;
        return std::to_wstring(mins) + L" 分 " + std::to_wstring(secs) + L" 秒";
    } else {
        int hours = seconds / 3600;
        int mins = (seconds % 3600) / 60;
        return std::to_wstring(hours) + L" 小时 " + std::to_wstring(mins) + L" 分";
    }
}

std::wstring CProgressDialog::FormatSpeed(double bytesPerSec) {
    std::wstringstream ss;
    ss << std::fixed << std::setprecision(1);

    if (bytesPerSec >= 1073741824.0) {
        ss << (bytesPerSec / 1073741824.0) << L" GB/s";
    } else if (bytesPerSec >= 1048576.0) {
        ss << (bytesPerSec / 1048576.0) << L" MB/s";
    } else if (bytesPerSec >= 1024.0) {
        ss << (bytesPerSec / 1024.0) << L" KB/s";
    } else {
        ss << bytesPerSec << L" B/s";
    }
    return ss.str();
}

void CProgressDialog::Close() {
    m_closed.store(true);

    if (m_hWnd) {
        KillTimer(m_hWnd, TIMER_UPDATE);
        PostMessage(m_hWnd, WM_CLOSE, 0, 0);
        DestroyWindow(m_hWnd);
        m_hWnd = NULL;
    }

    if (m_hThread) {
        WaitForSingleObject(m_hThread, 2000);
        CloseHandle(m_hThread);
        m_hThread = NULL;
    }
}
