// ProgressDialog.h - Windows 11 style progress dialog with speed graph
#pragma once

#include <Windows.h>
#include <string>
#include <atomic>
#include <vector>
#include <chrono>
#include <mutex>

class CProgressDialog {
public:
    CProgressDialog();
    ~CProgressDialog();

    // Show the dialog
    void Show(const std::wstring& title, const std::wstring& operation = L"");

    // Update progress
    void SetProgress(uint64_t completed, uint64_t total);

    // Update current item name
    void SetCurrentItem(const std::wstring& itemName);

    // Legacy SetStatus for compatibility
    void SetStatus(const std::wstring& status) { SetCurrentItem(status); }

    // Check if user cancelled
    bool IsCancelled() const { return m_cancelled.load(); }

    // Check if paused
    bool IsPaused() const { return m_paused.load(); }

    // Close the dialog
    void Close();

private:
    static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
    static DWORD WINAPI DialogThread(LPVOID lpParam);

    void CreateDialogWindow();
    void UpdateUI();
    void PaintGraph(HDC hdc, const RECT& rect);
    void DrawButton(LPDRAWITEMSTRUCT pDIS);
    void ToggleDetails();
    void CalculateSpeed();
    std::wstring FormatSize(uint64_t bytes);
    std::wstring FormatTime(int seconds);
    std::wstring FormatSpeed(double bytesPerSec);

    // Window handles
    HWND m_hWnd = NULL;
    HWND m_hOperationLabel = NULL;      // "正在将 X 个项目..."
    HWND m_hPercentLabel = NULL;        // "已完成 78%"
    HWND m_hPauseBtn = NULL;            // || button
    HWND m_hCancelBtn = NULL;           // X button
    HWND m_hGraphPanel = NULL;          // Green graph area
    HWND m_hNameLabel = NULL;           // "名称: xxx"
    HWND m_hTimeLabel = NULL;           // "剩余时间: xxx"
    HWND m_hItemsLabel = NULL;          // "剩余项目: xxx"
    HWND m_hSeparator = NULL;           // Separator line
    HWND m_hDetailsBtn = NULL;          // "简略信息" toggle

    // State
    std::wstring m_title;
    std::wstring m_operation;
    std::wstring m_currentItem;
    uint64_t m_completed = 0;
    uint64_t m_total = 0;
    uint32_t m_itemCount = 0;
    uint32_t m_processedItems = 0;

    std::atomic<bool> m_cancelled{ false };
    std::atomic<bool> m_paused{ false };
    std::atomic<bool> m_closed{ false };
    bool m_detailsExpanded = true;

    // Thread
    HANDLE m_hThread = NULL;
    DWORD m_threadId = 0;

    // Speed calculation
    struct SpeedSample {
        uint64_t bytes;
        std::chrono::steady_clock::time_point time;
    };
    std::vector<SpeedSample> m_samples;
    std::vector<double> m_speedHistory;
    std::mutex m_dataMutex;
    double m_currentSpeed = 0;
    int m_estimatedSeconds = -1;
    std::chrono::steady_clock::time_point m_startTime;

    static const size_t SPEED_HISTORY_SIZE = 100;
    double m_maxSpeedInHistory = 0;

    // Fonts
    HFONT m_hSmallFont = NULL;
    HFONT m_hNormalFont = NULL;
    HFONT m_hLargeFont = NULL;
    HFONT m_hIconFont = NULL;

    // Window dimensions
    int m_windowWidth = 480;
    int m_windowHeight = 390;
    int m_collapsedHeight = 270;

    // Control IDs
    enum {
        IDC_OPERATION = 1001,
        IDC_PERCENT,
        IDC_PAUSE,
        IDC_CANCEL,
        IDC_GRAPH,
        IDC_NAME,
        IDC_TIME,
        IDC_ITEMS,
        IDC_SEPARATOR,
        IDC_DETAILS,
    };

    static const UINT WM_UPDATE_PROGRESS = WM_USER + 1;
    static const UINT TIMER_UPDATE = 1;
};
