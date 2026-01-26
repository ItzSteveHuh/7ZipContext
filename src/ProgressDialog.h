// ProgressDialog.h - Progress dialog for extraction/compression
#pragma once

#include <Windows.h>
#include <string>
#include <atomic>
#include <thread>
#include <functional>

class CProgressDialog {
public:
    CProgressDialog();
    ~CProgressDialog();

    // Show the dialog (creates window in separate thread)
    void Show(const std::wstring& title);

    // Update progress (0-100 or absolute values)
    void SetProgress(uint64_t completed, uint64_t total);

    // Update status text
    void SetStatus(const std::wstring& status);

    // Check if user cancelled
    bool IsCancelled() const { return m_cancelled.load(); }

    // Close the dialog
    void Close();

private:
    static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
    static DWORD WINAPI DialogThread(LPVOID lpParam);

    void CreateDialogWindow();
    void UpdateUI();

    HWND m_hWnd = NULL;
    HWND m_hProgress = NULL;
    HWND m_hStatus = NULL;
    HWND m_hCancelBtn = NULL;

    std::wstring m_title;
    std::wstring m_status;
    uint64_t m_completed = 0;
    uint64_t m_total = 0;

    std::atomic<bool> m_cancelled{ false };
    std::atomic<bool> m_closed{ false };
    HANDLE m_hThread = NULL;
    DWORD m_threadId = 0;

    static const int IDC_PROGRESS = 1001;
    static const int IDC_STATUS = 1002;
    static const int IDC_CANCEL = 1003;
    static const int WM_UPDATE_PROGRESS = WM_USER + 1;
};
