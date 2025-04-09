#include "pch.h"
#include "framework.h"
#include "resource.h"
#include "ComExplorer_i.h"

using namespace ATL;

class CComExplorerModule : public ATL::CAtlExeModuleT<CComExplorerModule>
{
public:
    DECLARE_LIBID(LIBID_ComExplorerLib)
    DECLARE_REGISTRY_APPID_RESOURCEID(IDR_COMEXPLORER, "{00c8a28a-f581-4b00-8663-bdb995221441}")
};

CComExplorerModule _AtlModule;

// Window Procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    default:
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
}

extern "C" int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE /*hPrevInstance*/, LPTSTR /*lpCmdLine*/, int nShowCmd)
{
    // Register the window class
    const wchar_t CLASS_NAME[] = L"Sample Window Class";

    WNDCLASS wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;

    if (!RegisterClass(&wc))
    {
        return 0;
    }

    // Create the window
    HWND hwnd = CreateWindowEx(
        0,                              // Optional window styles
        CLASS_NAME,                     // Window class
        L"ATL Window",                  // Window text
        WS_OVERLAPPEDWINDOW,            // Window style
        CW_USEDEFAULT, CW_USEDEFAULT,   // Position
        CW_USEDEFAULT, CW_USEDEFAULT,   // Size
        nullptr,                        // Parent window
        nullptr,                        // Menu
        hInstance,                      // Instance handle
        nullptr                         // Additional application data
    );

    if (!hwnd)
    {
        return 0;
    }

    ShowWindow(hwnd, nShowCmd);

    // Message loop
    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return static_cast<int>(msg.wParam);
}