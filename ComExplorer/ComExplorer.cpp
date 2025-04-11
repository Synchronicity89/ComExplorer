#include "pch.h"
#include "framework.h"
#include "resource.h"
#include "ComExplorer_i.h"
#include <windows.h>
#include <atlbase.h>
#include <atlwin.h>
#include <atlcom.h>
#include <commctrl.h>
#pragma comment(lib, "comctl32.lib")

using namespace ATL;

// Resource IDs for our controls
#define IDC_LIST_OBJECTS   101
#define IDC_RADIO_DISPATCH 102
#define IDC_RADIO_DIRECT   103
#define IDC_BTN_CREATE     104
#define IDC_BTN_INVOKE     105
#define IDC_EDIT_OUTPUT    106
#define IDC_PANEL          107
#define IDC_BTN_LIST_MEMBERS 108
#define IDC_LIST_MEMBERS     109


// Structure to hold basic information about a COM object for testing.
struct ComObjectInfo {
    const wchar_t* displayName; // Friendly name as shown in the list box.
    const wchar_t* progID;      // ProgID (or CLSID string) used to create the object.
    bool hasUI;                 // Whether the object has a visual interface.
};

// For demonstration we hard-code two sample objects.
ComObjectInfo g_ComObjects[] = {
    { L"Scripting.Dictionary", L"Scripting.Dictionary", false },
    { L"Microsoft Web Browser", L"Shell.Explorer", true }
};

const int g_NumComObjects = sizeof(g_ComObjects) / sizeof(g_ComObjects[0]);

// Global variables:
HINSTANCE g_hInst = NULL;

// Handles for various controls in our main window.
HWND hList = NULL;
HWND hRadioDispatch = NULL;
HWND hRadioDirect = NULL;
HWND hButtonCreate = NULL;
HWND hButtonInvoke = NULL;
HWND hEditOutput = NULL;
HWND hPanel = NULL;
HWND hButtonListMembers = NULL;
HWND hListMembers = NULL;

// Global COM pointer for storing the automation object’s IDispatch interface.
CComPtr<IDispatch> g_pDispatch;

void InvokeDispatchMethod()
{
    if (!g_pDispatch)
    {
        MessageBox(NULL, L"No COM object available for invocation.", L"Error", MB_OK);
        return;
    }

    // Wrap the IDispatch pointer in a CComDispatchDriver.
    CComDispatchDriver dispatchDriver(g_pDispatch);

    // Look up the DISPID for the "Count" property.
    CComVariant result;
    HRESULT hr = dispatchDriver.GetPropertyByName(L"Count", &result);
    if (SUCCEEDED(hr))
    {
        // Expect an integer result.
        if (result.vt == VT_I4)
        {
            wchar_t buffer[100];
            swprintf_s(buffer, L"'Count' returned: %d", result.intVal);
            MessageBox(NULL, buffer, L"Success", MB_OK);
        }
        else
        {
            MessageBox(NULL, L"'Count' returned a value of an unexpected type.", L"Error", MB_OK);
        }
    }
    else
    {
        MessageBox(NULL, L"Failed to invoke the 'Count' property.", L"Error", MB_OK);
    }
}


void CreateCOMObject(const ComObjectInfo& info, bool useDispatch)
{
    // Clear any previous instance.
    g_pDispatch.Release();

    // Convert the ProgID to a CLSID.
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(info.progID, &clsid);
    if (FAILED(hr))
    {
        MessageBox(NULL, L"Failed to retrieve CLSID from ProgID.", L"Error", MB_OK);
        return;
    }

    if (useDispatch)
    {
        // Use CComPtr to create the COM object and request IDispatch.
        hr = g_pDispatch.CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER);
        if (SUCCEEDED(hr))
        {
            MessageBox(NULL, L"COM object created with IDispatch successfully.", L"Success", MB_OK);
        }
        else
        {
            MessageBox(NULL, L"Failed to create COM object with IDispatch.", L"Error", MB_OK);
        }
    }
    else
    {
        // Use CComPtr to create the COM object and request IUnknown.
        CComPtr<IUnknown> pUnk;
        hr = pUnk.CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER);
        if (SUCCEEDED(hr))
        {
            MessageBox(NULL, L"COM object created with IUnknown successfully.", L"Success", MB_OK);

            // Optionally query for IDispatch.
            hr = pUnk->QueryInterface(IID_IDispatch, (void**)&g_pDispatch);
            if (SUCCEEDED(hr))
            {
                MessageBox(NULL, L"IDispatch is also available on this object.", L"Success", MB_OK);
            }
            else
            {
                MessageBox(NULL, L"IDispatch is not available on this object.", L"Error", MB_OK);
            }
        }
        else
        {
            MessageBox(NULL, L"Failed to create COM object via IUnknown.", L"Error", MB_OK);
        }
    }
}


void CreateActiveXControl(const wchar_t* progID)
{
    // Ensure the panel handle is valid.
    if (!hPanel)
    {
        MessageBox(NULL, L"Panel handle is invalid.", L"Error", MB_OK);
        return;
    }

    // Clear any existing child windows in the panel.
    HWND hChild = GetWindow(hPanel, GW_CHILD);
    while (hChild)
    {
        DestroyWindow(hChild);
        hChild = GetWindow(hPanel, GW_CHILD);
    }

    // Create the ActiveX control.
    CComPtr<IAxWinHostWindow> spHost;
    CComPtr<IUnknown> spControl;

    // Initialize ATL module if not already initialized.
    AtlAxWinInit();

    // Create the ActiveX container window.
    HWND hAxWindow = CreateWindowEx(0, CAxWindow::GetWndClassName(), NULL,
        WS_CHILD | WS_VISIBLE, 0, 0, 400, 330, hPanel, NULL, g_hInst, NULL);

    if (!hAxWindow)
    {
        MessageBox(NULL, L"Failed to create ActiveX container window.", L"Error", MB_OK);
        return;
    }

    // Attach the window to an ATL CAxWindow object.
    CAxWindow axWindow(hAxWindow);

    // Query for the IAxWinHostWindow interface.
    HRESULT hr = axWindow.QueryHost(&spHost);
    if (FAILED(hr))
    {
        MessageBox(NULL, L"Failed to query IAxWinHostWindow interface.", L"Error", MB_OK);
        return;
    }

    // Create the ActiveX control using the provided ProgID.
    hr = spHost->CreateControl(progID, hAxWindow, NULL);
    if (FAILED(hr))
    {
        MessageBox(NULL, L"Failed to create ActiveX control.", L"Error", MB_OK);
        return;
    }

    // Query for the control's IUnknown interface.
    hr = axWindow.QueryControl(&spControl);
    if (FAILED(hr))
    {
        MessageBox(NULL, L"Failed to query ActiveX control interface.", L"Error", MB_OK);
        return;
    }

    // Successfully created and hosted the ActiveX control.
    MessageBox(NULL, L"ActiveX control created successfully.", L"Success", MB_OK);
}


// Helper functions and WndProc implementation (from the provided code)...
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
    {
        // Create a ListBox for selecting the COM object.
        hList = CreateWindowEx(WS_EX_CLIENTEDGE, L"LISTBOX", NULL,
            WS_CHILD | WS_VISIBLE | LBS_STANDARD,
            10, 10, 250, 100,
            hWnd, (HMENU)IDC_LIST_OBJECTS, g_hInst, NULL);
        // Populate the list box.
        for (int i = 0; i < g_NumComObjects; i++)
        {
            SendMessage(hList, LB_ADDSTRING, 0, (LPARAM)g_ComObjects[i].displayName);
        }
        SendMessage(hList, LB_SETCURSEL, 0, 0);

        // Create two radio buttons to select call mode.
        hRadioDispatch = CreateWindow(L"BUTTON", L"IDispatch",
            WS_CHILD | WS_VISIBLE | BS_RADIOBUTTON,
            10, 120, 100, 25,
            hWnd, (HMENU)IDC_RADIO_DISPATCH, g_hInst, NULL);
        hRadioDirect = CreateWindow(L"BUTTON", L"Direct",
            WS_CHILD | WS_VISIBLE | BS_RADIOBUTTON,
            120, 120, 100, 25,
            hWnd, (HMENU)IDC_RADIO_DIRECT, g_hInst, NULL);
        // Default to using IDispatch.
        SendMessage(hRadioDispatch, BM_SETCHECK, BST_CHECKED, 0);
        SendMessage(hRadioDirect, BM_SETCHECK, BST_UNCHECKED, 0);

        // Create a button to instantiate the selected COM object.
        hButtonCreate = CreateWindow(L"BUTTON", L"Create Object",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            10, 150, 150, 30,
            hWnd, (HMENU)IDC_BTN_CREATE, g_hInst, NULL);

        // Create a button to invoke a method on the object.
        hButtonInvoke = CreateWindow(L"BUTTON", L"Invoke Method",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            170, 150, 150, 30,
            hWnd, (HMENU)IDC_BTN_INVOKE, g_hInst, NULL);

        // Create an Edit control to display output messages.
        hEditOutput = CreateWindow(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
            10, 190, 310, 150,
            hWnd, (HMENU)IDC_EDIT_OUTPUT, g_hInst, NULL);

        // Create a panel that will host the ActiveX control (if needed).
        hPanel = CreateWindow(L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_WHITERECT,
            330, 10, 400, 330,
            hWnd, (HMENU)IDC_PANEL, g_hInst, NULL);
        // Create a button to list members of the selected COM object.
        hButtonListMembers = CreateWindow(L"BUTTON", L"List Members",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            10, 350, 150, 30,
            hWnd, (HMENU)IDC_BTN_LIST_MEMBERS, g_hInst, NULL);

        // Create a dropdown listbox to display members.
        hListMembers = CreateWindow(L"COMBOBOX", L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
            170, 350, 310, 150,
            hWnd, (HMENU)IDC_LIST_MEMBERS, g_hInst, NULL);

        break;
    }
    case WM_COMMAND:
    {
        int wmId = LOWORD(wParam);
        // Handle the Create Object button.
        if (wmId == IDC_BTN_CREATE)
        {
            int sel = (int)SendMessage(hList, LB_GETCURSEL, 0, 0);
            if (sel == LB_ERR)
            {
                MessageBox(hWnd, L"Please select a COM object.", L"Error", MB_OK);
                break;
            }
            ComObjectInfo selectedObj = g_ComObjects[sel];

            if (selectedObj.hasUI)
            {
                // For objects with a UI, instantiate an ActiveX control in the panel.
                CreateActiveXControl(selectedObj.progID);
                    // Add the definition of the CreateActiveXControl function to resolve the undefined identifier error.
            }
            else
            {
                // For non-UI objects, determine which call mode to use.
                bool useDispatch = (SendMessage(hRadioDispatch, BM_GETCHECK, 0, 0) == BST_CHECKED);
                CreateCOMObject(selectedObj, useDispatch);
            }
        }
        // Handle the Invoke Method button.
        else if (wmId == IDC_BTN_INVOKE)
        {
            InvokeDispatchMethod();
        }
        // Handle the List Members button.
        else if (wmId == IDC_BTN_LIST_MEMBERS)
        {
            // Clear the dropdown listbox.
            SendMessage(hListMembers, CB_RESETCONTENT, 0, 0);

            // Add placeholder items (actual member listing will be implemented later).
            SendMessage(hListMembers, CB_ADDSTRING, 0, (LPARAM)L"Method1()");
            SendMessage(hListMembers, CB_ADDSTRING, 0, (LPARAM)L"Property1");
            SendMessage(hListMembers, CB_ADDSTRING, 0, (LPARAM)L"Method2(int param)");
            SendMessage(hListMembers, CB_ADDSTRING, 0, (LPARAM)L"Property2");

            // Optionally, select the first item.
            SendMessage(hListMembers, CB_SETCURSEL, 0, 0);
        }
        break;
    }
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}


// Entry point: initializes COM in an STA, registers the window class,
// creates the main window, and starts the message loop.
int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR    lpCmdLine,
    _In_ int       nCmdShow)
{
    g_hInst = hInstance;
    // Initialize COM in the single-threaded apartment mode.
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr))
    {
        MessageBox(NULL, L"COM initialization failed", L"Error", MB_OK);
        return 0;
    }

    // Register the main window class.
    WNDCLASSEX wcex = { sizeof(WNDCLASSEX) };
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    // Add the WndProc function declaration and definition to resolve the errors.

    // Update the assignment to wcex.lpfnWndProc to use the WndProc function.
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszClassName = L"COMExplorerWindowClass";
    RegisterClassEx(&wcex);

    // Create the main application window.
    HWND hWnd = CreateWindow(L"COMExplorerWindowClass", L"COM and ActiveX Explorer",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, 0, 750, 400,
        NULL, NULL, hInstance, NULL);
    if (!hWnd)
    {
        CoUninitialize();
        return FALSE;
    }

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    // Standard message loop.
    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // Clean up.
    g_pDispatch.Release();
    CoUninitialize();
    return (int)msg.wParam;
}


