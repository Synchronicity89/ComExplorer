#include <vector>
#include <string>
#include <algorithm>

// Structure to hold COM object information.
struct ComObjectInfo {
    std::wstring displayName; // Friendly name as shown in the list box.
    std::wstring progID;      // ProgID (or CLSID string) used to create the object.
    bool hasUI;               // Whether the object has a visual interface.
};

// Function to enumerate COM objects and populate a vector.
std::vector<ComObjectInfo> EnumerateCOMObjects()
{
    std::vector<ComObjectInfo> comObjects;

    // Open the CLSID registry key.
    HKEY hKey;
    if (RegOpenKeyEx(HKEY_CLASSES_ROOT, L"CLSID", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
    {
        MessageBox(NULL, L"Failed to open CLSID registry key.", L"Error", MB_OK);
        return comObjects;
    }

    // Enumerate all subkeys under CLSID.
    wchar_t clsidKeyName[256];
    DWORD index = 0;
    DWORD keyNameSize = _countof(clsidKeyName);
    while (RegEnumKeyEx(hKey, index, clsidKeyName, &keyNameSize, NULL, NULL, NULL, NULL) == ERROR_SUCCESS)
    {
        // Open the CLSID subkey.
        HKEY hSubKey;
        if (RegOpenKeyEx(hKey, clsidKeyName, 0, KEY_READ, &hSubKey) == ERROR_SUCCESS)
        {
            // Query the friendly name (default value).
            wchar_t friendlyName[256];
            DWORD friendlyNameSize = sizeof(friendlyName);
            if (RegQueryValueEx(hSubKey, NULL, NULL, NULL, (LPBYTE)friendlyName, &friendlyNameSize) == ERROR_SUCCESS)
            {
                // Check if the object has a "Control" subkey (indicating it's an ActiveX control).
                HKEY hControlKey;
                bool hasUI = (RegOpenKeyEx(hSubKey, L"Control", 0, KEY_READ, &hControlKey) == ERROR_SUCCESS);
                if (hasUI)
                {
                    RegCloseKey(hControlKey);
                }

                // Query the ProgID (if available).
                wchar_t progID[256];
                DWORD progIDSize = sizeof(progID);
                if (RegQueryValueEx(hSubKey, L"ProgID", NULL, NULL, (LPBYTE)progID, &progIDSize) == ERROR_SUCCESS)
                {
                    // Add the COM object to the list.
                    comObjects.push_back({ friendlyName, progID, hasUI });
                }
                else
                {
                    // If no ProgID, use the CLSID as the identifier.
                    comObjects.push_back({ friendlyName, clsidKeyName, hasUI });
                }
            }
            RegCloseKey(hSubKey);
        }

        // Prepare for the next iteration.
        keyNameSize = _countof(clsidKeyName);
        index++;
    }

    RegCloseKey(hKey);

    // Sort the list alphabetically by display name.
    std::sort(comObjects.begin(), comObjects.end(), [](const ComObjectInfo& a, const ComObjectInfo& b) {
        return a.displayName < b.displayName;
    });

    return comObjects;
}
