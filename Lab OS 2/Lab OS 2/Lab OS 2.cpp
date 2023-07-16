#include <iostream>
#include <windows.h>
using namespace std;

int main()
{
    setlocale(LC_ALL, "Russian");
    HKEY hKey;
    LONG lResult = RegOpenKeyEx(HKEY_CURRENT_USER, L"SOFTWARE\\Microsoft\\Office\\16.0\\Excel", NULL, KEY_QUERY_VALUE, &hKey);
    if (lResult == ERROR_SUCCESS)
    {
        WCHAR buffer[MAX_PATH];
        DWORD bufferSize = MAX_PATH;
        LONG result = RegQueryValueEx(hKey,L"ExcelPreviousSessionVersion", NULL, NULL, (LPBYTE)&buffer, &bufferSize);
        if (result == ERROR_SUCCESS)
        {
            //std::cout << "The reg kes \'ExcelPreviousSessionVersion\' setup to value = " << buffer << "\n";
            wstring ws(buffer);
            string str(ws.begin(), ws.end());
            cout << "Установленная версия MS Excel: " << str << endl;
        }
        else
        {
            std::cout << "Error: " << result << "\n";
        }
    }
    else if (lResult != ERROR_SUCCESS)
    {
        if (lResult == ERROR_FILE_NOT_FOUND) {
            printf("Key not found.\n");
            return TRUE;
        }
        else {
            printf("Error opening key.\n");
            return FALSE;
        }
    }
    return 0;
}
