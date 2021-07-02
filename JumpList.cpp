#include <atlbase.h>
#include <atlcom.h>

#include <propvarutil.h>
#include <propkey.h>

#include <Shlobj.h>

HRESULT _CreateShellLink(PCWSTR pszArguments, PCWSTR pszTitle, PCWSTR pszIconLocation, int iIcon, IShellLink** ppsl)
{
    CComPtr<IShellLink> psl;
    ATLENSURE_RETURN(SUCCEEDED(psl.CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER)));

    // Determine our executable's file path so the task will execute this application
    WCHAR szAppPath[MAX_PATH];
    ATLENSURE_RETURN_HR(GetModuleFileName(NULL, szAppPath, ARRAYSIZE(szAppPath)), HRESULT_FROM_WIN32(GetLastError()));

    ATLVERIFY(SUCCEEDED(psl->SetPath(szAppPath)));
    ATLVERIFY(SUCCEEDED(psl->SetArguments(pszArguments)));
    if (pszIconLocation != nullptr)
        ATLVERIFY(SUCCEEDED(psl->SetIconLocation(pszIconLocation, iIcon)));

    // The title property is required on Jump List items provided as an IShellLink
    // instance.  This value is used as the display name in the Jump List.
    CComPtr<IPropertyStore> pps;
    ATLVERIFY(SUCCEEDED(psl.QueryInterface(&pps)));

    PROPVARIANT propvar;
    ATLVERIFY(SUCCEEDED(InitPropVariantFromString(pszTitle, &propvar)));
    ATLVERIFY(SUCCEEDED(pps->SetValue(PKEY_Title, propvar)));
    ATLVERIFY(SUCCEEDED(pps->Commit()));
    ATLVERIFY(SUCCEEDED(psl.QueryInterface(ppsl)));
    ATLVERIFY(SUCCEEDED(PropVariantClear(&propvar)));

    return S_OK;
}

HRESULT _CreateSeparatorLink(IShellLink** ppsl)
{
    CComPtr<IPropertyStore> pps;
    ATLENSURE_RETURN(SUCCEEDED(pps.CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER)));

    PROPVARIANT propvar;
    ATLVERIFY(SUCCEEDED(InitPropVariantFromBoolean(TRUE, &propvar)));
    ATLVERIFY(SUCCEEDED(pps->SetValue(PKEY_AppUserModel_IsDestListSeparator, propvar)));
    ATLVERIFY(SUCCEEDED(pps->Commit()));
    ATLVERIFY(SUCCEEDED(pps.QueryInterface(ppsl)));
    ATLVERIFY(SUCCEEDED(PropVariantClear(&propvar)));

    return S_OK;
}

HRESULT _AddShellLink(IObjectCollection* poc, PCWSTR pszArguments, PCWSTR pszTitle, PCWSTR pszIconLocation, int iIcon)
{
    CComPtr<IShellLink> psl;
    ATLENSURE_RETURN(SUCCEEDED(_CreateShellLink(pszArguments, pszTitle, pszIconLocation, iIcon, &psl)));
    ATLENSURE_RETURN(SUCCEEDED(poc->AddObject(psl)));
    return S_OK;
}

HRESULT _AddSeparatorLink(IObjectCollection* poc)
{
    CComPtr<IShellLink> psl;
    ATLENSURE_RETURN(SUCCEEDED(_CreateSeparatorLink(&psl)));
    ATLENSURE_RETURN(SUCCEEDED(poc->AddObject(psl)));
    return S_OK;
}

HRESULT _AddTasksToList(ICustomDestinationList* pcdl)
{
    CComPtr<IObjectCollection> poc;
    ATLENSURE_RETURN(SUCCEEDED(poc.CoCreateInstance(CLSID_EnumerableObjectCollection, NULL, CLSCTX_INPROC)));

    ATLVERIFY(SUCCEEDED(_AddShellLink(poc, L"/Add", L"Add Window", nullptr, -1)));
    ATLVERIFY(SUCCEEDED(_AddShellLink(poc, L"/Remove", L"Remove Window", nullptr, -1)));
    ATLVERIFY(SUCCEEDED(_AddSeparatorLink(poc)));
    {
        CRegKey reg;
        reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\Windows"));

        {
            TCHAR name[1024];
            DWORD szname = 0;
            DWORD index = 0;
            while (szname = ARRAYSIZE(name), ERROR_SUCCESS == reg.EnumKey(index++, name, &szname))
            {
                CRegKey childreg;
                childreg.Open(reg, name);

                TCHAR cmd[1024];
                _stprintf_s(cmd, _T("/Open %s"), name);

                CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
                ULONG bytes = 0;
                childreg.QueryBinaryValue(_T("pidl"), nullptr, &bytes);
                spidl.AllocateBytes(bytes);
                childreg.QueryBinaryValue(_T("pidl"), spidl, &bytes);

                CComHeapPtr<wchar_t> folder_name;
                ATLVERIFY(SUCCEEDED(SHGetNameFromIDList(spidl, SIGDN_NORMALDISPLAY, &folder_name)));

                SHFILEINFO sfi = {};
                SHGetFileInfoW(reinterpret_cast<LPCWSTR>(static_cast<ITEMIDLIST_ABSOLUTE*>(spidl)), 0, &sfi, sizeof(sfi), SHGFI_PIDL | SHGFI_ICONLOCATION);

                ATLVERIFY(SUCCEEDED(_AddShellLink(poc, cmd, folder_name, sfi.szDisplayName, sfi.iIcon)));
            }
        }
    }

    CComPtr<IObjectArray> poa;
    ATLVERIFY(SUCCEEDED(poc.QueryInterface(&poa)));

    // Add the tasks to the Jump List. Tasks always appear in the canonical "Tasks"
    // category that is displayed at the bottom of the Jump List, after all other
    // categories.
    ATLVERIFY(SUCCEEDED(pcdl->AddUserTasks(poa)));

    return S_OK;
}

void CreateJumpList()
{
    CComPtr<ICustomDestinationList> pcdl;
    ATLVERIFY(SUCCEEDED(pcdl.CoCreateInstance(CLSID_DestinationList, nullptr, CLSCTX_INPROC_SERVER)));

    UINT cMinSlots;
    CComPtr<IObjectArray> poaRemoved;
    ATLVERIFY(SUCCEEDED(pcdl->BeginList(&cMinSlots, IID_PPV_ARGS(&poaRemoved))));
    ATLVERIFY(SUCCEEDED(_AddTasksToList(pcdl)));
    ATLVERIFY(SUCCEEDED(pcdl->CommitList()));
}
