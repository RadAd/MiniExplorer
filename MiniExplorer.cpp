#define _ATL_APARTMENT_THREADED

// https://docs.microsoft.com/en-us/windows/win32/lwef/nse-folderview
// https://github.com/pauldotknopf/WindowsSDK7-Samples/blob/master/winui/shell/appshellintegration/CustomJumpList/CustomJumpListSample.cpp

//#define USE_EXPLORER_BROWSER // https://www.codeproject.com/Articles/17809/Host-Windows-Explorer-in-your-applications-using-t

// TODO
// Store and use colours from registry
// Switching between details and not doesn't show headings
// Where to store settings for adhoc windows???

#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <atlbase.h>

#include <tchar.h>
#include <string>
#include <set>
#include <vector>

#include "MiniExplorerWnd.h"

#define CD_COMMAND_LINE 524

HWND s_MainWnd = NULL;

//EXTERN_C const GUID DECLSPEC_SELECTANY IID_IShellBrowserService = { 0XDFBC7E30L, 0XF9E5, 0x455F, 0x88, 0xF8, 0xFA, 0x98, 0xC1, 0xE4, 0x94, 0xCA };

void CreateJumpList();

CMiniExplorerWnd* GetMiniExplorer(int id)
{
    for (CMiniExplorerWnd* wnd : Registered<CMiniExplorerWnd>::get())
    {
        if (wnd->GetId() == id)
            return wnd;
    }
    return nullptr;
}

int Find(CRegKey reg, const ITEMIDLIST_ABSOLUTE* pidl, bool& isnew)
{
    std::set<int> used;
    isnew = false;

    TCHAR name[1024];
    DWORD szname = 0;
    DWORD index = 0;
    while (szname = ARRAYSIZE(name), ERROR_SUCCESS == reg.EnumKey(index++, name, &szname))
    {
        CRegKey childreg;
        childreg.Open(reg, name);

        CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
        ULONG bytes = 0;
        childreg.QueryBinaryValue(_T("pidl"), nullptr, &bytes);
        if (ILGetSize(pidl) == bytes)
        {
            spidl.AllocateBytes(bytes);
            childreg.QueryBinaryValue(_T("pidl"), spidl, &bytes);

            const int id = std::stoi(name);
            if (ILIsEqual(pidl, spidl))
                return std::stoi(name);
            else
                used.insert(id);
        }
    }
    isnew = true;
    int newid = 0;
    while (used.find(newid) != used.end())
        ++newid;
    return newid;
}

HRESULT BrowseFolder(int id, CComPtr<IShellFolder> pShellFolder, std::wstring name, HICON hIcon, FOLDERFLAGS flags, FOLDERVIEWMODE ViewMode, _In_ int nShowCmd, RECT* pRect)
{
    // TODO Put flags, ViewMode in create params
    CMiniExplorerWnd* mainwnd = new CMiniExplorerWnd(id, pShellFolder, flags, ViewMode);
    if (!mainwnd->Create(s_MainWnd, pRect))
        return AtlHresultFromWin32(GetLastError());
    mainwnd->SetWindowText((name + L" - " + CMiniExplorerWnd::GetWndCaption()).c_str());
    mainwnd->SetIcon(hIcon);
    mainwnd->ShowWindow(nShowCmd);
    return S_OK;
}

HRESULT BrowseFolder(int id, const ITEMIDLIST_ABSOLUTE* pidl, FOLDERFLAGS flags, FOLDERVIEWMODE ViewMode, _In_ int nShowCmd, RECT* pRect)
{
    CComPtr<IShellFolder> pShellFolder;
    ATLVERIFY(SUCCEEDED(SHGetDesktopFolder(&pShellFolder)));

    CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
    ATLVERIFY(SUCCEEDED(SHGetIDListFromObject(pShellFolder, &spidl)));

    CComPtr<IShellFolder> pFolder;
    if (ILIsEqual(pidl, spidl))
        pFolder = pShellFolder;
    else
        ATLVERIFY(SUCCEEDED(pShellFolder->BindToObject(pidl, 0, IID_PPV_ARGS(&pFolder))));

#if 0
    CComHeapPtr<wchar_t> folder_name;
    ATLVERIFY(SUCCEEDED(SHGetNameFromIDList(pidl, SIGDN_NORMALDISPLAY, &folder_name)));

    // TODO Get hIcon

    return BrowseFolder(id, pFolder, static_cast<wchar_t*>(folder_name), NULL, ViewMode, nShowCmd, pRect);
#else
    SHFILEINFO sfi = {};
    SHGetFileInfo(reinterpret_cast<LPCTSTR>(pidl), 0, &sfi, sizeof(sfi), SHGFI_PIDL | SHGFI_DISPLAYNAME | SHGFI_ICON);

    return BrowseFolder(id, pFolder, sfi.szDisplayName, sfi.hIcon, flags, ViewMode, nShowCmd, pRect);
#endif
}

bool OpenMiniExplorer(HKEY hKeyParent, LPCTSTR lpszKeyName, const int id, const int nShowCmd)
{
    ATLVERIFY(GetMiniExplorer(id) == nullptr);

    CRegKey childreg;
    if (childreg.Open(hKeyParent, lpszKeyName) != ERROR_SUCCESS)
    {
        MessageBox(NULL, _T("No entry."), _T("Mini Explorer"), MB_ICONERROR | MB_OK);
        return false;
    }

    DWORD temp;

    CRect rc = CMiniExplorerWnd::rcDefault;
    ATLVERIFY(ERROR_SUCCESS == childreg.QueryDWORDValue(_T("top"), temp));
    rc.top = temp;
    ATLVERIFY(ERROR_SUCCESS == childreg.QueryDWORDValue(_T("left"), temp));
    rc.left = temp;
    ATLVERIFY(ERROR_SUCCESS == childreg.QueryDWORDValue(_T("bottom"), temp));
    rc.bottom = temp;
    ATLVERIFY(ERROR_SUCCESS == childreg.QueryDWORDValue(_T("right"), temp));
    rc.right = temp;

    FOLDERVIEWMODE ViewMode = FVM_AUTO;
    ATLVERIFY(ERROR_SUCCESS == childreg.QueryDWORDValue(_T("view"), temp));
    ViewMode = static_cast<FOLDERVIEWMODE>(temp);
    FOLDERFLAGS flags = FWF_NONE;
    ATLVERIFY(ERROR_SUCCESS == childreg.QueryDWORDValue(_T("flags"), temp));
    flags = static_cast<FOLDERFLAGS>(temp);

    CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
    ULONG bytes = 0;
    childreg.QueryBinaryValue(_T("pidl"), nullptr, &bytes);
    spidl.AllocateBytes(bytes);
    childreg.QueryBinaryValue(_T("pidl"), spidl, &bytes);

    ATLVERIFY(SUCCEEDED(BrowseFolder(id, spidl, flags, ViewMode, nShowCmd, &rc)));

    return true;
}

bool AddMiniExplorer(_In_ int nShowCmd)
{
    BROWSEINFO bi = {};
    bi.lpszTitle = _T("Select a folder to open");
    bi.ulFlags = BIF_EDITBOX | BIF_NEWDIALOGSTYLE | BIF_BROWSEFILEJUNCTIONS;
    CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl(SHBrowseForFolder(&bi));
    if (spidl)
    {
        std::set<int> used;

        CRegKey reg;
        reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\Windows"));

        bool isnew = false;
        const int id = Find(reg, spidl, isnew);

        CMiniExplorerWnd* wnd = GetMiniExplorer(id);
        if (wnd != nullptr)
            SetForegroundWindow(*wnd);
        else if (isnew)
            ATLVERIFY(SUCCEEDED(BrowseFolder(id, spidl, FWF_NONE, FVM_AUTO, nShowCmd, nullptr)));
        else
        {
            std::wstring name = std::to_wstring(id);
            OpenMiniExplorer(reg, name.c_str(), id, SW_SHOW);
        }

        return true;
    }
    else
        return false;
}

void OpenMRU(const ITEMIDLIST_ABSOLUTE* pidl, FOLDERFLAGS flags, FOLDERVIEWMODE ViewMode)
{
    int max = 0;
    CRegKey reg;
    reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\MRU"));

    // TODO Window may be open without saving to registry yet
    bool isnew = false;
    const int id = Find(reg, pidl, isnew);

    std::vector<unsigned char> mru;
    ULONG bytes = 0;
    reg.QueryBinaryValue(_T("mru"), nullptr, &bytes);
    mru.resize(bytes / sizeof(unsigned char));
    reg.QueryBinaryValue(_T("mru"), mru.data(), &bytes);

    mru.erase(std::remove(mru.begin(), mru.end(), id), mru.end());
    mru.insert(mru.begin(), id);

    // TODO if mru too big then reuse an id

    reg.SetBinaryValue(_T("mru"), mru.data(), (ULONG) mru.size() * sizeof(unsigned char));

    CMiniExplorerWnd* wnd = GetMiniExplorer(-id);
    if (wnd != nullptr)
        SetForegroundWindow(*wnd);
    else if (isnew)
        ATLVERIFY(SUCCEEDED(BrowseFolder(-id, pidl, flags, ViewMode, SW_SHOW, nullptr)));
    else
    {
        std::wstring name = std::to_wstring(id);
        OpenMiniExplorer(reg, name.c_str(), -id, SW_SHOW);
    }
}

bool ParseCommandLine(_In_ PCWSTR lpCmdLine, _In_ int nShowCmd)
{
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(lpCmdLine, &argc);
    if (argc > 1 && _wcsicmp(argv[1], _T("/Add")) == 0)
    {
        if (AddMiniExplorer(SW_SHOW))
        {
            CreateJumpList();
            return true;
        }
        else
            return false;
    }
    else if (argc > 1 && _wcsicmp(argv[1], _T("/Remove")) == 0)
    {
        // TODO Select an entry
        // Close window if open
        // Delete registry entry
        //CreateJumpList();
        return true;
    }
    else if (argc > 2 && _wcsicmp(argv[1], _T("/Open")) == 0)
    {
        const int id = std::stoi(argv[2]);

        CMiniExplorerWnd* wnd = GetMiniExplorer(id);

        if (wnd == nullptr)
        {
            CRegKey reg;
            reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\Windows"));

            if (!OpenMiniExplorer(reg, argv[2], id, nShowCmd))
                return false;
        }
        else
        {
            SetForegroundWindow(*wnd);
        }

        return true;
    }
    else if (argc == 1)
    {
        CRegKey reg;
        reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\Windows"));

        TCHAR name[1024];
        DWORD szname = 0;
        DWORD index = 0;
        while (szname = ARRAYSIZE(name), ERROR_SUCCESS == reg.EnumKey(index++, name, &szname))
        {
            const int id = std::stoi(name);

            CMiniExplorerWnd* wnd = GetMiniExplorer(id);

            if (wnd == nullptr)
                OpenMiniExplorer(reg, name, id, nShowCmd);
        }

        return true;
    }
    else
        return false;
}

class CMainWnd :
    public CWindowImpl<CMainWnd, CWindow, CFrameWinTraits>
{
public:
    DECLARE_WND_CLASS(_T("MINI_EXPLORER"))

    static LPCTSTR GetWndCaption()
    {
        return _T("Mini Explorer");
    }

    BEGIN_MSG_MAP(CMainWnd)
        MSG_WM_COPYDATA(OnCopyData)
    END_MSG_MAP()

private:
    BOOL OnCopyData(CWindow wnd, PCOPYDATASTRUCT pCopyDataStruct)
    {
        if (pCopyDataStruct->dwData == CD_COMMAND_LINE)
        {
            SetForegroundWindow(m_hWnd);
            ::ParseCommandLine((LPTSTR) pCopyDataStruct->lpData, SW_SHOW);
        }
        return TRUE;
    }
};

class CModule : public ATL::CAtlExeModuleT< CModule >
{
public:
    static HRESULT InitializeCom() throw()
    {
        return OleInitialize(NULL);
    }

    static void UninitializeCom() throw()
    {
        OleUninitialize();
    }

    HRESULT PreMessageLoop(_In_ int nShowCmd) throw()
    {
        HWND hWnd = FindWindow(CMainWnd::GetWndClassInfo().m_wc.lpszClassName, nullptr);
        if (hWnd != NULL)
        {
            COPYDATASTRUCT cds;
            cds.dwData = CD_COMMAND_LINE;
            cds.lpData = GetCommandLine();
            cds.cbData = (DWORD) (_tcslen(GetCommandLine()) + 1) * sizeof(TCHAR);
            ::SendMessage(hWnd, WM_COPYDATA, NULL, (LPARAM) &cds);
            return S_FALSE;
        }

        //AtlInitCommonControls(0xFFFF);

        CMainWnd* main_wnd = new CMainWnd();
        s_MainWnd = main_wnd->Create(NULL);

        if (!::ParseCommandLine(GetCommandLine(), nShowCmd))
            return -1;

        CreateJumpList();

        HRESULT hr = __super::PreMessageLoop(nShowCmd);
        if (FAILED(hr))
            return hr;

        return S_OK;
    }

    void RunMessageLoop() throw()
    {
        MSG msg;
        while (GetMessage(&msg, 0, 0, 0) > 0)
        {
            CMiniExplorerWnd* pMiniExplorerWnd = nullptr;
            for (CMiniExplorerWnd* wnd : Registered<CMiniExplorerWnd>::get())
            {
                if (*wnd == msg.hwnd || wnd->IsChild(msg.hwnd))
                {
                    pMiniExplorerWnd = wnd;
                    break;
                }
            }

            if (pMiniExplorerWnd == nullptr || pMiniExplorerWnd->TranslateAccelerator(&msg) != S_OK)
            {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
    }
};

CModule _AtlModule;

int APIENTRY _tWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPTSTR lpCmdLine, _In_ int nShowCmd)
{
    int r = _AtlModule.WinMain(nShowCmd);
    return r;
}
