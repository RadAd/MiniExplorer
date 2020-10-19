#define _ATL_APARTMENT_THREADED

// https://docs.microsoft.com/en-us/windows/win32/lwef/nse-folderview

//#define USE_EXPLORER_BROWSER // https://www.codeproject.com/Articles/17809/Host-Windows-Explorer-in-your-applications-using-t

#include <atlbase.h>
#include <atlwin.h>
#include <atlapp.h>
#include <atlcrack.h>
#include <atltypes.h>

#include <Shlobj.h>
#include <shobjidl_core.h>

#include <tchar.h>
#include <commctrl.h>
#include <string>
#include <set>

int s_id = 0;

EXTERN_C const GUID DECLSPEC_SELECTANY IID_IShellBrowserService = { 0XDFBC7E30L, 0XF9E5, 0x455F, 0x88, 0xF8, 0xFA, 0x98, 0xC1, 0xE4, 0x94, 0xCA };

HRESULT BrowseFolder(int id, CComPtr<IShellFolder> Child, std::wstring name, HICON hIcon, FOLDERVIEWMODE ViewMode, _In_ int nShowCmd, RECT* pRect);

HRESULT BrowseFolder(int id, const ITEMIDLIST_ABSOLUTE* pidl, FOLDERVIEWMODE ViewMode, _In_ int nShowCmd, RECT* pRect)
{
    CComPtr<IShellFolder> pShellFolder;
    ATLVERIFY(SUCCEEDED(SHGetDesktopFolder(&pShellFolder)));

    CComPtr<IShellFolder> pFolder;
    ATLVERIFY(SUCCEEDED(pShellFolder->BindToObject(pidl, 0, IID_PPV_ARGS(&pFolder))));

#if 0
    CComHeapPtr<wchar_t> folder_name;
    ATLVERIFY(SUCCEEDED(SHGetNameFromIDList(pidl, SIGDN_NORMALDISPLAY, &folder_name)));

    // TODO Get hIcon

    return BrowseFolder(s_id++, pFolder, static_cast<wchar_t*>(folder_name), NULL, ViewMode, nShowCmd, pRect);
#else
    SHFILEINFO sfi = {};
    SHGetFileInfo(reinterpret_cast<LPCTSTR>(pidl), 0, &sfi, sizeof(sfi), SHGFI_PIDL | SHGFI_DISPLAYNAME | SHGFI_ICON);

    return BrowseFolder(s_id++, pFolder, sfi.szDisplayName, sfi.hIcon, ViewMode, nShowCmd, pRect);
#endif
}

class Counter
{
public:
    Counter() { ++s_counter; }
    ~Counter() { --s_counter; }
    static int get() { return s_counter;  }
private:
    static int s_counter;
};

int Counter::s_counter = 0;

template <class T>
class Registered
{
public:
    Registered(T* pThis)
        : m_pThis(pThis)
    {
        ATLVERIFY(m_pThis != nullptr);
        if (m_pThis != nullptr)
            s_registry.insert(m_pThis);
    }

    ~Registered()
    {
        ATLVERIFY(m_pThis != nullptr);
        s_registry.erase(m_pThis);
    }

private:
    T* m_pThis;
    static std::set<T*> s_registry;
};

template <class T>
std::set<T*> Registered<T>::s_registry;


class CShellBrowser :
    public CComObjectRootEx<ATL::CComSingleThreadModel>,
    public IShellBrowser,
    public IServiceProvider,
    public Registered<CShellBrowser>
{
    BEGIN_COM_MAP(CShellBrowser)
        COM_INTERFACE_ENTRY_IID(IID_IUnknown, IOleWindow)
        COM_INTERFACE_ENTRY(IOleWindow)
        COM_INTERFACE_ENTRY(IShellBrowser)
        COM_INTERFACE_ENTRY(IServiceProvider)
    END_COM_MAP()

public:
    CShellBrowser()
        : Registered<CShellBrowser>(this)
    {
    }

    void Init(HWND hWnd, CComPtr<IShellFolder> pFolder)
    {
        m_hWnd = hWnd;
        m_pFolder = pFolder;
    }

    // IOleWindow

    STDMETHOD(GetWindow)(
        /* [out] */ __RPC__deref_out_opt HWND* phwnd) override
    {
        if (!phwnd)
            return E_INVALIDARG;

        *phwnd = m_hWnd;
        return S_OK;
    }

    STDMETHOD(ContextSensitiveHelp)(
        /* [in] */ BOOL fEnterMode) override
    {
        return E_NOTIMPL;
    }

    // IShellBrowser

    STDMETHOD(InsertMenusSB)(
        _In_ HMENU hmenuShared, _Inout_ LPOLEMENUGROUPWIDTHS lpMenuWidths) override
    {
        return E_NOTIMPL;
    }

    STDMETHOD(SetMenuSB)(
        _In_opt_ HMENU hmenuShared, _In_opt_ HOLEMENU holemenuRes, _In_opt_ HWND hwndActiveObject) override
    {
        return E_NOTIMPL;
    }

    STDMETHOD(RemoveMenusSB)(
        _In_ HMENU hmenuShared) override
    {
        return E_NOTIMPL;
    }

    STDMETHOD(SetStatusTextSB)(
        _In_opt_ LPCWSTR pszStatusText) override
    {
        return E_NOTIMPL;
    }

    STDMETHOD(EnableModelessSB)(
        BOOL fEnable) override
    {
        return S_OK;
    }

    STDMETHOD(TranslateAcceleratorSB)(
        _In_ MSG* pmsg, WORD wID) override
    {
        return E_NOTIMPL;
    }

    STDMETHOD(BrowseObject)(
        _In_opt_ PCUIDLIST_RELATIVE pidl, UINT wFlags) override
    {
        if (false)
        {
            if (pidl != nullptr)
            {
                // TODO Get parent ViewMode

                if ((wFlags & 0xF000) == SBSP_ABSOLUTE)
                {
                    ATLVERIFY(SUCCEEDED(BrowseFolder(s_id++, pidl, FVM_AUTO, SW_SHOW, nullptr)));
                }
                else if ((wFlags & 0xF000) == SBSP_RELATIVE)
                {
                    CComPtr<IShellFolder> pFolder;
                    // TODO get name and hIcon
                    std::wstring name;
                    HICON hIcon = NULL;
                    ATLVERIFY(SUCCEEDED(m_pFolder->BindToObject(pidl, 0, IID_PPV_ARGS(&pFolder))));
                    ATLVERIFY(SUCCEEDED(BrowseFolder(s_id++, pFolder, name, hIcon, FVM_AUTO, SW_SHOW, nullptr)));
                }
            }

            return S_OK;
        }
        else
            return E_NOTIMPL;
    }

    STDMETHOD(GetViewStateStream)(
        DWORD grfMode, _Out_ IStream** ppStrm) override
    {
        return E_NOTIMPL;
    }

    STDMETHOD(GetControlWindow)(
        UINT id, _Out_opt_ HWND* phwnd) override
    {
        if (phwnd != nullptr)
            *phwnd = NULL;
        return S_OK;
    }

    STDMETHOD(SendControlMsg)(
        _In_  UINT id, _In_  UINT uMsg, _In_  WPARAM wParam, _In_  LPARAM lParam, _Out_opt_  LRESULT* pret) override
    {
        return E_NOTIMPL;
    }

    STDMETHOD(QueryActiveShellView)(
        _Out_ IShellView** ppshv) override
    {
        return E_NOTIMPL;
    }

    STDMETHOD(OnViewWindowActive)(
        _In_ IShellView* pshv) override
    {
        return S_OK;
    }

    STDMETHOD(SetToolbarItems)(
        _In_ LPTBBUTTONSB lpButtons, _In_  UINT nButtons, _In_  UINT uFlags) override
    {
        return E_NOTIMPL;
    }

    // IServiceProvider

    STDMETHOD(QueryService)(
        _In_ REFGUID guidService,
        _In_ REFIID riid,
        _Outptr_ void __RPC_FAR* __RPC_FAR* ppvObject) override
    {
        return QueryInterface(riid, ppvObject);
    }

private:
    HWND m_hWnd = NULL;
    CComPtr<IShellFolder> m_pFolder;
};

typedef CWinTraits<WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, WS_EX_ACCEPTFILES>		CMiniExplorerTraits;

class CMiniExplorerWnd :
    public CWindowImpl<CMiniExplorerWnd, CWindow, CMiniExplorerTraits>
{
public:
    static LPCTSTR GetWndCaption()
    {
        return _T("Mini Explorer");
    }

    CMiniExplorerWnd(int id, CComPtr<IShellFolder> pShellFolder, FOLDERVIEWMODE ViewMode)
        : m_id(id), m_pShellFolder(pShellFolder), m_ViewMode(ViewMode)
    {
    }

    BEGIN_MSG_MAP(CMiniExplorerWnd)
        MSG_WM_CREATE(OnCreate)
        MSG_WM_DESTROY(OnDestroy)
        MSG_WM_KEYDOWN(OnKeyDown)
        MSG_WM_SIZE(OnSize)
    END_MSG_MAP()

    void OnFinalMessage(_In_ HWND /*hWnd*/) override
    {
        if (Counter::get() == 1)
            PostQuitMessage(0);
        delete this;
    }

private:
    int OnCreate(LPCREATESTRUCT lpCreateStruct)
    {
        FOLDERSETTINGS fs = {};
        fs.ViewMode = m_ViewMode;
        fs.fFlags = 0;  // https://docs.microsoft.com/en-gb/windows/win32/api/shobjidl_core/ne-shobjidl_core-folderflags

        CRect rc;
        GetClientRect(rc);

#ifdef USE_EXPLORER_BROWSER
        ATLVERIFY(SUCCEEDED(SHCoCreateInstance(NULL, &CLSID_ExplorerBrowser, NULL, IID_PPV_ARGS(&m_pExplorerBrowser))));

        ATLVERIFY(SUCCEEDED(m_pExplorerBrowser->SetOptions(EBO_NAVIGATEONCE | EBO_NOBORDER)));
        ATLVERIFY(SUCCEEDED(m_pExplorerBrowser->Initialize(m_hWnd, rc, &fs)));
        ATLVERIFY(SUCCEEDED(m_pExplorerBrowser->BrowseToObject(m_pShellFolder, 0)));

        ATLVERIFY(SUCCEEDED(m_pExplorerBrowser->GetCurrentView(IID_PPV_ARGS(&m_pShellView))));
        //HWND hListView = FindWindowEx(m_hWnd, NULL, WC_LISTVIEW, nullptr);  // Doesn't work as it has a "DirectUIHWND" instead
#else

#if 1
        ATLVERIFY(SUCCEEDED(m_pShellFolder->CreateViewObject(m_hWnd, IID_PPV_ARGS(&m_pShellView))));
#else
        SFV_CREATE sfvc = {};
        sfvc.cbSize = sizeof(sfvc);
        sfvc.pshf = m_pShellFolder;
        ATLVERIFY(SUCCEEDED(SHCreateShellFolderView(&sfvc, &m_pShellView)));
#endif

        CComObject<CShellBrowser>* pCom;
        ATLVERIFY(SUCCEEDED(CComObject<CShellBrowser>::CreateInstance(&pCom)));
        pCom->Init(m_hWnd, m_pShellFolder);
        m_pShellBrowser = pCom;

        CComPtr<IShellView3> m_pShellView3;
        if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&m_pShellView3))))
        {
            ATLVERIFY(SUCCEEDED(m_pShellView3->CreateViewWindow3(m_pShellBrowser, nullptr, SV3CVW3_FORCEFOLDERFLAGS | SV3CVW3_FORCEVIEWMODE, static_cast<FOLDERFLAGS>(-1), static_cast<FOLDERFLAGS>(fs.fFlags), m_ViewMode, nullptr, rc, &m_hViewWnd)));
        }
        // else TODO Implement IShellView2
        else
        {
            ATLVERIFY(SUCCEEDED(m_pShellView->CreateViewWindow(nullptr, &fs, m_pShellBrowser, rc, &m_hViewWnd)));
        }

#if 0
        CComPtr<IFolderView> m_pFolderView;
        ATLVERIFY(SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&m_pFolderView))));

        ATLVERIFY(SUCCEEDED(m_pFolderView->SetCurrentViewMode(fs.ViewMode)));
#endif

        ATLVERIFY(SUCCEEDED(m_pShellView->UIActivate(SVUIA_ACTIVATE_NOFOCUS)));

        //ATLVERIFY(SUCCEEDED(m_pFolderView->SetCurrentViewMode(fs.ViewMode)));

        HWND hListView = FindWindowEx(m_hViewWnd, NULL, WC_LISTVIEW, nullptr);

        CComPtr <IVisualProperties> pVisualProperties;
        ATLVERIFY(SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pVisualProperties))));
        if (pVisualProperties != nullptr)
        {
            if (true)
            {
                ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_TEXT, RGB(255, 255, 255))));
                ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_BACKGROUND, RGB(0, 0, 0))));
                ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_SORTCOLUMN, RGB(25, 25, 25))));
                ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_SUBTEXT, RGB(200, 200, 2000))));
                ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_TEXTBACKGROUND, RGB(255, 0, 0))));    // TODO Not sure what this is?
            }
            else
            {
                pVisualProperties->SetTheme(L"DarkMode_Explorer", nullptr);
            }
        }
        else if (hListView != NULL)
        {
            ListView_SetBkColor(hListView, RGB(0, 0, 0));
            ListView_SetTextColor(hListView, RGB(255, 255, 255));
            ListView_SetTextBkColor(hListView, RGB(0, 0, 0));
            ListView_SetOutlineColor(hListView, RGB(255, 0, 0));    // TODO Not sure what this is?

            //ListView_SetView(hListView, LV_VIEW_SMALLICON);
        }
#endif

        CComPtr<IDropTarget> pDropTarget;
        ATLVERIFY(SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pDropTarget))));
        ATLVERIFY(SUCCEEDED(RegisterDragDrop(m_hWnd, pDropTarget)));

        return 0;
    }

    void OnDestroy()
    {
        ATLVERIFY(SUCCEEDED(RevokeDragDrop(m_hWnd)));

        TCHAR keyname[1024];
        _stprintf_s(keyname, _T("Software\\RadSoft\\MiniExplorer\\Windows\\%d"), m_id);

        CRegKey reg;
        ATLVERIFY(ERROR_SUCCESS == reg.Create(HKEY_CURRENT_USER, keyname));

        CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
        ATLVERIFY(SUCCEEDED(SHGetIDListFromObject(m_pShellFolder, &spidl)));

        ATLVERIFY(ERROR_SUCCESS == reg.SetBinaryValue(_T("pidl"), reinterpret_cast<BYTE*>(static_cast<PIDLIST_ABSOLUTE>(spidl)), ILGetSize(spidl)));

        CRect rc;
        GetWindowRect(rc);
        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("top"), rc.top));
        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("left"), rc.left));
        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("bottom"), rc.bottom));
        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("right"), rc.right));

        // TODO May be better to use m_pShellView->SaveViewState()

        if (true)
        {
            FOLDERSETTINGS fs = {};
            ATLVERIFY(SUCCEEDED(m_pShellView->GetCurrentInfo(&fs)));

            ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("view"), fs.ViewMode));
        }
        else
        {
            CComPtr<IFolderView> m_pFolderView;
            ATLVERIFY(SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&m_pFolderView))));

            UINT viewmode = 0;
            ATLVERIFY(SUCCEEDED(m_pFolderView->GetCurrentViewMode(&viewmode)));

            ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("view"), viewmode));
        }

#ifdef USE_EXPLORER_BROWSER
        ATLVERIFY(SUCCEEDED(m_pExplorerBrowser->Destroy()));
#else
        ATLVERIFY(SUCCEEDED(m_pShellView->UIActivate(SVUIA_DEACTIVATE)));
        ATLVERIFY(SUCCEEDED(m_pShellView->DestroyViewWindow()));
#endif
    }

    void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
    {
        // TODO How to implement "Delete", "F2", "Ctrl+C", "Ctrl+V", etc
        SetMsgHandled(FALSE);
    }

    void OnSize(UINT nType, CSize size)
    {
        CRect rc(0, 0, size.cx, size.cy);
#ifdef USE_EXPLORER_BROWSER
        m_pExplorerBrowser->SetRect(NULL, rc);
#else
        if (m_hViewWnd != NULL)
        {
            HWND hListView = FindWindowEx(m_hViewWnd, NULL, WC_LISTVIEW, nullptr);
            if (hListView != NULL)
            {
                // ListView isn't correctly removing the header when not in details mode
                // TODO Need a way to detect when mode changes
                if (true)
                {
                    LONG style = ::GetWindowLong(hListView, GWL_STYLE);
                    if (ListView_GetView(hListView) != LV_VIEW_DETAILS)
                        style |= LVS_NOCOLUMNHEADER;
                    else
                        style &= ~LVS_NOCOLUMNHEADER;
                    ::SetWindowLong(hListView, GWL_STYLE, style);
                }
                else if (ListView_GetView(hListView) != LV_VIEW_DETAILS)
                {
                    HWND hHeader = ListView_GetHeader(hListView);
                    if (hHeader != NULL)
                    {
                        CRect wrc;
                        ::GetWindowRect(hHeader, wrc);
                        rc.top -= wrc.Height();
                    }
                }
            }
            ::SetWindowPos(m_hViewWnd, NULL, rc.left, rc.top, rc.Width(), rc.Height(), SWP_NOZORDER);
        }
#endif
    }

    int m_id;
    FOLDERVIEWMODE m_ViewMode;
    CComPtr<IShellFolder> m_pShellFolder;
    CComPtr<IShellView> m_pShellView;

#ifdef USE_EXPLORER_BROWSER
    CComPtr<IExplorerBrowser> m_pExplorerBrowser;
#else
    CComPtr<IShellBrowser> m_pShellBrowser;
    HWND m_hViewWnd = NULL;
#endif
    Counter m_counter;
};

HRESULT BrowseFolder(int id, CComPtr<IShellFolder> pShellFolder, std::wstring name, HICON hIcon, FOLDERVIEWMODE ViewMode, _In_ int nShowCmd, RECT* pRect)
{
    CMiniExplorerWnd* mainwnd = new CMiniExplorerWnd(id, pShellFolder, ViewMode);
    if (!mainwnd->Create(NULL, pRect))
        return AtlHresultFromWin32(GetLastError());
    mainwnd->SetWindowText((name + L" - " + CMiniExplorerWnd::GetWndCaption()).c_str());
    mainwnd->SetIcon(hIcon);
    mainwnd->ShowWindow(nShowCmd);
    return S_OK;
}

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
        //AtlInitCommonControls(0xFFFF);

        CComPtr<IShellFolder> pShellFolder;
        ATLVERIFY(SUCCEEDED(SHGetDesktopFolder(&pShellFolder)));

        CRegKey reg;
        reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\Windows"));

        {
            TCHAR name[1024];
            DWORD szname = ARRAYSIZE(name);
            DWORD index = 0;
            while (ERROR_SUCCESS == reg.EnumKey(index++, name, &szname))
            {
                int id = std::stoi(name);

                CRegKey childreg;
                childreg.Open(reg, name);

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

                CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
                ULONG bytes = 0;
                childreg.QueryBinaryValue(_T("pidl"), nullptr, &bytes);
                spidl.AllocateBytes(bytes);
                childreg.QueryBinaryValue(_T("pidl"), spidl, &bytes);

                ATLVERIFY(SUCCEEDED(BrowseFolder(id, spidl, ViewMode, nShowCmd, &rc)));

                if (s_id <= id)
                    s_id = id + 1;

                szname = ARRAYSIZE(name);
            }
            _putts(name);
        }

        if (false)
        {
            BROWSEINFO bi = {};
            bi.lpszTitle = _T("Select a folder to open");
            bi.ulFlags = BIF_EDITBOX | BIF_NEWDIALOGSTYLE | BIF_BROWSEFILEJUNCTIONS;
            CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl(SHBrowseForFolder(&bi));
            if (spidl)
                ATLVERIFY(SUCCEEDED(BrowseFolder(s_id++, spidl, FVM_AUTO, nShowCmd, nullptr)));
        }

        HRESULT hr = __super::PreMessageLoop(nShowCmd);
        if (FAILED(hr))
            return hr;

        return S_OK;
    }
};

CModule _AtlModule;

int APIENTRY _tWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPTSTR lpCmdLine, _In_ int nShowCmd)
{
    return _AtlModule.WinMain(nShowCmd);
}
