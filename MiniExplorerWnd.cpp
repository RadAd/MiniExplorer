#include "MiniExplorerWnd.h"

#include <string>

HRESULT BrowseFolder(int id, CComPtr<IShellFolder> Child, std::wstring name, HICON hIcon, FOLDERFLAGS flags, FOLDERVIEWMODE ViewMode, _In_ int nShowCmd, RECT* pRect);
void OpenMRU(PCUIDLIST_ABSOLUTE pidl, FOLDERFLAGS flags, FOLDERVIEWMODE ViewMode);

template <class T>
std::set<T*> Registered<T>::s_registry;

class CShellBrowser :
    public CComObjectRootEx<ATL::CComSingleThreadModel>,
    public IShellBrowser,
    public IServiceProvider
{
    BEGIN_COM_MAP(CShellBrowser)
        COM_INTERFACE_ENTRY_IID(IID_IUnknown, IOleWindow)
        COM_INTERFACE_ENTRY(IOleWindow)
        COM_INTERFACE_ENTRY(IShellBrowser)
        COM_INTERFACE_ENTRY(IServiceProvider)
    END_COM_MAP()

public:
    void Init(HWND hWnd, int id, CComPtr<IShellFolder> pFolder)
    {
        m_hWnd = hWnd;
        m_id = id;
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
        if (!(GetKeyState(VK_SHIFT) & 0x8000))
        {
            if (pidl != nullptr)
            {
                FOLDERSETTINGS fs = {};
                fs.ViewMode = FVM_AUTO;
                fs.fFlags = FWF_NONE;
                if (m_pShellView != nullptr)
                    ATLVERIFY(SUCCEEDED(m_pShellView->GetCurrentInfo(&fs)));

                if ((wFlags & 0xF000) == SBSP_ABSOLUTE)
                {
                    OpenMRU(pidl, (FOLDERFLAGS) fs.fFlags, (FOLDERVIEWMODE) fs.ViewMode);
                }
                else if ((wFlags & 0xF000) == SBSP_RELATIVE)
                {
                    CComPtr<IShellFolder> pFolder;
                    // TODO get name and hIcon
                    std::wstring name;
                    HICON hIcon = NULL;
                    ATLVERIFY(SUCCEEDED(m_pFolder->BindToObject(pidl, 0, IID_PPV_ARGS(&pFolder))));
                    // TODO Find the appropriate id
                    ATLVERIFY(SUCCEEDED(BrowseFolder(-100, pFolder, name, hIcon, (FOLDERFLAGS) fs.fFlags, (FOLDERVIEWMODE) fs.ViewMode, SW_SHOW, nullptr)));
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
#if 1
        if (ppStrm != nullptr)
        {
            TCHAR keyname[1024];
            if (m_id >= 0)
                _stprintf_s(keyname, _T("Software\\RadSoft\\MiniExplorer\\Windows\\%d"), m_id);
            else
                _stprintf_s(keyname, _T("Software\\RadSoft\\MiniExplorer\\MRU\\%d"), -m_id);
            *ppStrm = SHOpenRegStream2(HKEY_CURRENT_USER, keyname, _T("state"), grfMode);
            //ATLVERIFY(*ppStrm != nullptr);
            return *ppStrm != nullptr ? S_OK : E_FAIL;
        }
        else
#endif
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
        if (ppshv != nullptr)
            *ppshv = m_pShellView;
        return S_OK;
    }

    STDMETHOD(OnViewWindowActive)(
        _In_ IShellView* pshv) override
    {
        m_pShellView = pshv;
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
    int m_id = -1;
    CComPtr<IShellFolder> m_pFolder;
    IShellView* m_pShellView = nullptr;
};

int CMiniExplorerWnd::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
    FOLDERSETTINGS fs = {};
    fs.ViewMode = m_ViewMode;
    fs.fFlags = m_flags;  // https://docs.microsoft.com/en-gb/windows/win32/api/shobjidl_core/ne-shobjidl_core-folderflags

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
    pCom->Init(m_hWnd, m_id, m_pShellFolder);
    m_pShellBrowser = pCom;

    CComPtr<IShellView3> m_pShellView3;
    if (SUCCEEDED(m_pShellView.QueryInterface(&m_pShellView3)))
    {
        ATLVERIFY(SUCCEEDED(m_pShellView3->CreateViewWindow3(m_pShellBrowser, nullptr, SV3CVW3_FORCEFOLDERFLAGS | SV3CVW3_FORCEVIEWMODE, static_cast<FOLDERFLAGS>(-1), m_flags, m_ViewMode, nullptr, rc, &m_hViewWnd)));
    }
    // else TODO Implement IShellView2
    else
    {
        ATLVERIFY(SUCCEEDED(m_pShellView->CreateViewWindow(nullptr, &fs, m_pShellBrowser, rc, &m_hViewWnd)));
    }

#if 0
    CComPtr<IFolderView> m_pFolderView;
    ATLVERIFY(SUCCEEDED(m_pShellView.QueryInterface(&m_pFolderView)));

    ATLVERIFY(SUCCEEDED(m_pFolderView->SetCurrentViewMode(fs.ViewMode)));
#endif

    ATLVERIFY(SUCCEEDED(m_pShellView->UIActivate(SVUIA_ACTIVATE_NOFOCUS)));

    //ATLVERIFY(SUCCEEDED(m_pFolderView->SetCurrentViewMode(fs.ViewMode)));

    HWND hListView = FindWindowEx(m_hViewWnd, NULL, WC_LISTVIEW, nullptr);

    CComPtr <IVisualProperties> pVisualProperties;
    ATLVERIFY(SUCCEEDED(m_pShellView.QueryInterface(&pVisualProperties)));
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
    ATLVERIFY(SUCCEEDED(m_pShellView.QueryInterface(&pDropTarget)));
    ATLVERIFY(SUCCEEDED(RegisterDragDrop(m_hWnd, pDropTarget)));

    return 0;
}

void CMiniExplorerWnd::OnDestroy()
{
    ATLVERIFY(SUCCEEDED(RevokeDragDrop(m_hWnd)));

    if (true)
    {
        TCHAR keyname[1024];
        if (m_id >= 0)
            _stprintf_s(keyname, _T("Software\\RadSoft\\MiniExplorer\\Windows\\%d"), m_id);
        else
            _stprintf_s(keyname, _T("Software\\RadSoft\\MiniExplorer\\MRU\\%d"), -m_id);

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
        m_pShellView->SaveViewState();

        if (true)
        {
            FOLDERSETTINGS fs = {};
            ATLVERIFY(SUCCEEDED(m_pShellView->GetCurrentInfo(&fs)));

            ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("view"), fs.ViewMode));
            ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("flags"), fs.fFlags));
        }
        else if (false)
        {
            CComPtr<IFolderView> m_pFolderView;
            ATLVERIFY(SUCCEEDED(m_pShellView.QueryInterface(&m_pFolderView)));

            UINT viewmode = 0;
            ATLVERIFY(SUCCEEDED(m_pFolderView->GetCurrentViewMode(&viewmode)));

            ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("view"), viewmode));

            // TODO Use IFolderView2 to get more settings
        }
    }

#ifdef USE_EXPLORER_BROWSER
    ATLVERIFY(SUCCEEDED(m_pExplorerBrowser->Destroy()));
#else
    ATLVERIFY(SUCCEEDED(m_pShellView->UIActivate(SVUIA_DEACTIVATE)));
    ATLVERIFY(SUCCEEDED(m_pShellView->DestroyViewWindow()));
#endif
}

void CMiniExplorerWnd::OnSize(UINT nType, CSize size)
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

void CMiniExplorerWnd::OnSetFocus(CWindow wndOld)
{
    m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
    //::SetFocus(m_hViewWnd);
}

