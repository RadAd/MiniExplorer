#include "MiniExplorerWnd.h"

#include <shdispid.h>
#include <propvarutil.h>

#include <string>

#include "Trace.h"

std::wstring StringFromCLSID(REFCLSID rclsid)
{
    CComHeapPtr<OLECHAR> lpsz;
    ATLVERIFY(SUCCEEDED(::StringFromCLSID(rclsid, &lpsz)));
    return static_cast<const OLECHAR*>(lpsz);
}

inline void SetWindowBlur(HWND hWnd, bool bEnabled)
{
    const HINSTANCE hModule = GetModuleHandle(TEXT("user32.dll"));
    if (hModule)
    {
        typedef enum _ACCENT_STATE {
            ACCENT_DISABLED,
            ACCENT_ENABLE_GRADIENT,
            ACCENT_ENABLE_TRANSPARENTGRADIENT,
            ACCENT_ENABLE_BLURBEHIND,
            ACCENT_ENABLE_ACRYLICBLURBEHIND,
            ACCENT_INVALID_STATE
        } ACCENT_STATE;
        struct ACCENTPOLICY
        {
            ACCENT_STATE nAccentState;
            DWORD nFlags;
            DWORD nColor;
            DWORD nAnimationId;
        };
        typedef enum _WINDOWCOMPOSITIONATTRIB {
            WCA_UNDEFINED = 0,
            WCA_NCRENDERING_ENABLED = 1,
            WCA_NCRENDERING_POLICY = 2,
            WCA_TRANSITIONS_FORCEDISABLED = 3,
            WCA_ALLOW_NCPAINT = 4,
            WCA_CAPTION_BUTTON_BOUNDS = 5,
            WCA_NONCLIENT_RTL_LAYOUT = 6,
            WCA_FORCE_ICONIC_REPRESENTATION = 7,
            WCA_EXTENDED_FRAME_BOUNDS = 8,
            WCA_HAS_ICONIC_BITMAP = 9,
            WCA_THEME_ATTRIBUTES = 10,
            WCA_NCRENDERING_EXILED = 11,
            WCA_NCADORNMENTINFO = 12,
            WCA_EXCLUDED_FROM_LIVEPREVIEW = 13,
            WCA_VIDEO_OVERLAY_ACTIVE = 14,
            WCA_FORCE_ACTIVEWINDOW_APPEARANCE = 15,
            WCA_DISALLOW_PEEK = 16,
            WCA_CLOAK = 17,
            WCA_CLOAKED = 18,
            WCA_ACCENT_POLICY = 19,
            WCA_FREEZE_REPRESENTATION = 20,
            WCA_EVER_UNCLOAKED = 21,
            WCA_VISUAL_OWNER = 22,
            WCA_LAST = 23
        } WINDOWCOMPOSITIONATTRIB;
        struct WINCOMPATTRDATA
        {
            WINDOWCOMPOSITIONATTRIB nAttribute;
            PVOID pData;
            ULONG ulDataSize;
        };
        typedef BOOL(WINAPI* pSetWindowCompositionAttribute)(HWND, WINCOMPATTRDATA*);
        const pSetWindowCompositionAttribute SetWindowCompositionAttribute = (pSetWindowCompositionAttribute)GetProcAddress(hModule, "SetWindowCompositionAttribute");
        if (SetWindowCompositionAttribute)
        {
            ACCENTPOLICY policy = { bEnabled ? ACCENT_ENABLE_ACRYLICBLURBEHIND : ACCENT_DISABLED, 0, 0x80000000, 0 };
            //ACCENTPOLICY policy = { ACCENT_ENABLE_BLURBEHIND };
            WINCOMPATTRDATA data = { WCA_ACCENT_POLICY, &policy, sizeof(ACCENTPOLICY) };
            SetWindowCompositionAttribute(hWnd, &data);
            //DwmSetWindowAttribute(hWnd, WCA_ACCENT_POLICY, &policy, sizeof(ACCENTPOLICY));
        }
        //FreeLibrary(hModule);
    }
}

#define SM_FAVOURITE 1236
#define SM_OPEN_IN_EXPLORER 1234
#define SM_OPEN_UP 1235

int Find(CRegKey& reg, PCUIDLIST_ABSOLUTE pidl, bool& isnew);
HRESULT BrowseFolder(int id, CComPtr<IShellFolder> Child, std::wstring name, HICON hIcon, const MiniExplorerSettings& settings, _In_ int nShowCmd, RECT* pRect);
bool OpenFavourite(PCUIDLIST_ABSOLUTE pidl);
void OpenMRU(PCUIDLIST_ABSOLUTE pidl, const MiniExplorerSettings& settings);
void CreateJumpList();

MiniExplorerSettings GetSettings(IShellView* pShellView)
{
    MiniExplorerSettings settings = {};
    settings.ViewMode = FVM_AUTO;
    settings.flags = FWF_NONE;

    if (pShellView != nullptr)
    {
        //CComQIPtr<IFolderView> pFolderView(pShellView);
        CComQIPtr<IFolderView2> pFolderView2(pShellView);

        if (pFolderView2)
        {
            ATLVERIFY(SUCCEEDED(pFolderView2->GetViewModeAndIconSize(&settings.ViewMode, &settings.iIconSize)));

            DWORD FolderFlags;
            ATLVERIFY(SUCCEEDED(pFolderView2->GetCurrentFolderFlags(&FolderFlags)));
            settings.flags = static_cast<FOLDERFLAGS>(FolderFlags);
        }
#if 0
        else if (pFolderView)
        {
            UINT ViewMode = 0;
            ATLVERIFY(SUCCEEDED(pFolderView->GetCurrentViewMode(&ViewMode)));
            settings.ViewMode = static_cast<FOLDERVIEWMODE>(ViewMode);

            // TODO Missing flags
        }
#endif
        else
        {
            FOLDERSETTINGS fs = {};
            fs.ViewMode = FVM_AUTO;
            fs.fFlags = FWF_NONE;
            ATLVERIFY(SUCCEEDED(pShellView->GetCurrentInfo(&fs)));
            settings.ViewMode = static_cast<FOLDERVIEWMODE>(fs.ViewMode);
            settings.flags = static_cast<FOLDERFLAGS>(fs.fFlags);
        }
    }

    return settings;
}

void FixListViewHeader(HWND hViewWnd)
{
    CWindow hListView = FindWindowEx(hViewWnd, NULL, WC_LISTVIEW, nullptr);
    if (hListView)
    {
        // ListView isn't correctly removing the header when not in details mode
        LONG style = hListView.GetWindowLong(GWL_STYLE);
        if (ListView_GetView(hListView) != LV_VIEW_DETAILS)
            style |= LVS_NOCOLUMNHEADER;
        else
            style &= ~LVS_NOCOLUMNHEADER;
        hListView.SetWindowLong(GWL_STYLE, style);
    }
}

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
        TRACE(__FUNCTIONW__ _T(" %p\n"), phwnd);
        if (!phwnd)
            return E_INVALIDARG;

        *phwnd = m_hWnd;
        return S_OK;
    }

    STDMETHOD(ContextSensitiveHelp)(
        /* [in] */ BOOL fEnterMode) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        return E_NOTIMPL;
    }

    // IShellBrowser

    STDMETHOD(InsertMenusSB)(
        _In_ HMENU hmenuShared, _Inout_ LPOLEMENUGROUPWIDTHS lpMenuWidths) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        return E_NOTIMPL;
    }

    STDMETHOD(SetMenuSB)(
        _In_opt_ HMENU hmenuShared, _In_opt_ HOLEMENU holemenuRes, _In_opt_ HWND hwndActiveObject) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        return E_NOTIMPL;
    }

    STDMETHOD(RemoveMenusSB)(
        _In_ HMENU hmenuShared) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        return E_NOTIMPL;
    }

    STDMETHOD(SetStatusTextSB)(
        _In_opt_ LPCWSTR pszStatusText) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        return E_NOTIMPL;
    }

    STDMETHOD(EnableModelessSB)(
        BOOL fEnable) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        return S_OK;
    }

    STDMETHOD(TranslateAcceleratorSB)(
        _In_ MSG* pmsg, WORD wID) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        return E_NOTIMPL;
    }

    STDMETHOD(BrowseObject)(
        _In_opt_ PCUIDLIST_RELATIVE pidl, UINT wFlags) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        if (!(GetKeyState(VK_SHIFT) & 0x8000))
        {
            if (pidl != nullptr)
            {
                const MiniExplorerSettings settings = GetSettings(m_pShellView);

                if (wFlags & SBSP_RELATIVE)
                {
                    CComPtr<IShellFolder> pFolder;
                    ATLVERIFY(SUCCEEDED(m_pFolder->BindToObject(pidl, 0, IID_PPV_ARGS(&pFolder))));

                    CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
                    ATLVERIFY(SUCCEEDED(SHGetIDListFromObject(pFolder, &spidl)));

                    if (!OpenFavourite(spidl))
                        OpenMRU(spidl, settings);
                }
                else /* SBSP_ABSOLUTE */
                {
                    if (!OpenFavourite(pidl))
                        OpenMRU(pidl, settings);
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
        TRACE(__FUNCTIONW__ _T(" \n"));
#if 1
        if (ppStrm != nullptr)
        {
            const std::wstring keyname =  m_id >= 0
                ? Format(_T("Software\\RadSoft\\MiniExplorer\\Windows\\%d"), m_id)
                : Format(_T("Software\\RadSoft\\MiniExplorer\\MRU\\%d"), -m_id);
            *ppStrm = SHOpenRegStream2(HKEY_CURRENT_USER, keyname.c_str(), _T("state"), grfMode);
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
        TRACE(__FUNCTIONW__ _T(" \n"));
        if (phwnd != nullptr)
            *phwnd = NULL;
        return S_OK;
    }

    STDMETHOD(SendControlMsg)(
        _In_  UINT id, _In_  UINT uMsg, _In_  WPARAM wParam, _In_  LPARAM lParam, _Out_opt_  LRESULT* pret) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        return E_NOTIMPL;
    }

    STDMETHOD(QueryActiveShellView)(
        _Out_ IShellView** ppshv) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        if (ppshv != nullptr)
            *ppshv = m_pShellView;
        return S_OK;
    }

    STDMETHOD(OnViewWindowActive)(
        _In_ IShellView* pshv) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        m_pShellView = pshv;
        return S_OK;
    }

    STDMETHOD(SetToolbarItems)(
        _In_ LPTBBUTTONSB lpButtons, _In_  UINT nButtons, _In_  UINT uFlags) override
    {
        TRACE(__FUNCTIONW__ _T(" \n"));
        return E_NOTIMPL;
    }

    // IServiceProvider

    STDMETHOD(QueryService)(
        _In_ REFGUID guidService,
        _In_ REFIID riid,
        _Outptr_ void __RPC_FAR* __RPC_FAR* ppvObject) override
    {
        TRACE(__FUNCTIONW__ _T(" %s %s\n"), StringFromCLSID(guidService).c_str(), StringFromCLSID(riid).c_str());
        return QueryInterface(riid, ppvObject);
    }

private:
    HWND m_hWnd = NULL;
    int m_id = -1;
    CComPtr<IShellFolder> m_pFolder;
    IShellView* m_pShellView = nullptr;
};

class CShellFolderViewEvents :
    public CComObjectRootEx<ATL::CComSingleThreadModel>,
    public IDispatchImpl<DShellFolderViewEvents>
{
    BEGIN_COM_MAP(CShellFolderViewEvents)
        COM_INTERFACE_ENTRY(DShellFolderViewEvents)
        COM_INTERFACE_ENTRY(IDispatch)
        COM_INTERFACE_ENTRY(IUnknown)
    END_COM_MAP()

public:
    void Init(HWND hViewWnd)
    {
        m_hViewWnd = hViewWnd;
    }

    // DShellFolderViewEvents

    STDMETHOD(Invoke)(
        _In_  DISPID dispIdMember,
        _In_  REFIID riid,
        _In_  LCID lcid,
        _In_  WORD wFlags,
        _In_  DISPPARAMS* pDispParams,
        _Out_opt_  VARIANT* pVarResult,
        _Out_opt_  EXCEPINFO* pExcepInfo,
        _Out_opt_  UINT* puArgErr) override
    {
        TRACE(__FUNCTIONW__ _T(" %d \n"), dispIdMember);
        switch (dispIdMember)
        {
        case DISPID_VIEWMODECHANGED:
            FixListViewHeader(m_hViewWnd);
            break;
        }
        return S_OK;
        //return IDispatchImpl<DShellFolderViewEvents>::Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
    }

private:
    HWND m_hViewWnd = NULL;
};

void CMiniExplorerWnd::SetId(int id)
{
    m_id = id;
    static_cast<CShellBrowser*>(static_cast<IShellBrowser*>(m_pShellBrowser))->Init(m_hWnd, m_id, m_pShellFolder);
}


int CMiniExplorerWnd::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
    TRACE(__FUNCTIONW__ _T(" \n"));
    const MiniExplorerSettings* pSettings = static_cast<MiniExplorerSettings*>(lpCreateStruct->lpCreateParams);

    SetWindowBlur(*this, true);

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

    CComQIPtr<IShellView3> pShellView3(m_pShellView);
    if (pShellView3)
    {
        HWND hViewWnd;
        ATLVERIFY(SUCCEEDED(pShellView3->CreateViewWindow3(m_pShellBrowser, nullptr, SV3CVW3_FORCEFOLDERFLAGS | SV3CVW3_FORCEVIEWMODE, static_cast<FOLDERFLAGS>(-1), pSettings->flags, pSettings->ViewMode, nullptr, rc, &hViewWnd)));
        m_hViewWnd = hViewWnd;
    }
    // else TODO Implement IShellView2
    else
    {
        FOLDERSETTINGS fs = {};
        fs.ViewMode = pSettings->ViewMode;
        fs.fFlags = pSettings->flags;  // https://docs.microsoft.com/en-gb/windows/win32/api/shobjidl_core/ne-shobjidl_core-folderflags
        HWND hViewWnd;
        ATLVERIFY(SUCCEEDED(m_pShellView->CreateViewWindow(nullptr, &fs, m_pShellBrowser, rc, &hViewWnd)));
        m_hViewWnd = hViewWnd;
    }
    FixListViewHeader(m_hViewWnd);

    //CComQIPtr<IFolderView> pFolderView(m_pShellView);
    CComQIPtr<IFolderView2> pFolderView2(m_pShellView);

    if (pFolderView2)
    {
        //ATLVERIFY(SUCCEEDED(pFolderView2->SetCurrentFolderFlags(static_cast<FOLDERFLAGS>(-1), pSettings->flags)));
        if (pSettings->iIconSize > 0)
            ATLVERIFY(SUCCEEDED(pFolderView2->SetViewModeAndIconSize(pSettings->ViewMode, pSettings->iIconSize)));
        // TODO Set sort columns
    }
#if 0
    else if (pFolderView)
    {
        //ATLVERIFY(SUCCEEDED(pFolderView->SetCurrentViewMode(pSettings->ViewMode)));
    }
#endif

    ATLVERIFY(SUCCEEDED(m_pShellView->UIActivate(SVUIA_ACTIVATE_NOFOCUS)));

    HWND hListView = FindWindowEx(m_hViewWnd, NULL, WC_LISTVIEW, nullptr);

    CComQIPtr<IVisualProperties> pVisualProperties(m_pShellView);
    if (pVisualProperties)
    {
        if (true)
        {
            ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_TEXT, RGB(250, 250, 250))));
            ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_BACKGROUND, RGB(0, 0, 0))));
            ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_SORTCOLUMN, RGB(25, 25, 25))));
            ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_SUBTEXT, RGB(100, 100, 100))));
            ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_TEXTBACKGROUND, RGB(255, 0, 255))));    // TODO Not sure what this is?
        }
        else
        {
            pVisualProperties->SetTheme(L"DarkMode_Explorer", nullptr);
        }
    }
    else if (hListView != NULL)
    {
        ListView_SetBkColor(hListView, RGB(0, 0, 0));
        ListView_SetTextColor(hListView, RGB(250, 250, 250));
        ListView_SetTextBkColor(hListView, RGB(61, 61, 61));
        ListView_SetOutlineColor(hListView, RGB(255, 0, 255));    // TODO Not sure what this is?

        //ListView_SetView(hListView, LV_VIEW_SMALLICON);
    }
#endif

    CComQIPtr<IDropTarget> pDropTarget(m_pShellView);
    ATLVERIFY(pDropTarget);
    ATLVERIFY(SUCCEEDED(RegisterDragDrop(m_hWnd, pDropTarget)));

    CComPtr<IDispatch> viewDisp;
    ATLVERIFY(SUCCEEDED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&viewDisp))));

    CComObject<CShellFolderViewEvents>* pEvents;
    ATLVERIFY(SUCCEEDED(CComObject<CShellFolderViewEvents>::CreateInstance(&pEvents)));
    pEvents->Init(m_hViewWnd);
    m_pEvents = pEvents;

    ATLVERIFY(SUCCEEDED(AtlAdvise(viewDisp, m_pEvents, __uuidof(DShellFolderViewEvents), &m_dwCookie)));

    CMenuHandle hMenu = GetSystemMenu(FALSE);
    if (hMenu)
    {
        // TODO Shell_MergeMenus(hMenu, hMenuMerge, )
        hMenu.AppendMenu(MF_SEPARATOR);
        hMenu.AppendMenu(MF_STRING, SM_FAVOURITE, _T("Favourite"));
        hMenu.AppendMenu(MF_STRING, SM_OPEN_IN_EXPLORER, _T("Open in Explorer..."));
        hMenu.AppendMenu(MF_STRING, SM_OPEN_UP, _T("Up..."));

        hMenu.CheckMenuItem(SM_FAVOURITE, MF_BYCOMMAND | (GetId() >= 0 ? MF_CHECKED : MF_UNCHECKED));
    }

    return 0;
}

void CMiniExplorerWnd::OnDestroy()
{
    TRACE(__FUNCTIONW__ _T(" \n"));
    ATLVERIFY(SUCCEEDED(RevokeDragDrop(m_hWnd)));

    CComPtr<IDispatch> viewDisp;
    ATLVERIFY(SUCCEEDED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&viewDisp))));

    ATLVERIFY(SUCCEEDED(AtlUnadvise(viewDisp, __uuidof(DShellFolderViewEvents), m_dwCookie)));

    if (true)
    {
        const std::wstring keyname = m_id >= 0
            ? Format(_T("Software\\RadSoft\\MiniExplorer\\Windows\\%d"), m_id)
            : Format(_T("Software\\RadSoft\\MiniExplorer\\MRU\\%d"), -m_id);

        CRegKey reg;
        ATLVERIFY(ERROR_SUCCESS == reg.Create(HKEY_CURRENT_USER, keyname.c_str()));

        CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
        ATLVERIFY(SUCCEEDED(SHGetIDListFromObject(m_pShellFolder, &spidl)));

        ATLVERIFY(ERROR_SUCCESS == reg.SetBinaryValue(_T("pidl"), reinterpret_cast<BYTE*>(static_cast<PIDLIST_ABSOLUTE>(spidl)), ILGetSize(spidl)));

        CRect rc;
        GetWindowRect(rc);
        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("top"), rc.top));
        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("left"), rc.left));
        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("bottom"), rc.bottom));
        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("right"), rc.right));

        m_pShellView->SaveViewState();

        const MiniExplorerSettings settings = GetSettings(m_pShellView);

        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("view"), settings.ViewMode));
        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("flags"), settings.flags));
        ATLVERIFY(ERROR_SUCCESS == reg.SetDWORDValue(_T("iconsize"), settings.iIconSize));
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
    TRACE(__FUNCTIONW__ _T(" \n"));
    CRect rc(0, 0, size.cx, size.cy);
#ifdef USE_EXPLORER_BROWSER
    m_pExplorerBrowser->SetRect(NULL, rc);
#else
    if (m_hViewWnd)
        m_hViewWnd.SetWindowPos(NULL, rc.left, rc.top, rc.Width(), rc.Height(), SWP_NOZORDER);
#endif
}

void CMiniExplorerWnd::OnEnterSizeMove()
{
    SetWindowBlur(*this, false);
}

void CMiniExplorerWnd::OnExitSizeMove()
{
    SetWindowBlur(*this, true);
}

void CMiniExplorerWnd::OnSetFocus(CWindow wndOld)
{
    TRACE(__FUNCTIONW__ _T(" \n"));
    m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
    //m_hViewWnd.SetFocus();
}

void CMiniExplorerWnd::OnDpiChanged(UINT nDpiX, UINT nDpiY, PRECT pRect)
{
    TRACE(__FUNCTIONW__ _T(" \n"));
    SetWindowPos(NULL, pRect, SWP_NOZORDER | SWP_NOACTIVATE);
}

void CMiniExplorerWnd::OnSysCommand(UINT nID, CPoint point)
{
    TRACE(__FUNCTIONW__ _T(" \n"));
    switch (nID)
    {
    case SM_FAVOURITE:
        if (GetId() >= 0)
        {
            // TODO Remove from favourites
            {
                CRegKey reg;
                reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\Windows"));
                reg.DeleteSubKey(Format(_T("%d"), m_id).c_str());
            }

            CRegKey reg;
            reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\MRU"));
            
            CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
            GetPidl(&spidl);

            bool isnew = false;
            const int id = Find(reg, spidl, isnew);

            SetId(-id);

            if (isnew)
            {
                CRegKey regwnd;
                ATLVERIFY(ERROR_SUCCESS == regwnd.Create(reg, Format(_T("%d"), id).c_str()));
                ATLVERIFY(ERROR_SUCCESS == regwnd.SetBinaryValue(_T("pidl"), reinterpret_cast<BYTE*>(static_cast<PIDLIST_ABSOLUTE>(spidl)), ILGetSize(spidl)));
            }
        }
        else
        {
            CRegKey reg;
            reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\Windows"));

            CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
            GetPidl(&spidl);

            bool isnew = false;
            const int id = Find(reg, spidl, isnew);
            ATLASSERT(isnew);

            SetId(id);

            CRegKey regwnd;
            ATLVERIFY(ERROR_SUCCESS == regwnd.Create(reg, Format(_T("%d"), m_id).c_str()));
            ATLVERIFY(ERROR_SUCCESS == regwnd.SetBinaryValue(_T("pidl"), reinterpret_cast<BYTE*>(static_cast<PIDLIST_ABSOLUTE>(spidl)), ILGetSize(spidl)));
        }

        {
            CreateJumpList();
            CMenuHandle hMenu = GetSystemMenu(FALSE);
            if (hMenu)
                hMenu.CheckMenuItem(SM_FAVOURITE, MF_BYCOMMAND | (GetId() >= 0 ? MF_CHECKED : MF_UNCHECKED));
        }
        break;

    case SM_OPEN_IN_EXPLORER:
        {
            CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
            ATLVERIFY(SUCCEEDED(SHGetIDListFromObject(m_pShellFolder, &spidl)));

#if 0
            CComHeapPtr<WCHAR> name;
            ATLVERIFY(SUCCEEDED(SHGetNameFromIDList(spidl, SIGDN_FILESYSPATH, &name)));

            // %SystemRoot%\Explorer.exe /e,::{MyExtension CLSID},objectname

            STARTUPINFO startup_info = {};
            startup_info.cb = sizeof(STARTUPINFO);
            PROCESS_INFORMATION process_info = {};

            WCHAR cmd[1024];
            //wsprintf(cmd, L"%SystemRoot%\Explorer.exe /e,::{MyExtension CLSID},objectname", );
            wsprintf(cmd, L"Explorer.exe \"%s\"", static_cast<const OLECHAR*>(name));

            BOOL rv = ::CreateProcess(
                NULL,                // LPCTSTR lpApplicationName
                cmd,                 // LPTSTR lpCommandLine
                NULL,                // LPSECURITY_ATTRIBUTES lpProcessAttributes
                NULL,                // LPSECURITY_ATTRIBUTES lpThreadAttributes
                FALSE,               // BOOL bInheritHandles
                0,                   // DWORD dwCreationFlags
                NULL,                // LPVOID lpEnvironment
                NULL,                // LPCTSTR lpCurrentDirectory
                &startup_info,       // LPSTARTUPINFO lpStartupInfo
                &process_info        // LPPROCESS_INFORMATION lpProcessInformation
            );

            if (rv)
            {
                ::CloseHandle(process_info.hThread);
                ::CloseHandle(process_info.hProcess);
            }
            else
            {
            }
#else
            CComPtr<IWebBrowser2> pwb;
            ATLVERIFY(SUCCEEDED(pwb.CoCreateInstance(CLSID_ShellBrowserWindow, nullptr, CLSCTX_LOCAL_SERVER)));

            ATLVERIFY(SUCCEEDED(CoAllowSetForegroundWindow(pwb, 0)));

            CComVariant varTarget;
            ATLVERIFY(SUCCEEDED(InitVariantFromBuffer(spidl, ILGetSize(spidl), &varTarget)));

            CComVariant vEmpty;
            ATLVERIFY(SUCCEEDED(pwb->Navigate2(&varTarget, &vEmpty, &vEmpty, &vEmpty, &vEmpty)));
            ATLVERIFY(SUCCEEDED(pwb->put_Visible(VARIANT_TRUE)));
#endif
        }
        break;

    case SM_OPEN_UP:
        {
            CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
            ATLVERIFY(SUCCEEDED(SHGetIDListFromObject(m_pShellFolder, &spidl)));

            if (ILRemoveLastID(spidl))
            {
                const MiniExplorerSettings settings = GetSettings(m_pShellView);
                if (!OpenFavourite(spidl))
                    OpenMRU(spidl, settings);
            }
        }
        break;

    default:
        SetMsgHandled(FALSE);
        break;
    }
}
