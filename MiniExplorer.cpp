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

#include <commctrl.h>

HRESULT BrowseFolder(CComPtr<IShellFolder> Child, _In_ int nShowCmd);

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
        if (pidl != nullptr)
        {
            if ((wFlags & 0xF000) == SBSP_ABSOLUTE)
            {
                CComPtr<IShellFolder> pShellFolder;
                ATLVERIFY(SUCCEEDED(SHGetDesktopFolder(&pShellFolder)));
                CComPtr<IShellFolder> Child;
                ATLVERIFY(SUCCEEDED(pShellFolder->BindToObject(pidl, 0, IID_IShellFolder, (LPVOID*) &Child)));
                BrowseFolder(Child, SW_SHOW);
            }
            else if ((wFlags & 0xF000) == SBSP_RELATIVE)
            {
                CComPtr<IShellFolder> Child;
                ATLVERIFY(SUCCEEDED(m_pFolder->BindToObject(pidl, 0, IID_IShellFolder, (LPVOID*) &Child)));
                BrowseFolder(Child, SW_SHOW);
            }
        }

        return S_OK;
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
        if (IID_IShellBrowser == riid)
        {
            *ppvObject = static_cast<IShellBrowser*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = NULL;
        return E_NOINTERFACE;
    }

private:
    HWND m_hWnd = NULL;
    CComPtr<IShellFolder> m_pFolder;
};

// SHGetIDListFromObject

class CMainWnd :
    public CWindowImpl<CMainWnd, CWindow, CFrameWinTraits>
{
public:
    static LPCTSTR GetWndCaption()
    {
        return _T("Mini Explorer");
    }

    CMainWnd(CComPtr<IShellFolder> pShellFolder)
        : m_pShellFolder(pShellFolder)
    {
    }

    BEGIN_MSG_MAP(CMainWnd)
        MSG_WM_CREATE(OnCreate)
        MSG_WM_DESTROY(OnDestroy)
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
        fs.ViewMode = FVM_AUTO;
        fs.fFlags = 0;  // https://docs.microsoft.com/en-gb/windows/win32/api/shobjidl_core/ne-shobjidl_core-folderflags

        CRect rc;
        GetClientRect(rc);

#ifdef USE_EXPLORER_BROWSER
        ATLVERIFY(SUCCEEDED(SHCoCreateInstance(NULL, &CLSID_ExplorerBrowser, NULL, IID_PPV_ARGS(&m_pExplorerBrowser))));

        ATLVERIFY(SUCCEEDED(m_pExplorerBrowser->SetOptions(EBO_NAVIGATEONCE | EBO_NOBORDER)));
        ATLVERIFY(SUCCEEDED(m_pExplorerBrowser->Initialize(m_hWnd, rc, &fs)));
        ATLVERIFY(SUCCEEDED(m_pExplorerBrowser->BrowseToObject(m_pShellFolder, 0)));

        HWND hListView = FindWindowEx(m_hWnd, NULL, WC_LISTVIEW, nullptr);  // Doesn't work as it has a "DirectUIHWND" instead
#else
        ATLVERIFY(SUCCEEDED(m_pShellFolder->CreateViewObject(m_hWnd, IID_IShellView, (void **) &m_pShellView)));

        CComObject<CShellBrowser>* pCom;
        ATLVERIFY(SUCCEEDED(CComObject<CShellBrowser>::CreateInstance(&pCom)));
        pCom->Init(m_hWnd, m_pShellFolder);
        m_pShellBrowser = pCom;

        ATLVERIFY(SUCCEEDED(m_pShellView->CreateViewWindow(nullptr, &fs, m_pShellBrowser, rc, &m_hViewWnd)));
        ATLVERIFY(SUCCEEDED(m_pShellView->UIActivate(SVUIA_ACTIVATE_NOFOCUS)));

        HWND hListView = FindWindowEx(m_hViewWnd, NULL, WC_LISTVIEW, nullptr);
#endif
        if (hListView != NULL)
        {
            ListView_SetBkColor(hListView, RGB(0, 0, 0));
            ListView_SetTextColor(hListView, RGB(255, 255, 255));
            ListView_SetTextBkColor(hListView, RGB(0, 0, 0));
        }

        return 0;
    }

    void OnDestroy()
    {
#ifdef USE_EXPLORER_BROWSER
        m_pExplorerBrowser->Destroy();
#else
        ATLVERIFY(SUCCEEDED(m_pShellView->DestroyViewWindow()));
#endif
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
            if (hListView != NULL  && ListView_GetView(hListView) != LV_VIEW_DETAILS)
            {
                HWND hHeader = ListView_GetHeader(hListView);
                if (hHeader != NULL)
                {
                    CRect wrc;
                    ::GetWindowRect(hHeader, wrc);
                    rc.top -= wrc.Height();
                }
            }
            ::SetWindowPos(m_hViewWnd, NULL, rc.left, rc.top, rc.Width(), rc.Height(), SWP_NOZORDER);
        }
#endif
    }

    CComPtr<IShellFolder> m_pShellFolder;

#ifdef USE_EXPLORER_BROWSER
    CComPtr<IExplorerBrowser> m_pExplorerBrowser;
#else
    CComPtr<IShellView> m_pShellView;
    CComPtr<IShellBrowser> m_pShellBrowser;
    HWND m_hViewWnd = NULL;
#endif
    Counter m_counter;
};

HRESULT BrowseFolder(CComPtr<IShellFolder> pShellFolder, _In_ int nShowCmd)
{
    CMainWnd* mainwnd = new CMainWnd(pShellFolder);
    if (!mainwnd->Create(NULL))
        return AtlHresultFromWin32(GetLastError());
    mainwnd->ShowWindow(nShowCmd);
    return S_OK;
}

class CModule : public ATL::CAtlExeModuleT< CModule >
{
public:
    HRESULT PreMessageLoop(_In_ int nShowCmd) throw()
    {
        //AtlInitCommonControls(0xFFFF);

        CComPtr<IShellFolder> pShellFolder;
        ATLVERIFY(SUCCEEDED(SHGetDesktopFolder(&pShellFolder)));
        BrowseFolder(pShellFolder, nShowCmd);

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
