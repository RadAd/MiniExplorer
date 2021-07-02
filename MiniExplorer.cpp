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
#include <atlwin.h>
#include <atlapp.h>
#include <atlcrack.h>
#include <atltypes.h>

#include <Shlobj.h>
#include <shobjidl_core.h>

#include <propvarutil.h>
#include <propkey.h>

#include <strsafe.h>

#include <tchar.h>
#include <commctrl.h>
#include <string>
#include <set>
#include <vector>

int s_id = 0;

#define CD_COMMAND_LINE 524

EXTERN_C const GUID DECLSPEC_SELECTANY IID_IShellBrowserService = { 0XDFBC7E30L, 0XF9E5, 0x455F, 0x88, 0xF8, 0xFA, 0x98, 0xC1, 0xE4, 0x94, 0xCA };

HRESULT BrowseFolder(int id, CComPtr<IShellFolder> Child, std::wstring name, HICON hIcon, FOLDERFLAGS flags, FOLDERVIEWMODE ViewMode, _In_ int nShowCmd, RECT* pRect);
void OpenMRU(const ITEMIDLIST_ABSOLUTE* pidl, FOLDERFLAGS flags, FOLDERVIEWMODE ViewMode);
void CreateJumpList();

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

void AddMiniExplorer(_In_ int nShowCmd)
{
    BROWSEINFO bi = {};
    bi.lpszTitle = _T("Select a folder to open");
    bi.ulFlags = BIF_EDITBOX | BIF_NEWDIALOGSTYLE | BIF_BROWSEFILEJUNCTIONS;
    CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl(SHBrowseForFolder(&bi));
    if (spidl)
        ATLVERIFY(SUCCEEDED(BrowseFolder(s_id++, spidl, FWF_NONE, FVM_AUTO, nShowCmd, nullptr)));
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

    static const std::set<T*>& get() { return s_registry; }

private:
    T* m_pThis;
    static std::set<T*> s_registry;
};

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

typedef CWinTraits<WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, WS_EX_ACCEPTFILES | WS_EX_APPWINDOW /*| WS_EX_TOOLWINDOW*/>		CMiniExplorerTraits;

class CMiniExplorerWnd : public CWindowImpl<CMiniExplorerWnd, CWindow, CMiniExplorerTraits>,
    public Registered<CMiniExplorerWnd>
{
public:
    static LPCTSTR GetWndCaption()
    {
        return _T("Mini Explorer");
    }

    CMiniExplorerWnd(int id, CComPtr<IShellFolder> pShellFolder, FOLDERFLAGS flags, FOLDERVIEWMODE ViewMode)
        : Registered<CMiniExplorerWnd>(this)
        , m_id(id), m_pShellFolder(pShellFolder), m_flags(flags), m_ViewMode(ViewMode)
    {
    }

    BEGIN_MSG_MAP(CMiniExplorerWnd)
        MSG_WM_CREATE(OnCreate)
        MSG_WM_DESTROY(OnDestroy)
        MSG_WM_SIZE(OnSize)
        MSG_WM_SETFOCUS(OnSetFocus)
    END_MSG_MAP()

    int GetId() const
    {
        return m_id;
    }

    void OnFinalMessage(_In_ HWND /*hWnd*/) override
    {
        //if (Counter::get() == 1)
            //PostQuitMessage(0);
        delete this;
    }

    HRESULT TranslateAccelerator(MSG* pmsg)
    {
        return m_pShellView->TranslateAccelerator(pmsg);
    }

private:
    int OnCreate(LPCREATESTRUCT lpCreateStruct)
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

    void OnDestroy()
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

    void OnSetFocus(CWindow wndOld)
    {
        m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
        //::SetFocus(m_hViewWnd);
    }

    int m_id;
    FOLDERFLAGS m_flags;
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

HWND s_MainWnd = NULL;

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

CMiniExplorerWnd* GetMiniExplorer(int id)
{
    for (CMiniExplorerWnd* wnd : Registered<CMiniExplorerWnd>::get())
    {
        if (wnd->GetId() == id)
            return wnd;
    }
    return nullptr;
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

void OpenMRU(const ITEMIDLIST_ABSOLUTE* pidl, FOLDERFLAGS flags, FOLDERVIEWMODE ViewMode)
{
    int max = 0;
    CRegKey reg;
    reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\MRU"));

    {
        TCHAR name[1024];
        DWORD szname = 0;
        DWORD index = 0;
        while (szname = ARRAYSIZE(name), ERROR_SUCCESS == reg.EnumKey(index++, name, &szname))
        {
            CRegKey childreg;
            childreg.Open(reg, name);

            const int id = std::stoi(name);
            if (id > max)
                max = id;

            CComHeapPtr<ITEMIDLIST_ABSOLUTE> spidl;
            ULONG bytes = 0;
            childreg.QueryBinaryValue(_T("pidl"), nullptr, &bytes);
            if (ILGetSize(pidl) == bytes)
            {
                spidl.AllocateBytes(bytes);
                childreg.QueryBinaryValue(_T("pidl"), spidl, &bytes);

                std::vector<unsigned char> mru;
                ULONG bytes = 0;
                reg.QueryBinaryValue(_T("mru"), nullptr, &bytes);
                mru.resize(bytes / sizeof(unsigned char));
                reg.QueryBinaryValue(_T("mru"), mru.data(), &bytes);

                mru.erase(std::remove(mru.begin(), mru.end(), id), mru.end());
                mru.insert(mru.begin(), id);

                reg.SetBinaryValue(_T("mru"), mru.data(), (ULONG) mru.size() * sizeof(unsigned char));

                if (ILIsEqual(pidl, spidl))
                {
                    CMiniExplorerWnd* wnd = GetMiniExplorer(-id);
                    if (wnd == nullptr)
                        OpenMiniExplorer(reg, name, -id, SW_SHOW);
                    else
                        SetForegroundWindow(*wnd);
                    return;
                }
            }
        }
    }
    // TODO Potentially resuse an id if there are too many (need to keep an MRU list)
    const int id = -(max + 1);
    ATLVERIFY(SUCCEEDED(BrowseFolder(id, pidl, flags, ViewMode, SW_SHOW, nullptr)));
}

bool ParseCommandLine(_In_ PCWSTR lpCmdLine, _In_ int nShowCmd)
{
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(lpCmdLine, &argc);
    if (argc > 1 && _wcsicmp(argv[1], _T("/Add")) == 0)
    {
        AddMiniExplorer(SW_SHOW);
        CreateJumpList();
        return true;
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
        int id = std::stoi(argv[2]);

        CMiniExplorerWnd* wnd = GetMiniExplorer(id);

        if (wnd == nullptr)
        {
            CRegKey reg;
            reg.Create(HKEY_CURRENT_USER, _T("Software\\RadSoft\\MiniExplorer\\Windows"));

            if (!OpenMiniExplorer(reg, argv[2], std::stoi(argv[2]), nShowCmd))
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
            int id = std::stoi(name);

            CMiniExplorerWnd* wnd = GetMiniExplorer(id);

            if (wnd == nullptr)
                OpenMiniExplorer(reg, name, id, nShowCmd);

            if (s_id <= id)
                s_id = id + 1;
        }
        _putts(name);

        return true;
    }
    else
        return false;
}

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
