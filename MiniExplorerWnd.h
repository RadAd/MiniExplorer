#pragma once

#include <atlbase.h>
#include <atlwin.h>
#include <atltypes.h>
#include <atlapp.h>
#include <atlcrack.h>

#include <Shlobj.h>

#include <set>

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

typedef CWinTraits<WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, WS_EX_ACCEPTFILES | WS_EX_APPWINDOW /*| WS_EX_TOOLWINDOW*/>		CMiniExplorerTraits;

struct MiniExplorerSettings
{
    FOLDERVIEWMODE ViewMode;
    FOLDERFLAGS flags;
    int iIconSize;
};

class CMiniExplorerWnd : public CWindowImpl<CMiniExplorerWnd, CWindow, CMiniExplorerTraits>,
    public Registered<CMiniExplorerWnd>
{
public:
    static LPCTSTR GetWndCaption()
    {
        return _T("Mini Explorer");
    }

    CMiniExplorerWnd(int id, CComPtr<IShellFolder> pShellFolder)
        : Registered<CMiniExplorerWnd>(this)
        , m_id(id), m_pShellFolder(pShellFolder)
    {
        ATLVERIFY(m_pShellFolder != nullptr);
    }

    BEGIN_MSG_MAP(CMiniExplorerWnd)
        MSG_WM_CREATE(OnCreate)
        MSG_WM_DESTROY(OnDestroy)
        MSG_WM_SIZE(OnSize)
        MSG_WM_SETFOCUS(OnSetFocus)
        MSG_WM_DPICHANGED(OnDpiChanged)
    END_MSG_MAP()

    int GetId() const
    {
        return m_id;
    }

    void SetId(int id)
    {
        m_id = id;
    }

    void GetPidl(_Outptr_ PIDLIST_ABSOLUTE* ppidl) const
    {
        ATLVERIFY(SUCCEEDED(SHGetIDListFromObject(m_pShellFolder, ppidl)));
    }

    void OnFinalMessage(_In_ HWND /*hWnd*/) override
    {
        delete this;
    }

    HRESULT TranslateAccelerator(MSG* pmsg)
    {
        return m_pShellView->TranslateAccelerator(pmsg);
    }

private:
    int OnCreate(LPCREATESTRUCT lpCreateStruct);
    void OnDestroy();
    void OnSize(UINT nType, CSize size);
    void OnSetFocus(CWindow wndOld);
    void OnDpiChanged(UINT nDpiX, UINT nDpiY, PRECT pRect);

    int m_id;
    CComPtr<IShellFolder> m_pShellFolder;
    CComPtr<IShellView> m_pShellView;

#ifdef USE_EXPLORER_BROWSER
    CComPtr<IExplorerBrowser> m_pExplorerBrowser;
#else
    CComPtr<IShellBrowser> m_pShellBrowser;
    CWindow m_hViewWnd;
#endif

    CComPtr<DShellFolderViewEvents> m_pEvents;
    DWORD m_dwCookie;
};
