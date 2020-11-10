#include <windows.h>
#include <oledlg.h>
#include "CContainer.h"
#include "CSite.h"

CSite::CSite()
{
    m_cRef = 1;
    m_dwConnection = 0;
    m_bSelect = FALSE;
    m_bShowWindow = FALSE;
    m_pOleObject = NULL;
    m_pStorage = NULL;
    m_pNext = NULL;
    SetRectEmpty(&m_rc);
}

CSite::~CSite()
{
    if (m_pOleObject != NULL)
        m_pOleObject->Release();
    if (m_pStorage != NULL)
        m_pStorage->Release();
}

STDMETHODIMP CSite::QueryInterface(REFIID riid, void **ppvObject)
{
    *ppvObject = NULL;  

    if (riid == IID_IUnknown || riid == IID_IOleClientSite)
        *ppvObject = static_cast<IOleClientSite *>(this);
    else if (riid == IID_IAdviseSink)
        *ppvObject = static_cast<IAdviseSink *>(this);
    else
        return E_NOINTERFACE;

    AddRef();
    
    return S_OK;
}

STDMETHODIMP_(ULONG) CSite::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CSite::Release()
{
    if (InterlockedDecrement(&m_cRef) == 0) {
        delete this;
        return 0;
    }

    return m_cRef;
}

STDMETHODIMP CSite::SaveObject()
{
    HWND hwnd;

    DoSave();

    DoUpdateRect();
    g_pContainer->DoGetWindow(&hwnd);
    InvalidateRect(hwnd, NULL, TRUE);

    return S_OK;
}

STDMETHODIMP CSite::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk)
{
    return E_NOTIMPL;
}

STDMETHODIMP CSite::GetContainer(IOleContainer **ppContainer)
{
    *ppContainer = NULL;
    return E_NOINTERFACE;
}

STDMETHODIMP CSite::ShowObject()
{
    return S_OK;
}

STDMETHODIMP CSite::OnShowWindow(BOOL fShow)
{
    HWND hwnd;

    m_bShowWindow = fShow;

    g_pContainer->DoGetWindow(&hwnd);
    InvalidateRect(hwnd, NULL, TRUE);

    return S_OK;
}

STDMETHODIMP CSite::RequestNewObjectLayout()
{
    return E_NOTIMPL;
}

STDMETHODIMP_(void) CSite::OnDataChange(FORMATETC *pFormatetc, STGMEDIUM *pStgmed)
{
}

STDMETHODIMP_(void) CSite::OnViewChange(DWORD dwAspect, LONG lindex)
{
    HWND hwnd;

    DoUpdateRect();

    g_pContainer->DoGetWindow(&hwnd);
    InvalidateRect(hwnd, NULL, TRUE);
}

STDMETHODIMP_(void) CSite::OnRename(IMoniker *pmk)
{
}

STDMETHODIMP_(void) CSite::OnSave()
{   
}

STDMETHODIMP_(void) CSite::OnClose()
{
}

void CSite::DoInit(IOleObject *pOlebject, IStorage *pStorage, LPWSTR lpszObjectName, LPRECT lprc)
{
    m_pStorage = pStorage;

    m_pOleObject = pOlebject;
    m_pOleObject->SetClientSite(static_cast<IOleClientSite *>(this));
    m_pOleObject->SetHostNames(L"sample", lpszObjectName);

    if (lprc != NULL)
        m_rc = *lprc;
    else
        DoUpdateRect();
}

void CSite::DoSave()
{
    IPersistStorage *pPersistStorage;

    m_pOleObject->QueryInterface(IID_PPV_ARGS(&pPersistStorage));

    OleSave(pPersistStorage, m_pStorage, TRUE);

    pPersistStorage->SaveCompleted(NULL);
    pPersistStorage->Release();
}

void CSite::DoRunObject(LONG lVerb)
{
    HWND        hwnd;
    IViewObject *pViewObject;
    
    m_pOleObject->Advise(static_cast<IAdviseSink *>(this), &m_dwConnection);

    m_pOleObject->QueryInterface(IID_PPV_ARGS(&pViewObject));
    pViewObject->SetAdvise(DVASPECT_CONTENT, 0, static_cast<IAdviseSink *>(this));
    pViewObject->Release();

    g_pContainer->DoGetWindow(&hwnd);
    m_pOleObject->DoVerb(lVerb, NULL, static_cast<IOleClientSite *>(this), 0, hwnd, &m_rc);
}

void CSite::DoClose()
{
    if (OleIsRunning(m_pOleObject))
    {
        m_pOleObject->Close(OLECLOSE_NOSAVE);
        m_pOleObject->Unadvise(m_dwConnection);
    }
}

void CSite::DoDraw(HDC hdc)
{
    OleDraw(m_pOleObject, DVASPECT_CONTENT, hdc, &m_rc);

    if (m_bShowWindow) {
        HBRUSH hbr;
        hbr = CreateHatchBrush(HS_BDIAGONAL, RGB(128, 128, 128));
        SelectObject(hdc, hbr);
        SetBkMode(hdc, TRANSPARENT);
        Rectangle(hdc, m_rc.left, m_rc.top, m_rc.right, m_rc.bottom);
        DeleteObject(hbr);
    }
    
    if (m_bSelect)
        DrawFocusRect(hdc, &m_rc);
    else
        FrameRect(hdc, &m_rc, (HBRUSH)GetStockObject(BLACK_BRUSH));
}

void CSite::DoAddVerbMenu(HMENU hmenu, UINT uIdVerbMin, UINT uIdVerbMax)
{
    HMENU hmenuTmp;
    OleUIAddVerbMenu(m_pOleObject, NULL, hmenu, 0, uIdVerbMin, uIdVerbMax, FALSE, 0, &hmenuTmp);
}

void CSite::DoChangeSelect(BOOL bSelect)
{
    m_bSelect = bSelect;
}

void CSite::DoGetRect(LPRECT lprc)
{
    *lprc = m_rc;
}

void CSite::DoUpdateRect()
{
    SIZEL sizel;

    m_pOleObject->GetExtent(DVASPECT_CONTENT, &sizel);

    DoHIMETRICtoDP(&sizel);
    m_rc.right = m_rc.left + sizel.cx;
    m_rc.bottom = m_rc.top + sizel.cy;
}

void CSite::DoMoveRect(LPARAM lParam)
{
    int x, y;

    x = LOWORD(lParam) - m_ptClick.x;
    y = HIWORD(lParam) - m_ptClick.y;

    m_ptClick.x = LOWORD(lParam);
    m_ptClick.y = HIWORD(lParam);

    OffsetRect(&m_rc, x, y);
}

void CSite::DoSetClickPos(LPARAM lParam)
{
    m_ptClick.x = LOWORD(lParam);
    m_ptClick.y = HIWORD(lParam);
}

void CSite::DoHIMETRICtoDP(LPSIZEL lpSizel)
{
    HDC hdc;
    const int HIMETRIC_INCH = 2540;

    hdc = GetDC(NULL);
    lpSizel->cx = lpSizel->cx * GetDeviceCaps(hdc, LOGPIXELSX) / HIMETRIC_INCH;
    lpSizel->cy = lpSizel->cy * GetDeviceCaps(hdc, LOGPIXELSY) / HIMETRIC_INCH;
    ReleaseDC(NULL, hdc);
}
