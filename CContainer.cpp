#include <windows.h>
#include <oledlg.h>
#include "CContainer.h"
#include "CSite.h"

#define ID_FILE 100
#define ID_OPEN 200
#define ID_SAVE 300
#define ID_INSERT 400
#define ID_VERBMIN 0
#define ID_VERBMAX 5

#pragma comment (lib, "oledlg.lib")

const CLSID CLSID_SampleFile = {0x8f05f2d2, 0x9ec2, 0x43ac, {0xb6, 0x69, 0x0, 0x7e, 0x3f, 0x80, 0x7, 0xfc}};

CContainer *g_pContainer = NULL;

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	return g_pContainer->WindowProc(hwnd, uMsg, wParam, lParam);
}
// CContainer

CContainer::CContainer()
{
	m_cRef = 1;
	m_hwnd = NULL;
	m_hmenu = NULL;
	m_nObjectCount = 0;
	m_pRootStorage = NULL;
	m_pStream = NULL;
	m_pSite = NULL;
	m_pSiteMoving = NULL;
	m_szFilePath[0] = '\0';
}

CContainer::~CContainer()
{
}

STDMETHODIMP CContainer::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;	

	if (IsEqualIID(riid, IID_IUnknown))
		*ppvObject = static_cast<IUnknown *>(this);
	else
		return E_NOINTERFACE;

	AddRef();
	
	return S_OK;
}

STDMETHODIMP_(ULONG) CContainer::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CContainer::Release()
{
	if (InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

int CContainer::DoRun(HINSTANCE hinst, int nCmdShow)
{
	TCHAR      szAppName[] = TEXT("sample");
	HWND       hwnd;
	MSG        msg;
	WNDCLASSEX wc;

	wc.cbSize        = sizeof(WNDCLASSEX);
	wc.style         = CS_DBLCLKS;
	wc.lpfnWndProc   = ::WindowProc;
	wc.cbClsExtra    = 0;
	wc.cbWndExtra    = 0;
	wc.hInstance     = hinst;
	wc.hIcon         = (HICON)LoadImage(NULL, IDI_APPLICATION, IMAGE_ICON, 0, 0, LR_SHARED);
	wc.hCursor       = (HCURSOR)LoadImage(NULL, IDC_ARROW, IMAGE_CURSOR, 0, 0, LR_SHARED);
	wc.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
	wc.lpszMenuName  = NULL;
	wc.lpszClassName = szAppName;
	wc.hIconSm       = (HICON)LoadImage(NULL, IDI_APPLICATION, IMAGE_ICON, 0, 0, LR_SHARED);
	
	if (RegisterClassEx(&wc) == 0)
		return 0;

	hwnd = CreateWindowEx(0, szAppName, szAppName, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hinst, NULL);
	if (hwnd == NULL)
		return 0;

	ShowWindow(hwnd, nCmdShow);
	UpdateWindow(hwnd);
	
	while (GetMessage(&msg, NULL, 0, 0) > 0) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return (int)msg.wParam;
}

LRESULT CContainer::WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg) {

	case WM_CREATE: {
		HMENU hmenu;
		HMENU hmenuPopupFile;
		HMENU hmenuPopupHelp;

		hmenu  = CreateMenu();
		hmenuPopupFile = CreatePopupMenu();
		hmenuPopupHelp = CreatePopupMenu();

		DoInitMenuItem(hmenuPopupFile, TEXT("&Open"), ID_OPEN, NULL);
		DoInitMenuItem(hmenuPopupFile, TEXT("&Save"), ID_SAVE, NULL);

		DoInitMenuItem(hmenu, TEXT("&File"), ID_FILE, hmenuPopupFile);
		DoInitMenuItem(hmenu, TEXT("&Insert Object"), ID_INSERT, NULL);
		
		m_hwnd = hwnd;
		m_hmenu = hmenu;
		SetMenu(m_hwnd, m_hmenu);

		if (!DoCreateRootStorage(NULL))
			return -1;

		return 0;
	}
	
	case WM_COMMAND: {
		int nId = LOWORD(wParam);

		if (nId == ID_OPEN)
			DoLoadObject();
		else if (nId == ID_SAVE)
			DoSaveObject();
		else if (nId == ID_INSERT)
			DoInsertObject();
		else if (nId >= ID_VERBMIN || nId <= ID_VERBMAX)
			m_pSiteMoving->DoRunObject(nId);
		else
			;
		InvalidateRect(m_hwnd, NULL, TRUE);
		
		return 0;
	}

	case WM_LBUTTONDOWN: {
		CSite *p = NULL;
		BOOL  bChangeSelect = TRUE;

		if (DoGetSiteFormPos(lParam, &p)) {
			if (p == m_pSiteMoving)
				bChangeSelect = FALSE;
			SetCapture(hwnd);
			m_pSiteMoving = p;
			m_pSiteMoving->DoSetClickPos(lParam);
		}

		if (bChangeSelect) {
			DoChangeSelection(p);
			InvalidateRect(m_hwnd, NULL, TRUE);
		}

		return 0;
	}

	case WM_LBUTTONUP:
		ReleaseCapture();
		m_pSiteMoving = NULL;
		return 0;
	
	case WM_LBUTTONDBLCLK: {
		CSite *p;

		if (DoGetSiteFormPos(lParam, &p))
			p->DoRunObject(OLEIVERB_PRIMARY);

		return 0;
	}
	
	case WM_RBUTTONDOWN: {
		CSite *p = NULL;
		POINT pt;
		HMENU hmenu;

		if (DoGetSiteFormPos(lParam, &p)) {
			if (p != m_pSiteMoving) {
				m_pSiteMoving = p;
				DoChangeSelection(p);
				InvalidateRect(m_hwnd, NULL, TRUE);
			}

			hmenu = CreatePopupMenu();
			p->DoAddVerbMenu(hmenu, ID_VERBMIN, ID_VERBMAX);
			GetCursorPos(&pt);
			TrackPopupMenu(hmenu, 0, pt.x, pt.y, 0, hwnd, NULL);
			DestroyMenu(hmenu);
		}
		else {
			DoChangeSelection(p);
			InvalidateRect(m_hwnd, NULL, TRUE);
		}
		
		return 0;
	}

	case WM_MOUSEMOVE:
		if (GetCapture() != NULL) {
			m_pSiteMoving->DoMoveRect(lParam);
			InvalidateRect(m_hwnd, NULL, TRUE);
		}
		return 0;
	
	case WM_PAINT: {
		HDC         hdc;
		PAINTSTRUCT ps;
		CSite       *p = m_pSite;

		hdc = BeginPaint(hwnd, &ps);

		while (p != NULL) {
			p->DoDraw(hdc);
			p = p->m_pNext;
		}

		EndPaint(hwnd, &ps);

		return 0;
	}

	case WM_DESTROY:
		DoDestroy();
		PostQuitMessage(0);
		return 0;

	default:
		break;

	}

	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

BOOL CContainer::DoCreateRootStorage(LPWSTR lpszFileName)
{
	CLSID clsid;
	WCHAR szStreamName[] = L"stream";

	if (lpszFileName == NULL) {
		StgCreateDocfile(NULL, STGM_READWRITE | STGM_TRANSACTED, 0, &m_pRootStorage);
		if (m_pRootStorage == NULL)
			return FALSE;

		WriteClassStg(m_pRootStorage, CLSID_SampleFile);
		
		m_pRootStorage->CreateStream(szStreamName, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, &m_pStream);
		if (m_pStream == NULL)
			return FALSE;
	}
	else {
		IStorage *pStorage;
		IStream  *pStream;

		StgOpenStorage(lpszFileName, NULL, STGM_READWRITE | STGM_TRANSACTED, NULL, 0, &pStorage);
		if (pStorage == NULL)
			return FALSE;

		ReadClassStg(pStorage, &clsid);
		if (!IsEqualCLSID(clsid, CLSID_SampleFile)) {
			pStorage->Release();
			return FALSE;
		}

		pStorage->OpenStream(szStreamName, NULL, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, &pStream);
		if (pStream == NULL) {
			pStorage->Release();
			return FALSE;
		}

		DoDestroy();
	
		m_pRootStorage = pStorage;
		m_pStream = pStream;
	}

	return TRUE;
}

BOOL CContainer::DoInsertObject()
{
	UINT              uResult;
	WCHAR             szFilePath[MAX_PATH];
	WCHAR             szObjectName[256];
	IStorage          *pStorage;
	IOleObject        *pOleObject;
	CSite             *p;
	OLEUIINSERTOBJECT insertObject;

	wsprintfW(szObjectName, L"Object %d", m_nObjectCount + 1);
	m_pRootStorage->CreateStorage(szObjectName, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, &pStorage);
	if (pStorage == NULL)
		return FALSE;

	szFilePath[0] = '\0';

	ZeroMemory(&insertObject, sizeof(OLEUIINSERTOBJECT));
	insertObject.cbStruct   = sizeof(OLEUIINSERTOBJECT);
	insertObject.dwFlags    = IOF_DISABLEDISPLAYASICON | IOF_SELECTCREATENEW | /*IOF_CREATENEWOBJECT |*/ IOF_CREATEFILEOBJECT | IOF_DISABLELINK;
	insertObject.hWndOwner  = m_hwnd;
	insertObject.lpszFile   = szFilePath;
	insertObject.cchFile    = MAX_PATH;
	insertObject.iid        = IID_IOleObject;
	insertObject.oleRender  = OLERENDER_DRAW;
	insertObject.lpIStorage = pStorage;
	//insertObject.ppvObj     =  (void **)&pOleObject;
	
	uResult = OleUIInsertObject(&insertObject);
	if (uResult != OLEUI_OK) {
		m_pRootStorage->DestroyElement(szObjectName);
		return FALSE;
	}

    OleCreate(insertObject.clsid,
              insertObject.iid,
              insertObject.oleRender,
              insertObject.lpFormatEtc,
              insertObject.lpIOleClientSite,
              insertObject.lpIStorage,
              (void **)&pOleObject);

	m_nObjectCount++;

	if (m_pSite == NULL) {
		m_pSite = new CSite;
		p = m_pSite;
	}
	else {
		p = m_pSite;
		while (p->m_pNext != NULL)
			p = p->m_pNext;

		p->m_pNext = new CSite;
		p = p->m_pNext;
	}
	
	p->DoInit(pOleObject, pStorage, szObjectName, NULL);

	return TRUE;
}

BOOL CContainer::DoSaveObject()
{
	int           i;
	ULONG         uResult;
	RECT          rc;
	CSite         *p = m_pSite;
	WCHAR         szFilePath[MAX_PATH];
	LARGE_INTEGER li = {0};

	if (m_nObjectCount == 0) {
		MessageBox(NULL, TEXT("Object not created"), NULL, MB_ICONWARNING);
		return FALSE;
	}

	if (!DoFileDialog(FALSE, szFilePath, MAX_PATH))
		return FALSE;

	m_pStream->Seek(li, STREAM_SEEK_SET, NULL);
	m_pStream->Write(&m_nObjectCount, sizeof(int), &uResult);

	for (i = 0; i < m_nObjectCount; i++) {
		p->DoGetRect(&rc);
		m_pStream->Write(&rc, sizeof(RECT), &uResult);

		p->DoSave();
		p = p->m_pNext;
	}

	if (lstrcmpW(szFilePath, m_szFilePath) != 0) {
		HRESULT      hr;
		IRootStorage *pRootStorage;

		m_pRootStorage->QueryInterface(IID_PPV_ARGS(&pRootStorage));
		hr = pRootStorage->SwitchToFile(szFilePath);
		pRootStorage->Release();
		if (FAILED(hr)) {
			m_pRootStorage->Revert();
			return FALSE;
		}

		lstrcpyW(m_szFilePath, szFilePath);
	}

	m_pRootStorage->Commit(STGC_DEFAULT);

	return TRUE;
}

BOOL CContainer::DoLoadObject()
{
	int        i;
	ULONG      uResult;
	RECT       rc;
	WCHAR      szObjectName[256];
	WCHAR      szFilePath[MAX_PATH];
	CSite      *p = NULL;
	IOleObject *pOleObject;
	IStorage   *pStorage;

	if (!DoFileDialog(TRUE, szFilePath, MAX_PATH))
		return FALSE;

	if (!DoCreateRootStorage(szFilePath))
		return FALSE;

	lstrcpyW(m_szFilePath, szFilePath);

	m_pStream->Read(&m_nObjectCount, sizeof(int), &uResult);
	for (i = 0; i < m_nObjectCount; i++) {
		wsprintf(szObjectName, L"Object %d", i + 1); 
		m_pRootStorage->OpenStorage(szObjectName, NULL, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, NULL, 0, &pStorage);
		OleLoad(pStorage, IID_IOleObject, NULL, (void **)&pOleObject);

		m_pStream->Read(&rc, sizeof(RECT), &uResult);

		if (i == 0) {
			m_pSite = new CSite;
			p = m_pSite;
		}
		else {
			p->m_pNext = new CSite;
			p = p->m_pNext;
		}

		p->DoInit(pOleObject, pStorage, szObjectName, &rc);
	}

	return TRUE;
}

void CContainer::DoDestroy()
{
	CSite *p = m_pSite;
	CSite *p2;
	
	while (p != NULL) {
		p2 = p->m_pNext;
		
		p->DoClose();
		p->Release();
		p = p2;
	}
	
	if (m_pStream != NULL)
		m_pStream->Release();

	if (m_pRootStorage != NULL)
		m_pRootStorage->Release();
}

void CContainer::DoChangeSelection(CSite *pNew)
{
	CSite *p = m_pSite;

	while (p != NULL) {
		p->DoChangeSelect(FALSE);
		p = p->m_pNext;
	}

	if (pNew != NULL)
		pNew->DoChangeSelect(TRUE);
}

BOOL CContainer::DoGetSiteFormPos(LPARAM lParam, CSite **pp)
{
	POINT pt;
	RECT  rc;
	CSite *p = m_pSite;
	
	pt.x = LOWORD(lParam);
	pt.y = HIWORD(lParam);
	
	while (p != NULL) {
		p->DoGetRect(&rc);
		if (PtInRect(&rc, pt)) {
			*pp = p;
			return TRUE;
		}
		p = p->m_pNext;
	}
	
	*pp = NULL;

	return FALSE;
}

void CContainer::DoGetWindow(HWND *phwnd)
{
	*phwnd = m_hwnd;
}

void CContainer::DoInitMenuItem(HMENU hmenu, LPCTSTR lpszItemName, int nId, HMENU hmenuSub)
{
	MENUITEMINFO mii;
	
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask  = MIIM_ID | MIIM_TYPE;
	mii.wID    = nId;

	if (lpszItemName != NULL) {
		mii.fType      = MFT_STRING;
	    mii.dwTypeData = const_cast<LPTSTR>(lpszItemName);
	}
	else
		mii.fType = MFT_SEPARATOR;

	if (hmenuSub != NULL) {
		mii.fMask   |= MIIM_SUBMENU;
		mii.hSubMenu = hmenuSub;
	}

	InsertMenuItem(hmenu, nId, FALSE, &mii);
}

BOOL CContainer::DoFileDialog(BOOL bOpen, LPWSTR lpszFilePath, DWORD dwLength)
{
	OPENFILENAMEW ofn;
	LPCWSTR        lpszFilter = L"ols File (*.ols)\0*.ols\0\0";

	lpszFilePath[0] = '\0';

	if (bOpen) {
		ZeroMemory(&ofn, sizeof(OPENFILENAME));
		ofn.lStructSize = sizeof(OPENFILENAME);
		ofn.hwndOwner   = m_hwnd;
		ofn.lpstrFilter = lpszFilter;
		ofn.lpstrFile   = lpszFilePath;
		ofn.nMaxFile    = dwLength;
		ofn.lpstrTitle  = L"Load File";
		ofn.Flags       = OFN_FILEMUSTEXIST;

		if (!GetOpenFileNameW(&ofn))
			return FALSE;
	}
	else {
		ZeroMemory(&ofn, sizeof(OPENFILENAME));
		ofn.lStructSize = sizeof(OPENFILENAME);
		ofn.hwndOwner   = m_hwnd;
		ofn.lpstrFilter = lpszFilter;
		ofn.lpstrFile   = lpszFilePath;
		ofn.nMaxFile    = dwLength;
		ofn.lpstrTitle  = L"Save File";
		ofn.Flags       = OFN_FILEMUSTEXIST;
		ofn.lpstrDefExt = L"ols";

		if (!GetSaveFileNameW(&ofn))
			return FALSE;
	}

	return TRUE;
}

int WINAPI WinMain(HINSTANCE hinst, HINSTANCE hinstPrev, LPSTR lpszCmdLine, int nCmdShow)
{
	int nResult = 0;

	OleInitialize(NULL);
	
	g_pContainer = new CContainer;
	if (g_pContainer != NULL) {
		nResult = g_pContainer->DoRun(hinst, nCmdShow);
		g_pContainer->Release();
	}

	OleUninitialize();

	return nResult;
}
