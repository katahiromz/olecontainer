class CSite;

class CContainer : public IUnknown
{
public:
    // IUnknown interface
    STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    CContainer();
    ~CContainer();
    int DoRun(HINSTANCE hinst, int nCmdShow);
    LRESULT WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    BOOL DoCreateRootStorage(LPWSTR lpszFileName);
    BOOL DoInsertObject();
    BOOL DoSaveObject();
    BOOL DoLoadObject();
    void DoDestroy();
    void DoChangeSelection(CSite *pNew);
    BOOL DoGetSiteFormPos(LPARAM lParam, CSite **pp);
    void DoGetWindow(HWND *phwnd);
    void DoInitMenuItem(HMENU hmenu, LPCTSTR lpszItemName, int nId, HMENU hmenuSub);
    BOOL DoFileDialog(BOOL bOpen, LPWSTR lpszFilePath, DWORD dwLength);

private:
    LONG     m_cRef;
    HWND     m_hwnd;
    HMENU    m_hmenu;
    int      m_nObjectCount;
    IStorage *m_pRootStorage;
    IStream  *m_pStream;
    CSite    *m_pSite;
    CSite    *m_pSiteMoving;
    WCHAR    m_szFilePath[MAX_PATH];
};

extern CContainer *g_pContainer;
