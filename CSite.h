class CSite : public IOleClientSite, public IAdviseSink
{
public:
    // IUnknown interface
    STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    // IOleClientSite interface
    STDMETHODIMP SaveObject() override;
    STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk) override;
    STDMETHODIMP GetContainer(IOleContainer **ppContainer) override;
    STDMETHODIMP ShowObject() override;
    STDMETHODIMP OnShowWindow(BOOL fShow) override;
    STDMETHODIMP RequestNewObjectLayout() override;

    // IAdviseSink interface
    STDMETHODIMP_(void) OnDataChange(FORMATETC *pFormatetc, STGMEDIUM *pStgmed) override;
    STDMETHODIMP_(void) OnViewChange(DWORD dwAspect, LONG lindex) override;
    STDMETHODIMP_(void) OnRename(IMoniker *pmk) override;
    STDMETHODIMP_(void) OnSave() override;
    STDMETHODIMP_(void) OnClose() override;

    CSite();
    ~CSite();
    void DoInit(IOleObject *pOlebject, IStorage *pStorage, LPWSTR lpszObjectName, LPRECT lprc);
    void DoSave();
    void DoRunObject(LONG lVerb);
    void DoClose();
    void DoDraw(HDC hdc);
    void DoAddVerbMenu(HMENU hmenu, UINT uIdVerbMin, UINT uIdVerbMax);
    void DoChangeSelect(BOOL bSelect);
    void DoGetRect(LPRECT lprc);
    void DoUpdateRect();
    void DoMoveRect(LPARAM lParam);
    void DoSetClickPos(LPARAM lParam);
    void DoHIMETRICtoDP(LPSIZEL lpSizel);
    CSite *m_pNext;

protected:
    LONG       m_cRef;
    RECT       m_rc;
    POINT      m_ptClick;
    DWORD      m_dwConnection;
    BOOL       m_bSelect;
    BOOL       m_bShowWindow;
    IOleObject *m_pOleObject;
    IStorage   *m_pStorage;
};
