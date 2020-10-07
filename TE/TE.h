#pragma once

class CteFolderItem : public FolderItem, public IPersistFolder2, public IParentAndItem
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IPersist
    STDMETHODIMP GetClassID(CLSID *pClassID);
	//IPersistFolder
    STDMETHODIMP Initialize(PCIDLIST_ABSOLUTE pidl);
	//IPersistFolder2
    STDMETHODIMP GetCurFolder(PIDLIST_ABSOLUTE *ppidl);
	//IParentAndItem
	STDMETHODIMP GetParentAndItem(PIDLIST_ABSOLUTE *ppidlParent, IShellFolder **ppsf, PITEMID_CHILD *ppidlChild);
	STDMETHODIMP SetParentAndItem(PCIDLIST_ABSOLUTE pidlParent, IShellFolder *psf, PCUITEMID_CHILD pidlChild);
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//FolderItem
    STDMETHODIMP get_Application(IDispatch **ppid);
    STDMETHODIMP get_Parent(IDispatch **ppid);
    STDMETHODIMP get_Name(BSTR *pbs);
    STDMETHODIMP put_Name(BSTR bs);
    STDMETHODIMP get_Path(BSTR *pbs);
    STDMETHODIMP get_GetLink(IDispatch **ppid);
    STDMETHODIMP get_GetFolder(IDispatch **ppid);
    STDMETHODIMP get_IsLink(VARIANT_BOOL *pb);
    STDMETHODIMP get_IsFolder(VARIANT_BOOL *pb);
    STDMETHODIMP get_IsFileSystem(VARIANT_BOOL *pb);
    STDMETHODIMP get_IsBrowsable(VARIANT_BOOL *pb);
    STDMETHODIMP get_ModifyDate(DATE *pdt);
    STDMETHODIMP put_ModifyDate(DATE dt);
    STDMETHODIMP get_Size(LONG *pul);
    STDMETHODIMP get_Type(BSTR *pbs);
    STDMETHODIMP Verbs(FolderItemVerbs **ppfic);
    STDMETHODIMP InvokeVerb(VARIANT vVerb);
/*	//FolderItem2
    STDMETHODIMP InvokeVerbEx(VARIANT vVerb, VARIANT vArgs);
	STDMETHODIMP ExtendedProperty(BSTR bstrPropName, VARIANT *pvRet);
*/
	CteFolderItem(VARIANT *pv);
	~CteFolderItem();
	LPITEMIDLIST GetPidl();
	LPITEMIDLIST GetAlt();
	BOOL GetFolderItem();
	VOID Clear();
	BSTR GetStrPath();
	VOID MakeUnavailable();

	VARIANT			m_v;
	LPITEMIDLIST	m_pidl, m_pidlAlt;
	FolderItem		*m_pFolderItem;
	LPITEMIDLIST	m_pidlFocused;
	IDispatch		*m_pEnum;
	int				m_nSelected;
	BOOL			m_bStrict;
	DWORD			m_dwUnavailable;
private:
	LONG			m_cRef;
	DWORD			m_dwSessionId;
	DISPID			m_dispExtendedProperty;
};

class CteFolderItems : public FolderItems, public IDispatchEx, public IDataObject
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IDispatchEx
	STDMETHODIMP GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid);
	STDMETHODIMP InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller);
	STDMETHODIMP DeleteMemberByName(BSTR bstrName, DWORD grfdex);
	STDMETHODIMP DeleteMemberByDispID(DISPID id);
	STDMETHODIMP GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex);
	STDMETHODIMP GetMemberName(DISPID id, BSTR *pbstrName);
	STDMETHODIMP GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid);
	STDMETHODIMP GetNameSpaceParent(IUnknown **ppunk);
	//FolderItems
	STDMETHODIMP get_Count(long *plCount);
    STDMETHODIMP get_Application(IDispatch **ppid);
    STDMETHODIMP get_Parent(IDispatch **ppid);
    STDMETHODIMP Item(VARIANT index, FolderItem **ppid);
    STDMETHODIMP _NewEnum(IUnknown **ppunk);
	//IDataObject
	STDMETHODIMP GetData(FORMATETC *pformatetcIn, STGMEDIUM *pmedium);
	STDMETHODIMP GetDataHere(FORMATETC *pformatetc, STGMEDIUM *pmedium);
	STDMETHODIMP QueryGetData(FORMATETC *pformatetc);
	STDMETHODIMP GetCanonicalFormatEtc(FORMATETC *pformatectIn, FORMATETC *pformatetcOut);
	STDMETHODIMP SetData(FORMATETC *pformatetc, STGMEDIUM *pmedium, BOOL fRelease);
	STDMETHODIMP EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenumFormatEtc);
	STDMETHODIMP DAdvise(FORMATETC *pformatetc, DWORD advf, IAdviseSink *pAdvSink, DWORD *pdwConnection);
	STDMETHODIMP DUnadvise(DWORD dwConnection);
	STDMETHODIMP EnumDAdvise(IEnumSTATDATA **ppenumAdvise);

	CteFolderItems(IDataObject *pDataObj, FolderItems *pFolderItems);
	~CteFolderItems();

	HDROP GethDrop(int x, int y, BOOL fNC);
	VOID Regenerate(BOOL bFull);
	VOID ItemEx(long nIndex, VARIANT *pVarResult, VARIANT *pVarNew);
	BOOL CanIDListFormat();
	VOID AdjustIDListEx();
	VOID Clear();
	HRESULT QueryGetData2(FORMATETC *pformatetc);
public:
	int				m_nIndex;
	DWORD			m_dwEffect;
	LPITEMIDLIST	*m_pidllist;

private:
	VARIANT m_vData;
	IDataObject		*m_pDataObj;
	FolderItems		*m_pFolderItems;
	BSTR			m_bsText;
	std::vector<FolderItem *>	m_ovFolderItems;

	LONG			m_cRef;
	LONG			m_nCount;
	int				m_nUseIDListFormat;
	BOOL			m_bUseText;
	BOOL			m_oFolderItems;
};

class CteDropTarget2 : public IDropTarget
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDropTarget
	STDMETHODIMP DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
	STDMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
	STDMETHODIMP DragLeave();
	STDMETHODIMP Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);

	CteDropTarget2(HWND hwnd, IUnknown *punk, BOOL bUseHelper);
	~CteDropTarget2();
public:
	CteFolderItems *m_pDragItems;
	IDropTarget *m_pDropTarget;
	HWND        m_hwnd;
	IUnknown	*m_punk;

	DWORD	m_grfKeyState;
	HRESULT m_DragLeave;
	LONG	m_cRef;
	BOOL	m_bUseHelper;
};

class CTE : public IDispatch, public IDropSource
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IDropSource
	STDMETHODIMP QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState);
	STDMETHODIMP GiveFeedback(DWORD dwEffect);

	CTE(HWND hwnd);
	~CTE();
private:
	LONG	m_cRef;
	HWND	m_hwnd;
};
//
class CteInternetSecurityManager : public IInternetSecurityManager
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IInternetSecurityManager
	STDMETHODIMP SetSecuritySite(IInternetSecurityMgrSite *pSite);
	STDMETHODIMP GetSecuritySite(IInternetSecurityMgrSite **ppSite);
	STDMETHODIMP MapUrlToZone(LPCWSTR pwszUrl, DWORD *pdwZone, DWORD dwFlags);
	STDMETHODIMP GetSecurityId(LPCWSTR pwszUrl,	BYTE *pbSecurityId, DWORD *pcbSecurityId, DWORD_PTR dwReserved);
	STDMETHODIMP ProcessUrlAction(LPCWSTR pwszUrl, DWORD dwAction, BYTE *pPolicy, DWORD cbPolicy, BYTE *pContext, DWORD cbContext, DWORD dwFlags, DWORD dwReserved);
	STDMETHODIMP QueryCustomPolicy(LPCWSTR pwszUrl, REFGUID guidKey, BYTE **ppPolicy, DWORD *pcbPolicy, BYTE *pContext, DWORD cbContext, DWORD dwReserved);
	STDMETHODIMP SetZoneMapping(DWORD dwZone, LPCWSTR lpszPattern, DWORD dwFlags);
	STDMETHODIMP GetZoneMappings(DWORD dwZone, IEnumString **ppenumString, DWORD dwFlags);

	CteInternetSecurityManager();
	~CteInternetSecurityManager();
private:
	LONG		m_cRef;
};

class CteNewWindowManager : public INewWindowManager
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//INewWindowManager
	STDMETHODIMP EvaluateNewWindow(LPCWSTR pszUrl, LPCWSTR pszName, LPCWSTR pszUrlContext, LPCWSTR pszFeatures, BOOL fReplace, DWORD dwFlags, DWORD dwUserActionTime);

	CteNewWindowManager();
	~CteNewWindowManager();
private:
	LONG		m_cRef;
};

class CteWebBrowser : public IDispatch, public IOleClientSite, public IOleInPlaceSite,
	public IDocHostUIHandler, public IDropTarget, public IDocHostShowUI//, public IOleCommandTarget
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IOleClientSite
	STDMETHODIMP SaveObject();
	STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk);
	STDMETHODIMP GetContainer(IOleContainer **ppContainer);
	STDMETHODIMP ShowObject();
	STDMETHODIMP OnShowWindow(BOOL fShow);
	STDMETHODIMP RequestNewObjectLayout();
	//IOleWindow
	STDMETHODIMP GetWindow(HWND *phwnd);
	STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
	//IOleInPlaceSite
	STDMETHODIMP CanInPlaceActivate();
	STDMETHODIMP OnInPlaceActivate();
	STDMETHODIMP OnUIActivate();
	STDMETHODIMP GetWindowContext(IOleInPlaceFrame **ppFrame, IOleInPlaceUIWindow **ppDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo);
	STDMETHODIMP Scroll(SIZE scrollExtant);
	STDMETHODIMP OnUIDeactivate(BOOL fUndoable);
	STDMETHODIMP OnInPlaceDeactivate();
	STDMETHODIMP DiscardUndoState();
	STDMETHODIMP DeactivateAndUndo();
	STDMETHODIMP OnPosRectChange(LPCRECT lprcPosRect);
	//IOleCommandTarget
//    STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD *prgCmds, OLECMDTEXT *pCmdText);
//    STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut);
	//IDocHostUIHandler
	STDMETHODIMP ShowContextMenu(DWORD dwID, POINT *ppt, IUnknown *pcmdtReserved, IDispatch *pdispReserved);
	STDMETHODIMP GetHostInfo(DOCHOSTUIINFO *pInfo);
	STDMETHODIMP ShowUI(DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc);
	STDMETHODIMP HideUI(VOID);
	STDMETHODIMP UpdateUI(VOID);
	STDMETHODIMP EnableModeless(BOOL fEnable);
	STDMETHODIMP OnDocWindowActivate(BOOL fActivate);
	STDMETHODIMP OnFrameWindowActivate(BOOL fActivate);
	STDMETHODIMP ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fFrameWindow);
	STDMETHODIMP TranslateAccelerator(LPMSG lpMsg, const GUID *pguidCmdGroup, DWORD nCmdID);
	STDMETHODIMP GetOptionKeyPath(LPOLESTR *pchKey, DWORD dw);
	STDMETHODIMP GetDropTarget(IDropTarget *pDropTarget, IDropTarget **ppDropTarget);
	STDMETHODIMP GetExternal(IDispatch **ppDispatch);
	STDMETHODIMP TranslateUrl(DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut);
	STDMETHODIMP FilterDataObject(IDataObject *pDO, IDataObject **ppDORet);
	//IDocHostShowUI
	STDMETHODIMP ShowMessage(HWND hwnd, LPOLESTR lpstrText, LPOLESTR lpstrCaption, DWORD dwType, LPOLESTR lpstrHelpFile, DWORD dwHelpContext, LRESULT *plResult);
	STDMETHODIMP ShowHelp(HWND hwnd, LPOLESTR pszHelpFile, UINT uCommand, DWORD dwData, POINT ptMouse, IDispatch *pDispatchObjectHit);//*/
	//IDropTarget
	STDMETHODIMP DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
	STDMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
	STDMETHODIMP DragLeave();
	STDMETHODIMP Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);

	CteWebBrowser(HWND hwndParent, WCHAR *szPath, VARIANT *pvArg);
	~CteWebBrowser();
	void Close();
	HWND get_HWND();
	VOID write(LPWSTR pszPath);
	BOOL IsBusy();
public:
	VARIANT m_vData;
	IWebBrowser2 *m_pWebBrowser;
	IDispatch	*m_ppDispatch[Count_WBFunc];
	BSTR	m_bstrPath;
	HWND	m_hwndParent;
	HWND	m_hwndBrowser;

	HRESULT m_DragLeave;
	BOOL	m_bRedraw;
	BOOL	m_nClose;
private:
	CteFolderItems *m_pDragItems;
	IDropTarget *m_pDropTarget;
	LONG	m_cRef;
	DWORD	m_dwCookie;
	DWORD	m_grfKeyState;
	DWORD	m_dwEffect;
	DWORD	m_dwEffectTE;
};

struct TEWebBrowsers
{
	CteWebBrowser	*pWB;
	HWND			hwnd;
};

class CteTabCtrl : public IDispatchEx
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IDispatchEx
	STDMETHODIMP GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid);
	STDMETHODIMP InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller);
	STDMETHODIMP DeleteMemberByName(BSTR bstrName, DWORD grfdex);
	STDMETHODIMP DeleteMemberByDispID(DISPID id);
	STDMETHODIMP GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex);
	STDMETHODIMP GetMemberName(DISPID id, BSTR *pbstrName);
	STDMETHODIMP GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid);
	STDMETHODIMP GetNameSpaceParent(IUnknown **ppunk);

	VOID TabChanged(BOOL bSameTC);
	CteShellBrowser * GetShellBrowser(int nPage);

	CteTabCtrl();
	~CteTabCtrl();

	VOID CreateTC();
	BOOL Create();
	VOID Init();
	BOOL Close(BOOL bForce);
	VOID Move(int nSrc, int nDest, CteTabCtrl *pDestTab);
	VOID LockUpdate();
	VOID UnlockUpdate();
	VOID RedrawUpdate();
	VOID Show(BOOL bVisible);
	VOID GetItem(int i, VARIANT *pVarResult);
	DWORD GetStyle();
	VOID SetItemSize();
	BOOL SetDefault();
public:
	SCROLLINFO m_si;
	HWND	m_hwnd;
	HWND	m_hwndStatic;
	HWND	m_hwndButton;

	DWORD	m_dwSize;
	LONG	m_nLockUpdate;
	int		m_nIndex;
	int		m_param[9];
	int		m_nTC;
	int		m_nScrollWidth;
	long	m_lNavigating;
	BOOL	m_bRedraw;
	BOOL	m_bEmpty;
	BOOL	m_bVisible;
private:
	VARIANT m_vData;
	CteDropTarget2 *m_pDropTarget2;
	LONG	m_cRef;
};

class CteServiceProvider : public IServiceProvider
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IServiceProvider
	STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppv);

	CteServiceProvider(IUnknown *punk, IUnknown *punk2);
	~CteServiceProvider();
public:
	IUnknown *m_pUnk;
	IUnknown *m_pUnk2;
	LONG	m_cRef;
};

class CteShellBrowser : public IShellBrowser, public ICommDlgBrowser2,
	public IShellFolderViewDual,
//	public IFolderFilter,
	public IExplorerBrowserEvents, public IExplorerPaneVisibility,
	public IShellFolderViewCB,
#ifdef _2000XP
	public IShellFolder2,
#elif _USE_SHELLBROWSER
	public IShellFolder,
#endif
	public IPersistFolder2
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IOleWindow
	STDMETHODIMP GetWindow(HWND *phwnd);
	STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
	//IShellBrowser
	STDMETHODIMP InsertMenusSB(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths);
	STDMETHODIMP SetMenuSB(HMENU hmenuShared, HOLEMENU holemenuRes, HWND hwndActiveObject);
	STDMETHODIMP RemoveMenusSB(HMENU hmenuShared);
	STDMETHODIMP SetStatusTextSB(LPCWSTR lpszStatusText);
	STDMETHODIMP EnableModelessSB(BOOL fEnable);
	STDMETHODIMP TranslateAcceleratorSB(LPMSG lpmsg, WORD wID);
	STDMETHODIMP BrowseObject(PCUIDLIST_RELATIVE pidl, UINT wFlags);
	STDMETHODIMP GetViewStateStream(DWORD grfMode, IStream **ppStrm);
	STDMETHODIMP GetControlWindow(UINT id, HWND *lphwnd);
	STDMETHODIMP SendControlMsg(UINT id, UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT *pret);
	STDMETHODIMP QueryActiveShellView(IShellView **ppshv);
	STDMETHODIMP OnViewWindowActive(IShellView *ppshv);
	STDMETHODIMP SetToolbarItems(LPTBBUTTONSB lpButtons, UINT nButtons, UINT uFlags);
	//ICommDlgBrowser
	STDMETHODIMP OnDefaultCommand(IShellView *ppshv);
	STDMETHODIMP OnStateChange(IShellView *ppshv, ULONG uChange);
	STDMETHODIMP IncludeObject(IShellView *ppshv, LPCITEMIDLIST pidl);
	//ICommDlgBrowser2
    STDMETHODIMP Notify(IShellView *ppshv, DWORD dwNotifyType);
    STDMETHODIMP GetDefaultMenuText(IShellView *ppshv, LPWSTR pszText, int cchMax);
    STDMETHODIMP GetViewFlags(DWORD *pdwFlags);
	/*/ICommDlgBrowser3
	STDMETHODIMP OnColumnClicked(IShellView *ppshv, int iColumn);
    STDMETHODIMP GetCurrentFilter(LPWSTR pszFileSpec, int cchFileSpec);
    STDMETHODIMP OnPreViewCreated(IShellView *ppshv);//*/
	/*/IFolderFilter
	STDMETHODIMP ShouldShow(IShellFolder *psf, PCIDLIST_ABSOLUTE pidlFolder, PCUITEMID_CHILD pidlItem);
	STDMETHODIMP GetEnumFlags(IShellFolder *psf, PCIDLIST_ABSOLUTE pidlFolder, HWND *phwnd, DWORD *pgrfFlags);//*/
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IShellFolderViewDual
    STDMETHODIMP get_Application(IDispatch **ppid);
    STDMETHODIMP get_Parent(IDispatch **ppid);
	STDMETHODIMP get_Folder(Folder **ppid);
	STDMETHODIMP SelectedItems(FolderItems **ppid);
	STDMETHODIMP get_FocusedItem(FolderItem **ppid);
	STDMETHODIMP SelectItem(VARIANT *pvfi, int dwFlags);
	STDMETHODIMP PopupItemMenu(FolderItem *pfi, VARIANT vx, VARIANT vy, BSTR *pbs);
	STDMETHODIMP get_Script(IDispatch **ppDisp);
    STDMETHODIMP get_ViewOptions(long *plViewOptions);
	//IExplorerBrowserEvents
	STDMETHODIMP OnNavigationPending(PCIDLIST_ABSOLUTE pidlFolder);
	STDMETHODIMP OnViewCreated(IShellView *psv);
	STDMETHODIMP OnNavigationComplete(PCIDLIST_ABSOLUTE pidlFolder);
	STDMETHODIMP OnNavigationFailed(PCIDLIST_ABSOLUTE pidlFolder);
	//IExplorerPaneVisibility
	STDMETHODIMP GetPaneState(REFEXPLORERPANE ep, EXPLORERPANESTATE *peps);
#if defined(_USE_SHELLBROWSER) || defined(_2000XP)
	//IShellFolder
	STDMETHODIMP ParseDisplayName(HWND hwnd, IBindCtx *pbc, LPWSTR pszDisplayName, ULONG *pchEaten, PIDLIST_RELATIVE *ppidl, ULONG *pdwAttributes);
	STDMETHODIMP EnumObjects(HWND hwndOwner, SHCONTF grfFlags, IEnumIDList **ppenumIDList);
	STDMETHODIMP BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppvOut);
	STDMETHODIMP BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppvOut);
	STDMETHODIMP CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2);
	STDMETHODIMP CreateViewObject(HWND hwndOwner, REFIID riid, void **ppv);
	STDMETHODIMP GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, SFGAOF *rgfInOut);
	STDMETHODIMP GetUIObjectOf(HWND hwndOwner, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid, UINT *rgfReserved, void **ppv);
	STDMETHODIMP GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET *pName);
	STDMETHODIMP SetNameOf(HWND hwndOwner, PCUITEMID_CHILD pidl, LPCWSTR pszName, SHGDNF uFlags, PITEMID_CHILD *ppidlOut);
	//IShellFolder2
	STDMETHODIMP GetDefaultSearchGUID(GUID *pguid);
	STDMETHODIMP EnumSearches(IEnumExtraSearch **ppEnum);
	STDMETHODIMP GetDefaultColumn(DWORD dwReserved, ULONG *pSort, ULONG *pDisplay);
	STDMETHODIMP GetDefaultColumnState(UINT iColumn, SHCOLSTATEF *pcsFlags);
	STDMETHODIMP GetDetailsEx(PCUITEMID_CHILD pidl, const SHCOLUMNID *pscid, VARIANT *pv);
	STDMETHODIMP GetDetailsOf(PCUITEMID_CHILD pidl, UINT iColumn, SHELLDETAILS *psd);
	STDMETHODIMP MapColumnToSCID(UINT iColumn, SHCOLUMNID *pscid);
#endif
	//IShellFolderViewCB
	STDMETHODIMP MessageSFVCB(UINT uMsg, WPARAM wParam, LPARAM lParam);
	//IDropTarget
	STDMETHODIMP DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
	STDMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
	STDMETHODIMP DragLeave();
	STDMETHODIMP Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
	//IPersist
    STDMETHODIMP GetClassID(CLSID *pClassID);
	//IPersistFolder
    STDMETHODIMP Initialize(PCIDLIST_ABSOLUTE pidl);
	//IPersistFolder2
    STDMETHODIMP GetCurFolder(PIDLIST_ABSOLUTE *ppidl);

	CteShellBrowser(CteTabCtrl *pTabs);
	~CteShellBrowser();

	void Init(CteTabCtrl *pTC, BOOL bNew);
	void Clear();
	void Show(BOOL bShow, DWORD dwOptions);
	VOID Suspend(int nMode);
	VOID SetPropEx();
	VOID ResetPropEx();
	int GetTabIndex();
	BOOL Close(BOOL bForce);
	VOID DestroyView(int nFlags);
	HWND GetListHandle(HWND *hList);
	HRESULT BrowseObject2(FolderItem *pid, UINT wFlags);
	BOOL Navigate1(FolderItem *pFolderItem, UINT wFlags, FolderItems *pFolderItems, FolderItem *pPrevious, int nErrorHandling);
	VOID Navigate1Ex(LPOLESTR pstr, FolderItems *pFolderItems, UINT wFlags, FolderItem *pPrevious, int nErrorHandleing);
	HRESULT Navigate2(FolderItem *pFolderItem, UINT wFlags, DWORD *param, FolderItems *pFolderItems, FolderItem *pPrevious, CteShellBrowser *pHistSB);
	HRESULT Navigate3(FolderItem *pFolderItem, UINT wFlags, DWORD *param, CteShellBrowser **ppSB, FolderItems *pFolderItems);
	HRESULT NavigateEB(DWORD dwFrame);
	HRESULT OnBeforeNavigate(FolderItem *pPrevious, UINT wFlags);
//	void InitializeMenuItem(HMENU hmenu, LPTSTR lpszItemName, int nId, HMENU hmenuSub);
	VOID GetSort(BSTR* pbs, int nFormat);
	VOID GetSort2(BSTR* pbs, int nFormat);
	VOID SetSort(BSTR bs);
	VOID GetGroupBy(BSTR* pbs);
	VOID SetGroupBy(BSTR bs);
	VOID SetRedraw(BOOL bRedraw);
	VOID SetColumnsStr(BSTR bsColumns);
	BSTR GetColumnsStr(int nFormat);
	VOID GetDefaultColumns();
	VOID SaveFocusedItemToHistory();
	VOID FocusItem(BOOL bFree);
	HRESULT GetAbsPath(FolderItem *pid, UINT wFlags, FolderItems *pFolderItems, FolderItem *pPrevious, CteShellBrowser *pHistSB);
	VOID EBNavigate();
	VOID SetHistory(FolderItems *pFolderItems, UINT wFlags);
	VOID GetVariantPath(FolderItem **ppFolderItem, FolderItems **ppFolderItems, VARIANT *pv);
	VOID Refresh(BOOL bCheck);
	BOOL SetActive(BOOL bForce);
	VOID SetTitle(BSTR szName, int nIndex);
	VOID GetShellFolderView();
	VOID GetFocusedIndex(int *piItem);
	VOID SetFolderFlags(BOOL bGetIconSize);
	VOID GetViewModeAndIconSize(BOOL bGetIconSize);
	HRESULT Items(UINT uItem, FolderItems **ppid);
	HRESULT SelectItemEx(LPITEMIDLIST *ppidl, int dwFlags, BOOL bFree);
	VOID InitFolderSize();
	VOID SetSize(LPCITEMIDLIST pidl, LPWSTR szText, int cch);
	VOID SetFolderSize(IShellFolder2 *pSF2, LPCITEMIDLIST pidl, LPWSTR szText, int cch);
	VOID SetLabel(LPCITEMIDLIST pidl, LPWSTR szText, int cch);
	BOOL ReplaceColumns(IDispatch *pdisp, NMLVDISPINFO *lpDispInfo);
	HRESULT PropertyKeyFromName(BSTR bs, PROPERTYKEY *pkey);
	FOLDERVIEWOPTIONS teGetFolderViewOptions(LPITEMIDLIST pidl, UINT uViewMode);
	VOID OnNavigationComplete2();
	HRESULT BrowseToObject();
	HRESULT GetShellFolder2(LPITEMIDLIST *ppidl);
	HRESULT CreateViewWindowEx(IShellView *pPreviousView);
	VOID SetTabName();
	HRESULT RemoveAll();
	VOID SetViewModeAndIconSize(BOOL bSetIconSize);
	VOID SetListColumnWidth();
	VOID AddItem(LPITEMIDLIST pidl);
	VOID NavigateComplete(BOOL bBeginNavigate);
	VOID InitFilter();
	HRESULT SetTheme();
	LPWSTR GetThemeName();
	VOID FixColumnEmphasis();
	HRESULT OnNavigationPending2(LPITEMIDLIST pidlFolder);
	HRESULT IncludeObject2(IShellFolder *pSF, LPCITEMIDLIST pidl);
	BOOL HasFilter();
	int GetSizeFormat();
#if defined(_USE_SHELLBROWSER) || defined(_2000XP)
	HRESULT NavigateSB(IShellView *pPreviousView, FolderItem *pPrevious);
#endif
#ifdef _2000XP
	VOID AddPathXP(CteFolderItems *pFolderItems, IShellFolderView *pSFV, int nIndex, BOOL bResultsFolder);
	int PSGetColumnIndexXP(LPWSTR pszName, int *pcxChar);
	BSTR PSGetNameXP(LPWSTR pszName, int nFormat);
	VOID AddColumnDataXP(LPWSTR pszColumns, LPWSTR pszName, int nWidth, int nFormat);
#endif
public:
	HWND		m_hwnd;
	HWND		m_hwndDV;
	HWND		m_hwndLV;
	HWND		m_hwndDT;
	HWND		m_hwndAlt;
	CteTabCtrl		*m_pTC;
	CteTreeView	*m_pTV;
	IShellView  *m_pShellView;
	IDispatch	*m_ppDispatch[Count_SBFunc];
	std::vector<IDispatch *> m_ppColumns;
	FolderItem *m_pFolderItem;
	FolderItem *m_pFolderItem1;
	IExplorerBrowser *m_pExplorerBrowser;
	LPITEMIDLIST m_pidl;
	IShellFolder2 *m_pSF2;
	std::vector<UINT> m_pDTColumns;
	BSTR		m_bsNextGroup;
	int			m_nForceViewMode;
	int			m_nFolderSizeIndex;
	int			m_nLabelIndex;
	int			m_nSizeIndex;
	int			m_nLinkTargetIndex;
	int			m_nSB;
	int			m_nUnload;
	int			m_nFocusItem;
	int			m_nSizeFormat;
	int			m_nGroupByDelay;
	LONG		m_dwUnavailable;
	DWORD		m_param[SB_Count];
	DWORD		m_nOpenedType;
	DWORD		m_dwCookie;
	DWORD		m_dwSessionId;
	DWORD		m_dwTickNotify;
	COLORREF	m_clrBk, m_clrTextBk, m_clrText;
	DISPID      m_dispidSortColumns;
	BOOL		m_bSetListColumnWidth;
	BOOL		m_bEmpty;
	BOOL		m_bInit;
	BOOL		m_bVisible;
	BOOL		m_bCheckLayout;
	BOOL		m_bRefreshLator;
	BOOL		m_bRefreshing;
	BOOL		m_bAutoVM;
	BOOL		m_bBeforeNavigate;
	BOOL		m_bRedraw;
	BOOL		m_bViewCreated;
	BOOL		m_bFiltered;
	BOOL		m_bNavigateComplete;
	BOOL		m_bSetRedraw;
#ifdef _2000XP
	TEColumn	*m_pColumns;
	UINT		m_nColumns;
#endif
private:
	VARIANT		m_vData;
	std::vector<FolderItem *> m_ppLog;
	IDispatch	*m_pdisp;
	std::vector< PROPERTYKEY> m_pDefultColumns;
	CteDropTarget2 *m_pDropTarget2;
	BSTR		m_bsFilter;
	BSTR		m_bsNextFilter;
	BSTR		m_bsAltSortColumn;

	LONG		m_cRef;
	LONG		m_nCreate;
	DWORD		m_dwEventCookie;
	UINT		m_uLogIndex;
	UINT		m_uPrevLogIndex;
	int			m_nSuspendMode;
	BOOL		m_bIconSize;
	BOOL		m_bRegenerateItems;
	IShellFolderViewCB	*m_pSFVCB;
#ifdef _2000XP
	int			m_nFolderName;
#endif
};

class CteMemory : public IDispatchEx
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IDispatchEx
	STDMETHODIMP GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid);
	STDMETHODIMP InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller);
	STDMETHODIMP DeleteMemberByName(BSTR bstrName, DWORD grfdex);
	STDMETHODIMP DeleteMemberByDispID(DISPID id);
	STDMETHODIMP GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex);
	STDMETHODIMP GetMemberName(DISPID id, BSTR *pbstrName);
	STDMETHODIMP GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid);
	STDMETHODIMP GetNameSpaceParent(IUnknown **ppunk);

	VOID ItemEx(int i, VARIANT *pVarResult, VARIANT *pVarNew);
	BSTR AddBSTR(BSTR bs);
	VOID Read(int nIndex, int nLen, VARIANT *pVarResult);
	VOID Write(int nIndex, int nLen, VARTYPE vt, VARIANT *pv);
	void SetPoint(int x, int y);
	void Free(BOOL bpbs);

	CteMemory(int nSize, void *pc, int nCount, LPWSTR lpStruct);
	~CteMemory();
public:
	char	*m_pc;

	int		m_nSize;
	int		m_nCount;
private:
	std::vector<BSTR> m_ppbs;
	BSTR	m_bsStruct;
	BSTR	m_bsAlloc;

	LONG	m_cRef;
	int		m_nStructIndex;
};

class CteContextMenu : public IDispatch//, IContextMenu
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
/*	//IContexMenun
	STDMETHODIMP QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
	STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax);
	STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
//*/
	CteContextMenu(IUnknown *punk, IDataObject *pDataObj, IUnknown *pSB);
	~CteContextMenu();

	BOOL GetFolderVew(IShellBrowser **ppSB);
public:
	IContextMenu *m_pContextMenu;
	HMODULE m_hDll;
private:
	IDataObject *m_pDataObj;

	teParam	m_param[5];

	LONG	m_cRef;
};

class CteDropTarget : public IDispatch
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	CteDropTarget(IDropTarget *pDropTarget, FolderItem *pFolderItem);
	~CteDropTarget();
private:
	IDropTarget *m_pDropTarget;
	FolderItem *m_pFolderItem;

	LONG	m_cRef;
};

class CteTreeView : public IDispatch,
#ifdef _2000XP
	public IOleClientSite, public IOleInPlaceSite,
#endif
	public INameSpaceTreeControlEvents, public INameSpaceTreeControlCustomDraw, public IShellItemFilter
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//INameSpaceTreeControlEvents
	STDMETHODIMP OnItemClick(IShellItem *psi, NSTCEHITTEST nstceHitTest, NSTCSTYLE nsctsFlags);
	STDMETHODIMP OnPropertyItemCommit(IShellItem *psi);
	STDMETHODIMP OnItemStateChanging(IShellItem *psi, NSTCITEMSTATE nstcisMask, NSTCITEMSTATE nstcisState);
	STDMETHODIMP OnItemStateChanged(IShellItem *psi, NSTCITEMSTATE nstcisMask, NSTCITEMSTATE nstcisState);
	STDMETHODIMP OnSelectionChanged(IShellItemArray *psiaSelection);
	STDMETHODIMP OnKeyboardInput(UINT uMsg, WPARAM wParam, LPARAM lParam);
	STDMETHODIMP OnBeforeExpand(IShellItem *psi);
	STDMETHODIMP OnAfterExpand(IShellItem *psi);
	STDMETHODIMP OnBeginLabelEdit(IShellItem *psi);
	STDMETHODIMP OnEndLabelEdit(IShellItem *psi);
	STDMETHODIMP OnGetToolTip(IShellItem *psi, LPWSTR pszTip, int cchTip);
	STDMETHODIMP OnBeforeItemDelete(IShellItem *psi);
	STDMETHODIMP OnItemAdded(IShellItem *psi, BOOL fIsRoot);
	STDMETHODIMP OnItemDeleted(IShellItem *psi, BOOL fIsRoot);
	STDMETHODIMP OnBeforeContextMenu(IShellItem *psi, REFIID riid, void **ppv);
	STDMETHODIMP OnAfterContextMenu(IShellItem *psi, IContextMenu *pcmIn, REFIID riid, void **ppv);
	STDMETHODIMP OnBeforeStateImageChange(IShellItem *psi);
	STDMETHODIMP OnGetDefaultIconIndex(IShellItem *psi, int *piDefaultIcon, int *piOpenIcon);
	//INameSpaceTreeControlCustomDraw
	STDMETHODIMP PrePaint(HDC hdc, RECT *prc, LRESULT *plres);
	STDMETHODIMP PostPaint(HDC hdc, RECT *prc);
	STDMETHODIMP ItemPrePaint(HDC hdc, RECT *prc, NSTCCUSTOMDRAW *pnstccdItem, COLORREF *pclrText, COLORREF *pclrTextBk, LRESULT *plres);
	STDMETHODIMP ItemPostPaint(HDC hdc, RECT *prc, NSTCCUSTOMDRAW *pnstccdItem);
	//IShellItemFilter
	STDMETHODIMP IncludeItem(IShellItem *psi);
	STDMETHODIMP GetEnumFlagsForItem(IShellItem *psi, SHCONTF *pgrfFlags);
#ifdef _2000XP
	//IOleClientSite
	STDMETHODIMP SaveObject();
	STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk);
	STDMETHODIMP GetContainer(IOleContainer **ppContainer);
	STDMETHODIMP ShowObject();
	STDMETHODIMP OnShowWindow(BOOL fShow);
	STDMETHODIMP RequestNewObjectLayout();
	//IOleWindow
	STDMETHODIMP GetWindow(HWND *phwnd);
	STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
	//IOleInPlaceSite
	STDMETHODIMP CanInPlaceActivate();
	STDMETHODIMP OnInPlaceActivate();
	STDMETHODIMP OnUIActivate();
	STDMETHODIMP GetWindowContext(IOleInPlaceFrame **ppFrame, IOleInPlaceUIWindow **ppDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo);
	STDMETHODIMP Scroll(SIZE scrollExtant);
	STDMETHODIMP OnUIDeactivate(BOOL fUndoable);
	STDMETHODIMP OnInPlaceDeactivate();
	STDMETHODIMP DiscardUndoState();
	STDMETHODIMP DeactivateAndUndo();
	STDMETHODIMP OnPosRectChange(LPCRECT lprcPosRect);
#endif
	//IPersist
    STDMETHODIMP GetClassID(CLSID *pClassID);
	//IPersistFolder
    STDMETHODIMP Initialize(PCIDLIST_ABSOLUTE pidl);
	//IPersistFolder2
    STDMETHODIMP GetCurFolder(PIDLIST_ABSOLUTE *ppidl);

	CteTreeView();
	~CteTreeView();

	VOID Init(CteShellBrowser *pFV, HWND hwnd);
	VOID Close();
	VOID Create();
	HRESULT getSelected(IDispatch **pItem);
	VOID SetRootV(VARIANT *pv);
	VOID SetRoot();
	VOID Show();
	VOID Expand();
#ifdef _2000XP
	VOID SetObjectRect();
#endif
public:
	BSTR		m_bsRoot;
	HWND        m_hwnd;
	HWND        m_hwndTV;
	INameSpaceTreeControl	*m_pNameSpaceTreeControl;
	CteShellBrowser	*m_pFV;
	IShellItem *m_psiFocus;
	DWORD	*m_param;
	DWORD	m_dwState;
	BOOL	m_bEmpty;
	BOOL	m_bExpand;
#ifdef _2000XP
	IShellNameSpace *m_pShellNameSpace;
#endif
	BOOL		m_bMain;
	BOOL		m_bSetRoot;
private:
	VARIANT m_vData;
	LPWSTR	lplpVerbs;
	CteDropTarget2 *m_pDropTarget2;
	HWND	m_hwndParent;
#ifdef _W2000
	VARIANT m_vSelected;
#endif
	LONG	m_cRef;
	int		m_nType, m_nOpenedType;
	DWORD	m_dwCookie;
};

class CteCommonDialog : public IDispatch
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	CteCommonDialog();
	~CteCommonDialog();
private:
	OPENFILENAME m_ofn;

	LONG	m_cRef;
};

class CteWICBitmap : public IDispatch
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	CteWICBitmap();
	~CteWICBitmap();
	VOID FromStreamRelease(IStream *pStream, LPWSTR lpfn, BOOL bExtend, int cx);
	HRESULT GetArchive(LPWSTR lpfn, int cx);
	VOID GetFrameFromStream(IStream *pStream, UINT uFrame, BOOL bInit);
	BOOL HasImage();
	CteWICBitmap* GetBitmapObj();
	VOID ClearImage(BOOL bAll);
	HBITMAP GetHBITMAP(COLORREF clBk);
	BOOL Get(WICPixelFormatGUID guidNewPF);
	BOOL GetEncoderClsid(LPWSTR pszName, CLSID* pClsid, LPWSTR pszMimeType);
	HRESULT CreateBitmapFromHBITMAP(HBITMAP hBitmap, HPALETTE hPalette, int nAlpha);
	HRESULT CreateStream(IStream *pStream, CLSID encoderClsid, LONG lQuality);
//	HRESULT CreateBMPStream(IStream *pStream, LPWSTR szMime);
private:
	IWICBitmap *m_pImage;
	IWICImagingFactory *m_pWICFactory;
	IStream *m_pStream;
	CLSID m_guidSrc;
	IWICMetadataQueryReader *m_ppMetadataQueryReader[2];
	IDispatch	*m_ppDispatch[Count_WICFunc];
	LONG	m_cRef;
	UINT m_uFrameCount, m_uFrame;
};

#ifdef _USE_BSEARCHAPI
class CteWindowsAPI : public IDispatch
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	CteWindowsAPI(TEDispatchApi *pApi);
	~CteWindowsAPI();
private:
	TEDispatchApi *m_pApi;
	LONG	m_cRef;
};
#endif
#ifdef _USE_OBJECTAPI
class CteAPI : public IDispatch
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	CteAPI(TEDispatchApi *pApi);
	~CteAPI();
private:
	TEDispatchApi *m_pApi;
	LONG		m_cRef;
};
#endif

class CteDispatch : public IDispatch, public IEnumVARIANT
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IEnumVARIANT
	STDMETHODIMP Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched);
	STDMETHODIMP Skip(ULONG celt);
	STDMETHODIMP Reset(void);
	STDMETHODIMP Clone(IEnumVARIANT **ppEnum);

	CteDispatch(IDispatch *pDispatch, int nMode, DISPID dispId);
	~CteDispatch();

	VOID Clear();
public:
	IActiveScript *m_pActiveScript;
	HMODULE m_hDll;
	DISPID		m_dispIdMember;
	int			m_nIndex;
private:
	IDispatch	*m_pDispatch, *m_pDispatch2;

	LONG		m_cRef;
	int			m_nMode;//0: Clone 1:Collection 2:IActiveScript 4:DLL
};

class CteDispatchEx : public IDispatchEx
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IDispatchEx
	STDMETHODIMP GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid);
	STDMETHODIMP InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller);
	STDMETHODIMP DeleteMemberByName(BSTR bstrName, DWORD grfdex);
	STDMETHODIMP DeleteMemberByDispID(DISPID id);
	STDMETHODIMP GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex);
	STDMETHODIMP GetMemberName(DISPID id, BSTR *pbstrName);
	STDMETHODIMP GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid);
	STDMETHODIMP GetNameSpaceParent(IUnknown **ppunk);
	//
	VOID GetLegacyDispId(BSTR bstrName, DISPID *pid, BOOL bDispatchEx, HRESULT *phr);
	VOID InvokeLegacy(DISPID id, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, HRESULT *phr);

	CteDispatchEx(IUnknown *punk, BOOL bLegacy);
	~CteDispatchEx();
private:
	IDispatchEx *m_pdex;
	LONG	m_cRef;
	BOOL	m_bDispathEx;
	BOOL	m_bLegacy;
};

class CteActiveScriptSite : public IActiveScriptSite, public IActiveScriptSiteWindow
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	CteActiveScriptSite(IUnknown *punk, EXCEPINFO *pExcepInfo, HRESULT *phr);
	~CteActiveScriptSite();
	//ActiveScriptSite
	STDMETHODIMP GetLCID(LCID *plcid);
	STDMETHODIMP GetItemInfo(LPCOLESTR pstrName,
                           DWORD dwReturnMask,
                           IUnknown **ppiunkItem,
                           ITypeInfo **ppti);
	STDMETHODIMP GetDocVersionString(BSTR *pbstrVersion);
	STDMETHODIMP OnScriptError(IActiveScriptError *pscripterror);
	STDMETHODIMP OnStateChange(SCRIPTSTATE ssScriptState);
	STDMETHODIMP OnScriptTerminate(const VARIANT *pvarResult,const EXCEPINFO *pexcepinfo);
	STDMETHODIMP OnEnterScript(void);
	STDMETHODIMP OnLeaveScript(void);
	//IActiveScriptSiteWindow
	STDMETHODIMP GetWindow(HWND *phwnd);
	STDMETHODIMP EnableModeless(BOOL fEnable);

public:
	IDispatchEx	*m_pDispatchEx;
	EXCEPINFO *m_pExcepInfo;
	HRESULT	*m_phr;
	LONG		m_cRef;
};

#ifdef _USE_HTMLDOC
// DocHostUIHandler
class CteDocHostUIHandler : public IDocHostUIHandler
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDocHostUIHandler
	STDMETHODIMP ShowContextMenu(DWORD dwID, POINT *ppt, IUnknown *pcmdtReserved, IDispatch *pdispReserved);
	STDMETHODIMP GetHostInfo(DOCHOSTUIINFO *pInfo);
	STDMETHODIMP ShowUI(DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc);
	STDMETHODIMP HideUI(VOID);
	STDMETHODIMP UpdateUI(VOID);
	STDMETHODIMP EnableModeless(BOOL fEnable);
	STDMETHODIMP OnDocWindowActivate(BOOL fActivate);
	STDMETHODIMP OnFrameWindowActivate(BOOL fActivate);
	STDMETHODIMP ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fFrameWindow);
	STDMETHODIMP TranslateAccelerator(LPMSG lpMsg, const GUID *pguidCmdGroup, DWORD nCmdID);
	STDMETHODIMP GetOptionKeyPath(LPOLESTR *pchKey, DWORD dw);
	STDMETHODIMP GetDropTarget(IDropTarget *pDropTarget, IDropTarget **ppDropTarget);
	STDMETHODIMP GetExternal(IDispatch **ppDispatch);
	STDMETHODIMP TranslateUrl(DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut);
	STDMETHODIMP FilterDataObject(IDataObject *pDO, IDataObject **ppDORet);

	CteDocHostUIHandler();
	~CteDocHostUIHandler();
public:
	LONG	m_cRef;
};
#endif
#ifdef _USE_TESTOBJECT
class CteTest : public IDispatch
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	CteTest();
	~CteTest();
public:
	LONG	m_cRef;
};
#endif
class CteProgressDialog : public IDispatch
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	CteProgressDialog(IProgressDialog *ppd);
	~CteProgressDialog();
private:
	IProgressDialog *m_ppd;
	LONG	m_cRef;
};

class CteObject : public IDispatch
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	CteObject(PVOID pObj);
	~CteObject();

	VOID Clear();
public:
	LONG	m_cRef;
	IUnknown *m_punk;
	LARGE_INTEGER m_liOffset;
};

class CteFileSystemBindData : public IFileSystemBindData
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IFileSystemBindData
	STDMETHODIMP SetFindData(const WIN32_FIND_DATAW *pfd);
	STDMETHODIMP GetFindData(WIN32_FIND_DATAW *pfd);

	CteFileSystemBindData();
	~CteFileSystemBindData();
public:
	WIN32_FIND_DATA m_wfd;
	LONG	m_cRef;
};

class CteEnumerator : public IDispatch
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	VOID GetItem();

	CteEnumerator(VARIANT *pvObject);
	~CteEnumerator();
public:
	VARIANT	m_vItem;
	IEnumVARIANT *m_pEnumVARIANT;
	IDispatchEx *m_pdex;
	LONG	m_cRef;
	HRESULT m_hr;
	DISPID m_dispIdMember;
};
