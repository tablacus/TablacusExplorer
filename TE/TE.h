#pragma once

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

class CteShellBrowser : public IShellBrowser, public ICommDlgBrowser2,
	public IShellFolderViewDual,
//	public IFolderFilter,
	public IExplorerBrowserEvents, public IExplorerPaneVisibility,
	public IShellFolderViewCB,
#ifdef _2000XP
	public IShellFolder2,
#else
#ifdef USE_SHELLBROWSER
	public IShellFolder,
#endif
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
#if defined(USE_SHELLBROWSER) || defined(_2000XP)
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
	HRESULT Navigate2(FolderItem *pFolderItem, UINT wFlags, DWORD *param, FolderItems *pFolderItems, CteFolderItem *pPrevious, CteShellBrowser *pHistSB);
	HRESULT Navigate3(FolderItem *pFolderItem, UINT wFlags, DWORD *param, CteShellBrowser **ppSB, FolderItems *pFolderItems);
	VOID GetInitFS(FOLDERSETTINGS *pfs);
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
	VOID GetFocusedItem(LPITEMIDLIST *ppidl);
	VOID SaveFocusedItemToHistory();
	VOID FocusItem();
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
	HRESULT SelectItemEx(LPITEMIDLIST pidl, int dwFlags);
	VOID InitFolderSize();
	HRESULT GetPropertyKey(HWND hHeader, int iItem, PROPERTYKEY *pPropKey);
	BOOL SetColumnsProperties(HWND hHeader, int iItem);
	VOID SetSize(LPCITEMIDLIST pidl, LPWSTR szText, int cch);
	VOID SetFolderSize(IShellFolder2 *pSF2, LPCITEMIDLIST pidl, LPWSTR szText, int cch);
	BOOL SetFolderSizeProperties();
	VOID SetLabel(LPCITEMIDLIST pidl, LPWSTR szText, int cch);
	VOID SetLabelProperties();
	BOOL ReplaceColumns(IDispatch *pdisp, LVITEM *plvi, LPITEMIDLIST pidl);
	VOID ReplaceColumnsEx(IDispatch *pdisp, LVITEM *plvi, LPITEMIDLIST *ppidl);
#ifdef USE_VS_COLUMNSREPLACE
	VOID SetReplaceColumnsProperties(IDispatch *pdisp, PROPERTYKEY propKey);
#endif
	HRESULT PropertyKeyFromName(BSTR bs, PROPERTYKEY *pkey);
	FOLDERVIEWOPTIONS teGetFolderViewOptions(LPITEMIDLIST pidl, UINT uViewMode);
	VOID OnNavigationComplete2();
	HRESULT BrowseToObject(DWORD dwFrame, FOLDERVIEWOPTIONS fvoFlags);
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
	HRESULT IncludeObject2(IShellFolder *pSF, LPCITEMIDLIST pidl, LPITEMIDLIST pidlFull);
	BOOL HasFilter();
	int GetSizeFormat();
	VOID SetObjectRect();
	int GetFolderViewAndItemCount(IFolderView **ppFV, UINT uFlags);
	HRESULT Search(LPWSTR pszSearch);
	VOID RedrawUpdate();
	VOID SetEmptyText();
#if defined(USE_SHELLBROWSER) || defined(_2000XP)
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
	CteFolderItem *m_pFolderItem;
	CteFolderItem *m_pFolderItem1;
	IExplorerBrowser *m_pExplorerBrowser;
	LPITEMIDLIST m_pidl;
	IShellFolder2 *m_pSF2;
	std::vector<UINT> m_pDTColumns;
	BSTR		m_bsNextGroup;
	RECT		m_rc;
	int			m_nForceViewMode;
	int			m_nForceIconSize;
	int			m_nFolderSizeIndex;
	int			m_nLabelIndex;
	int			m_nSizeIndex;
	int			m_nLinkTargetIndex;
	int			m_nSB;
	int			m_nUnload;
	int			m_nFocusItem;
	int			m_nSizeFormat;
	int			m_nGroupByDelay;
	DWORD		m_param[SB_Count];
	DWORD		m_nOpenedType;
	DWORD		m_dwCookie;
	DWORD		m_dwTickNotify;
	DWORD		m_dwRedraw;
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
	BOOL		m_bViewCreated;
	BOOL		m_bFiltered;
	BOOL		m_bNavigateComplete;
#ifdef _2000XP
	TEColumn	*m_pColumns;
	UINT		m_nColumns;
#endif
private:
	VARIANT		m_vData;
	std::vector<CteFolderItem *> m_ppLog;
	IDispatch	*m_pdisp;
	std::vector< PROPERTYKEY> m_pDefultColumns;
	CteDropTarget2 *m_pDropTarget2;
	BSTR		m_bsFilter;
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

class CteActiveScriptSite : public IActiveScriptSite, public IActiveScriptSiteWindow
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	CteActiveScriptSite(IUnknown *punk, EXCEPINFO *pExcepInfo);
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
	LONG		m_cRef;
	HRESULT m_hr;
};

