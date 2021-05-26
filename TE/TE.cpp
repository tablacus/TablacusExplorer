// TE.cpp
// Tablacus Explorer (C)2011 Gaku
// MIT Lisence
// Visual Studio Express 2017 for Windows Desktop
// Windows SDK v7.1
// https://tablacus.github.io/

#include "stdafx.h"
#include "common.h"
#include "idsofnames.h"
#include "TE.h"

// Global Variables:
HINSTANCE hInst;								// current instance
WCHAR	g_szTE[MAX_LOADSTRING];
WCHAR	g_szText[MAX_STATUS];
WCHAR	g_szStatus[MAX_STATUS];
BSTR	g_bsTitle = NULL;
BSTR	g_bsName = NULL;
BSTR	g_bsAnyText = NULL;
BSTR	g_bsAnyTextId = NULL;
HWND	g_hwndMain = NULL;
CteTabCtrl *g_pTC = NULL;
std::vector<CteTabCtrl *> g_ppTC;
HINSTANCE	g_hShell32 = NULL;
HINSTANCE	g_hTEWV = NULL;
HWND		g_hDialog = NULL;
IShellWindows *g_pSW = NULL;

LPFNSHRunDialog _SHRunDialog = NULL;
LPFNRegenerateUserEnvironment _RegenerateUserEnvironment = NULL;
LPFNChangeWindowMessageFilterEx _ChangeWindowMessageFilterEx = NULL;
LPFNRtlGetVersion _RtlGetVersion = NULL;
LPFNSetDefaultDllDirectories _SetDefaultDllDirectories = NULL;
LPFNSetPreferredAppMode _SetPreferredAppMode = NULL;
LPFNAllowDarkModeForWindow _AllowDarkModeForWindow = NULL;
LPFNSetWindowCompositionAttribute _SetWindowCompositionAttribute = NULL;
LPFNDwmSetWindowAttribute _DwmSetWindowAttribute = NULL;
LPFNShouldAppsUseDarkMode _ShouldAppsUseDarkMode = NULL;
LPFNRefreshImmersiveColorPolicyState _RefreshImmersiveColorPolicyState = NULL;
//LPFNIsDarkModeAllowedForWindow _IsDarkModeAllowedForWindow = NULL;
LPFNGetDpiForMonitor _GetDpiForMonitor = NULL;
#ifdef _2000XP
LPFNSetDllDirectoryW _SetDllDirectoryW = NULL;
LPFNIsWow64Process _IsWow64Process = NULL;
LPFNPSPropertyKeyFromString _PSPropertyKeyFromString = NULL;
LPFNPSGetPropertyKeyFromName _PSGetPropertyKeyFromName = NULL;
LPFNPSPropertyKeyFromString _PSPropertyKeyFromStringEx = NULL;
LPFNPSGetPropertyDescription _PSGetPropertyDescription = NULL;
LPFNPSStringFromPropertyKey _PSStringFromPropertyKey = NULL;
LPFNPropVariantToVariant _PropVariantToVariant = NULL;
LPFNVariantToPropVariant _VariantToPropVariant = NULL;
LPFNSHCreateItemFromIDList _SHCreateItemFromIDList = NULL;
LPFNSHGetIDListFromObject _SHGetIDListFromObject = NULL;
LPFNChangeWindowMessageFilter _ChangeWindowMessageFilter = NULL;
LPFNAddClipboardFormatListener _AddClipboardFormatListener = NULL;
LPFNRemoveClipboardFormatListener _RemoveClipboardFormatListener = NULL;
LPFNSHCreateShellItemArrayFromShellItem _SHCreateShellItemArrayFromShellItem = NULL;
#else
#define _PSPropertyKeyFromString PSPropertyKeyFromString
#define _PSGetPropertyKeyFromName PSGetPropertyKeyFromName
#define _PSGetPropertyDescription PSGetPropertyDescription
#define _PSPropertyKeyFromStringEx tePSPropertyKeyFromStringEx
#define _SHCreateItemFromIDList SHCreateItemFromIDList
#define _SHGetIDListFromObject SHGetIDListFromObject
#define _PropVariantToVariant PropVariantToVariant
#define _VariantToPropVariant VariantToPropVariant
#define _SHCreateShellItemArrayFromShellItem SHCreateShellItemArrayFromShellItem
#endif
#ifdef _USE_APIHOOK
LPFNRegQueryValueExW _RegQueryValueExW = NULL;
LPFNRegQueryValueW _RegQueryValueW = NULL;
LPFNGetSysColor _GetSysColor = NULL;
LPFNOpenNcThemeData _OpenNcThemeData = NULL;
#endif
#define teVariantCopy(d, s)     if (d) { VariantCopy(d, s); }

FORMATETC HDROPFormat = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC IDLISTFormat = {0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC DROPEFFECTFormat = {0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC UNICODEFormat = {CF_UNICODETEXT, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC TEXTFormat = {CF_TEXT, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC DRAGWINDOWFormat = {0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC IsShowingLayeredFormat = {0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
GUID		g_ClsIdStruct;
GUID		g_ClsIdSB;
GUID		g_ClsIdTC;
GUID		g_ClsIdFI;
IUIAutomation *g_pAutomation = NULL;
PROPERTYID	g_PID_ItemIndex;
EXCEPINFO   g_ExcepInfo;

CTE			*g_pTE;
std::vector<CteShellBrowser *> g_ppSB;
std::vector<CteTreeView *> g_ppTV;
std::vector<LONG_PTR> g_ppGetArchive;
std::vector<LONG_PTR> g_ppGetImage;
std::vector<LONG_PTR> g_ppMessageSFVCB;
CteWebBrowser *g_pWebBrowser = NULL;
#ifdef _USE_OBJECTAPI
IDispatch	*g_pAPI = NULL;
#endif

LPITEMIDLIST g_pidls[MAX_CSIDL2];
BSTR		 g_bsPidls[MAX_CSIDL2];

IDispatch	*g_pOnFunc[Count_OnFunc];
std::vector<HMODULE>	g_pFreeLibrary;
std::vector<HMODULE>	g_phModule;
IDispatch	*g_pJS = NULL;
std::vector<TEWebBrowsers>	g_pSubWindows;
std::vector<TEFS *>	g_pFolderSize;
CRITICAL_SECTION g_csFolderSize;
IDropTargetHelper *g_pDropTargetHelper = NULL;
IUnknown	*g_pDraggingCtrl = NULL;
IDataObject	*g_pDraggingItems = NULL;
IDispatchEx *g_pCM = NULL;
IServiceProvider *g_pSP = NULL;
ULONG_PTR g_Token;
HHOOK	g_hHook;
HHOOK	g_hMouseHook;
HHOOK	g_hMessageHook;
HHOOK	g_hMenuKeyHook;
HMENU	g_hMenu = NULL;
BSTR	g_bsCmdLine = NULL;
BSTR	g_bsDocumentWrite = NULL;
BSTR	g_bsClipRoot = NULL;
BSTR	g_bsDateTimeFormat = NULL;
BSTR	g_bsHiddenFilter = NULL;
BSTR	g_bsExplorerBrowserFilter = NULL;
HTREEITEM	g_hItemDown = NULL;
SORTCOLUMN g_pSortColumnNull[3];
VARIANT g_vData;
VARIANT g_vArguments;
CteDropTarget2 *g_pDropTarget2;

UINT	g_uCrcTable[256];
LONG	g_nLockUpdate = 0;
LONG    g_nCountOfThreadFolderSize = 0;
DWORD	g_paramFV[SB_Count];
DWORD	g_dwMainThreadId;
DWORD	g_dwCookieSW = 0;
DWORD   g_dwCookieJS = 0;
DWORD	g_dwSessionId = 0;
DWORD	g_dwTickKey;
DWORD	g_dwFreeLibrary = 0;
DWORD	g_dwTickMount = 0;
long	g_nProcKey	   = 0;
long	g_nProcMouse   = 0;
long	g_nProcDrag    = 0;
long	g_nProcFV      = 0;
long	g_nProcTV      = 0;
long	g_nProcFI      = 0;
long	g_nThreads	   = 0;
long	g_lSWCookie = 0;
int		g_nReload = 0;
int		g_nException = 256;
int		*g_maps[MAP_LENGTH];
int		g_param[Count_TE_params];
int		g_x = MAXINT;
int		g_nTCCount = 0;
int		g_nTCIndex = 0;
int		g_nWindowTheme = 0;
int		g_nBlink = 0;
int		g_nCreateTimer = 100;
int		g_nDropState = 0;
BOOL	g_bArrange = FALSE;
BOOL	g_bLabelsMode;
BOOL	g_bMessageLoop = TRUE;
BOOL	g_bDialogOk = FALSE;
BOOL	g_bInit = FALSE;
BOOL	g_bShowParseError = TRUE;
BOOL	g_bDragging = FALSE;
BOOL	g_bCanLayout = FALSE;
BOOL	g_bUpper10;
BOOL	g_bDarkMode = FALSE;
BOOL	g_bDragIcon = TRUE;
#ifdef _2000XP
std::vector <IUnknown *> g_pRelease;
int		g_nCharWidth = 7;
BOOL	g_bCharWidth = FALSE;
BOOL	g_bUpperVista;
HWND	g_hwndNextClip = NULL;
LPWSTR  g_szTotalFileSizeXP = GetUserDefaultLCID() == 1041 ? L"総ファイル サイズ" : L"Total file size";
LPWSTR  g_szLabelXP = GetUserDefaultLCID() == 1041 ? L"ラベル" : L"Label";
LPWSTR	g_szTotalFileSizeCodeXP = L"System.TotalFileSize";
LPWSTR	g_szLabelCodeXP = L"System.Contact.Label";
#endif
#ifdef _W2000
BOOL	g_bIsXP;
BOOL	g_bIs2000;
#endif
#ifdef _DEBUG
LPWSTR	g_strException;
WCHAR	g_pszException[MAX_PATH];
HHOOK	g_hMenuGMHook;
#endif
#ifdef _CHECK_HANDLELEAK
int g_nLeak = 0;
#endif

// Forward declarations of functions included in this code module:
ATOM MyRegisterClass(HINSTANCE hInstance, LPWSTR szClassName, WNDPROC lpfnWndProc);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
VOID ArrangeWindow();
BOOL teGetIDListFromObject(IUnknown *punk, LPITEMIDLIST *ppidl);
VOID CALLBACK teTimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
VOID CALLBACK teTimerProcParse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
VOID CALLBACK teTimerProcForTree(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
static void threadParseDisplayName(void *args);
LPITEMIDLIST teILCreateFromPath(LPWSTR pszPath);
int teGetModuleFileName(HMODULE hModule, BSTR *pbsPath);
BOOL GetVarArrayFromIDList(VARIANT *pv, LPITEMIDLIST pidl);
HRESULT teGetDisplayNameFromIDList(BSTR *pbs, LPITEMIDLIST pidl, SHGDNF uFlags);
HRESULT STDAPICALLTYPE tePSPropertyKeyFromStringEx(__in LPCWSTR pszString, __out PROPERTYKEY *pkey);
VOID teGetDarkMode();
IDispatch* teCreateObj(LONG lId, VARIANT *pvArg);
VOID teArrayPush(IDispatch *pdisp, PVOID pObj);
HRESULT teExecMethod(IDispatch *pdisp, LPOLESTR sz, VARIANT *pvResult, int nArg, VARIANTARG *pvArgs);
VOID Invoke4(IDispatch *pdisp, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs);
int teGetObjectLength(IDispatch *pdisp);
BOOL teILIsSearchFolder(LPCITEMIDLIST pidl);

//Unit

DWORD teGetDWordFromDataObj(IDataObject *pDataObj, FORMATETC *pformatetcIn, DWORD dw)
{
	if (pDataObj) {
		STGMEDIUM Medium;
		if (pDataObj->GetData(pformatetcIn, &Medium) == S_OK) {
			if (::GlobalSize(Medium.hGlobal) >= sizeof(DWORD)) {
				dw = *(DWORD *)::GlobalLock(Medium.hGlobal);
				::GlobalUnlock(Medium.hGlobal);
			}
			::ReleaseStgMedium(&Medium);
		}
	}
	return dw;
}

VOID teAddRemoveProc(std::vector<LONG_PTR> *pppProc, LONG_PTR lpProc, BOOL bAdd)
{
	if (bAdd) {
		pppProc->push_back(lpProc);
	} else {//Remove
		for (size_t u = pppProc->size(); u--;) {
			if (pppProc->at(u) == lpProc) {
				pppProc->erase(pppProc->begin() + u);
			}
		}
	}
}

VOID teSetExStyleOr(HWND hwnd, LONG l) {
	LONG exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
	if (exStyle != (exStyle | l)) {
		SetWindowLong(hwnd, GWL_EXSTYLE, exStyle | l);
	}
}

VOID teSetExStyleAnd(HWND hwnd, LONG l) {
	LONG exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
	if (exStyle != (exStyle & l)) {
		SetWindowLongPtr(hwnd, GWL_EXSTYLE, exStyle & l);
	}
}

VOID teFixListState(HWND hwnd, int dwFlags) {
	if (hwnd && (dwFlags & SVSI_FOCUSED)) {// bug fix
#ifdef _2000XP
		if (g_bUpperVista) {
#endif
			int nFocused = ListView_GetNextItem(hwnd, -1, LVNI_ALL | LVNI_FOCUSED);
			if (nFocused >= 0) {
				ListView_SetItemState(hwnd, nFocused, 0, LVIS_FOCUSED);
			}
#ifdef _2000XP
		}
#endif
	}
}

IDispatch* teAddLegacySupport(IDispatch *pdisp)
{
	DISPID id;
	LPWSTR sz = L"Item";
	if SUCCEEDED(pdisp->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &id)) {
		if (id != DISPID_UNKNOWN) {
			return pdisp;
		}
	}
	IDispatch *pArray = new CteDispatchEx(pdisp, TRUE);
	pdisp->Release();
	return pArray;
}

BOOL teIsDarkColor(COLORREF cl)
{
	return 299 * GetRValue(cl) + 587 * GetGValue(cl) + 114 * GetBValue(cl) < 127500;
}

VOID teSetTreeTheme(HWND hwnd, COLORREF cl)
{
	if (_AllowDarkModeForWindow) {
		BOOL bDarkMode = teIsDarkColor(cl);
		_AllowDarkModeForWindow(hwnd, bDarkMode);
		SetWindowTheme(hwnd, bDarkMode ? L"darkmode_explorer" : L"explorer", NULL);
	}
}

HRESULT teDoDragDrop(HWND hwnd, IDataObject *pDataObj, DWORD *pdwEffect, BOOL bDropState)
{
	HRESULT hr;
	try {
		g_nDropState = bDropState ? 2 : 1;
		IDragSourceHelper *pDragSourceHelper;
		if SUCCEEDED(g_pDropTargetHelper->QueryInterface(IID_PPV_ARGS(&pDragSourceHelper))) {
			pDragSourceHelper->InitializeFromWindow(hwnd, NULL, pDataObj);
			pDragSourceHelper->Release();
		}
		g_pDraggingItems = pDataObj;
		hr = DoDragDrop(pDataObj, static_cast<IDropSource *>(g_pTE), *pdwEffect, pdwEffect);
	} catch(...) {
		hr = E_UNEXPECTED;
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"DoDragDrop";
#endif
	}
	g_pDraggingItems = NULL;
	return hr;
}

BOOL teGetSubWindow(HWND hwnd, CteWebBrowser **ppWB)
{
	for (size_t u = g_pSubWindows.size(); u--;) {
		if (g_pSubWindows[u].hwnd == hwnd) {
			*ppWB = g_pSubWindows[u].pWB;
			return TRUE;
		}
	}
	return FALSE;
}

BOOL teFindSubWindow(HWND hwnd, size_t *pu)
{
	for (size_t u = g_pSubWindows.size(); u--;) {
		if (g_pSubWindows[u].hwnd == hwnd) {
			*pu = u;
			return TRUE;
		}
	}
	return FALSE;
}

int teGetthreadCount(DWORD dwProcessId)
{
	int nThread = 0;
	HANDLE hSnapshot;
	if ((hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, dwProcessId)) != INVALID_HANDLE_VALUE) {
		THREADENTRY32  te32 = { sizeof(THREADENTRY32) };

		if (Thread32First(hSnapshot, &te32)) {
			do {
				if (dwProcessId == te32.th32OwnerProcessID) {
					++nThread;
				}
			} while (Thread32Next(hSnapshot, &te32));
		}		CloseHandle(hSnapshot);
	}
	return nThread;
}

VOID teTranslateAccelerator(IDispatch *pdisp, MSG *pMsg, HRESULT *phr)
{
	if (pMsg->message == WM_KEYDOWN && GetKeyState(VK_CONTROL) < 0 && StrChr(L"LNOP\x6b\x6d\xbb\xbd", (WCHAR)pMsg->wParam)) {
		return;
	}
	IWebBrowser2 *pWB = NULL;
	if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pWB))) {
		IOleInPlaceActiveObject *pActiveObject = NULL;
		if SUCCEEDED(pWB->QueryInterface(IID_PPV_ARGS(&pActiveObject))) {
			if (pActiveObject->TranslateAcceleratorW(pMsg) == S_OK) {
				*phr = S_OK;
			}
			pActiveObject->Release();
		}
		pWB->Release();
	}
}

VOID PushFolderSizeList(TEFS* pFS)
{
	EnterCriticalSection(&g_csFolderSize);
	try {
		g_pFolderSize.push_back(pFS);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"PushFolderSizeList";
#endif
	}
	LeaveCriticalSection(&g_csFolderSize);
}

BOOL PopFolderSizeList(TEFS** ppFS)
{
	BOOL bResult = FALSE;
	EnterCriticalSection(&g_csFolderSize);
	try {
		if (bResult = !g_pFolderSize.empty()) {
			*ppFS = g_pFolderSize.back();
			g_pFolderSize.pop_back();
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"PopFolderSizeList";
#endif
	}
	LeaveCriticalSection(&g_csFolderSize);
	return bResult;
}

BOOL teIsSameSort(IFolderView2 *pFV2, SORTCOLUMN *pCol1, int nCount1)
{
	int i;
	SORTCOLUMN pCol2[8];
	BOOL bResult = SUCCEEDED(pFV2->GetSortColumnCount(&i)) && nCount1 == i && i <= _countof(pCol2) && SUCCEEDED(pFV2->GetSortColumns(pCol2, i));
	while (bResult && --i >= 0) {
		bResult = (pCol1[i].direction == pCol2[i].direction && IsEqualPropertyKey(pCol1[i].propkey, pCol2[i].propkey));
	}
	return bResult;
}

HRESULT teGetCodecInfo(IWICBitmapCodecInfo *pCodecInfo, int nMode, UINT cch, WCHAR *wz, UINT *pcchActual)
{
	switch (nMode) {
		case 0:
			return pCodecInfo->GetFileExtensions(cch, wz, pcchActual);
		case 1:
			return pCodecInfo->GetMimeTypes(cch, wz, pcchActual);
		case 2:
			return pCodecInfo->GetFriendlyName(cch, wz, pcchActual);
		case 3:
			return pCodecInfo->GetAuthor(cch, wz, pcchActual);
		case 4:
			return pCodecInfo->GetSpecVersion(cch, wz, pcchActual);
	}
	return E_NOTIMPL;
}

VOID teQueryFolderItem(FolderItem **ppfi, CteFolderItem **ppid)
{
	if FAILED((*ppfi)->QueryInterface(g_ClsIdFI, (LPVOID *)ppid)) {
		*ppid = new CteFolderItem(NULL);
		LPITEMIDLIST pidl;
		if (teGetIDListFromObject(*ppfi, &pidl)) {
			if SUCCEEDED((*ppid)->Initialize(pidl)) {
				(*ppfi)->Release();
				(*ppid)->QueryInterface(IID_PPV_ARGS(ppfi));
			}
			::CoTaskMemFree(pidl);
		}
	}
}

BOOL teGetHTMLWindow(IWebBrowser2 *pWB2, REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;
	IDispatch *pdisp;
	if (pWB2->get_Document(&pdisp) == S_OK) {
		IHTMLDocument2 *pdoc;
		if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pdoc))) {
			IHTMLWindow2 *pwin;
			if SUCCEEDED(pdoc->get_parentWindow(&pwin)) {
				pwin->QueryInterface(riid, ppvObject);
			}
			pdoc->Release();
		}
		pdisp->Release();
	}
	return *ppvObject != NULL;
}

BSTR teMultiByteToWideChar(UINT CodePage, LPCSTR lpA, int nLenA)
{
	int nLenW = ::MultiByteToWideChar(CodePage, 0, lpA, nLenA, NULL, NULL);
	BSTR bs = ::SysAllocStringLen(NULL, nLenW);
	::MultiByteToWideChar(CodePage, 0, lpA, nLenA, bs, nLenW);
	return bs;
}


LPSTR teWideCharToMultiByte(UINT CodePage, LPCWSTR lpW, int nLenW)
{
	int nLenA = WideCharToMultiByte(CodePage, 0, lpW, nLenW, NULL, 0, NULL, NULL);
	LPSTR bsA = (LPSTR)::SysAllocStringByteLen(NULL, nLenA);
	WideCharToMultiByte(CodePage, 0, lpW, nLenW, bsA, nLenA, NULL, NULL);
	return bsA;
}

BSTR teSysAllocStringByteLen(LPCSTR psz, UINT len, UINT org)
{
	if (!len) {
		return NULL;
	}
	BSTR bs = ::SysAllocStringByteLen(NULL, len);
	if (psz && bs && org) {
		::CopyMemory(bs, psz, org < len ? org : len);
	}
	return bs;
}

BSTR teLoadFromFile(BSTR bsFile)
{
	BSTR bsResult = NULL;
	HANDLE hFile = CreateFile(bsFile, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile != INVALID_HANDLE_VALUE) {
		DWORD dwFileSize, dwFileSize2;
		dwFileSize = GetFileSize(hFile, &dwFileSize2);
		if(dwFileSize != INVALID_FILE_SIZE) {
			BSTR bsUTF8 = ::SysAllocStringByteLen(NULL, dwFileSize);
			if (::ReadFile(hFile, bsUTF8, dwFileSize, &dwFileSize2, NULL)) {
				bsResult = teMultiByteToWideChar(CP_UTF8, (LPCSTR)bsUTF8, dwFileSize2);
			}
			::SysFreeString(bsUTF8);
		}
		CloseHandle(hFile);
	}
	return bsResult;
}

LONGLONG teGetU(LONGLONG ll)
{
	if ((ll & (~MAXINT)) == (~MAXINT)) {
		return ll & MAXUINT;
	}
	return ll & MAXINT64;
}

VOID SafeRelease(PVOID ppObj)
{
	try {
		IUnknown **ppunk = static_cast<IUnknown **>(ppObj);
		if (*ppunk) {
			(*ppunk)->Release();
			*ppunk = NULL;
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"SafeRelease";
#endif
	}
}

HRESULT teSetSFVCB(IShellView *pSV, IShellFolderViewCB* pNewCB, IShellFolderViewCB** ppOldCB)
{
	HRESULT hr = E_NOTIMPL;
#ifdef _2000XP
	if (!g_bUpperVista) {
		return hr;
	}
#endif
	IShellFolderView *pSFV;
	if (pSV->QueryInterface(IID_PPV_ARGS(&pSFV)) == S_OK) {
		IShellFolderViewCB *pSFVCB = NULL;
		if (ppOldCB == NULL) {
			ppOldCB = &pSFVCB;
		}
		hr = pSFV->SetCallback(pNewCB, ppOldCB);
		SafeRelease(&pSFVCB);
		pSFV->Release();
	}
	return hr;
}

LONGLONG teGetStreamPos(IStream *pStream)
{
	ULARGE_INTEGER uliPos;
	LARGE_INTEGER liOffset;
	liOffset.QuadPart = 0;
	pStream->Seek(liOffset, STREAM_SEEK_CUR, &uliPos);
	return uliPos.QuadPart;
}

DWORD teGetStreamSize(IStream *pStream)
{
	LARGE_INTEGER liOffset;
	liOffset.QuadPart = 0;
	ULARGE_INTEGER uliSize;
	pStream->Seek(liOffset, STREAM_SEEK_END, &uliSize);
	pStream->Seek(liOffset, STREAM_SEEK_SET, NULL);
	return uliSize.LowPart;
}

VOID teCopyStream(IStream *pSrc, IStream *pDst)
{
	LARGE_INTEGER liOffset;
	liOffset.QuadPart = 0;
	LARGE_INTEGER liSrc, liDst;
	liSrc.QuadPart = teGetStreamPos(pSrc);
	liDst.QuadPart = teGetStreamPos(pDst);
	pSrc->Seek(liOffset, STREAM_SEEK_SET, NULL);
	pDst->Seek(liOffset, STREAM_SEEK_SET, NULL);
	ULONG cbRead;
	BYTE pszData[SIZE_BUFF];
	while (SUCCEEDED(pSrc->Read(pszData, SIZE_BUFF, &cbRead)) && cbRead) {
		pDst->Write(pszData, cbRead, NULL);
	}
	pSrc->Seek(liSrc, STREAM_SEEK_SET, NULL);
	pDst->Seek(liDst, STREAM_SEEK_SET, NULL);
}

VOID teCommaSize(LPWSTR pszIn, LPWSTR pszOut, UINT cchBuf, int nDigits)
{
	NUMBERFMT nbf;
	WCHAR pszDecimal[4];
	WCHAR pszThousand[4];
	GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, pszDecimal, _countof(pszDecimal));
	GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, pszThousand, _countof(pszThousand));

	nbf.NumDigits = nDigits;
	nbf.LeadingZero = 1;
	nbf.Grouping = 3;
	nbf.lpDecimalSep = pszDecimal;
	nbf.lpThousandSep = pszThousand;
	nbf.NegativeOrder = 1;
	GetNumberFormat(LOCALE_USER_DEFAULT, 0, pszIn, &nbf, pszOut, cchBuf);
}

VOID teStrFormatSize(DWORD dwFormat, LONGLONG qdw, LPWSTR pszBuf, UINT cchBuf)
{
	if (dwFormat <= 1) {
		StrFormatKBSize(qdw, pszBuf, cchBuf);
		return;
	}
	if (dwFormat == 2) {
		StrFormatByteSize(qdw, pszBuf, cchBuf);
		return;
	}
	if (dwFormat == 3) {
		WCHAR pszNum[32];
		swprintf_s(pszNum, 32, L"%llu", qdw);
		teCommaSize(pszNum, pszBuf, cchBuf, 0);
		return;
	}
	LPCSTR pszPrefix = "\0KMGTPE";
	int flag = (dwFormat / 100) % 10;
	int i = (dwFormat / 10) % 10;
	int j = i;
	LONGLONG llPow = 1;
	if (i < 7) {
		while (--i >= 0) {
			llPow *= (flag != 2) ? 1024 : 1000;
		}
	} else {
		j = 0;
	}
	WCHAR pszNum[32];
	swprintf_s(pszNum, 32, L"%f", double(qdw) / llPow);
	teCommaSize(pszNum, pszBuf, cchBuf, dwFormat > 10 ? dwFormat % 10 : 0);
	swprintf_s(&pszBuf[lstrlen(pszBuf)], 5, L" %c%sB", pszPrefix[j], flag == 1 ? L"i" : L"");
}

BSTR teSysAllocStringLen(const OLECHAR *strIn, UINT uSize)
{
	UINT uOrg = lstrlen(strIn);
	if (strIn && uSize > uOrg) {
		BSTR bs = ::SysAllocStringLen(NULL, uSize);
		lstrcpy(bs, strIn);
		return bs;
	}
	return ::SysAllocStringLen(strIn, uSize);
}

BSTR teSysAllocStringLenEx(const BSTR strIn, UINT uSize)
{
	UINT uOrg = ::SysStringLen(strIn);
	if (uSize > uOrg) {
		BSTR bs = ::SysAllocStringLen(NULL, uSize);
		if (bs) {
			::CopyMemory(bs, strIn, ::SysStringByteLen(strIn) + sizeof(WORD));
		}
		return bs;
	}
	return ::SysAllocStringLen(strIn, uSize);
}

VOID teSysFreeString(BSTR *pbs)
{
	if (*pbs) {
		::SysFreeString(*pbs);
		*pbs = NULL;
	}
}

VOID teGetSearchArg(BSTR *pbs, LPWSTR pszPath, LPWSTR pszArg) {//L"&crumb=location:"
	*pbs = NULL;
	LPWSTR pszPath1 = StrStrI(pszPath, pszArg);
	if (pszPath1) {
		BSTR bsPath3 = ::SysAllocString(pszPath1 + lstrlen(pszArg));
		if (pszPath1 = StrChr(bsPath3, '&')) {
			pszPath1[0] = NULL;
		}
		UrlUnescape(bsPath3, NULL, NULL, URL_UNESCAPE_INPLACE);
		*pbs = ::SysAllocString(bsPath3[0] ? bsPath3 : L"*");
		teSysFreeString(&bsPath3);
	}
}

BOOL teSetWindowText(HWND hwnd, LPCWSTR lpcwstr)
{
	if (hwnd == g_hwndMain) {
		teSysFreeString(&g_bsTitle);
		g_bsTitle = ::SysAllocString(lpcwstr);
		SetTimer(g_hwndMain, TET_Title, 100, teTimerProc);
		return TRUE;
	}
	return SetWindowText(hwnd, lpcwstr);
}


VOID teSetLong(VARIANT *pv, LONG i)
{
	if (pv) {
		pv->lVal = i;
		pv->vt = VT_I4;
	}
}

VOID teSetBool(VARIANT *pv, BOOL b)
{
	if (pv) {
		pv->boolVal = b ? VARIANT_TRUE : VARIANT_FALSE;
		pv->vt = VT_BOOL;
	}
}

VOID teSetBSTR(VARIANT *pv, BSTR *pbs, int nLen)
{
	if (pv) {
		pv->vt = VT_BSTR;
		if (*pbs) {
			if (nLen < 0) {
				nLen = lstrlen(*pbs);
			}
			if (::SysStringLen(*pbs) == nLen) {
				pv->bstrVal = *pbs;
				*pbs = NULL;
				return;
			}
			pv->bstrVal = teSysAllocStringLenEx(*pbs, nLen);
			teSysFreeString(pbs);
			return;
		}
		pv->bstrVal = NULL;
		return;
	}
	teSysFreeString(pbs);
}

VOID teSetLL(VARIANT *pv, LONGLONG ll)
{
	if (pv) {
		pv->lVal = static_cast<int>(ll);
		if (ll == static_cast<LONGLONG>(pv->lVal)) {
			pv->vt = VT_I4;
			return;
		}
		pv->dblVal = static_cast<DOUBLE>(ll);
		if (ll == static_cast<LONGLONG>(pv->dblVal)) {
			pv->vt = VT_R8;
			return;
		}
		if (g_nBlink == 1) {// Next version
			pv->bstrVal = ::SysAllocStringLen(NULL, 18);
			swprintf_s(pv->bstrVal, 19, L"0x%016llx", ll);
			pv->vt = VT_BSTR;
			return;
		}
		SAFEARRAY *psa = SafeArrayCreateVector(VT_I4, 0, sizeof(LONGLONG) / sizeof(int));
		if (psa) {
			PVOID pvData;
			if (::SafeArrayAccessData(psa, &pvData) == S_OK) {
				::CopyMemory(pvData, &ll, sizeof(LONGLONG));
				::SafeArrayUnaccessData(psa);
				pv->vt = VT_ARRAY | VT_I4;
				pv->parray = psa;
			}
		}
	}
}

VOID teSetULL(VARIANT *pv, LONGLONG ll)
{
	if (pv) {
		pv->ulVal = static_cast<UINT>(ll);
		if (ll == static_cast<LONGLONG>(pv->ulVal)) {
			pv->vt = VT_UI4;
			return;
		}
		teSetLL(pv, ll);
	}
}

VOID teSetSZ(VARIANT *pv, LPCWSTR lpstr)
{
	if (pv) {
		pv->bstrVal = ::SysAllocString(lpstr);
		pv->vt = VT_BSTR;
	}
}

VOID teCoTaskMemFree(LPVOID pv)
{
	if (pv) {
		try {
			::CoTaskMemFree(pv);
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"CoTaskMemFree";
#endif
		}
	}
}

HRESULT teCreateInstance(CLSID clsid, LPWSTR lpszDllFile, HMODULE *phDll, REFIID riid, PVOID *ppvObj)
{
	*ppvObj = NULL;
	HMODULE hDll = NULL;
	HRESULT hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, riid, ppvObj);
	if FAILED(hr) {
		BSTR bsDllFile = NULL;
		WCHAR pszPath[MAX_PATH];
		if (lstrlen(lpszDllFile) == 0) {
			HKEY hKey;
			lstrcpy(pszPath, L"CLSID\\");
			LPOLESTR lpClsid;
			StringFromCLSID(clsid, &lpClsid);
			lstrcat(pszPath, lpClsid);
			teCoTaskMemFree(lpClsid);
			lstrcat(pszPath, L"\\InprocServer32");
			if (RegOpenKeyEx(HKEY_CLASSES_ROOT, pszPath, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
				DWORD dwSize = ARRAYSIZE(pszPath);
				if (RegQueryValueEx(hKey, NULL, NULL, NULL, (LPBYTE)&pszPath, &dwSize) == ERROR_SUCCESS) {
					ExpandEnvironmentStrings(pszPath, pszPath, MAX_PATH);
					lpszDllFile = pszPath;
				}
				RegCloseKey(hKey);
			}
		}
		hDll = LoadLibrary(lpszDllFile);
		if (hDll) {
			LPFNDllGetClassObject _DllGetClassObject = (LPFNDllGetClassObject)GetProcAddress(hDll, "DllGetClassObject");
			if (_DllGetClassObject) {
				IClassFactory *pCF;
				hr = _DllGetClassObject(clsid, IID_PPV_ARGS(&pCF));
				if (hr == S_OK) {
					hr = pCF->CreateInstance(NULL, riid, ppvObj);
					pCF->Release();
					if (hr == S_OK && phDll) {
						*phDll = hDll;
					}
				}
			} else {
				FreeLibrary(hDll);
			}
		}
	}
	return hr;
}

BOOL teStrSameIFree(BSTR bs, LPWSTR lpstr2)
{
	BOOL b = lstrcmpi(bs, lpstr2) == 0;
	::SysFreeString(bs);
	return b;
}

BOOL teCompareSFClass(IShellFolder *pSF, const CLSID *pclsid)
{
	BOOL bResult = FALSE;
	if (pSF) {
		IPersist *pPersist;
		if SUCCEEDED(pSF->QueryInterface(IID_PPV_ARGS(&pPersist))) {
			CLSID clsid;
			if SUCCEEDED(pPersist->GetClassID(&clsid)) {
				bResult = IsEqualCLSID(clsid, *pclsid);
			}
			pPersist->Release();
		}
	}
	return bResult;
}

HRESULT teGetPropertyKeyFromName(IShellFolder2 *pSF2, BSTR bs, PROPERTYKEY *pkey)
{
	HRESULT hr = _PSPropertyKeyFromStringEx(bs, pkey);
	if FAILED(hr) {
		if (pSF2) {
			SHELLDETAILS sd;
			for (UINT i = 0; i < MAX_COLUMNS && pSF2->GetDetailsOf(NULL, i, &sd) == S_OK; ++i) {
				BSTR bs1;
				if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs1)) {
					if (teStrSameIFree(bs1, bs)) {
						hr = pSF2->MapColumnToSCID(i, pkey);
						if (hr == S_OK) {
							break;
						}
					}
				}
			}
		}
		if FAILED(hr) {
			if (!teCompareSFClass(pSF2, &CLSID_ShellFSFolder) && !teCompareSFClass(pSF2, &CLSID_ShellDesktop)) {
				IShellFolder *pSF;
				if SUCCEEDED(SHGetDesktopFolder(&pSF)) {
					if (teCompareSFClass(pSF, &CLSID_ShellDesktop)) {
						IShellFolder2 *pSF2a;
						if SUCCEEDED(pSF->QueryInterface(IID_PPV_ARGS(&pSF2a))) {
							hr = teGetPropertyKeyFromName(pSF2a, bs, pkey);
							pSF2a->Release();
						}
					}
					pSF->Release();
				}
			}
		}
	}
	return hr;
}

BSTR tePSGetNameFromPropertyKeyEx(PROPERTYKEY propKey, int nFormat, CteShellBrowser *pSB)
{
	if (nFormat == 2) {
		WCHAR szProp[64];
#ifdef _2000XP
		if (_PSStringFromPropertyKey) {
			_PSStringFromPropertyKey(propKey, szProp, 64);
		} else {
			StringFromGUID2(propKey.fmtid, szProp, 39);
			wchar_t pszId[8];
			swprintf_s(pszId, 8, L" %u", propKey.pid);
			lstrcat(szProp, pszId);
		}
#else
		PSStringFromPropertyKey(propKey, szProp, 64);
#endif
		return ::SysAllocString(szProp);
	}
#ifdef _2000XP
	if (_PSGetPropertyDescription) {
#endif
		BSTR bs = NULL;
		IPropertyDescription *pdesc;
		if SUCCEEDED(_PSGetPropertyDescription(propKey, IID_PPV_ARGS(&pdesc))) {
			LPWSTR psz = NULL;
			CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_NAME };
			cmci.wszName[0] = NULL;
			if (nFormat) {
				pdesc->GetCanonicalName(&psz);
			} else {
				IColumnManager *pColumnManager;
				if (pSB && pSB->m_pShellView && SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager)))) {
					pColumnManager->GetColumnInfo(propKey, &cmci);
					pColumnManager->Release();
				}
				if (cmci.wszName[0]) {
					bs = ::SysAllocString(cmci.wszName);
				} else {
					pdesc->GetDisplayName(&psz);
				}
			}
			if (psz) {
				bs = ::SysAllocString(psz);
				teCoTaskMemFree(psz);
			}
			if (nFormat == 0 && pSB && bs) {
				PROPERTYKEY propKey1;
				if SUCCEEDED(teGetPropertyKeyFromName(pSB->m_pSF2, bs, &propKey1)) {
					if (!IsEqualPropertyKey(propKey, propKey1)) {
						teSysFreeString(&bs);
						pdesc->GetCanonicalName(&psz);
						bs = ::SysAllocString(psz);
						teCoTaskMemFree(psz);
					}
				}
			}
			pdesc->Release();
			return bs;
		}
		return tePSGetNameFromPropertyKeyEx(propKey, 2, pSB);
#ifdef _2000XP
	}
	if (IsEqualPropertyKey(propKey, PKEY_TotalFileSize)) {
		return ::SysAllocString(nFormat ? g_szTotalFileSizeCodeXP : g_szTotalFileSizeXP);
	}
	if (IsEqualPropertyKey(propKey, PKEY_Contact_Label)) {
		return ::SysAllocString(nFormat ? g_szLabelCodeXP : g_szLabelXP);
	}
	WCHAR szProp[MAX_PROP];
	IPropertyUI *pPUI;
	if SUCCEEDED(teCreateInstance(CLSID_PropertiesUI, NULL, NULL, IID_PPV_ARGS(&pPUI))) {
		if (nFormat) {
			pPUI->GetCannonicalName(propKey.fmtid, propKey.pid, szProp, MAX_PROP);
		} else {
			pPUI->GetDisplayName(propKey.fmtid, propKey.pid, PUIFNF_DEFAULT, szProp, MAX_PROP);
		}
		pPUI->Release();
	}
	return ::SysAllocString(szProp);
#endif
	return NULL;
}

BOOL teSetForegroundWindow(HWND hwnd)
{
	BOOL bResult = SetForegroundWindow(hwnd);
	if (!bResult) {
		DWORD dwForeThreadId = GetWindowThreadProcessId(GetForegroundWindow(), NULL);
		AttachThreadInput(g_dwMainThreadId, dwForeThreadId, TRUE);
		SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
		bResult = SetForegroundWindow(hwnd);
		AttachThreadInput(g_dwMainThreadId, dwForeThreadId, FALSE);
		if (!bResult) {
			DWORD dwTimeout = 0;
			SystemParametersInfo(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, &dwTimeout, 0);
			SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, NULL, 0);
			SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
			bResult = SetForegroundWindow(hwnd);
			SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, &dwTimeout, 0);
			if (!bResult) {
				Sleep(500);
				bResult = SetForegroundWindow(hwnd);
				if (!bResult && hwnd == g_hwndMain) {
					teSetExStyleAnd(hwnd, ~WS_EX_TOOLWINDOW);
					ShowWindow(hwnd, SW_SHOWNORMAL);
					bResult = SetForegroundWindow(hwnd);
				}
			}
		}
	}
	return bResult;
}

BOOL teChangeWindowMessageFilterEx(HWND hwnd, UINT message, DWORD action, PCHANGEFILTERSTRUCT pChangeFilterStruct)
{
	if (_ChangeWindowMessageFilterEx && hwnd) {
		return _ChangeWindowMessageFilterEx(hwnd, message, action, pChangeFilterStruct);
	}
#ifdef _2000XP
	if (_ChangeWindowMessageFilter) {
		return _ChangeWindowMessageFilter(message, action);
	}
#else
	return ChangeWindowMessageFilter(message, action);
#endif
	return FALSE;
}

BOOL teGetStrFromFolderItem(BSTR *pbs, IUnknown *punk)
{
	*pbs = NULL;
	CteFolderItem *pid;
	if (punk && SUCCEEDED(punk->QueryInterface(g_ClsIdFI, (LPVOID *)&pid))) {
		*pbs = pid->GetStrPath();
		pid->Release();
	}
	return *pbs != NULL;
}

VOID teAdvise(IUnknown *punk, IID diid, IUnknown *punk2, PDWORD pdwCookie)
{
	IConnectionPointContainer *pCPC;
	if (SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCPC)))) {
		IConnectionPoint *pCP;
		if (SUCCEEDED(pCPC->FindConnectionPoint(diid, &pCP))) {
			pCP->Advise(punk2, pdwCookie);
			pCP->Release();
		}
		pCPC->Release();
	}
}

VOID teUnadviseAndRelease(IUnknown *punk, IID diid, PDWORD pdwCookie)
{
	if (punk) {
		IConnectionPointContainer *pCPC;
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCPC))) {
			IConnectionPoint *pCP;
			if (SUCCEEDED(pCPC->FindConnectionPoint(diid, &pCP))) {
				pCP->Unadvise(*pdwCookie);
				pCP->Release();
			}
			pCPC->Release();
		}
		punk->Release();
	}
	*pdwCookie = 0;
}

VOID teRevoke(IShellWindows *pSW)
{
	if (pSW) {
		if (pSW == g_pSW) {
			g_pSW = NULL;
		}
		try {
			if (g_lSWCookie) {
				pSW->Revoke(g_lSWCookie);
				g_lSWCookie = 0;
			}
			if (g_dwCookieSW) {
				teUnadviseAndRelease(pSW, DIID_DShellWindowsEvents, &g_dwCookieSW);
				g_dwCookieSW = 0;
			}
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"Revoke";
#endif
		}
		SafeRelease(&pSW);
	}
}

VOID teRegister()
{
	IShellWindows *pSW;
	if (SUCCEEDED(teCreateInstance(CLSID_ShellWindows,
		NULL, NULL, IID_PPV_ARGS(&pSW)))) {
		if (pSW == g_pSW) {
			pSW->Release();
		} else {
			IShellWindows *pSW2 = g_pSW;
			g_pSW = pSW;
			teRevoke(pSW2);
			if (g_pWebBrowser) {
				pSW->Register(g_pWebBrowser->m_pWebBrowser, HandleToLong(g_hwndMain), SWC_3RDPARTY, &g_lSWCookie);
				BSTR bsPath;
				teGetModuleFileName(NULL, &bsPath);//Executable Path
				LPITEMIDLIST pidl = SHSimpleIDListFromPath(bsPath);
				VARIANT v;
				if (GetVarArrayFromIDList(&v, pidl)) {
#ifdef _DEBUG
					HRESULT hr = pSW->OnNavigate(g_lSWCookie, &v);
					if (hr != S_OK) {
						WCHAR pszNum[10];
						swprintf_s(pszNum, 9, L"%08x", hr);
						MessageBox(g_hwndMain, pszNum, _T(PRODUCTNAME), MB_OK | MB_ICONERROR);
					}
#else
					pSW->OnNavigate(g_lSWCookie, &v);
#endif
					::VariantClear(&v);
				}
				teCoTaskMemFree(pidl);
				teSysFreeString(&bsPath);
			}
			if (g_dwCookieSW) {
				teAdvise(pSW, DIID_DShellWindowsEvents, static_cast<IDispatch *>(g_pTE), &g_dwCookieSW);
			}
		}
	}
}

VOID teLockUpdate(int nStep)
{
	g_nLockUpdate += nStep;
}

VOID teUnlockUpdate(int nStep)
{
	g_nLockUpdate -= nStep;
	if (g_nLockUpdate <= 0) {
		g_nLockUpdate = 0;
		for (size_t i = 0; i < g_ppTC.size(); ++i) {
			CteTabCtrl *pTC = g_ppTC[i];
			if (pTC->m_bVisible) {
				CteShellBrowser *pSB = pTC->GetShellBrowser(pTC->m_nIndex);
				if (pSB) {
					if (pSB->m_nUnload & 5) {
						pSB->Show(TRUE, 0);
					}
					pTC->RedrawUpdate();
				}
			}
		}
	}
}

HRESULT teExceptionEx(EXCEPINFO *pExcepInfo, LPCSTR pszObjA, LPCSTR pszNameA)
{
	g_nException = 0;
#ifdef _DEBUG
	if (!g_strException) {
		g_strException = L"Exception";
	}
#endif
	if (pExcepInfo) {
		BSTR bsObj = teMultiByteToWideChar(CP_UTF8, pszObjA, -1);
		BSTR bsName = teMultiByteToWideChar(CP_UTF8, pszNameA, -1);
		int nLen = ::SysStringLen(bsObj) + ::SysStringLen(bsName) + 15;
		pExcepInfo->bstrDescription = ::SysAllocStringLen(NULL, nLen);
		swprintf_s(pExcepInfo->bstrDescription, nLen, L"Exception in %s.%s", bsObj, pszNameA ? bsName : L"");
		pExcepInfo->bstrSource = ::SysAllocString(g_szTE);
		pExcepInfo->scode = E_UNEXPECTED;
		::SysFreeString(bsName);
		::SysFreeString(bsObj);
		return DISP_E_EXCEPTION;
	}
	return E_UNEXPECTED;
}

HRESULT teException(EXCEPINFO *pExcepInfo, LPCSTR pszObjA, TEmethod* pMethod, DISPID dispIdMember)
{
	LPSTR pszNameA = NULL;
	if (pMethod) {
		int i = 0;
		while (pMethod[i].id && pMethod[i].id != dispIdMember) {
			++i;
		}
		pszNameA = pMethod[i].name;
	}
	return teExceptionEx(pExcepInfo, pszObjA, pszNameA);
}

BOOL teILIsBlank(FolderItem *pFolderItem)
{
	BSTR bs = NULL;
	return !pFolderItem || (SUCCEEDED(pFolderItem->get_Path(&bs)) && teStrSameIFree(bs, PATH_BLANK));
}

VOID teILCloneReplace(LPITEMIDLIST *ppidl, LPCITEMIDLIST pidl)
{
	if (ppidl) {
		LPITEMIDLIST pidl2 = *ppidl;
		*ppidl = ILClone(pidl);
		teCoTaskMemFree(pidl2);
	}
}

VOID teILFreeClear(LPITEMIDLIST *ppidl)
{
	if (ppidl) {
		teCoTaskMemFree(*ppidl);
		*ppidl = NULL;
	}
}

#ifdef _USE_TESTPATHMATCHSPEC
BOOL tePathMatchSpec2(LPCWSTR pszFile, LPWSTR pszSpec)
{
	switch (*pszSpec)
	{
		case NULL:
		case ';':
			return !pszFile[0];
		case '*':
			return tePathMatchSpec2(pszFile, pszSpec + 1) || (pszFile[0] && tePathMatchSpec2(pszFile + 1, pszSpec));
		case '?':
			return pszFile[0] && tePathMatchSpec2(pszFile + 1, pszSpec + 1);
		default:
			return (towlower(pszFile[0]) == towlower(pszSpec[0])) && tePathMatchSpec2(pszFile + 1, pszSpec + 1);
	}
}
#endif

BOOL teStartsText(LPWSTR pszSub, LPCWSTR pszFile)
{
	BOOL bResult = pszFile ? TRUE : FALSE;
	WCHAR wc;
	while (bResult && (wc = *pszSub++)) {
		bResult = towlower(wc) == towlower(*pszFile++);
	}
	return bResult;
}

BOOL teIsSearchFolder(LPWSTR lpszPath)
{
	return teStartsText(L"search-ms:", lpszPath);
}

VOID teGetJunctionLinkTarget(BSTR bsPath, LPWSTR *ppszText, int cchTextMax)
{
	HANDLE hFile = CreateFile(bsPath, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT, NULL);
	if (hFile != INVALID_HANDLE_VALUE) {
		BYTE buf[MAXIMUM_REPARSE_DATA_BUFFER_SIZE];
		REPARSE_DATA_BUFFER *tpReparseDataBuffer = (_REPARSE_DATA_BUFFER*)buf;
		DWORD dwRet = 0;
		if (DeviceIoControl(hFile, FSCTL_GET_REPARSE_POINT, NULL, 0, tpReparseDataBuffer,
			MAXIMUM_REPARSE_DATA_BUFFER_SIZE, &dwRet, NULL)) {
			if (IsReparseTagMicrosoft(tpReparseDataBuffer->ReparseTag)) {
				if (tpReparseDataBuffer->ReparseTag == IO_REPARSE_TAG_MOUNT_POINT) {
					LPWSTR lpPathBuffer = (LPWSTR)tpReparseDataBuffer->MountPointReparseBuffer.PathBuffer;
					LPWSTR lpPath = &lpPathBuffer[tpReparseDataBuffer->MountPointReparseBuffer.SubstituteNameOffset / 2];
					if (teStartsText(L"\\??\\", lpPath)) {
						lpPath += 4;
					}
					if (*ppszText) {
						lstrcpyn(*ppszText, lpPath, cchTextMax);
					} else {
						*ppszText = ::SysAllocString(lpPath);
					}
				}
			}
		}
		CloseHandle(hFile);
	}
}

BOOL tePathMatchSpec1(LPCWSTR pszFile, LPWSTR pszSpec, WCHAR wSpecEnd)
{
	WCHAR wc = *pszSpec;
	if (wc == wSpecEnd) {
		return !*pszFile;
	}
	if (!*pszFile && wc != '*') {
		return FALSE;
	}
	for (; *pszFile; ++pszFile) {
		wc = *pszSpec++;
		if (wc == '*') {
			wc = towlower(*pszSpec++);
			if (wc == wSpecEnd) {
				return TRUE;
			}
			do {
				if (!*pszFile) {
					return FALSE;
				}
				if (wc != '*' && wc != '?') {
					while (towlower(*pszFile) != wc) {
						if (!*(++pszFile)) {
							return FALSE;
						}
					}
				}
			} while (!tePathMatchSpec1(++pszFile, pszSpec, wSpecEnd));
			return TRUE;
		}
		if (wc != '?') {
			if (wc == wSpecEnd || towlower(*pszFile) != towlower(wc)) {
				return FALSE;
			}
		}
	}
	for (; (wc = *pszSpec) == '*'; pszSpec++);
	return *pszFile == (wc == wSpecEnd ? NULL : wc);
}

BOOL tePathMatchSpec(LPCWSTR pszFile, LPWSTR pszSpec)
{
	LPWSTR pszSpecEnd;
	if (!pszSpec || !pszSpec[0]) {
		return TRUE;
	}
	if (!pszFile) {
		return FALSE;
	}
	do {
		pszSpecEnd = StrChr(pszSpec, ';');
#ifdef _USE_TESTPATHMATCHSPEC
		BOOL b1 = !!tePathMatchSpec1(pszFile, pszSpec, pszSpecEnd ? ';' : NULL);
		BOOL b2 = !!tePathMatchSpec2(pszFile, pszSpec);
		if (b1 != b2) {
			b2 = !!tePathMatchSpec1(pszFile, pszSpec, pszSpecEnd ? ';' : NULL);
		}
#endif
		if (tePathMatchSpec1(pszFile, pszSpec, pszSpecEnd ? ';' : NULL)) {
			return TRUE;
		}
		pszSpec = pszSpecEnd + 1;
	} while (pszSpecEnd);
	return FALSE;
}

BOOL tePathIsNetworkPath(LPCWSTR pszPath)//PathIsNetworkPath is slow in DRIVE_NO_ROOT_DIR.
{
	if (!pszPath) {
		return FALSE;
	}
	WCHAR pszDrive[4];
	lstrcpyn(pszDrive, pszPath, 4);
	if (pszDrive[0] >= 'A' && pszDrive[1] == ':' && pszDrive[2] == '\\') {
		UINT uDriveType = GetDriveType(pszDrive);
		return uDriveType == DRIVE_REMOTE || uDriveType == DRIVE_NO_ROOT_DIR;
	}
	return tePathMatchSpec(pszPath, L"\\\\*;*://*") && !teStartsText(L"\\\\\\", pszPath);
}

HRESULT STDAPICALLTYPE teGetDpiForMonitor(HMONITOR hmonitor, MONITOR_DPI_TYPE dpiType, UINT *dpiX, UINT *dpiY)
{
	HDC hdc = GetDC(g_hwndMain);
	*dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
	*dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
	ReleaseDC(g_hwndMain, hdc);
	return S_OK;
}

HRESULT STDAPICALLTYPE tePSPropertyKeyFromStringEx(__in LPCWSTR pszString, __out PROPERTYKEY *pkey)
{
	HRESULT hr = _PSPropertyKeyFromString(pszString, pkey);
	return SUCCEEDED(hr) ? hr : _PSGetPropertyKeyFromName(pszString, pkey);
}

int CalcCrc32(BYTE *pc, int nLen, UINT c)
{
	c ^= MAXUINT;
	for (int i = 0; i < nLen; ++i) {
		c = g_uCrcTable[LOBYTE(c ^ pc[i])] ^ (c >> 8);
	}
	return c ^ MAXUINT;
}

int teGetModuleFileName(HMODULE hModule, BSTR *pbsPath)
{
	int i = 0;
	*pbsPath = NULL;
	for (int nSize = MAX_PATH; nSize < MAX_PATHEX; nSize += MAX_PATH) {
		SysReAllocStringLen(pbsPath, NULL, nSize);
		i = GetModuleFileName(hModule, *pbsPath, nSize);
		if (i + 1 < nSize) {
			(*pbsPath)[0] = towupper((*pbsPath)[0]);
			(*pbsPath)[i] = NULL;
			break;
		}
	}
	return i;
}

int teDragQueryFile(HDROP hDrop, UINT iFile, BSTR *pbsPath)
{
	UINT i = 0;
	for (UINT nSize = MAX_PATH; nSize < MAX_PATHEX; nSize += MAX_PATH) {
		*pbsPath = ::SysAllocStringLen(NULL, nSize);
		i = DragQueryFile(hDrop, iFile, *pbsPath, nSize);
		if (i + 1 < nSize) {
			(*pbsPath)[i] = NULL;
			break;
		}
		::SysFreeString(*pbsPath);
	}
	return i;
}

VOID tePathAppend(BSTR *pbsPath, LPCWSTR pszPath, LPWSTR pszFile)
{
	BSTR bsPath = teSysAllocStringLen(pszPath, lstrlen(pszPath) + lstrlen(pszFile) + 1);
	PathAppend(bsPath, pszFile);
	*pbsPath = ::SysAllocString(bsPath);
	teSysFreeString(&bsPath);
}

BOOL GetShellFolder(IShellFolder **ppSF, LPCITEMIDLIST pidl)
{
	if (ILIsEmpty(pidl)) {
		SHGetDesktopFolder(ppSF);
		return TRUE;
	}
	IShellFolder *pSF;
	LPCITEMIDLIST pidlPart;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
		pSF->BindToObject(pidlPart, NULL, IID_PPV_ARGS(ppSF));
		pSF->Release();
	}
	if (*ppSF == NULL) {
		return FALSE;
	}
#ifdef _2000XP
	if (!g_bUpperVista) {
		if (ILIsEqual(pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
			IPersistFolder *pPF;
			if SUCCEEDED((*ppSF)->QueryInterface(IID_PPV_ARGS(&pPF))) {
				pPF->Initialize(pidl);
				pPF->Release();
			}
		}
	}
#endif
	return TRUE;
}

LPITEMIDLIST teSHSimpleIDListFromPathEx(LPWSTR lpstr, DWORD dwAttr, DWORD nSizeLow, DWORD nSizeHigh, FILETIME *pft)
{
	LPITEMIDLIST pidl = NULL;
	IBindCtx *pbc = NULL;
	ULONG chEaten;
	ULONG dwAttributes;
	IShellFolder *pSF;
	if SUCCEEDED(SHGetDesktopFolder(&pSF)) {
		if SUCCEEDED(CreateBindCtx(0, &pbc)) {
			pbc->RegisterObjectParam(STR_PARSE_PREFER_FOLDER_BROWSING, static_cast<IDropSource *>(g_pTE));
			WIN32_FIND_DATA wfd = { 0 };
			int n = lstrlen(lpstr);
			wfd.dwFileAttributes = dwAttr | ((n > 0 && lpstr[n - 1] == '\\') ? FILE_ATTRIBUTE_DIRECTORY : 0);
			wfd.nFileSizeLow = nSizeLow;
			wfd.nFileSizeHigh = nSizeHigh;
			if (pft) {
				wfd.ftLastWriteTime = *pft;
			}
			IFileSystemBindData *pFSBD = new CteFileSystemBindData();
			pFSBD->SetFindData(&wfd);
			pbc->RegisterObjectParam(STR_FILE_SYS_BIND_DATA, pFSBD);
			pFSBD->Release();
		}
		try {
			pSF->ParseDisplayName(NULL, pbc, lpstr, &chEaten, &pidl, &dwAttributes);
		} catch (...) {
			pidl = NULL;
		}
		SafeRelease(&pbc);
		pSF->Release();
	}
	return pidl;
}

ULONGLONG teGetFolderSize(LPCWSTR szPath, int nItems, PDWORD pdwSessionId, DWORD dwSessionId)
{
	if (PathIsRoot(szPath)) {
		ULARGE_INTEGER FreeBytesOfAvailable;
		ULARGE_INTEGER TotalNumberOfBytes;
		ULARGE_INTEGER TotalNumberOfFreeBytes;
		return GetDiskFreeSpaceEx(szPath, &FreeBytesOfAvailable, &TotalNumberOfBytes, &TotalNumberOfFreeBytes) ? TotalNumberOfBytes.QuadPart - TotalNumberOfFreeBytes.QuadPart : 0;
	}
	ULONGLONG Result = 0;
	std::vector<BSTR>	pFolders;
	WIN32_FIND_DATA wfd;
	pFolders.push_back(::SysAllocString(szPath));

	while (g_bMessageLoop && *pdwSessionId == dwSessionId && !pFolders.empty()) {
		BSTR bsPath = pFolders.back();
		pFolders.pop_back();
		BSTR bs2 = NULL;
		tePathAppend(&bs2, bsPath, L"*");
		HANDLE hFind = FindFirstFile(bs2, &wfd);
		teSysFreeString(&bs2);
		if (hFind != INVALID_HANDLE_VALUE) {
			do {
				if (nItems-- < 0) {
					++dwSessionId;
					break;
				}
				if (tePathMatchSpec(wfd.cFileName, L".;..")) {
					continue;
				}
				if ((wfd.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) == FILE_ATTRIBUTE_DIRECTORY) {
					tePathAppend(&bs2, bsPath, wfd.cFileName);
					pFolders.push_back(bs2);
					continue;
				}
				ULARGE_INTEGER uli;
				uli.HighPart = wfd.nFileSizeHigh;
				uli.LowPart = wfd.nFileSizeLow;
				Result += uli.QuadPart;
			} while (g_bMessageLoop && *pdwSessionId == dwSessionId && FindNextFile(hFind, &wfd));
			FindClose(hFind);
		}
		teSysFreeString(&bsPath);
	}
	for (int i = int(pFolders.size()); --i >= 0;) {
		::SysFreeString(pFolders[i]);
	}
	pFolders.clear();
	return Result;
}

BOOL teSetObject(VARIANT *pv, PVOID pObj)
{
	if (pv && pObj) {
		try {
			IUnknown *punk = static_cast<IUnknown *>(pObj);
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->pdispVal))) {
				pv->vt = VT_DISPATCH;
				return TRUE;
			}
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->punkVal))) {
				pv->vt = VT_UNKNOWN;
				return TRUE;
			}
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"SetObject";
#endif
		}
	}
	return false;
}

BOOL teSetObjectRelease(VARIANT *pv, PVOID pObj)
{
	if (pObj) {
		try {
			IUnknown *punk = static_cast<IUnknown *>(pObj);
			if (pv) {
				if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->pdispVal))) {
					pv->vt = VT_DISPATCH;
					punk->Release();
					return TRUE;
				}
				if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->punkVal))) {
					pv->vt = VT_UNKNOWN;
					punk->Release();
					return TRUE;
				}
			}
			punk->Release();
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"SetObjectRelease";
#endif
		}
	}
	return FALSE;
}

BOOL teVariantTimeToFileTime(DOUBLE dt, LPFILETIME pft)
{
	SYSTEMTIME SysTime;
	if (::VariantTimeToSystemTime(dt, &SysTime)) {
		FILETIME ft;
		if (::SystemTimeToFileTime(&SysTime, &ft)) {
			return ::LocalFileTimeToFileTime(&ft, pft);
		}
	}
	return FALSE;
}

BOOL teFileTimeToVariantTime(LPFILETIME pft, DOUBLE *pdt)
{
	FILETIME ft;
	if (::FileTimeToLocalFileTime(pft, &ft)) {
		SYSTEMTIME SysTime;
		if (::FileTimeToSystemTime(&ft, &SysTime)) {
			return ::SystemTimeToVariantTime(&SysTime, pdt);
		}
	}
	return FALSE;
}

VOID GetDragItems(CteFolderItems **ppDragItems, IDataObject *pDataObj)
{
	SafeRelease(ppDragItems);
	*ppDragItems = new CteFolderItems(pDataObj, NULL);
}

BOOL teIsFileSystem(LPOLESTR pszPath)
{
	return tePathMatchSpec(pszPath, L"?:\\*;\\\\*\\*") && !teStartsText(L"\\\\\\", pszPath);
}

HRESULT tePathGetFileName(BSTR *pbs, LPWSTR pszPath)
{
	if (pszPath && lstrlen(pszPath) > 2) {
		if (teIsSearchFolder(pszPath)) {
			teGetSearchArg(pbs, pszPath, L"crumb=");
			return S_OK;
		}
		LPWSTR pszName = PathFindFileName(pszPath);
		if (pszName && pszName[0] != ':') {
			*pbs = ::SysAllocString(pszName);
			return S_OK;
		}
	}
	return E_FAIL;
}

BOOL teGetArchiveSpec(LPWSTR pszPath, BSTR *pbsArcPath, BSTR *pbsItem)
{
	BOOL bResult = FALSE;
	*pbsArcPath = ::SysAllocString(pszPath);
	*pbsItem = NULL;
	while (teIsFileSystem(*pbsArcPath)) {
		WIN32_FIND_DATA wfd;
		HANDLE hFind = FindFirstFile(*pbsArcPath, &wfd);
		if (hFind != INVALID_HANDLE_VALUE) {
			FindClose(hFind);
			bResult = !(wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && *pbsItem;
			break;
		}
		BSTR bsItem2 = NULL;
		LPWSTR lpszItem = PathFindFileName(*pbsArcPath);
		tePathAppend(&bsItem2, lpszItem, *pbsItem);
		teSysFreeString(pbsItem);
		*pbsItem = bsItem2;
		if (!PathRemoveFileSpec(*pbsArcPath)) {
			break;
		}
	}
	return bResult;
}

VARIANTARG* GetNewVARIANT(int n)
{
	if (!n) {
		return NULL;
	}
	VARIANT *pv = new VARIANTARG[n];
	while (n--) {
		VariantInit(&pv[n]);
	}
	return pv;
}

BOOL FindUnknown(VARIANT *pv, IUnknown **ppunk)
{
	if (pv) {
		if (pv->vt == VT_DISPATCH || pv->vt == VT_UNKNOWN) {
			*ppunk = pv->punkVal;
			return *ppunk != NULL;
		}
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return FindUnknown(pv->pvarVal, ppunk);
		}
		if (pv->vt == (VT_DISPATCH | VT_BYREF) || pv->vt == (VT_UNKNOWN | VT_BYREF)) {
			*ppunk = *pv->ppunkVal;
			return *ppunk != NULL;
		}
	}
	*ppunk = NULL;
	return false;
}

VOID GetFolderItemFromVariant(FolderItem **ppid, VARIANT *pv)
{
	IUnknown *punk;
	if (!FindUnknown(pv, &punk) || FAILED(punk->QueryInterface(IID_PPV_ARGS(ppid)))) {
		*ppid = new CteFolderItem(pv);
	}
}

BOOL GetDispatch(VARIANT *pv, IDispatch **ppdisp)
{
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		return SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(ppdisp)));
	}
	if (pv->vt == (VT_ARRAY | VT_VARIANT)) {
		*ppdisp = teCreateObj(TE_ARRAY, pv);
		return TRUE;
	}
	return FALSE;
}

VOID teRegisterDragDrop(HWND hwnd, IDropTarget *pDropTarget, IDropTarget **ppDropTarget)
{
	*ppDropTarget = static_cast<IDropTarget *>(GetPropA(hwnd, "OleDropTargetInterface"));
	if (*ppDropTarget) {
		(*ppDropTarget)->AddRef();
		RevokeDragDrop(hwnd);
		RegisterDragDrop(hwnd, pDropTarget);
	}
}

VOID teSetParent(HWND hwnd, HWND hwndParent)
{
	if (GetParent(hwnd) != hwndParent) {
		SetParent(hwnd, hwndParent);
	}
}

int ILGetCount(LPCITEMIDLIST pidl)
{
	if (ILIsEmpty(pidl)) {
		return 0;
	}
	return ILGetCount(::ILGetNext(pidl)) + 1;
}

HRESULT teGetDisplayNameBSTR(IShellFolder *pSF, PCUITEMID_CHILD pidl, SHGDNF uFlags, BSTR *pbs)
{
	STRRET strret;
	HRESULT hr = pSF->GetDisplayNameOf(pidl, uFlags & 0x3fffffff, &strret);
	if SUCCEEDED(hr) {
		hr = StrRetToBSTR(&strret, pidl, pbs);
		if (hr == S_OK && (uFlags & SHGDN_FORADDRESSBAR) && teIsSearchFolder(*pbs) && StrChr(*pbs, '\\')) {
			BSTR bs;
			if (teGetDisplayNameBSTR(pSF, pidl, uFlags & ~SHGDN_FORADDRESSBAR, &bs) == S_OK) {
				if (teIsFileSystem(bs)) {
					teSysFreeString(pbs);
					*pbs = bs;
				} else {
					teSysFreeString(&bs);
				}
			}
		}
	}
	return hr;
}

LPITEMIDLIST teILCreateFromPathEx(LPWSTR pszPath)
{
	LPITEMIDLIST pidl = NULL;
	IBindCtx *pbc = NULL;
	ULONG chEaten;
	ULONG dwAttributes;
	IShellFolder *pSF;
	if SUCCEEDED(SHGetDesktopFolder(&pSF)) {
		if SUCCEEDED(CreateBindCtx(0, &pbc)) {
			pbc->RegisterObjectParam(STR_PARSE_PREFER_FOLDER_BROWSING, static_cast<IDropSource *>(g_pTE));
		}
		try {
			pSF->ParseDisplayName(NULL, pbc, pszPath, &chEaten, &pidl, &dwAttributes);
		} catch (...) {
			pidl = NULL;
		}
		SafeRelease(&pbc);
		pSF->Release();
	}
	return pidl;
}

LPITEMIDLIST teILCreateFromPath3(IShellFolder *pSF, LPWSTR pszPath, HWND hwnd, int nDog)
{
	LPITEMIDLIST pidlResult = NULL;
	IEnumIDList *peidl = NULL;
	if (nDog-- < 0) {
		return NULL;
	}
	LPWSTR lpDelimiter = StrChr(pszPath, '\\');
	int nDelimiter = 0;
	if (lpDelimiter) {
		nDelimiter = (int)(lpDelimiter - pszPath);
	}
	if SUCCEEDED(pSF->EnumObjects(hwnd, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN | SHCONTF_INCLUDESUPERHIDDEN | SHCONTF_NAVIGATION_ENUM, &peidl)) {
		int ashgdn[] = { SHGDN_FORPARSING, SHGDN_INFOLDER, SHGDN_INFOLDER | SHGDN_FORPARSING, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_INFOLDER };
		BSTR bstr = NULL;
		LPWSTR lpfn = NULL;
		LPITEMIDLIST pidlPart = NULL;
		while (!pidlResult && peidl->Next(1, &pidlPart, NULL) == S_OK) {
			for (int j = _countof(ashgdn); !pidlResult && j--; teSysFreeString(&bstr)) {
				if SUCCEEDED(teGetDisplayNameBSTR(pSF, pidlPart, ashgdn[j], &bstr)) {
					lpfn = ashgdn[j] & SHGDN_INFOLDER ? bstr : PathFindFileName(bstr);
					if (nDelimiter && nDelimiter == lstrlen(lpfn) && StrCmpNI(const_cast<LPCWSTR>(pszPath), (LPCWSTR)lpfn, nDelimiter) == 0) {
						IShellFolder *pSF1;
						if SUCCEEDED(pSF->BindToObject(pidlPart, NULL, IID_PPV_ARGS(&pSF1))) {
							pidlResult = teILCreateFromPath3(pSF1, &lpDelimiter[1], NULL, nDog);
							pSF1->Release();
						}
						if (!pidlResult) {
							BSTR bsFull;
							if SUCCEEDED(teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_FORPARSING, &bsFull)) {
								LPITEMIDLIST pidlFull = teILCreateFromPathEx(bsFull);
								teSysFreeString(&bsFull);
								if (pidlFull) {
									if (GetShellFolder(&pSF1, pidlFull)) {
										pidlResult = teILCreateFromPath3(pSF1, &lpDelimiter[1], NULL, nDog);
										pSF1->Release();
									}
									teILFreeClear(&pidlFull);
								}
							}
						}
						continue;
					}
					if (lstrcmpi(lpfn, pszPath) == 0) {
						LPITEMIDLIST pidlParent;
						if (teGetIDListFromObject(pSF, &pidlParent)) {
							pidlResult = ILCombine(pidlParent, pidlPart);
						}
						teCoTaskMemFree(pidlParent);
						continue;
					}
				}
			}
			teCoTaskMemFree(pidlPart);
		}
		peidl->Release();
	}
	return pidlResult;
}

LPITEMIDLIST teILCreateFromPath2(LPITEMIDLIST pidlParent, LPWSTR pszPath, HWND hwnd)
{
	LPITEMIDLIST pidlResult = NULL;
	IShellFolder *pSF;
	if (GetShellFolder(&pSF, pidlParent)) {
		pidlResult = teILCreateFromPath3(pSF, pszPath, hwnd, 256);
		pSF->Release();
	}
	return pidlResult;
}

VOID teReleaseILCreate(TEILCreate *pILC)
{
	if (::InterlockedDecrement(&pILC->cRef) == 0) {
		teCoTaskMemFree(pILC->pidlResult);
		CloseHandle(pILC->hEvent);
		delete [] pILC;
	}
}

BOOL GetCSIDLFromPath(int *i, LPWSTR pszPath)
{
	int n = lstrlen(pszPath);
	if (n < 3 && pszPath[0] >= '0' && pszPath[0] <= '9') {
		swscanf_s(pszPath, L"%d", i);
		return (*i <= 9 && n == 1) || (*i >= 10 && *i < MAX_CSIDL2);
	}
	return FALSE;
}

VOID teCreateSearchPath(BSTR *pbs, LPWSTR pszPath, LPWSTR pszSearch)
{
	LPWSTR pszHeader = L"search-ms:crumb=";
	LPWSTR pszLocation = L"&crumb=location:";
	DWORD dw = 2048;
	BSTR bsSearch3 = teSysAllocStringLen(L"", dw);
	UrlEscape(pszSearch, bsSearch3, &dw, 0);
	dw = 2048;
	BSTR bsPath3 = ::SysAllocStringLen(NULL, dw);
	UrlEscape(pszPath, bsPath3, &dw, 0);

	teSysFreeString(pbs);
	*pbs = ::SysAllocStringLen(NULL, lstrlen(pszHeader) + lstrlen(bsSearch3) + lstrlen(pszLocation) + lstrlen(bsPath3));
	lstrcpy(*pbs, pszHeader);
	lstrcat(*pbs, bsSearch3);
	lstrcat(*pbs, pszLocation);
	lstrcat(*pbs, bsPath3);
	teSysFreeString(&bsPath3);
	teSysFreeString(&bsSearch3);
}

BSTR teGetSearchText1(LPCITEMIDLIST pidl, BOOL bColumnFilter)
{
	BSTR bsSearch = NULL;
	IShellFolder *pSF;
	LPCITEMIDLIST pidlPart;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
		teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_INFOLDER | SHGDN_FORPARSING, &bsSearch);
		BSTR bsSearch1 = NULL;
		BSTR bsSearch2 = NULL;
		teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_INFOLDER | SHGDN_FORPARSING, &bsSearch);
		teGetSearchArg(&bsSearch1, bsSearch, L"displayname=");
		teGetSearchArg(&bsSearch2, bsSearch, L"crumb=");
		teSysFreeString(&bsSearch);
		if (!bsSearch1 || !bsSearch2) {
			teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &bsSearch);
			if (!bsSearch1) {
				teGetSearchArg(&bsSearch1, bsSearch, L"displayname=");
			}
			if (!bsSearch2) {
				LPWSTR psz = StrStr(bsSearch, L"&crumb=location:");
				if (psz) {
					*psz = NULL;
				}
				teGetSearchArg(&bsSearch2, bsSearch, L"crumb=");
			}
			teSysFreeString(&bsSearch);
		}
		pSF->Release();
		if (bsSearch2) {
			int nPos = ::SysStringLen(bsSearch2) - ::SysStringLen(bsSearch1);
			if (!lstrcmpi(bsSearch1, &bsSearch2[nPos >= 0 && StrChr(bsSearch2, '~') ? nPos : 0]) || StrStrI(bsSearch2, g_bsAnyText)|| StrStrI(bsSearch2, g_bsAnyTextId)) {
				bsSearch = bsSearch1;
				bsSearch1 = NULL;
			} else if (bColumnFilter) {
				bsSearch = bsSearch2;
				bsSearch2 = NULL;
				if (!StrChr(bsSearch, '"')) {
					for (int i = 0; i < ::SysStringLen(bsSearch); ++i) {
						if (bsSearch[i] == ':' || bsSearch[i] == 0xff1a) {
							if (StrCmpNI(bsSearch, bsSearch1, i)) {
								if (lstrlen(bsSearch1) <= lstrlen(&bsSearch[i + 1]) - 2) {
									BOOL bName = !StrCmpNI(bsSearch, g_bsName, i);
									if (bName || StrChr(&bsSearch[i + 1], ':')) {
										LPWSTR pszHasComma = StrChr(bsSearch1, ',');
										if (pszHasComma) {
											lstrcpy(&bsSearch[i + 1], L"(");
										} else {
											bsSearch[i + 1] = NULL;
										}
										for (LPWSTR pszSearch = bsSearch1;;) {
											while (pszSearch[0] == ' ') {
												++pszSearch;
											}
											LPWSTR pszComma = StrChr(pszSearch, ',');
											i = lstrlen(bsSearch);
											if (bName && pszSearch[1] == ' ' && pszSearch[3] == ' ') {// name:0 - 9
//												swprintf_s(&bsSearch[i], ::SysStringLen(bsSearch) - i, L"%c..%c%c", pszSearch[0], pszSearch[4], 0xd7fb);
												swprintf_s(&bsSearch[i], ::SysStringLen(bsSearch) - i, L"(%c..%c OR ~<%c)", pszSearch[0], pszSearch[4], pszSearch[4]);
											} else {
												if (pszComma) {
													StrNCpy(&bsSearch[i], pszSearch, pszComma - pszSearch + 1);
												} else {
													lstrcpy(&bsSearch[i], pszSearch);
												}
												for (LPWSTR pszSpace = &bsSearch[i]; pszSpace = StrChr(pszSpace, ' '); lstrcpy(pszSpace, &pszSpace[1]));
												if (LPWSTR psz = StrChr(&bsSearch[i], '(')) {
													psz[0] = NULL;
												}
											}
											if (pszComma) {
												lstrcat(bsSearch, L" OR ");
												pszSearch = pszComma + 1;
											} else {
												break;
											}
										}
										if (pszHasComma) {
											lstrcat(bsSearch, L")");
										}
									}
								}
							}
							break;
						}
					}
				}
			}
		}
		teSysFreeString(&bsSearch2);
		teSysFreeString(&bsSearch1);
	}
	return bsSearch;
}

VOID teGetSearchText(BSTR *pbs, LPCITEMIDLIST pidl)
{
	LPITEMIDLIST pidl1 = ILClone(pidl);
	*pbs = NULL;
	while (!*pbs && teILIsSearchFolder(pidl1)) {
		*pbs = teGetSearchText1(pidl1, FALSE);
		ILRemoveLastID(pidl1);
	}
	teILFreeClear(&pidl1);
}

HRESULT teGetDisplayNameFromIDList(BSTR *pbs, LPITEMIDLIST pidl, SHGDNF uFlags)
{
	HRESULT hr = E_FAIL;
	IShellFolder *pSF;
	LPCITEMIDLIST pidlPart;

	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
		hr = teGetDisplayNameBSTR(pSF, pidlPart, uFlags & (~SHGDN_FORPARSINGEX), pbs);
		if (hr == S_OK) {
			BSTR bs = NULL;
			teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &bs);
			if (teIsSearchFolder(bs)) {
				if (uFlags & SHGDN_FORPARSING) {
					BSTR bsSearchX = NULL;
					LPITEMIDLIST pidl1 = ILClone(pidl);
					while (teILIsSearchFolder(pidl1)) {
						BSTR bsSearch = teGetSearchText1(pidl1, TRUE);
						if (bsSearch) {
							if (bsSearchX) {
								BSTR bs = teSysAllocStringLen(bsSearchX, ::SysStringLen(bsSearchX) + ::SysStringLen(bsSearch) + 1);
								lstrcat(bs, L" ");
								lstrcat(bs, bsSearch);
								teSysFreeString(&bsSearch);
								teSysFreeString(&bsSearchX);
								bsSearchX = bs;
							} else {
								bsSearchX = bsSearch;
							}
						}
						ILRemoveLastID(pidl1);
					}
					teILFreeClear(&pidl1);
					BSTR bsPath = NULL;
					teGetSearchArg(&bsPath, bs, L"&crumb=location:");
					/*
					WCHAR pszTemp[MAX_LOADSTRING], pszNull[MAX_PATH];
					LoadString(g_hShell32, 34132, pszTemp, MAX_LOADSTRING);
					swprintf_s(pszNull, MAX_PATH, pszTemp, PathFindFileName(bsPath));
					if (lstrcmpi(bsSearchX, pszNull) == 0) {
						teSysFreeString(&bsSearchX);
					}*/
					teCreateSearchPath(pbs, bsPath, bsSearchX);
					teSysFreeString(&bsPath);
					teSysFreeString(&bsSearchX);
				}
			} else {
				teSysFreeString(&bs);
				if (tePathMatchSpec(*pbs, L"?:\\*")) {
					if (uFlags & SHGDN_FORPARSINGEX) {
						BOOL bSpecial = FALSE;
						if (teIsFileSystem(*pbs)) {
							LPITEMIDLIST pidl2 = SHSimpleIDListFromPath(*pbs);
							if (!ILIsEqual(pidl, pidl2)) {
								teILFreeClear(&pidl2);
								int n;
								if (!GetCSIDLFromPath(&n, *pbs)) {
									teILCloneReplace(&pidl2, pidl);
									BSTR bs0 = ::SysAllocString(*pbs);
									teSysFreeString(pbs);
									while (!ILIsEmpty(pidl2)) {
										BSTR bs, bs2;
										if SUCCEEDED(teGetDisplayNameFromIDList(&bs, pidl2, SHGDN_INFOLDER | SHGDN_FORPARSING)) {
											if (teIsSearchFolder(bs)) {
												teSysFreeString(&bs);
												teSysFreeString(pbs);
												*pbs = bs0;
												bs0 = NULL;
												break;
											}
											if (*pbs) {
												tePathAppend(&bs2, bs, *pbs);
												teSysFreeString(&bs);
												teSysFreeString(pbs);
												*pbs = bs2;
											} else {
												*pbs = bs;
											}
											::ILRemoveLastID(pidl2);
										}
									}
									teSysFreeString(&bs0);
								}
							}
							teILFreeClear(&pidl2);
						}
					}
				} else if (((uFlags & (SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) == (SHGDN_FORADDRESSBAR | SHGDN_FORPARSING))) {
					if (ILGetCount(pidl) == 1) {
						LPITEMIDLIST pidl2 = teILCreateFromPath2(g_pidls[CSIDL_DESKTOP], *pbs, NULL);
						if (!ILIsEqual(pidl, pidl2)) {
							teSysFreeString(pbs);
							hr = teGetDisplayNameBSTR(pSF, pidlPart, uFlags & (~(SHGDN_FORPARSINGEX | SHGDN_FORADDRESSBAR)), pbs);
						}
						teCoTaskMemFree(pidl2);
					}
				}
			}
		}
		pSF->Release();
	}
 	return hr;
}

BOOL teILIsSearchFolder(LPCITEMIDLIST pidl)
{
	BOOL bResult = FALSE;
	IShellFolder *pSF;
	LPCITEMIDLIST pidlPart;
	BSTR bs;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
		teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &bs);
		bResult = teIsSearchFolder(bs);
		::SysFreeString(bs);
		pSF->Release();
	}
	return bResult;
}

BOOL teILIsFileSystemEx(LPCITEMIDLIST pidl)
{
	LPCITEMIDLIST pidlPart;
	IShellFolder *pSF;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
		SFGAOF sfAttr = SFGAO_FILESYSTEM | SFGAO_FILESYSANCESTOR | SFGAO_STORAGEANCESTOR | SFGAO_NONENUMERATED | SFGAO_DROPTARGET;
		if FAILED(pSF->GetAttributesOf(1, &pidlPart, &sfAttr)) {
			sfAttr = 0;
		}
		pSF->Release();
		if (sfAttr & (SFGAO_FILESYSTEM | SFGAO_FILESYSANCESTOR | SFGAO_STORAGEANCESTOR | SFGAO_NONENUMERATED | SFGAO_DROPTARGET)) {
			return TRUE;
		}
	}
	return FALSE;
}

BOOL teVarIsNumber(VARIANT *pv) {
	return pv->vt == VT_I4 || pv->vt == VT_R8 || pv->vt == (VT_ARRAY | VT_I4) || (pv->vt == VT_BSTR && ::SysStringLen(pv->bstrVal) == 18 && teStartsText(L"0x", pv->bstrVal));
}

void GetVarPathFromIDList(VARIANT *pVarResult, LPITEMIDLIST pidl, int uFlags)
{
	int i;

	if (uFlags & SHGDN_FORPARSINGEX) {
		for (i = 0; i < MAX_CSIDL; ++i) {
			if (g_pidls[i] && ::ILIsEqual(pidl, g_pidls[i])) {
				LPITEMIDLIST pidl2 = SHSimpleIDListFromPath(g_bsPidls[i]);
				if (!::ILIsEqual(pidl, pidl2)) {
					teSetLong(pVarResult, i);
				}
				teCoTaskMemFree(pidl2);
				break;
			}
		}
	}
	if (!teVarIsNumber(pVarResult)) {
		if SUCCEEDED(teGetDisplayNameFromIDList(&pVarResult->bstrVal, pidl, uFlags)) {
			pVarResult->vt = VT_BSTR;
		}
	}
}

// VARIANT Clean-up of an array
VOID teClearVariantArgs(int nArgs, VARIANTARG *pvArgs)
{
	if (pvArgs && nArgs > 0) {
		for (int i = nArgs ; i-- >  0;){
			VariantClear(&pvArgs[i]);
		}
		delete[] pvArgs;
		pvArgs = NULL;
	}
}

HRESULT Invoke5(IDispatch *pdisp, DISPID dispid, WORD wFlags, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs)
{
	HRESULT hr;
	// DISPPARAMS
	DISPPARAMS dispParams;
	dispParams.rgvarg = pvArgs;
	dispParams.cArgs = abs(nArgs);
	DISPID dispidName = DISPID_PROPERTYPUT;
	if (wFlags & DISPATCH_PROPERTYPUT) {
		dispParams.cNamedArgs = 1;
		dispParams.rgdispidNamedArgs = &dispidName;
	} else {
		dispParams.rgdispidNamedArgs = NULL;
		dispParams.cNamedArgs = 0;
	}
	try {
		hr = pdisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
			wFlags, &dispParams, pvResult, NULL, NULL);
#ifdef _DEBUG
		if (hr) {
			WCHAR pszFunc[99];
			swprintf_s(pszFunc, 99, L"Invoke: %x\n", hr);
			::OutputDebugString(pszFunc);
			if (hr == E_ACCESSDENIED) {
				Sleep(1);
			}
			if (hr == E_UNEXPECTED) {
				Sleep(1);
			}
		}
#endif
	} catch (...) {
		hr = E_FAIL;
	}
#ifdef _USE_SYNC
	if (hr == HRESULT_FROM_WIN32(ERROR_POSSIBLE_DEADLOCK)) {
		if (wFlags & DISPATCH_METHOD) {
			if (g_nBlink == 1 && dispid == DISPID_VALUE) {
				int nArgsA = abs(nArgs);
				SAFEARRAY *psa = SafeArrayCreateVector(VT_VARIANT, 0, nArgsA + 1);
				LONG l = 0;
				VARIANT v;
				v.pdispVal = pdisp;
				v.vt = VT_DISPATCH;
				SafeArrayPutElement(psa, &l, &v);
				while (l < nArgsA) {
					++l;
					SafeArrayPutElement(psa, &l, &pvArgs[nArgsA - l]);
				}
				v.parray = psa;
				v.vt = VT_ARRAY | VT_VARIANT;
				if (!g_pOnFunc[TE_FN]) {
					g_pOnFunc[TE_FN] = teCreateObj(TE_ARRAY, NULL);
				}
				teExecMethod(g_pOnFunc[TE_FN], L"push", NULL, -1, &v);
				VariantClear(&v);
				if (teGetObjectLength(g_pOnFunc[TE_FN]) < 2) {
					g_pWebBrowser->m_pWebBrowser->GetProperty(L"InvokeMethod", NULL);
				}
			}
		}
	}
#endif
	teClearVariantArgs(nArgs, pvArgs);
	return hr;
}

VOID teInvokeObject(IDispatch **ppdisp, WORD wFlags, VARIANT *pVarResult, int nArg, VARIANTARG *pvArgs)
{
	if (wFlags & DISPATCH_METHOD) {
		if (*ppdisp) {
			Invoke5(*ppdisp, DISPID_VALUE, wFlags, pVarResult, - nArg - 1, pvArgs);
		}
		return;
	}
	if (nArg >= 0) {
		SafeRelease(ppdisp);
		GetDispatch(&pvArgs[nArg], ppdisp);
	}
	teSetObject(pVarResult, *ppdisp);
}

HRESULT teDelProperty(IUnknown *punk, LPOLESTR sz)
{
	HRESULT hr = E_FAIL;
	IDispatchEx *pdex;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pdex))) {
		BSTR bs = ::SysAllocString(sz);
		hr = pdex->DeleteMemberByName(bs, fdexNameCaseSensitive);
		::SysFreeString(bs);
		pdex->Release();
	}
	return hr;
}

HRESULT tePutProperty0(IUnknown *punk, LPOLESTR sz, VARIANT *pv, DWORD grfdex)
{
	HRESULT hr = E_FAIL;
	DISPID dispid, putid;
	DISPPARAMS dispparams;
	IDispatchEx *pdex;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pdex))) {
		BSTR bs = ::SysAllocString(sz);
		hr = pdex->GetDispID(bs, grfdex, &dispid);
		if SUCCEEDED(hr) {
			putid = DISPID_PROPERTYPUT;
			dispparams.rgvarg = pv;
			dispparams.rgdispidNamedArgs = &putid;
			dispparams.cArgs = 1;
			dispparams.cNamedArgs = 1;
			hr = pdex->InvokeEx(dispid, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUTREF, &dispparams, NULL, NULL, NULL);
		}
		::SysFreeString(bs);
		pdex->Release();
	}
	if FAILED(hr) {
		IDispatch *pdisp;
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pdisp))) {
			DISPID dispIdMember;
			if SUCCEEDED(pdisp->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &dispIdMember)) {
				hr = Invoke5(pdisp, dispIdMember, DISPATCH_PROPERTYPUT, NULL, -1, pv);
			}
			pdisp->Release();
		}
	}
	return hr;
}

HRESULT tePutProperty(IUnknown *punk, LPOLESTR sz, VARIANT *pv)
{
	HRESULT hr = tePutProperty0(punk, sz, pv, fdexNameEnsure);
	VariantClear(pv);
	return hr;
}

BOOL teHasProperty(IDispatch *pdisp, LPOLESTR sz)
{
	DISPID dispid;
	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &dispid);
	return hr == S_OK;
}

HRESULT teGetProperty(IDispatch *pdisp, LPOLESTR sz, VARIANT *pv)
{
	DISPID dispid;
	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &dispid);
	if (hr == S_OK) {
		hr = Invoke5(pdisp, dispid, DISPATCH_PROPERTYGET, pv, 0, NULL);
	}
	return hr;
}

//GetProperty(Case-insensitive)
HRESULT teGetPropertyI(IDispatch *pdisp, LPOLESTR sz, VARIANT *pv)
{
	DISPID dispid;
	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &dispid);
	if FAILED(hr) {
		IDispatchEx *pdex;
		if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pdex))) {
			BSTR bs = ::SysAllocString(sz);
			hr = pdex->GetDispID(bs, fdexNameCaseInsensitive, &dispid);
			::SysFreeString(bs);
			pdex->Release();
		}
	}
	if (hr == S_OK) {
		hr = Invoke5(pdisp, dispid, DISPATCH_PROPERTYGET, pv, 0, NULL);
	}
	return hr;
}

HRESULT teExecMethod(IDispatch *pdisp, LPOLESTR sz, VARIANT *pvResult, int nArg, VARIANTARG *pvArgs)
{
	DISPID dispid;
	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &dispid);
	if (hr == S_OK) {
		return Invoke5(pdisp, dispid, DISPATCH_METHOD, pvResult, nArg, pvArgs);
	}
	teClearVariantArgs(nArg, pvArgs);
	return hr;
}

BOOL teGetDispatchFromString(VARIANT *pv, IDispatch **ppdisp)
{
	if (pv->vt == VT_BSTR) {
		VARIANT v;
		VariantInit(&v);
		teExecMethod(g_pJS, L"_c", &v, -1, pv);
		if (v.vt == VT_DISPATCH) {
			*ppdisp = v.pdispVal;
			return TRUE;
		}
	}
	return FALSE;
}

HRESULT teGetPropertyAt(IDispatch *pdisp, int i, VARIANT *pv)
{
	wchar_t pszName[8];
	swprintf_s(pszName, 8, L"%d", i);
	return teGetProperty(pdisp, pszName, pv);
}

HRESULT tePutPropertyAt(PVOID pObj, int i, VARIANT *pv)
{
	wchar_t pszName[8];
	swprintf_s(pszName, 8, L"%d", i);
	return tePutProperty((IUnknown *)pObj, pszName, pv);
}

VOID teCreateSafeArrayFromVariantArray(IDispatch *pdisp, VARIANT *pVarResult) {
	LONG nLen = teGetObjectLength(pdisp);
	SAFEARRAY *psa = SafeArrayCreateVector(VT_VARIANT, 0, nLen);
	VARIANT v;
	VariantInit(&v);
	while (--nLen >= 0) {
		if SUCCEEDED(teGetPropertyAt(pdisp, nLen, &v)) {
			::SafeArrayPutElement(psa, &nLen, &v);
			::VariantClear(&v);
		}
	}
	pVarResult->vt = VT_ARRAY | VT_VARIANT;
	pVarResult->parray = psa;
}

char* GetpcFromVariant(VARIANT *pv, VARIANT *pvMem)
{
	char *pc1 = NULL;
	CteMemory *pMem;
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		if SUCCEEDED(punk->QueryInterface(g_ClsIdStruct, (PVOID *)&pMem)) {
			char *pc = pMem->m_pc;
			pMem->Release();
			return pc;
		}
	}
	if (pvMem) {
		IDispatch *pdisp;
		if (GetDispatch(pv, &pdisp)) {
			teGetProperty(pdisp, L"_BLOB", pvMem);
			pdisp->Release();
			if (pvMem->vt == VT_BSTR) {
				return (char *)pvMem->bstrVal;
			}
		}
	}
	return NULL;
}

BOOL GetLLFromVariant2(LONGLONG *pll, VARIANT *pv) {
	if (pv->vt == VT_I4) {
		*pll = pv->lVal;
		return TRUE;
	}
	if (pv->vt == VT_R8) {
		*pll = (LONGLONG)pv->dblVal;
		return TRUE;
	}
	if (pv->vt == (VT_ARRAY | VT_I4)) {
		LONGLONG ll = 0;
		PVOID pvData;
		if (::SafeArrayAccessData(pv->parray, &pvData) == S_OK) {
			::CopyMemory(&ll, pvData, sizeof(LONGLONG));
			::SafeArrayUnaccessData(pv->parray);
			return ll;
		}
	}
	if (teVarIsNumber(pv)) {
		if (swscanf_s(pv->bstrVal, L"0x%016llx", pll) > 0) {
			return TRUE;
		}
	}
	return FALSE;
}

LONGLONG GetLLFromVariant(VARIANT *pv)
{
	if (pv) {
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return GetLLFromVariant(pv->pvarVal);
		}
		LONGLONG ll = 0;
		if (GetLLFromVariant2(&ll, pv)) {
			return ll;
		}
		if (pv->vt != VT_DISPATCH) {
			VARIANT vo;
			VariantInit(&vo);
			if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8)) {
				return vo.llVal;
			}
		}
#ifdef _W2000
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_R8)) {
			return (int)(LONGLONG)vo.dblVal;
		}
#endif
		char *pc = GetpcFromVariant(pv, NULL);
		if (pc) {
			return (LONGLONG)pc;
		}
	}
	return 0;
}

LONGLONG GetLLFromVariantClear(VARIANT *pv)
{
	LONGLONG ll = GetLLFromVariant(pv);
	VariantClear(pv);
	return ll;
}

LONGLONG GetParamFromVariant(VARIANT *pv, VARIANT *pvMem)
{
	if (pv) {
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return GetParamFromVariant(pv->pvarVal, pvMem);
		}
		LONGLONG ll = 0;
		if (GetLLFromVariant2(&ll, pv)) {
			return ll;
		}
		char *pc = GetpcFromVariant(pv, pvMem);
		if (pc) {
			return (LONGLONG)pc;
		}
		if (pv->vt == VT_DISPATCH || pv->vt == VT_UNKNOWN) {
			return TRUE;
		}
		VARIANT vo;
		VariantInit(&vo);
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8)) {
			return vo.llVal;
		}
#ifdef _W2000
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_R8)) {
			return (LONGLONG)vo.dblVal;
		}
#endif
	}
	return 0;
}

VOID teWriteBack(VARIANT *pvArg, VARIANT *pvDat)
{
	IDispatch *pdisp;
	if (pvDat->vt == VT_BSTR && pvDat->bstrVal && GetDispatch(pvArg, &pdisp)) {
		VARIANT v;
		if SUCCEEDED(teGetProperty(pdisp, L"_BLOB", &v)) {
			if (v.vt == VT_BSTR) {
				tePutProperty(pdisp, L"_BLOB", pvDat);
			}
		}
		pdisp->Release();
	}
	VariantClear(pvDat);
}

VOID teVariantChangeType(VARIANTARG * pvargDest, const VARIANTARG * pvarSrc, VARTYPE vt)
{
	VariantInit(pvargDest);
	if (pvarSrc->vt == VT_BSTR && ::SysStringLen(pvarSrc->bstrVal) == 18 && teStartsText(L"0x", pvarSrc->bstrVal)) {
		VARIANT v;
		if (GetLLFromVariant2(&v.llVal, (VARIANT *)pvarSrc)) {
			v.vt = VT_I8;
			if FAILED(VariantChangeType(pvargDest, &v, 0, vt)) {
				pvargDest->llVal = 0;
			}
			return;
		}
	}
	if (pvarSrc->vt == VT_DISPATCH || pvarSrc->vt == VT_EMPTY || pvarSrc->vt == VT_NULL || FAILED(VariantChangeType(pvargDest, pvarSrc, 0, vt))) {
		pvargDest->llVal = 0;
	}
}

BOOL teGetVariantTime(VARIANT *pv, DATE *pdt)
{
	VARIANT v;
	VariantInit(&v);
	VariantCopy(&v, pv);
	if (teVarIsNumber(pv)) {//UNIX EPOCH
		teVariantChangeType(&v, pv, VT_UI8);
		v.ullVal = v.ullVal * 10000 + 116444736000000000LL;
		teFileTimeToVariantTime((FILETIME *)&v.ullVal, pdt);
		return TRUE;
	}
	IDispatch *pdisp;
	if (GetDispatch(&v, &pdisp)) {
		VariantClear(&v);
		teExecMethod(pdisp, L"getVarDate", &v, 0, NULL);
		pdisp->Release();
	} else if (v.vt == VT_BSTR) {
		VariantClear(&v);
		teVariantChangeType(&v, pv, VT_DATE);
	}
	if (v.vt == VT_DATE && v.date) {
		*pdt = v.date;
		return TRUE;
	}
	VariantClear(&v);
	return FALSE;
}

BOOL teGetSystemTime(VARIANT *pv, SYSTEMTIME *pst)
{
	DATE dt;
	if (teGetVariantTime(pv, &dt)) {
		return ::VariantTimeToSystemTime(dt, pst);
	}
	return FALSE;
}

BOOL teVariantTimeToSystemTime(DATE dt, SYSTEMTIME *pst)
{
	if (::VariantTimeToSystemTime(dt, pst)) {
		FILETIME ft, ft2;
		if (::SystemTimeToFileTime(pst, &ft)) {
			if (::FileTimeToLocalFileTime(&ft, &ft2)) {
				return ::FileTimeToSystemTime(&ft2, pst);
			}
		}
	}
	return FALSE;
}

VOID teExtraLongPath(BSTR *pbs)
{
	int nLen = lstrlen(*pbs);
	if (nLen < MAX_PATH || StrChr(*pbs, '?')) {
		return;
	}
	BSTR bs = tePathMatchSpec(*pbs, L"\\\\*\\*") ? teSysAllocStringLen(L"\\\\?\\UNC", nLen + 7) : teSysAllocStringLen(L"\\\\?\\", nLen + 4);
	lstrcat(bs, *pbs);
	teSysFreeString(pbs);
	*pbs = bs;
}

HRESULT teInvokeAPI(TEDispatchApi *pApi, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo)
{
	int nArg = pDispParams ? pDispParams->cArgs : 0;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	teParam param[14] = { 0 };
	param[TE_EXCEPINFO].pExcepInfo = pExcepInfo;
	if (nArg-- >= pApi->nArgs) {
		VARIANT vParam[12] = { 0 };
		for (int i = nArg < 11 ? nArg : 11; i >= 0; --i) {
			VariantInit(&vParam[i]);
			if (i == (pApi->nStr1 % 10) || i == (pApi->nStr2 % 10) || i == (pApi->nStr3 % 10)) {
				teVariantChangeType(&vParam[i], &pDispParams->rgvarg[nArg - i], VT_BSTR);
				if (i == (pApi->nStr1 - 10) || i == (pApi->nStr2 - 10) || i == (pApi->nStr3 - 10)) {
					teExtraLongPath(&vParam[i].bstrVal);
				}
				param[i].bstrVal = vParam[i].bstrVal;
			} else {
				param[i].llVal = GetParamFromVariant(&pDispParams->rgvarg[nArg - i], &vParam[i]);
			}
		}
#ifdef _CHECK_HANDLELEAK
		HANDLE hProcess;
		DWORD dwHandle1, dwHandle2;
		if (hProcess = OpenProcess(PROCESS_QUERY_INFORMATION,FALSE, GetCurrentProcessId())) {
			GetProcessHandleCount(hProcess, &dwHandle1);
			++g_nLeak;
			if (g_nLeak == 19357) {
				Sleep(0);
			}
#endif
			pApi->fn(nArg, param, pDispParams, pVarResult);
#ifdef _CHECK_HANDLELEAK
			GetProcessHandleCount(hProcess, &dwHandle2);
			if (dwHandle2 > dwHandle1) {
				LPWSTR lpwstr = param[0].lpwstr;
				Sleep(0 * param[0].uintVal * param[1].uintVal * param[2].uintVal * param[3].uintVal * param[4].uintVal * g_nLeak);
			}
			CloseHandle(hProcess);
		}
#endif
		for (int i = nArg < 11 ? nArg : 11; i >= 0; --i) {
			if (i != pApi->nStr1 && i != pApi->nStr2 && i != pApi->nStr3) {
				teWriteBack(&pDispParams->rgvarg[nArg - i], &vParam[i]);
			} else {
				VariantClear(&vParam[i]);
			}
		}
	}
	return param[TE_API_RESULT].lVal;
}

LPWSTR teGetCommandLine()
{
	if (!g_bsCmdLine) {
		LPWSTR strCmdLine = GetCommandLine();
		int nSize = lstrlen(strCmdLine) + MAX_PROP;
		BSTR bsCmdLine = ::SysAllocStringLen(NULL, nSize);
		int j = 0;
		int i = 0;
		while (i < nSize && strCmdLine[j]) {
			if (StrCmpNI(&strCmdLine[j], L"/idlist,:", 9) == 0) {
				LONGLONG hData;
				DWORD dwProcessId;
				WCHAR sz[MAX_PATH + 2];
				swscanf_s(&strCmdLine[j + 9], L"%32[1234567890]s", sz, UINT(_countof(sz)));
				swscanf_s(sz, L"%lld", &hData);
				int k = lstrlen(sz);
				swscanf_s(&strCmdLine[j + k + 10], L"%32[1234567890]s", sz, UINT(_countof(sz)));
				swscanf_s(sz, L"%d", &dwProcessId);
				j += k + lstrlen(sz) + 11;
#ifdef _2000XP
				LPITEMIDLIST pidl = (LPITEMIDLIST)SHLockShared((HANDLE)hData, dwProcessId);
				if (pidl) {
					VARIANT v, v2;
					GetVarPathFromIDList(&v, pidl, SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
					SHUnlockShared(pidl);
					SHFreeShared((HANDLE)hData, dwProcessId);
					teVariantChangeType(&v2, &v, VT_BSTR);
					lstrcpy(sz, v2.bstrVal);
					PathQuoteSpaces(sz);
					lstrcpy(&bsCmdLine[i], sz);
					i += lstrlen(sz);
					VariantClear(&v);
					VariantClear(&v2);
					while (strCmdLine[j] && strCmdLine[j] != _T(',')) {
						++j;
					}
				}
#endif
			} else {
				bsCmdLine[i++] = strCmdLine[j++];
			}
		}
		g_bsCmdLine = teSysAllocStringLenEx(bsCmdLine, i);
		::SysFreeString(bsCmdLine);
	}
	return ::SysAllocStringLen(g_bsCmdLine, ::SysStringLen(g_bsCmdLine));
}

HRESULT teCLSIDFromProgID(__in LPCOLESTR lpszProgID, __out LPCLSID lpclsid)
{
	if (lstrcmpi(lpszProgID, L"ads") == 0) {
		*lpclsid = CLSID_ADODBStream;
		return S_OK;
	}
	if (lstrcmpi(lpszProgID, L"fso") == 0) {
		*lpclsid = CLSID_ScriptingFileSystemObject;
		return S_OK;
	}
	if (lstrcmpi(lpszProgID, L"sha") == 0) {
		*lpclsid = CLSID_Shell;
		return S_OK;
	}
	if (lstrcmpi(lpszProgID, L"wsh") == 0) {
		*lpclsid = CLSID_WScriptShell;
		return S_OK;
	}
	if (StrCmpNI(lpszProgID, L"new:", 4) == 0) {
		return CLSIDFromString(&lpszProgID[4], lpclsid);
	}
	return CLSIDFromProgID(lpszProgID, lpclsid);
}

HRESULT teCLSIDFromString(__in LPCOLESTR lpsz, __out LPCLSID lpclsid)
{
	HRESULT hr = CLSIDFromString(lpsz, lpclsid);
	if (hr == S_OK) {
		return S_OK;
	}
	return teCLSIDFromProgID(lpsz, lpclsid);
}

HWND FindTreeWindow(HWND hwnd)
{
	HWND hwnd1 = FindWindowExA(hwnd, 0, WC_TREEVIEWA, NULL);
	return hwnd1 ? hwnd1 : FindWindowExA(FindWindowExA(hwnd, 0, "NamespaceTreeControl", NULL), 0, WC_TREEVIEWA, NULL);
}

HWND teFindChildByClassA(HWND hwnd, LPCSTR lpClassA)
{
	HWND hwnd1 = NULL;
	while (hwnd1 = FindWindowEx(hwnd, hwnd1, NULL, NULL)) {
		CHAR szClassA[MAX_CLASS_NAME];
		GetClassNameA(hwnd1, szClassA, MAX_CLASS_NAME);
		if (lstrcmpiA(szClassA, lpClassA) == 0) {
			return hwnd1;
		}
		HWND hwnd2;
		if (hwnd2 = teFindChildByClassA(hwnd1, lpClassA)) {
			return hwnd2;
		}
	}
	return NULL;
}

BOOL teSetRect(HWND hwnd, int left, int top, int right, int bottom)
{
	RECT rc;
	GetWindowRect(hwnd, &rc);
	MoveWindow(hwnd, left, top, right - left, bottom - top, TRUE);
	return (rc.right - rc.left != right - left || rc.bottom - rc.top != bottom - top);
}

HRESULT teILFolderExists(LPITEMIDLIST pidl)
{
	LPCITEMIDLIST pidlPart;
	IShellFolder *pSF;
	HRESULT hr = SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart);
	if FAILED(hr) {
		return hr & MAXWORD;
	}
	SFGAOF sfAttr = SFGAO_FOLDER | SFGAO_FILESYSTEM;
	if FAILED(pSF->GetAttributesOf(1, &pidlPart, &sfAttr)) {
		sfAttr = 0;
	}
	hr = sfAttr & SFGAO_FILESYSTEM ? E_FAIL : S_FALSE;
	if ((sfAttr & SFGAO_FOLDER) && FAILED(hr)) {
		IShellFolder *pSF1;
		hr = pSF->BindToObject(pidlPart, NULL, IID_PPV_ARGS(&pSF1));
		if SUCCEEDED(hr) {
			IEnumIDList *peidl = NULL;
			hr = pSF1->EnumObjects(NULL, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN | SHCONTF_INCLUDESUPERHIDDEN | SHCONTF_NAVIGATION_ENUM, &peidl);
			if (peidl) {
				LPITEMIDLIST pidlPart = NULL;
				hr = peidl->Next(1, &pidlPart, NULL);
				if (hr == S_FALSE) {
					hr = S_OK;
				}
				teCoTaskMemFree(pidlPart);
				peidl->Release();
			}
			if (hr == E_INVALID_PASSWORD || hr == E_CANCELLED) {
				hr &= MAXWORD;
			}
			pSF1->Release();
		}
	}
	pSF->Release();
	return hr;
}

HRESULT tePathIsDirectory2(LPWSTR pszPath, int iUseFS)
{
	if (!pszPath) {
		return E_FAIL;
	}
	if (!(iUseFS & 2)) {
		WCHAR pszDrive[0x80];
		lstrcpyn(pszDrive, pszPath, 4);
		if (pszDrive[0] >= 'A' && pszDrive[1] == ':' && pszDrive[2] == '\\') {
			if (!GetVolumeInformation(pszDrive, NULL, 0, NULL, NULL, NULL, pszDrive, ARRAYSIZE(pszDrive))) {
				return E_NOT_READY;
			}
		}
	}
	LPITEMIDLIST pidl = iUseFS ? teILCreateFromPathEx(pszPath) : teILCreateFromPath(pszPath);
	if (pidl) {
		HRESULT hr = teILFolderExists(pidl);
		teCoTaskMemFree(pidl);
		return hr;
	}
	return E_FAIL;
}

VOID teReleaseExists(TEExists *pExists)
{
	if (::InterlockedDecrement(&pExists->cRef) == 0) {
		CloseHandle(pExists->hEvent);
		delete [] pExists;
	}
}

static void threadExists(void *args)
{
	::CoInitialize(NULL);
	try {
		TEExists *pExists = (TEExists *)args;
		pExists->hr = tePathIsDirectory2(pExists->pszPath, pExists->iUseFS);
		SetEvent(pExists->hEvent);
		teReleaseExists(pExists);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"threadExists";
#endif
	}
	::CoUninitialize();
	::_endthread();
}

HRESULT tePathIsDirectory(LPWSTR pszPath, int dwms, int iUseFS)
{
	if (g_dwMainThreadId == GetCurrentThreadId()) {
		if (dwms <= 0) {
			dwms = g_param[TE_NetworkTimeout];
		}
#ifdef _CHECK_HANDLELEAK
		dwms = 0;
#endif
		if (dwms) {
			TEExists *pExists = new TEExists[1];
			pExists->cRef = 2;
			pExists->hr = E_ABORT;
			pExists->pszPath = pszPath;
			pExists->hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
			pExists->iUseFS = iUseFS;
			if (_beginthread(&threadExists, 0, pExists) != -1) {
				WaitForSingleObject(pExists->hEvent, dwms);
			} else {
				--pExists->cRef;
			}
			HRESULT hr = pExists->hr;
			teReleaseExists(pExists);
			return hr;
		}
	}
	return tePathIsDirectory2(pszPath, iUseFS);
}

LPITEMIDLIST teILCreateFromPath1(LPWSTR pszPath)
{
	LPITEMIDLIST pidl = NULL;
	BSTR bstr = NULL;
	int n;

	if (pszPath && !teStartsText(L"\\\\\\", pszPath)) {
		BSTR bsPath2 = NULL;
		if (pszPath[0] == _T('"')) {
			bsPath2 = teSysAllocStringLen(pszPath, lstrlen(pszPath) + 1);
			PathUnquoteSpaces(bsPath2);
			pszPath = bsPath2;
		}
		BSTR bsPath3 = NULL;
		if (teIsSearchFolder(pszPath)) {
			teGetSearchArg(&bsPath3, pszPath, L"&crumb=location:");
			ISearchFolderItemFactory *psfif = NULL;
			IQueryParserManager *pqpm = NULL;
			if SUCCEEDED(teCreateInstance(CLSID_SearchFolderItemFactory, NULL, NULL, IID_PPV_ARGS(&psfif))) {
				if SUCCEEDED(CoCreateInstance(__uuidof(QueryParserManager), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pqpm))) {
					if (pidl = teILCreateFromPath(bsPath3)) {
						teSysFreeString(&bsPath3);
						IShellItem *psi = NULL;
						if SUCCEEDED(SHCreateShellItem(NULL, NULL, pidl, &psi)) {
							teILFreeClear(&pidl);
							IShellItemArray *psia;
							if (
#ifdef _2000XP
								_SHCreateShellItemArrayFromShellItem &&
#endif
								SUCCEEDED(_SHCreateShellItemArrayFromShellItem(psi, IID_PPV_ARGS(&psia)))) {
								teGetSearchArg(&bsPath3, pszPath, L"crumb=");
								psfif->SetScope(psia);
								psia->Release();
								IQueryParser *pqp;
								if SUCCEEDED(pqpm->CreateLoadedParser(L"SystemIndex", LOCALE_USER_DEFAULT, IID_PPV_ARGS(&pqp))) {
									IQuerySolution *pqs = NULL;
									ICondition *pCondition;
									if (SUCCEEDED(pqp->Parse(bsPath3, NULL, &pqs)) && SUCCEEDED(pqs->GetQuery(&pCondition, NULL))) {
										psfif->SetDisplayName(bsPath3);
										psfif->SetCondition(pCondition);
										psfif->GetIDList(&pidl);
									}
									pqp->Release();
									teSysFreeString(&bsPath3);
									SafeRelease(&pqs);
								}
							}
							SafeRelease(&psi);
						} else {
							teILFreeClear(&pidl);
						}
					}
				}
				SafeRelease(&pqpm);
			}
			SafeRelease(&psfif);
		} else if (tePathMatchSpec(pszPath, L"*\\..\\*;*\\..;*\\.\\*;*\\.;*%*%*")) {
			DWORD dwLen = lstrlen(pszPath) + MAX_PATH;
			bsPath3 = ::SysAllocStringLen(NULL, dwLen);
			if (PathSearchAndQualify(pszPath, bsPath3, dwLen)) {
				pszPath = bsPath3;
			}
		} else if (lstrlen(pszPath) == 1 && pszPath[0] >= 'A') {
			bsPath3 = teMultiByteToWideChar(CP_UTF8, "?:\\", -1);
			bsPath3[0] = pszPath[0];
			pszPath = bsPath3;
		}
		if (!pidl && GetCSIDLFromPath(&n, pszPath)) {
			pidl = ::ILClone(g_pidls[n]);
			pszPath = NULL;
		}
		if (pszPath) {
			if (!pidl) {
				pidl = teILCreateFromPathEx(pszPath);
				if (pidl) {
					if (tePathIsNetworkPath(pszPath) && PathIsRoot(pszPath) && FAILED(tePathIsDirectory(pszPath, 0, 3))) {
						teILFreeClear(&pidl);
					}
				} else if (PathGetDriveNumber(pszPath) >= 0 && !PathIsRoot(pszPath)) {
					WCHAR pszDrive[4];
					lstrcpyn(pszDrive, pszPath, 4);
					int n = GetDriveType(pszDrive);
					if (n == DRIVE_NO_ROOT_DIR && SUCCEEDED(tePathIsDirectory(pszDrive, 0, 3))) {
						pidl = teILCreateFromPathEx(pszPath);
					}
				} else if (tePathMatchSpec(pszPath, L"\\\\*\\*")) {
					WIN32_FIND_DATA wfd;
					HANDLE hFind = FindFirstFile(pszPath, &wfd);
					if (hFind != INVALID_HANDLE_VALUE) {
						FindClose(hFind);
						LPWSTR lpDelimiter = StrChr(&pszPath[2], '\\');
						BSTR bsServer = teSysAllocStringLen(pszPath, int(lpDelimiter - pszPath));
						LPITEMIDLIST pidlServer = teILCreateFromPathEx(bsServer);
						if (pidlServer) {
							pidl = teILCreateFromPath2(pidlServer, &lpDelimiter[1], g_hwndMain);
							teCoTaskMemFree(pidlServer);
						}
						::SysFreeString(bsServer);
					}
				}
			}
			if (pidl == NULL && !teIsFileSystem(pszPath)) {
				for (int i = 0; i < MAX_CSIDL; ++i) {
					if (g_pidls[i]) {
						if (!lstrcmpi(pszPath, g_bsPidls[i])) {
							pidl = ILClone(g_pidls[i]);
							break;
						}
					}
				}
				if (!pidl) {
					pidl = teILCreateFromPath2(g_pidls[CSIDL_DESKTOP], pszPath, NULL);
					if (!pidl) {
						pidl = teILCreateFromPath2(g_pidls[CSIDL_DRIVES], pszPath, NULL);
						if (!pidl) {
							pidl = teILCreateFromPath2(g_pidls[CSIDL_USER], pszPath, NULL);
							if (!pidl && PathMatchSpec(pszPath, L"::{*")) {
								int nSize = lstrlen(pszPath) + 6;
								BSTR bsPath4 = teSysAllocStringLen(L"shell:", nSize);
								lstrcat(bsPath4, pszPath);
								pidl = teILCreateFromPathEx(bsPath4);
								::SysFreeString(bsPath4);
								if SUCCEEDED(teGetDisplayNameFromIDList(&bsPath4, pidl, SHGDN_INFOLDER)) {
									::SysFreeString(bsPath4);
								} else {
									teILFreeClear(&pidl);
								}
							}
						}
					}
				}
			}
		}
		teSysFreeString(&bsPath3);
		teSysFreeString(&bsPath2);
	}
	return pidl;
}

static void threadILCreate(void *args)
{
	::CoInitialize(NULL);
	try {
		TEILCreate *pILC = (TEILCreate *)args;
		pILC->pidlResult = teILCreateFromPath1(pILC->pszPath);
		SetEvent(pILC->hEvent);
		teReleaseILCreate(pILC);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"threadILCreate";
#endif
	}
	::CoUninitialize();
	::_endthread();
}

LPITEMIDLIST teILCreateFromPath0(LPWSTR pszPath, BOOL bForceLimit)
{
	DWORD dwms = g_pTE ? g_param[TE_NetworkTimeout] : 2000;
	if (bForceLimit && (!dwms || dwms > 500)) {
		dwms = 500;
	}
#ifdef _CHECK_HANDLELEAK
	dwms = 0;
#endif
	if (dwms && (bForceLimit || g_dwMainThreadId == GetCurrentThreadId())) {
		TEILCreate *pILC = new TEILCreate[1];
		pILC->cRef = 2;
		pILC->pidlResult = NULL;
		pILC->pszPath = pszPath;
		pILC->hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
		LPITEMIDLIST pidl = NULL;
		if (_beginthread(&threadILCreate, 0, pILC) != -1) {
			if (WaitForSingleObject(pILC->hEvent, dwms) != WAIT_TIMEOUT) {
				pidl = pILC->pidlResult;
				if (pidl) {
					pILC->pidlResult = NULL;
				} else {
					pidl = teILCreateFromPath1(pszPath);
				}
			}
		} else {
			--pILC->cRef;
			pidl = teILCreateFromPath1(pszPath);
		}
		teReleaseILCreate(pILC);
		return pidl;
	}
	return teILCreateFromPath1(pszPath);
}

LPITEMIDLIST teILCreateFromPath(LPWSTR pszPath)
{
	return teILCreateFromPath0(pszPath, FALSE);
}

BOOL teCreateItemFromPath(LPWSTR pszPath, IShellItem **ppSI)
{
	BOOL Result = FALSE;
	LPITEMIDLIST pidl = teILCreateFromPath(const_cast<LPWSTR>(pszPath));
	if (pidl) {
		Result = SUCCEEDED(_SHCreateItemFromIDList(pidl, IID_PPV_ARGS(ppSI)));
		teCoTaskMemFree(pidl);
	}
	return Result;
}

BOOL GetDestAndName(LPWSTR pszPath, IShellItem **ppSI, LPWSTR *ppsz)
{
	if (teCreateItemFromPath(pszPath, ppSI)) {
		*ppsz = NULL;
		return TRUE;
	}
	BOOL Result = FALSE;
	BSTR bs = ::SysAllocString(pszPath);
	if (PathRemoveFileSpec(bs)) {
		Result = teCreateItemFromPath(bs, ppSI);
		*ppsz = PathFindFileName(pszPath);
	}
	::SysFreeString(bs);
	return Result;
}

int teSHFileOperation(LPSHFILEOPSTRUCT pFOS)
{
	HRESULT hr = E_NOTIMPL;
	LPWSTR pszFrom = const_cast<LPWSTR>(pFOS->pFrom);
	if (pszFrom && !(pFOS->fFlags & FOF_WANTMAPPINGHANDLE)) {
		IFileOperation *pFO;
		if SUCCEEDED(teCreateInstance(CLSID_FileOperation, NULL, NULL, IID_PPV_ARGS(&pFO))) {
			if SUCCEEDED(pFO->SetOperationFlags(pFOS->fFlags & ~FOF_MULTIDESTFILES)) {
				try {
					pFO->SetOwnerWindow(pFOS->hwnd);
					LPWSTR pszTo = const_cast<LPWSTR>(pFOS->pTo);
					IShellItem *psiFrom = NULL;
					IShellItem *psiTo = NULL;
					LPWSTR pszName = NULL;
					while (*pszFrom) {
						if (teCreateItemFromPath(pszFrom, &psiFrom)) {
							if (pFOS->wFunc == FO_DELETE) {
								hr = pFO->DeleteItem(psiFrom, NULL);
							} else if (pszTo && *pszTo) {
								if (pFOS->wFunc == FO_RENAME) {
									hr = pFO->RenameItem(psiFrom, pszTo, NULL);
								} else if (psiTo || GetDestAndName(pszTo, &psiTo, &pszName)) {
									if (pFOS->wFunc == FO_COPY) {
										hr = pFO->CopyItem(psiFrom, psiTo, pszName, NULL);
									} else if (pFOS->wFunc == FO_MOVE) {
										if (!pszName || PathFindFileName(pszFrom) - pszFrom != pszName - pszTo || StrCmpNI(pszFrom, pszTo, (int)(pszName - pszTo))) {
											hr = pFO->MoveItem(psiFrom, psiTo, pszName, NULL);
										} else {
											hr = pFO->RenameItem(psiFrom, pszName, NULL);
										}
									}
								}
								if (pFOS->fFlags & FOF_MULTIDESTFILES) {
									SafeRelease(&psiTo);
									while (*(pszTo++));
								}
							}
							psiFrom->Release();
						}
						while (*(pszFrom++));
					}
					SafeRelease(&psiTo);
					if SUCCEEDED(hr) {
						hr = pFO->PerformOperations();
						pFO->GetAnyOperationsAborted(&pFOS->fAnyOperationsAborted);
					}
				} catch (...) {
					g_nException = 0;
#ifdef _DEBUG
					g_strException = L"SHFileOperation";
#endif
				}
			}
			pFO->Release();
		}
	}
	return hr == E_NOTIMPL ? ::SHFileOperation(pFOS) : hr;
}

static void threadFileOperation(void *args)
{
	::CoInitialize(NULL);
	LPSHFILEOPSTRUCT pFO = (LPSHFILEOPSTRUCT)args;
	::InterlockedIncrement(&g_nThreads);
	try {
		teSHFileOperation(pFO);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"threadFileOperation";
#endif
	}
	::InterlockedDecrement(&g_nThreads);
	::SysFreeString(const_cast<BSTR>(pFO->pTo));
	::SysFreeString(const_cast<BSTR>(pFO->pFrom));
	delete [] pFO;
	SetTimer(g_hwndMain, TET_EndThread, 100, teTimerProc);
	::CoUninitialize();
	::_endthread();
}

int teGetMenuString(BSTR *pbs, HMENU hMenu, UINT uIDItem, BOOL fByPosition)
{
	MENUITEMINFO mii = { sizeof(MENUITEMINFO), MIIM_STRING };
	GetMenuItemInfo(hMenu, uIDItem, fByPosition, &mii);
	if (mii.cch) {
		*pbs = SysAllocStringLen(NULL, mii.cch++);
		mii.dwTypeData = *pbs;
		GetMenuItemInfo(hMenu, uIDItem, fByPosition, &mii);
	}
	return mii.cch;
}

VOID teMenuText(LPWSTR sz)
{
	int i = 0, j = 0;
	while (sz[i]) {
		if (sz[i] == '&') {
			++i;
			continue;
		}
		if (sz[i] == '(') {
			sz[j++] = NULL;
			break;
		}
		sz[j++] = sz[i++];
	}
}

int tePos(WCHAR wc, WCHAR *sz)
{
	int i = 0;
	while (sz[i]) {
		if (sz[i] == wc) {
			return i;
		}
		++i;
	}
	return -1;
}

VOID ToMinus(BSTR *pbs)
{
	int nLen = SysStringByteLen(*pbs) + sizeof(WCHAR);
	BSTR bs = SysAllocStringByteLen(NULL, nLen);
	bs[0] = L'-';
	if (*pbs) {
		::CopyMemory(&bs[1], *pbs, nLen);
		teSysFreeString(pbs);
	}
	*pbs = bs;
}

int SizeOfvt(VARTYPE vt)
{
	switch (LOBYTE(vt)) {
		case VT_BOOL:
			return sizeof(VARIANT_BOOL);
		case VT_I1:
		case VT_UI1:
			return sizeof(CHAR);
		case VT_I2:
		case VT_UI2:
			return sizeof(SHORT);
		case VT_I4:
		case VT_UI4:
			return sizeof(LONG);
		case VT_I8:
		case VT_UI8:
			return sizeof(LONGLONG);
		case VT_PTR:
		case VT_LPWSTR:
		case VT_BSTR:
		case VT_LPSTR:
			return sizeof(HANDLE);
		case VT_R4:
			return sizeof(FLOAT);
		case VT_R8:
		case VT_FILETIME:
		case VT_DATE:
			return sizeof(double);
		case VT_VARIANT:
			return sizeof(VARIANT);
	}//end_switch
	return 0;
}

void teCopyMenu(HMENU hDest, HMENU hSrc, UINT fState)
{
	int n = GetMenuItemCount(hSrc);
	while (--n >= 0) {
		MENUITEMINFO mii = { sizeof(MENUITEMINFO), MIIM_STRING };
		GetMenuItemInfo(hSrc, n, TRUE, &mii);
		BSTR bsTypeData = SysAllocStringLen(NULL, mii.cch++);
		mii.dwTypeData = bsTypeData;
		mii.fMask  = MIIM_ID | MIIM_TYPE | MIIM_SUBMENU | MIIM_STATE;
		GetMenuItemInfo(hSrc, n, TRUE, &mii);
		HMENU hSubMenu = mii.hSubMenu;
		if (hSubMenu) {
			mii.hSubMenu = CreateMenu();
		}
		mii.fState = fState;
		InsertMenuItem(hDest, 0, TRUE, &mii);
		if (hSubMenu) {
			teCopyMenu(mii.hSubMenu, hSubMenu, fState);
		}
		::SysFreeString(bsTypeData);
	}
}

CteShellBrowser* SBfromhwnd(HWND hwnd)
{
	for (size_t i = 0; i < g_ppSB.size(); ++i) {
		CteShellBrowser *pSB = g_ppSB[i];
		if (pSB->m_hwnd == hwnd || IsChild(pSB->m_hwnd, hwnd)) {
			return pSB;
		}
	}
	return NULL;
}

CteTabCtrl* TCfromhwnd(HWND hwnd)
{
	for (size_t i = 0; i < g_ppTC.size(); ++i) {
		CteTabCtrl *pTC = g_ppTC[i];
		if (!pTC) {
			continue;
		}
		if (pTC->m_hwnd == hwnd || pTC->m_hwndStatic == hwnd || pTC->m_hwndButton == hwnd) {
			return pTC;
		}
	}
	return NULL;
}

CteTreeView* TVfromhwnd2(HWND hwnd)
{
	for (size_t i = 0; i < g_ppTV.size(); ++i) {
		CteTreeView *pTV = g_ppTV[i];
		HWND hwndTV = pTV->m_hwnd;
		if (hwndTV == hwnd || IsChild(hwndTV, hwnd)) {
			return pTV->m_bMain ? pTV : NULL;
		}
	}
	return NULL;
}

CteTreeView* TVfromhwnd(HWND hwnd)
{
	for (size_t i = 0; i < g_ppSB.size(); ++i) {
		CteShellBrowser *pSB = g_ppSB[i];
		HWND hwndTV = pSB->m_pTV->m_hwnd;
		if (hwndTV == hwnd || IsChild(hwndTV, hwnd)) {
			return pSB->m_pTV->m_bMain ? pSB->m_pTV : NULL;
		}
	}
	return TVfromhwnd2(hwnd);
}

void CheckChangeTabSB(HWND hwnd)
{
	if (g_pTC) {
		CteShellBrowser *pSB = SBfromhwnd(hwnd);
		if (!pSB) {
			CteTreeView *pTV = TVfromhwnd(hwnd);
			if (pTV) {
				pSB = pTV->m_pFV;
			}
		}
		if (pSB) {
			if (pSB->m_pTC->SetDefault()) {
				pSB->m_pTC->TabChanged(FALSE);
			}
		}
	}
}

VOID CheckChangeTabTC(HWND hwnd)
{
	CteTabCtrl *pTC = TCfromhwnd(hwnd);
	if (pTC && pTC->m_bVisible) {
		if (pTC->SetDefault()) {
			pTC->TabChanged(FALSE);
		}
	}
}

VOID AdjustIDList(LPITEMIDLIST *ppidllist, int nCount)
{
	if (ppidllist == NULL || nCount <= 0) {
		return;
	}
	if (ppidllist[0]) {
#ifdef _2000XP
		if (g_bUpperVista || !ILIsEqual(ppidllist[0], g_pidls[CSIDL_RESULTSFOLDER])) {
			return;
		}
		for (int i = nCount; i > 1; --i) {
			teILCloneReplace(&ppidllist[i], ILCombine(ppidllist[0], ppidllist[i]));
		}
		teILFreeClear(&ppidllist[0]);
#else
		return;
#endif
	}
	ppidllist[0] = ::ILClone(ppidllist[1]);
	ILRemoveLastID(ppidllist[0]);
	int nCommon = ILGetCount(ppidllist[0]);
	int nBase = nCommon;

	for (int i = nCount; i > 1 && nCommon; --i) {
		LPITEMIDLIST pidl = ::ILClone(ppidllist[i]);
		ILRemoveLastID(pidl);
		int nLevel = ILGetCount(pidl);
		while (nLevel > nCommon) {
			ILRemoveLastID(pidl);
			--nLevel;
		}
		while (nLevel < nCommon) {
			ILRemoveLastID(ppidllist[0]);
			--nCommon;
		}
		while (nCommon > 0 && !ILIsEqual(pidl, ppidllist[0])) {
			ILRemoveLastID(pidl);
			ILRemoveLastID(ppidllist[0]);
			--nCommon;
		}
		teCoTaskMemFree(pidl);
	}

	if (nCommon) {
		for (int i = nCount; i > 0; --i) {
			LPITEMIDLIST pidl = ppidllist[i];
			for (int j = nCommon; --j >= 0;) {
				pidl = ILGetNext(pidl);
			}
			teILCloneReplace(&ppidllist[i], pidl);
		}
	}
}

LPITEMIDLIST* IDListFormDataObj(IDataObject *pDataObj, long *pnCount)
{
	LPITEMIDLIST *ppidllist = NULL;
	*pnCount = 0;
	if (pDataObj) {
		STGMEDIUM Medium;
		if (pDataObj->GetData(&IDLISTFormat, &Medium) == S_OK) {
			try {
				CIDA *pIDA = (CIDA *)::GlobalLock(Medium.hGlobal);
				if (pIDA) {
					*pnCount = pIDA->cidl;
					ppidllist = new LPITEMIDLIST[*pnCount + 1];
					for (int i = *pnCount; i >= 0; --i) {
						ppidllist[i] = ::ILClone((LPCITEMIDLIST)((BYTE *)pIDA + pIDA->aoffset[i]));
					}
				}
			} catch(...) {
				if (ppidllist) {
					delete [] ppidllist;
					ppidllist = NULL;
				}
				*pnCount = 0;
			}
			::GlobalUnlock(Medium.hGlobal);
			::ReleaseStgMedium(&Medium);
			if (ppidllist) {
				return ppidllist;
			}
		}
		if (pDataObj->GetData(&HDROPFormat, &Medium) == S_OK) {
			*pnCount = DragQueryFile((HDROP)Medium.hGlobal, (UINT)-1, NULL, 0);
			ppidllist = new LPITEMIDLIST[*pnCount + 1];
			ppidllist[0] = NULL;
			for (int i = *pnCount; --i >= 0;) {
				BSTR bsPath;
				teDragQueryFile((HDROP)Medium.hGlobal, i, &bsPath);
				ppidllist[i + 1] = teILCreateFromPath(bsPath);
				::SysFreeString(bsPath);
			}
			::ReleaseStgMedium(&Medium);
		}
	}
	return ppidllist;
}

#ifdef _USE_APIHOOK

HTHEME WINAPI teOpenNcThemeData(HWND hWnd, LPCWSTR pszClassList)
{
	if (lstrcmpi(pszClassList, L"ScrollBar") == 0) {
		hWnd = NULL;
		pszClassList = L"Explorer::ScrollBar";
	}
	return _OpenNcThemeData(hWnd, pszClassList);
}

LSTATUS APIENTRY teRegQueryValueExW(HKEY hKey, LPCWSTR lpValueName, LPDWORD lpReserved, LPDWORD lpType, LPBYTE lpData, LPDWORD lpcbData)
{
	return _RegQueryValueExW(hKey, lpValueName, lpReserved, lpType, lpData, lpcbData);
}

LSTATUS APIENTRY teRegQueryValueW(HKEY hKey, LPCWSTR lpSubKey, LPWSTR lpData, PLONG lpcbData)
{
	return _RegQueryValueW(hKey, lpSubKey, lpData, lpcbData);
}

DWORD WINAPI teGetSysColor(int nIndex)
{
	return _GetSysColor(nIndex);
}
#endif

#ifdef _2000XP
HRESULT STDAPICALLTYPE teGetIDListFromObjectXP(IUnknown *punk, PIDLIST_ABSOLUTE *ppidl)
{
	IPersistFolder2 *pPF2;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pPF2))) {
		pPF2->GetCurFolder(ppidl);
		pPF2->Release();
		return *ppidl ? S_OK : E_FAIL;
	}
	IPersistIDList *pPI;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pPI))) {
		pPI->GetIDList(ppidl);
		pPI->Release();
		return *ppidl ? S_OK : E_FAIL;
	}
	FolderItem *pFI;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pFI))) {
		BSTR bstr;
		if SUCCEEDED(pFI->get_Path(&bstr)) {
			*ppidl = teILCreateFromPath(bstr);
			::SysFreeString(bstr);
		}
		pFI->Release();
		return *ppidl ? S_OK : E_FAIL;
	}
	return E_NOTIMPL;
}

HRESULT STDAPICALLTYPE teSHCreateItemFromIDListXP(PCIDLIST_ABSOLUTE pidl, REFIID riid, void **ppv)
{
	return SHCreateShellItem(NULL, NULL, pidl, (IShellItem **)ppv);
}

LPITEMIDLIST teILCreateResultsXP(LPITEMIDLIST pidl)
{
	LPITEMIDLIST pidl2 = NULL;
	LPCITEMIDLIST pidlLast;
	IShellFolder *pSF;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlLast)) {
		SFGAOF sfAttr = SFGAO_FILESYSTEM | SFGAO_FILESYSANCESTOR | SFGAO_STORAGE | SFGAO_STREAM;
		if (SUCCEEDED(pSF->GetAttributesOf(1, &pidlLast, &sfAttr)) &&
			(sfAttr & (SFGAO_FILESYSTEM | SFGAO_FILESYSANCESTOR | SFGAO_STORAGE | SFGAO_STREAM))) {
			UINT uSize = ::ILGetSize(pidl) + 28;
			pidl2 = (LPITEMIDLIST)::CoTaskMemAlloc(uSize + sizeof(USHORT));
			::ZeroMemory(pidl2, uSize + sizeof(USHORT));

			UINT uSize2 = ::ILGetSize(pidlLast);
			::CopyMemory(pidl2, pidlLast, uSize2);
			*(PUSHORT)pidl2 = uSize - 2;
			UINT uSize3 = uSize - uSize2 - 28;
			PBYTE p = (PBYTE)pidl2;
			*(PUSHORT)&p[uSize2 - 2] = uSize3 + 28;
			*(PDWORD)&p[uSize2 + 2] = 0xbeef0005;
			::CopyMemory(&p[uSize2 + 22], pidl, uSize3);
			*(PUSHORT)&p[uSize - 4] = uSize2 - 2;
			if (teCompareSFClass(pSF, &CLSID_ShellFSFolder)) {
				*(PUSHORT)&p[uSize2 + 24 + uSize3] = *(PUSHORT)&p[uSize2 - 4];
			}
			STRRET strret;
			if SUCCEEDED(pSF->GetDisplayNameOf(pidl2, SHGDN_NORMAL, &strret)) {
				if (strret.uType == STRRET_WSTR) {
					teCoTaskMemFree(strret.pOleStr);
				}
			} else {
				teILFreeClear(&pidl2);
			}
		}
		pSF->Release();
	}
	return pidl2;
}

HRESULT STDAPICALLTYPE tePSPropertyKeyFromStringXP(__in LPCWSTR pszString,  __out PROPERTYKEY *pkey)
{
	if (lstrcmpi(pszString, g_szTotalFileSizeCodeXP) == 0) {
		*pkey = PKEY_TotalFileSize;
		return S_OK;
	}
	if (lstrcmpi(pszString, g_szLabelCodeXP) == 0) {
		*pkey = PKEY_Contact_Label;
		return S_OK;
	}
	HRESULT hr = E_NOTIMPL;
	ULONG chEaten = 0;
	IPropertyUI *pPUI;
	if SUCCEEDED(teCreateInstance(CLSID_PropertiesUI, NULL, NULL, IID_PPV_ARGS(&pPUI))) {
		hr = pPUI->ParsePropertyName(pszString, &pkey->fmtid, &pkey->pid, &chEaten);
		pPUI->Release();
	}
	return hr;
}

HRESULT STDAPICALLTYPE teVariantToVariantXP(VARIANTARG *pvSrc, VARIANTARG *pvDest)
{
	return VariantCopy(pvDest, pvSrc);
}

VOID teStripAmp(LPWSTR lpstr)
{
	LPWSTR lp1 = lpstr;
	WCHAR wc;
	while (wc = *lp1 = *(lpstr++)) {
		if (wc == '(') {
			*lp1 = NULL;
			return;
		}
		if (wc != '&') {
			++lp1;
		}
	}
}

#endif
#ifdef _W2000
HRESULT STDAPICALLTYPE teSHGetImageList2000(int iImageList, REFIID riid, void **ppvObj)
{
	SHFILEINFO sfi;
	UINT uFlags = 0;
	switch (iImageList) {
		case SHIL_LARGE:
			uFlags = SHGFI_SYSICONINDEX | SHGFI_LARGEICON;
			break;
		case SHIL_SMALL:
			uFlags = SHGFI_SYSICONINDEX | SHGFI_SMALLICON;
			break;
		case SHIL_SYSSMALL:
			uFlags = SHGFI_SYSICONINDEX | SHGFI_SHELLICONSIZE;
			break;
		default:
			return E_NOTIMPL;
	}//end_switch
	*(DWORD_PTR *)ppvObj = SHGetFileInfo(NULL, 0, &sfi, sizeof(sfi), uFlags);
	return S_OK;
}

#endif

HMODULE teCreateInstanceV(LPVARIANTARG pvDllFile, LPVARIANTARG pvClass, REFIID riid, PVOID *ppvObj)
{
	VARIANT v;
	CLSID clsid;
	teVariantChangeType(&v, pvClass, VT_BSTR);
	teCLSIDFromString(v.bstrVal, &clsid);
	VariantClear(&v);
	teVariantChangeType(&v, pvDllFile, VT_BSTR);
	HMODULE hDll = NULL;
	teCreateInstance(clsid, v.bstrVal, &hDll, riid, ppvObj);
	VariantClear(&v);
	return hDll;
}

BOOL teGetIDListFromObject(IUnknown *punk, LPITEMIDLIST *ppidl)
{
	*ppidl = NULL;
	if (!punk) {
		return FALSE;
	}
#ifdef _CHECK_HANDLELEAK
	if SUCCEEDED(teGetIDListFromObjectXP(punk, ppidl)) {
		return TRUE;
	}
#endif
	if SUCCEEDED(_SHGetIDListFromObject(punk, ppidl)) {
		return TRUE;
	}
	IShellBrowser *pSB;
	if SUCCEEDED(IUnknown_QueryService(punk, SID_SShellBrowser, IID_PPV_ARGS(&pSB))) {
		IShellView *pSV;
		if SUCCEEDED(pSB->QueryActiveShellView(&pSV)) {
			if FAILED(_SHGetIDListFromObject(pSV, ppidl)) {
#ifdef _W2000
				//Windows 2000
				IDataObject *pDataObj;
				if SUCCEEDED(pSV->GetItemObject(SVGIO_ALLVIEW, IID_PPV_ARGS(&pDataObj))) {
					long nCount;
					LPITEMIDLIST *pidllist = IDListFormDataObj(pDataObj, &nCount);
					*ppidl = pidllist[0];
					while (--nCount >= 1) {
						if (pidllist) {
							teCoTaskMemFree(pidllist[nCount]);
						}
					}
					delete[] pidllist;
					pDataObj->Release();
				}
#endif
			}
			pSV->Release();
		}
		pSB->Release();
	}
	return *ppidl != NULL;
}

BOOL teGetIDListFromObjectForFV(IUnknown *punk, LPITEMIDLIST *ppidl)
{
	*ppidl = NULL;
	if (!punk) {
		return FALSE;
	}
	CteFolderItem *pid1 = NULL;
	if SUCCEEDED(punk->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) {
		*ppidl = ILClone(pid1->GetAlt());
		pid1->Release();
	}
	return *ppidl ? TRUE : teGetIDListFromObject(punk, ppidl);
}

VOID teGetPath(BSTR *pbs, IUnknown *punk)
{
	*pbs = NULL;
	FolderItem *pid;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pid))) {
		pid->get_Path(pbs);
	}
}


BOOL teILIsEqual(IUnknown *punk1, IUnknown *punk2)
{
	BOOL bResult = FALSE;
	if (punk1 && punk2) {
		LPITEMIDLIST pidl1, pidl2;
		if (teGetIDListFromObject(punk1, &pidl1)) {
			if (teGetIDListFromObject(punk2, &pidl2)) {
				BSTR bs1, bs2;
				teGetPath(&bs1, punk1);
				teGetPath(&bs2, punk2);
				if (::ILIsEqual(pidl1, g_pidls[CSIDL_RESULTSFOLDER]) || ILIsEmpty(pidl1) || teIsSearchFolder(bs1) || teIsSearchFolder(bs2)) {
					bResult = lstrcmpi(bs1, bs2) == 0;
				} else {
					bResult = ::ILIsEqual(pidl1, pidl2);
				}
				teSysFreeString(&bs2);
				teSysFreeString(&bs1);
				teCoTaskMemFree(pidl2);
			}
			teCoTaskMemFree(pidl1);
		}
	}
	return bResult;
}

int GetIntFromVariant(VARIANT *pv)
{
	if (pv) {
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return GetIntFromVariant(pv->pvarVal);
		}
		if (pv->vt == VT_I4) {
			return pv->lVal;
		}
		if (pv->vt == VT_R8) {
			return (int)(LONGLONG)pv->dblVal;
		}
		if (pv->vt != VT_DISPATCH) {
			VARIANT vo;
			VariantInit(&vo);
			if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I4)) {
				return vo.lVal;
			}
			if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8)) {
				return (int)vo.llVal;
			}
			if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_R8)) {
				return (int)(LONGLONG)vo.dblVal;
			}
		}
	}
	return 0;
}

BOOL GetBoolFromVariant(VARIANT *pv)
{
	if (pv) {
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return GetBoolFromVariant(pv->pvarVal);
		}
		return pv->vt == VT_DISPATCH || pv->vt == VT_UNKNOWN || GetIntFromVariant(pv);
	}
	return 0;
}

int GetIntFromVariantClear(VARIANT *pv)
{
	int i = GetIntFromVariant(pv);
	VariantClear(pv);
	return i;
}

VOID teGetFileTimeFromItem(IShellFolder2 *pSF2, LPCITEMIDLIST pidl, const SHCOLUMNID *pscid, FILETIME *pft)
{
	VARIANT v;
	VariantInit(&v);
	SYSTEMTIME SysTime;
	if (SUCCEEDED(pSF2->GetDetailsEx(pidl, pscid, &v)) && v.vt == VT_DATE) {
		if (::VariantTimeToSystemTime(v.date, &SysTime)) {
			::SystemTimeToFileTime(&SysTime, pft);
		}
	}
	VariantClear(&v);
}

HRESULT teSHGetDataFromIDList(IShellFolder *pSF, LPCITEMIDLIST pidlPart, int nFormat, void *pv, int cb)
{
	HRESULT hr = SHGetDataFromIDList(pSF, pidlPart, nFormat, pv, cb);
	if (hr != S_OK && nFormat == SHGDFIL_FINDDATA) {
		::ZeroMemory(pv, cb);
		WIN32_FIND_DATA *pwfd = (WIN32_FIND_DATA *)pv;

		BSTR bs = NULL;
		if SUCCEEDED(teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_INFOLDER, &bs)) {
			lstrcpyn(pwfd->cFileName, bs, MAX_PATH);
			::SysFreeString(bs);
			hr = S_OK;

			IShellFolder2 *pSF2;
			if SUCCEEDED(pSF->QueryInterface(IID_PPV_ARGS(&pSF2))) {
				VARIANT v;
				VariantInit(&v);
				if SUCCEEDED(pSF2->GetDetailsEx(pidlPart, &PKEY_FileAttributes, &v)) {
					int nAttritutes = GetIntFromVariantClear(&v);
					if (nAttritutes != -1) {
						pwfd->dwFileAttributes = nAttritutes;
					}
				}
				if SUCCEEDED(pSF2->GetDetailsEx(pidlPart, &PKEY_Size, &v)) {
					ULARGE_INTEGER uli;
					uli.QuadPart = GetLLFromVariantClear(&v);
					pwfd->nFileSizeHigh = uli.HighPart;
					pwfd->nFileSizeLow = uli.LowPart;
				}
				teGetFileTimeFromItem(pSF2, pidlPart, &PKEY_DateCreated, &pwfd->ftCreationTime);
				teGetFileTimeFromItem(pSF2, pidlPart, &PKEY_DateAccessed, &pwfd->ftLastAccessTime);
				teGetFileTimeFromItem(pSF2, pidlPart, &PKEY_DateModified, &pwfd->ftLastWriteTime);
				pSF2->Release();
			}
		}
	}
	return hr;
}

BOOL teGetIDListFromVariant(LPITEMIDLIST *ppidl, VARIANT *pv, BOOL bForceLimit = FALSE)
{
	*ppidl = NULL;
	switch(pv->vt) {
		case VT_UNKNOWN:
		case VT_DISPATCH:
			teGetIDListFromObject(pv->punkVal, ppidl);
			break;
		case VT_UNKNOWN | VT_BYREF:
		case VT_DISPATCH | VT_BYREF:
			teGetIDListFromObject(*pv->ppunkVal, ppidl);
			break;
		case VT_ARRAY | VT_I1:
		case VT_ARRAY | VT_UI1:
		case VT_ARRAY | VT_I1 | VT_BYREF:
		case VT_ARRAY | VT_UI1 | VT_BYREF:
			LONG lUBound, lLBound, nSize;
			SAFEARRAY *psa;
			PVOID pvData;
			psa = (pv->vt & VT_BYREF) ? pv->pparray[0] : pv->parray;
			if (::SafeArrayAccessData(psa, &pvData) == S_OK) {
				SafeArrayGetUBound(psa, 1, &lUBound);
				SafeArrayGetLBound(psa, 1, &lLBound);
				nSize = lUBound - lLBound + 1;
				*ppidl = (LPITEMIDLIST)::CoTaskMemAlloc(nSize);
				::CopyMemory(*ppidl, pvData, nSize);
				::SafeArrayUnaccessData(psa);
			}
			break;
		case VT_NULL:
		case VT_EMPTY:
			return FALSE;
		case VT_BSTR:
		case VT_LPWSTR:
			*ppidl = teILCreateFromPath0(pv->bstrVal, bForceLimit);
			break;
		case VT_BSTR | VT_BYREF:
		case VT_LPWSTR | VT_BYREF:
			*ppidl = teILCreateFromPath0(*pv->pbstrVal, bForceLimit);
			break;
		case VT_VARIANT | VT_BYREF:
			return teGetIDListFromVariant(ppidl, pv->pvarVal, bForceLimit);
	}
	if (*ppidl == NULL && pv->vt != VT_DISPATCH) {
		VARIANT v;
		VariantInit(&v);
		if (SUCCEEDED(VariantChangeType(&v, pv, 0, VT_I4))) {
			if (v.uintVal < MAX_CSIDL2) {
				*ppidl = ::ILClone(g_pidls[v.uintVal]);
			}
		}
	}
	return (*ppidl != NULL);
}

BOOL teCreateSafeArray(VARIANT *pv, PVOID pSrc, DWORD dwSize, BOOL bBSTR)
{
	if (bBSTR) {
		pv->vt = VT_BSTR;
		pv->bstrVal = ::SysAllocStringByteLen(NULL, dwSize);
		if (pSrc && pv->bstrVal) {
			::CopyMemory(pv->bstrVal, pSrc, dwSize);
		}
		return TRUE;
	} else {
		pv->parray = SafeArrayCreateVector(VT_UI1, 0, dwSize);
		if (pv->parray) {
			pv->vt = VT_ARRAY | VT_UI1;
			PVOID pvData;
			if (pSrc && ::SafeArrayAccessData(pv->parray, &pvData) == S_OK) {
				::CopyMemory(pvData, pSrc, dwSize);
				::SafeArrayUnaccessData(pv->parray);
			}
			return TRUE;
		}
	}
	pv->vt = VT_EMPTY;
	return FALSE;
}

BOOL GetVarArrayFromIDList(VARIANT *pv, LPITEMIDLIST pidl)
{
	return teCreateSafeArray(pv, pidl, ::ILGetSize(pidl), FALSE);
}

HRESULT GetFolderObjFromIDList(LPITEMIDLIST pidl, Folder** ppsdf)
{
	VARIANT v;
	GetVarArrayFromIDList(&v, pidl);
	IShellDispatch *psha;
	HRESULT hr = teCreateInstance(CLSID_Shell, NULL, NULL, IID_PPV_ARGS(&psha));
	if SUCCEEDED(hr) {
		hr = psha->NameSpace(v, ppsdf);
	}
	psha->Release();
	VariantClear(&v);
	return hr;
}

BOOL GetFolderItemFromIDList2(FolderItem **ppid, LPITEMIDLIST pidl)
{
	BOOL Result;
	Result = FALSE;
	*ppid = NULL;
	IPersistFolder *pPF = NULL;
	if SUCCEEDED(teCreateInstance(CLSID_FolderItem, NULL, NULL, IID_PPV_ARGS(&pPF))) {
		if SUCCEEDED(pPF->Initialize(pidl)) {
			Result = SUCCEEDED(pPF->QueryInterface(IID_PPV_ARGS(ppid)));
		}
		pPF->Release();
		return Result;
	}
#ifdef _W2000
	Folder *pFolder = NULL;
	if (GetFolderObjFromIDList(pidl, &pFolder) == S_OK) {
		Folder2 *pFolder2 = NULL;
		if SUCCEEDED(pFolder->QueryInterface(IID_PPV_ARGS(&pFolder2))) {
			pFolder2->get_Self(ppid);
			pFolder2->Release();
		}
		pFolder->Release();
		if (*ppid) {
			return TRUE;
		}
	}
	LPITEMIDLIST pidlParent = ::ILClone(pidl);
	if (ILRemoveLastID(pidlParent) && GetFolderObjFromIDList(pidlParent, &pFolder) == S_OK) {
		BSTR bs;
		if SUCCEEDED(teGetDisplayNameFromIDList(&bs, pidl, SHGDN_FORPARSING | SHGDN_INFOLDER)) {
			pFolder->ParseName(bs, ppid);
			::SysFreeString(bs);
		}
		pFolder->Release();
	}
	teCoTaskMemFree(pidlParent);
	if (*ppid) {
		return TRUE;
	}
#endif
	return FALSE;
}

BOOL GetFolderItemFromIDList(FolderItem **ppid, LPITEMIDLIST pidl)
{
	BOOL Result = FALSE;
	CteFolderItem *pPF = new CteFolderItem(NULL);
	if SUCCEEDED(pPF->Initialize(pidl)) {
		Result = SUCCEEDED(pPF->QueryInterface(IID_PPV_ARGS(ppid)));
	}
	pPF->Release();
	return Result;
}

BOOL GetFolderItemFromFolderItems(FolderItem **ppFolderItem, FolderItems *pFolderItems, int nIndex)
{
	VARIANT v;
	teSetLong(&v, nIndex);
	if (pFolderItems->Item(v, ppFolderItem) == S_OK) {
		return TRUE;
	}
	GetFolderItemFromIDList(ppFolderItem, g_pidls[CSIDL_DESKTOP]);
	return TRUE;
}

/*//
VOID GetFolderItemsFromPCZZSTR(CteFolderItems **ppFolderItems, LPCWSTR pszPath)
{
	*ppFolderItems = new CteFolderItems(NULL, NULL, TRUE);
	VARIANT v;
	VariantInit(&v);
	while (pszPath[0]) {
		v.vt = VT_BSTR;
		v.bstrVal = ::SysAllocString(pszPath);
		(*ppFolderItems)->ItemEx(-1, NULL, &v);
		VariantClear(&v);
		while (*(pszPath++));
	}
}
//*/


HRESULT GetFolderItemFromObject(FolderItem **ppid, IUnknown *pid)
{
	if (pid) {
		LPITEMIDLIST pidl;
		if (teGetIDListFromObject(pid, &pidl)) {
			GetFolderItemFromIDList(ppid, pidl);
			teCoTaskMemFree(pidl);
			return S_OK;
		}
	}
	return E_FAIL;
}

VOID teSetIDList(VARIANT *pv, LPITEMIDLIST pidl)
{
	FolderItem *pFI;
	if (GetFolderItemFromIDList(&pFI, pidl)) {
		teSetObjectRelease(pv, pFI);
	}
}

VOID teSetIDListRelease(VARIANT *pv, LPITEMIDLIST *ppidl)
{
	teSetIDList(pv, *ppidl);
	teILFreeClear(ppidl);
}

int teGetObjectLength(IDispatch *pdisp)
{
	VARIANT v;
	VariantInit(&v);
	teGetProperty(pdisp, L"length", &v);
	return GetIntFromVariantClear(&v);
}

BOOL GetDataObjFromObject(IDataObject **ppDataObj, IUnknown *punk)
{
	*ppDataObj = NULL;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(ppDataObj))) {
		return TRUE;
	}
	FolderItems *pItems;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pItems))) {
		long lCount = 0;
		pItems->get_Count(&lCount);
		if (lCount) {
			CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL);
			VARIANT v, v2;
			VariantInit(&v2);
			v.vt = VT_I4;
			for (v.lVal = 0; v.lVal < lCount; ++v.lVal) {
				FolderItem *pItem;
				if SUCCEEDED(pItems->Item(v, &pItem)) {
					teSetObjectRelease(&v2, pItem);
					pFolderItems->ItemEx(-1, NULL, &v2);
					VariantClear(&v2);
				}
			}
			pFolderItems->QueryInterface(IID_PPV_ARGS(ppDataObj));
			pFolderItems->Release();
		}
		pItems->Release();
		if (*ppDataObj) {
			return TRUE;
		}
	}
	IDispatchEx *pdex;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pdex))) {
		long lCount = teGetObjectLength(pdex);
		if (lCount) {
			VARIANT v;
			VariantInit(&v);
			CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL);
			for (int i = 0; i < lCount; ++i) {
				teGetPropertyAt(pdex, i, &v);
				pFolderItems->ItemEx(-1, NULL, &v);
				VariantClear(&v);
			}
			pFolderItems->QueryInterface(IID_PPV_ARGS(ppDataObj));
			pFolderItems->Release();
		}
		pdex->Release();
		if (*ppDataObj) {
			return TRUE;
		}
	}
	return FALSE;
}

BOOL GetDataObjFromVariant2(IDataObject **ppDataObj, VARIANT *pv)
{
	*ppDataObj = NULL;
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		if (GetDataObjFromObject(ppDataObj, punk)) {
			return TRUE;
		}
	}
	LPITEMIDLIST pidl = NULL;
	if (teGetIDListFromVariant(&pidl, pv)) {
		IShellFolder *pSF;
		LPCITEMIDLIST pidlPart;
		if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
			pSF->GetUIObjectOf(g_hwndMain, 1, &pidlPart, IID_IDataObject, NULL, (LPVOID*)ppDataObj);
			pSF->Release();
		}
		teCoTaskMemFree(pidl);
	}
	return *ppDataObj != NULL;
}

BOOL GetDataObjFromVariant(IDataObject **ppDataObj, VARIANT *pv)
{
	if (GetDataObjFromVariant2(ppDataObj, pv)) {
		return TRUE;
	}
	IDispatch *pdisp;
	if (GetDispatch(pv, &pdisp)) {
		VARIANT vResult;
		VariantInit(&vResult);
		Invoke4(pdisp, &vResult, 0, NULL);
		if (vResult.vt != VT_EMPTY) {
			GetDataObjFromVariant2(ppDataObj, &vResult);
			VariantClear(&vResult);
		}
		pdisp->Release();
	}
	return *ppDataObj != NULL;
}

#ifdef _USE_BSEARCHAPI
int* teSortDispatchApi(TEDispatchApi *method, int nCount)
{
	int *pi = new int[nCount];

	for (int j = 0; j < nCount; ++j) {
		LPSTR pszNameA = method[j].name;
		int nMin = 0;
		int nMax = j - 1;
		int nIndex;
		while (nMin <= nMax) {
			nIndex = (nMin + nMax) / 2;
			if (lstrcmpiA(pszNameA, method[pi[nIndex]].name) < 0) {
				nMax = nIndex - 1;
			} else {
				nMin = nIndex + 1;
			}
		}
		for (int i = j; i > nMin; --i) {
			pi[i] = pi[i - 1];
		}
		pi[nMin] = j;
	}
	return pi;
}

int teBSearchApi(TEDispatchApi *method, int nSize, int* pMap, LPOLESTR bs)
{
	int nMin = 0;
	int nMax = nSize - 1;
	int nIndex, nCC;
	CHAR pszNameA[32];
	::WideCharToMultiByte(CP_TE, 0, (LPCWSTR)bs, -1, pszNameA, ARRAYSIZE(pszNameA) - 1, NULL, NULL);

	while (nMin <= nMax) {
		nIndex = (nMin + nMax) / 2;
		nCC = lstrcmpiA(pszNameA, method[pMap[nIndex]].name);
		if (nCC < 0) {
			nMax = nIndex - 1;
			continue;
		}
		if (nCC > 0) {
			nMin = nIndex + 1;
			continue;
		}
		return pMap[nIndex];
	}
	return -1;
}
#endif

int* SortTEMethod(TEmethod *method, int nCount)
{
	int *pi = new int[nCount];

	for (int j = 0; j < nCount; ++j) {
		LPSTR pszNameA = method[j].name;
		int nMin = 0;
		int nMax = j - 1;
		int nIndex;
		while (nMin <= nMax) {
			nIndex = (nMin + nMax) / 2;
			if (lstrcmpiA(pszNameA, method[pi[nIndex]].name) < 0) {
				nMax = nIndex - 1;
			} else {
				nMin = nIndex + 1;
			}
		}
		for (int i = j; i > nMin; --i) {
			pi[i] = pi[i - 1];
		}
		pi[nMin] = j;
	}
	return pi;
}

int teBSearch(TEmethod *method, int nSize, int* pMap, LPOLESTR bs)
{
	int nMin = 0;
	int nMax = nSize - 1;
	int nIndex, nCC;
	CHAR pszNameA[32];
	::WideCharToMultiByte(CP_TE, 0, (LPCWSTR)bs, -1, pszNameA, ARRAYSIZE(pszNameA) - 1, NULL, NULL);

	while (nMin <= nMax) {
		nIndex = (nMin + nMax) / 2;
		nCC = lstrcmpiA(pszNameA, method[pMap[nIndex]].name);
		if (nCC < 0) {
			nMax = nIndex - 1;
			continue;
		}
		if (nCC > 0) {
			nMin = nIndex + 1;
			continue;
		}
		return pMap[nIndex];
	}
	return -1;
}

int* SortTEStruct(TEStruct *method, int nCount)
{
	int *pi = new int[nCount];

	for (int j = 0; j < nCount; ++j) {
		LPSTR pszNameA = method[j].name;
		int nMin = 0;
		int nMax = j - 1;
		int nIndex;
		while (nMin <= nMax) {
			nIndex = (nMin + nMax) / 2;
			if (lstrcmpiA(pszNameA, method[pi[nIndex]].name) < 0) {
				nMax = nIndex - 1;
			} else {
				nMin = nIndex + 1;
			}
		}
		for (int i = j; i > nMin; --i) {
			pi[i] = pi[i - 1];
		}
		pi[nMin] = j;
	}
	return pi;
}

int teBSearchStruct(TEStruct *method, int nSize, int* pMap, LPOLESTR bs)
{
	int nMin = 0;
	int nMax = nSize - 1;
	int nIndex, nCC;
	CHAR pszNameA[32];
	::WideCharToMultiByte(CP_TE, 0, (LPCWSTR)bs, -1, pszNameA, ARRAYSIZE(pszNameA) - 1, NULL, NULL);

	while (nMin <= nMax) {
		nIndex = (nMin + nMax) / 2;
		nCC = lstrcmpiA(pszNameA, method[pMap[nIndex]].name);
		if (nCC < 0) {
			nMax = nIndex - 1;
			continue;
		}
		if (nCC > 0) {
			nMin = nIndex + 1;
			continue;
		}
		return pMap[nIndex];
	}
	return -1;
}

int* SortTEMethodW(TEmethodW *method, int nCount)
{
	int *pi = new int[nCount];

	for (int j = 0; j < nCount; ++j) {
		BSTR bs = method[j].name;
		int nMin = 0;
		int nMax = j - 1;
		int nIndex;
		while (nMin <= nMax) {
			nIndex = (nMin + nMax) / 2;
			if (lstrcmpi(bs, method[pi[nIndex]].name) < 0) {
				nMax = nIndex - 1;
			} else {
				nMin = nIndex + 1;
			}
		}
		for (int i = j; i > nMin; --i) {
			pi[i] = pi[i - 1];
		}
		pi[nMin] = j;
	}
	return pi;
}

int teBSearchW(TEmethodW *method, int nSize, int* pMap, LPOLESTR bs)
{
	int nMin = 0;
	int nMax = nSize - 1;
	int nIndex, nCC;

	while (nMin <= nMax) {
		nIndex = (nMin + nMax) / 2;
		nCC = lstrcmpi(bs, method[pMap[nIndex]].name);
		if (nCC < 0) {
			nMax = nIndex - 1;
			continue;
		}
		if (nCC > 0) {
			nMin = nIndex + 1;
			continue;
		}
		return pMap[nIndex];
	}
	return -1;
}

HRESULT teGetDispId(TEmethod *method, int nCount, int* pMap, LPOLESTR bs, DISPID *rgDispId, BOOL bNum)
{
#ifdef _DEBUG_
	::OutputDebugString(bs);
	::OutputDebugString(L"\n");
#endif
	if (pMap) {
		int nIndex = teBSearch(method, nCount, pMap, bs);
		if (nIndex >= 0) {
			*rgDispId = method[nIndex].id;
			return S_OK;
		}
	} else {
		CHAR pszNameA[32];
		::WideCharToMultiByte(CP_TE, 0, (LPCWSTR)bs, -1, pszNameA, ARRAYSIZE(pszNameA) - 1, NULL, NULL);
		for (int i = 0; method[i].name; ++i) {
			if (lstrcmpiA(pszNameA, method[i].name) == 0) {
				*rgDispId = method[i].id;
				return S_OK;
			}
		}
	}
	if (bNum) {
		VARIANT v, vo;
		teSetSZ(&v, bs);
		VariantInit(&vo);
		if (SUCCEEDED(VariantChangeType(&vo, &v, 0, VT_I4))) {
			*rgDispId = vo.lVal + DISPID_COLLECTION_MIN;
			VariantClear(&vo);
			VariantClear(&v);
			return *rgDispId < DISPID_TE_MAX ? S_OK : S_FALSE;
		}
		VariantClear(&v);
	}
	return DISP_E_UNKNOWNNAME;
}
/*
HRESULT teGetDispIdUm(std::unorderd_map<std::wstring, LONG> method, LPOLESTR bs, DISPID *rgDispId, BOOL bNum)
{
	auto itr = method.find(bs);
	if(itr != method.end()) {
		*rgDispId = itr->second;
		return S_OK;
	}
	if (bNum) {
		VARIANT v, vo;
		teSetSZ(&v, bs);
		VariantInit(&vo);
		if (SUCCEEDED(VariantChangeType(&vo, &v, 0, VT_I4))) {
			*rgDispId = vo.lVal + DISPID_COLLECTION_MIN;
			VariantClear(&vo);
			VariantClear(&v);
			return *rgDispId < DISPID_TE_MAX ? S_OK : S_FALSE;
		}
		VariantClear(&v);
	}
	return DISP_E_UNKNOWNNAME;
}
*/
HRESULT teGetMemberName(DISPID id, BSTR *pbstrName)
{
	if (id >= DISPID_COLLECTION_MIN && id <= DISPID_TE_MAX) {
		wchar_t pszName[8];
		swprintf_s(pszName, 8, L"%d", id - DISPID_COLLECTION_MIN);
		*pbstrName = ::SysAllocString(pszName);
		return S_OK;
	}
	return E_NOTIMPL;
}

int GetSizeOfStruct(LPOLESTR bs)
{
	int nIndex = teBSearchStruct(pTEStructs, _countof(pTEStructs), g_maps[MAP_SS], bs);
	if (nIndex >= 0) {
		return pTEStructs[nIndex].lSize;
	}
	return 0;
}

int GetSizeOf(VARIANT *pv)
{
	if (pv->vt == VT_BSTR) {
		return GetSizeOfStruct(pv->bstrVal);
	}
	int i = SizeOfvt(static_cast<VARTYPE>(GetIntFromVariant(pv)));
	if (i) {
		return i;
	}
	if (pv->vt & VT_ARRAY) {
		return pv->parray->rgsabound[0].cElements * SizeOfvt(pv->vt);
	}
	return SizeOfvt(pv->vt);
}

//Left-Bottom "%" is 100 times, on the other 0x4000 times
int GetIntFromVariantPP(VARIANT *pv, int nIndex)
{
	if (pv->vt == (VT_VARIANT | VT_BYREF)) {
		return GetIntFromVariantPP(pv->pvarVal, nIndex);
	}
	if (nIndex >= TE_Left && nIndex <= TE_Bottom) {
		int i = 0;
		if (pv->vt == VT_BSTR) {
			if (tePathMatchSpec(pv->bstrVal, L"*%")) {
				float f = 0;
				if (swscanf_s(pv->bstrVal, L"%g%%", &f) != EOF) {
					i = (int)(f * 100);
					if (i > 10000) {
						i = 10000;	//Not exceed 100%
					}
				}
			}
		}
		if (i == 0) {
			return GetIntFromVariant(pv) * 0x4000;
		}
		return i;
	} else {
		return GetIntFromVariant(pv);
	}
}

//There is no need for SysFreeString
BSTR GetLPWSTRFromVariant(VARIANT *pv)
{
	switch (pv->vt) {
		case VT_BSTR:
		case VT_LPWSTR:
			return pv->bstrVal;
		case VT_VARIANT | VT_BYREF:
			return GetLPWSTRFromVariant(pv->pvarVal);
		default:
			return (BSTR)GetPtrFromVariant(pv);
	}//end_switch
}

UINT GetpDataFromVariant(UCHAR **ppc, VARIANT *pv, VARIANT *pvMem)
{
	if (pv->vt == (VT_VARIANT | VT_BYREF)) {
		return GetpDataFromVariant(ppc, pv->pvarVal, pvMem);
	}
	if (pv->vt & VT_ARRAY) {
		if (SizeOfvt(pv->vt) == 1) {
			*ppc = (UCHAR *)pv->parray->pvData;
			return pv->parray->rgsabound[0].cElements;
		}
		if (pvMem) {
			pvMem->bstrVal = ::SysAllocStringByteLen(NULL, pv->parray->rgsabound[0].cElements);
			pvMem->vt = VT_BSTR;
			*ppc = (UCHAR *)pvMem->bstrVal;
			VARIANT v;
			VariantInit(&v);
			for (LONG i = pv->parray->rgsabound[0].cElements; i-- > 0;) {
				SafeArrayGetElement(pv->parray, &i, &v);
				(*ppc)[i] = (UCHAR)GetIntFromVariant(&v);
			}
			return pv->parray->rgsabound[0].cElements;
		}
	}
	char *pc = GetpcFromVariant(pv, NULL);
	if (pc) {
		IDispatch *pdisp;
		if (GetDispatch(pv, &pdisp)) {
			VARIANT v;
			VariantInit(&v);
			if (teGetProperty(pdisp, L"Size", &v) == S_OK) {
				*ppc = (UCHAR *)pc;
			}
			pdisp->Release();
			return GetIntFromVariantClear(&v);
		}
	}
	*ppc = (UCHAR *)GetLPWSTRFromVariant(pv);
	if (pv->vt == VT_BSTR) {
		return ::SysStringByteLen(pv->bstrVal);
	}
	return 0;
}

VOID Invoke4(IDispatch *pdisp, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs)
{
	Invoke5(pdisp, DISPID_VALUE, DISPATCH_METHOD, pvResult, nArgs, pvArgs);
}

VOID teCustomDraw(int nFunc, CteShellBrowser *pSB, CteTreeView *pTV, IShellItem *psi, LPNMCUSTOMDRAW lpnmcd, PVOID pvcd, LRESULT *plres)
{
	if (lpnmcd->rc.top == 0 && lpnmcd->rc.bottom == 0) {
#ifdef _2000XP
		if (g_bUpperVista) {
			return;
		}
#else
		return;
#endif
	}
	if (pSB) {
		HWND hHeader = ListView_GetHeader(pSB->m_hwndLV);
		if (hHeader && IsWindowVisible(hHeader)) {
			RECT rc;
			GetWindowRect(hHeader, &rc);
			if (lpnmcd->rc.bottom <= rc.bottom - rc.top) {
				return;
			}
		}
	}
	LPITEMIDLIST pidl = NULL;
	PVOID pvcd2 = NULL;
	if (nFunc == TE_OnItemPrePaint && g_pOnFunc[TE_OnItemPostPaint]) {
		*plres = CDRF_NOTIFYPOSTPAINT;
	}
	if (g_pOnFunc[nFunc]) {
		VARIANTARG *pv = GetNewVARIANT(5);
		if (pSB) {
			teSetObject(&pv[4], pSB);
			pvcd2 = new CteMemory(sizeof(NMLVCUSTOMDRAW), pvcd, 1, L"NMLVCUSTOMDRAW");
			IFolderView *pFV;
			if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
				LPITEMIDLIST pidlChild;
				if SUCCEEDED(pFV->Item((int)lpnmcd->dwItemSpec, &pidlChild)) {
					pidl = ILCombine(pSB->m_pidl, pidlChild);
					teCoTaskMemFree(pidlChild);
				}
				pFV->Release();
			}
		} else {
			teSetObject(&pv[4], pTV);
			pvcd2 = new CteMemory(sizeof(NMTVCUSTOMDRAW), pvcd, 1, L"NMTVCUSTOMDRAW");
			if (psi && lpnmcd->rc.bottom) {
				teGetIDListFromObject(psi, &pidl);
				TVHITTESTINFO ht = { 0 };
				ht.pt.x = (lpnmcd->rc.left + lpnmcd->rc.right) / 2;
				ht.pt.y = (lpnmcd->rc.top + lpnmcd->rc.bottom) / 2;
				TreeView_HitTest(pTV->m_hwndTV, &ht);
				lpnmcd->dwItemSpec = (DWORD_PTR)ht.hItem;
			}
#ifdef _2000XP
			if (!pidl && pTV->m_pShellNameSpace && (lpnmcd->uItemState & CDIS_SELECTED)) {
				IDispatch *pid;
				if (pTV->m_pShellNameSpace->get_SelectedItem(&pid) == S_OK) {
					teGetIDListFromObject(pid, &pidl);
					pid->Release();
				}
			}
#endif
		}
		if (pidl) {
			teSetIDListRelease(&pv[3], &pidl);
			teSetObjectRelease(&pv[2], new CteMemory(sizeof(NMCUSTOMDRAW), lpnmcd, 1, L"NMCUSTOMDRAW"));
			teSetObjectRelease(&pv[1], pvcd2);
			teSetObjectRelease(&pv[0], new CteMemory(sizeof(HANDLE), plres, 1, L"HANDLE"));
			Invoke4(g_pOnFunc[nFunc], NULL, 5, pv);
		} else {
			VariantClear(&pv[4]);
			SafeRelease(&pvcd2);
			delete [] pv;
		}
	}
}

#ifdef _LOG
VOID teLog(HANDLE hFile, LPWSTR lpLog)
{
	DWORD dwWriteByte;
	WriteFile(hFile, lpLog, lstrlen(lpLog) * 2, &dwWriteByte, NULL);
}
#endif

BOOL teFreeLibrary2(HMODULE hDll)
{
	LPFNDllCanUnloadNow _DllCanUnloadNow = (LPFNDllCanUnloadNow)GetProcAddress(hDll, "DllCanUnloadNow");
	if (g_nReload || (_DllCanUnloadNow && _DllCanUnloadNow() != S_OK)) {
		g_pFreeLibrary.insert(g_pFreeLibrary.begin(), hDll);
		SetTimer(g_hwndMain, TET_FreeLibrary, (++g_dwFreeLibrary) * 100, teTimerProc);
		return FALSE;
	}
	g_dwFreeLibrary = 0;
	return FreeLibrary(hDll);
}

BOOL teFreeLibrary(HMODULE hDll, UINT uElpase)
{
	if (hDll) {
		if (uElpase) {
			g_dwFreeLibrary = 0;
			g_pFreeLibrary.push_back(hDll);
			SetTimer(g_hwndMain, TET_FreeLibrary, uElpase, teTimerProc);
			return TRUE;
		}
		return FreeLibrary(hDll);
	}
	return FALSE;
}

HRESULT ParseScript(LPOLESTR lpScript, LPOLESTR lpLang, VARIANT *pv, IDispatch **ppdisp, EXCEPINFO *pExcepInfo)
{
	HRESULT hr = E_FAIL;
	CLSID clsid;
	IActiveScript *pas = NULL;

	BOOL bChakra = tePathMatchSpec(lpLang, L"J*Script");
	if (g_nBlink == 1 && bChakra) {
		teCreateInstance(CLSID_JScriptEdgeChakra, NULL, NULL, IID_PPV_ARGS(&pas));
	}
	if (pas == NULL && bChakra) {
		teCreateInstance(CLSID_JScript9Chakra, NULL, NULL, IID_PPV_ARGS(&pas));
		bChakra = FALSE;
	}
	if (pas == NULL && teCLSIDFromProgID(lpLang, &clsid) == S_OK) {
		teCreateInstance(clsid, NULL, NULL, IID_PPV_ARGS(&pas));
		bChakra = FALSE;
	}
	if (pas) {
		if (!bChakra) {
			IActiveScriptProperty *paspr;
			if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&paspr))) {
				VARIANT v;
				teSetLong(&v, 0);
				while (++v.lVal <= 256 && paspr->SetProperty(SCRIPTPROP_INVOKEVERSIONING, NULL, &v) == S_OK) {
				}
				paspr->Release();
			}
		}
		IUnknown *punk = NULL;
		FindUnknown(pv, &punk);
		CteActiveScriptSite *pass = new CteActiveScriptSite(punk, pExcepInfo, &hr);
		pas->SetScriptSite(pass);
		pass->Release();
		IActiveScriptParse *pasp;
		if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&pasp))) {
			hr = pasp->InitNew();
			if (punk) {
				IDispatchEx *pdex;
				if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pdex))) {
					DISPID dispid;
					for (HRESULT hr2 = pdex->GetNextDispID(fdexEnumAll, DISPID_STARTENUM, &dispid); hr2 == S_OK; hr2 = pdex->GetNextDispID(fdexEnumAll, dispid, &dispid)) {
						BSTR bs;
						if (pdex->GetMemberName(dispid, &bs) == S_OK) {
							pas->AddNamedItem(bs, SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE | SCRIPTITEM_GLOBALMEMBERS);
							::SysFreeString(bs);
						}
					}
					pdex->Release();
				}
			} else if (pv && pv->boolVal) {
				if (g_pWebBrowser && g_dwMainThreadId == GetCurrentThreadId()) {
					pas->AddNamedItem(L"window", SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE | SCRIPTITEM_GLOBALMEMBERS);
				}
			}
			if (g_dwMainThreadId != GetCurrentThreadId()) {
				pas->AddNamedItem(L"api", SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE);
			}
			VARIANT v;
			VariantInit(&v);
			pasp->ParseScriptText(lpScript, NULL, NULL, NULL, 0, 0, SCRIPTTEXT_ISPERSISTENT | SCRIPTTEXT_ISVISIBLE, &v, NULL);
			if (hr == S_OK) {
				pas->SetScriptState(SCRIPTSTATE_CONNECTED);
				if (ppdisp) {
					hr = E_FAIL;
					IDispatch *pdisp;
					if (pas->GetScriptDispatch(NULL, &pdisp) == S_OK) {
						CteDispatch *odisp = new CteDispatch(pdisp, 2, DISPID_UNKNOWN);
						pdisp->Release();
						if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&odisp->m_pActiveScript))) {
							hr = odisp->QueryInterface(IID_PPV_ARGS(ppdisp));
						}
						odisp->Release();
					}
				}
			}
			pasp->Release();
		}
		if (!ppdisp || !*ppdisp) {
			pas->SetScriptState(SCRIPTSTATE_CLOSED);
			pas->Close();
		}
		pas->Release();
	}
	return hr;
}

VOID GetPointFormVariant(POINT *ppt, VARIANT *pv)
{
	IDispatch *pdisp;
	if (GetDispatch(pv, &pdisp)) {
		VARIANT v;
		VariantInit(&v);
		teGetProperty(pdisp, L"x", &v);
		ppt->x = GetIntFromVariantClear(&v);
		teGetProperty(pdisp, L"y", &v);
		ppt->y = GetIntFromVariantClear(&v);
		pdisp->Release();
		return;
	}
	int npt = GetIntFromVariant(pv);
	ppt->x = GET_X_LPARAM(npt);
	ppt->y = GET_Y_LPARAM(npt);
}

VOID PutPointToVariant(POINT *ppt, VARIANT *pv)
{
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		VARIANT v;
		VariantInit(&v);
		teSetLong(&v, ppt->x);
		tePutProperty(punk, L"x", &v);
		VariantClear(&v);
		teSetLong(&v, ppt->y);
		tePutProperty(punk, L"y", &v);
		VariantClear(&v);
	}
}

HRESULT DoFunc1(int nFunc, PVOID pObj, VARIANT *pVarResult)
{
	if (g_pOnFunc[nFunc]) {
		VARIANTARG *pv = GetNewVARIANT(1);
		teSetObject(&pv[0], pObj);
		Invoke4(g_pOnFunc[nFunc], pVarResult, 1, pv);
		return S_OK;
	}
	return E_FAIL;
}

HRESULT DoFunc(int nFunc, PVOID pObj, HRESULT hr)
{
	VARIANT vResult;
	VariantInit(&vResult);
	if (SUCCEEDED(DoFunc1(nFunc, pObj, &vResult)) && vResult.vt != VT_EMPTY) {
		hr = GetIntFromVariantClear(&vResult);
	}
	return hr;
}

VARIANT_BOOL OpenNewWindowV(VARIANT *pv) {
	VARIANT_BOOL bFolder = VARIANT_FALSE;
	FolderItem *pid = new CteFolderItem(pv);
	pid->get_IsFolder(&bFolder);
	if (bFolder && g_pOnFunc[TE_OnNewWindow]) {
		VARIANTARG *pv = GetNewVARIANT(2);
		teSetObject(&pv[1], g_pWebBrowser);
		teSetObject(&pv[0], pid);
		Invoke4(g_pOnFunc[TE_OnNewWindow], NULL, 2, pv);
	}
	pid->Release();
	return bFolder;
}

VOID teSetProgress(IProgressDialog *ppd, ULONGLONG ullCurrent, ULONGLONG ullTotal, int nMode)
{
	WCHAR pszMsg[128];
	ppd->SetProgress64(ullCurrent, ullTotal);
	if (nMode && ullTotal) {
		if (nMode > 1 && ullTotal != 100) {
			WCHAR pszCurrent[32], pszTotal[32];
			swprintf_s(pszMsg, 32, L"%llu", ullCurrent);
			teCommaSize(pszMsg, pszCurrent, 32, 0);
			swprintf_s(pszMsg, 32, L"%llu", ullTotal);
			teCommaSize(pszMsg, pszTotal, 32, 0);
			swprintf_s(pszMsg, 128, L"%llu%% ( %s / %s )", 100 * ullCurrent / ullTotal, pszCurrent, pszTotal);
		} else {
			swprintf_s(pszMsg, 128, L"%llu%%", 100 * ullCurrent / ullTotal);
		}
		ppd->SetTitle(pszMsg);
	}
}

BOOL teSetProgressEx(IProgressDialog *ppd, ULONGLONG ullCurrent, ULONGLONG ullTotal, int nMode)
{
	teSetProgress(ppd, ullCurrent, ullTotal, nMode);
	return ullCurrent < ullTotal && !ppd->HasUserCancelled();
}

static void threadAddItems(void *args)
{
	::CoInitialize(NULL);
	WCHAR pszMsg[MAX_PATH];
	LPITEMIDLIST pidl;
	IProgressDialog *ppd = NULL;

	TEAddItems *pAddItems = (TEAddItems *)args;
	IDispatch *pSB;
	CoGetInterfaceAndReleaseStream(pAddItems->pStrmSB, IID_PPV_ARGS(&pSB));
	IDispatch *pArray;
	CoGetInterfaceAndReleaseStream(pAddItems->pStrmArray, IID_PPV_ARGS(&pArray));
	IDispatch *pOnCompleted = NULL;
	if (pAddItems->pStrmOnCompleted) {
		CoGetInterfaceAndReleaseStream(pAddItems->pStrmOnCompleted, IID_PPV_ARGS(&pOnCompleted));
	}
	teCreateInstance(CLSID_ProgressDialog, NULL, NULL, IID_PPV_ARGS(&ppd));
	if (ppd) {
		try {
			ppd->StartProgressDialog(g_hwndMain, NULL, PROGDLG_NORMAL | PROGDLG_AUTOTIME, NULL);
#ifdef _2000XP
			ppd->SetAnimation(g_hShell32, 150);
#endif
			if (!LoadString(g_hShell32, 13576, pszMsg, MAX_PATH)) {
				LoadString(g_hShell32, 4223, pszMsg, MAX_PATH);
			}
			ppd->SetLine(1, pszMsg, TRUE, NULL);

			FolderItems *pItems;
			if SUCCEEDED(pArray->QueryInterface(IID_PPV_ARGS(&pItems))) {
				long lCount = 0;
				pItems->get_Count(&lCount);
				if (lCount) {
					VARIANT v;
					v.vt = VT_I4;
					for (v.lVal = 0; teSetProgressEx(ppd, v.lVal, lCount, 2); ++v.lVal) {
						FolderItem *pItem;
						if SUCCEEDED(pItems->Item(v, &pItem)) {
							BOOL bAdd = TRUE;
							if (!pAddItems->bDeleted) {
								CteFolderItem *pid;
								if SUCCEEDED(pItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid)) {
									bAdd = !pid->m_dwUnavailable;
									pid->Release();
								}
							}
							if (bAdd) {
								VariantClear(&pAddItems->pv[1]);
								teSetObjectRelease(&pAddItems->pv[1], pItem);
								if (Invoke5(pSB, TE_METHOD + 0xf501, DISPATCH_METHOD, NULL, -2, pAddItems->pv) == E_ACCESSDENIED) {
									break;
								}
							} else {
								SafeRelease(&pItem);
							}
						}
					}
					VariantClear(&pAddItems->pv[1]);
				}
				pItems->Release();
			} else {
				int nCount = teGetObjectLength(pArray);
				for (int nCurrent = 0; teSetProgressEx(ppd, nCurrent, nCount, 2); ++nCurrent) {
					if (SUCCEEDED(teGetPropertyAt(pArray, nCurrent, &pAddItems->pv[1])) && pAddItems->pv[1].vt != VT_EMPTY) {
						if (!teGetIDListFromVariant(&pidl, &pAddItems->pv[1], TRUE) && pAddItems->bDeleted && pAddItems->pv[1].vt == VT_BSTR) {
							pidl = teSHSimpleIDListFromPathEx(pAddItems->pv[1].bstrVal, FILE_ATTRIBUTE_HIDDEN, -1, -1, NULL);
						}
						if (pidl) {
							VariantClear(&pAddItems->pv[1]);
							teSetIDListRelease(&pAddItems->pv[1], &pidl);
							if (Invoke5(pSB, TE_METHOD + 0xf501, DISPATCH_METHOD, NULL, -2, pAddItems->pv) == E_ACCESSDENIED) {
								break;
							}
						}
						VariantClear(&pAddItems->pv[1]);
					}
				}
			}
			ppd->SetLine(2, L"", TRUE, NULL);

			if (pAddItems->bNavigateComplete) {
				Invoke5(pSB, TE_METHOD + 0xf500, DISPATCH_METHOD, NULL, 0, NULL);
			}
			if (pOnCompleted) {
				VARIANT *pv = GetNewVARIANT(3);
				teSetObject(&pv[2], pSB);
				teSetObject(&pv[1], pArray);
				teSetObjectRelease(&pv[0], new CteProgressDialog(ppd));
				Invoke4(pOnCompleted, NULL, 3, pv);
			}
			SetTimer(g_hwndMain, TET_Status, 500, teTimerProc);
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"threadAddItems";
#endif
		}
	}
	teClearVariantArgs(2, pAddItems->pv);
	delete [] pAddItems;
	SafeRelease(&pOnCompleted);
	SafeRelease(&pArray);
	SafeRelease(&pSB);
	if (ppd) {
		ppd->SetLine(2, L"", TRUE, NULL);
		ppd->StopProgressDialog();
		SafeRelease(&ppd);
	}
	::CoUninitialize();
	::_endthread();
}

BOOL teILGetParent(FolderItem *pid, FolderItem **ppid)
{
	BOOL bResult =FALSE;
	VARIANT v;
	VariantInit(&v);
	if SUCCEEDED(DoFunc1(TE_OnILGetParent, pid, &v)) {
		if (v.vt != VT_EMPTY) {
			GetFolderItemFromVariant(ppid, &v);
			return TRUE;
		}
	}
	LPITEMIDLIST pidl = NULL;
	if (teGetIDListFromObject(pid, &pidl)) {
		bResult = ::ILRemoveLastID(pidl);
		if (::ILIsEqual(pidl, g_pidls[CSIDL_INTERNET])) {
			bResult = ::ILRemoveLastID(pidl);
		}
		if (!bResult || ILIsEmpty(pidl)) {
			BSTR bs;
			if SUCCEEDED(pid->get_Path(&bs)) {
				if (tePathMatchSpec(bs, L"?:\\*")) {
					if (PathRemoveFileSpec(bs)) {
						CteFolderItem *pid1 = new CteFolderItem(NULL);
						pid1->m_v.vt = VT_BSTR;
						pid1->m_v.bstrVal = bs;
						pid1->QueryInterface(IID_PPV_ARGS(ppid));
						pid1->Release();
						teCoTaskMemFree(pidl);
						return TRUE;
					}
				}
				::SysFreeString(bs);
			}
		}
		GetFolderItemFromIDList(ppid, pidl);
		teCoTaskMemFree(pidl);
	}
	return bResult;
}

VOID GetNewObject1(LPOLESTR sz, IDispatch **ppdisp)
{
	VARIANT v;
	VariantInit(&v);
	IDispatch *pdisp = *ppdisp;
	if (g_dwMainThreadId == GetCurrentThreadId()) {
		if (teExecMethod(g_pJS, sz, &v, 0, NULL) == S_OK) {
			GetDispatch(&v, ppdisp);
			VariantClear(&v);
			SafeRelease(&pdisp);
			return;
		}
	}
	BOOL bSuccess = FALSE;
	IGlobalInterfaceTable *pGlobalInterfaceTable;
	CoCreateInstance(CLSID_StdGlobalInterfaceTable, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pGlobalInterfaceTable));
	IDispatch *pJS = NULL;
	if SUCCEEDED(pGlobalInterfaceTable->GetInterfaceFromGlobal(g_dwCookieJS, IID_PPV_ARGS(&pJS))) {
		if (teExecMethod(pJS, sz, &v, 0, NULL) == S_OK) {
			bSuccess = GetDispatch(&v, ppdisp);
			VariantClear(&v);
		}
		pJS->Release();
	}
	pGlobalInterfaceTable->Release();
	if (bSuccess) {
		SafeRelease(&pdisp);
	} else {
		SafeRelease(ppdisp);
#ifdef _DEBUG
		MessageBox(g_hwndMain, L"Failed to get object.", NULL, MB_ICONERROR | MB_OK);
#endif
	}
}


VOID GetNewArray(IDispatch **ppArray)
{
	GetNewObject1(L"_a", ppArray);
}

VOID GetNewObject(IDispatch **ppObj)
{
	GetNewObject1(L"_o", ppObj);
}

IDispatch* teCreateObj(LONG lId, VARIANT *pvArg)
{
	IDispatch *pdisp = NULL;
	IDispatch *pdispResult = NULL;
	switch (lId) {
	case TE_OBJECT:
		if (g_nBlink != 1 || g_pSP->QueryService(SID_TablacusObject, IID_PPV_ARGS(&pdisp)) != S_OK) {
			GetNewObject(&pdisp);
		}
		return pdisp;

	case TE_ARRAY:
		if (g_nBlink != 1 || g_pSP->QueryService(SID_TablacusArray, IID_PPV_ARGS(&pdisp)) != S_OK) {
			if (pvArg && pvArg->vt == VT_DISPATCH) {
				pvArg->pdispVal->QueryInterface(IID_PPV_ARGS(&pdisp));
			} else {
				GetNewArray(&pdisp);
			}
		} else if (pvArg) {
			if (pvArg->vt == VT_DISPATCH) {
				pvArg->pdispVal->QueryInterface(IID_PPV_ARGS(&pdisp));
			} else {
				Invoke5(pdisp, DISPID_PROPERTYPUT, DISPATCH_PROPERTYPUT, NULL, -1, pvArg);
			}
		}
		return pdisp;

	case 3://Enum
		return new CteEnumerator(pvArg);

	case 4://api
		return new CteWindowsAPI(NULL);

	case 5://CommonDialog
		return new CteCommonDialog();

	case 6://FolderItems
		IDataObject *pDataObj;
		pDataObj = NULL;
		if (pvArg) {
			GetDataObjFromVariant(&pDataObj, pvArg);
		}
		pdispResult = static_cast<FolderItems *>(new CteFolderItems(pDataObj, NULL));
		SafeRelease(&pDataObj);
		return pdispResult;

	case 7://ProgressDialog
		return new CteProgressDialog(NULL);

	case 8://WICBitmap
		return new CteWICBitmap();

	}
	return NULL;
}

VOID ClearEvents()
{
	++g_dwSessionId;
	for (int j = Count_OnFunc; j-- > 1;) {
		SafeRelease(&g_pOnFunc[j]);
	}

	for (size_t i = 0; i < g_ppSB.size(); ++i) {
		CteShellBrowser *pSB = g_ppSB[i];
		for (int j = Count_SBFunc; j-- > 1;) {
			SafeRelease(&pSB->m_ppDispatch[j]);
		}
	}

	g_ppGetImage.clear();
	g_ppGetArchive.clear();
	g_ppMessageSFVCB.clear();

	g_param[TE_Tab] = TRUE;
	g_param[TE_Version] = 20000000 + VER_Y * 10000 + VER_M * 100 + VER_D;
	g_param[TE_ColumnEmphasis] = FALSE;
	g_param[TE_ViewOrder] = FALSE;
	g_param[TE_LibraryFilter] = FALSE;
	g_param[TE_AutoArrange] = FALSE;
	g_param[TE_ShowInternet] = FALSE;

	if (g_dwCookieSW) {
		teUnadviseAndRelease(g_pSW, DIID_DShellWindowsEvents, &g_dwCookieSW);
		g_dwCookieSW = 0;
	}

	EnableWindow(g_pWebBrowser->get_HWND(), TRUE);
	SetWindowPos(g_hwndMain, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
}

VOID teArrayPush(IDispatch *pdisp, PVOID pObj)
{
	VARIANT *pv = GetNewVARIANT(1);
	teSetObject(pv, pObj);
	teExecMethod(pdisp, L"push", NULL, 1, pv);
}

HRESULT DoStatusText(PVOID pObj, LPCWSTR lpszStatusText, int iPart)
{
	HRESULT hr = E_NOTIMPL;
	if (g_pOnFunc[TE_OnStatusText] && g_pOnFunc[TE_OnArrange]) {
		VARIANT vResult;
		VariantInit(&vResult);
		VARIANTARG *pv = GetNewVARIANT(3);
		teSetObject(&pv[2], pObj);
		teSetSZ(&pv[1], lpszStatusText);
		teSetLong(&pv[0], iPart);
		Invoke4(g_pOnFunc[TE_OnStatusText], &vResult, 3, pv);
		hr = GetIntFromVariantClear(&vResult);
	}
	g_szStatus[0] = NULL;
	return hr;
}

LONG_PTR DoHitTest(PVOID pObj, POINT pt, UINT flags)
{
	if (g_pOnFunc[TE_OnHitTest]) {
		VARIANT vResult;
		VariantInit(&vResult);
		VARIANTARG *pv = GetNewVARIANT(3);
		teSetObject(&pv[2], pObj);
		CteMemory *pstPt = new CteMemory(2 * sizeof(int), NULL, 1, L"POINT");
		pstPt->SetPoint(pt.x, pt.y);
		teSetObject(&pv[1], pstPt);
		pstPt->Release();
		teSetLong(&pv[0], flags);
		Invoke4(g_pOnFunc[TE_OnHitTest], &vResult, 3, pv);
		if (vResult.vt != VT_EMPTY) {
			return GetPtrFromVariantClear(&vResult);
		}
	}
	return -1;
}

HRESULT DragSub(int nFunc, PVOID pObj, CteFolderItems *pDragItems, PDWORD pgrfKeyState, POINTL pt, PDWORD pdwEffect)
{
	HRESULT hr = E_FAIL;
	VARIANT vResult;
	VARIANTARG *pv;
	CteMemory *pstPt;

	if (g_pOnFunc[nFunc]) {
		try {
			if (InterlockedIncrement(&g_nProcDrag) < 10) {
				pv = GetNewVARIANT(5);
				teSetObject(&pv[4], pObj);
				teSetObject(&pv[3], pDragItems);
				teSetObjectRelease(&pv[2], new CteMemory(sizeof(int), pgrfKeyState, 1, L"DWORD"));
				pstPt = new CteMemory(2 * sizeof(int), NULL, 1, L"POINT");
				pstPt->SetPoint(pt.x, pt.y);
				teSetObjectRelease(&pv[1], pstPt);
				teSetObjectRelease(&pv[0], new CteMemory(sizeof(int), pdwEffect, 1, L"DWORD"));
				Invoke4(g_pOnFunc[nFunc], &vResult, 5, pv);
				hr = GetIntFromVariant(&vResult);
			}
		} catch(...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"DragSub";
#endif
		}
		::InterlockedDecrement(&g_nProcDrag);
	}
	return hr;
}

HRESULT DragLeaveSub(PVOID pObj, CteFolderItems **ppDragItems)
{
	if (ppDragItems && *ppDragItems) {
		CteFolderItems *pDragItems = *ppDragItems;
		pDragItems->Release();
		*ppDragItems = NULL;
	}
	return DoFunc(TE_OnDragLeave, pObj, E_NOTIMPL);
}

BOOL MessageSub(int nFunc, PVOID pObj, MSG *pMsg, HRESULT *phr)
{
	VARIANT vResult;
	VARIANTARG *pv;

	if (g_pOnFunc[nFunc]) {
		pv = GetNewVARIANT(5);
		teSetObject(&pv[4], pObj);
		teSetPtr(&pv[3], pMsg->hwnd);
		teSetLong(&pv[2], pMsg->message);
		teSetPtr(&pv[1], pMsg->wParam);
		teSetPtr(&pv[0], pMsg->lParam);
		Invoke4(g_pOnFunc[nFunc], &vResult, 5, pv);
		if (vResult.vt != VT_EMPTY) {
			if (phr) {
				*phr = GetIntFromVariant(&vResult);
			}
			VariantClear(&vResult);
			return TRUE;
		}
	}
	return FALSE;
}

HRESULT teDoCommand(PVOID pObj, HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	if (g_pOnFunc[TE_OnCommand]) {
		MSG msg1;
		msg1.hwnd = hwnd;
		msg1.message = msg;
		msg1.wParam = wParam;
		msg1.lParam = lParam;
		HRESULT hr;
		if (MessageSub(TE_OnCommand, pObj, &msg1, &hr)) {
			return hr;
		}
	}
	return S_FALSE;
}

HRESULT tePSFormatForDisplay(PROPERTYKEY *ppropKey, VARIANT *pv, DWORD pdfFlags, LPWSTR *ppszDisplay, int nSizeFormat)
{
	HRESULT hr = E_FAIL;
	*ppszDisplay = NULL;
	if (g_bsDateTimeFormat && pv->vt == VT_DATE) {
		SYSTEMTIME st;
		if (teVariantTimeToSystemTime(pv->date, &st)) {
			LPWSTR pszDisplay = (LPWSTR)::CoTaskMemAlloc(MAX_PROP);
			pszDisplay[0] = NULL;
			*ppszDisplay = pszDisplay;
			BOOL bDate = TRUE;
			WCHAR wc = 0xffff;
			for (int i = 0, j = 0; wc; ++i) {
				wc = g_bsDateTimeFormat[i];
				if (bDate) {
					if (!wc || StrChr(L"hHmst", wc)) {
						if (i != j) {
							g_bsDateTimeFormat[i] = NULL;
							int k = lstrlen(*ppszDisplay);
							GetDateFormat(LOCALE_USER_DEFAULT, 0, &st, &g_bsDateTimeFormat[j], &pszDisplay[k], MAX_PROP - k);
							g_bsDateTimeFormat[i] = wc;
							j = i;
						}
						bDate = FALSE;
					}
				} else {
					if (!wc || StrChr(L"dMyg", wc)) {
						if (i != j) {
							g_bsDateTimeFormat[i] = NULL;
							int k = lstrlen(*ppszDisplay);
							GetTimeFormat(LOCALE_USER_DEFAULT, 0, &st, &g_bsDateTimeFormat[j], &pszDisplay[k], MAX_PROP - k);
							g_bsDateTimeFormat[i] = wc;
							j = i;
						}
						bDate = TRUE;
					}
				}
			}
			hr = S_OK;
		}
		return hr;
	}
	if (nSizeFormat && IsEqualPropertyKey(*ppropKey, PKEY_Size) && teVarIsNumber(pv)) {
		*ppszDisplay = (LPWSTR)::CoTaskMemAlloc(MAX_PROP);
		teStrFormatSize(nSizeFormat, GetLLFromVariant(pv), *ppszDisplay, MAX_PROP);
		return S_OK;
	}
	PROPVARIANT propVar;
	PropVariantInit(&propVar);
	_VariantToPropVariant(pv, &propVar);
#ifdef _2000XP
	if (_PSGetPropertyDescription) {
#endif
		IPropertyDescription *pdesc;
		hr = _PSGetPropertyDescription(*ppropKey, IID_PPV_ARGS(&pdesc));
		if SUCCEEDED(hr) {
			hr = pdesc->FormatForDisplay(propVar, (PROPDESC_FORMAT_FLAGS)pdfFlags, ppszDisplay);
			pdesc->Release();
		}
#ifdef _2000XP
	} else {
		*ppszDisplay = (LPWSTR)::CoTaskMemAlloc(MAX_PROP);
		IPropertyUI *pPUI;
		hr = teCreateInstance(CLSID_PropertiesUI, NULL, NULL, IID_PPV_ARGS(&pPUI));
		if SUCCEEDED(hr) {
			hr = pPUI->FormatForDisplay(ppropKey->fmtid, ppropKey->pid, &propVar, pdfFlags, *ppszDisplay, MAX_PROP);
			pPUI->Release();
		}
	}
#endif
	PropVariantClear(&propVar);
	return hr;
}
HRESULT MessageSubPtV(int nFunc, PVOID pObj, MSG *pMsg, VARIANT* pv0)
{
	VARIANT vResult;
	if (g_pOnFunc[nFunc]) {
		VARIANTARG *pv = GetNewVARIANT(6);
		teSetObject(&pv[5], pObj);
		teSetPtr(&pv[4], pMsg->hwnd);
		teSetLong(&pv[3], pMsg->message);
		teSetLong(&pv[2], (LONG)pMsg->wParam);
		CteMemory *pstPt = new CteMemory(2 * sizeof(int), NULL, 1, L"POINT");
		pstPt->SetPoint(pMsg->pt.x, pMsg->pt.y);
		teSetObjectRelease(&pv[1], pstPt);
		VariantCopy(&pv[0], pv0);
		VariantClear(pv0);
		Invoke4(g_pOnFunc[nFunc], &vResult, 6, pv);
		if (vResult.vt != VT_EMPTY) {
			return GetIntFromVariantClear(&vResult);
		}
	}
	VariantClear(pv0);
	return S_FALSE;
}

HRESULT MessageSubPt(int nFunc, PVOID pObj, MSG *pMsg)
{
	VARIANT vResult;
	if (g_pOnFunc[nFunc]) {
		VARIANTARG *pv = GetNewVARIANT(5);
		teSetObject(&pv[4], pObj);
		teSetPtr(&pv[3], pMsg->hwnd);
		teSetLong(&pv[2], pMsg->message);
		teSetLong(&pv[1], (LONG)pMsg->wParam);
		CteMemory *pstPt = new CteMemory(2 * sizeof(int), NULL, 1, L"POINT");
		pstPt->SetPoint(pMsg->pt.x, pMsg->pt.y);
		teSetObjectRelease(&pv[0], pstPt);
		Invoke4(g_pOnFunc[nFunc], &vResult, 5, pv);
		if (vResult.vt != VT_EMPTY) {
			return GetIntFromVariantClear(&vResult);
		}
	}
	return S_FALSE;
}

int GetShellBrowser2(CteShellBrowser *pSB1)
{
	for (size_t i = 0; i < g_ppSB.size(); ++i) {
		CteShellBrowser *pSB = g_ppSB[i];
		if (pSB == pSB1 && !pSB->m_bEmpty) {
			return i;
		}
	}
	return 0;
}

CteShellBrowser* GetNewShellBrowser(CteTabCtrl *pTC)
{
	if (pTC->m_bEmpty) {
		return NULL;
	}
	CteShellBrowser *pSB = NULL;
	BOOL bNew = TRUE;
	for (size_t i = 0; i < g_ppSB.size(); ++i) {
		if (g_ppSB[i]->m_bEmpty) {
			pSB = g_ppSB[i];
			pSB->Init(pTC, FALSE);
			break;
		}
	}
	if (!pSB) {
		pSB = new CteShellBrowser(pTC);
		g_ppSB.push_back(pSB);
		pSB->m_nSB = static_cast<int>(g_ppSB.size());
	}
	if (pTC) {
		TC_ITEM tcItem;
		tcItem.mask = TCIF_PARAM;
		tcItem.lParam = pSB->m_nSB;
		if (pTC) {
			TabCtrl_InsertItem(pTC->m_hwnd, pTC->m_nIndex + 1, &tcItem);
			if (pTC->m_nIndex < 0) {
				pTC->m_nIndex = 0;
			}
		}
		return pSB;
	}
	return NULL;
}

CteTabCtrl* GetNewTC()
{
	for (size_t i = 0; i < g_ppTC.size(); ++i) {
		if (g_ppTC[i]->m_bEmpty) {
			return g_ppTC[i];
		}
	}
	CteTabCtrl *pTC = new CteTabCtrl();
	g_ppTC.push_back(pTC);
	pTC->m_nTC = static_cast<int>(g_ppTC.size());
	return pTC;
}

CteTreeView* GetNewTV()
{
	for (size_t i = 0; i < g_ppTV.size(); ++i) {
		if (g_ppTV[i]->m_bEmpty) {
			return g_ppTV[i];
		}
	}
	CteTreeView *pTV = new CteTreeView();
	g_ppTV.push_back(pTV);
	return pTV;
}

HRESULT ControlFromhwnd(IDispatch **ppdisp, HWND hwnd)
{
	*ppdisp = NULL;
	CteShellBrowser *pSB = SBfromhwnd(hwnd);
	if (pSB) {
		return pSB->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	CteTabCtrl *pTC = TCfromhwnd(hwnd);
	if (pTC) {
		return pTC->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	CteTreeView *pTV = TVfromhwnd(hwnd);
	if (pTV) {
		return pTV->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	if (g_hwndMain == hwnd && g_pTE) {
		return g_pTE->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	if (g_pWebBrowser) {
		HWND hwndBrowser = g_pWebBrowser->get_HWND();
		if (hwnd == hwndBrowser || IsChild(hwndBrowser, hwnd)) {
			return g_pWebBrowser->QueryInterface(IID_PPV_ARGS(ppdisp));
		}
	}
	return E_FAIL;
}

VOID teSetDarkMode(HWND hwnd)
{
	if (_AllowDarkModeForWindow) {
		_AllowDarkModeForWindow(hwnd, g_bDarkMode);
	}
	if (_SetWindowCompositionAttribute) {
		WINCOMPATTRDATA wcpad = { 26, &g_bDarkMode, sizeof(g_bDarkMode) };
		_SetWindowCompositionAttribute(hwnd, &wcpad);
	} else if (_DwmSetWindowAttribute) {
		_DwmSetWindowAttribute(hwnd, 19, &g_bDarkMode, sizeof(g_bDarkMode));
	}
}

BOOL teVerifyVersion(DWORD dwMajor, DWORD dwMinor, DWORD dwBuild)
{
	DWORDLONG dwlConditionMask = 0;
	VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, VER_GREATER_EQUAL);
	VER_SET_CONDITION(dwlConditionMask, VER_MINORVERSION, VER_GREATER_EQUAL);
	VER_SET_CONDITION(dwlConditionMask, VER_BUILDNUMBER, VER_GREATER_EQUAL);
	OSVERSIONINFOEX osvi = { sizeof(OSVERSIONINFOEX), dwMajor, dwMinor, dwBuild };
	return VerifyVersionInfo(&osvi, VER_MAJORVERSION | VER_MINORVERSION | VER_BUILDNUMBER, dwlConditionMask);
}

LRESULT CALLBACK MenuKeyProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = 1;
	if (nCode >= 0) {
		if (nCode == HC_ACTION) {
			g_dwTickKey = GetTickCount();
			MSG msg;
			msg.hwnd = GetFocus();
			msg.message = (lParam & 0xc0000000) ? WM_KEYUP : WM_KEYDOWN;
			msg.wParam = wParam;
			msg.lParam = lParam;
			try {
				if (InterlockedIncrement(&g_nProcKey) < 5) {
					MessageSub(TE_OnKeyMessage, g_pTE, &msg, (HRESULT *)&lResult);
				}
			} catch(...) {
				g_nException = 0;
#ifdef _DEBUG
				g_strException = L"MenuKeyProc";
#endif
			}
			::InterlockedDecrement(&g_nProcKey);
		}
	}
	return lResult ? CallNextHookEx(g_hMenuKeyHook, nCode, wParam, lParam) : TRUE;
}

#ifdef _DEBUG
LRESULT CALLBACK MenuGMProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = 1;
	if (nCode >= 0) {
		if (nCode == HC_ACTION) {
			MSG *msg = (MSG *)lParam;
			if(msg) {
				WCHAR pszNum[99];
				swprintf_s(pszNum, 99, L"%x\n", msg->message);
//				::OutputDebugString(pszNum);
			}
		}
	}
	return lResult ? CallNextHookEx(g_hMenuGMHook, nCode, wParam, lParam) : TRUE;
}
#endif
LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	static UINT dwDoubleTime;
	static WPARAM wParam2;

	HRESULT hrResult = S_FALSE;
	try {
		IDispatch *pdisp = NULL;
		if (nCode >= 0 && g_x == MAXINT) {
			if (nCode == HC_ACTION) {
				g_dwTickKey = 0;
				if (!g_nDropState) {
					if (g_bDragging) {
						if (wParam == WM_LBUTTONUP || wParam == WM_RBUTTONUP) {
							g_bDragging = FALSE;
						}
					}
					if (g_pOnFunc[TE_OnMouseMessage]) {
						MOUSEHOOKSTRUCTEX *pMHS = (MOUSEHOOKSTRUCTEX *)lParam;
						if (ControlFromhwnd(&pdisp, pMHS->hwnd) == S_OK) {
							try {
								if (InterlockedIncrement(&g_nProcMouse) < 5 || wParam != wParam2) {
									wParam2 = wParam;
									CteTreeView *pTV = NULL;
									TVHITTESTINFO ht;
									pTV = TVfromhwnd(pMHS->hwnd);
									if (pTV) {
										if (wParam != WM_MOUSEMOVE && wParam != WM_MOUSEWHEEL) {
											ht.pt.x = pMHS->pt.x;
											ht.pt.y = pMHS->pt.y;
											ht.flags = TVHT_ONITEM;
											ScreenToClient(pMHS->hwnd, &ht.pt);
											TreeView_HitTest(pMHS->hwnd, &ht);
											if (wParam == WM_LBUTTONDOWN || wParam == WM_RBUTTONDOWN ||
												wParam == WM_MBUTTONDOWN || wParam == WM_XBUTTONDOWN) {
												g_hItemDown = ht.hItem;
											} else if (wParam == WM_LBUTTONUP || wParam == WM_LBUTTONDBLCLK ||
												wParam == WM_RBUTTONUP || wParam == WM_RBUTTONDBLCLK ||
												wParam == WM_MBUTTONUP || wParam == WM_MBUTTONDBLCLK ||
												wParam == WM_XBUTTONUP || wParam == WM_XBUTTONDBLCLK) {
												if (ht.hItem == g_hItemDown) {
													SetFocus(pMHS->hwnd);
													if (ht.flags & TVHT_ONITEM) {
														TreeView_SelectItem(pMHS->hwnd, ht.hItem);
													}
												}
											}
										}
									}
									MSG msg;
									msg.hwnd = pMHS->hwnd;
									msg.message = (LONG)wParam;
									if (msg.message == WM_LBUTTONDOWN) {
										CHAR szClassA[MAX_CLASS_NAME];
										GetClassNameA(msg.hwnd, szClassA, MAX_CLASS_NAME);
										if (lstrcmpA(szClassA, WC_LISTVIEWA) == 0) {
											msg.hwnd = WindowFromPoint(pMHS->pt);
										} else if (lstrcmpA(szClassA, "DirectUIHWND") == 0) {
											DWORD dwTick = GetTickCount();
											if (dwTick - dwDoubleTime < GetDoubleClickTime()) {
												msg.message = WM_LBUTTONDBLCLK;
											} else {
												dwDoubleTime = dwTick;
											}
											msg.hwnd = WindowFromPoint(pMHS->pt);
										}
									}

									msg.wParam = pMHS->mouseData;
									msg.pt.x = pMHS->pt.x;
									msg.pt.y = pMHS->pt.y;
									hrResult = MessageSubPt(TE_OnMouseMessage, pdisp, &msg);
#ifdef _2000XP
									if (hrResult != S_OK && pTV) {
										if (wParam == WM_LBUTTONUP || wParam == WM_LBUTTONDBLCLK || wParam == WM_MBUTTONUP) {
											if (pTV->m_pShellNameSpace && ht.hItem) {
												NSTCSTYLE nStyle = NSTCECT_LBUTTON;
												if (wParam == WM_LBUTTONDBLCLK) {
													nStyle = NSTCECT_LBUTTON | NSTCECT_DBLCLICK;
												} else if (wParam == WM_MBUTTONUP) {
													nStyle = NSTCECT_MBUTTON;
												}
												hrResult = pTV->OnItemClick(NULL, ht.flags, nStyle);
											}
										}
									}
#endif
								}
							} catch(...) {
								g_nException = 0;
#ifdef _DEBUG
								g_strException = L"MouseProc1";
#endif
							}
							::InterlockedDecrement(&g_nProcMouse);
							pdisp->Release();
						}
					}
				}
			}
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"MouseProc2";
#endif
	}
	return (hrResult != S_OK) ? CallNextHookEx(g_hMouseHook, nCode, wParam, lParam) : TRUE;
}

#ifdef _2000XP
LRESULT CALLBACK TETVProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	BOOL bHandled = FALSE;
	try {
		CteTreeView *pTV = (CteTreeView *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (msg == WM_NOTIFY) {
			LPNMTVCUSTOMDRAW lptvcd = (LPNMTVCUSTOMDRAW)lParam;
/// Custom Draw
			if (lptvcd->nmcd.hdr.code == NM_CUSTOMDRAW) {
				if (g_pOnFunc[TE_OnItemPrePaint] || g_pOnFunc[TE_OnItemPostPaint]) {
					if (lptvcd->nmcd.dwDrawStage == CDDS_PREPAINT) {
						return CDRF_NOTIFYITEMDRAW;
					}
					if (lptvcd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
						LRESULT lRes = CDRF_DODEFAULT;
						teCustomDraw(TE_OnItemPrePaint, NULL, pTV, NULL, &lptvcd->nmcd, lptvcd, &lRes);
						return lRes;
					}
					if (lptvcd->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT) {
						LRESULT lRes = CDRF_DODEFAULT;
						teCustomDraw(TE_OnItemPostPaint, NULL, pTV, NULL, &lptvcd->nmcd, lptvcd, &lRes);
						return lRes;
					}
				}
			}
			if (lptvcd->nmcd.hdr.code == NM_RCLICK) {
				try {
					if (InterlockedIncrement(&g_nProcTV) < 5) {
						MSG msg1;
						msg1.hwnd = hwnd;
						msg1.message = WM_CONTEXTMENU;
						msg1.wParam = (WPARAM)hwnd;
						GetCursorPos(&msg1.pt);
						if (MessageSubPt(TE_OnShowContextMenu, pTV, &msg1) == S_OK) {
							bHandled = TRUE;
						}
					}
				} catch(...) {
					g_nException = 0;
#ifdef _DEBUG
					g_strException = L"NM_RCLICK";
#endif
				}
				::InterlockedDecrement(&g_nProcTV);
			}
		}
		return bHandled ? 1 : DefSubclassProc(hwnd, msg, wParam, lParam);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"TETVProc";
#endif
	}
	return 0;
}
#endif

LRESULT CALLBACK TETVProc2(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	try {
		CteTreeView *pTV = (CteTreeView *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (msg == TVM_SETBKCOLOR) {
			if (lParam != SendMessage(hwnd, TVM_GETBKCOLOR, 0, 0)) {
				teSetTreeTheme(pTV->m_hwndTV, (COLORREF)lParam);
			}
		}
		if (pTV->m_bMain) {
			if (msg == WM_ENTERMENULOOP) {
				pTV->m_bMain = FALSE;
			}
		} else {
			if (msg == WM_EXITMENULOOP) {
				pTV->m_bMain = TRUE;
			} else if (msg >= TV_FIRST && msg < TV_FIRST + 0x100) {
				return 0;
			}
		}
		if (msg == WM_COMMAND) {
			LRESULT Result = S_FALSE;
			try {
				if (InterlockedIncrement(&g_nProcTV) == 1) {
					Result = teDoCommand(pTV, hwnd, msg, wParam, lParam);
				}
			} catch(...) {
				g_nException = 0;
#ifdef _DEBUG
				g_strException = L"TETVProc2/DoCommand";
#endif
			}
			::InterlockedDecrement(&g_nProcTV);
			if (Result == 0) {
				return 0;
			}
		}
#ifdef _2000XP
		if (msg == TVM_INSERTITEM) {
			TVINSERTSTRUCT *pTVInsert = (TVINSERTSTRUCT *)lParam;
			if (pTVInsert->item.cChildren == 1) {
				pTVInsert->item.cChildren = I_CHILDRENCALLBACK;
			}
		}
#endif
		return DefSubclassProc(hwnd, msg, wParam, lParam);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"TETVProc2";
#endif
	}
	return 0;
}

VOID CALLBACK teTimerProcEnableWindow(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	try {
		KillTimer(hwnd, idEvent);
		if (hwnd == (HWND)idEvent) {
			EnableWindow(hwnd, TRUE);
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teTimerProcEnableWindow";
#endif
	}
}

LRESULT CALLBACK TELVProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	static HWND hwndEdit = NULL;
	static FolderItem *pidEdit = NULL;

	CteShellBrowser *pSB = (CteShellBrowser *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
	try {
		BOOL bDoCallProc = TRUE;
		LRESULT lResult = S_FALSE;
		if (msg == WM_NOTIFY) {
			if (((LPNMHDR)lParam)->code == LVN_GETDISPINFO) {
				NMLVDISPINFO *lpDispInfo = (NMLVDISPINFO *)lParam;
				if (lpDispInfo->item.mask & LVIF_TEXT) {
					if (pSB->m_param[SB_ViewMode] == FVM_DETAILS || lpDispInfo->item.iSubItem == 0) {
						IDispatch *pdisp = (size_t)lpDispInfo->item.iSubItem < pSB->m_ppColumns.size() ? pSB->m_ppColumns[lpDispInfo->item.iSubItem] : NULL;
						if (pdisp) {
							bDoCallProc = FALSE;
							lResult = DefSubclassProc(hwnd, msg, wParam, lParam);
							if (teHasProperty(pdisp, L"push")) {
								VARIANT v;
								VariantInit(&v);
								for (int nLen = teGetObjectLength(pdisp); --nLen >= 0;) {
									if SUCCEEDED(teGetPropertyAt(pdisp, nLen, &v)) {
										IDispatch *pdisp2;
										if (GetDispatch(&v, &pdisp2)) {
											if (pSB->ReplaceColumns(pdisp2, lpDispInfo)) {
												nLen = 0;
											}
											pdisp2->Release();
										}
										VariantClear(&v);
									}
								}
							} else {
								pSB->ReplaceColumns(pdisp, lpDispInfo);
							}
						}
						if (lpDispInfo->item.iSubItem == pSB->m_nFolderSizeIndex) {
							if (bDoCallProc) {
								bDoCallProc = FALSE;
								lResult = DefSubclassProc(hwnd, msg, wParam, lParam);
							}
							IFolderView *pFV;
							LPITEMIDLIST pidl;
							if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
								if SUCCEEDED(pFV->Item(lpDispInfo->item.iItem, &pidl)) {
									pSB->SetFolderSize(pSB->m_pSF2, pidl, lpDispInfo->item.pszText, lpDispInfo->item.cchTextMax);
									teCoTaskMemFree(pidl);
								}
								pFV->Release();
							}
						}
						if (lpDispInfo->item.iSubItem == pSB->m_nLabelIndex && g_pOnFunc[TE_Labels]) {
							IFolderView *pFV;
							LPITEMIDLIST pidl;
							if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
								if SUCCEEDED(pFV->Item(lpDispInfo->item.iItem, &pidl)) {
									pSB->SetLabel(pidl, lpDispInfo->item.pszText, lpDispInfo->item.cchTextMax);
									teCoTaskMemFree(pidl);
									pFV->Release();
								}
							}
							return 0;
						}
						if (lpDispInfo->item.iSubItem == pSB->m_nSizeIndex && pSB->GetSizeFormat()) {
							if (bDoCallProc) {
								bDoCallProc = FALSE;
								lResult = DefSubclassProc(hwnd, msg, wParam, lParam);
							}
							if (lpDispInfo->item.pszText[0]) {
								IFolderView *pFV;
								LPITEMIDLIST pidl;
								if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
									if SUCCEEDED(pFV->Item(lpDispInfo->item.iItem, &pidl)) {
										pSB->SetSize(pidl, lpDispInfo->item.pszText, lpDispInfo->item.cchTextMax);
									}
									teCoTaskMemFree(pidl);
									pFV->Release();
								}
							}
						}
						if (lpDispInfo->item.iSubItem == pSB->m_nLinkTargetIndex) {
							if (bDoCallProc) {
								bDoCallProc = FALSE;
								lResult = DefSubclassProc(hwnd, msg, wParam, lParam);
							}
							if (!lpDispInfo->item.pszText[0]) {
								IFolderView *pFV;
								LPITEMIDLIST pidl;
								if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
									if SUCCEEDED(pFV->Item(lpDispInfo->item.iItem, &pidl)) {
										WIN32_FIND_DATA wfd;
										if SUCCEEDED(SHGetDataFromIDList(pSB->m_pSF2, pidl, SHGDFIL_FINDDATA, &wfd, sizeof(WIN32_FIND_DATA))) {
											if (wfd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) {
												BSTR bs;
												if SUCCEEDED(teGetDisplayNameBSTR(pSB->m_pSF2, pidl, SHGDN_FORPARSING, &bs)) {
													teGetJunctionLinkTarget(bs, &lpDispInfo->item.pszText, lpDispInfo->item.cchTextMax);
													teSysFreeString(&bs);
												}
											}
										}
									}
									teCoTaskMemFree(pidl);
									pFV->Release();
								}
							}
						}
					}
					if (g_bsDateTimeFormat && (size_t)lpDispInfo->item.iSubItem < pSB->m_pDTColumns.size()) {
						UINT ix = pSB->m_pDTColumns[lpDispInfo->item.iSubItem];
						if (ix) {
							VARIANT v;
							VariantInit(&v);
							IFolderView *pFV;
							LPITEMIDLIST pidl;
							if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
								if SUCCEEDED(pFV->Item(lpDispInfo->item.iItem, &pidl)) {
									SHCOLUMNID scid;
									if SUCCEEDED(pSB->m_pSF2->MapColumnToSCID(ix, &scid)) {
										if FAILED(pSB->m_pSF2->GetDetailsEx(pidl, &scid, &v)) {
											LPITEMIDLIST pidlFull = ILCombine(pSB->m_pidl, pidl);// for Search folder (search-ms:)
											IShellFolder2 *pSF2;
											LPCITEMIDLIST pidlPart;
											if SUCCEEDED(SHBindToParent(pidlFull, IID_PPV_ARGS(&pSF2), &pidlPart)) {
												pSF2->GetDetailsEx(pidlPart, &scid, &v);
												pSF2->Release();
											}
											teCoTaskMemFree(pidlFull);
										}
									}
									teCoTaskMemFree(pidl);
									pFV->Release();
								}
							}
							if (v.vt == VT_DATE) {
								LPWSTR pszDisplay;
								if SUCCEEDED(tePSFormatForDisplay(NULL, &v, PDFF_DEFAULT, &pszDisplay, pSB->GetSizeFormat())) {
									lstrcpyn(lpDispInfo->item.pszText, pszDisplay, lpDispInfo->item.cchTextMax);
									teCoTaskMemFree(pszDisplay);
								}
								VariantClear(&v);
								return 0;
							}
							if (v.vt != VT_EMPTY) {
								VariantClear(&v);
								if ((size_t)lpDispInfo->item.iSubItem < pSB->m_pDTColumns.size()) {
									pSB->m_pDTColumns[lpDispInfo->item.iSubItem] = 0;
								}
							}
						}
					}
				}
				if (lpDispInfo->item.mask & LVIF_IMAGE && g_pOnFunc[TE_OnHandleIcon]) {
					IFolderView *pFV;
					LPITEMIDLIST pidl;
					if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
						VARIANT vResult;
						VariantInit(&vResult);
						if SUCCEEDED(pFV->Item(lpDispInfo->item.iItem, &pidl)) {
							VARIANTARG *pv = GetNewVARIANT(3);
							teSetObject(&pv[2], pSB);
							LPITEMIDLIST pidlFull = ILCombine(pSB->m_pidl, pidl);
							teSetIDListRelease(&pv[1], &pidlFull);
							teCoTaskMemFree(pidl);
							teSetLong(&pv[0], lpDispInfo->item.iItem);
							Invoke4(g_pOnFunc[TE_OnHandleIcon], &vResult, 3, pv);
						}
						pFV->Release();
						if (GetIntFromVariantClear(&vResult)) {
							lpDispInfo->item.mask &= ~LVIF_IMAGE;
							if (bDoCallProc) {
								bDoCallProc = FALSE;
								lResult = DefSubclassProc(hwnd, msg, wParam, lParam);
							}
							lpDispInfo->item.mask |= LVIF_IMAGE;
							lpDispInfo->item.iImage = -1;
						}
					}
				}
			} else if (((LPNMHDR)lParam)->code == LVN_COLUMNCLICK) {
				if (g_pOnFunc[TE_OnColumnClick]) {
					LPNMLISTVIEW pnmv = (LPNMLISTVIEW)lParam;
					HWND hHeader = ListView_GetHeader(pSB->m_hwndLV);
					if (hHeader) {
						UINT nHeader = Header_GetItemCount(hHeader);
						if (nHeader) {
							int *piOrderArray = new int[nHeader];
							ListView_GetColumnOrderArray(pSB->m_hwndLV, nHeader, piOrderArray);
							VARIANTARG *pv = GetNewVARIANT(2);
							teSetObject(&pv[1], pSB);
							teSetLong(&pv[0], piOrderArray[pnmv->iSubItem]);
							delete [] piOrderArray;
							VARIANT vResult;
							VariantInit(&vResult);
							Invoke4(g_pOnFunc[TE_OnColumnClick], &vResult, 2, pv);
							if (vResult.vt != VT_EMPTY) {
								if (!GetIntFromVariantClear(&vResult)) {
									return 1;
								}
							}
						}
					}
					pSB->FixColumnEmphasis();
				}
			} else if (((LPNMHDR)lParam)->code == LVN_INCREMENTALSEARCH) {
				NMLVFINDITEM *lpFindItem = (NMLVFINDITEM *)lParam;
				if (lpFindItem->lvfi.flags & (LVFI_STRING | LVFI_PARTIAL)) {
					DoStatusText(pSB, lpFindItem->lvfi.psz, 0);
				}
			} else if (((LPNMHDR)lParam)->code == LVN_BEGINLABELEDIT) {
				VARIANT vResult;
				VariantInit(&vResult);
				if SUCCEEDED(DoFunc1(TE_OnBeginLabelEdit, pSB, &vResult)) {
					if (GetIntFromVariantClear(&vResult)) {
						return TRUE;
					}
				}
			} else if (((LPNMHDR)lParam)->code == LVN_ENDLABELEDIT) {
				if (g_nWindowTheme == 1) { // Bug: Classic style and Shift + Tab
					HWND hHeader = ListView_GetHeader(pSB->m_hwndLV);
					if (hHeader) {
						EnableWindow(hHeader, FALSE);
						SetTimer(hHeader, (UINT_PTR)hHeader, 100, teTimerProcEnableWindow);
					}
				}
				NMLVDISPINFO *lpDispInfo = (NMLVDISPINFO *)lParam;
				if (g_pOnFunc[TE_OnEndLabelEdit]) {
					VARIANTARG *pv = GetNewVARIANT(2);
					teSetObject(&pv[1], pSB);
					teSetSZ(&pv[0], lpDispInfo->item.pszText);
					VARIANT vResult;
					VariantInit(&vResult);
					Invoke4(g_pOnFunc[TE_OnEndLabelEdit], &vResult, 2, pv);
					if (GetIntFromVariantClear(&vResult) && lpDispInfo->item.pszText) {
						lpDispInfo->item.pszText[0] = NULL;
					}
				}
				if (lpDispInfo->item.pszText) {
					if (lpDispInfo->item.pszText[0] == '.' && !StrChr(&lpDispInfo->item.pszText[1], '.')) {
						int i = lstrlen(lpDispInfo->item.pszText);
						if (i > 1 && i < lpDispInfo->item.cchTextMax - 1) {
							lpDispInfo->item.pszText[i] = '.';
							lpDispInfo->item.pszText[i + 1] = NULL;
						}
					}
				}
			} else if (((LPNMHDR)lParam)->code == LVN_GETINFOTIP) {
				LPNMLVGETINFOTIP lpInfotip = ((LPNMLVGETINFOTIP)lParam);
				if (g_pOnFunc[TE_OnToolTip] && lpInfotip->cchTextMax) {
					VARIANT vResult;
					VariantInit(&vResult);
					VARIANTARG *pv = GetNewVARIANT(2);
					teSetObject(&pv[1], pSB);
					teSetLong(&pv[0], lpInfotip->iItem);
					Invoke4(g_pOnFunc[TE_OnToolTip], &vResult, 2, pv);
					if (vResult.vt == VT_BSTR) {
						lstrcpyn(lpInfotip->pszText, vResult.bstrVal, lpInfotip->cchTextMax);
						lResult = 0;
					}
					VariantClear(&vResult);
				}
			} else if (((LPNMHDR)lParam)->code == HDN_ITEMCHANGED) {
				pSB->FixColumnEmphasis();
			}
/// Custom Draw
			if (pSB->m_pShellView) {
				LPNMLVCUSTOMDRAW lplvcd = (LPNMLVCUSTOMDRAW)lParam;
				if (lplvcd->nmcd.hdr.code == NM_CUSTOMDRAW) {
					if (g_nWindowTheme != 2 && teIsDarkColor(pSB->m_clrBk) && lplvcd->dwItemType == LVCDI_GROUP) { //Fix groups in dark background
						if (lplvcd->nmcd.dwDrawStage == CDDS_PREPAINT) {
							FillRect(lplvcd->nmcd.hdc, &lplvcd->rcText, GetStockBrush(WHITE_BRUSH));
						} else if (lplvcd->nmcd.dwDrawStage == CDDS_POSTPAINT) {
							int w = lplvcd->rcText.right - lplvcd->rcText.left;
							int h = lplvcd->rcText.bottom - lplvcd->rcText.top;
							BYTE r0 = GetRValue(pSB->m_clrBk);
							BYTE g0 = GetGValue(pSB->m_clrBk);
							BYTE b0 = GetBValue(pSB->m_clrBk);
							BITMAPINFO bmi;
							RGBQUAD *pcl = NULL;
							::ZeroMemory(&bmi.bmiHeader, sizeof(BITMAPINFOHEADER));
							bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
							bmi.bmiHeader.biWidth = w;
							bmi.bmiHeader.biHeight = -(LONG)h;
							bmi.bmiHeader.biPlanes = 1;
							bmi.bmiHeader.biBitCount = 32;
							HBITMAP hBM = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, (void **)&pcl, NULL, 0);
							HDC hmdc = CreateCompatibleDC(lplvcd->nmcd.hdc);
							HGDIOBJ hOld = SelectObject(hmdc, hBM);
							BitBlt(hmdc, 0, 0, w, h, lplvcd->nmcd.hdc, lplvcd->rcText.left, lplvcd->rcText.top, NOTSRCCOPY);
							for (int i = w * h; --i >= 0; ++pcl) {
								if (pcl->rgbRed || pcl->rgbGreen || pcl->rgbBlue) {
									WORD cl = pcl->rgbRed > pcl->rgbGreen ? pcl->rgbRed : pcl->rgbGreen;
									if (cl < pcl->rgbBlue) {
										cl = pcl->rgbBlue;
									}
									cl += 48;
									pcl->rgbRed = pcl->rgbGreen = pcl->rgbBlue = cl > 0xff ? 0xff : cl;
								} else {
									pcl->rgbRed = r0;
									pcl->rgbGreen = g0;
									pcl->rgbBlue = b0;
								}
							}
							BitBlt(lplvcd->nmcd.hdc, lplvcd->rcText.left, lplvcd->rcText.top, w, h, hmdc, 0, 0, SRCCOPY);
							SelectObject(hmdc, hOld);
							DeleteDC(hmdc);
							DeleteObject(hBM);
						}
					}
					if (g_pOnFunc[TE_OnItemPrePaint] || g_pOnFunc[TE_OnItemPostPaint]) {
						if (lplvcd->nmcd.dwDrawStage == CDDS_PREPAINT) {
							return CDRF_NOTIFYITEMDRAW;
						}
						if (lplvcd->dwItemType != LVCDI_GROUP) {
							if (lplvcd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
								LRESULT lRes = CDRF_DODEFAULT;
								teCustomDraw(TE_OnItemPrePaint, pSB, NULL, NULL, &lplvcd->nmcd, lplvcd, &lRes);
								return lRes;
							}
							if (lplvcd->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT) {
								LRESULT lRes = CDRF_DODEFAULT;
								teCustomDraw(TE_OnItemPostPaint, pSB, NULL, NULL, &lplvcd->nmcd, lplvcd, &lRes);
								return lRes;
							}
						}
					}
				}
			}
		}
		if (msg == WM_CONTEXTMENU || msg == SB_SETTEXT || msg == WM_COMMAND || msg == WM_COPYDATA) {
			if (pSB->m_hwndAlt && msg == WM_CONTEXTMENU) {
				return TRUE;
			}
			try {
				if (InterlockedIncrement(&g_nProcFV) < 5) {
					switch (msg) {
						case WM_CONTEXTMENU:
							IDispatch *pdisp;
							if (ControlFromhwnd(&pdisp, hwnd) == S_OK) {
								MSG msg1;
								msg1.hwnd = hwnd;
								msg1.message = msg;
								msg1.wParam = wParam;
								HWND hwndLV = pSB->m_hwndLV;
								if ((lParam & MAXDWORD) == MAXDWORD) {
									msg1.pt.x = 0;
									msg1.pt.y = 0;
									if (hwndLV) {
										ClientToScreen(hwndLV, &msg1.pt);
										int nSelected = ListView_GetNextItem(hwndLV, -1, LVNI_ALL | LVNI_SELECTED);
										if (nSelected >= 0) {
											RECT rc;
											ListView_GetItemRect(hwndLV, nSelected, &rc, LVIR_ICON);
											OffsetRect(&rc, msg1.pt.x, msg1.pt.y);
											msg1.pt.x = (rc.left + rc.right) / 2;
											msg1.pt.y = (rc.top + rc.bottom) / 2;
										}
									} else {
										ClientToScreen(pSB->m_hwndDV, &msg1.pt);
									}
								} else {
									if (GetKeyState(VK_SHIFT) < 0 && ListView_GetNextItem(hwndLV, -1, LVNI_ALL | LVNI_SELECTED) >= 0) {
										int nFocused = ListView_GetNextItem(hwndLV, -1, LVNI_ALL | LVNI_FOCUSED);
										if (!ListView_GetItemState(hwndLV, nFocused, LVIS_SELECTED)) {
											ListView_SetItemState(hwndLV, ListView_GetNextItem(hwndLV, -1, LVNI_ALL | LVNI_SELECTED), LVIS_FOCUSED, LVIS_FOCUSED);
											ListView_SetItemState(hwndLV, nFocused, 0, LVIS_FOCUSED);
										}
									}
									msg1.pt.x = GET_X_LPARAM(lParam);
									msg1.pt.y = GET_Y_LPARAM(lParam);
								}
								lResult = MessageSubPt(TE_OnShowContextMenu, pdisp, &msg1);
							}
							break;
						case SB_SETTEXT:
							pSB->SetStatusTextSB((LPCWSTR)lParam);
							break;
						case WM_COMMAND:
							lResult = teDoCommand(pSB, hwnd, msg, wParam, lParam);
							break;
						case WM_COPYDATA:
							if (g_pOnFunc[TE_OnSystemMessage]) {
								MSG msg1;
								msg1.hwnd = hwnd;
								msg1.message = msg;
								msg1.wParam = wParam;
								msg1.lParam = lParam;
								MessageSub(TE_OnSystemMessage, pSB, &msg1, (HRESULT *)&lResult);
							}
							break;
					}//end_switch
				}
			} catch(...) {
				g_nException = 0;
#ifdef _DEBUG
				g_strException = L"TELVProc";
#endif
			}
			::InterlockedDecrement(&g_nProcFV);
		}
		if (msg == WM_WINDOWPOSCHANGED) {
			pSB->SetFolderFlags(FALSE);
		}
		if (hwnd != pSB->m_hwndDV) {
			return 0;
		}
		if (bDoCallProc && lResult) {
			lResult = DefSubclassProc(hwnd, msg, wParam, lParam);
		}
		if (msg == WM_COMMAND) {
			IFolderView2 *pFV2;
			if (pSB->m_hwndLV && SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2)))) {
				UINT uViewMode;
				if SUCCEEDED(pFV2->GetCurrentViewMode(&uViewMode)) {
					if (uViewMode == FVM_ICON || uViewMode == FVM_SMALLICON || uViewMode == FVM_TILE) {
						DWORD dwFlags;
						if SUCCEEDED(pFV2->GetCurrentFolderFlags(&dwFlags)) {
							pSB->m_param[SB_FolderFlags] = (pSB->m_param[SB_FolderFlags] & ~(FWF_AUTOARRANGE | FWF_SNAPTOGRID)) | (dwFlags & (FWF_AUTOARRANGE | FWF_SNAPTOGRID));
						}
					}
				}
				pFV2->Release();
			}
		}
		return lResult;
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		DWORD dw = 0;
		if (msg == WM_NOTIFY) {
			LPNMTVCUSTOMDRAW lptvcd = (LPNMTVCUSTOMDRAW)lParam;
			dw = lptvcd->nmcd.hdr.code;
		}
		swprintf_s(g_pszException, MAX_PATH, L"TELVProc/%x %x", msg, dw);
		g_strException = g_pszException;
#endif
	}
	return 0;
}

LRESULT CALLBACK TELVProc2(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	HWND hTree;
	try {
		CteShellBrowser *pSB = (CteShellBrowser *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		switch (msg)
		{
		case LVM_SETTEXTCOLOR:
			pSB->m_clrText = (COLORREF)lParam;
			break;
		case LVM_SETBKCOLOR:
			if (pSB->m_clrBk != (COLORREF)lParam) {
				pSB->m_clrBk = (COLORREF)lParam;
				pSB->SetTheme();
			}
			if (hTree = teFindChildByClassA(pSB->m_hwnd, WC_TREEVIEWA)) {
				teSetTreeTheme(hTree, (COLORREF)lParam);
			}
			break;
		case LVM_SETTEXTBKCOLOR:
			pSB->m_clrTextBk = (COLORREF)lParam;
			break;
		case LVM_SETCOLUMNWIDTH:
			if (pSB->m_param[SB_ViewMode] == FVM_LIST) {
				if (lParam < 0) {
					pSB->m_bSetListColumnWidth = TRUE;
					pSB->SetListColumnWidth();
					return 0;
				}
				pSB->m_bSetListColumnWidth = FALSE;
			}
			break;
		case WM_NOTIFY:
			if (_AllowDarkModeForWindow && ((LPNMHDR)lParam)->code == NM_CUSTOMDRAW) {
				if (teIsDarkColor(pSB->m_clrBk)) {
					LPNMCUSTOMDRAW pnmcd = (LPNMCUSTOMDRAW)lParam;
					if (pnmcd->dwDrawStage == CDDS_PREPAINT) {
						return CDRF_NOTIFYITEMDRAW;
					}
					if (pnmcd->dwDrawStage == CDDS_ITEMPREPAINT) {
						SetTextColor(pnmcd->hdc, 0xffffff);
						return CDRF_DODEFAULT;
					}
				}
			}
			break;
		case WM_SETTINGCHANGE:
			pSB->InitFolderSize();
			ArrangeWindow();
			break;
		}
	} catch (...) {
		g_nException = 0;
	}
	return DefSubclassProc(hwnd, msg, wParam, lParam);
}

VOID teSetTreeWidth(CteShellBrowser	*pSB, HWND hwnd, LPARAM lParam)
{
	int nWidth = pSB->m_param[SB_TreeWidth] + GET_X_LPARAM(lParam) - g_x;
	RECT rc;
	GetWindowRect(hwnd, &rc);
	int nMax = rc.right - rc.left - 1;
	if (nWidth > nMax) {
		nWidth = nMax;
	} else if (nWidth < 3) {
		nWidth = 3;
	} else {
		g_x =GET_X_LPARAM(lParam);
	}
	if (pSB->m_param[SB_TreeWidth] != nWidth) {
		pSB->m_param[SB_TreeWidth] = nWidth;
		ArrangeWindow();
	}
}

VOID teSetTabWidth(CteTabCtrl *pTC, LPARAM lParam)
{
	int nWidth = pTC->m_param[TC_TabWidth];
	RECT rc;
	GetWindowRect(pTC->m_hwndStatic, &rc);
	int nMax = rc.right - rc.left - 1;
	int nDiff = GET_X_LPARAM(lParam) - g_x;
	nWidth += (pTC->m_param[TC_Align] == 4) ? nDiff : -nDiff;
	if (nWidth > nMax) {
		nWidth = nMax;
	} else if (nWidth < 1) {
		nWidth = 1;
	} else {
		g_x = GET_X_LPARAM(lParam);
	}
	if (pTC->m_param[TC_TabWidth] != nWidth) {
		pTC->m_param[TC_TabWidth] = nWidth;
		ArrangeWindow();
	}
}

VOID ArrangeWindowTC(CteTabCtrl *pTC, RECT *prc)
{
	if (!prc) {
		return;
	}
	if (!pTC->m_bEmpty && pTC->m_bVisible) {
		RECT rcTab;
		int nAlign = g_param[TE_Tab] ? pTC->m_param[TC_Align] : 0;
		MoveWindow(pTC->m_hwndStatic, prc->left, prc->top, prc->right - prc->left, prc->bottom - prc->top, TRUE);
		pTC->LockUpdate();
		try {
			int left = prc->left;
			int top = prc->top;
			OffsetRect(prc, -left, -top);
			CopyRect(&rcTab, prc);
			int nCount = TabCtrl_GetItemCount(pTC->m_hwnd);
			int h = prc->bottom;
			int nBottom = rcTab.bottom;
			if (nCount) {
				DWORD dwStyle = pTC->GetStyle();
				if (nAlign == 4 || nAlign == 5) {
					prc->right = pTC->m_param[TC_TabWidth];
				}
				if ((dwStyle & (TCS_BOTTOM | TCS_BUTTONS)) == (TCS_BOTTOM | TCS_BUTTONS)) {
					SetWindowLong(pTC->m_hwnd, GWL_STYLE, dwStyle & ~TCS_BOTTOM);
					TabCtrl_AdjustRect(pTC->m_hwnd, FALSE, prc);
					SetWindowLong(pTC->m_hwnd, GWL_STYLE, dwStyle);
					int i = rcTab.bottom - prc->bottom;
					prc->bottom -= prc->top - i;
					prc->top = i;
				} else {
					TabCtrl_AdjustRect(pTC->m_hwnd, FALSE, prc);
				}
				h = rcTab.bottom - (prc->bottom - prc->top) - 4;
				switch (nAlign) {
				case 0:						//none
					CopyRect(prc, &rcTab);
					rcTab.top = rcTab.bottom;
					rcTab.left = rcTab.right;
					nBottom = rcTab.bottom;
					break;
				case 2:						//top
					CopyRect(prc, &rcTab);
					rcTab.bottom = h;
					prc->top += h;
					nBottom = h;
					break;
				case 3:						//bottom
					CopyRect(prc, &rcTab);
					rcTab.top = rcTab.bottom - h;
					prc->bottom -= h;
					nBottom = rcTab.bottom;
					break;
				case 4:						//left
					SetRect(prc, pTC->m_param[TC_TabWidth], 0, rcTab.right, rcTab.bottom);
					rcTab.right = prc->left;
					nBottom = h;
					break;
				case 5:						//Right
					SetRect(prc, 0, 0, rcTab.right - pTC->m_param[TC_TabWidth], rcTab.bottom);
					rcTab.left = prc->right;
					nBottom = h;
					break;
				}
			}
			pTC->m_nScrollWidth = 0;
			if (nAlign == 4 || nAlign == 5) {
				int h2 = h - rcTab.bottom;
				if (h2 <= 0) {
					h2 = 0;
				} else {
					pTC->m_nScrollWidth = GetSystemMetrics(SM_CXHSCROLL);
				}
				if (h2 < pTC->m_si.nPos) {
					pTC->m_si.nPos = h2;
				}
				ShowScrollBar(pTC->m_hwndButton, SB_VERT, h2 ? TRUE : FALSE);
				pTC->m_si.fMask = SIF_RANGE | SIF_POS | SIF_PAGE;
				pTC->m_si.nMax = h;
				pTC->m_si.nPage = rcTab.bottom;
				SetScrollInfo(pTC->m_hwndButton, SB_VERT, &pTC->m_si, TRUE);
				if (pTC->m_param[TC_Flags] & TCS_FIXEDWIDTH) {
					pTC->SetItemSize();
				}
				rcTab.right -= pTC->m_nScrollWidth;
			} else {
				ShowScrollBar(pTC->m_hwndButton, SB_VERT, FALSE);
				if (pTC->m_si.nTrackPos && pTC->m_param[TC_Flags] & TCS_FIXEDWIDTH) {
					pTC->SetItemSize();
				}
				pTC->m_si.nTrackPos = 0;
				pTC->m_si.nPos = 0;
			}
			if (teSetRect(pTC->m_hwnd, 0, -pTC->m_si.nPos, rcTab.right - rcTab.left, (nBottom - rcTab.top) + pTC->m_si.nPos)) {
				ArrangeWindow();
				if (!(pTC->m_param[TC_Flags] & TCS_MULTILINE)) {
					int i = TabCtrl_GetCurSel(pTC->m_hwnd);
					TabCtrl_SetCurSel(pTC->m_hwnd, 0);
					TabCtrl_SetCurSel(pTC->m_hwnd, i);
				}
			}
			CteShellBrowser *pSB = pTC->GetShellBrowser(pTC->m_nIndex);
			if (!pSB) {
				pSB = GetNewShellBrowser(pTC);
			}
			if (pSB) {
				if (pSB->m_param[SB_TreeAlign] & 2) {
					MoveWindow(pSB->m_pTV->m_hwnd, prc->left, prc->top, pSB->m_param[SB_TreeWidth] - 2, prc->bottom - prc->top, FALSE);
#ifdef _2000XP
					pSB->m_pTV->SetObjectRect();
#endif
					prc->left += pSB->m_param[SB_TreeWidth];
				}
				if (pSB->m_bVisible) {
					teSetParent(pSB->m_pTV->m_hwnd, pSB->m_pTC->m_hwndStatic);
					ShowWindow(pSB->m_pTV->m_hwnd, (pSB->m_param[SB_TreeAlign] & 2) ? SW_SHOWNA : SW_HIDE);
				} else {
					pSB->Show(TRUE, 0);
				}
				pSB->SetFolderFlags(FALSE);
				if (pSB->m_hwnd) {
					CopyRect(&pSB->m_rc, prc);
					pSB->SetObjectRect();
				}
			}
			MoveWindow(pTC->m_hwndButton, rcTab.left, rcTab.top, rcTab.right - rcTab.left, rcTab.bottom - rcTab.top, TRUE);
			BringWindowToTop(pTC->m_hwndStatic);
			while (--nCount >= 0) {
				if (nCount != pTC->m_nIndex) {
					pSB = pTC->GetShellBrowser(nCount);
					if (pSB) {
						if (pSB->m_bVisible) {
							pSB->Show(FALSE, 1);
						}
					}
				}
			}
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"ArrangeWindowTC";
#endif
		}
		pTC->UnlockUpdate();
	}
}

VOID ArrangeWindowEx()
{
	if (!g_bArrange || g_nLockUpdate) {
		return;
	}
	g_bArrange = FALSE;
	DoFunc(TE_OnArrange, g_pTE, E_NOTIMPL);
}

LRESULT CALLBACK TESTProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	LRESULT lResult = 1;
	try {
		CteTabCtrl *pTC = (CteTabCtrl *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		CteShellBrowser	*pSB = pTC->GetShellBrowser(pTC->m_nIndex);
		RECT rc;
		if (pSB && pSB->m_pidl && GetWindowRect(pSB->m_hwnd, &rc) && rc.top != rc.bottom && IsWindowVisible(pSB->m_pTV->m_hwndTV)) {
			switch (msg) {
				case WM_MOUSEMOVE:
					if (g_x == MAXINT) {
						if (IsWindowVisible(pSB->m_hwndDV)) {
							SetCursor(LoadCursor(NULL, IDC_SIZEWE));
						}
					} else {
						teSetTreeWidth(pSB, hwnd, lParam);
					}
					break;
				case WM_LBUTTONDOWN:
					SetCapture(hwnd);
					g_x = GET_X_LPARAM(lParam);
					SetCursor(LoadCursor(NULL, IDC_SIZEWE));
					break;
				case WM_LBUTTONUP:
					if (g_x != MAXINT) {
						ReleaseCapture();
						teSetTreeWidth(pSB, hwnd, lParam);
						g_x = MAXINT;
					}
					break;
			}//end_switch
		}
		return lResult ? DefSubclassProc(hwnd, msg, wParam, lParam) : 0;
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"TESTProc";
#endif
	}
	return 0;
}

LRESULT TETCProc2(CteTabCtrl *pTC, HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = 1;
	switch (msg) {
		case WM_LBUTTONDOWN:
			if ((pTC->m_param[TC_Align] == 4 && WORD(lParam) >= pTC->m_param[TC_TabWidth] - pTC->m_nScrollWidth - 2 && WORD(lParam) < pTC->m_param[TC_TabWidth]) ||
				(pTC->m_param[TC_Align] == 5 && WORD(lParam) < 2)) {
				SetCapture(hwnd);
				g_x = GET_X_LPARAM(lParam);
				SetCursor(LoadCursor(NULL, IDC_SIZEWE));
				lResult = 0;
			}
		case WM_SETFOCUS:
		case WM_RBUTTONDOWN:
			CheckChangeTabTC(hwnd);
			CteShellBrowser *pSB;
			pSB = pTC->GetShellBrowser(pTC->m_nIndex);
			if (pSB) {
				pSB->SetActive(TRUE);
			}
			break;
		case WM_MOUSEMOVE:
			if (g_x == MAXINT) {
				if ((pTC->m_param[TC_Align] == 4 && WORD(lParam) >= pTC->m_param[TC_TabWidth] - pTC->m_nScrollWidth - 2 && WORD(lParam) < pTC->m_param[TC_TabWidth]) ||
					(pTC->m_param[TC_Align] == 5 && WORD(lParam) < 2)) {
					SetCursor(LoadCursor(NULL, IDC_SIZEWE));
				}
			} else {
				if (pTC->m_param[TC_Align] == 4) {
					teSetTabWidth(pTC, lParam);
				}
				lResult = 0;
			}
			break;
		case WM_LBUTTONUP:
			if (g_x != MAXINT) {
				ReleaseCapture();
				teSetTabWidth(pTC, lParam);
				g_x = MAXINT;
			}
			break;
	}
	return lResult;
}

LRESULT CALLBACK TEBTProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	LRESULT lResult = 1;
	try {
		CteTabCtrl *pTC = (CteTabCtrl *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		switch (msg) {
			case WM_NOTIFY:
				switch (((LPNMHDR)lParam)->code) {
					case TCN_SELCHANGE:
						pTC->SetDefault();
						pTC->TabChanged(TRUE);
						lResult = 0;
						break;
					case TCN_SELCHANGING:
						lResult = DoFunc(TE_OnSelectionChanging, pTC, 0);
						break;
					case TTN_GETDISPINFO:
						if (g_pOnFunc[TE_OnToolTip]) {
							VARIANT vResult;
							VariantInit(&vResult);
							VARIANTARG *pv = GetNewVARIANT(2);
							teSetObject(&pv[1], pTC);
							teSetPtr(&pv[0], ((LPNMHDR)lParam)->idFrom);
							Invoke4(g_pOnFunc[TE_OnToolTip], &vResult, 2, pv);
							VARIANT vText;
							teVariantChangeType(&vText, &vResult, VT_BSTR);
							lstrcpyn(g_szText, vText.bstrVal, MAX_STATUS);
							((LPTOOLTIPTEXT)lParam)->lpszText = g_szText;
							VariantClear(&vResult);
							VariantClear(&vText);
						}
						break;
				}
				break;
			case WM_CONTEXTMENU:
				MSG msg1;
				msg1.hwnd = pTC->m_hwnd;
				msg1.message = msg;
				msg1.wParam = wParam;
				msg1.pt.x = GET_X_LPARAM(lParam);
				msg1.pt.y = GET_Y_LPARAM(lParam);
				MessageSubPt(TE_OnShowContextMenu, pTC, &msg1);
				lResult = 0;
				break;
			case WM_VSCROLL:
				switch(LOWORD(wParam)){
					case SB_LINEUP:
						pTC->m_si.nPos -= 16;
						break;
					case SB_LINEDOWN:
						pTC->m_si.nPos += 16;
						break;
					case SB_PAGEUP:
						pTC->m_si.nPos -= pTC->m_si.nPage;
						break;
					case SB_PAGEDOWN:
						pTC->m_si.nPos += pTC->m_si.nPage;
						break;
					case SB_THUMBPOSITION:
					case SB_THUMBTRACK:
						pTC->m_si.nPos = HIWORD(wParam);
						break;
				}//end_switch
				if (pTC->m_si.nPos > pTC->m_si.nMax) {
					pTC->m_si.nPos = pTC->m_si.nMax;
				}
				if (pTC->m_si.nPos < 0) {
					pTC->m_si.nPos = 0;
				}
				SetScrollInfo(pTC->m_hwndButton, SB_VERT, &pTC->m_si, TRUE);
				ArrangeWindow();
				break;
		}//end_switch
		if (lResult) {
			lResult = TETCProc2(pTC, hwnd, msg, wParam, lParam);
		}
		return lResult ? DefSubclassProc(hwnd, msg, wParam, lParam) : 0;
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"TEBTProc";
#endif
	}
	return 0;
}

LRESULT CALLBACK TETCProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	LRESULT Result = 1;
	int nIndex, nCount;
	static BOOL bCancelPaint = FALSE;
	try {
		CteTabCtrl *pTC = (CteTabCtrl *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (hwnd != pTC->m_hwnd) {
			return 0;
		}
		switch (msg) {
			case TCM_DELETEITEM:
				if (lParam) {
					lParam = 0;
				} else {
					CteShellBrowser *pSB = NULL;
					pSB = pTC->GetShellBrowser((int)wParam);
					DefSubclassProc(hwnd, msg, wParam, lParam);
					Result = 0;
					if (pSB) {
						pSB->Close(TRUE);
					}
					nIndex = TabCtrl_GetCurSel(hwnd);
					if (nIndex >= 0) {
						pTC->m_nIndex = nIndex;
					} else {
						nCount = TabCtrl_GetItemCount(hwnd);
						nIndex = pTC->m_nIndex;
						if ((int)wParam > nIndex) {
							--nIndex;
						}
						if (nIndex >= nCount) {
							nIndex = nCount - 1;
						}
						if (nIndex < 0) {
							nIndex = 0;
						}
						pTC->m_nIndex = -1;
						DefSubclassProc(hwnd, TCM_SETCURSEL, nIndex, 0);
					}
					pTC->TabChanged(TRUE);
					ArrangeWindow();
				}
				break;
			case TCM_SETCURSEL:
				if (wParam != (UINT_PTR)TabCtrl_GetCurSel(hwnd)) {
					DefSubclassProc(hwnd, msg, wParam, lParam);
					Result = 0;
					pTC->TabChanged(TRUE);
					if (g_param[TE_Tab] && pTC->m_param[TC_Align] == 1) {
						if (pTC->m_param[TC_Flags] & TCS_SCROLLOPPOSITE) {
							ArrangeWindow();
						}
					}
					if (!pTC->m_nLockUpdate) {
						CteShellBrowser *pSB = pTC->GetShellBrowser(pTC->m_nIndex);
						if (pSB && pSB->m_hwndDV) {
							ArrangeWindowEx();
						}
					}
				}
				break;
			case TCM_SETITEM:
				if (pTC->m_param[TC_Flags] & TCS_FIXEDWIDTH) {
					DefSubclassProc(hwnd, msg, wParam, lParam);
					Result = 0;
					DefSubclassProc(pTC->m_hwnd, TCM_SETITEMSIZE, 0, MAKELPARAM(pTC->m_param[TC_TabWidth], pTC->m_param[TC_TabHeight] + 1));
					DefSubclassProc(pTC->m_hwnd, TCM_SETITEMSIZE, 0, pTC->m_dwSize);
				}
				break;
		}//end_switch

		//Fix for high CPU usage on single line tab.
		if (!(pTC->m_param[TC_Flags] & TCS_MULTILINE)) {
			if (msg == WM_PAINT) {
				if (bCancelPaint) {
					Result = 0;
				}
				bCancelPaint = TRUE;
			} else if (msg != WM_ERASEBKGND && msg != WM_NCPAINT && msg != TCM_GETITEM) {
				bCancelPaint = FALSE;
			}
		}

		if (Result) {
			Result = TETCProc2(pTC, hwnd, msg, wParam, lParam);
		}
		return Result ? DefSubclassProc(hwnd, msg, wParam, lParam) : 0;
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"TETCProc";
#endif
	}
	return 0;
}

HRESULT MessageProc(MSG *pMsg)
{
	HRESULT hrResult = S_FALSE;
	IDispatch *pdisp = NULL;
	IShellBrowser *pSB = NULL;
	IShellView *pSV = NULL;
	if (pMsg->message >= WM_KEYFIRST && pMsg->message <= WM_KEYLAST) {
		try {
			g_dwTickKey = GetTickCount();
			if (pMsg->message == WM_KEYUP && g_szStatus[0]) {
				SetTimer(g_hwndMain, TET_Status, 100, teTimerProc);
			}
			if (InterlockedIncrement(&g_nProcKey) < 5) {
				if (ControlFromhwnd(&pdisp, pMsg->hwnd) == S_OK) {
					MessageSub(TE_OnKeyMessage, pdisp, pMsg, &hrResult);
					if (hrResult != S_OK) {
						if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pSB))) {
							if SUCCEEDED(pSB->QueryActiveShellView(&pSV)) {
								if (pSV->TranslateAcceleratorW(pMsg) == S_OK) {
									hrResult = S_OK;
								}
								pSV->Release();
							}
							pSB->Release();
						} else {
							teTranslateAccelerator(pdisp, pMsg, &hrResult);
						}
					}
					pdisp->Release();
				}
				CteWebBrowser *pWB;
				if (teGetSubWindow(GetParent((HWND)GetWindowLongPtr(GetParent(pMsg->hwnd), GWLP_HWNDPARENT)), &pWB)) {
					teTranslateAccelerator(pWB, pMsg, &hrResult);
				}
				g_bDragging = FALSE;
			}
		} catch(...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"MessageProc2";
#endif
		}
		::InterlockedDecrement(&g_nProcKey);
	}
	return hrResult;
}

VOID Finalize()
{
	try {
		teRevoke(g_pSW);
		for (int i = _countof(g_maps); i--;) {
			delete[] g_maps[i];
		}
		for (int i = MAX_CSIDL2; i--;) {
			teILFreeClear(&g_pidls[i]);
			teSysFreeString(&g_bsPidls[i]);
		}
		teSysFreeString(&g_bsCmdLine);
		teSysFreeString(&g_bsTitle);
		teSysFreeString(&g_bsAnyTextId);
		teSysFreeString(&g_bsAnyText);
		teSysFreeString(&g_bsName);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"Finalize";
#endif
	}
	try {
		::OleUninitialize();
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"OleUninitialize";
#endif
	}
}

HRESULT teExtract(IStorage *pStorage, LPWSTR lpszFolderPath, IProgressDialog *ppd, int *pnItems, int nCount, int nBase)
{
	STATSTG statstg;
	IEnumSTATSTG *pEnumSTATSTG = NULL;
	BYTE pszData[SIZE_BUFF];
	IStream *pStream;
	ULONG uRead;
	DWORD dwWriteByte;
#ifdef _2000XP
	IShellFolder2 *pSF2;
	LPITEMIDLIST pidl;
	ULONG chEaten;
	ULONG dwAttributes;
#endif
	if (ppd->HasUserCancelled()) {
		return E_ABORT;
	}
	if (nCount >= 0) {
		CreateDirectory(lpszFolderPath, NULL);
	}
	HRESULT hr = pStorage->EnumElements(NULL, NULL, NULL, &pEnumSTATSTG);
	while (SUCCEEDED(hr) && pEnumSTATSTG->Next(1, &statstg, NULL) == S_OK) {
		BSTR bsPath = NULL;
		tePathAppend(&bsPath, lpszFolderPath, statstg.pwcsName);
		if (statstg.type == STGTY_STORAGE) {
			IStorage *pStorageNew;
			pStorage->OpenStorage(statstg.pwcsName, NULL, STGM_READ, 0, 0, &pStorageNew);
			hr = teExtract(pStorageNew, bsPath, ppd, pnItems, nCount, nBase);
			pStorageNew->Release();
		} else if (statstg.type == STGTY_STREAM) {
			ppd->SetLine(2, &bsPath[nBase], TRUE, NULL);
			++pnItems[0];
			if (ppd->HasUserCancelled()) {
				hr = E_ABORT;
			} else if (nCount >= 0) {
				teSetProgress(ppd, pnItems[0], nCount, 2);
				hr = pStorage->OpenStream(statstg.pwcsName, NULL, STGM_READ, NULL, &pStream);
				if SUCCEEDED(hr) {
					HANDLE hFile = CreateFile(bsPath, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
					if (hFile != INVALID_HANDLE_VALUE) {
						while (SUCCEEDED(pStream->Read(pszData, SIZE_BUFF, &uRead)) && uRead) {
							WriteFile(hFile, pszData, uRead, &dwWriteByte, NULL);
						}
#ifdef _2000XP
						if (statstg.mtime.dwLowDateTime == 0) {
							if SUCCEEDED(pStorage->QueryInterface(IID_PPV_ARGS(&pSF2))) {
								if SUCCEEDED(pSF2->ParseDisplayName(NULL, NULL, statstg.pwcsName, &chEaten, &pidl, &dwAttributes)) {
									teGetFileTimeFromItem(pSF2, pidl, &PKEY_DateCreated, &statstg.ctime);
									teGetFileTimeFromItem(pSF2, pidl, &PKEY_DateAccessed, &statstg.atime);
									teGetFileTimeFromItem(pSF2, pidl, &PKEY_DateModified, &statstg.mtime);
									teCoTaskMemFree(pidl);
								}
								pSF2->Release();
							}
						}
#endif
						SetFileTime(hFile, &statstg.ctime, &statstg.atime, &statstg.mtime);
						CloseHandle(hFile);
						hr = S_OK;
					} else {
						hr = E_ACCESSDENIED;
					}
					pStream->Release();
				}
			}
		}
		::SysFreeString(bsPath);
		teCoTaskMemFree(statstg.pwcsName);
	}
	SafeRelease(&pEnumSTATSTG);
	return hr;
}

BOOL teLocalizePath(LPWSTR pszPath, BSTR *pbsPath)
{
	*pbsPath = NULL;
	if (PathMatchSpec(pszPath, L"::{*}*")) {
		BSTR bsPath = ::SysAllocString(pszPath);
		LPWSTR lpEnd = &bsPath[::SysStringLen(bsPath) - 1];
		while (lpEnd > bsPath && !*pbsPath) {
			LPWSTR lp = StrRChr(bsPath, lpEnd, '\\');
			lpEnd = lp;
			if (lp) {
				lp[0] = NULL;
				LPITEMIDLIST pidl = teILCreateFromPathEx(bsPath);
				if (pidl) {
					BSTR bs;
					if SUCCEEDED(teGetDisplayNameFromIDList(&bs, pidl, SHGDN_FORADDRESSBAR)) {
						if (StrChr(bsPath, '\\') && !StrChr(bs, '\\')) {
							::SysFreeString(bs);
							teGetDisplayNameFromIDList(&bs, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
						}
						lp[0] = '\\';
						*pbsPath = teSysAllocStringLenEx(bs, SysStringLen(bs) + lstrlen(lp) + 1);
						lstrcat(*pbsPath, lp);
						::SysFreeString(bs);
					}
					teCoTaskMemFree(pidl);
				}
				lp[0] = '\\';
			}
		}
		::SysFreeString(bsPath);
	}
	return *pbsPath != NULL;
}

HMODULE teLoadLibrary(LPWSTR lpszName)
{
	BSTR bsPath = NULL;
	BSTR bsSystem;
	teGetDisplayNameFromIDList(&bsSystem, g_pidls[CSIDL_SYSTEM], SHGDN_FORPARSING);
	tePathAppend(&bsPath, bsSystem, lpszName);
	HMODULE hModule = GetModuleHandle(bsPath);
	if (!hModule) {
		hModule = LoadLibrary(bsPath);
		if (hModule) {
			g_phModule.push_back(hModule);
		}
	}
	teSysFreeString(&bsPath);
	teSysFreeString(&bsSystem);
	return hModule;
}

VOID teGetDisplayNameOf(VARIANT *pv, int uFlags, VARIANT *pVarResult)
{
	try {
		IUnknown *punk;
		if (pv->vt == VT_BSTR) {
			if (!(uFlags & SHGDN_INFOLDER)) {
				teVariantCopy(pVarResult, pv);
				return;
			}
		} else if (FindUnknown(pv, &punk)) {
			BSTR bs;
			if (teGetStrFromFolderItem(&bs, punk)) {
				if (uFlags & SHGDN_INFOLDER) {
					BSTR bsName = NULL;
					if SUCCEEDED(tePathGetFileName(&bsName, bs)) {
						teSetBSTR(pVarResult, &bsName, -1);
						return;
					}
				}
				if (uFlags & SHGDN_FORADDRESSBAR) {
					BSTR bsLocal;
					if (teLocalizePath(bs, &bsLocal)) {
						teSetSZ(pVarResult, bsLocal);
						::SysFreeString(bsLocal);
						return;
					}
				}
				if (!(uFlags & SHGDN_INFOLDER)) {
					teSetSZ(pVarResult, bs);
					return;
				}
			}
		}
		LPITEMIDLIST pidl;
		if (teGetIDListFromVariant(&pidl, pv)) {
			GetVarPathFromIDList(pVarResult, pidl, uFlags);
			teCoTaskMemFree(pidl);
		}
	} catch(...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teGetDisplayNameOf";
#endif
	}
	if (pVarResult && pVarResult->vt == VT_EMPTY) {
		teSetSZ(pVarResult, L"");
	}
}

BOOL GetVarPathFromFolderItem(FolderItem *pFolderItem, VARIANT *pVarResult)
{
	FolderItem *pid = NULL;
	VARIANT v;
	if (pFolderItem) {
		VariantInit(&v);
		if SUCCEEDED(pFolderItem->QueryInterface(IID_PPV_ARGS(&v.pdispVal))) {
			v.vt = VT_DISPATCH;
			VariantInit(pVarResult);
			teGetDisplayNameOf(&v, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX, pVarResult);
			VariantClear(&v);
			if (pVarResult->vt == VT_BSTR && teIsSearchFolder(pVarResult->bstrVal)) {
				LPITEMIDLIST pidl;
				if (teGetIDListFromObject(pFolderItem, &pidl)) {
					VariantClear(pVarResult);
					GetVarArrayFromIDList(pVarResult, pidl);
					teCoTaskMemFree(pidl);
				}
			}
			return TRUE;
		}
	}
	return FALSE;
}

VOID teGetRootFromDataObj(BSTR *pbs, IDataObject *pDataObj)
{
	FolderItems *pFolderItems;
	if SUCCEEDED(pDataObj->QueryInterface(IID_PPV_ARGS(&pFolderItems))) {
		VARIANT v;
		teSetLong(&v, -1);
		FolderItem *pFolderItem;
		if SUCCEEDED(pFolderItems->Item(v, &pFolderItem)) {
			pFolderItem->get_Path(pbs);
			pFolderItem->Release();
		}
		pFolderItems->Release();
	}
}

//WIC GetImage
HRESULT WINAPI teGetImage(IStream *pStream, LPWSTR lpfn, int cx, HBITMAP *phBM, int *pnAlpha)
{
	HRESULT hr = E_FAIL;
	CteWICBitmap *pWB = new CteWICBitmap();
	pStream->AddRef();
	pWB->FromStreamRelease(pStream, lpfn, FALSE, cx);
	if (pWB->HasImage()) {
		*phBM = pWB->GetHBITMAP(-1);
		*pnAlpha = 0;
		hr = S_OK;
	}
	pWB->Release();
	return hr;
}

//Dispatch API

VOID teApiMemory(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	CteMemory *pMem;
	char *pc = NULL;;
	LPWSTR sz = NULL;
	int nCount = 1;
	IUnknown *punk;
	IStream *pStream = NULL;
	int i = 0;
	if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
		sz = pDispParams->rgvarg[nArg].bstrVal;
		i = GetSizeOfStruct(sz);
	} else if (pDispParams->rgvarg[nArg].vt & VT_ARRAY) {
		pc = pDispParams->rgvarg[nArg].pcVal;
		nCount = pDispParams->rgvarg[nArg].parray->rgsabound[0].cElements;
		i = SizeOfvt(pDispParams->rgvarg[nArg].vt);
		sz = L"SAFEARRAY";
	} else if (nArg >= 1 && FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pStream))) {
			sz = L"IStream";
			i = 1;
			ULARGE_INTEGER uliSize;
			LARGE_INTEGER liOffset;
			liOffset.QuadPart = 0;
			pStream->Seek(liOffset, STREAM_SEEK_END, &uliSize);
			pStream->Seek(liOffset, STREAM_SEEK_SET, NULL);
			nCount = uliSize.QuadPart < param[1].lVal ? uliSize.LowPart : param[1].lVal;
		}
	}
	if (i == 0) {
		i = param[0].lVal;
		if (i == 0) {
			return;
		}
	}
	BSTR bs = NULL;
	if (nArg >= 1 && !pStream) {
		if (i == 2 && pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
			bs = pDispParams->rgvarg[nArg - 1].bstrVal;
			nCount = SysStringByteLen(bs) / 2 + 1;
		} else {
			nCount = param[1].lVal;
		}
		if (nCount < 1) {
			nCount = 1;
		}
		if (nArg >= 2) {
			pc = param[2].pcVal;
		}
	}
	pMem = new CteMemory(i * nCount, pc, nCount, sz);
	if (bs) {
		::CopyMemory(pMem->m_pc, bs, nCount * 2);
	}
	if (pStream) {
		ULONG ulBytesRead;
		pStream->Read(pMem->m_pc, nCount, &ulBytesRead);
		pStream->Release();
	}
	teSetObjectRelease(pVarResult, pMem);
}

VOID teApiContextMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDataObject *pDataObj = NULL;
	CteContextMenu *pdispCM;
	if (nArg >= 6) {
		IShellExtInit *pSEI;
		HMODULE hDll = teCreateInstanceV(&pDispParams->rgvarg[nArg], &pDispParams->rgvarg[nArg - 1], IID_PPV_ARGS(&pSEI));
		if (pSEI) {
			LPITEMIDLIST pidl;
			if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg - 2])) {
				if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg - 3])) {
					HKEY hKey;
					if (RegOpenKeyEx(param[4].hkeyVal, param[5].lpcwstr, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
						if SUCCEEDED(pSEI->Initialize(pidl, pDataObj, hKey)) {
							IUnknown *punk = NULL;
							FindUnknown(&pDispParams->rgvarg[nArg - 6], &punk);
							pdispCM = new CteContextMenu(pSEI, pDataObj, punk);
							pDataObj = NULL;
							pdispCM->m_hDll = hDll;
							hDll = NULL;
							teSetObjectRelease(pVarResult, pdispCM);
						}
						RegCloseKey(hKey);
					}
					SafeRelease(&pDataObj);
				}
				teCoTaskMemFree(pidl);
			}
			pSEI->Release();
		}
		teFreeLibrary(hDll, 0);
	} else if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
		long nCount;
		LPITEMIDLIST *ppidllist = IDListFormDataObj(pDataObj, &nCount);
#ifdef _2000XP
		if (nCount >= 2 && !g_bUpperVista && ppidllist[0] && ILIsEmpty(ppidllist[0])) {
			for (int i = nCount; --i >= 0;) {
				LPITEMIDLIST pidlFull = ILCombine(ppidllist[0], ppidllist[i + 1]);
				BSTR bs;
				if SUCCEEDED(teGetDisplayNameFromIDList(&bs, pidlFull, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
					teILCloneReplace(&ppidllist[i + 1], teILCreateFromPath(bs));
					SysFreeString(bs);
				}
				teCoTaskMemFree(ppidllist[i + 1]);
				teCoTaskMemFree(pidlFull);
			}
			teILFreeClear(&ppidllist[0]);
		}
#endif
		AdjustIDList(ppidllist, nCount);
		if (nCount >= 1) {
			IShellFolder *pSF;
			if (GetShellFolder(&pSF, ppidllist[0])) {
				IContextMenu *pCM;
//				if SUCCEEDED(CDefFolderMenu_Create2(ppidllist[0], g_hwndMain, nCount, const_cast<LPCITEMIDLIST *>(&ppidllist[1]), pSF, NULL, 0, NULL, &pCM)) {
				if SUCCEEDED(pSF->GetUIObjectOf(g_hwndMain, nCount, const_cast<LPCITEMIDLIST *>(&ppidllist[1]), IID_IContextMenu, NULL, (LPVOID*)&pCM)) {
					IUnknown *punk = NULL;
					if (nArg >= 1) {
						FindUnknown(&pDispParams->rgvarg[nArg - 1], &punk);
					}
					pdispCM = new CteContextMenu(pCM, pDataObj, punk);
					pDataObj = NULL;
					teSetObjectRelease(pVarResult, pdispCM);
					pCM->Release();
				}
			}
			do {
				teCoTaskMemFree(ppidllist[nCount]);
			} while (--nCount >= 0);
			delete [] ppidllist;
		}
	}
	if (pDataObj) {
		pDataObj->Release();
	}
}

VOID teApiDropTarget(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	FolderItem *pid;
	GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
	LPITEMIDLIST pidl;
	if (teGetIDListFromObject(pid, &pidl)) {
		IShellFolder *pSF;
		LPCITEMIDLIST pidlPart;
		if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
			IDropTarget *pDT = NULL;
			pSF->GetUIObjectOf(g_hwndMain, 1, &pidlPart, IID_IDropTarget, NULL, (LPVOID*)&pDT);
			teSetObjectRelease(pVarResult, new CteDropTarget(pDT, pid));
			pDT && pDT->Release();
			pSF->Release();
		}
		teCoTaskMemFree(pidl);
	}
	pid->Release();
}

VOID teApiDataObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetObjectRelease(pVarResult, new CteFolderItems(NULL, NULL));
}

VOID teApiOleCmdExec(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		IOleCommandTarget *pCT;
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCT))) {
			GUID *pguid = NULL;
			teCLSIDFromString(param[1].lpcolestr, pguid);
			pCT->Exec(pguid, param[2].olecmdid, param[3].olecmdexecopt, &pDispParams->rgvarg[nArg - 4], pVarResult);
			pCT->Release();
		}
	}
}

VOID teApisizeof(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetSizeOf(&pDispParams->rgvarg[nArg]));
}

VOID teApiLowPart(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, param[0].li.LowPart);
}

VOID teApiULowPart(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetULL(pVarResult, param[0].uli.LowPart);
}

VOID teApiHighPart(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, param[0].li.HighPart);
}

VOID teApiQuadPart(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 1) {
		LARGE_INTEGER li;
		li.LowPart = param[0].dword;
		li.HighPart = param[1].lVal;
		teSetLL(pVarResult, li.QuadPart);
		return;
	}
	teSetLL(pVarResult, param[0].llVal);
}

VOID teApiUQuadPart(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 1) {
		ULARGE_INTEGER li;
		li.LowPart = param[0].dword;
		li.HighPart = param[1].lVal;
		teSetULL(pVarResult, li.QuadPart);
		return;
	}
	teSetULL(pVarResult, teGetU(param[0].llVal));
}

VOID teApipvData(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pDispParams->rgvarg[nArg].vt & VT_ARRAY) {
		teSetPtr(pVarResult, pDispParams->rgvarg[nArg].parray->pvData);
	}
}

VOID teApiExecMethod(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDispatch *pdisp;
	if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
		IDispatch *pdisp1 = NULL;
		if (GetDispatch(&pDispParams->rgvarg[nArg - 2], &pdisp1)) {
			int nCount = teGetObjectLength(pdisp);
			VARIANTARG *pv = GetNewVARIANT(nCount);
			for (int i = nCount; --i >= 0;) {
				teGetPropertyAt(pdisp1, nCount - i - 1, &pv[i]);//Reverse argument
			}
			teExecMethod(pdisp, param[1].lpolestr, pVarResult, -nCount, pv);
			for (int i = nCount; --i >= 0;) {
				tePutPropertyAt(pdisp1, nCount - i - 1, &pv[i]);//for Reference
			}
			teClearVariantArgs(nCount, pv);
			pdisp1->Release();
		}
		pdisp->Release();
	}
}

HRESULT teInitStorage(LPVARIANTARG pvDllFile, LPVARIANTARG pvClass, LPWSTR lpwstr, HMODULE *phDll, IStorage **ppStorage)
{
	HRESULT hr = E_FAIL;
	*phDll = teCreateInstanceV(pvDllFile, pvClass, IID_PPV_ARGS(ppStorage));
	if (*ppStorage) {
		IPersistFile *pPF;
		hr = (*ppStorage)->QueryInterface(IID_PPV_ARGS(&pPF));
		if SUCCEEDED(hr) {
			hr = pPF->Load(lpwstr, STGM_READ);
			pPF->Release();
		} else {
			IPersistFolder *pPF2;
			hr = (*ppStorage)->QueryInterface(IID_PPV_ARGS(&pPF2));
			if SUCCEEDED(hr) {
				LPITEMIDLIST pidl = teILCreateFromPath(lpwstr);
				if (pidl) {
					hr = pPF2->Initialize(pidl);
					teCoTaskMemFree(pidl);
				}
				pPF2->Release();
			}
		}
	}
	return hr;
}

VOID teApiExtract(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HRESULT hr = E_FAIL;
	IStorage *pStorage = NULL;
	HMODULE hDll = NULL;
	IProgressDialog *ppd = NULL;
	teCreateInstance(CLSID_ProgressDialog, NULL, NULL, IID_PPV_ARGS(&ppd));
	try {
		if (ppd) {
			ppd->StartProgressDialog(g_hwndMain, NULL, PROGDLG_NORMAL | PROGDLG_AUTOTIME, NULL);
#ifdef _2000XP
			ppd->SetAnimation(g_hShell32, 161);
#endif
			ppd->SetLine(1, PathFindFileName(param[2].lpwstr), TRUE, NULL);
		}
		hr = teInitStorage(&pDispParams->rgvarg[nArg], &pDispParams->rgvarg[nArg - 1], param[2].lpwstr, &hDll, &pStorage);
		if SUCCEEDED(hr) {
			int nItems = 0, nCount = 0, nBase = lstrlen(param[3].lpwstr) + 1;
			hr = teExtract(pStorage, param[3].lpwstr, ppd, &nCount, -1, nBase);
			if SUCCEEDED(hr) {
				hr = teExtract(pStorage, param[3].lpwstr, ppd, &nItems, nCount, nBase);
			}
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"ApiExtract";
#endif
	}
	SafeRelease(&pStorage);
	teFreeLibrary(hDll, 0);
	if (ppd) {
		ppd->StopProgressDialog();
		SafeRelease(&ppd);
	}
	teSetLong(pVarResult, hr);
}

VOID teApiQuadAdd(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLL(pVarResult, param[0].llVal + param[1].llVal);
}

VOID teApiUQuadAdd(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetULL(pVarResult, teGetU(param[0].llVal) + teGetU(param[1].llVal));
}

VOID teApiUQuadSub(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLL(pVarResult, teGetU(param[0].llVal) - teGetU(param[1].llVal));
}

VOID teApiUQuadCmp(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LONGLONG ll = teGetU(param[0].llVal) - teGetU(param[1].llVal);
	int i = 0;
	if (ll > 0) {
		i = 1;
	} else if (ll < 0) {
		i = -1;
	}
	teSetLong(pVarResult, i);
}

VOID teApiFindWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, FindWindow(param[0].lpcwstr, param[1].lpcwstr));
}

VOID teApiFindWindowEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, FindWindowEx(param[0].hwnd, param[1].hwnd, param[2].lpcwstr, param[3].lpcwstr));
}

VOID teApiOleGetClipboard(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		IDataObject *pDataObj;
		if (::OleGetClipboard(&pDataObj) == S_OK) {
			teSetObjectRelease(pVarResult, new CteFolderItems(pDataObj, NULL));
			pDataObj->Release();
		}
	}
}

VOID teApiOleSetClipboard(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LONG lResult = E_FAIL;
	IDataObject *pDataObj;
	if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
		lResult = ::OleSetClipboard(pDataObj);
		teSysFreeString(&g_bsClipRoot);
		teGetRootFromDataObj(&g_bsClipRoot, pDataObj);
//		Don't release pDataObj.
//		pDataObj->Release();
	}
	teSetLong(pVarResult, lResult);
}

VOID teApiPathMatchSpec(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, tePathMatchSpec(param[0].lpcwstr, param[1].lpwstr));
}

VOID teSzToArgv(IDispatch *pArray, LPTSTR lpsz)
{
	BSTR bs = ::SysAllocString(lpsz);
	LPWSTR lpszQuote = StrChr(bs, '"');
	if (lpszQuote) {
		lpszQuote[0] = '\\';
		int i = 0;
		while (lpszQuote[++i] == ' ') {}
		if (i > 1) {
			lpszQuote[1] = NULL;
			teSzToArgv(pArray, bs);
			if (lpszQuote[i]) {
				teSzToArgv(pArray, &lpszQuote[i]);
			}
			teSysFreeString(&bs);
			return;
		}
	}
	VARIANT v;
	v.vt = VT_BSTR;
	v.bstrVal = bs;
	teExecMethod(pArray, L"push", NULL, -1, &v);
	VariantClear(&v);
}

VOID teApiCommandLineToArgv(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDispatch *pArray = NULL;
	if (!GetDispatch(&pDispParams->rgvarg[nArg], &pArray)) {
		pArray = teCreateObj(TE_ARRAY, NULL);
		int nLen = 0;
		LPTSTR *lplpszArgs = NULL;
		if (param[0].lpcwstr && param[0].lpcwstr[0]) {
			lplpszArgs = CommandLineToArgvW(param[0].lpcwstr, &nLen);
			for (int i = 0; i < nLen; ++i) {
				teSzToArgv(pArray, lplpszArgs[i]);
			}
			LocalFree(lplpszArgs);
		}
	}
	teSetObjectRelease(pVarResult, teAddLegacySupport(pArray));
}

VOID teApiILIsEqual(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	FolderItem *pid1, *pid2;
	GetFolderItemFromVariant(&pid1, &pDispParams->rgvarg[nArg]);
	GetFolderItemFromVariant(&pid2, &pDispParams->rgvarg[nArg - 1]);
	teSetBool(pVarResult, teILIsEqual(pid1, pid2));
	pid2->Release();
	pid1->Release();
}

VOID teApiILClone(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	FolderItem *pid;
	GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
	teSetObjectRelease(pVarResult, pid);
}

VOID teApiILIsParent(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl, pidl2;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		if (teGetIDListFromVariant(&pidl2, &pDispParams->rgvarg[nArg - 1])) {
			teSetBool(pVarResult, ::ILIsParent(pidl, pidl2, param[2].boolVal));
			teCoTaskMemFree(pidl2);
		}
		teCoTaskMemFree(pidl);
	}
}

VOID teApiILRemoveLastID(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		if (!ILIsEmpty(pidl)) {
			if (::ILRemoveLastID(pidl)) {
				teSetIDList(pVarResult, ILIsEqual(pidl, g_pidls[CSIDL_INTERNET]) ? g_pidls[CSIDL_DESKTOP] : pidl);
			}
		}
		teCoTaskMemFree(pidl);
	}
}

VOID teApiILFindLastID(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl, pidlLast;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		pidlLast = ILFindLastID(pidl);
		if (pidlLast) {
			teSetIDList(pVarResult, pidl);
		}
		teCoTaskMemFree(pidl);
	}
}

VOID teApiILIsEmpty(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	FolderItem *pid1;
	GetFolderItemFromVariant(&pid1, &pDispParams->rgvarg[nArg]);
	CteFolderItem *pid2 = new CteFolderItem(NULL);
	pid2->Initialize(g_pidls[CSIDL_DESKTOP]);
	teSetBool(pVarResult, teILIsEqual(pid1, static_cast<FolderItem *>(pid2)));
	pid2->Release();
	pid1->Release();
}

VOID teApiILGetParent(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		FolderItem *pFolderItem, *pFolderItem1;
		GetFolderItemFromVariant(&pFolderItem, &pDispParams->rgvarg[nArg]);
		if (teILGetParent(pFolderItem, &pFolderItem1)) {
			teSetObjectRelease(pVarResult, pFolderItem1);
		}
		pFolderItem->Release();
	}
}

VOID teApiFindFirstFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, FindFirstFile(param[0].lpcwstr, param[1].lpwin32_find_data));
}

VOID teApiWindowFromPoint(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg]);
	teSetPtr(pVarResult, WindowFromPoint(pt));
}

VOID teApiGetThreadCount(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, param[0].dword ? teGetthreadCount(param[0].dword): g_nThreads);
}

VOID teApiDoEvents(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	MSG msg;
	BOOL bResult = PeekMessage(&msg, NULL, 0, 0, PM_REMOVE);
	if (bResult) {
		if (MessageProc(&msg) != S_OK) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	teSetBool(pVarResult, bResult);
}

VOID teApisscanf(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		LPWSTR pszData = param[0].lpwstr;
		LPWSTR pszFormat = param[1].lpwstr;
		if (pszData && pszFormat) {
			int nPos = 0;
			while (pszFormat[nPos] && pszFormat[nPos++] != '%') {
			}
			while (WCHAR wc = pszFormat[nPos]) {
				++nPos;
				if (wc == '%') {
					break;
				}
				if (StrChrIW(L"diuoxc", wc)) {//Integer
					LONGLONG Result = 0;
					swscanf_s(pszData, pszFormat, &Result);
					teSetLL(pVarResult, Result);
					return;
				}
				if (StrChrIW(L"efga", wc)) {//Real
					pVarResult->vt = VT_R8;
					pVarResult->dblVal = 0;
					swscanf_s(pszData, pszFormat, &pVarResult->dblVal);
					return;
				}
				if (StrChrIW(L"s", wc)) {//String
					int nLen = lstrlen(pszData);
					BSTR bs = ::SysAllocStringLen(NULL, nLen);
					swscanf_s(pszData, pszFormat, bs, nLen);
					teSetBSTR(pVarResult, &bs, -1);
				}
			}
		}
	}
}

VOID teApiSetFileTime(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BOOL bResult = FALSE;
	BSTR bs = NULL;
	if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
		bs = ::SysAllocString(pDispParams->rgvarg[nArg].bstrVal);
	} else {
		LPITEMIDLIST pidl;
		if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
			teGetDisplayNameFromIDList(&bs, pidl, SHGDN_FORPARSING);
			teCoTaskMemFree(pidl);
		}
	}
	if (bs) {
		HANDLE hFile = ::CreateFile(bs, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0);
		if (hFile != INVALID_HANDLE_VALUE) {
			FILETIME *ppft[3];
			for (int i = _countof(ppft); i--;) {
				ppft[i] = NULL;
				DATE dt;
				if (teGetVariantTime(&pDispParams->rgvarg[nArg - i - 1], &dt) && dt >= 0) {
					ppft[i] = new FILETIME[1];
					teVariantTimeToFileTime(dt, ppft[i]);
				}
			}
			bResult = ::SetFileTime(hFile, ppft[0], ppft[1], ppft[2]);
			for (int i = _countof(ppft); i--;) {
				if (ppft[i]) {
					delete [] ppft[i];
				}
			}
			::CloseHandle(hFile);
		}
		teSysFreeString(&bs);
	}
	teSetBool(pVarResult, bResult);
}

VOID teApiILCreateFromPath(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[1].boolVal) {
		if (pVarResult) {
			LPITEMIDLIST pidl;
			if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
				GetVarArrayFromIDList(pVarResult, pidl);
				teCoTaskMemFree(pidl);
			}
		}
	} else {
		FolderItem *pid;
		GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
		teSetObjectRelease(pVarResult, pid);
	}
}

VOID teApiGetProcObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HMODULE hDll = LoadLibrary(param[0].lpcwstr);
	if (hDll) {
		CHAR szProcNameA[MAX_LOADSTRING];
		LPSTR lpProcNameA = szProcNameA;
		if (pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
			::WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)pDispParams->rgvarg[nArg - 1].bstrVal, -1, lpProcNameA, MAX_LOADSTRING, NULL, NULL);
		} else {
			lpProcNameA = MAKEINTRESOURCEA(param[1].word);
		}
		LPFNGetProcObjectW _GetProcObjectW = (LPFNGetProcObjectW)GetProcAddress(hDll, lpProcNameA);
		if (_GetProcObjectW) {
			if (nArg >= 2) {
				teVariantCopy(pVarResult, &pDispParams->rgvarg[nArg - 2]);
			}
			_GetProcObjectW(pVarResult);
			IDispatch *pdisp;
			if (GetDispatch(pVarResult, &pdisp)) {
				CteDispatch *odisp = new CteDispatch(pdisp, 4, DISPID_UNKNOWN);
				pdisp->Release();
				odisp->m_hDll = hDll;
				hDll = NULL;
				VariantClear(pVarResult);
				teSetObjectRelease(pVarResult, odisp);
			}
		}
		teFreeLibrary(hDll, 0);
	}
}

VOID teApiSetCurrentDirectory(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetCurrentDirectory(param[0].lpcwstr));//use wsh.CurrentDirectory
}

VOID teApiSetDllDirectory(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
#ifdef _2000XP
	if (_SetDllDirectoryW) {
		teSetBool(pVarResult, _SetDllDirectoryW(param[0].lpcwstr));
	}
#else
	teSetBool(pVarResult, SetDllDirectory(param[0].lpcwstr));
#endif
}

VOID teApiPathIsNetworkPath(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, tePathIsNetworkPath(param[0].lpcwstr));
}

VOID teApiRegisterWindowMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, RegisterWindowMessage(param[0].lpcwstr));
}

VOID teApiStrCmpI(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, lstrcmpi(param[0].lpcwstr, param[1].lpcwstr));
}

VOID teApiStrLen(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, lstrlen(param[0].lpcwstr));
}

VOID teApiStrCmpNI(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, StrCmpNI(param[0].lpcwstr, param[1].lpcwstr, param[2].intVal));
}

VOID teApiStrCmpLogical(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, StrCmpLogicalW(param[0].lpcwstr, param[1].lpcwstr));
}

VOID teApiPathQuoteSpaces(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].bstrVal) {
		int nLen = lstrlen(param[0].bstrVal);
		BSTR bsResult = teSysAllocStringLen(param[0].bstrVal, nLen + 2);
		PathQuoteSpaces(bsResult);
		if (!StrChr(bsResult, '"')) {
			if (StrChr(bsResult, ',') || StrChr(bsResult, '&')) {
				lstrcpyn(&bsResult[1], param[0].bstrVal, nLen + 1);
				bsResult[0] = bsResult[nLen + 1] = '"';
			}
		}
		teSetBSTR(pVarResult, &bsResult, -1);
	}
}

VOID teApiPathUnquoteSpaces(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].lpolestr) {
		BSTR bsResult = ::SysAllocString(param[0].lpolestr);
		PathUnquoteSpaces(bsResult);
		teSetBSTR(pVarResult, &bsResult, -1);
	}
}

VOID teApiGetShortPathName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].lpcwstr) {
		int nLen = GetShortPathName(param[0].lpcwstr, NULL, 0);
		BSTR bsResult = ::SysAllocStringLen(NULL, nLen + MAX_PATH);
		nLen = GetShortPathName(param[0].lpcwstr, bsResult, nLen);
		teSetBSTR(pVarResult, &bsResult, nLen);
	}
}

VOID teApiPathCreateFromUrl(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].bstrVal) {
		if (teIsFileSystem(param[0].bstrVal)) {
			teSetSZ(pVarResult, param[0].bstrVal);
			return;
		}
		DWORD dwLen = ::SysStringLen(param[0].bstrVal) + MAX_PATH;
		BSTR bsResult = ::SysAllocStringLen(NULL, dwLen);
		if SUCCEEDED(PathCreateFromUrl(param[0].lpcwstr, bsResult, &dwLen, NULL)) {
			teSetBSTR(pVarResult, &bsResult, dwLen);
		} else {
			::SysFreeString(bsResult);
			teSetSZ(pVarResult, NULL);
		}
	}
}

VOID teApiUrlCreateFromPath(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].bstrVal) {
		DWORD dwLen = ::SysStringLen(param[0].bstrVal) + MAX_PATH;
		BSTR bsResult = ::SysAllocStringLen(NULL, dwLen);
		if SUCCEEDED(UrlCreateFromPath(param[0].lpcwstr, bsResult, &dwLen, NULL)) {
			teSetBSTR(pVarResult, &bsResult, dwLen);
		} else {
			::SysFreeString(bsResult);
			teSetSZ(pVarResult, NULL);
		}
	}
}

VOID teApiPathSearchAndQualify(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].bstrVal) {
		UINT uLen = ::SysStringLen(param[0].bstrVal) + MAX_PATH;
		BSTR bsResult = teSysAllocStringLen(param[0].lpolestr, uLen);
		PathSearchAndQualify(param[0].lpcwstr, bsResult, uLen);
		teSetBSTR(pVarResult, &bsResult, -1);
	}
}

VOID teApiPSFormatForDisplay(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].lpcwstr) {
		PROPERTYKEY propKey;
		if SUCCEEDED(_PSPropertyKeyFromStringEx(param[0].lpcwstr, &propKey)) {
			LPWSTR psz;
			if SUCCEEDED(tePSFormatForDisplay(&propKey, &pDispParams->rgvarg[nArg - 1], param[2].dword, &psz, g_param[TE_SizeFormat])) {
				teSetSZ(pVarResult, psz);
				teCoTaskMemFree(psz);
			}
		}
	}
}

VOID teApiPSGetDisplayName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].lpcwstr && pVarResult) {
		PROPERTYKEY propKey;
		if (SUCCEEDED(teGetPropertyKeyFromName(NULL, param[0].bstrVal, &propKey))) {
			int nFormat = param[1].intVal;
			CteShellBrowser *pSB = NULL;
			if (nFormat == 0 && g_pTC) {
				pSB = g_pTC->GetShellBrowser(g_pTC->m_nIndex);
			}
			pVarResult->bstrVal = tePSGetNameFromPropertyKeyEx(propKey, nFormat, pSB);
			pVarResult->vt = VT_BSTR;
		}
	}
}

VOID teApiGetCursorPos(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	teSetBool(pVarResult, GetCursorPos(&pt));
	PutPointToVariant(&pt, &pDispParams->rgvarg[nArg]);
}

VOID teApiGetKeyboardState(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetKeyboardState(param[0].pbVal));
}

VOID teApiSetKeyboardState(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetKeyboardState(param[0].pbVal));
}

VOID teApiGetVersionEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].prtl_osversioninfoexw) {
		teSetBool(pVarResult, _RtlGetVersion ? !_RtlGetVersion(param[0].prtl_osversioninfoexw) : GetVersionEx(param[0].lposversioninfo));
	}
}

VOID teApiChooseFont(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ChooseFont(param[0].lpchoosefont));
}

VOID teApiChooseColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ::ChooseColor(param[0].lpchoosecolor));
}

VOID teApiTranslateMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ::TranslateMessage(param[0].lpmsg));
}

VOID teApiShellExecuteEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ::ShellExecuteEx(param[0].lpshellexecuteinfo));
}

VOID teApiCreateFontIndirect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::CreateFontIndirect(param[0].plogfont));
}

VOID teApiCreateIconIndirect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::CreateIconIndirect(param[0].piconinfo));
}

VOID teApiFindText(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::FindText(param[0].lpfindreplace));
}

VOID teApiFindReplace(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::ReplaceText(param[0].lpfindreplace));
}

VOID teApiDispatchMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::DispatchMessage(param[0].lpmsg));
}

VOID teApiPostMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	PostMessage(param[0].hwnd, param[1].uintVal, param[2].wparam, param[3].lparam);
}

VOID teApiSHFreeNameMappings(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SHFreeNameMappings(param[0].handle);
}

VOID teApiCoTaskMemFree(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teCoTaskMemFree(param[0].lpvoid);
}

VOID teApiSleep(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	Sleep(param[0].dword);
#ifdef _DEBUG
	teSetPtr(pVarResult, GetCurrentThreadId());
#endif
}

VOID teApiShRunDialog(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (_SHRunDialog) {
		_SHRunDialog(param[0].hwnd, param[1].hicon, param[2].lpwstr, param[3].lpwstr, param[4].lpwstr, param[5].dword);
	}
}

VOID teApiDragAcceptFiles(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DragAcceptFiles(param[0].hwnd, param[1].boolVal);
}

VOID teApiDragFinish(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DragFinish(param[0].hdrop);
}

VOID teApimouse_event(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	mouse_event(param[0].dword, param[1].dword, param[2].dword, param[3].dword, param[4].ulong_ptr);
}

VOID teApikeybd_event(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	keybd_event(param[0].bVal, param[1].bVal, param[2].dword, param[3].ulong_ptr);
}

VOID teApiSHChangeNotify(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	VARIANT pv[2];
	BYTE type = param[1].bVal;
	for (int i = 2; i--;) {
		VariantInit(&pv[i]);
		pv[i].bstrVal = NULL;
		if (type == SHCNF_IDLIST) {
			teGetIDListFromVariant((LPITEMIDLIST *)&pv[i].bstrVal, &pDispParams->rgvarg[nArg - i - 2]);
		} else if (type == SHCNF_PATH) {
			teVariantChangeType(&pv[i], &pDispParams->rgvarg[nArg - i - 2], VT_BSTR);
		}
	}
	SHChangeNotify(param[0].lVal, param[1].uintVal, pv[0].bstrVal, pv[1].bstrVal);
	for (int i = 2; i--;) {
		if (pv[i].vt != VT_BSTR) {
			teCoTaskMemFree(pv[i].bstrVal);
		} else {
			VariantClear(&pv[i]);
		}
	}
}

VOID teApiDrawIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DrawIcon(param[0].hdc, param[1].intVal, param[2].intVal, param[3].hicon));
}

VOID teApiDestroyMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyMenu(param[0].hmenu));
}

VOID teApiFindClose(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, FindClose(param[0].handle));
}

VOID teApiFreeLibrary(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teFreeLibrary(param[0].hmodule, param[1].uintVal));
}

VOID teApiImageList_Destroy(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 0 && param[1].boolVal) {
		try {
			if (param[0].pImageList) {
				teSetLong(pVarResult, param[0].pImageList->Release());
			}
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"ApiImageList_Destroy";
#endif
		}
		return;
	}
	teSetBool(pVarResult, ImageList_Destroy(param[0].himagelist));
}

VOID teApiDeleteObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteObject(param[0].hgdiobj));
}

VOID teApiDestroyIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyIcon(param[0].hicon));
}

VOID teApiIsWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsWindow(param[0].hwnd));
}

VOID teApiIsWindowVisible(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsWindowVisible(param[0].hwnd));
}

VOID teApiIsZoomed(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsZoomed(param[0].hwnd));
}

VOID teApiIsIconic(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsIconic(param[0].hwnd));
}

VOID teApiOpenIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, OpenIcon(param[0].hwnd));
}

VOID teApiSetForegroundWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teSetForegroundWindow(param[0].hwnd));
}

VOID teApiBringWindowToTop(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, BringWindowToTop(param[0].hwnd));
}

VOID teApiDeleteDC(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteDC(param[0].hdc));
}

VOID teApiCloseHandle(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, CloseHandle(param[0].handle));
}

VOID teApiIsMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsMenu(param[0].hmenu));
}

VOID teApiMoveWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, MoveWindow(param[0].hwnd, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal, param[5].boolVal));
#ifdef _2000XP
	CteTreeView *pTV = TVfromhwnd2(param[0].hwnd);
	if (pTV) {
		pTV->SetObjectRect();
	}
#endif
}

VOID teApiSetMenuItemBitmaps(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuItemBitmaps(param[0].hmenu, param[1].uintVal, param[2].uintVal, param[3].hbitmap, param[4].hbitmap));
}

VOID teApiShowWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ShowWindow(param[0].hwnd, param[1].intVal));
}

VOID teApiDeleteMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteMenu(param[0].hmenu, param[1].uintVal, param[2].uintVal));
}

VOID teApiRemoveMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, RemoveMenu(param[0].hmenu, param[1].uintVal, param[2].uintVal));
}

VOID teApiDrawIconEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DrawIconEx(param[0].hdc, param[1].intVal, param[2].intVal, param[3].hicon,
	param[4].intVal, param[5].intVal, param[6].uintVal, param[7].hbrush, param[8].uintVal));
}

VOID teApiEnableMenuItem(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, EnableMenuItem(param[0].hmenu, param[1].uintVal, param[2].uintVal));
}

VOID teApiImageList_Remove(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_Remove(param[0].himagelist, param[1].intVal));
}

VOID teApiImageList_SetIconSize(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_SetIconSize(param[0].himagelist, param[1].intVal, param[2].intVal));
}

VOID teApiImageList_Draw(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_Draw(param[0].himagelist, param[1].intVal, param[2].hdc,
	param[3].intVal, param[4].intVal, param[5].uintVal));
}

VOID teApiImageList_DrawEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_DrawEx(param[0].himagelist, param[1].intVal, param[2].hdc,
	param[3].intVal, param[4].intVal, param[5].intVal, param[6].intVal,
	param[7].colorref, param[8].colorref, param[9].uintVal));
}

VOID teApiImageList_SetImageCount(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_SetImageCount(param[0].himagelist, param[1].uintVal));
}

VOID teApiImageList_SetOverlayImage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_SetOverlayImage(param[0].himagelist, param[1].intVal, param[2].intVal));
}

VOID teApiImageList_Copy(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_Copy(param[0].himagelist, param[1].intVal,
	param[2].himagelist, param[3].intVal, param[4].uintVal));
}

VOID teApiDestroyWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyWindow(param[0].hwnd));
}

VOID teApiLineTo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, LineTo(param[0].hdc, param[1].intVal, param[2].intVal));
}

VOID teApiReleaseCapture(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ReleaseCapture());
}

VOID teApiSetCursorPos(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetCursorPos(param[0].intVal, param[1].intVal));
}

VOID teApiDestroyCursor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyCursor(param[0].hcursor));
}

VOID teApiSHFreeShared(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SHFreeShared(param[0].handle, param[1].dword));
}

VOID teApiEndMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, EndMenu());
}

VOID teApiSHChangeNotifyDeregister(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SHChangeNotifyDeregister(param[0].ulVal));
}

VOID teApiSHChangeNotification_Unlock(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, TRUE);
}

VOID teApiIsWow64Process(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BOOL bResult = FALSE;
#ifdef _2000XP
	if (_IsWow64Process) {
		_IsWow64Process(param[0].handle, &bResult);
	}
#else
	IsWow64Process(param[0].handle, &bResult);
#endif
	teSetBool(pVarResult, bResult);
}

VOID teApiBitBlt(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, BitBlt(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal,
	param[5].hdc, param[6].intVal, param[7].intVal, param[8].dword));
}

VOID teApiImmSetOpenStatus(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImmSetOpenStatus(param[0].himc, param[1].boolVal));
}

VOID teApiImmReleaseContext(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImmReleaseContext(param[0].hwnd, param[1].himc));
}

VOID teApiIsChild(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsChild(param[0].hwnd, param[1].hwnd));
}

VOID teApiKillTimer(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, KillTimer(param[0].hwnd, param[1].uint_ptr));
}

VOID teApiAllowSetForegroundWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, AllowSetForegroundWindow(param[0].dword));
}

VOID teApiSetWindowPos(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetWindowPos(param[0].hwnd, param[1].hwnd, param[2].intVal, param[3].intVal, param[4].intVal, param[5].intVal, param[6].uintVal));
}

VOID teApiInsertMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, InsertMenu(param[0].hmenu, param[1].uintVal, param[2].uintVal,
	param[3].uint_ptr, param[4].lpcwstr));
}

VOID teApiSetWindowText(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teSetWindowText(param[0].hwnd, param[1].lpcwstr));
}

VOID teApiRedrawWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, RedrawWindow(param[0].hwnd, param[1].lprect, param[2].hrgn, param[3].uintVal));
}

VOID teApiMoveToEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, MoveToEx(param[0].hdc, param[1].intVal, param[2].intVal, param[3].lppoint));
}

VOID teApiInvalidateRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, InvalidateRect(param[0].hwnd, param[1].lprect, param[2].boolVal));
}

VOID teApiSendNotifyMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SendNotifyMessage(param[0].hwnd, param[1].uintVal, param[2].wparam, param[3].lparam));
}

VOID teApiPeekMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, PeekMessage(param[0].lpmsg, param[1].hwnd, param[2].uintVal, param[3].uintVal, param[4].uintVal));
}

VOID teApiSHFileOperation(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 4) {
		LPSHFILEOPSTRUCT pFO = new SHFILEOPSTRUCT[1];
		::ZeroMemory(pFO, sizeof(SHFILEOPSTRUCT));
		pFO->hwnd = g_hwndMain;
		pFO->wFunc = param[0].uintVal;
		VARIANT *pv = &pDispParams->rgvarg[nArg - 1];
		BSTR bs = NULL;
		int nLen = 0;
		IDispatch *pdisp;
		if (GetDispatch(pv, &pdisp)) {
			VARIANT v;
			VariantInit(&v);
			int nCount = teGetObjectLength(pdisp);
			for (int i = 0; i < nCount; ++i) {
				if SUCCEEDED(teGetPropertyAt(pdisp, i, &v)) {
					if (v.vt == VT_BSTR) {
						nLen += ::SysStringByteLen(v.bstrVal);
					}
					VariantInit(&v);
				}
			}
			nLen += nCount * sizeof(WORD);
			bs = ::SysAllocStringByteLen(NULL, nLen);
			::ZeroMemory(bs, nLen);
			BYTE *p = (BYTE *)bs;
			for (int i = 0; i < nCount; ++i) {
				if SUCCEEDED(teGetPropertyAt(pdisp, i, &v)) {
					if (v.vt == VT_BSTR) {
						UINT nSize = ::SysStringByteLen(v.bstrVal);
						CopyMemory(p, v.bstrVal, nSize);
						p += nSize;
					}
					VariantInit(&v);
				}
				p += sizeof(WORD);
			}
			pdisp->Release();
		} else {
			bs = GetLPWSTRFromVariant(pv);
			nLen = (pv->vt == VT_BSTR ? ::SysStringLen(bs) : lstrlen(bs)) + 1;
			bs = teSysAllocStringLenEx(bs, nLen);
			bs[nLen] = 0;
		}
		pFO->pFrom = bs;
		pv = &pDispParams->rgvarg[nArg - 2];
		bs = GetLPWSTRFromVariant(pv);
		nLen = (pv->vt == VT_BSTR ? ::SysStringLen(bs): lstrlen(bs)) + 1;
		bs = pv->vt == VT_BSTR ? teSysAllocStringLenEx(bs, nLen) : teSysAllocStringLen(bs, nLen);
		bs[nLen] = 0;
		pFO->pTo = bs;
		pFO->fFlags = param[3].fileop_flags;
		if (param[4].boolVal) {
			teSetPtr(pVarResult, _beginthread(&threadFileOperation, 0, pFO));
			return;
		}
		try {
			teSetLong(pVarResult, teSHFileOperation(pFO));
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"ApiSHFileOperation";
#endif
		}
		::SysFreeString(const_cast<BSTR>(pFO->pTo));
		::SysFreeString(const_cast<BSTR>(pFO->pFrom));
		delete [] pFO;
		return;
	}
	teSetLong(pVarResult, teSHFileOperation(param[0].lpshfileopstruct));
}

VOID teApiGetMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMessage(param[0].lpmsg, param[1].hwnd, param[2].uintVal, param[3].uintVal));
}

VOID teApiGetWindowRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetWindowRect(param[0].hwnd, param[1].lprect));
}

VOID teApiGetWindowThreadProcessId(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetWindowThreadProcessId(param[0].hwnd, param[1].lpdword));
}

VOID teApiGetClientRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetClientRect(param[0].hwnd, param[1].lprect));
}

VOID teApiSendInput(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SendInput(param[0].uintVal, param[1].lpinput, param[2].intVal));
}

VOID teApiScreenToClient(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg - 1]);
	teSetBool(pVarResult, ScreenToClient(param[0].hwnd, &pt));
	PutPointToVariant(&pt, &pDispParams->rgvarg[nArg - 1]);
}

VOID teApiMsgWaitForMultipleObjectsEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, MsgWaitForMultipleObjectsEx(param[0].dword, param[1].lphandle, param[2].dword, param[3].dword, param[4].dword));
}

VOID teApiClientToScreen(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg - 1]);
	teSetBool(pVarResult, ClientToScreen(param[0].hwnd, &pt));
	PutPointToVariant(&pt, &pDispParams->rgvarg[nArg - 1]);
}

VOID teApiGetIconInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetIconInfo(param[0].hicon, param[1].piconinfo));
}

VOID teApiFindNextFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, FindNextFile(param[0].handle, param[1].lpwin32_find_data));
}

VOID teApiFillRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[2].hbrush) {
		teSetBool(pVarResult, FillRect(param[0].hdc, param[1].lprect, param[2].hbrush));
		return;
	}
	BITMAPINFO bmi;
	RGBQUAD *pcl = NULL;
	::ZeroMemory(&bmi.bmiHeader, sizeof(BITMAPINFOHEADER));
	bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bmi.bmiHeader.biWidth = 1;
	bmi.bmiHeader.biHeight = -1;
	bmi.bmiHeader.biPlanes = 1;
	bmi.bmiHeader.biBitCount = 32;
	HBITMAP hBM = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, (void **)&pcl, NULL, 0);
	HDC hmdc = CreateCompatibleDC(param[0].hdc);
	HGDIOBJ hOld = SelectObject(hmdc, hBM);
	pcl->rgbReserved = 0xff;
	pcl->rgbRed = param[3].rgbquad.rgbRed;
	pcl->rgbGreen = param[3].rgbquad.rgbGreen;
	pcl->rgbBlue = param[3].rgbquad.rgbBlue;
	BLENDFUNCTION blendFunction;
	blendFunction.BlendOp = AC_SRC_OVER;
	blendFunction.BlendFlags = 0;
	blendFunction.SourceConstantAlpha = param[3].rgbquad.rgbReserved;
	blendFunction.AlphaFormat = 0;
	AlphaBlend(param[0].hdc, param[1].lprect->left, param[1].lprect->top, param[1].lprect->right - param[1].lprect->left, param[1].lprect->bottom - param[1].lprect->top,
		hmdc, 0, 0, 1, 1, blendFunction);
	SelectObject(hmdc, hOld);
	DeleteDC(hmdc);
	DeleteObject(hBM);
}

VOID teApiShell_NotifyIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, Shell_NotifyIcon(param[0].dword, param[1].pnotifyicondata));
}

VOID teApiEndPaint(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, EndPaint(param[0].hwnd, param[1].lppaintstruct));
}

VOID teApiImageList_GetIconSize(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_GetIconSize(param[0].himagelist, (int *)&(param[1].lpsize->cx), (int *)&(param[1].lpsize->cy)));
}

VOID teApiGetMenuInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMenuInfo(param[0].hmenu, param[1].lpmenuinfo));
}

VOID teApiSetMenuInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuInfo(param[0].hmenu, param[1].lpcmenuinfo));
}

VOID teApiSystemParametersInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SystemParametersInfo(param[0].uintVal, param[1].uintVal, param[2].lpvoid, param[3].uintVal));
}

VOID teApiGetTextExtentPoint32(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetTextExtentPoint32(param[0].hdc, param[1].lpcwstr, lstrlen(param[1].lpcwstr), param[2].lpsize));
}

VOID teApiSHGetDataFromIDList(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl;
	teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg]);
	if (pidl) {
		HRESULT hr = E_FAIL;
		IShellFolder *pSF;
		LPCITEMIDLIST pidlPart;
		if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
			hr = teSHGetDataFromIDList(pSF, pidlPart, param[1].intVal, param[2].lpvoid, param[3].intVal);
			pSF->Release();
		}
		teSetLong(pVarResult, hr);
		teCoTaskMemFree(pidl);
	}
}

VOID teApiInsertMenuItem(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, InsertMenuItem(param[0].hmenu, param[1].uintVal, param[2].boolVal, param[3].lpcmenuiteminfo));
}

VOID teApiGetMenuItemInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMenuItemInfo(param[0].hmenu, param[1].uintVal, param[2].boolVal, param[3].lpmenuiteminfo));
}

VOID teApiSetMenuItemInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuItemInfo(param[0].hmenu, param[1].uintVal, param[2].boolVal, param[3].lpmenuiteminfo));
}

VOID teApiChangeWindowMessageFilterEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teChangeWindowMessageFilterEx(param[0].hwnd, param[1].uintVal, param[2].dword, param[3].pchangefilterstruct));
}

VOID teApiImageList_SetBkColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_SetBkColor(param[0].himagelist, param[1].colorref));
}

VOID teApiImageList_AddIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_AddIcon(param[0].himagelist, param[1].hicon));
}

VOID teApiImageList_Add(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_Add(param[0].himagelist, param[1].hbitmap, param[2].hbitmap));
}

VOID teApiImageList_AddMasked(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_AddMasked(param[0].himagelist, param[1].hbitmap, param[2].colorref));
}

VOID teApiGetKeyState(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetKeyState(param[0].intVal));
}

VOID teApiGetSystemMetrics(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetSystemMetrics(param[0].intVal));
}

VOID teApiGetSysColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
/*	HTHEME hTheme = OpenThemeData(g_hwndMain, L"menu");// Doesn't work on dark mode.
	if (hTheme) {
		teSetLong(pVarResult, GetThemeSysColor(hTheme, param[0].intVal + TMT_SCROLLBAR));
		CloseThemeData(hTheme);
		return;
	}*/
	teSetLong(pVarResult, GetSysColor(param[0].intVal));
}

VOID teApiGetMenuItemCount(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMenuItemCount(param[0].hmenu));
}

VOID teApiImageList_GetImageCount(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_GetImageCount(param[0].himagelist));
}

VOID teApiReleaseDC(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ReleaseDC(param[0].hwnd, param[1].hdc));
}

VOID teApiGetMenuItemID(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMenuItemID(param[0].hmenu, param[1].intVal));
}

VOID teApiImageList_Replace(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_Replace(param[0].himagelist, param[1].intVal,
		param[2].hbitmap, param[3].hbitmap));
}

VOID teApiImageList_ReplaceIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_ReplaceIcon(param[0].himagelist, param[1].intVal, param[2].hicon));
}

VOID teApiSetBkMode(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetBkMode(param[0].hdc, param[1].intVal));
}

VOID teApiSetBkColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetBkColor(param[0].hdc, param[1].colorref));
}

VOID teApiSetTextColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetTextColor(param[0].hdc, param[1].colorref));
}

VOID teApiMapVirtualKey(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, MapVirtualKey(param[0].uintVal, param[1].uintVal));
}

VOID teApiWaitForInputIdle(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, WaitForInputIdle(param[0].handle, param[1].dword));
}

VOID teApiWaitForSingleObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, WaitForSingleObject(param[0].handle, param[1].dword));
}

VOID teApiGetMenuDefaultItem(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMenuDefaultItem(param[0].hmenu, param[1].uintVal, param[2].uintVal));
}

VOID teApiSetMenuDefaultItem(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuDefaultItem(param[0].hmenu, param[1].uintVal, param[2].uintVal));
}

VOID teApiCRC32(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BYTE *pc;
	VARIANT vMem;
	VariantInit(&vMem);
	UINT nLen = GetpDataFromVariant(&pc, &pDispParams->rgvarg[nArg], &vMem);
	teSetLong(pVarResult, CalcCrc32(pc, nLen, param[1].uintVal));
	VariantClear(&vMem);
}

VOID teApiSHEmptyRecycleBin(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SHEmptyRecycleBin(param[0].hwnd, param[1].lpcwstr, param[2].dword));
}

VOID teApiGetMessagePos(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMessagePos());
}

VOID teApiImageList_GetOverlayImage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	int iResult = -1;
	try {
		param[0].pImageList->GetOverlayImage(param[1].intVal, &iResult);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"ApiImageList_GetOverlayImage";
#endif
	}
	teSetLong(pVarResult, iResult);
}

VOID teApiImageList_GetBkColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_GetBkColor(param[0].himagelist));
}

VOID teApiSetWindowTheme(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	CteShellBrowser *pSB = SBfromhwnd(param[0].hwnd);
	if (pSB) {
		if (param[1].lpcwstr == NULL || tePathMatchSpec(param[1].lpcwstr, L"darkmode_explorer")) {
			g_nWindowTheme = 1;
		} else if (tePathMatchSpec(param[1].lpcwstr, L"*itemsview")) {
			g_nWindowTheme = 2;
		} else {
			g_nWindowTheme = 0;
		}
		teSetLong(pVarResult, pSB->SetTheme());
	} else {
		teSetLong(pVarResult, SetWindowTheme(param[0].hwnd, param[1].lpcwstr, param[2].lpcwstr));
	}
}

VOID teApiImmGetVirtualKey(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImmGetVirtualKey(param[0].hwnd));
}

VOID teApiGetAsyncKeyState(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetAsyncKeyState(param[0].intVal));
}

VOID teApiTrackPopupMenuEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	if (nArg >= 6) {
		if (FindUnknown(&pDispParams->rgvarg[nArg - 6], &punk)) {
			IContextMenu *pCM;
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCM))) {
				IDispatch *pdisp = NULL;
				GetNewArray(&pdisp);
				teArrayPush(pdisp, punk);
				pCM->Release();
				pdisp->QueryInterface(IID_PPV_ARGS(&g_pCM));
				pdisp->Release();
			} else {
				punk->QueryInterface(IID_PPV_ARGS(&g_pCM));
			}
		}
	}
	g_hMenuKeyHook = SetWindowsHookEx(WH_KEYBOARD, (HOOKPROC)MenuKeyProc, hInst, g_dwMainThreadId);
#ifdef _DEBUG
	g_hMenuGMHook = SetWindowsHookEx(WH_GETMESSAGE, (HOOKPROC)MenuGMProc, NULL, g_dwMainThreadId);
#endif
	teSetLong(pVarResult, TrackPopupMenuEx(param[0].hmenu, param[1].uintVal, param[2].intVal, param[3].intVal,
		param[4].hwnd, param[5].lptpmparams));
#ifdef _DEBUG
	UnhookWindowsHookEx(g_hMenuGMHook);
#endif
	UnhookWindowsHookEx(g_hMenuKeyHook);
	SafeRelease(&g_pCM);
}

VOID teApiExtractIconEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ExtractIconEx(param[0].lpcwstr, param[1].intVal, param[2].phicon, param[3].phicon, param[4].uintVal));
}

VOID teApiGetObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetObject(param[0].handle, param[1].intVal, param[2].lpvoid));
}

VOID teApiMultiByteToWideChar(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		UCHAR *pc;
		VARIANT vMem;
		VariantInit(&vMem);
		UINT nLen = GetpDataFromVariant(&pc, &pDispParams->rgvarg[nArg - 1], &vMem);
		if (nLen) {
			pVarResult->bstrVal = teMultiByteToWideChar(param[0].uintVal, (LPSTR)pc, param[2].intVal ? param[2].intVal : nLen);
			pVarResult->vt = VT_BSTR;
		}
		VariantClear(&vMem);
	}
}

VOID teApiWideCharToMultiByte(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		UCHAR *pc = NULL;
		VARIANT vMem;
		VariantInit(&vMem);
		GetpDataFromVariant(&pc, &pDispParams->rgvarg[nArg - 1], &vMem);
		if (pc) {
			int nLenW = lstrlen((LPCWSTR)pc);
			int nLenA = ::WideCharToMultiByte(param[0].uintVal, 0, (LPCWSTR)pc, nLenW, NULL, 0, NULL, NULL);
			if (nLenA) {
				SAFEARRAY *psa = ::SafeArrayCreateVector(VT_UI1, 0, nLenA);
				if (psa) {
					PVOID pvData;
					if (::SafeArrayAccessData(psa, &pvData) == S_OK) {
						pVarResult->vt = VT_ARRAY | VT_UI1;
						pVarResult->parray = psa;
						::WideCharToMultiByte(param[0].uintVal, 0, (LPCWSTR)pc, nLenW, (LPSTR)pvData, nLenA, NULL, NULL);
						::SafeArrayUnaccessData(psa);
					}
				}
			}
		}
		VariantClear(&vMem);
	}
}

VOID teApiGetAttributesOf(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	ULONG lResult = 0;
	IDataObject *pDataObj;
	if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
		long nCount;
		LPITEMIDLIST *ppidllist = IDListFormDataObj(pDataObj, &nCount);
		if (ppidllist) {
			AdjustIDList(ppidllist, nCount);
			if (nCount >= 1) {
				IShellFolder *pSF;
				if (GetShellFolder(&pSF, ppidllist[0])) {
					lResult = param[1].ulVal;
					if SUCCEEDED(pSF->GetAttributesOf(nCount, (LPCITEMIDLIST *)&ppidllist[1], (SFGAOF *)(&lResult))) {
						lResult &= param[1].ulVal;
					} else {
						lResult = 0;
					}
					pSF->Release();
				}
			}
			do {
				teCoTaskMemFree(ppidllist[nCount]);
			} while (--nCount >= 0);
			delete [] ppidllist;
		}
		pDataObj->Release();
	}
	teSetLong(pVarResult, lResult);
}

VOID teApiDoDragDrop(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDataObject *pDataObj;
	if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
		DWORD dwEffect = param[1].dword;
		teSetLong(pVarResult, teDoDragDrop(g_hwndMain, pDataObj, &dwEffect, FALSE));
		IDispatch *pEffect;
		if (nArg >= 2 && GetDispatch(&pDispParams->rgvarg[nArg - 2], &pEffect)) {
			VARIANT v;
			teSetLong(&v, dwEffect);
			tePutPropertyAt(pEffect, 0, &v);
			pEffect->Release();
		}
		pDataObj->Release();
	}
}

VOID teApiSHDoDragDrop(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDataObject *pDataObj;
	if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg - 1])) {
		FindUnknown(&pDispParams->rgvarg[nArg - 2], &g_pDraggingCtrl);
		DWORD dwEffect = param[3].dword;
		teSetLong(pVarResult, teDoDragDrop(param[0].hwnd, pDataObj, &dwEffect, param[5].boolVal));
		g_pDraggingCtrl = NULL;
		IDispatch *pEffect;
		if (nArg >= 4 && GetDispatch(&pDispParams->rgvarg[nArg - 4], &pEffect)) {
			VARIANT v;
			teSetLong(&v, dwEffect);
			tePutPropertyAt(pEffect, 0, &v);
			pEffect->Release();
		}
		pDataObj->Release();
	}
}

VOID teApiCompareIDs(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl1, pidl2;
	LPCITEMIDLIST pidlPart;
	if (teGetIDListFromVariant(&pidl1, &pDispParams->rgvarg[nArg - 1])) {
		if (teGetIDListFromVariant(&pidl2, &pDispParams->rgvarg[nArg - 2])) {
			HRESULT hr = E_FAIL;
			IShellFolder *pSF = NULL;
			LPITEMIDLIST pidlParent = ::ILClone(pidl1);
			::ILRemoveLastID(pidlParent);
			BOOL bInternet1 = ::ILIsParent(g_pidls[CSIDL_INTERNET], pidl1, FALSE);
			BOOL bInternet2 = ::ILIsParent(g_pidls[CSIDL_INTERNET], pidl2, FALSE);
			if (bInternet1 || bInternet2) {
				if (!bInternet1) {
					hr = (WORD)-1;
				} else if (!bInternet2) {
					hr = 1;
				} else {
					BSTR bs1, bs2;
					if SUCCEEDED(teGetDisplayNameFromIDList(&bs1, pidl1, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
						if SUCCEEDED(teGetDisplayNameFromIDList(&bs2, pidl2, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
							hr = (WORD)StrCmpLogicalW(bs1, bs2);
							teSysFreeString(&bs2);
						}
						teSysFreeString(&bs1);
					}
				}
			} else if (ILIsParent(pidlParent, pidl2, TRUE) && SUCCEEDED(SHBindToParent(pidl1, IID_PPV_ARGS(&pSF), &pidlPart))) {
				hr = pSF->CompareIDs(param[0].lparam, pidlPart, ILFindLastID(pidl2));
			} else {
				SHGetDesktopFolder(&pSF);
				hr = pSF->CompareIDs(param[0].lparam, pidl1, pidl2);
			}
			SafeRelease(&pSF);
			teCoTaskMemFree(pidlParent);
			teCoTaskMemFree(pidl2);
			if SUCCEEDED(hr) {
				teSetLong(pVarResult, (short)hr);
			}
		}
		teCoTaskMemFree(pidl1);
	}
}

static void threadExecScript(void *args)
{
	::CoInitialize(NULL);
	VARIANT v;
	VariantInit(&v);
	TEExecScript *pES = (TEExecScript *)args;
	try {
		if (pES->pStream) {
			if SUCCEEDED(CoGetInterfaceAndReleaseStream(pES->pStream, IID_PPV_ARGS(&v.pdispVal))) {
				v.vt = VT_DISPATCH;
			}
		}
		ParseScript(pES->bsScript, pES->bsLang, &v, NULL, NULL);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"threadExecScript";
#endif
	}
	VariantClear(&v);
	teSysFreeString(&pES->bsLang);
	teSysFreeString(&pES->bsScript);
	::CoUninitialize();
	::_endthread();
}

VOID teApiExecScript(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	VARIANT *pv = NULL;
	if (nArg >= 2) {
		pv = &pDispParams->rgvarg[nArg - 2];
		if (param[3].boolVal) {
			TEExecScript *pES = new TEExecScript();
			pES->bsScript = ::SysAllocString(param[0].lpolestr);
			pES->bsLang = ::SysAllocString(param[1].lpolestr);
			if (pv && pv->vt == VT_DISPATCH) {
				CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pv->pdispVal, &pES->pStream);
			}
			teSetPtr(pVarResult, _beginthread(&threadExecScript, 0, pES));
			return;
		}
	}
	param[TE_API_RESULT].lVal = ParseScript(param[0].lpolestr, param[1].lpolestr, pv, NULL, param[TE_EXCEPINFO].pExcepInfo);
	teSetLong(pVarResult, param[TE_API_RESULT].lVal);
}

VOID teApiGetScriptDispatch(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		VARIANT *pv = NULL;
		if (nArg >= 2) {
			pv = &pDispParams->rgvarg[nArg - 2];
		}
		IDispatch *pdisp = NULL;
		param[TE_API_RESULT].lVal = ParseScript(param[0].lpolestr, param[1].lpolestr, pv, &pdisp, param[TE_EXCEPINFO].pExcepInfo);
		teSetObjectRelease(pVarResult, pdisp);
	}
}

VOID teApiGetDispatch(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		IDispatch *pdisp;
		if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
			DISPID dispid = DISPID_UNKNOWN;
			LPOLESTR lp = param[1].lpolestr;
			if (pdisp->GetIDsOfNames(IID_NULL, &lp, 1, LOCALE_USER_DEFAULT, &dispid) == S_OK) {
				teSetObjectRelease(pVarResult, new CteDispatch(pdisp, 0, dispid));
				pdisp->Release();
			}
			pdisp->Release();
		}
	}
}

VOID teApiSHChangeNotifyRegister(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SHChangeNotifyEntry entry;
	teGetIDListFromVariant(const_cast<LPITEMIDLIST *>(&entry.pidl), &pDispParams->rgvarg[nArg - 4]);
	entry.fRecursive = param[5].boolVal;
	if (entry.pidl) {
		teSetLong(pVarResult, SHChangeNotifyRegister(param[0].hwnd, param[1].intVal, param[2].lVal, param[3].uintVal, 1, &entry));
		teChangeWindowMessageFilterEx(param[0].hwnd, param[3].uintVal, MSGFLT_ALLOW, NULL);
	}
}

VOID teApiMessageBox(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, MessageBox(param[0].hwnd, param[1].lpcwstr, param[2].lpcwstr, param[3].uintVal));
}

VOID teApiImageList_GetIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_GetIcon(param[0].himagelist, param[1].intVal, param[2].uintVal));
}

VOID teApiImageList_Create(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_Create(param[0].intVal, param[1].intVal, param[2].uintVal, param[3].intVal, param[4].intVal));
}

VOID teApiGetWindowLongPtr(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetWindowLongPtr(param[0].hwnd, param[1].intVal));
}

VOID teApiGetClassLongPtr(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetClassLongPtr(param[0].hwnd, param[1].intVal));
}

VOID teApiGetSubMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetSubMenu(param[0].hmenu, param[1].intVal));
}

VOID teApiSelectObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SelectObject(param[0].hdc, param[1].hgdiobj));
}

VOID teApiGetStockObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetStockObject(param[0].intVal));
}

VOID teApiGetSysColorBrush(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetSysColorBrush(param[0].intVal));
}

VOID teApiSetFocus(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetFocus(param[0].hwnd));
	if (param[0].hwnd == g_hwndMain) {
		CteShellBrowser *pSB = g_pTC->GetShellBrowser(g_pTC->m_nIndex);
		if (pSB) {
			pSB->SetActive(FALSE);
		}
	}
}

VOID teApiGetDC(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetDC(param[0].hwnd));
}

VOID teApiCreateCompatibleDC(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateCompatibleDC(param[0].hdc));
}

VOID teCreateMenu(HMENU hMenu, VARIANT *pVarResult)
{
	MENUINFO menuInfo;
	menuInfo.cbSize = sizeof(MENUINFO);
	menuInfo.fMask = MIM_STYLE;
	GetMenuInfo(hMenu, &menuInfo);
	menuInfo.dwStyle = (menuInfo.dwStyle & ~MNS_NOCHECK) | MNS_CHECKORBMP;
	SetMenuInfo(hMenu, &menuInfo);
	teSetPtr(pVarResult, hMenu);
}

VOID teApiCreatePopupMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teCreateMenu(CreatePopupMenu(), pVarResult);
}

VOID teApiCreateMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teCreateMenu(CreateMenu(), pVarResult);
}

VOID teApiCreateCompatibleBitmap(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateCompatibleBitmap(param[0].hdc, param[1].intVal, param[2].intVal));
}

VOID teApiSetWindowLongPtr(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetWindowLongPtr(param[0].hwnd, param[1].intVal, param[2].long_ptr));
}

VOID teApiSetClassLongPtr(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetClassLongPtr(param[0].hwnd, param[1].intVal, param[2].long_ptr));
}

VOID teApiImageList_Duplicate(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_Duplicate(param[0].himagelist));
}

VOID teApiSendMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SendMessage(param[0].hwnd, param[1].uintVal, param[2].wparam, param[3].lparam));
}

VOID teApiGetSystemMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetSystemMenu(param[0].hwnd, param[1].boolVal));
}

VOID teApiGetWindowDC(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetWindowDC(param[0].hwnd));
}

VOID teApiCreatePen(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreatePen(param[0].intVal, param[1].intVal, param[2].colorref));
}

VOID teApiSetCapture(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetCapture(param[0].hwnd));
}

VOID teApiSetCursor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetCursor(param[0].hcursor));
}

VOID teApiCallWindowProc(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CallWindowProc(param[0].wndproc, param[1].hwnd, param[2].uintVal, param[3].wparam, param[4].lparam));
}

VOID teApiGetWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 1) {
		teSetPtr(pVarResult, GetWindow(param[0].hwnd, param[1].uintVal));
		return;
	}
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		HWND hwnd;
		IUnknown_GetWindow(punk, &hwnd);
		teSetPtr(pVarResult, hwnd);
	}
}

VOID teApiGetTopWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetTopWindow(param[0].hwnd));
}

VOID teApiOpenProcess(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, OpenProcess(param[0].dword, param[1].boolVal, param[2].dword));
}

VOID teApiGetParent(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetParent(param[0].hwnd));
}

VOID teApiGetCapture(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetCapture());
}

VOID teApiGetModuleHandle(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetModuleHandle(param[0].lpcwstr));
}

VOID teApiSHGetImageList(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HANDLE h = 0;
	if FAILED(SHGetImageList(param[0].intVal, IID_IImageList, (LPVOID *)&h)) {
		h = 0;
	}
	teSetPtr(pVarResult, h);
}

VOID teApiCopyImage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CopyImage(param[0].handle, param[1].uintVal, param[2].intVal, param[3].intVal, param[4].uintVal));
}

VOID teApiGetCurrentProcess(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetCurrentProcess());
}

VOID teApiImmGetContext(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImmGetContext(param[0].hwnd));
}

VOID teApiGetFocus(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetFocus());
}

VOID teApiGetForegroundWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetForegroundWindow());
}

VOID teApiSetTimer(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetTimer(param[0].hwnd, param[1].uint_ptr, param[2].uintVal, NULL));
}

VOID teApiLoadMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadMenu(param[0].hinstance, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teApiLoadIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadIcon(param[0].hinstance, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teApiLoadLibraryEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadLibraryEx(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), param[1].handle, param[2].dword));
}

VOID teApiLoadImage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadImage(param[0].hinstance, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
		param[2].uintVal,	param[3].intVal, param[4].intVal, param[5].uintVal));
}

VOID teApiImageList_LoadImage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_LoadImage(param[0].hinstance, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
	param[2].intVal, param[3].intVal, param[4].colorref, param[5].uintVal, param[6].uintVal));
}

VOID teApiSHGetFileInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teParam	Path;
	IUnknown *punk;
	BOOL bItemIDList = FindUnknown(&pDispParams->rgvarg[nArg], &punk);
	if (bItemIDList) {
		bItemIDList = teGetIDListFromVariant(&Path.lpitemidlist, &pDispParams->rgvarg[nArg]);
	} else {
		Path.lpcwstr = param[0].lpcwstr;
	}
	teSetPtr(pVarResult, SHGetFileInfo(Path.lpcwstr, param[1].dword, param[2].lpshfileinfo, param[3].uintVal, param[4].uintVal));
	if (bItemIDList) {
		teCoTaskMemFree(Path.lpitemidlist);
	}
}

VOID teApiCreateWindowEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateWindowEx(param[0].dword, param[1].lpcwstr, param[2].lpcwstr,
		param[3].dword, param[4].intVal, param[5].intVal, param[6].intVal, param[7].intVal,
		param[8].hwnd, param[9].hmenu, param[10].hinstance, param[11].lpvoid));
}

VOID teApiShellExecute(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ShellExecute(param[0].hwnd, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
		GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 2]), param[3].lpcwstr, param[4].lpcwstr, param[5].intVal));
}

VOID teApiBeginPaint(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, BeginPaint(param[0].hwnd, param[1].lppaintstruct));
}

VOID teApiLoadCursor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadCursor(param[0].hinstance, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teApiLoadCursorFromFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadCursorFromFile(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teReleaseInvoke(TEInvoke *pInvoke)
{
	try {
		if (::InterlockedDecrement(&pInvoke->cRef) == 0) {
			SafeRelease(&pInvoke->pdisp);
			teClearVariantArgs(pInvoke->cArgs, pInvoke->pv);
			teILFreeClear(&pInvoke->pidl);
			teSysFreeString(&pInvoke->bsPath);
			delete [] pInvoke;
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teReleaseInvoke";
#endif
	}
}

VOID teAsyncInvoke(WORD wMode, int nArg, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDispatch *pdisp;
	if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
		TEInvoke *pInvoke = new TEInvoke[1];
		pInvoke->pdisp = pdisp;
		int dwms = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
		if (dwms == 0) {
			dwms = g_param[TE_NetworkTimeout];
		}
		pInvoke->dispid = DISPID_VALUE;
		pInvoke->cRef = 2;
		pInvoke->cDo = 1;
		pInvoke->cArgs = nArg - 1;
		pInvoke->pv = GetNewVARIANT(nArg);
		pInvoke->wErrorHandling = 3;
		pInvoke->wMode = wMode;
		pInvoke->pidl = NULL;
		pInvoke->hr = E_ABORT;
		for (int i = nArg - 1; i--;) {
			VariantCopy(&pInvoke->pv[i], &pDispParams->rgvarg[i]);
		}
		if ((pInvoke->pv[pInvoke->cArgs - 1].vt == VT_BSTR) || (wMode && teGetIDListFromVariant(&pInvoke->pidl, &pInvoke->pv[pInvoke->cArgs - 1]))) {
			pInvoke->bsPath = (pInvoke->pv[pInvoke->cArgs - 1].vt == VT_BSTR) ? ::SysAllocString(pInvoke->pv[pInvoke->cArgs - 1].bstrVal) : NULL;
			if (dwms > 0) {
				SetTimer(g_hwndMain, (UINT_PTR)pInvoke, dwms, teTimerProcParse);
			}
			uintptr_t uResult = _beginthread(&threadParseDisplayName, 0, pInvoke);
			if (uResult != -1) {
				teSetPtr(pVarResult, uResult);
				return;
			}
		}
		FolderItem *pid;
		GetFolderItemFromVariant(&pid, &pInvoke->pv[pInvoke->cArgs - 1]);
		VariantClear(&pInvoke->pv[pInvoke->cArgs - 1]);
		teSetObjectRelease(&pInvoke->pv[pInvoke->cArgs - 1], pid);
		Invoke5(pdisp, DISPID_VALUE, DISPATCH_METHOD, NULL, -pInvoke->cArgs, pInvoke->pv);
		--pInvoke->cRef;
		teReleaseInvoke(pInvoke);
		teSetPtr(pVarResult, 0);
	}
}

VOID teApiSHParseDisplayName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teAsyncInvoke(0, nArg, pDispParams, pVarResult);
}

VOID teApiSHChangeNotification_Lock(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDispatch *pdisp;
	if (GetDispatch(&pDispParams->rgvarg[nArg - 2], &pdisp)) {
		PIDLIST_ABSOLUTE *ppidl;
		LONG lEvent;
		HANDLE hLock;
		try {
			if (hLock = SHChangeNotification_Lock(param[0].handle, param[1].dword, &ppidl, &lEvent)) {
				VARIANT v;
				VariantInit(&v);
				if (ppidl) {
					for (int i = (lEvent & (SHCNE_RENAMEFOLDER | SHCNE_RENAMEITEM)) ? 2 : 1; i--;) {
						CteFolderItem *pPF = new CteFolderItem(NULL);
						pPF->Initialize(ppidl[i]);
						teSetObjectRelease(&v, pPF);
						tePutPropertyAt(pdisp, i, &v);
					}
				}
				SHChangeNotification_Unlock(hLock);
				teSetPtr(pVarResult, hLock);
				teSetLong(&v, lEvent);
				tePutProperty(pdisp, L"lEvent", &v);
			}
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"SHChangeNotification_Lock";
#endif
		}
		pdisp->Release();
	}
}

VOID teApiGetWindowText(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].hwnd == g_hwndMain) {
		teSetSZ(pVarResult, g_bsTitle);
		return;
	}
	BSTR bs = NULL;
	int nLen = GetWindowTextLength(param[0].hwnd);
	if (nLen) {
		bs = ::SysAllocStringLen(NULL, nLen);
		nLen = GetWindowText(param[0].hwnd, bs, nLen + 1);
	}
	teSetBSTR(pVarResult, &bs, nLen);
}

VOID teApiGetClassName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = ::SysAllocStringLen(NULL, MAX_CLASS_NAME);
	int nLen = GetClassName(param[0].hwnd, bsResult, MAX_CLASS_NAME);
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiGetModuleFileName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = NULL;
	int nLen = teGetModuleFileName(param[0].hmodule, &bsResult);
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiGetCommandLine(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		pVarResult->bstrVal = teGetCommandLine();
		pVarResult->vt = VT_BSTR;
	}
}

VOID teApiGetCurrentDirectory(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = ::SysAllocStringLen(NULL, MAX_PATH);
	int nLen = GetCurrentDirectory(MAX_PATH, bsResult);
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiGetMenuString(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = NULL;
	int nLen = teGetMenuString(&bsResult, param[0].hmenu, param[1].uintVal, param[2].lVal != MF_BYCOMMAND);
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiGetDisplayNameOf(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teGetDisplayNameOf(&pDispParams->rgvarg[nArg], param[1].intVal, pVarResult);
}

VOID teApiGetKeyNameText(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = ::SysAllocStringLen(NULL, MAX_PATH);
	int nLen = GetKeyNameText(param[0].lVal, bsResult, MAX_PATH);
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiDragQueryFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[1].lVal == -1) {
		teSetLong(pVarResult, DragQueryFile(param[0].hdrop, param[1].uintVal, NULL, 0));
		return;
	}
	BSTR bsResult = NULL;
	int nLen = teDragQueryFile(param[0].hdrop, param[1].uintVal, &bsResult);
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiSysAllocString(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 1 && teVarIsNumber(&pDispParams->rgvarg[nArg - 1])) {
		if (pVarResult) {
			UCHAR *pc;
			VARIANT vMem;
			VariantInit(&vMem);
			UINT nLen = GetpDataFromVariant(&pc, &pDispParams->rgvarg[nArg], &vMem);
			pVarResult->vt = VT_BSTR;
			pVarResult->bstrVal = teMultiByteToWideChar(param[1].lVal, (LPCSTR)pc, nLen);
			VariantClear(&vMem);
		}
		return;
	}
	teSetSZ(pVarResult, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]));
}

VOID teApiSysAllocStringLen(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		VARIANT *pv = &pDispParams->rgvarg[nArg];
		pVarResult->bstrVal = (pv->vt == VT_BSTR) ? teSysAllocStringLenEx(pv->bstrVal, param[1].uintVal) : teSysAllocStringLen(GetLPWSTRFromVariant(pv), param[1].uintVal);
		pVarResult->vt = VT_BSTR;
	}
}

VOID teApiSysAllocStringByteLen(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		UCHAR *pc;
		VARIANT vMem;
		VariantInit(&vMem);
		UINT nLen = GetpDataFromVariant(&pc, &pDispParams->rgvarg[nArg], &vMem);
		pVarResult->bstrVal = teSysAllocStringByteLen((LPCSTR)pc, param[1].uintVal, nLen ? nLen : param[2].intVal);
		pVarResult->vt = VT_BSTR;
		VariantClear(&vMem);
	}
}

VOID teApisprintf(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	int nSize = param[0].intVal;
	LPWSTR pszFormat = param[1].lpwstr;
	if (pszFormat && nSize > 0) {
		BSTR bsResult = ::SysAllocStringLen(NULL, nSize);
		int nLen = 0;
		int nIndex = 1;
		int nPos = 0;
		while (pszFormat[nPos] && nLen < nSize) {
			while (pszFormat[nPos] && pszFormat[nPos++] != '%') {
			}
			while (WCHAR wc = pszFormat[nPos]) {
				++nPos;
				if (wc == '%') {//Escape %
					lstrcpyn(&bsResult[nLen], pszFormat, nPos);
					nLen += nPos - 1;
					pszFormat += nPos;
					nPos = 0;
					break;
				}
				if (nIndex <= nArg) {
					if (StrChrIW(L"diuoxc", wc)) {//Integer
						wc = pszFormat[nPos];
						pszFormat[nPos] = NULL;
						swprintf_s(&bsResult[nLen], nSize - nLen, pszFormat, GetLLFromVariant(&pDispParams->rgvarg[nArg - ++nIndex]));
						nLen += lstrlen(&bsResult[nLen]);
						pszFormat += nPos;
						nPos = 0;
						pszFormat[0] = wc;
						break;
					}
					if (StrChrIW(L"efga", wc)) {//Real
						wc = pszFormat[nPos];
						pszFormat[nPos] = NULL;
						VARIANT v;
						teVariantChangeType(&v, &pDispParams->rgvarg[nArg - ++nIndex], VT_R8);
						swprintf_s(&bsResult[nLen], nSize - nLen, pszFormat, v.dblVal);
						nLen += lstrlen(&bsResult[nLen]);
						pszFormat += nPos;
						nPos = 0;
						pszFormat[0] = wc;
						break;
					}
					if (towlower(wc) == 's') {//String
						wc = pszFormat[nPos];
						pszFormat[nPos] = NULL;
						VARIANT v;
						teVariantChangeType(&v, &pDispParams->rgvarg[nArg - ++nIndex], VT_BSTR);
						swprintf_s(&bsResult[nLen], nSize - nLen, pszFormat, v.bstrVal);
						VariantClear(&v);
						nLen += lstrlen(&bsResult[nLen]);
						pszFormat += nPos;
						nPos = 0;
						pszFormat[0] = wc;
						break;
					}
					if (!StrChrIW(L"0123456789-+#.hljzt ", wc)) {//not Specifier
						lstrcpyn(&bsResult[nLen], pszFormat, nPos + 1);
						nLen += nPos;
						pszFormat += nPos;
						nPos = 0;
						break;
					}
				}
			}
		}
		if (nLen < nSize && pszFormat[0]) {
			lstrcpyn(&bsResult[nLen], pszFormat, nSize - nLen);
			nLen += lstrlen(&bsResult[nLen]);
		}
		teSetBSTR(pVarResult, &bsResult, nLen);
	}
}

VOID teApibase64_encode(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	UCHAR *pc;
	VARIANT vMem;
	VariantInit(&vMem);
	UINT nLen = GetpDataFromVariant(&pc, &pDispParams->rgvarg[nArg], &vMem);
	if (nLen) {
		DWORD dwSize;
		CryptBinaryToString(pc, nLen, CRYPT_STRING_BASE64, NULL, &dwSize);
		BSTR bsResult = NULL;
		if (dwSize > 0) {
			bsResult = ::SysAllocStringLen(NULL, dwSize - 1);
			CryptBinaryToString(pc, nLen, CRYPT_STRING_BASE64, bsResult, &dwSize);
		}
		teSetBSTR(pVarResult, &bsResult, -1);
	}
	VariantClear(&vMem);
}

VOID teApiLoadString(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = ::SysAllocStringLen(NULL, SIZE_BUFF);
	int nLen = LoadString(param[0].hinstance, param[1].uintVal, bsResult, SIZE_BUFF);
	teSetBSTR(pVarResult, &bsResult, -1);
}

VOID teApiAssocQueryString(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IQueryAssociations *pQA = NULL;
	DWORD cch = 0;
	BSTR bsResult = NULL;
	int nLen = -1;
	IDataObject *pDataObj;
	if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg - 2])) {
		long nCount;
		LPITEMIDLIST *ppidllist = IDListFormDataObj(pDataObj, &nCount);
		AdjustIDList(ppidllist, nCount);
		if (nCount >= 1) {
			IShellFolder *pSF;
			if (GetShellFolder(&pSF, ppidllist[0])) {
				if SUCCEEDED(pSF->GetUIObjectOf(g_hwndMain, nCount, const_cast<LPCITEMIDLIST *>(&ppidllist[1]), IID_IQueryAssociations, NULL, (LPVOID*)&pQA)) {
					if SUCCEEDED(pQA->GetString(param[0].assocf, param[1].assocstr, param[3].lpcwstr, NULL, &cch)) {
						if (cch) {
							nLen = cch - 1;
							bsResult = ::SysAllocStringLen(NULL, nLen);
							pQA->GetString(param[0].assocf, param[1].assocstr, param[3].lpcwstr, bsResult, &cch);
						}
						SafeRelease(&pQA);
					}
				}
				pSF->Release();
			}
			do {
				teCoTaskMemFree(ppidllist[nCount]);
			} while (--nCount >= 0);
		}
		delete [] ppidllist;
		pDataObj->Release();
	}
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiStrFormatByteSize(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = ::SysAllocStringLen(NULL, MAX_COLUMN_NAME_LEN);
	teStrFormatSize(nArg >= 1 ? param[1].dword : 2, param[0].llVal, bsResult, MAX_COLUMN_NAME_LEN);
	teSetBSTR(pVarResult, &bsResult, -1);
}

VOID teApiGetDateFormat(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SYSTEMTIME SysTime;
	if (teGetSystemTime(&pDispParams->rgvarg[nArg - 2], &SysTime)) {
		int cch = GetDateFormat(param[0].lcid, param[1].dword, &SysTime, param[3].lpcwstr, NULL, 0);
		if (cch > 0) {
			BSTR bsResult = ::SysAllocStringLen(NULL, cch - 1);
			GetDateFormat(param[0].lcid, param[1].dword, &SysTime, param[3].lpcwstr, bsResult, cch);
			teSetBSTR(pVarResult, &bsResult, cch - 1);
		}
	}
}

VOID teApiGetTimeFormat(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SYSTEMTIME SysTime;
	if (teGetSystemTime(&pDispParams->rgvarg[nArg - 2], &SysTime)) {
		int cch = GetTimeFormat(param[0].lcid, param[1].dword, &SysTime, param[3].lpcwstr, NULL, 0);
		if (cch > 0) {
			BSTR bsResult = ::SysAllocStringLen(NULL, cch - 1);
			GetTimeFormat(param[0].lcid, param[1].dword, &SysTime, param[3].lpcwstr, bsResult, cch);
			teSetBSTR(pVarResult, &bsResult, cch - 1);
		}
	}
}

VOID teApiGetLocaleInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	int cch = GetLocaleInfo(param[0].lcid, param[1].lctype, NULL, 0);
	if (cch > 0) {
		BSTR bsResult = ::SysAllocStringLen(NULL, cch - 1);
		GetLocaleInfo(param[0].lcid, param[1].lctype, bsResult, cch);
		teSetBSTR(pVarResult, &bsResult, -1);
	}
}

VOID teApiMonitorFromPoint(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg]);
	teSetPtr(pVarResult, MonitorFromPoint(pt, param[1].dword));
}

VOID teApiMonitorFromRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, MonitorFromRect(param[0].lpcrect, param[1].dword));
}

VOID teApiGetMonitorInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMonitorInfo(param[0].hmonitor, param[1].lpmonitorinfo));
}

VOID teApiGlobalAddAtom(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GlobalAddAtom(param[0].lpcwstr));
}

VOID teApiGlobalGetAtomName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	TCHAR szBuf[256];
	UINT uSize = GlobalGetAtomName(param[0].atom, szBuf, ARRAYSIZE(szBuf));
	teSetSZ(pVarResult, szBuf);
}

VOID teApiGlobalFindAtom(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GlobalFindAtom(param[0].lpcwstr));
}

VOID teApiGlobalDeleteAtom(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GlobalDeleteAtom(param[0].atom));
}

VOID teApiRegisterHotKey(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, RegisterHotKey(param[0].hwnd, param[1].intVal, param[2].uintVal, param[3].uintVal));
}

VOID teApiUnregisterHotKey(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, UnregisterHotKey(param[0].hwnd, param[1].intVal));
}

VOID teApiILGetCount(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		teSetLong(pVarResult, ILGetCount(pidl));
		teCoTaskMemFree(pidl);
	}
}

VOID teApiSHTestTokenMembership(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SHTestTokenMembership(param[0].handle, param[1].ulVal));
}

VOID teApiObjGetI(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDispatch *pdisp;
	if (pVarResult && GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
		teGetPropertyI(pdisp, param[1].lpolestr, pVarResult);
		if (pVarResult->vt == VT_DATE) {//UNIX EPOCH
			ULONGLONG ft;
			teVariantTimeToFileTime(pVarResult->date, (FILETIME *)&ft);
			teSetLL(pVarResult, ft / 10000 - 11644473600000LL);
		}
		pdisp->Release();
	}
}

VOID teApiObjPutI(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		teSetLong(pVarResult, tePutProperty0(punk, param[1].lpolestr, &pDispParams->rgvarg[nArg - 2], fdexNameEnsure | fdexNameCaseInsensitive));
	}
}

VOID teApiObjDispId(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DISPID dispid = DISPID_UNKNOWN;
	IDispatch *pdisp;
	if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
		LPOLESTR lpsz = param[1].lpolestr;
		HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &lpsz, 1, LOCALE_USER_DEFAULT, &dispid);
		pdisp->Release();
	}
	teSetLong(pVarResult, dispid);
}

VOID teApiDrawText(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, DrawText(param[0].hdc, param[1].lpcwstr, param[2].intVal, param[3].lprect, param[4].uintVal));
}

VOID teApiRectangle(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, Rectangle(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal));
}

VOID teApiPathIsDirectory(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 2 && pDispParams->rgvarg[nArg].vt == VT_DISPATCH) {
		teAsyncInvoke(1, nArg, pDispParams, pVarResult);
		return;
	}
	HRESULT hr = tePathIsDirectory(param[0].lpwstr, param[1].dword, param[2].intVal);
	if (hr != E_ABORT) {
		if (hr <= 0) {
			teSetBool(pVarResult, SUCCEEDED(hr));
		} else {
			teSetLong(pVarResult, hr);
		}
	}
}

VOID teApiPathIsSameRoot(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, PathIsSameRoot(param[0].lpcwstr, param[1].lpcwstr));
}

VOID teApiSHSimpleIDListFromPath(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	FILETIME ft = { 0 };
	if (nArg >= 2) {
		DATE dt;
		if (teGetVariantTime(&pDispParams->rgvarg[nArg - 2], &dt) && dt >= 0) {
			teVariantTimeToFileTime(dt, &ft);
		}
	}
	LPITEMIDLIST pidl = teSHSimpleIDListFromPathEx(param[0].lpwstr, param[1].dword, param[3].uintVal, param[3].uli.HighPart, &ft);
	teSetIDListRelease(pVarResult, &pidl);
}

VOID teApiOutputDebugString(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	::OutputDebugString(param[0].lpcwstr);
}

VOID teApiDllGetClassObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDispatch *pdisp;
	HMODULE hDll = teCreateInstanceV(&pDispParams->rgvarg[nArg], &pDispParams->rgvarg[nArg - 1], IID_PPV_ARGS(&pdisp));
	if (hDll) {
		CteDispatch *odisp = new CteDispatch(pdisp, 4, DISPID_UNKNOWN);
		pdisp->Release();
		odisp->m_hDll = hDll;
		teSetObjectRelease(pVarResult, odisp);
	}
}

VOID teApiIsAppThemed(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsAppThemed());
}

VOID teApiSHDefExtractIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SHDefExtractIcon(param[0].lpcwstr, param[1].intVal, param[2].uintVal, param[3].phicon, param[4].phicon, param[5].uintVal));
}

VOID teApiURLDownloadToFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HRESULT hr = E_FAIL;
	IUnknown *punk = NULL;
	if (FindUnknown(&pDispParams->rgvarg[nArg - 1], &punk)) {
		IStream *pDst, *pSrc;
		hr = punk->QueryInterface(IID_PPV_ARGS(&pSrc));
		if SUCCEEDED(hr) {
			hr = SHCreateStreamOnFileEx(param[2].bstrVal, STGM_WRITE | STGM_CREATE | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, TRUE, NULL, &pDst);
			if SUCCEEDED(hr) {
				teCopyStream(pSrc, pDst);
				pDst->Release();
			}
			pSrc->Release();
		}
		if FAILED(hr) {
			IDispatch *pdisp = NULL;
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pdisp))) {
				VARIANT v;
				VariantInit(&v);
				if SUCCEEDED(teGetProperty(pdisp, L"responseBody", &v)) {
					UCHAR *pc;
					VARIANT vMem;
					VariantInit(&vMem);
					UINT nLen = GetpDataFromVariant(&pc, &v, &vMem);
					if (nLen) {
						HANDLE hFile = CreateFile(param[2].bstrVal, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
						if (hFile != INVALID_HANDLE_VALUE) {
							DWORD dwWriteByte;
							if (WriteFile(hFile, pc, nLen, &dwWriteByte, NULL)) {
								hr = S_OK;
							}
							CloseHandle(hFile);
						}
					}
					VariantClear(&vMem);
					VariantClear(&v);
				}
			}
			pdisp->Release();
			teSetLong(pVarResult, hr);
			return;
		}
	}
	if (teStartsText(L"data:", param[1].bstrVal)) {
		LPWSTR lpBase64 = StrChr(param[1].bstrVal, ',');
		if (lpBase64) {
			DWORD dwData;
			DWORD dwLen = lstrlen(++lpBase64);
			if (CryptStringToBinary(lpBase64, dwLen, CRYPT_STRING_BASE64, NULL, &dwData, NULL, NULL) && dwData > 0) {
				BSTR bs = ::SysAllocStringByteLen(NULL, dwData);
				CryptStringToBinary(lpBase64, dwLen, CRYPT_STRING_BASE64, (BYTE *)bs, &dwData, NULL, NULL);
				HANDLE hFile = CreateFile(param[2].bstrVal, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
				if (hFile != INVALID_HANDLE_VALUE) {
					DWORD dwWriteByte;
					if (WriteFile(hFile, bs, dwData, &dwWriteByte, NULL)) {
						hr = S_OK;
					}
					CloseHandle(hFile);
				}
				::SysFreeString(bs);
			}
		}
		teSetLong(pVarResult, hr);
		return;
	}
	FindUnknown(&pDispParams->rgvarg[nArg], &punk);
	teSetLong(pVarResult, URLDownloadToFile(punk, param[1].bstrVal, param[2].bstrVal, param[3].dword, NULL));
}

VOID teApiMoveFileEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, MoveFileEx(param[0].lpcwstr, param[1].lpcwstr, param[2].dword));
}

VOID teApiObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetObjectRelease(pVarResult, teCreateObj(TE_OBJECT, NULL));
}

VOID teApiArray(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetObjectRelease(pVarResult, teCreateObj(TE_ARRAY, NULL));
}

VOID teApiPathIsRoot(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, PathIsRoot(param[0].lpcwstr));
}

VOID teApiEnableWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, EnableWindow(param[0].hwnd, param[1].boolVal));
}

VOID teApiIsWindowEnabled(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsWindowEnabled(param[0].hwnd));
}

VOID teApiCryptProtectData(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DATA_BLOB blob[3];
	for (int i = 2; i--;) {
		blob[i].cbData = ::SysStringByteLen(param[i].bstrVal);
		blob[i].pbData = param[i].pbVal;
	}
	if (CryptProtectData(&blob[0], NULL, &blob[1], NULL, NULL, 0, &blob[2])) {
		teCreateSafeArray(pVarResult, blob[2].pbData, blob[2].cbData, param[2].boolVal);
		LocalFree(blob[2].pbData);
	}
}

VOID teApiCryptUnprotectData(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	UCHAR *pc;
	VARIANT vMem;
	VariantInit(&vMem);
	UINT nLen = GetpDataFromVariant(&pc, &pDispParams->rgvarg[nArg], &vMem);
	if (nLen && pVarResult) {
		DATA_BLOB blob[3];
		blob[0].cbData = nLen;
		blob[0].pbData = pc;
		blob[1].cbData = ::SysStringByteLen(param[1].bstrVal);
		blob[1].pbData = param[1].pbVal;
		if (CryptUnprotectData(&blob[0], NULL, &blob[1], NULL, NULL, 0, &blob[2])) {
			teCreateSafeArray(pVarResult, blob[2].pbData, blob[2].cbData, param[2].boolVal);
			LocalFree(blob[2].pbData);
		}
	}
	VariantClear(&vMem);
}

VOID teApibase64_decode(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DWORD dwData;
	if (pVarResult) {
		if (CryptStringToBinary(param[0].bstrVal, ::SysStringLen(param[0].bstrVal), CRYPT_STRING_BASE64, NULL, &dwData, NULL, NULL) && dwData > 0) {
			if (param[1].boolVal) {
				pVarResult->vt = VT_BSTR;
				pVarResult->bstrVal = ::SysAllocStringByteLen(NULL, dwData);
				CryptStringToBinary(param[0].bstrVal, ::SysStringLen(param[0].bstrVal), CRYPT_STRING_BASE64, (BYTE *)pVarResult->bstrVal, &dwData, NULL, NULL);
			} else {
				SAFEARRAY *psa = SafeArrayCreateVector(VT_UI1, 0, dwData);
				if (psa) {
					PVOID pvData;
					if (::SafeArrayAccessData(psa, &pvData) == S_OK) {
						pVarResult->vt = VT_ARRAY | VT_UI1;
						pVarResult->parray = psa;
						CryptStringToBinary(param[0].bstrVal, ::SysStringLen(param[0].bstrVal), CRYPT_STRING_BASE64, (BYTE *)pvData, &dwData, NULL, NULL);
						::SafeArrayUnaccessData(psa);
					}
				}
			}
		}
	}
}

VOID teApiAlphaBlend(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, AlphaBlend(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal,
		param[5].hdc, param[6].intVal, param[7].intVal, param[8].intVal, param[9].intVal, *(BLENDFUNCTION*)&param[10].lVal));
}

VOID teApiStretchBlt(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SetStretchBltMode(param[0].hdc, HALFTONE);
	teSetBool(pVarResult, StretchBlt(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal,
		param[5].hdc, param[6].intVal, param[7].intVal, param[8].intVal, param[9].intVal, param[10].dword));
}

VOID teApiTransparentBlt(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, TransparentBlt(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal,
		param[5].hdc, param[6].intVal, param[7].intVal, param[8].intVal, param[9].intVal, param[10].uintVal));
}

VOID teApiFormatMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPWSTR lpBuf;
	if (FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ARGUMENT_ARRAY,
		param[0].lpwstr, 0, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&lpBuf, MAX_STATUS, (va_list *)&param[1])) {
		teSetSZ(pVarResult, lpBuf);
		LocalFree(lpBuf);
	}
}

VOID teApiGetGUIThreadInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetGUIThreadInfo(param[0].dword, param[1].pgui));
}

VOID teApiCreateSolidBrush(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateSolidBrush(param[0].colorref));
}

VOID teApiPlgBlt(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SetStretchBltMode(param[0].hdc, HALFTONE);
	teSetBool(pVarResult, PlgBlt(param[0].hdc, param[1].lppoint, param[2].hdc, param[3].intVal, param[4].intVal,
		param[5].intVal, param[6].intVal, param[7].hbitmap, param[8].intVal, param[9].intVal));
}

VOID teApiRoundRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, RoundRect(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal, param[5].intVal, param[6].intVal));
}

VOID teApiCreateProcess(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HANDLE hRead, hWrite;
 	SECURITY_ATTRIBUTES	sa;
	STARTUPINFO si = { sizeof(STARTUPINFO) };
	PROCESS_INFORMATION pi = { 0 };
	DWORD dwLen;

	sa.nLength = sizeof(sa);
	sa.lpSecurityDescriptor	= 0;
	sa.bInheritHandle = TRUE;

	teSetLong(pVarResult, -1);
	if(!CreatePipe(&hRead, &hWrite, &sa, 0)) {
		return;
	}

	si.dwFlags = STARTF_USESTDHANDLES;
	si.hStdOutput = hWrite;
	si.hStdError = hWrite;
	DWORD dwms =  param[3].dword;
	if (dwms == 0) {
		dwms = INFINITE;
	}
	if (CreateProcess(NULL, param[0].lpwstr, NULL, NULL, TRUE, nArg >= 4 ? param[4].dword : CREATE_NO_WINDOW, NULL, param[1].lpwstr, &si, &pi)) {
		BSTR bsStdOut = NULL;
		if (WaitForInputIdle(pi.hProcess, dwms)) {
			if (WaitForSingleObject(pi.hProcess, dwms) != WAIT_TIMEOUT) {
				if (PeekNamedPipe(hRead, NULL, 0, NULL, &dwLen, NULL)) {
					BSTR bs = ::SysAllocStringByteLen(NULL, dwLen);
					if (dwLen) {
						::ReadFile(hRead, bs, dwLen, &dwLen, NULL);
						if (param[2].intVal == 2 || (param[2].intVal != 1 && IsTextUnicode(bs, dwLen, NULL))) {
							bsStdOut = bs;
							bs = NULL;
						} else {
							bsStdOut = teMultiByteToWideChar(param[2].uintVal < 3 ? CP_ACP : param[2].uintVal, (LPCSTR)bs, dwLen);
						}
					}
					teSysFreeString(&bs);
				}
			}
		}
		if (param[5].boolVal) {
			IDispatch *pdisp = teCreateObj(TE_OBJECT, NULL);
			VARIANT v;
			VariantInit(&v);
			teSetSZ(&v, bsStdOut);
			tePutProperty(pdisp, L"StdOut", &v);
			teSetPtr(&v, GetProcessId(pi.hProcess));
			tePutProperty(pdisp, L"ProcessId", &v);
			teSetObjectRelease(pVarResult, pdisp);
		} else {
			teSetSZ(pVarResult, bsStdOut);
		}
		teSysFreeString(&bsStdOut);
		CloseHandle(pi.hThread);
		CloseHandle(pi.hProcess);
	}
	CloseHandle(hRead);
	CloseHandle(hWrite);
}

VOID teApiDeleteFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteFile(param[0].lpcwstr));
}

VOID teApiCreateObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	CHAR pszNameA[16];
	::WideCharToMultiByte(CP_TE, 0, param[0].lpcwstr, -1, pszNameA, ARRAYSIZE(pszNameA) - 1, NULL, NULL);
	for (int i = 0; methodObject[i].name; ++i) {
		if (lstrcmpiA(pszNameA, methodObject[i].name) == 0) {
			teSetObjectRelease(pVarResult, teCreateObj(methodObject[i].id, nArg >= 1 ? &pDispParams->rgvarg[nArg - 1] : NULL));
			return;
		}
	}
	if (nArg >= 1 && lstrcmpiA(pszNameA, "SafeArray") == 0) {
		if (pVarResult) {
			IDispatch *pdisp;
			if (GetDispatch(&pDispParams->rgvarg[nArg - 1], &pdisp)) {
				teCreateSafeArrayFromVariantArray(pdisp, pVarResult);
				SafeRelease(&pdisp);
				return;
			}
		}
	}
	CLSID clsid;
	IUnknown *punk;
	if SUCCEEDED(teCLSIDFromProgID(param[0].lpwstr, &clsid)) {
		if SUCCEEDED(teCreateInstance(clsid, NULL, NULL, IID_PPV_ARGS(&punk))) {
			teSetObjectRelease(pVarResult, punk);
		}
	}
}

VOID teApiCreateEvent(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateEvent(param[0].lpSecurityAttributes, param[1].boolVal, param[2].boolVal, param[3].lpcwstr));
}

VOID teApiSetEvent(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetEvent(param[0].handle));
}

VOID teApiResetEvent(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ResetEvent(param[0].handle));
}

VOID teApiPlaySound(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, PlaySound(param[0].lpcwstr, param[1].hmodule, param[2].dword));
}

VOID teApiSHCreateStreamOnFileEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk = NULL;
	IStream *pStream0 = NULL;
	if (nArg >= 4 && FindUnknown(&pDispParams->rgvarg[nArg - 4], &punk)) {
		punk->QueryInterface(IID_PPV_ARGS(&pStream0));
	}
	IStream *pStream = NULL;
	if (nArg >= 5) {
		teSetLong(pVarResult, SHCreateStreamOnFileEx(param[0].lpcwstr, param[1].dword, param[2].dword, param[3].boolVal, pStream0, &pStream));
		if (pStream) {
			if (FindUnknown(&pDispParams->rgvarg[nArg - 5], &punk)) {
				VARIANT v;
				v.vt = VT_DISPATCH;
				v.pdispVal = new CteObject(pStream);
				tePutPropertyAt(punk, 0, &v);
			}
			pStream->Release();
		}
		return;
	}
	SHCreateStreamOnFileEx(param[0].lpcwstr, param[1].dword, param[2].dword, param[3].boolVal, pStream0, &pStream);
	teSetObjectRelease(pVarResult, new CteObject(pStream));
	SafeRelease(&pStream);
}

VOID teApiHasThumbnail(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	int nResult = 0;
	LPITEMIDLIST pidl;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		IShellFolder *pSF;
		LPCITEMIDLIST pidlPart;
		if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
			IExtractImage *pEI;
			IThumbnailProvider *pTP;
			if (pSF->GetUIObjectOf(NULL, 1, &pidlPart, IID_IThumbnailProvider, NULL, (void **)&pTP) == S_OK) {
				nResult |= 1;
				pTP->Release();
			}
			if (pSF->GetUIObjectOf(NULL, 1, &pidlPart, IID_IExtractImage, NULL, (void **)&pEI) == S_OK) {
				nResult |= 2;
				pEI->Release();
			}
			pSF->Release();
		}
/*		IThumbnailCache *pThumbnailCache;
		ISharedBitmap *pSharedBitmap;
		WTS_CACHEFLAGS  cacheFlags;
		WTS_THUMBNAILID thumbnailId;
		if SUCCEEDED(CoCreateInstance(CLSID_LocalThumbnailCache, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&pThumbnailCache))) {
			IShellItem *pShellItem;
			if SUCCEEDED(_SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&pShellItem))) {
				if (pThumbnailCache->GetThumbnail(pShellItem, 96, WTS_EXTRACTDONOTCACHE, &pSharedBitmap, &cacheFlags, &thumbnailId) == S_OK) {
					nResult |= 4;
					pSharedBitmap->Release();
				}
				pShellItem->Release();
			}
		}*/
		teILFreeClear(&pidl);
	}
	teSetLong(pVarResult, nResult);
}

VOID teApiGetDiskFreeSpaceEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	ULARGE_INTEGER FreeBytesOfAvailable;
	ULARGE_INTEGER TotalNumberOfBytes;
	ULARGE_INTEGER TotalNumberOfFreeBytes;
	if (pVarResult && GetDiskFreeSpaceEx(param[0].lpcwstr, &FreeBytesOfAvailable, &TotalNumberOfBytes, &TotalNumberOfFreeBytes)) {
		teSetObjectRelease(pVarResult, teCreateObj(TE_OBJECT, NULL));
		VARIANT v;
		teSetLL(&v, FreeBytesOfAvailable.QuadPart);
		tePutProperty(pVarResult->pdispVal, L"FreeBytesOfAvailable", &v);
		teSetLL(&v, TotalNumberOfBytes.QuadPart);
		tePutProperty(pVarResult->pdispVal, L"TotalNumberOfBytes", &v);
		teSetLL(&v, TotalNumberOfFreeBytes.QuadPart);
		tePutProperty(pVarResult->pdispVal, L"TotalNumberOfFreeBytes", &v);
	}
}

VOID teApiSetSysColors(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetSysColors(param[0].intVal, (INT *)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1], NULL), (COLORREF *)GetpcFromVariant(&pDispParams->rgvarg[nArg - 2], NULL)));
}

VOID teApiIsClipboardFormatAvailable(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsClipboardFormatAvailable(param[0].uintVal));
}

VOID teApiCreateFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::CreateFile(param[0].lpcwstr, param[1].dword, param[2].dword, param[3].lpSecurityAttributes, param[4].dword, param[5].dword, param[6].handle));
}

VOID teApiGetPixel(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ::GetPixel(param[0].hdc, param[1].intVal, param[2].intVal));
}

VOID teApiShouldAppsUseDarkMode(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, g_bDarkMode);
}

VOID teApiGetProcAddress(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	CHAR szProcNameA[MAX_LOADSTRING];
	LPSTR lpProcNameA = szProcNameA;
	if (pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
		::WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)pDispParams->rgvarg[nArg - 1].bstrVal, -1, lpProcNameA, MAX_LOADSTRING, NULL, NULL);
		if (!param[0].hmodule) {
			if (!lstrcmpiA(lpProcNameA, "GetImage")) {
				teSetPtr(pVarResult, teGetImage);
				return;
			}
		}
	} else {
		lpProcNameA = MAKEINTRESOURCEA(param[1].word);
	}
	teSetPtr(pVarResult, GetProcAddress(param[0].hmodule, lpProcNameA));
}

VOID teApiRunDLL(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	CHAR szProcNameA[MAX_LOADSTRING];
	LPSTR lpProcNameA = szProcNameA;
	if (pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
		::WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)pDispParams->rgvarg[nArg - 1].bstrVal, -1, lpProcNameA, MAX_LOADSTRING, NULL, NULL);
	} else {
		lpProcNameA = MAKEINTRESOURCEA(param[1].word);
	}
	LPFNEntryPointW _EntryPointW = (LPFNEntryPointW)GetProcAddress(param[0].hmodule, lpProcNameA);
	if (_EntryPointW) {
		_EntryPointW(param[2].hwnd, param[3].hinstance, param[4].lpwstr, param[5].intVal);
	}
}

VOID teApiGetCurrentThreadId(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetCurrentThreadId());
}

VOID teApiGetDpiForMonitor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HMONITOR hMonitor = param[0].hmonitor;
	if (!hMonitor || nArg == 0) {
		RECT rc;
		POINT pt;
		GetWindowRect(g_hwndMain, &rc);
		pt.x = rc.left;
		pt.y = rc.top;
		hMonitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
	}
	CteMemory *pstPt = new CteMemory(2 * sizeof(int), NULL, 1, L"POINT");
	UINT ux, uy;
	_GetDpiForMonitor(hMonitor, nArg >= 1 ? param[1].MonitorDpiType : MDT_EFFECTIVE_DPI, &ux, &uy);
	pstPt->SetPoint(ux, uy);
	teSetObjectRelease(pVarResult, pstPt);
}

VOID teApiHashData(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		BYTE *pc;
		VARIANT vMem;
		VariantInit(&vMem);
		UINT nLen = GetpDataFromVariant(&pc, &pDispParams->rgvarg[nArg], &vMem);
		BSTR bs = ::SysAllocStringByteLen(NULL, param[1].uintVal);
		HashData(pc, nLen, (BYTE *)bs, param[1].dword);
		DWORD dwSize;
		CryptBinaryToString(pc, param[1].uintVal, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, NULL, &dwSize);
		BSTR bsResult = NULL;
		if (dwSize > 0) {
			bsResult = ::SysAllocStringLen(NULL, dwSize - 1);
			CryptBinaryToString((BYTE *)bs, param[1].uintVal, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, bsResult, &dwSize);
		}
		::SysFreeString(bs);
		teSetBSTR(pVarResult, &bsResult, -1);
		VariantClear(&vMem);
	}
}

VOID teApiSetDCPenColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetDCPenColor(param[0].hdc, param[1].colorref));
}

VOID teApiSetDCBrushColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetDCBrushColor(param[0].hdc, param[1].colorref));
}

VOID teApiWriteFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	UCHAR *pc;
	VARIANT vMem;
	VariantInit(&vMem);
	UINT nLen = GetpDataFromVariant(&pc, &pDispParams->rgvarg[nArg - 1], &vMem);
	if (nLen) {
		DWORD dwWriteByte;
		IUnknown *punk;
		IStream *pStream;
		if (FindUnknown(&pDispParams->rgvarg[nArg], &punk) && SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pStream)))) {
			teSetLong(pVarResult, pStream->Write(pc, nLen, &dwWriteByte));
		} else {
			teSetBool(pVarResult, ::WriteFile(param[0].handle, pc, nLen, &dwWriteByte, NULL));
		}
	}
	VariantClear(&vMem);
}

VOID teApiReadFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		pVarResult->vt = VT_BSTR;
		pVarResult->bstrVal = ::SysAllocStringByteLen(NULL, param[1].uintVal);
		DWORD dwReadByte;
		BOOL bOK = FALSE;
		IUnknown *punk;
		IStream *pStream;
		if (FindUnknown(&pDispParams->rgvarg[nArg], &punk) && SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pStream)))) {
			bOK = SUCCEEDED(pStream->Read(pVarResult->bstrVal, param[1].uintVal, &dwReadByte));
		} else {
			bOK = ::ReadFile(param[0].handle, pVarResult->bstrVal, param[1].uintVal, &dwReadByte, NULL);
		}
		if (bOK) {
			BSTR bs = ::SysAllocStringByteLen(NULL, dwReadByte);
			::CopyMemory(bs, pVarResult->bstrVal, dwReadByte);
			teSysFreeString(&pVarResult->bstrVal);
			pVarResult->bstrVal = bs;
		} else {
			teSysFreeString(&pVarResult->bstrVal);
			pVarResult->vt = VT_EMPTY;
		}
	}
}

VOID teApiSetFilePointer(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	IStream *pStream;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk) && SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pStream)))) {
		ULARGE_INTEGER uli;
		if SUCCEEDED(pStream->Seek(param[1].li, param[2].dword, &uli)) {
			teSetLL(pVarResult, uli.QuadPart);
		}
	} else {
		LARGE_INTEGER li;
		li.HighPart = param[1].li.HighPart;
		li.LowPart = ::SetFilePointer(param[0].handle, param[1].li.LowPart, &li.HighPart, param[2].dword);
		teSetLL(pVarResult, li.QuadPart);
	}
}

VOID teApiSysStringByteLen(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	UCHAR *pc;
	VARIANT vMem;
	VariantInit(&vMem);
	teSetLong(pVarResult, GetpDataFromVariant(&pc, &pDispParams->rgvarg[nArg], &vMem));
	VariantClear(&vMem);
}

VOID teApiGetProp(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetProp(param[0].hwnd, param[1].lpcwstr));
}


VOID teApiInvoke(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDispatch *pdisp, *pArgs;
	if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
		int nLen = 0;
		if (nArg == 1 && GetDispatch(&pDispParams->rgvarg[nArg - 1], &pArgs)) {
			nLen = teGetObjectLength(pArgs);
			if (nLen) {
				VARIANT *pvArgs = GetNewVARIANT(nLen);
				for (int i = nLen; i > 0; --i) {
					teGetPropertyAt(pArgs, i - 1, &pvArgs[nLen - i]);
				}
				Invoke4(pdisp, pVarResult, nLen, pvArgs);
			}
			SafeRelease(&pArgs);
		}
		if (nLen == 0) {
			Invoke4(pdisp, pVarResult, -nArg, pDispParams->rgvarg);
		}
		pdisp->Release();
	}
}

VOID teApiGetClipboardData(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{

	if (!::IsClipboardFormatAvailable(CF_UNICODETEXT)) {
		return;
	}
	::OpenClipboard(g_hwndMain);
	HGLOBAL hGlobal = (HGLOBAL)::GetClipboardData(CF_UNICODETEXT);
	if (hGlobal) {
		teSetSZ(pVarResult, (LPWSTR)::GlobalLock(hGlobal));
	}
	::CloseClipboard();
	::GlobalUnlock(hGlobal);
}

VOID teApiSetClipboardData(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	int nSize = ::SysStringByteLen(param[0].bstrVal) + sizeof(WORD);
	if (nSize == sizeof(WORD)) {
		return;
	}
	HGLOBAL hGlobal = ::GlobalAlloc(GHND, nSize);
	if (!hGlobal) {
		return;
	}
	::CopyMemory(::GlobalLock(hGlobal), param[0].bstrVal, nSize);
	::GlobalUnlock(hGlobal);
	if (::OpenClipboard(g_hwndMain) == 0) {
		::GlobalFree(hGlobal);
		return;
	}
	::EmptyClipboard();
	::SetClipboardData(CF_UNICODETEXT, hGlobal);
	CloseClipboard();
}

VOID teApiObjDelete(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		teSetLong(pVarResult, teDelProperty(punk, param[1].lpolestr));
	}
}

VOID teApiGetDeviceCaps(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetDeviceCaps(param[0].hdc, param[1].intVal));
}

VOID teApiSetLayeredWindowAttributes(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetLayeredWindowAttributes(param[0].hwnd, param[1].colorref, param[2].bVal, param[3].dword));
}

VOID teApiSetRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SetRect(param[0].lprect, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal);
	teVariantCopy(pVarResult, &pDispParams->rgvarg[nArg]);
}

VOID teApiSendMessageTimeout(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DWORD_PTR dwResult = 0;
	teSetPtr(pVarResult, SendMessageTimeout(param[0].hwnd, param[1].uintVal, param[2].wparam, param[3].lparam, param[4].uintVal, param[5].uintVal, &dwResult));
	IUnknown *punk = NULL;
	if (nArg >= 6 && FindUnknown(&pDispParams->rgvarg[nArg - 6], &punk)) {
		VARIANT v;
		VariantInit(&v);
		teSetPtr(&v, dwResult);
		tePutPropertyAt(punk, 0, &v);
		VariantClear(&v);
	}
}

#ifdef _DEBUG
VOID teApimciSendString(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, mciSendString(param[0].lpcwstr, param[1].lpwstr, param[2].uintVal, param[3].hwnd));
}

VOID teApiGetGlyphIndices(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	VARIANT vMem;
	VariantInit(&vMem);
	teSetLong(pVarResult, GetGlyphIndices(param[0].hdc, param[1].lpcwstr, lstrlen(param[1].lpcwstr), (LPWORD)GetpcFromVariant(&pDispParams->rgvarg[nArg - 2], &vMem), param[3].dword));
	VariantClear(&vMem);
}

VOID teApiTest(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		CteShellBrowser *pSB;
		if SUCCEEDED(punk->QueryInterface(g_ClsIdSB, (LPVOID *)&pSB)) {
			IShellFolderView *pSFV;
			if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
				LPITEMIDLIST pidl;
				if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg - 1])) {
					LPITEMIDLIST pidl2;
					if (teGetIDListFromVariant(&pidl2, &pDispParams->rgvarg[nArg - 2])) {
						UINT u;
						HRESULT hr = pSFV->UpdateObject(ILFindLastID(pidl), ILFindLastID(pidl2), &u);
						Sleep(hr * 0);
						teILFreeClear(&pidl2);
					}
					teILFreeClear(&pidl);
				}
				pSFV->Release();
			}
			pSFV->Release();
		}
	}
}
#endif
//*/
/*
VOID teApi(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
}

*/

//
TEDispatchApi dispAPI[] = {
	{ 1,  0, -1, -1, "CreateObject", teApiCreateObject },
	{ 1, -1, -1, -1, "Memory", teApiMemory },
	{ 1,  5, -1, -1, "ContextMenu", teApiContextMenu },
	{ 1, -1, -1, -1, "DropTarget", teApiDropTarget },
	{ 0, -1, -1, -1, "DataObject", teApiDataObject },
	{ 5,  1, -1, -1, "OleCmdExec", teApiOleCmdExec },
	{ 1, -1, -1, -1, "sizeof", teApisizeof },
	{ 1, -1, -1, -1, "LowPart", teApiLowPart },
	{ 1, -1, -1, -1, "ULowPart", teApiULowPart },
	{ 1, -1, -1, -1, "HighPart", teApiHighPart },
	{ 1, -1, -1, -1, "QuadPart", teApiQuadPart },
	{ 1, -1, -1, -1, "UQuadPart", teApiUQuadPart },
	{ 1, -1, -1, -1, "pvData", teApipvData },
	{ 3,  1, -1, -1, "ExecMethod", teApiExecMethod },
	{ 4,  2,  3, -1, "Extract", teApiExtract },
	{ 2, -1, -1, -1, "Add", teApiQuadAdd },
	{ 2, -1, -1, -1, "UQuadAdd", teApiUQuadAdd },
	{ 2, -1, -1, -1, "UQuadSub", teApiUQuadSub },
	{ 2, -1, -1, -1, "UQuadCmp", teApiUQuadCmp },
	{ 2,  0,  1, -1, "FindWindow", teApiFindWindow },
	{ 4,  2,  3, -1, "FindWindowEx", teApiFindWindowEx },
	{ 0, -1, -1, -1, "OleGetClipboard", teApiOleGetClipboard },
	{ 1, -1, -1, -1, "OleSetClipboard", teApiOleSetClipboard },
	{ 2,  0,  1, -1, "PathMatchSpec", teApiPathMatchSpec },
	{ 1,  0, -1, -1, "CommandLineToArgv", teApiCommandLineToArgv },
	{ 2, -1, -1, -1, "ILIsEqual", teApiILIsEqual },
	{ 1, -1, -1, -1, "ILClone", teApiILClone },
	{ 3, -1, -1, -1, "ILIsParent", teApiILIsParent },
	{ 1, -1, -1, -1, "ILRemoveLastID", teApiILRemoveLastID },
	{ 1, -1, -1, -1, "ILFindLastID", teApiILFindLastID },
	{ 1, -1, -1, -1, "ILIsEmpty", teApiILIsEmpty },
	{ 1, -1, -1, -1, "ILGetParent", teApiILGetParent },
	{ 2, 10, -1, -1, "FindFirstFile", teApiFindFirstFile },
	{ 1, -1, -1, -1, "WindowFromPoint", teApiWindowFromPoint },
	{ 0, -1, -1, -1, "GetThreadCount", teApiGetThreadCount },
	{ 0, -1, -1, -1, "DoEvents", teApiDoEvents },
	{ 2,  0,  1, -1, "sscanf", teApisscanf },
	{ 4, -1, -1, -1, "SetFileTime", teApiSetFileTime },
	{ 1, -1, -1, -1, "ILCreateFromPath", teApiILCreateFromPath },
	{ 2,  0, -1, -1, "GetProcObject", teApiGetProcObject },
	{ 1,  0, -1, -1, "SetDllDirectory", teApiSetDllDirectory },
	{ 1,  0, -1, -1, "PathIsNetworkPath", teApiPathIsNetworkPath },
	{ 1,  0, -1, -1, "RegisterWindowMessage", teApiRegisterWindowMessage },
	{ 2,  0,  1, -1, "StrCmpI", teApiStrCmpI },
	{ 1,  0, -1, -1, "StrLen", teApiStrLen },
	{ 3,  0,  1, -1, "StrCmpNI", teApiStrCmpNI },
	{ 2,  0,  1, -1, "StrCmpLogical", teApiStrCmpLogical },
	{ 1,  0, -1, -1, "PathQuoteSpaces", teApiPathQuoteSpaces },
	{ 1,  0, -1, -1, "PathUnquoteSpaces", teApiPathUnquoteSpaces },
	{ 1,  0, -1, -1, "GetShortPathName", teApiGetShortPathName },
	{ 1,  0, -1, -1, "PathCreateFromUrl", teApiPathCreateFromUrl },
	{ 1,  0, -1, -1, "UrlCreateFromPath", teApiUrlCreateFromPath },
	{ 1,  0, -1, -1, "PathSearchAndQualify", teApiPathSearchAndQualify },
	{ 3,  0, -1, -1, "PSFormatForDisplay", teApiPSFormatForDisplay },
	{ 1,  0, -1, -1, "PSGetDisplayName", teApiPSGetDisplayName },
	{ 1, -1, -1, -1, "GetCursorPos", teApiGetCursorPos },
	{ 1, -1, -1, -1, "GetKeyboardState", teApiGetKeyboardState },
	{ 1, -1, -1, -1, "SetKeyboardState", teApiSetKeyboardState },
	{ 1, -1, -1, -1, "GetVersionEx", teApiGetVersionEx },
	{ 1, -1, -1, -1, "ChooseFont", teApiChooseFont },
	{ 1, -1, -1, -1, "ChooseColor", teApiChooseColor },
	{ 1, -1, -1, -1, "TranslateMessage", teApiTranslateMessage },
	{ 1, -1, -1, -1, "ShellExecuteEx", teApiShellExecuteEx },
	{ 1, -1, -1, -1, "CreateFontIndirect", teApiCreateFontIndirect },
	{ 1, -1, -1, -1, "CreateIconIndirect", teApiCreateIconIndirect },
	{ 1, -1, -1, -1, "FindText", teApiFindText },
	{ 1, -1, -1, -1, "FindReplace", teApiFindReplace },
	{ 1, -1, -1, -1, "DispatchMessage", teApiDispatchMessage },
	{ 1, -1, -1, -1, "PostMessage", teApiPostMessage },
	{ 1, -1, -1, -1, "SHFreeNameMappings", teApiSHFreeNameMappings },
	{ 1, -1, -1, -1, "CoTaskMemFree", teApiCoTaskMemFree },
	{ 1, -1, -1, -1, "Sleep", teApiSleep },
	{ 6,  2,  3,  4, "ShRunDialog", teApiShRunDialog },
	{ 2, -1, -1, -1, "DragAcceptFiles", teApiDragAcceptFiles },
	{ 1, -1, -1, -1, "DragFinish", teApiDragFinish },
	{ 2, -1, -1, -1, "mouse_event", teApimouse_event },
	{ 4, -1, -1, -1, "keybd_event", teApikeybd_event },
	{ 4, -1, -1, -1, "SHChangeNotify", teApiSHChangeNotify },
	{ 4, -1, -1, -1, "DrawIcon", teApiDrawIcon },
	{ 1, -1, -1, -1, "DestroyMenu", teApiDestroyMenu },
	{ 1, -1, -1, -1, "FindClose", teApiFindClose },
	{ 1, -1, -1, -1, "FreeLibrary", teApiFreeLibrary },
	{ 1, -1, -1, -1, "ImageList_Destroy", teApiImageList_Destroy },
	{ 1, -1, -1, -1, "DeleteObject", teApiDeleteObject },
	{ 1, -1, -1, -1, "DestroyIcon", teApiDestroyIcon },
	{ 1, -1, -1, -1, "IsWindow", teApiIsWindow },
	{ 1, -1, -1, -1, "IsWindowVisible", teApiIsWindowVisible },
	{ 1, -1, -1, -1, "IsZoomed", teApiIsZoomed },
	{ 1, -1, -1, -1, "IsIconic", teApiIsIconic },
	{ 1, -1, -1, -1, "OpenIcon", teApiOpenIcon },
	{ 1, -1, -1, -1, "SetForegroundWindow", teApiSetForegroundWindow },
	{ 1, -1, -1, -1, "BringWindowToTop", teApiBringWindowToTop },
	{ 1, -1, -1, -1, "DeleteDC", teApiDeleteDC },
	{ 1, -1, -1, -1, "CloseHandle", teApiCloseHandle },
	{ 1, -1, -1, -1, "IsMenu", teApiIsMenu },
	{ 6, -1, -1, -1, "MoveWindow", teApiMoveWindow },
	{ 5, -1, -1, -1, "SetMenuItemBitmaps", teApiSetMenuItemBitmaps },
	{ 2, -1, -1, -1, "ShowWindow", teApiShowWindow },
	{ 3, -1, -1, -1, "DeleteMenu", teApiDeleteMenu },
	{ 3, -1, -1, -1, "RemoveMenu", teApiRemoveMenu },
	{ 9, -1, -1, -1, "DrawIconEx", teApiDrawIconEx },
	{ 3, -1, -1, -1, "EnableMenuItem", teApiEnableMenuItem },
	{ 2, -1, -1, -1, "ImageList_Remove", teApiImageList_Remove },
	{ 3, -1, -1, -1, "ImageList_SetIconSize", teApiImageList_SetIconSize },
	{ 6, -1, -1, -1, "ImageList_Draw", teApiImageList_Draw },
	{10, -1, -1, -1, "ImageList_DrawEx", teApiImageList_DrawEx },
	{ 2, -1, -1, -1, "ImageList_SetImageCount", teApiImageList_SetImageCount },
	{ 3, -1, -1, -1, "ImageList_SetOverlayImage", teApiImageList_SetOverlayImage },
	{ 5, -1, -1, -1, "ImageList_Copy", teApiImageList_Copy },
	{ 1, -1, -1, -1, "DestroyWindow", teApiDestroyWindow },
	{ 3, -1, -1, -1, "LineTo", teApiLineTo },
	{ 0, -1, -1, -1, "ReleaseCapture", teApiReleaseCapture },
	{ 2, -1, -1, -1, "SetCursorPos", teApiSetCursorPos },
	{ 2, -1, -1, -1, "DestroyCursor", teApiDestroyCursor },
	{ 2, -1, -1, -1, "SHFreeShared", teApiSHFreeShared },
	{ 0, -1, -1, -1, "EndMenu", teApiEndMenu },
	{ 1, -1, -1, -1, "SHChangeNotifyDeregister", teApiSHChangeNotifyDeregister },
	{ 1, -1, -1, -1, "SHChangeNotification_Unlock", teApiSHChangeNotification_Unlock },
	{ 1, -1, -1, -1, "IsWow64Process", teApiIsWow64Process },
	{ 9, -1, -1, -1, "BitBlt", teApiBitBlt },
	{ 2, -1, -1, -1, "ImmSetOpenStatus", teApiImmSetOpenStatus },
	{ 2, -1, -1, -1, "ImmReleaseContext", teApiImmReleaseContext },
	{ 2, -1, -1, -1, "IsChild", teApiIsChild },
	{ 2, -1, -1, -1, "KillTimer", teApiKillTimer },
	{ 1, -1, -1, -1, "AllowSetForegroundWindow", teApiAllowSetForegroundWindow },
	{ 7, -1, -1, -1, "SetWindowPos", teApiSetWindowPos },
	{ 5,  4, -1, -1, "InsertMenu", teApiInsertMenu },
	{ 2, -1,  1, -1, "SetWindowText", teApiSetWindowText },
	{ 4, -1, -1, -1, "RedrawWindow", teApiRedrawWindow },
	{ 4, -1, -1, -1, "MoveToEx", teApiMoveToEx },
	{ 3, -1, -1, -1, "InvalidateRect", teApiInvalidateRect },
	{ 4, -1, -1, -1, "SendNotifyMessage", teApiSendNotifyMessage },
	{ 5, -1, -1, -1, "PeekMessage", teApiPeekMessage },
	{ 1, -1, -1, -1, "SHFileOperation", teApiSHFileOperation },
	{ 4, -1, -1, -1, "GetMessage", teApiGetMessage },
	{ 2, -1, -1, -1, "GetWindowRect", teApiGetWindowRect },
	{ 2, -1, -1, -1, "GetWindowThreadProcessId", teApiGetWindowThreadProcessId },
	{ 2, -1, -1, -1, "GetClientRect", teApiGetClientRect },
	{ 3, -1, -1, -1, "SendInput", teApiSendInput },
	{ 2, -1, -1, -1, "ScreenToClient", teApiScreenToClient },
	{ 5, -1, -1, -1, "MsgWaitForMultipleObjectsEx", teApiMsgWaitForMultipleObjectsEx },
	{ 2, -1, -1, -1, "ClientToScreen", teApiClientToScreen },
	{ 2, -1, -1, -1, "GetIconInfo", teApiGetIconInfo },
	{ 2, -1, -1, -1, "FindNextFile", teApiFindNextFile },
	{ 3, -1, -1, -1, "FillRect", teApiFillRect },
	{ 2, -1, -1, -1, "Shell_NotifyIcon", teApiShell_NotifyIcon },
	{ 2, -1, -1, -1, "EndPaint", teApiEndPaint },
	{ 2, -1, -1, -1, "ImageList_GetIconSize", teApiImageList_GetIconSize },
	{ 2, -1, -1, -1, "GetMenuInfo", teApiGetMenuInfo },
	{ 2, -1, -1, -1, "SetMenuInfo", teApiSetMenuInfo },
	{ 4, -1, -1, -1, "SystemParametersInfo", teApiSystemParametersInfo },
	{ 3,  1, -1, -1, "GetTextExtentPoint32", teApiGetTextExtentPoint32 },
	{ 4, -1, -1, -1, "SHGetDataFromIDList", teApiSHGetDataFromIDList },
	{ 4, -1, -1, -1, "InsertMenuItem", teApiInsertMenuItem },
	{ 4, -1, -1, -1, "GetMenuItemInfo", teApiGetMenuItemInfo },
	{ 4, -1, -1, -1, "SetMenuItemInfo", teApiSetMenuItemInfo },
	{ 4, -1, -1, -1, "ChangeWindowMessageFilterEx", teApiChangeWindowMessageFilterEx },
	{ 2, -1, -1, -1, "ImageList_SetBkColor", teApiImageList_SetBkColor },
	{ 2, -1, -1, -1, "ImageList_AddIcon", teApiImageList_AddIcon },
	{ 3, -1, -1, -1, "ImageList_Add", teApiImageList_Add },
	{ 3, -1, -1, -1, "ImageList_AddMasked", teApiImageList_AddMasked },
	{ 1, -1, -1, -1, "GetKeyState", teApiGetKeyState },
	{ 1, -1, -1, -1, "GetSystemMetrics", teApiGetSystemMetrics },
	{ 1, -1, -1, -1, "GetSysColor", teApiGetSysColor },
	{ 1, -1, -1, -1, "GetMenuItemCount", teApiGetMenuItemCount },
	{ 1, -1, -1, -1, "ImageList_GetImageCount", teApiImageList_GetImageCount },
	{ 2, -1, -1, -1, "ReleaseDC", teApiReleaseDC },
	{ 2, -1, -1, -1, "GetMenuItemID", teApiGetMenuItemID },
	{ 4, -1, -1, -1, "ImageList_Replace", teApiImageList_Replace },
	{ 3, -1, -1, -1, "ImageList_ReplaceIcon", teApiImageList_ReplaceIcon },
	{ 2, -1, -1, -1, "SetBkMode", teApiSetBkMode },
	{ 2, -1, -1, -1, "SetBkColor", teApiSetBkColor },
	{ 2, -1, -1, -1, "SetTextColor", teApiSetTextColor },
	{ 2, -1, -1, -1, "MapVirtualKey", teApiMapVirtualKey },
	{ 2, -1, -1, -1, "WaitForInputIdle", teApiWaitForInputIdle },
	{ 2, -1, -1, -1, "WaitForSingleObject", teApiWaitForSingleObject },
	{ 3, -1, -1, -1, "GetMenuDefaultItem", teApiGetMenuDefaultItem },
	{ 3, -1, -1, -1, "SetMenuDefaultItem", teApiSetMenuDefaultItem },
	{ 1, -1, -1, -1, "CRC32", teApiCRC32 },
	{ 3,  1, -1, -1, "SHEmptyRecycleBin", teApiSHEmptyRecycleBin },
	{ 0, -1, -1, -1, "GetMessagePos", teApiGetMessagePos },
	{ 2, -1, -1, -1, "ImageList_GetOverlayImage", teApiImageList_GetOverlayImage },
	{ 1, -1, -1, -1, "ImageList_GetBkColor", teApiImageList_GetBkColor },
	{ 3,  1,  2, -1, "SetWindowTheme", teApiSetWindowTheme },
	{ 1, -1, -1, -1, "ImmGetVirtualKey", teApiImmGetVirtualKey },
	{ 1, -1, -1, -1, "GetAsyncKeyState", teApiGetAsyncKeyState },
	{ 6, -1, -1, -1, "TrackPopupMenuEx", teApiTrackPopupMenuEx },
	{ 5,  0, -1, -1, "ExtractIconEx", teApiExtractIconEx },
	{ 3, -1, -1, -1, "GetObject", teApiGetObject },
	{ 2, -1, -1, -1, "MultiByteToWideChar", teApiMultiByteToWideChar },
	{ 2, -1, -1, -1, "WideCharToMultiByte", teApiWideCharToMultiByte },
	{ 2, -1, -1, -1, "GetAttributesOf", teApiGetAttributesOf },
	{ 2, -1, -1, -1, "DoDragDrop", teApiDoDragDrop },//Deprecated
	{ 4, -1, -1, -1, "SHDoDragDrop", teApiSHDoDragDrop },
	{ 3, -1, -1, -1, "CompareIDs", teApiCompareIDs },
	{ 2,  0,  1, -1, "ExecScript", teApiExecScript },
	{ 2,  0,  1, -1, "GetScriptDispatch", teApiGetScriptDispatch },
	{ 2,  1, -1, -1, "GetDispatch", teApiGetDispatch },
	{ 6, -1, -1, -1, "SHChangeNotifyRegister", teApiSHChangeNotifyRegister },
	{ 4,  1,  2, -1, "MessageBox", teApiMessageBox },
	{ 3, -1, -1, -1, "ImageList_GetIcon", teApiImageList_GetIcon },
	{ 5, -1, -1, -1, "ImageList_Create", teApiImageList_Create },
	{ 2, -1, -1, -1, "GetWindowLongPtr", teApiGetWindowLongPtr },
	{ 2, -1, -1, -1, "GetClassLongPtr", teApiGetClassLongPtr },
	{ 2, -1, -1, -1, "GetSubMenu", teApiGetSubMenu },
	{ 2, -1, -1, -1, "SelectObject", teApiSelectObject },
	{ 1, -1, -1, -1, "GetStockObject", teApiGetStockObject },
	{ 1, -1, -1, -1, "GetSysColorBrush", teApiGetSysColorBrush },
	{ 1, -1, -1, -1, "SetFocus", teApiSetFocus },
	{ 1, -1, -1, -1, "GetDC", teApiGetDC },
	{ 1, -1, -1, -1, "CreateCompatibleDC", teApiCreateCompatibleDC },
	{ 0, -1, -1, -1, "CreatePopupMenu", teApiCreatePopupMenu },
	{ 0, -1, -1, -1, "CreateMenu", teApiCreateMenu },
	{ 3, -1, -1, -1, "CreateCompatibleBitmap", teApiCreateCompatibleBitmap },
	{ 3, -1, -1, -1, "SetWindowLongPtr", teApiSetWindowLongPtr },
	{ 3, -1, -1, -1, "SetClassLongPtr", teApiSetClassLongPtr },
	{ 1, -1, -1, -1, "ImageList_Duplicate", teApiImageList_Duplicate },
	{ 2, -1, -1, -1, "SendMessage", teApiSendMessage },
	{ 2, -1, -1, -1, "GetSystemMenu", teApiGetSystemMenu },
	{ 1, -1, -1, -1, "GetWindowDC", teApiGetWindowDC },
	{ 3, -1, -1, -1, "CreatePen", teApiCreatePen },
	{ 1, -1, -1, -1, "SetCapture", teApiSetCapture },
	{ 1, -1, -1, -1, "SetCursor", teApiSetCursor },
	{ 5, -1, -1, -1, "CallWindowProc", teApiCallWindowProc },
	{ 1, -1, -1, -1, "GetWindow", teApiGetWindow },
	{ 1, -1, -1, -1, "GetTopWindow", teApiGetTopWindow },
	{ 3, -1, -1, -1, "OpenProcess", teApiOpenProcess },
	{ 1, -1, -1, -1, "GetParent", teApiGetParent },
	{ 0, -1, -1, -1, "GetCapture", teApiGetCapture },
	{ 1, -1,  0, -1, "GetModuleHandle", teApiGetModuleHandle },
	{ 1, -1, -1, -1, "SHGetImageList", teApiSHGetImageList },
	{ 5, -1, -1, -1, "CopyImage", teApiCopyImage },
	{ 0, -1, -1, -1, "GetCurrentProcess", teApiGetCurrentProcess },
	{ 1, -1, -1, -1, "ImmGetContext", teApiImmGetContext },
	{ 0, -1, -1, -1, "GetFocus", teApiGetFocus },
	{ 0, -1, -1, -1, "GetForegroundWindow", teApiGetForegroundWindow },
	{ 3, -1, -1, -1, "SetTimer", teApiSetTimer },
	{ 2, -1, -1, -1, "LoadMenu", teApiLoadMenu },
	{ 2, -1, -1, -1, "LoadIcon", teApiLoadIcon },
	{ 3, -1, -1, -1, "LoadLibraryEx", teApiLoadLibraryEx },
	{ 6, -1, -1, -1, "LoadImage", teApiLoadImage },
	{ 7, -1, -1, -1, "ImageList_LoadImage", teApiImageList_LoadImage },
	{ 5, 10, -1, -1, "SHGetFileInfo", teApiSHGetFileInfo },
	{12, -1,  1,  2, "CreateWindowEx", teApiCreateWindowEx },
	{ 6, -1,  3,  4, "ShellExecute", teApiShellExecute },
	{ 2, -1, -1, -1, "BeginPaint", teApiBeginPaint },
	{ 2, -1, -1, -1, "LoadCursor", teApiLoadCursor },
	{ 2, -1, -1, -1, "LoadCursorFromFile", teApiLoadCursorFromFile },
	{ 3, -1, -1, -1, "SHParseDisplayName", teApiSHParseDisplayName },
	{ 3, -1, -1, -1, "SHChangeNotification_Lock", teApiSHChangeNotification_Lock },
	{ 1, -1, -1, -1, "GetWindowText", teApiGetWindowText },
	{ 1, -1, -1, -1, "GetClassName", teApiGetClassName },
	{ 1, -1, -1, -1, "GetModuleFileName", teApiGetModuleFileName },
	{ 0, -1, -1, -1, "GetCommandLine", teApiGetCommandLine },
	{ 0, -1, -1, -1, "GetCurrentDirectory", teApiGetCurrentDirectory },//use wsh.CurrentDirectory
	{ 3, -1, -1, -1, "GetMenuString", teApiGetMenuString },
	{ 2, -1, -1, -1, "GetDisplayNameOf", teApiGetDisplayNameOf },
	{ 1, -1, -1, -1, "GetKeyNameText", teApiGetKeyNameText },
	{ 2, -1, -1, -1, "DragQueryFile", teApiDragQueryFile },
	{ 1, -1, -1, -1, "SysAllocString", teApiSysAllocString },
	{ 2, -1, -1, -1, "SysAllocStringLen", teApiSysAllocStringLen },
	{ 2, -1, -1, -1, "SysAllocStringByteLen", teApiSysAllocStringByteLen },
	{ 3, -1,  1, -1, "sprintf", teApisprintf },
	{ 1, -1, -1, -1, "base64_encode", teApibase64_encode },
	{ 2, -1, -1, -1, "LoadString", teApiLoadString },
	{ 4,  2,  3, -1, "AssocQueryString", teApiAssocQueryString },
	{ 1, -1, -1, -1, "StrFormatByteSize", teApiStrFormatByteSize },
	{ 4,  3, -1, -1, "GetDateFormat", teApiGetDateFormat },
	{ 4,  3, -1, -1, "GetTimeFormat", teApiGetTimeFormat },
	{ 2, -1, -1, -1, "GetLocaleInfo", teApiGetLocaleInfo },
	{ 2, -1, -1, -1, "MonitorFromPoint", teApiMonitorFromPoint },
	{ 2, -1, -1, -1, "MonitorFromRect", teApiMonitorFromRect },
	{ 2, -1, -1, -1, "GetMonitorInfo", teApiGetMonitorInfo },
	{ 1,  0, -1, -1, "GlobalAddAtom", teApiGlobalAddAtom },
	{ 1, -1, -1, -1, "GlobalGetAtomName", teApiGlobalGetAtomName },
	{ 1,  0, -1, -1, "GlobalFindAtom", teApiGlobalFindAtom },
	{ 1, -1, -1, -1, "GlobalDeleteAtom", teApiGlobalDeleteAtom },
	{ 4, -1, -1, -1, "RegisterHotKey", teApiRegisterHotKey },
	{ 2, -1, -1, -1, "UnregisterHotKey", teApiUnregisterHotKey },
	{ 1, -1, -1, -1, "ILGetCount", teApiILGetCount },
	{ 2, -1, -1, -1, "SHTestTokenMembership", teApiSHTestTokenMembership },
	{ 2,  1, -1, -1, "ObjGetI", teApiObjGetI },
	{ 3,  1, -1, -1, "ObjPutI", teApiObjPutI },
	{ 5,  1, -1, -1, "DrawText", teApiDrawText },
	{ 5, -1, -1, -1, "Rectangle", teApiRectangle },
	{ 1, 10, -1, -1, "PathIsDirectory", teApiPathIsDirectory },
	{ 2,  0,  1, -1, "PathIsSameRoot", teApiPathIsSameRoot },
	{ 2,  1, -1, -1, "ObjDispId", teApiObjDispId },
	{ 1,  0, -1, -1, "SHSimpleIDListFromPath", teApiSHSimpleIDListFromPath },
	{ 1,  0, -1, -1, "OutputDebugString", teApiOutputDebugString },
	{ 2, -1, -1, -1, "DllGetClassObject", teApiDllGetClassObject },
	{ 0, -1, -1, -1, "IsAppThemed", teApiIsAppThemed },
	{ 6,  0, -1, -1, "SHDefExtractIcon", teApiSHDefExtractIcon },
	{ 3,  1,  2, -1, "URLDownloadToFile", teApiURLDownloadToFile },
	{ 3,  0,  1, -1, "MoveFileEx", teApiMoveFileEx },
	{ 0, -1, -1, -1, "Object", teApiObject },
	{ 0, -1, -1, -1, "Array", teApiArray },
	{ 1,  0, -1, -1, "PathIsRoot", teApiPathIsRoot },
	{ 2, -1, -1, -1, "EnableWindow", teApiEnableWindow },
	{ 1, -1, -1, -1, "IsWindowEnabled", teApiIsWindowEnabled },
	{ 2,  0,  1, -1, "CryptProtectData", teApiCryptProtectData },
	{ 2, -1,  1, -1, "CryptUnprotectData", teApiCryptUnprotectData },
	{ 1,  0, -1, -1, "base64_decode", teApibase64_decode },
	{11, -1, -1, -1, "AlphaBlend", teApiAlphaBlend },
	{11, -1, -1, -1, "StretchBlt", teApiStretchBlt },
	{11, -1, -1, -1, "TransparentBlt", teApiTransparentBlt },
	{ 2,  0,  1,  2, "FormatMessage", teApiFormatMessage },
	{ 2, -1, -1, -1, "GetGUIThreadInfo", teApiGetGUIThreadInfo },
	{ 1, -1, -1, -1, "CreateSolidBrush", teApiCreateSolidBrush },
	{10, -1, -1, -1, "PlgBlt", teApiPlgBlt },
	{ 7, -1, -1, -1, "RoundRect", teApiRoundRect },
	{ 1,  0,  1, -1, "CreateProcess", teApiCreateProcess },
	{ 1,  0, -1, -1, "DeleteFile", teApiDeleteFile },
	{ 4,  3, -1, -1, "CreateEvent", teApiCreateEvent },
	{ 1, -1, -1, -1, "SetEvent", teApiSetEvent },
	{ 1, -1, -1, -1, "ResetEvent", teApiResetEvent },
	{ 3,  0, -1, -1, "PlaySound", teApiPlaySound },
	{ 4,  0, -1, -1, "SHCreateStreamOnFileEx", teApiSHCreateStreamOnFileEx },
	{ 1, -1, -1, -1, "HasThumbnail", teApiHasThumbnail },
	{ 1,  0, -1, -1, "GetDiskFreeSpaceEx", teApiGetDiskFreeSpaceEx },
	{ 3, -1, -1, -1, "SetSysColors", teApiSetSysColors },
	{ 1, -1, -1, -1, "IsClipboardFormatAvailable", teApiIsClipboardFormatAvailable },
	{ 7,  0, -1, -1, "CreateFile", teApiCreateFile },
	{ 3, -1, -1, -1, "GetPixel", teApiGetPixel },
	{ 0, -1, -1, -1, "ShouldAppsUseDarkMode", teApiShouldAppsUseDarkMode },
	{ 2, -1, -1, -1, "GetProcAddress", teApiGetProcAddress },
	{ 6,  4, -1, -1, "RunDLL", teApiRunDLL },
	{ 0, -1, -1, -1, "GetCurrentThreadId", teApiGetCurrentThreadId },
	{ 0, -1, -1, -1, "GetDpiForMonitor", teApiGetDpiForMonitor },
	{ 2, -1, -1, -1, "HashData", teApiHashData },
	{ 2, -1, -1, -1, "SetDCPenColor", teApiSetDCPenColor },
	{ 2, -1, -1, -1, "SetDCBrushColor", teApiSetDCBrushColor },
	{ 2, -1, -1, -1, "WriteFile", teApiWriteFile },
	{ 2, -1, -1, -1, "ReadFile", teApiReadFile },
	{ 3, -1, -1, -1, "SetFilePointer", teApiSetFilePointer },
	{ 1, -1, -1, -1, "SysStringByteLen", teApiSysStringByteLen },
	{ 2,  1, -1, -1, "GetProp", teApiGetProp },
	{ 1, -1, -1, -1, "Invoke", teApiInvoke },
	{ 0, -1, -1, -1, "GetClipboardData", teApiGetClipboardData },
	{ 1,  0, -1, -1, "SetClipboardData", teApiSetClipboardData },
	{ 2,  1, -1, -1, "ObjDelete", teApiObjDelete },
	{ 2, -1, -1, -1, "GetDeviceCaps", teApiGetDeviceCaps },
	{ 4, -1, -1, -1, "SetLayeredWindowAttributes", teApiSetLayeredWindowAttributes },
	{ 1, -1, -1, -1, "SetRect", teApiSetRect },
	{ 6, -1, -1, -1, "SendMessageTimeout", teApiSendMessageTimeout },
#ifdef _DEBUG
	{ 4,  0, -1, -1, "mciSendString", teApimciSendString },
	{ 4,  1, -1, -1, "GetGlyphIndices", teApiGetGlyphIndices },
	{ 2, -1, -1, -1, "Test", teApiTest },
#endif
};

VOID Initlize()
{
	g_maps[MAP_TE] = SortTEMethod(methodTE, _countof(methodTE));
	g_maps[MAP_SB] = SortTEMethod(methodSB, _countof(methodSB));
	g_maps[MAP_TC] = SortTEMethod(methodTC, _countof(methodTC));
	g_maps[MAP_TV] = SortTEMethod(methodTV, _countof(methodTV));
	g_maps[MAP_Mem] = SortTEMethod(methodMem, _countof(methodMem));
	g_maps[MAP_GB] = SortTEMethod(methodGB, _countof(methodGB));
	g_maps[MAP_FIs] = SortTEMethod(methodFIs, _countof(methodFIs));
	g_maps[MAP_CD] = SortTEMethod(methodCD, _countof(methodCD));
	g_maps[MAP_SS] = SortTEStruct(pTEStructs, _countof(pTEStructs));
/*/// For order check
	int *map = g_maps[MAP_SS];
	for (int i = 0; i < _countof(pTEStructs); ++i) {
		int j = map[i];
		if (i != j) {
			LPSTR lpi = pTEStructs[i].name;
			LPSTR lpj = pTEStructs[j].name;
			Sleep(0);
		}
	}
*///
#ifdef _USE_BSEARCHAPI
	g_maps[MAP_API] = teSortDispatchApi(dispAPI, _countof(dispAPI));
#endif
	IExplorerBrowser *pEB;
	if (SUCCEEDED(teCreateInstance(CLSID_ExplorerBrowser, NULL, NULL, IID_PPV_ARGS(&pEB)))) {
		IFolderViewOptions *pOptions;
		if SUCCEEDED(pEB->QueryInterface(IID_PPV_ARGS(&pOptions))) {
			g_bCanLayout = TRUE;
			pOptions->Release();
		}
		pEB->Release();
	}
}

//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//	  This function and its usage are only necessary if you want this code
//    to be compatible with Win32 systems prior to the 'RegisterClassEx'
//    function that was added to Windows 95. It is important to call this function
//    so that the application will get 'well formed' small icons associated
//    with it.
//
ATOM MyRegisterClass(HINSTANCE hInstance, LPWSTR szClassName, WNDPROC lpfnWndProc)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);
	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= lpfnWndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_TE));
	wcex.hIconSm		= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_TE));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground  = (HBRUSH)(COLOR_BTNFACE + 1);
	wcex.lpszMenuName	= 0;
	wcex.lpszClassName	= szClassName;

	return RegisterClassEx(&wcex);
}

BOOL teFindType(LPWSTR pszText, LPWSTR pszFind)
{
	if (!pszFind) {
		return FALSE;
	}
	LPWSTR pszExt1 = pszText;
	LPWSTR pszExt2;
	if (pszFind[0]) {
		while (pszExt2 = StrChr(pszExt1, ',')) {
			if (StrCmpNI(pszFind, pszExt1, (int)(pszExt2 - pszExt1)) == 0) {
				return TRUE;
			}
			pszExt1 = pszExt2 + 1;
		}
		return lstrcmpi(pszText, pszFind) == 0;
	}
	return FALSE;
}

BOOL teIsHighContrast()
{
	HIGHCONTRAST highContrast = { sizeof(HIGHCONTRAST) };
	return SystemParametersInfo(SPI_GETHIGHCONTRAST, sizeof(highContrast), &highContrast, FALSE) && (highContrast.dwFlags & HCF_HIGHCONTRASTON);
}

VOID teGetDarkMode()
{
	if (_ShouldAppsUseDarkMode && _AllowDarkModeForWindow && _SetPreferredAppMode) {
		g_bDarkMode = _ShouldAppsUseDarkMode() && IsAppThemed() && !teIsHighContrast();
		if (teVerifyVersion(10, 0, 18334)) {
			_SetPreferredAppMode(g_bDarkMode ? APPMODE_FORCEDARK : APPMODE_FORCELIGHT);
		} else {
			((LPFNAllowDarkModeForApp)_SetPreferredAppMode)(g_bDarkMode);
		}
		if (_RefreshImmersiveColorPolicyState) {
			_RefreshImmersiveColorPolicyState();
		}
		teSetDarkMode(g_hwndMain);
	}
}

BOOL CanClose(PVOID pObj)
{
	return DoFunc(TE_OnClose, pObj, S_OK) != S_FALSE;
}

VOID teSetObjectRects(IUnknown *punk, HWND hwnd)
{
	IOleInPlaceObject *pOleInPlaceObject;
	punk->QueryInterface(IID_PPV_ARGS(&pOleInPlaceObject));
	RECT rc;
	GetClientRect(hwnd, &rc);
	pOleInPlaceObject->SetObjectRects(&rc, &rc);
	pOleInPlaceObject->Release();
	if (g_nBlink == 1) {
		IWebBrowser2 *pWB2;
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pWB2))) {
			POINT pt = { 0, 0 };
			ClientToScreen(hwnd, &pt);
			pWB2->put_Left(pt.x);
			pWB2->put_Top(pt.y);
			pWB2->Release();
		}
	}
}

VOID CALLBACK teTimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	CteShellBrowser *pSB;
	KillTimer(hwnd, idEvent);
	try {
		switch (idEvent) {
			case TET_Create:
				if (g_pOnFunc[TE_OnCreate]) {
					DoFunc(TE_OnCreate, g_pTE, E_NOTIMPL);
					DoFunc(TE_OnCreate, g_pWebBrowser, E_NOTIMPL);
					break;
				}
				g_nCreateTimer *= 2;
				if (g_nCreateTimer < 10000) {
					SetTimer(g_hwndMain, TET_Create, g_nCreateTimer, teTimerProc);
					break;
				}
				g_bsDocumentWrite = ::SysAllocStringLen(NULL, MAX_STATUS);
				MultiByteToWideChar(CP_UTF8, 0, lstrcmpi(g_pWebBrowser->m_bstrPath, L"TE32.exe") ? PathFileExists(g_pWebBrowser->m_bstrPath) ?
					"<h1>500 Internal Script Error</h1>" : "<h1>404 File Not Found</h1>" : "<h1>303 Exec Other</h1>",
					-1, g_bsDocumentWrite, MAX_STATUS);
				lstrcat(g_bsDocumentWrite, g_pWebBrowser->m_bstrPath);
				g_nLockUpdate = 0;
				::SysReAllocString(&g_pWebBrowser->m_bstrPath, PATH_BLANK);
				g_pWebBrowser->m_pWebBrowser->Navigate(g_pWebBrowser->m_bstrPath, NULL, NULL, NULL, NULL);
				break;
			case TET_Reload:
				if (g_nReload) {
					teRevoke(g_pSW);
#ifndef _DEBUG
					BSTR bsPath, bsQuoted;
					int i = teGetModuleFileName(NULL, &bsPath);
					bsQuoted = teSysAllocStringLen(bsPath, i + 2);
					SysFreeString(bsPath);
					PathQuoteSpaces(bsQuoted);
					ShellExecute(hwnd, NULL, bsQuoted, NULL, NULL, SW_SHOWNOACTIVATE);
					SysFreeString(bsQuoted);
#endif
					PostMessage(hwnd, WM_CLOSE, 0, 0);
				}
				break;
			case TET_Refresh:
				if (g_pWebBrowser) {
					g_nLockUpdate = 0;
					if (g_pWebBrowser->IsBusy()) {
						SetTimer(g_hwndMain, TET_Refresh, 100, teTimerProc);
						return;
					}
					g_pWebBrowser->m_pWebBrowser->Refresh();
				}
				break;
			case TET_Size:
				ArrangeWindowEx();
				break;
			case TET_FreeLibrary:
				while (!g_pFreeLibrary.empty()) {
					HMODULE hDll = g_pFreeLibrary.back();
					g_pFreeLibrary.pop_back();
					teFreeLibrary2(hDll);
					if (g_dwFreeLibrary) {
						break;
					}
				}
				break;
			case TET_Status:
				if (!g_pTC) {
					return;
				}
				pSB = g_pTC->GetShellBrowser(g_pTC->m_nIndex);
				if (pSB) {
					if (pSB->m_bBeforeNavigate) {
						SetTimer(g_hwndMain, TET_Status, 1000, teTimerProc);
						break;
					}
					if (pSB->m_bRedraw) {
						pSB->m_bRedraw = FALSE;
						pSB->SetRedraw(TRUE);
						RedrawWindow(pSB->m_hwnd, NULL, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
					}
					g_szStatus[0] = NULL;
					int nCount = pSB->GetFolderViewAndItemCount(NULL, SVGIO_SELECTION);
					UINT uID = 0;
					if (nCount) {
						uID = nCount > 1 ? 38194 : 38195;
					} else {
						nCount = pSB->GetFolderViewAndItemCount(NULL, SVGIO_ALLVIEW);
						uID = nCount > 1 ? 38192 : 38193;
					}
					if (uID) {
						BSTR bsCount = ::SysAllocStringLen(NULL, 16);
						WCHAR pszNum[16];
						swprintf_s(pszNum, 12, L"%d", nCount);
						teCommaSize(pszNum, bsCount, 16, 0);

						WCHAR psz[MAX_STATUS];
						if (LoadString(g_hShell32, uID, psz, _countof(psz)) > 2 && tePathMatchSpec(psz, L"*%s*")) {
							swprintf_s(g_szStatus, _countof(g_szStatus), psz, bsCount);
						} else {
							uID = uID < 38194 ? 6466 : 6477;
							if (LoadString(g_hShell32, uID, psz, MAX_STATUS) > 6) {
								FormatMessage(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ARGUMENT_ARRAY,
									psz, 0, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), g_szStatus, _countof(g_szStatus), (va_list *)&bsCount);
							}
						}
						::SysFreeString(bsCount);
					}
					DoStatusText(pSB, g_szStatus, 0);
				}
				break;
			case TET_Title:
				SetWindowText(hwnd, g_bsTitle);
				break;
			case TET_EndThread:
				DoFunc(TE_OnEndThread, g_pTE, S_OK);
				break;
		}//end_switch
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teTimerProc";
#endif
	}
}

VOID CALLBACK teTimerProcSW(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	try {
		KillTimer(hwnd, idEvent);
		CteWebBrowser *pWB = (CteWebBrowser *)idEvent;
		if (GetFocus() == pWB->m_hwndParent) {
			IHTMLWindow2 *pwin;
			if (teGetHTMLWindow(pWB->m_pWebBrowser, IID_PPV_ARGS(&pwin))) {
				pwin->focus();
				pwin->Release();
			}
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teTimerProcSW";
#endif
	}
}

VOID CALLBACK teTimerProcGroupBy(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	try {
		KillTimer(hwnd, idEvent);
		CteShellBrowser *pSB = SBfromhwnd(hwnd);
		if (pSB && pSB->m_bsNextGroup && idEvent == (UINT_PTR)&pSB->m_bsNextGroup) {
			--pSB->m_nGroupByDelay;
			pSB->SetGroupBy(pSB->m_bsNextGroup);
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teTimerProcGroupBy";
#endif
	}
}

VOID CALLBACK teTimerProcSW2(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	try {
		KillTimer(hwnd, idEvent);
		if (hwnd == (HWND)idEvent) {
			teSetExStyleAnd(hwnd, ~WS_EX_LAYERED);
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teTimerProcSW2";
#endif
	}
}

LRESULT CALLBACK HookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LPCWPSTRUCT pcwp;
	try {
		if (nCode >= 0) {
			if (nCode == HC_ACTION) {
				if (wParam == NULL) {
					pcwp = (LPCWPSTRUCT)lParam;
					switch (pcwp->message)
					{
						case WM_SETFOCUS:
							CheckChangeTabSB(pcwp->hwnd);
						case WM_KILLFOCUS:
							if (g_pWebBrowser && !IsChild(g_pWebBrowser->get_HWND(), pcwp->hwnd)) {
								break;
							}
						case WM_SHOWWINDOW:
						case WM_ACTIVATE:
						case WM_ACTIVATEAPP:
						case WM_INITDIALOG:
							if (g_pOnFunc[TE_OnSystemMessage]) {
								IDispatch *pdisp;
								if (ControlFromhwnd(&pdisp, pcwp->hwnd) == S_OK) {
									MSG msg1;
									msg1.hwnd = pcwp->hwnd;
									msg1.message = pcwp->message;
									msg1.wParam = pcwp->wParam;
									msg1.lParam = pcwp->lParam;
									MessageSub(TE_OnSystemMessage, pdisp, &msg1, NULL);
								}
							}
							break;
						case WM_COMMAND:
							if (pcwp->message == WM_COMMAND && g_hDialog == pcwp->hwnd) {
								if (LOWORD(pcwp->wParam) == IDOK) {
									g_bDialogOk = TRUE;
								}
							}
							break;
					}//end_switch
				}
			}
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"HookProc";
#endif
	}
	return CallNextHookEx(g_hHook, nCode, wParam, lParam);
}

UINT_PTR CALLBACK OFNHookProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LPOFNOTIFY pNotify;
	try {
		switch(msg){
			case WM_INITDIALOG:
				g_hDialog = GetParent(hwnd);
				return TRUE;
			case WM_NOTIFY:
				pNotify = (LPOFNOTIFY)lParam;
				if (pNotify->hdr.code == CDN_FOLDERCHANGE) {
					if (g_bDialogOk) {
						HWND hDlg = GetParent(hwnd);
						LRESULT nLen = SendMessage(hDlg, CDM_GETFOLDERIDLIST, 0, NULL);
						if (nLen) {
							LPITEMIDLIST pidl = (LPITEMIDLIST)::CoTaskMemAlloc(nLen);
							SendMessage(hDlg, CDM_GETFOLDERIDLIST, nLen, (LPARAM)pidl);
							BSTR bs;
							teGetDisplayNameFromIDList(&bs, pidl, SHGDN_FORPARSING | SHGDN_FORADDRESSBAR);
							teCoTaskMemFree(pidl);
							lstrcpyn(pNotify->lpOFN->lpstrFile, bs, pNotify->lpOFN->nMaxFile);
							::SysFreeString(bs);
							PostMessage(GetParent(hwnd), WM_CLOSE, 0, 0);
							return TRUE;
						}
					}
				}
				break;
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"OFNHookProc";
#endif
	}
    return FALSE;
}
#ifdef _USE_APIHOOK
VOID teAPIHook(LPWSTR pszTargetDll, LPVOID lpfnSrcProc, LPVOID lpfnNewProc)
{
	HMODULE hDll = pszTargetDll ? teLoadLibrary(pszTargetDll) : GetModuleHandle(NULL);
	if (hDll) {
		DWORD_PTR dwBase = (DWORD_PTR)hDll;
		DWORD dwIdataSize, dwOldProtect;
		for (PIMAGE_IMPORT_DESCRIPTOR pImgDesc = (PIMAGE_IMPORT_DESCRIPTOR)ImageDirectoryEntryToData(hDll,
			TRUE, IMAGE_DIRECTORY_ENTRY_IMPORT, &dwIdataSize); pImgDesc->Name; ++pImgDesc) {
			::OutputDebugStringA((char*)(dwBase + pImgDesc->Name));
			::OutputDebugStringA("\n");
			PIMAGE_THUNK_DATA pIAT = (PIMAGE_THUNK_DATA)(dwBase + pImgDesc->FirstThunk);
			PIMAGE_THUNK_DATA pINT = (PIMAGE_THUNK_DATA)(dwBase + pImgDesc->OriginalFirstThunk);
			while (pIAT->u1.Function) {
				PIMAGE_IMPORT_BY_NAME pImportName = (PIMAGE_IMPORT_BY_NAME)(dwBase+(DWORD)pINT->u1.AddressOfData);
				if (!IMAGE_SNAP_BY_ORDINAL(pINT->u1.Ordinal)) {
					::OutputDebugStringA((char*)pImportName->Name);
					::OutputDebugStringA("\n");
				}
				if ((LPVOID)pIAT->u1.Function == lpfnSrcProc) {
					VirtualProtect(&pIAT->u1.Function, sizeof(pIAT->u1.Function), PAGE_EXECUTE_READWRITE, &dwOldProtect);
					WriteProcessMemory(GetCurrentProcess(), &pIAT->u1.Function, &lpfnNewProc, sizeof(pIAT->u1.Function), &dwOldProtect);
					VirtualProtect(&pIAT->u1.Function, sizeof(pIAT->u1.Function), dwOldProtect, &dwOldProtect);
				}
				++pIAT;
				++pINT;
			}
			//			}
		}
	}
}
#endif

HRESULT teCreateWebView2(IWebBrowser2 **ppWebBrowser)
{
	HRESULT hr = E_NOTIMPL;
	if (g_nBlink == 2) {
		return hr;
	}
	WCHAR pszTEWV2[MAX_PATH];
	GetModuleFileName(NULL, pszTEWV2, MAX_PATH);
	PathRemoveFileSpec(pszTEWV2);
#ifdef _WIN64
	PathAppend(pszTEWV2, L"lib\\tewv64.dll");
#else
	PathAppend(pszTEWV2, L"lib\\tewv32.dll");
#endif
#ifdef _DEBUG
	if (!PathFileExists(pszTEWV2)) {
#ifdef _WIN64
		lstrcpy(pszTEWV2, L"C:\\cpp\\tewv\\x64\\Debug\\tewv64d.dll");
#else
		lstrcpy(pszTEWV2, L"C:\\cpp\\tewv\\Debug\\tewv32d.dll");
#endif
	}
#endif
	hr = teCreateInstance(CLSID_WebBrowserExt, pszTEWV2, &g_hTEWV, IID_PPV_ARGS(ppWebBrowser));
	if (hr == S_OK && !g_pSP) {
		VARIANT v;
		VariantInit(&v);
		(*ppWebBrowser)->GetProperty(L"version", &v);
		if (GetIntFromVariantClear(&v) >= 20112000) {
			(*ppWebBrowser)->QueryInterface(IID_PPV_ARGS(&g_pSP));
		} else {
			(*ppWebBrowser)->Release();
			hr = E_FAIL;
		}
	}
	g_nBlink = hr == S_OK ? 1 : 2;
	return hr;
}

int APIENTRY _tWinMain(HINSTANCE hInstance,
					HINSTANCE hPrevInstance,
					LPTSTR    lpCmdLine,
					int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	BSTR bsIndex = NULL;
	BSTR bsScript = NULL;
	hInst = hInstance;

	//Eliminates the vulnerable to a DLL pre-loading attack.
	//Late Binding
	HINSTANCE hDll = GetModuleHandleA("kernel32.dll");
	if (hDll) {
		*(FARPROC *)&_SetDefaultDllDirectories = GetProcAddress(hDll, "SetDefaultDllDirectories");
		if (_SetDefaultDllDirectories) {
			_SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_SYSTEM32 | LOAD_LIBRARY_SEARCH_USER_DIRS);
		}
#ifdef _2000XP
		*(FARPROC *)&_SetDllDirectoryW = GetProcAddress(hDll, "SetDllDirectoryW");
		if (_SetDllDirectoryW) {
			_SetDllDirectoryW(L"");
		}
		*(FARPROC *)&_IsWow64Process = GetProcAddress(hDll, "IsWow64Process");
#else
		SetDllDirectory(L"");
#endif
	}
	::OleInitialize(NULL);
	for (int i = MAX_CSIDL2; i--;) {
		g_pidls[i] = NULL;
		g_bsPidls[i] = NULL;
		if (i < MAX_CSIDL) {
			SHGetFolderLocation(NULL, i, NULL, NULL, &g_pidls[i]);
			if (g_pidls[i]) {
				IShellFolder *pSF;
				LPCITEMIDLIST pidlPart;
				if SUCCEEDED(SHBindToParent(g_pidls[i], IID_PPV_ARGS(&pSF), &pidlPart)) {
					teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &g_bsPidls[i]);
					pSF->Release();
				}
			}
		}
	}
	g_bUpper10 = teVerifyVersion(10, 0, 0);
#ifdef _2000XP
	g_bUpperVista = teVerifyVersion(6, 0, 0);
#else
	g_pidls[CSIDL_RESULTSFOLDER] = ILCreateFromPathA("shell:::{2965E715-EB66-4719-B53F-1672673BBEFA}");
#endif
#ifdef _W2000
	g_bIsXP = teVerifyVersion(5, 1, 0);
	g_bIs2000 = !g_bUpperVista && !g_bIsXP;
#endif
#ifdef _2000XP
	if (g_bUpperVista) {
		g_pidls[CSIDL_RESULTSFOLDER] = ILCreateFromPathA("shell:::{2965E715-EB66-4719-B53F-1672673BBEFA}");
	} else {
		g_pidls[CSIDL_RESULTSFOLDER] = ILCreateFromPathA("::{E17D4FC0-5564-11D1-83F2-00A0C90DC849}");
		g_bCharWidth = TRUE;
	}
#endif
	g_pidls[CSIDL_LIBRARY] = ILCreateFromPathA("shell:libraries");
	g_pidls[CSIDL_USER] = ILCreateFromPathA("shell:UsersFilesFolder");

	if (g_hShell32 = teLoadLibrary(L"shell32.dll")) {
#ifdef _2000XP
		*(FARPROC *)&_SHCreateItemFromIDList = GetProcAddress(g_hShell32, "SHCreateItemFromIDList");
		*(FARPROC *)&_SHGetIDListFromObject = GetProcAddress(g_hShell32, "SHGetIDListFromObject");
		*(FARPROC *)&_SHCreateShellItemArrayFromShellItem = GetProcAddress(g_hShell32, "SHCreateShellItemArrayFromShellItem");
#endif
		*(FARPROC *)&_SHRunDialog = GetProcAddress(g_hShell32, MAKEINTRESOURCEA(61));
		*(FARPROC *)&_RegenerateUserEnvironment = GetProcAddress(g_hShell32, "RegenerateUserEnvironment");
	}
#ifdef _2000XP
	if (!_SHGetIDListFromObject) {
		_SHGetIDListFromObject = teGetIDListFromObjectXP;
	}
	if (!_SHCreateItemFromIDList) {
		_SHCreateItemFromIDList = teSHCreateItemFromIDListXP;
	}
#endif
#ifdef _W2000
	if (!_SHGetImageList) {
		_SHGetImageList = teSHGetImageList2000;
	}
#endif
#ifdef _2000XP
	if (hDll = teLoadLibrary(L"propsys.dll")) {
		*(FARPROC *)&_PSPropertyKeyFromString = GetProcAddress(hDll, "PSPropertyKeyFromString");
		*(FARPROC *)&_PSGetPropertyKeyFromName = GetProcAddress(hDll, "PSGetPropertyKeyFromName");
		*(FARPROC *)&_PSGetPropertyDescription = GetProcAddress(hDll, "PSGetPropertyDescription");
		*(FARPROC *)&_PSStringFromPropertyKey = GetProcAddress(hDll, "PSStringFromPropertyKey");
		*(FARPROC *)&_PropVariantToVariant = GetProcAddress(hDll, "PropVariantToVariant");
		*(FARPROC *)&_VariantToPropVariant = GetProcAddress(hDll, "VariantToPropVariant");
		_PSPropertyKeyFromStringEx = tePSPropertyKeyFromStringEx;
	} else {
		_PropVariantToVariant = (LPFNPropVariantToVariant)teVariantToVariantXP;
		_VariantToPropVariant = (LPFNVariantToPropVariant)teVariantToVariantXP;
		_PSPropertyKeyFromStringEx = tePSPropertyKeyFromStringXP;
	}
#endif
	if (hDll = teLoadLibrary(L"user32.dll")) {
		*(FARPROC *)&_ChangeWindowMessageFilterEx = GetProcAddress(hDll, "ChangeWindowMessageFilterEx");
		*(FARPROC *)&_SetWindowCompositionAttribute = GetProcAddress(hDll, "SetWindowCompositionAttribute");
#ifdef _2000XP
		*(FARPROC *)&_ChangeWindowMessageFilter = GetProcAddress(hDll, "ChangeWindowMessageFilter");
		*(FARPROC *)&_RemoveClipboardFormatListener = GetProcAddress(hDll, "RemoveClipboardFormatListener");
		if (_RemoveClipboardFormatListener) {
			*(FARPROC *)&_AddClipboardFormatListener = GetProcAddress(hDll, "AddClipboardFormatListener");
		}
#endif
#ifdef _USE_APIHOOK
		*(FARPROC *)&_GetSysColor = GetProcAddress(hDll, "GetSysColor");
//		teAPIHook(L"shell32.dll", (LPVOID)_GetSysColor, &teGetSysColor);
//		teAPIHook(L"ieframe.dll", (LPVOID)_GetSysColor, &teGetSysColor);
//		teAPIHook(L"shlwapi.dll", (LPVOID)_GetSysColor, &teGetSysColor);
//		teAPIHook(L"user32.dll", (LPVOID)_GetSysColor, &teGetSysColor);
//		teAPIHook(L"comctl32.dll", (LPVOID)_GetSysColor, &teGetSysColor);
//		teAPIHook(NULL, (LPVOID)_GetSysColor, &teGetSysColor);
		if (hDll = teLoadLibrary(L"advapi32.dll")) {
			*(FARPROC *)&_RegQueryValueW = GetProcAddress(hDll, "RegQueryValueW");
//			teAPIHook(L"comctl32.dll", (LPVOID)_RegQueryValueW, &teRegQueryValueW);
//			teAPIHook(L"ieframe.dll", (LPVOID)_RegQueryValueW, &teRegQueryValueW);
//			teAPIHook(L"shlwapi.dll", (LPVOID)_RegQueryValueW, &teRegQueryValueW);
//			teAPIHook(L"shell32.dll", (LPVOID)_RegQueryValueW, &teRegQueryValueW);
			*(FARPROC *)&_RegQueryValueExW = GetProcAddress(hDll, "RegQueryValueExW");
//			teAPIHook(L"ieframe.dll", (LPVOID)_RegQueryValueExW, &teRegQueryValueExW);
//			teAPIHook(NULL, (LPVOID)_RegQueryValueW, &teRegQueryValueW);
		}
#endif
	}
	if (hDll = teLoadLibrary(L"ntdll.dll")) {
		*(FARPROC *)&_RtlGetVersion = GetProcAddress(hDll, "RtlGetVersion");
	}

	if (teVerifyVersion(10, 0, 17763)) {
		if (hDll = teLoadLibrary(L"uxtheme.dll")) {
			*(FARPROC *)&_ShouldAppsUseDarkMode = GetProcAddress(hDll, MAKEINTRESOURCEA(132));
			*(FARPROC *)&_AllowDarkModeForWindow = GetProcAddress(hDll, MAKEINTRESOURCEA(133));
			*(FARPROC *)&_SetPreferredAppMode = GetProcAddress(hDll, MAKEINTRESOURCEA(135));
			*(FARPROC *)&_RefreshImmersiveColorPolicyState = GetProcAddress(hDll, MAKEINTRESOURCEA(104));
//			*(FARPROC *)&_IsDarkModeAllowedForWindow = (LPFNIsDarkModeAllowedForWindow)GetProcAddress(hDll, MAKEINTRESOURCEA(137));
#ifdef _USE_APIHOOK
			*(FARPROC *)&_OpenNcThemeData = GetProcAddress(hDll, MAKEINTRESOURCEA(49));
			teAPIHook(L"comctl32.dll", (LPVOID)_OpenNcThemeData, &teOpenNcThemeData);
#endif
			teGetDarkMode();
		}
	}
	if (!_SetWindowCompositionAttribute) {
		if (hDll = teLoadLibrary(L"dwmapi.dll")) {
			*(FARPROC *)&_DwmSetWindowAttribute = GetProcAddress(hDll, "DwmSetWindowAttribute");
		}
	}
	if (hDll = teLoadLibrary(L"shcore.dll")) {
		*(FARPROC *)&_GetDpiForMonitor = GetProcAddress(hDll, "GetDpiForMonitor");
	}
	if (!_GetDpiForMonitor) {
		_GetDpiForMonitor = teGetDpiForMonitor;
	}	
#ifdef _2000XP
	if (g_bUpperVista) {
#endif
		teChangeWindowMessageFilterEx(NULL, WM_COPYDATA, MSGFLT_ALLOW, NULL);
#ifdef _2000XP
	}
#endif
	g_bsName = tePSGetNameFromPropertyKeyEx(PKEY_ItemNameDisplay, 0, NULL);
	PROPERTYKEY AnyText;
	AnyText.fmtid = PKEY_FullText.fmtid;
	AnyText.pid = 12;
	g_bsAnyText = tePSGetNameFromPropertyKeyEx(AnyText, 0, NULL);
	for (LPWSTR pszSpace = g_bsAnyText; pszSpace = StrChr(pszSpace, ' '); lstrcpy(pszSpace, &pszSpace[1]));
	g_bsAnyTextId = tePSGetNameFromPropertyKeyEx(AnyText, 1, NULL);

	InitCommonControls();

	for (UINT i = 0; i < 256; ++i) {
		UINT c = i;
		for (int j = 8; j--;) {
			c = (c & 1) ? (0xedb88320 ^ (c >> 1)) : (c >> 1);
		}
		g_uCrcTable[i] = c;
	}
	g_dwMainThreadId = GetCurrentThreadId();
	VariantInit(&g_vArguments);
	BSTR bsPath = NULL;
	teGetModuleFileName(NULL, &bsPath);//Executable Path
	for (int i = 0; bsPath[i]; ++i) {
		if (bsPath[i] == '\\') {
			bsPath[i] = '/';
		}
		bsPath[i] = towupper(bsPath[i]);
	}
	LONG_PTR nHash;
	HashData((BYTE *)bsPath, lstrlen(bsPath) * sizeof(WCHAR), (LPBYTE)&nHash, sizeof(LONG_PTR));
	//Command Line
	BOOL bVisible = !teStartsText(L"/run ", lpCmdLine);
	BOOL bNewProcess = teStartsText(L"/open ", lpCmdLine);
	LPWSTR szClass = WINDOW_CLASS2;
	if (bVisible && !bNewProcess) {
		//Multiple Launch
		szClass = WINDOW_CLASS;
		HWND hwndTE = NULL;
		while (hwndTE = FindWindowEx(NULL, hwndTE, szClass, NULL)) {
			if (GetWindowLongPtr(hwndTE, GWLP_USERDATA) == nHash) {
				BSTR bs = teGetCommandLine();
				COPYDATASTRUCT cd;
				cd.dwData = 0;
				cd.lpData = bs;
				cd.cbData = ::SysStringByteLen(bs) + sizeof(WCHAR);
				DWORD_PTR dwResult;
				LRESULT lResult = SendMessageTimeout(hwndTE, WM_COPYDATA, nCmdShow, LPARAM(&cd), SMTO_ABORTIFHUNG, 1000 * 30, &dwResult);
				teSysFreeString(&bs);
				if (lResult && dwResult == S_OK) {
					SysFreeString(bsPath);
					Finalize();
					teSetForegroundWindow(hwndTE);
					return FALSE;
				}
			}
		}
	}
	SysFreeString(bsPath);
	Initlize();
	MSG msg;

	//Initialize FolderView & TreeView settings
	g_paramFV[SB_Type] = 1;
	g_paramFV[SB_ViewMode] = FVM_DETAILS;
	g_paramFV[SB_FolderFlags] = FWF_SHOWSELALWAYS;
	g_paramFV[SB_IconSize] = 0;
	g_paramFV[SB_Options] = EBO_ALWAYSNAVIGATE;
	g_paramFV[SB_ViewFlags] = 0;
	g_paramFV[SB_TreeAlign] = 1;
	g_paramFV[SB_TreeWidth] = 200;
	g_paramFV[SB_TreeFlags] = NSTCS_HASEXPANDOS | NSTCS_SHOWSELECTIONALWAYS | NSTCS_BORDER | NSTCS_HASLINES | NSTCS_NOINFOTIP | NSTCS_HORIZONTALSCROLL;
	g_paramFV[SB_EnumFlags] = SHCONTF_FOLDERS;
	g_paramFV[SB_RootStyle] = NSTCRS_VISIBLE | NSTCRS_EXPANDED;

	MyRegisterClass(hInstance, szClass, WndProc);

	g_pSortColumnNull[0].propkey = PKEY_Search_Rank;
	g_pSortColumnNull[0].direction = SORT_DESCENDING;
	g_pSortColumnNull[1].propkey = PKEY_DateModified;
	g_pSortColumnNull[1].direction = SORT_DESCENDING;
	g_pSortColumnNull[2].propkey = PKEY_ItemNameDisplay;
	g_pSortColumnNull[2].direction = SORT_ASCENDING;
	// Title & Version
	lstrcpy(g_szTE, _T(PRODUCTNAME));
	g_bsTitle = ::SysAllocString(g_szTE);
	if (bVisible) {
		g_bInit = !bNewProcess;
		g_hwndMain = CreateWindowEx(WS_EX_LAYERED, szClass, g_szTE, WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		  CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInstance, NULL);
		SetLayeredWindowAttributes(g_hwndMain, 0, 0, LWA_ALPHA);
		if (!bNewProcess) {
			SetWindowLongPtr(g_hwndMain, GWLP_USERDATA, nHash);
		}
	} else {
		g_hwndMain = CreateWindowEx(WS_EX_TOOLWINDOW, szClass, g_szTE, WS_POPUP,
		  CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, NULL, NULL, hInstance, NULL);
	}
	if (!g_hwndMain)
	{
		Finalize();
		return FALSE;
	}
	teGetDarkMode();
	// ClipboardFormat
	IDLISTFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_SHELLIDLIST);
	DROPEFFECTFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
	DRAGWINDOWFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(L"DragWindow");
	IsShowingLayeredFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(L"IsShowingLayered");
	// Hook
	g_hHook = SetWindowsHookEx(WH_CALLWNDPROC, (HOOKPROC)HookProc, hInst, g_dwMainThreadId);
	g_hMouseHook = SetWindowsHookEx(WH_MOUSE, (HOOKPROC)MouseProc, hInst, g_dwMainThreadId);
	// Create own class
	CoCreateGuid(&g_ClsIdSB);
	CoCreateGuid(&g_ClsIdTC);
	CoCreateGuid(&g_ClsIdStruct);
	CoCreateGuid(&g_ClsIdFI);
	// IUIAutomation
	if FAILED(teCreateInstance(CLSID_CUIAutomation, NULL, NULL, IID_PPV_ARGS(&g_pAutomation))) {
		g_pAutomation = NULL;
	}
	IUIAutomationRegistrar *pUIAR;
	if SUCCEEDED(teCreateInstance(CLSID_CUIAutomationRegistrar, NULL, NULL, IID_PPV_ARGS(&pUIAR))) {
		UIAutomationPropertyInfo info = {ItemIndex_Property_GUID, L"ItemIndex", UIAutomationType_Int};
		pUIAR->RegisterProperty(&info, &g_PID_ItemIndex);
		pUIAR->Release();
	}
#ifdef _2000XP
	if (_AddClipboardFormatListener) {
		_AddClipboardFormatListener(g_hwndMain);
	} else {
		g_hwndNextClip = SetClipboardViewer(g_hwndMain);
	}
#else
	AddClipboardFormatListener(g_hwndMain);
#endif
	//WebView2 Check
	IWebBrowser2 *pWB;
	if (teCreateWebView2(&pWB) == S_OK) {
		pWB->Release();
	}
	//Get JScript Object
	bsScript = teMultiByteToWideChar(CP_UTF8, "\
function _o() {\
	return {};\
}\
function _a() {\
	return [];\
}\
function _t(o) {\
	return {window: {external: o}};\
}\
function _c(s) {\
	return s.replace(/\\0$/, '').split(/\\0/);\
}\
", -1);
	if (ParseScript(bsScript, L"JScript", NULL, &g_pJS, NULL) != S_OK) {
		teSysFreeString(&bsScript);
		PostMessage(g_hwndMain, WM_CLOSE, 0, 0);
		Finalize();
		MessageBoxA(NULL, "503 Script Engine Unavalable", NULL, MB_OK | MB_ICONERROR);
		return FALSE;
	}
	teSysFreeString(&bsScript);
	IGlobalInterfaceTable *pGlobalInterfaceTable;
	CoCreateInstance(CLSID_StdGlobalInterfaceTable, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pGlobalInterfaceTable));
	pGlobalInterfaceTable->RegisterInterfaceInGlobal(g_pJS, IID_IDispatch, &g_dwCookieJS);
	InitializeCriticalSection(&g_csFolderSize);
	//WindowsAPI
#ifdef _USE_OBJECTAPI
	GetNewObject(&g_pAPI);
	VARIANT v;
	for (int i = _countof(dispAPI); i--;) {
		teSetObjectRelease(&v, new CteAPI(&dispAPI[i]));
		BSTR bs = teMultiByteToWideChar(CP_UTF8, dispAPI[i].name, -1);
		tePutProperty(g_pAPI, bs, &v);
		::SysFreeString(bs);
	}
	teSetSZ(&v, L"ADODB.Stream");
	tePutProperty(g_pAPI, L"ADBSTRM", &v);
#endif
	// CTE
	::ZeroMemory(g_param, sizeof(g_param));
	g_param[TE_Type] = CTRL_TE;
	g_param[TE_CmdShow] = nCmdShow;
	g_param[TE_Layout] = 0x80;
	g_param[TE_UseHiddenFilter] = TRUE;
	VariantInit(&g_vData);
	g_pTE = new CTE(g_hwndMain);
	g_pDropTarget2 = new CteDropTarget2(g_hwndMain, static_cast<IDispatch *>(g_pTE), TRUE);
	RegisterDragDrop(g_hwndMain, g_pDropTarget2);

	IDispatch *pJS = NULL;
#ifdef _USE_HTMLDOC
	IHTMLDocument2	*pHtmlDoc = NULL;
#else
	VARIANT vWindow;
	VariantInit(&vWindow);
#endif
	teGetModuleFileName(NULL, &bsPath);//Executable Path
	PathRemoveFileSpec(bsPath);
	if (bVisible) {
		LPTSTR *lplpszArgs = NULL;
		LPWSTR lpFile = L"script\\index.html";
		if (bNewProcess && lpCmdLine && lpCmdLine[0]) {
			int nCount = 0;
			lplpszArgs = CommandLineToArgvW(lpCmdLine, &nCount);
			if (nCount > 1) {
				lpFile = lplpszArgs[1];
			}
		}
		if (teIsFileSystem(lpFile) || teStartsText(L"file://", lpFile)) {
			bsIndex = ::SysAllocString(lpFile);
		} else {
			tePathAppend(&bsIndex, bsPath, lpFile);
		}
		if (lplpszArgs) {
			LocalFree(lplpszArgs);
		}
		CoInternetSetFeatureEnabled(FEATURE_DISABLE_NAVIGATION_SOUNDS, SET_FEATURE_ON_PROCESS, TRUE);
		g_pWebBrowser = new CteWebBrowser(g_hwndMain, bsIndex, NULL);
		SetTimer(g_hwndMain, TET_Create, 10000, teTimerProc);
	} else {
		tePathAppend(&bsIndex, bsPath, L"script\\background.js");
		bsScript = teLoadFromFile(bsIndex);
		VARIANT vResult;
		VariantInit(&vResult);
#ifdef _USE_HTMLDOC
		// CLSID_HTMLDocument
		if SUCCEEDED(teCreateInstance(CLSID_HTMLDocument, NULL, NULL, IID_PPV_ARGS(&pHtmlDoc))) {
			ICustomDoc *pCD = NULL;
			if SUCCEEDED(pHtmlDoc->QueryInterface(IID_PPV_ARGS(&pCD))) {
				CteDocHostUIHandler *pDHUIH = new CteDocHostUIHandler();
				pCD->SetUIHandler(pDHUIH);
				pDHUIH->Release();
				pCD->Release();
			}
			SAFEARRAY *psa = SafeArrayCreateVector(VT_VARIANT, 0, 1);
			if (psa) {
				VARIANT *pv;
				SafeArrayAccessData(psa, (LPVOID *)&pv);
				pv->vt = VT_BSTR;
				LPOLESTR lpHtml = L"<html><body><script type='text/javascript'>%s external.Data=_s;</script></body></html>";
				int nSize = lstrlen(lp) + lstrlen(lpHtml) - 1;
				pv->bstrVal = ::SysAllocStringLen(NULL, nSize);
				swprintf_s(pv->bstrVal, nSize, lpHtml, lp);
				SafeArrayUnaccessData(psa);
				pHtmlDoc->write(psa);
				VariantClear(pv);
				pHtmlDoc->close();
				SafeArrayDestroy(psa);
				IDispatch *pdisp;
				if (GetDispatch(&g_pTE->m_Data, &pdisp)) {
					Invoke5(pdisp, DISPID_VALUE, DISPATCH_METHOD, &vResult, 0, NULL);
					pdisp->Release();
				}
			}
		}
#else
		// IActiveScript
		VARIANTARG *pv = GetNewVARIANT(1);
		teSetObject(&pv[0], g_pTE);
		teExecMethod(g_pJS, L"_t", &vWindow, 1, pv);
		ParseScript(bsScript, L"JScript", &vWindow, &pJS, NULL);
		teExecMethod(pJS, L"_s", &vResult, 0, NULL);
		teSysFreeString(&bsScript);
#endif
		bVisible = (vResult.vt == VT_BSTR) && (lstrcmpi(vResult.bstrVal, L"wait") == 0);
		VariantClear(&vResult);
	}
	if (!bVisible) {
		PostMessage(g_hwndMain, WM_CLOSE, 0, 0);
	}

	teSysFreeString(&bsIndex);
	teSysFreeString(&bsPath);
	teCreateInstance(CLSID_DragDropHelper, NULL, NULL, IID_PPV_ARGS(&g_pDropTargetHelper));

	// Main message loop:
	while (g_bMessageLoop) {
		try {
			if (!GetMessage(&msg, NULL, 0, 0)) {
				break;
			}
			if (MessageProc(&msg) == S_OK) {
				continue;
			}
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"MainMessageLoop";
#endif
		}
	}
	//At the end of processing
	try {
		g_bMessageLoop = FALSE;
		RevokeDragDrop(g_hwndMain);
		SafeRelease(&g_pDropTarget2);
#ifdef _USE_HTMLDOC
		SafeRelease(&pHtmlDoc);
#else
		SafeRelease(&pJS);
		VariantClear(&vWindow);
#endif
		for (size_t u = 0; u < g_pFreeLibrary.size(); ++u) {
			FreeLibrary(g_pFreeLibrary[u]);
		}
		g_pFreeLibrary.clear();
		for (size_t u = 0; u < g_phModule.size(); ++u) {
			FreeLibrary(g_phModule[u]);
		}
		g_phModule.clear();
		for (size_t u = g_pSubWindows.size(); u--;) {
			g_pSubWindows[u].pWB->Release();
		}
		g_pSubWindows.clear();
#ifdef _USE_OBJECTAPI
		SafeRelease(&g_pAPI);
#endif
		pGlobalInterfaceTable->RevokeInterfaceFromGlobal(g_dwCookieJS);
		pGlobalInterfaceTable->Release();
		SafeRelease(&g_pJS);
		SafeRelease(&g_pAutomation);
		SafeRelease(&g_pDropTargetHelper);
		DeleteCriticalSection(&g_csFolderSize);
		UnhookWindowsHookEx(g_hMouseHook);
		UnhookWindowsHookEx(g_hHook);
		if (g_hMenu) {
			DestroyMenu(g_hMenu);
		}
#ifdef	_2000XP
		if (_AddClipboardFormatListener) {
			_RemoveClipboardFormatListener(g_hwndMain);
		} else {
			ChangeClipboardChain(g_hwndMain, g_hwndNextClip);
		}
#else
		RemoveClipboardFormatListener(g_hwndMain);
#endif
	} catch (...) {}
	Finalize();
	SafeRelease(&g_pTE);
	VariantClear(&g_vData);
	return (int) msg.wParam;
}
VOID ArrangeWindow()
{
	if (g_bArrange || !g_pOnFunc[TE_OnArrange]) {
		return;
	}
	g_bArrange = TRUE;
	SetTimer(g_hwndMain, TET_Size, 100, teTimerProc);
}

#ifdef _2000XP
VOID CALLBACK teTimerProcRelease(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	try {
		KillTimer(hwnd, idEvent);
		for (int i = g_pRelease.size(); i--;) {
			g_pRelease[i]->Release();
		}
		g_pRelease.clear();
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teTimerProcRelease";
#endif
	}
}

VOID teDelayRelease(PVOID ppObj)
{
	try {
		IUnknown **ppunk = static_cast<IUnknown **>(ppObj);
		if (*ppunk) {
			g_pRelease.push_back(*ppunk);
			SetTimer(g_hwndMain, (UINT_PTR)teTimerProcRelease, 100, teTimerProcRelease);
			*ppunk = NULL;
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teDelayRelease";
#endif
	}
}
#endif

VOID CALLBACK teTimerProcParse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);
	TEInvoke *pInvoke = (TEInvoke *)idEvent;
	try {
		if (::InterlockedDecrement(&pInvoke->cDo) == 0) {
			if (pInvoke->pidl || pInvoke->wErrorHandling == 3) {
				VariantClear(&pInvoke->pv[pInvoke->cArgs - 1]);
				if (pInvoke->wMode) {
					teSetLong(&pInvoke->pv[pInvoke->cArgs - 1], pInvoke->hr);
				} else if (pInvoke->pidl) {
					CteFolderItem *pid1 = new CteFolderItem(NULL);
					if SUCCEEDED(pid1->Initialize(pInvoke->pidl)) {
						teSetSZ(&pid1->m_v, pInvoke->bsPath);
					}
					teSetObjectRelease(&pInvoke->pv[pInvoke->cArgs - 1], pid1);
				}
				Invoke5(pInvoke->pdisp, pInvoke->dispid, DISPATCH_METHOD, NULL, -pInvoke->cArgs, pInvoke->pv);
			} else if (pInvoke->wErrorHandling == 1) {
				VARIANT *pv = &pInvoke->pv[pInvoke->cArgs - 1];
				CteFolderItem *pid = new CteFolderItem(pv);
				pid->MakeUnavailable();
				VariantClear(pv);
				teSetSZ(&pid->m_v, pInvoke->bsPath);
				teSetObjectRelease(pv, pid);
				Invoke5(pInvoke->pdisp, pInvoke->dispid, DISPATCH_METHOD, NULL, -pInvoke->cArgs, pInvoke->pv);
			} else if (pInvoke->wErrorHandling == 2) {
				if (g_bShowParseError) {
					CteShellBrowser *pSB = NULL;
					pInvoke->pdisp->QueryInterface(g_ClsIdSB, (LPVOID *)&pSB);
					if (pSB == g_pTC->GetShellBrowser(g_pTC->m_nIndex)) {
						g_bShowParseError = FALSE;
						LPWSTR lpBuf;
						WCHAR pszFormat[MAX_STATUS];
						if (LoadString(g_hShell32, 6456, pszFormat, MAX_STATUS)) {
							if (FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ARGUMENT_ARRAY,
								pszFormat, 0, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&lpBuf, MAX_STATUS, (va_list *)&pInvoke->bsPath)) {
								int r = MessageBox(g_hwndMain, lpBuf, _T(PRODUCTNAME), MB_ABORTRETRYIGNORE);
								LocalFree(lpBuf);
								if (r == IDRETRY) {
									::InterlockedIncrement(&pInvoke->cDo);
									::InterlockedIncrement(&pInvoke->cRef);
									_beginthread(&threadParseDisplayName, 0, pInvoke);
									g_bShowParseError = TRUE;
								} else {
									g_bShowParseError = (r != IDIGNORE);
								}
							}
						}
					}
					SafeRelease(&pSB);
				}
			}
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teTimerProcParse";
#endif
	}
	teReleaseInvoke(pInvoke);
}

VOID CALLBACK teTimerProcForTree(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);
	try {
		CteTreeView *pTV = TVfromhwnd(hwnd);
		if (pTV) {
			if (g_nLockUpdate && (pTV->m_pFV && pTV->m_pFV->m_pTC->m_nLockUpdate)) {
				return;
			}
			switch (idEvent) {
			case TET_SetRoot:
				pTV->SetRoot();
				//break;//Expand after SetRoot
			case TET_Expand:
				pTV->Expand();
				break;
			}
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"teTimerProcForTree";
#endif
	}
}

static void threadParseDisplayName(void *args)
{
	::CoInitialize(NULL);
	TEInvoke *pInvoke = (TEInvoke *)args;
#ifdef _DEBUG
	WCHAR pszDebug[2048];
	swprintf_s(pszDebug, 2048, L"%s:%d:%d\n", pInvoke->bsPath, pInvoke->wMode, pInvoke->wErrorHandling);
	::OutputDebugString(pszDebug);
#endif
	try {
		if (!pInvoke->pidl && pInvoke->bsPath) {
			pInvoke->pidl = teILCreateFromPath1(pInvoke->bsPath);
		}
		if (pInvoke->wMode) {
			pInvoke->hr = E_PATH_NOT_FOUND;
			if (pInvoke->pidl) {
				pInvoke->hr = teILFolderExists(pInvoke->pidl);
			}
		}
		KillTimer(g_hwndMain, (UINT_PTR)pInvoke);
		pInvoke->cRef = 1;
		SetTimer(g_hwndMain, (UINT_PTR)pInvoke, 1, teTimerProcParse);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"threadParseDisplayName";
#endif
	}
	::CoUninitialize();
	::_endthread();
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	MSG msg1;
	HRESULT hr;

	try {
		msg1.hwnd = hWnd;
		msg1.message = message;
		msg1.wParam = wParam;
		msg1.lParam = lParam;
		switch (message)
		{
			case WM_CLOSE:
				g_nLockUpdate += 10000;
				try {
					if (CanClose(g_pTE)) {
						teRevoke(g_pSW);
						try {
							DestroyWindow(hWnd);
						} catch (...) {
							g_bMessageLoop = FALSE;
						}
					}
				} catch (...) {
					DestroyWindow(hWnd);
				}
				g_nLockUpdate -= 10000;
				return 0;
			case WM_SIZE:
				if (g_pWebBrowser) {
					teSetObjectRects(g_pWebBrowser->m_pWebBrowser, hWnd);
				}
				ArrangeWindow();
				break;
			case WM_DESTROY:
				try {
					MessageSub(TE_OnSystemMessage, g_pTE, &msg1, NULL);
					//Close Tab controls
					g_pTC = NULL;
					for (int i = int(g_ppTC.size()); --i >= 0;) {
						g_ppTC[i]->Close(TRUE);
						SafeRelease(&g_ppTC[i]);
					}
					g_ppTC.clear();
					//Close Tree controls
					for (int i = int(g_ppTV.size()); --i >= 0;) {
						SafeRelease(&g_ppTV[i]);
					}
					g_ppTV.clear();
					//Close Browser
					if (g_pWebBrowser) {
						g_pWebBrowser->Close();
						SafeRelease(&g_pWebBrowser);
					}
				} catch (...) {
#ifdef _DEBUG
					g_strException = L"WM_DESTROY";
#endif
				}
				PostQuitMessage(0);
				return 0;
			case WM_CONTEXTMENU:	//System Menu
				msg1.pt.x = GET_X_LPARAM(lParam);
				msg1.pt.y = GET_Y_LPARAM(lParam);
				MessageSubPt(TE_OnShowContextMenu, g_pTE, &msg1);
				break;
			//Menu
			case WM_MENUSELECT:
			case WM_HELP:
			case WM_INITMENU:
			case WM_MEASUREITEM:
			case WM_DRAWITEM:
			case WM_ENTERMENULOOP:
			case WM_INITMENUPOPUP:
			case WM_MENUCHAR:
			case WM_EXITMENULOOP:
			case WM_MENURBUTTONUP:
			case WM_MENUDRAG:
			case WM_MENUGETOBJECT:
			case WM_UNINITMENUPOPUP:
			case WM_MENUCOMMAND:
				BOOL bDone;
				bDone = TRUE;

				if (g_pOnFunc[TE_OnMenuMessage]) {
					HRESULT hr = S_FALSE;
					if (MessageSub(TE_OnMenuMessage, g_pCM, &msg1, &hr) && hr == S_OK) {
						bDone = FALSE;
					}
					if (message == WM_MENUCHAR && hr & 0xffff0000) {
						return hr;
					}
				}
				if (bDone && g_pCM) {
					LRESULT lResult = 0;
					BOOL bRet = FALSE;
					IContextMenu3 *pCM3;
					IContextMenu2 *pCM2;
					DISPID dispid;
					VARIANT v;
					VariantInit(&v);
					IUnknown *punk;
					HRESULT hr = g_pCM->GetNextDispID(fdexEnumAll, DISPID_STARTENUM, &dispid);
					while (hr == S_OK) {
						if (Invoke5(g_pCM, dispid, DISPATCH_PROPERTYGET, &v, 0, NULL) == S_OK) {
							if (FindUnknown(&v, &punk)) {
								if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCM3))) {
									bRet = SUCCEEDED(pCM3->HandleMenuMsg2(message, wParam, lParam, &lResult));
									pCM3->Release();
								} else if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCM2))) {
									pCM2->HandleMenuMsg(message, wParam, lParam);
									pCM2->Release();
								}
							}
							VariantClear(&v);
						}
						hr = g_pCM->GetNextDispID(fdexEnumAll, dispid, &dispid);
					}
					if (bRet) {
						return lResult;
					}
				}
				break;
			//System
			case WM_SETTINGCHANGE:
				if (g_pSW) {
					teRegister();
				}
				teGetDarkMode();
				if (_RegenerateUserEnvironment) {
					try {
						if (lstrcmpi((LPCWSTR)lParam, L"Environment") == 0) {
							LPVOID lpEnvironment;
							_RegenerateUserEnvironment(&lpEnvironment, TRUE);
							//Not permitted to free lpEnvironment!
							//FreeEnvironmentStrings((LPTSTR)lpEnvironment);
						}
					} catch (...) {
						g_nException = 0;
#ifdef _DEBUG
						g_strException = L"RegenerateUserEnvironment";
#endif
					}
				}
			case WM_COPYDATA:
				if (g_nReload) {
					return E_FAIL;
				}
			case WM_POWERBROADCAST:
			case WM_TIMER:
			case WM_DROPFILES:
			case WM_DEVICECHANGE:
			case WM_FONTCHANGE:
			case WM_ENABLE:
			case WM_NOTIFY:
			case WM_SETCURSOR:
			case WM_SYSCOLORCHANGE:
			case WM_THEMECHANGED:
			case WM_USERCHANGED:
			case WM_QUERYENDSESSION:
			case WM_HOTKEY:
			case WM_DPICHANGED:
			case WM_NCCALCSIZE:
			case TWM_CLIPBOARDUPDATE:
				if (MessageSub(TE_OnSystemMessage, g_pTE, &msg1, &hr)) {
					return hr;
				}
				break;
			case WM_MOVE:
				if (g_pWebBrowser) {
					teSetObjectRects(g_pWebBrowser->m_pWebBrowser, hWnd);
				}
				if (MessageSub(TE_OnSystemMessage, g_pTE, &msg1, &hr)) {
					return hr;
				}
				break;
			case WM_SYSCOMMAND:
				if (wParam >= 0xf000) {
					if ((wParam & 0xFFF0) == SC_MINIMIZE) {
						if (GetWindowLong(hWnd, GWL_EXSTYLE) & WS_EX_LAYERED) {
							BYTE alpha;
							GetLayeredWindowAttributes(hWnd, NULL, &alpha, NULL);
							if (alpha == 0) {
								teSetExStyleAnd(hWnd, ~WS_EX_LAYERED);
								return 1;
							}
						}
					}
				}
				if (MessageSub(TE_OnSystemMessage, g_pTE, &msg1, &hr)) {
					return hr;
				}
				break;
			case WM_COMMAND:
				if (teDoCommand(g_pTE, hWnd, message, wParam, lParam) == S_OK) {
					return 1;
				}
				break;
#ifdef	_2000XP
			case WM_CHANGECBCHAIN:
				if ((HWND)wParam == g_hwndNextClip) {
					g_hwndNextClip = (HWND)lParam;
				} else if (g_hwndNextClip) {
					SendMessage(g_hwndNextClip, message, wParam, lParam);
				}
				return 0;
			case WM_DRAWCLIPBOARD:
				if (g_hwndNextClip) {
					SendMessage(g_hwndNextClip, message, wParam, lParam);
				}
#endif
			case WM_CLIPBOARDUPDATE:
				PostMessage(g_hwndMain, TWM_CLIPBOARDUPDATE, 0, 0);
				break;
			case WM_CTLCOLORSTATIC:
				return (LRESULT)GetClassLongPtr(hWnd, GCLP_HBRBACKGROUND);
			default:
				if ((message > WM_APP && message <= MAXWORD)) {
					if (MessageSub(TE_OnAppMessage, g_pTE, &msg1, &hr) && hr == S_OK) {
						return hr;
					}
				}
				return DefWindowProc(hWnd, message, wParam, lParam);
		}//end_switch
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"WndProc";
#endif
	}
	return DefWindowProc(hWnd, message, wParam, lParam);
}

// Sub window

LRESULT CALLBACK WndProc2(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	try {
		CteWebBrowser *pWB;
		switch (message)
		{
			case WM_CLOSE:
				try {
					if (teGetSubWindow(hWnd, &pWB)) {
						if (pWB->m_nClose == 0) {
							pWB->m_nClose = 1;
							if (pWB->m_ppDispatch[WB_OnClose]) {
								VARIANT v;
								v.vt = VT_DISPATCH;
								v.pdispVal = pWB;
								Invoke4(pWB->m_ppDispatch[WB_OnClose], NULL, -1, &v);
								return 0;
							}
						}
						if (pWB->m_nClose == 1) {
							pWB->m_nClose = 2;
							::SysReAllocString(&pWB->m_bstrPath, PATH_BLANK);
							pWB->m_pWebBrowser->Navigate(pWB->m_bstrPath, NULL, NULL, NULL, NULL);
							return 0;
						}
						if (GetForegroundWindow() == pWB->m_hwndParent) {
							teSetForegroundWindow((HWND)GetWindowLongPtr(pWB->m_hwndParent, GWLP_HWNDPARENT));
						}
					}
				} catch(...) {
					g_nException = 0;
#ifdef _DEBUG
					g_strException = L"WM_CLOSE";
#endif
				}
				try {
					DestroyWindow(hWnd);
				} catch (...) {
					g_nException = 0;
#ifdef _DEBUG
					g_strException = L"DestroyWindow";
#endif
				}
				return 0;
			case WM_SIZE:
			case WM_MOVE:
				if (teGetSubWindow(hWnd, &pWB)) {
					teSetObjectRects(pWB->m_pWebBrowser, hWnd);
				}
				break;
			case WM_ACTIVATE:
				if (LOWORD(wParam)) {
					if (teGetSubWindow(hWnd, &pWB)) {
						SetTimer(pWB->m_hwndParent, (UINT_PTR)pWB, 100, teTimerProcSW);
					}
				}
				break;
			case WM_ERASEBKGND:
				return 1;
			default:
				return DefWindowProc(hWnd, message, wParam, lParam);
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"WndProc2";
#endif
	}
	return 0;
}

// CteShellBrowser

CteShellBrowser::CteShellBrowser(CteTabCtrl *pTC)
{
	m_cRef = 1;
	m_bInit = FALSE;
	m_bsFilter = NULL;
	m_bsAltSortColumn = NULL;
	m_bsNextGroup = NULL;
	m_pTV = new CteTreeView();
	m_dwCookie = 0;
	VariantInit(&m_vData);
	m_nForceViewMode = FVM_AUTO;
	Init(pTC, TRUE);
}

void CteShellBrowser::Init(CteTabCtrl *pTC, BOOL bNew)
{
	m_bEmpty = FALSE;
	m_hwnd = NULL;
	m_hwndLV = NULL;
	m_hwndDV = NULL;
	m_hwndDT = NULL;
	m_hwndAlt = NULL;
	m_pShellView = NULL;
	m_pExplorerBrowser = NULL;
	m_dwEventCookie = 0;
	m_pidl = NULL;
	m_pFolderItem = NULL;
	m_pFolderItem1 = NULL;
	m_nUnload = 0;
	m_pDropTarget2 = NULL;
	m_pdisp = NULL;
	m_pSF2 = NULL;
	m_bRefreshing = FALSE;
	m_pSFVCB = NULL;
#ifdef _2000XP
	m_nColumns = 0;
	m_dispidSortColumns = 0;
#endif
	m_bCheckLayout = FALSE;
	m_bRefreshLator = FALSE;
	m_nCreate = 0;
	m_bNavigateComplete = FALSE;
	m_dwUnavailable = 0;
	m_dwTickNotify = 0;
	VariantClear(&m_vData);

	for (int i = SB_Count; i--;) {
		m_param[i] = g_paramFV[i];
	}

	m_pTC = pTC;
	for (int i = Count_SBFunc; i--;) {
		m_ppDispatch[i] = NULL;
	}
	m_bVisible = FALSE;
	m_uLogIndex = -1;
	m_nFolderSizeIndex = MAXINT;
	m_nLabelIndex = MAXINT;
	m_nSizeIndex = MAXINT;
	m_nLinkTargetIndex = MAXINT;

	m_pTV->Init(this, m_pTC->m_hwndStatic);
	DoFunc(TE_OnCreate, this, E_NOTIMPL);
}

CteShellBrowser::~CteShellBrowser()
{
	Clear();
}

void CteShellBrowser::Clear()
{
	for (int i = _countof(m_ppDispatch); --i >= 0;) {
		SafeRelease(&m_ppDispatch[i]);
	}
	try {
		DestroyView(0);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"DestroyView";
#endif
	}
	try {
		while (!m_ppLog.empty()) {
			m_ppLog.back()->Release();
			m_ppLog.pop_back();
		}
		teSysFreeString(&m_bsFilter);
		teSysFreeString(&m_bsAltSortColumn);
		teSysFreeString(&m_bsNextGroup);
		teUnadviseAndRelease(m_pdisp, DIID_DShellFolderViewEvents, &m_dwCookie);
		m_pdisp = NULL;
#ifdef _2000XP
		if (m_nColumns) {
			delete [] m_pColumns;
			m_nColumns = 0;
		}
#endif
		SafeRelease(&m_pSF2);
		teILFreeClear(&m_pidl);
		SafeRelease(&m_pFolderItem);
		SafeRelease(&m_pFolderItem1);
		SafeRelease(&m_pTV->m_psiFocus);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"ShellBrowser::Clear";
#endif
	}
	VariantClear(&m_vData);
	for (size_t n = m_ppColumns.size(); n;) {
		SafeRelease(&m_ppColumns[--n]);
	}
	m_dwSessionId = 0;
}

STDMETHODIMP CteShellBrowser::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteShellBrowser, IShellBrowser),
		QITABENT(CteShellBrowser, ICommDlgBrowser),
		QITABENT(CteShellBrowser, ICommDlgBrowser2),
//		QITABENT(CteShellBrowser, ICommDlgBrowser3),
//		QITABENT(CteShellBrowser, IFolderFilter),
		QITABENT(CteShellBrowser, IOleWindow),
		QITABENT(CteShellBrowser, IDispatch),
		QITABENT(CteShellBrowser, IShellFolderViewDual),
		QITABENT(CteShellBrowser, IExplorerBrowserEvents),
		QITABENT(CteShellBrowser, IExplorerPaneVisibility),
		QITABENT(CteShellBrowser, IPersist),
		QITABENT(CteShellBrowser, IPersistFolder),
		QITABENT(CteShellBrowser, IPersistFolder2),
		QITABENT(CteShellBrowser, IShellFolderViewCB),
		{ 0 },
	};
	HRESULT hr = QISearch(this, qit, riid, ppvObject);
	if SUCCEEDED(hr) {
		return hr;
	}
	*ppvObject = NULL;
	if (IsEqualIID(riid, g_ClsIdSB)) {
		*ppvObject = this;
	} else if (IsEqualIID(riid, IID_IShellFolder) || IsEqualIID(riid, IID_IShellFolder2)) {
		if (m_pSF2) {
#ifdef _2000XP
			*ppvObject = static_cast<IShellFolder2 *>(this);
#else
			return m_pSF2->QueryInterface(riid, ppvObject);
#endif
		} else {
			if (m_pShellView) {
				IFolderView *pFV;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
					hr = pFV->GetFolder(riid, ppvObject);
					pFV->Release();
				}
			}
			return hr;
		}
	} else if (IsEqualIID(riid, IID_IServiceProvider)) {
		*ppvObject = static_cast<IServiceProvider *>(new CteServiceProvider(static_cast<IShellBrowser *>(this), m_pShellView));
		return S_OK;
	} else if (m_pShellView && IsEqualIID(riid, IID_IShellView)) {
		return m_pShellView->QueryInterface(riid, ppvObject);
	} else if (m_pFolderItem && (IsEqualIID(riid, IID_FolderItem) || IsEqualIID(riid, g_ClsIdFI))) {
		return m_pFolderItem->QueryInterface(riid, ppvObject);
	} else {
		if (m_pSF2) {
			return m_pSF2->QueryInterface(riid, ppvObject);
		}
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
}

STDMETHODIMP_(ULONG) CteShellBrowser::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteShellBrowser::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		return 0;
	}

	return m_cRef;
}

//IOleWindow
STDMETHODIMP CteShellBrowser::GetWindow(HWND *phwnd)
{
	*phwnd = m_pTC->m_hwndStatic;
	return S_OK;
}

STDMETHODIMP CteShellBrowser::ContextSensitiveHelp(BOOL fEnterMode)
{
	return E_NOTIMPL;
}

//IShellBrowser
/*//
void CteShellBrowser::InitializeMenuItem(HMENU hmenu, LPTSTR lpszItemName, int nId, HMENU hmenuSub)
{
	MENUITEMINFO mii = { sizeof(MENUITEMINFO), MIIM_ID | MIIM_STRING | MIIM_SUBMENU, 0, 0, nId, hmenuSub, NULL, NULL, 0, lpszItemName };
	InsertMenuItem(hmenu, nId, FALSE, &mii);
}
*///
STDMETHODIMP CteShellBrowser::InsertMenusSB(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths)
{
/*//
	if (!g_hMenu) {
		g_hMenu = LoadMenu(g_hShell32, MAKEINTRESOURCE(216));
		if (!g_hMenu) {
			TEmethod pull_downs[] = {
				{ FCIDM_MENU_FILE, L"&File" },
				{ FCIDM_MENU_EDIT, L"&Edit" },
				{ FCIDM_MENU_VIEW, L"&View" },
				{ FCIDM_MENU_TOOLS, L"&Tools" },
				{ FCIDM_MENU_HELP, L"&Help" },
			};
			for (size_t i = 0; i < 5; ++i) {
				InitializeMenuItem(hmenuShared, pull_downs[i].name, pull_downs[i].id, CreatePopupMenu());
				lpMenuWidths->width[i] = 0;
			}
			lpMenuWidths->width[0] = 2; // FILE: File, Edit
			lpMenuWidths->width[2] = 2; // VIEW: View, Tools
			lpMenuWidths->width[4] = 1; // WINDOW: Help
		}
	}
	return S_OK;
*///
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::SetMenuSB(HMENU hmenuShared, HOLEMENU holemenuRes, HWND hwndActiveObject)
{
/*//
	if (!g_hMenu) {
		if (hwndActiveObject == m_hwnd) {
			g_hMenu = CreatePopupMenu();
			teCopyMenu(g_hMenu, hmenuShared, MF_ENABLED);
			return S_OK;
		}
	}
*///
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::RemoveMenusSB(HMENU hmenuShared)
{
/*//
	for (int nCount = GetMenuItemCount(hmenuShared); --nCount >= 0;) {
		DeleteMenu(hmenuShared, nCount, MF_BYPOSITION);
	}
	return S_OK;
*///
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::SetStatusTextSB(LPCWSTR lpszStatusText)
{
	g_szStatus[0] = 0x20;
	SetTimer(g_hwndMain, TET_Status, g_dwTickKey && (GetTickCount() - g_dwTickKey < 3000) ? 500 : 100, teTimerProc);
	return S_OK;
}

STDMETHODIMP CteShellBrowser::EnableModelessSB(BOOL fEnable)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::TranslateAcceleratorSB(LPMSG lpmsg, WORD wID)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::BrowseObject(PCUIDLIST_RELATIVE pidl, UINT wFlags)
{
	if (ILIsEqual(pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
		return ILIsEqual(m_pidl, pidl) ? S_FALSE : E_FAIL;
	}
	FolderItem *pid = NULL;
	if (wFlags & SBSP_REDIRECT) {
		if (m_pidl) {
			if (m_uLogIndex < m_ppLog.size()) {
				m_ppLog[m_uLogIndex]->QueryInterface(IID_PPV_ARGS(&pid));
			} else {
				VARIANT v;
				if (!GetVarPathFromFolderItem(m_pFolderItem, &v)) {
					GetVarPathFromIDList(&v, m_pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
				}
				pid = new CteFolderItem(&v);
				VariantClear(&v);
			}
			teILFreeClear(&m_pidl);
			wFlags &= ~(SBSP_RELATIVE | SBSP_REDIRECT);
		} else if (m_pFolderItem) {
			CteFolderItem *pid1 = NULL;
			if SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) {
				if (pid1->m_v.vt == VT_BSTR) {
					m_pFolderItem->QueryInterface(IID_PPV_ARGS(&pid));
					wFlags &= ~(SBSP_RELATIVE | SBSP_REDIRECT);
				} else if (pid1->m_dwUnavailable || m_dwUnavailable) {
					if (pid1->m_v.vt == VT_EMPTY && pid1->m_pidl) {
						GetVarPathFromFolderItem(m_pFolderItem, &pid1->m_v);
						teILFreeClear(&pid1->m_pidl);
						pid1->QueryInterface(IID_PPV_ARGS(&pid));
						wFlags &= ~(SBSP_RELATIVE | SBSP_REDIRECT);
					}
				}
				pid1->Release();
			}
		}
	}
	if (!pid) {
		GetFolderItemFromIDList(&pid, const_cast<LPITEMIDLIST>(pidl));
	}
	HRESULT hr = BrowseObject2(pid, wFlags);
	SafeRelease(&pid);
	return hr;
}

HRESULT CteShellBrowser::BrowseObject2(FolderItem *pid, UINT wFlags)
{
	HRESULT hr = E_FAIL;
	m_pTC->LockUpdate();
	try {
		CteShellBrowser *pSB = NULL;
		DWORD param[SB_Count];
		for (int i = SB_Count; i--;) {
			param[i] = m_param[i];
		}
		param[SB_DoFunc] = 1;
		hr = Navigate3(pid, wFlags, param, &pSB, NULL);
		SafeRelease(&pSB);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"BrowseObject2";
#endif
	}
	m_pTC->UnlockUpdate();
	return hr;
}

HRESULT CteShellBrowser::Navigate3(FolderItem *pFolderItem, UINT wFlags, DWORD *param, CteShellBrowser **ppSB, FolderItems *pFolderItems)
{
	HRESULT hr = E_FAIL;
	try {
		if ((wFlags & SBSP_NEWBROWSER) && teILIsBlank(m_pFolderItem) && teILIsBlank(m_pFolderItem1)) {
			wFlags &= ~SBSP_NEWBROWSER;
		}
		if (!(wFlags & SBSP_NEWBROWSER)) {
			hr = Navigate2(pFolderItem, wFlags, param, pFolderItems, m_pFolderItem, this);
			if (hr == E_ACCESSDENIED) {
				wFlags = (wFlags | SBSP_NEWBROWSER) & (~SBSP_ACTIVATE_NOFOCUS);
			}
		}
		if (hr != E_ABORT && ((wFlags & SBSP_NEWBROWSER) || !m_pFolderItem)) {
			CteShellBrowser *pSB = this;
			BOOL bNew = wFlags & SBSP_NEWBROWSER;
			if (bNew) {
				pSB = GetNewShellBrowser(m_pTC);
				*ppSB = pSB;
				teSysFreeString(&pSB->m_pTV->m_bsRoot);
				pSB->m_pTV->m_bsRoot = ::SysAllocString(m_pTV->m_bsRoot);
			}
			for (int i = SB_Count; i--;) {
				pSB->m_param[i] = param[i];
			}
			hr = pSB->Navigate2(pFolderItem, wFlags, param, pFolderItems, m_pFolderItem, this);
			if SUCCEEDED(hr) {
				if (bNew && (wFlags & SBSP_ACTIVATE_NOFOCUS) == 0) {
					TabCtrl_SetCurSel(m_pTC->m_hwnd, m_pTC->m_nIndex + 1);
					m_pTC->m_nIndex = TabCtrl_GetCurSel(m_pTC->m_hwnd);
					pSB->Show(TRUE, 0);
				} else {
					m_pTC->TabChanged(TRUE);
				}
			} else {
				pSB->Close(TRUE);
			}
		}
	} catch (...) {
		hr = E_FAIL;
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"Navigate3";
#endif
	}
	return hr;
}

BOOL CteShellBrowser::Navigate1(FolderItem *pFolderItem, UINT wFlags, FolderItems *pFolderItems, FolderItem *pPrevious, int nErrorHandling)
{
	BOOL bResult = FALSE;
	CteFolderItem *pid = NULL;
	if (pFolderItem) {
		if SUCCEEDED(pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid)) {
			if (!pid->m_pidl && !pid->m_pidlAlt) {
				if (pid->m_v.vt == VT_BSTR) {
					if ((tePathIsNetworkPath(pid->m_v.bstrVal) && tePathIsDirectory(pid->m_v.bstrVal, 100, 3) != S_OK) || (m_pShellView && pid->m_dwUnavailable)) {
						if (m_nUnload || g_nLockUpdate <= 1) {
							Navigate1Ex(pid->m_v.bstrVal, pFolderItems, wFlags, pPrevious, nErrorHandling);
						} else {
							m_nUnload = 9;
						}
						bResult = TRUE;
					}
				}
			}
			pid->Release();
		}
	}
	return bResult;
}

VOID CteShellBrowser::Navigate1Ex(LPOLESTR pstr, FolderItems *pFolderItems, UINT wFlags, FolderItem *pPrevious, int nErrorHandling)
{
	TEInvoke *pInvoke = new TEInvoke[1];
	QueryInterface(IID_PPV_ARGS(&pInvoke->pdisp));
	pInvoke->cRef = 1;
	pInvoke->cDo = 1;
	pInvoke->dispid = TE_METHOD + 0xfb07;//Navigate2Ex
	pInvoke->cArgs = 4;
	pInvoke->pv = GetNewVARIANT(4);
	pInvoke->wErrorHandling = g_nLockUpdate || m_nUnload == 2 || !g_bShowParseError ? 1 : nErrorHandling;
	pInvoke->wMode = 0;
	pInvoke->pidl = NULL;
	pInvoke->bsPath = ::SysAllocString(pstr);
	teSetLong(&pInvoke->pv[2], wFlags);
	teSetObject(&pInvoke->pv[1], pFolderItems);
	teSetObject(&pInvoke->pv[0], pPrevious);
	_beginthread(&threadParseDisplayName, 0, pInvoke);
}

HRESULT CteShellBrowser::Navigate2(FolderItem *pFolderItem, UINT wFlags, DWORD *param, FolderItems *pFolderItems, FolderItem *pPrevious, CteShellBrowser *pHistSB)
{
	RECT rc;
	LPITEMIDLIST pidl = NULL;
	HRESULT hr;

	SafeRelease(&m_pFolderItem1);
	if (!(wFlags & SBSP_RELATIVE)) {
		m_dwUnavailable = 0;
	}
	m_bRefreshLator = FALSE;
	m_nFolderSizeIndex = MAXINT;
	m_nLabelIndex = MAXINT;
	m_nLinkTargetIndex = MAXINT;
	m_bRegenerateItems = FALSE;
	m_nSizeFormat = -1;
	SafeRelease(&m_ppDispatch[SB_TotalFileSize]);
	m_ppDispatch[SB_TotalFileSize] = teCreateObj(TE_OBJECT, NULL);
	SafeRelease(&m_ppDispatch[SB_AltSelectedItems]);
	teSysFreeString(&m_bsAltSortColumn);
	CteFolderItem *pid1 = NULL;
	if (m_hwnd) {
		KillTimer(m_hwnd, (UINT_PTR)this);
	}
	if (m_bInit) {
		return E_FAIL;
	}
	for (int i = SB_Count; --i >= 0;) {
		m_param[i] = param[i];
		g_paramFV[i] = param[i];
	}
	m_uPrevLogIndex = m_uLogIndex;
	SaveFocusedItemToHistory();
	if ((wFlags & (SBSP_PARENT | SBSP_NAVIGATEBACK | SBSP_NAVIGATEFORWARD | SBSP_RELATIVE)) == 0) {
		CteFolderItem *pid1;
		if SUCCEEDED(pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) {
			if (pid1->m_pidl && pid1->m_v.vt == VT_EMPTY) {
				VARIANT_BOOL bFolder = VARIANT_FALSE;
				pid1->get_IsFolder(&bFolder);
				if (!bFolder) {
					if (::ILIsEqual(pid1->m_pidl, g_pidls[CSIDL_INTERNET])) {
						::ILRemoveLastID(pid1->m_pidl);
					} else if SUCCEEDED(pid1->get_Path(&pid1->m_v.bstrVal)) {
						pid1->m_v.vt = VT_BSTR;
						teILFreeClear(&pid1->m_pidl);
					}
				}
			}
			if (g_pOnFunc[TE_OnReplacePath] && pid1->m_v.vt == VT_BSTR) {
				VARIANTARG *pv = GetNewVARIANT(2);
				teSetObject(&pv[1], this);
				teSetSZ(&pv[0], pid1->m_v.bstrVal);
				VARIANT vResult;
				VariantInit(&vResult);
				Invoke4(g_pOnFunc[TE_OnReplacePath], &vResult, 2, pv);
				if (vResult.vt != VT_EMPTY) {
					VariantClear(&pid1->m_v);
					VariantCopy(&pid1->m_v, &vResult);
					VariantClear(&vResult);
				}
			}
			pid1->Release();
		}
	}
	hr = GetAbsPath(pFolderItem, wFlags, pFolderItems, pPrevious, pHistSB);
	if (hr == S_OK) {
		if (!teGetIDListFromObjectForFV(m_pFolderItem1, &pidl)) {
			hr = E_FAIL;
		}
		if (teILIsBlank(m_pFolderItem1)) {
			teILCloneReplace(&pidl, g_pidls[CSIDL_RESULTSFOLDER]);
		}
	}
	if (!m_dwUnavailable) {
		if SUCCEEDED(pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) {
			m_dwUnavailable = pid1->m_dwUnavailable;
			pid1->Release();
		}
	}
	if (m_dwUnavailable) {
		if (m_pShellView) {
			hr = S_FALSE;
		} else if (ILIsEmpty(pidl)) {
			teILCloneReplace(&pidl, g_pidls[CSIDL_RESULTSFOLDER]);
		}
	}
	if (hr != S_OK) {
		if (hr == S_FALSE) {
			if (!m_pFolderItem) {
				if SUCCEEDED(pFolderItem->QueryInterface(IID_PPV_ARGS(&m_pFolderItem))) {
					CteFolderItem *pid1;
					teQueryFolderItem(&m_pFolderItem, &pid1);
					pid1->MakeUnavailable();
					pid1->Release();
				}
			}
		}
		teCoTaskMemFree(pidl);
		return m_pidl ? S_OK : hr;
	}
	hr = S_FALSE;
	if (!pidl) {
		return S_OK;
	}
	IShellFolder2 *pSF2;
	LPCITEMIDLIST pidlPart;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF2), &pidlPart)) {
		SFGAOF sfAttr = SFGAO_LINK | SFGAO_FOLDER;
		if SUCCEEDED(pSF2->GetAttributesOf(1, &pidlPart, &sfAttr)) {
			if ((sfAttr & (SFGAO_LINK | SFGAO_FOLDER)) == SFGAO_LINK) {
				VARIANT v;
				VariantInit(&v);
				if (SUCCEEDED(pSF2->GetDetailsEx(pidlPart, &PKEY_Link_TargetParsingPath, &v))) {
					if (v.vt == VT_BSTR) {
						LPITEMIDLIST pidl2 = teILCreateFromPath(v.bstrVal);
						if (pidl2) {
							teILFreeClear(&pidl);
							pidl = pidl2;
							if (m_pFolderItem1) {
								m_pFolderItem1->Release();
								GetFolderItemFromIDList(&m_pFolderItem1, pidl);
							}
						}
					}
					VariantClear(&v);
				}
			}
		}
		pSF2->Release();
	}
	DWORD dwFrame = 0;
	UINT uViewMode = m_param[SB_ViewMode];
	m_nGroupByDelay = 0;
#ifdef _2000XP
	if (g_bUpperVista) {
#endif
		if (m_param[SB_Type] == CTRL_EB) {
			dwFrame = EBO_SHOWFRAMES;
		}
		//ExplorerBrowser
		if (m_pExplorerBrowser && !pFolderItems
#ifdef _USE_SHELLBROWSER
			&& (dwFrame || teGetFolderViewOptions(pidl, m_param[SB_ViewMode]) == FVO_DEFAULT)
#endif
		) {
			m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
			if (!ILIsEqual(pidl, g_pidls[CSIDL_RESULTSFOLDER]) || !ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
				if (GetShellFolder2(&pidl) == S_OK) {
					IFolderViewOptions *pOptions;
					if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
						pOptions->SetFolderViewOptions(FVO_VISTALAYOUT, teGetFolderViewOptions(pidl, m_param[SB_ViewMode]));
						pOptions->Release();
					}
					BrowseToObject();
					teCoTaskMemFree(pidl);
					return S_OK;
				}
			} else {
				m_bBeforeNavigate = TRUE;
				hr = OnNavigationPending2(pidl);
				if SUCCEEDED(hr) {
					RemoveAll();
					OnViewCreated(NULL);
					OnNavigationComplete2();
				}
				return hr;
			}
		}
#ifdef _2000XP
	}
#endif
	m_clrText = GetSysColor(COLOR_WINDOWTEXT);
	m_clrBk = GetSysColor(COLOR_WINDOW);
	m_clrTextBk = m_clrBk;
	m_bSetListColumnWidth = FALSE;
	teCoTaskMemFree(m_pidl);
	m_pidl = pidl;
	m_pFolderItem = m_pFolderItem1;
	m_pFolderItem1 = NULL;
	if (g_nLockUpdate > 1 || !m_pTC->m_bVisible || (wFlags & SBSP_ACTIVATE_NOFOCUS)) {
		m_nUnload = 1;
		//History / Management
		SetHistory(pFolderItems, wFlags);
		SetTabName();
		return S_OK;
	}
	//Previous display
	IShellView *pPreviousView = m_pShellView;
	m_pShellView = NULL;
	EXPLORER_BROWSER_OPTIONS dwflag;
	dwflag = static_cast<EXPLORER_BROWSER_OPTIONS>(m_param[SB_Options]);

	m_bIconSize = FALSE;

	if (param[SB_DoFunc]) {
		hr = OnBeforeNavigate(pPrevious, wFlags);
		if FAILED(hr) {
			if (m_pFolderItem == pFolderItem) {
				m_pFolderItem = NULL;
			}
			if (teILIsBlank(pPrevious) && Close(FALSE)) {
				return hr;
			}
		}
		if (hr != S_OK) {
			teILFreeClear(&m_pidl);
			if (pPrevious) {
				teGetIDListFromObject(pPrevious, &m_pidl);
				pPrevious->AddRef();
				SafeRelease(&m_pFolderItem);
				m_pFolderItem = pPrevious;
				SafeRelease(&m_pShellView);
				m_pShellView = pPreviousView;
				pPreviousView = NULL;
			}
			m_uLogIndex = m_uPrevLogIndex;
			return hr;
		}
	}
	if (!teILIsEqual(m_pFolderItem, pPrevious)) {
		InitFilter();
	}
	SetRectEmpty(&rc);
	teSysFreeString(&m_bsNextGroup);
	//History / Management
	SetHistory(pFolderItems, wFlags);
	m_bBeforeNavigate = TRUE;
	m_pTC->LockUpdate();
	try {
		DestroyView(m_param[SB_Type]);
		Show(FALSE, 2);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"DestroyView1";
#endif
	}
	try {
#ifdef _USE_SHELLBROWSER
		hr = E_FAIL;
		if (dwFrame || teGetFolderViewOptions(m_pidl, m_param[SB_ViewMode]) == FVO_DEFAULT) {
			//ExplorerBrowser
			hr = NavigateEB(dwFrame);
		}
		if FAILED(hr) {
			//ShellBrowser
			hr = NavigateSB(pPreviousView, pPrevious);
		}
#else
#ifdef _2000XP
		hr = g_bUpperVista ? NavigateEB(dwFrame) : NavigateSB(pPreviousView, pPrevious);
#else
		hr =  NavigateEB(dwFrame);
#endif
#endif
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"NavigateSB/EB";
#endif
	}
	m_pTC->UnlockUpdate();
	if (m_pTC && GetTabIndex() == m_pTC->m_nIndex) {
		m_nUnload = 0;
		Show(TRUE, 0);
	}
	if (hr != S_OK) {
		SafeRelease(&m_pShellView);
		m_pShellView = pPreviousView;

		teILFreeClear(&m_pidl);
		teGetIDListFromObject(pPrevious, &m_pidl);
		if (pPrevious) {
			pPrevious->AddRef();
			SafeRelease(&m_pFolderItem);
			m_pFolderItem = pPrevious;
		}
		m_uLogIndex = m_uPrevLogIndex;
		return hr;
	}

	if (pPreviousView) {
		pPreviousView->DestroyViewWindow();
		SafeRelease(&pPreviousView);
	}

	m_nOpenedType = m_param[SB_Type];
	if (!m_pExplorerBrowser) {
		OnViewCreated(NULL);
	}
	return S_OK;
}

HRESULT CteShellBrowser::NavigateEB(DWORD dwFrame)
{
	HRESULT hr = E_FAIL;
	RECT rc;
	try {
		if (::InterlockedIncrement(&m_nCreate) <= 1) {
			if (m_pExplorerBrowser) {
				DestroyView(CTRL_SB);
			}
			if (SUCCEEDED(teCreateInstance(CLSID_ExplorerBrowser, NULL, NULL, IID_PPV_ARGS(&m_pExplorerBrowser)))) {
				if (m_nForceViewMode != FVM_AUTO) {
					m_param[SB_ViewMode] = m_nForceViewMode;
					m_bAutoVM = FALSE;
					m_nForceViewMode = FVM_AUTO;
				}
				FOLDERSETTINGS fs = { !m_bAutoVM || ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER]) ? m_param[SB_ViewMode] : FVM_AUTO, (m_param[SB_FolderFlags] | FWF_USESEARCHFOLDER) & ~FWF_NOENUMREFRESH };
				if (SUCCEEDED(m_pExplorerBrowser->Initialize(m_pTC->m_hwndStatic, &rc, &fs))) {
					hr = GetShellFolder2(&m_pidl);
					if (hr == S_OK) {
						m_pExplorerBrowser->Advise(static_cast<IExplorerBrowserEvents *>(this), &m_dwEventCookie);
						IFolderViewOptions *pOptions;
						if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
							pOptions->SetFolderViewOptions(FVO_VISTALAYOUT, teGetFolderViewOptions(m_pidl, m_param[SB_ViewMode]));
							pOptions->Release();
						}
						m_pExplorerBrowser->SetOptions(static_cast<EXPLORER_BROWSER_OPTIONS>((m_param[SB_Options] & ~(EBO_SHOWFRAMES | EBO_NAVIGATEONCE | EBO_ALWAYSNAVIGATE)) | dwFrame | EBO_NOTRAVELLOG));
						IServiceProvider *pSP = new CteServiceProvider(static_cast<IShellBrowser *>(this), NULL);
						IUnknown_SetSite(m_pExplorerBrowser, pSP);
						pSP->Release();
/*///
						IFolderFilterSite *pFolderFilterSite;
						if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pFolderFilterSite))) {
							pFolderFilterSite->SetFilter(static_cast<IFolderFilter *>(this));
						}
*///
						try {
							hr = BrowseToObject();
						} catch (...) {
							g_nException = 0;
#ifdef _DEBUG
							g_strException = L"NavigateEB/BrowseToObject";
#endif
						}
					}
					if (hr == S_OK) {
						if (m_pShellView) {
							GetDefaultColumns();
						}
					}
				}
			}
		}
	} catch (...) {
		hr = E_FAIL;
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"NavigateEB";
#endif
	}
	InterlockedDecrement(&m_nCreate);
	return hr;
}

HRESULT CteShellBrowser::OnBeforeNavigate(FolderItem *pPrevious, UINT wFlags)
{
	HRESULT hr = S_OK;
	if (m_pShellView) {
		m_bViewCreated = TRUE;
		GetViewModeAndIconSize(TRUE);
		if (m_pExplorerBrowser) {
			m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
		}
	}
	m_bFiltered = FALSE;
#ifdef _2000XP
	m_bAutoVM = g_bUpperVista && (g_param[TE_Layout] & TE_AutoViewMode);
#else
	m_bAutoVM = (g_param[TE_Layout] & TE_AutoViewMode);
#endif
	if (g_pOnFunc[TE_OnBeforeNavigate]) {
		VARIANT vResult;
		VariantInit(&vResult);

		VARIANTARG *pv;
		pv = GetNewVARIANT(4);
		teSetObject(&pv[3], this);
		teSetObjectRelease(&pv[2], new CteMemory(4 * sizeof(int), &m_param[SB_ViewMode], 1, L"FOLDERSETTINGS"));
		teSetLong(&pv[1], wFlags);
		if (!teILIsEqual(m_pFolderItem, pPrevious)) {
			teSetObject(&pv[0], pPrevious);
		}
		DWORD nViewMode = m_param[SB_ViewMode];
		m_param[SB_ViewMode] = FVM_AUTO;
		m_bInit = TRUE;
		try {
			Invoke4(g_pOnFunc[TE_OnBeforeNavigate], &vResult, 4, pv);
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"OnBeforeNavigate";
#endif
		}
		m_bInit = FALSE;
		hr = GetIntFromVariantClear(&vResult);
		if (m_param[SB_ViewMode] == FVM_AUTO) {
			m_param[SB_ViewMode] = nViewMode;
		} else {
			m_bAutoVM = FALSE;
		}
	}
	if (hr == E_ACCESSDENIED || (hr == E_FAIL && m_nUnload == 2)) {
		if (teILIsEqual(m_pFolderItem, pPrevious)) {
			hr = S_OK;
		}
	}
	m_bViewCreated = SUCCEEDED(hr);
	return hr;
}

VOID CteShellBrowser::GetFocusedItem(LPITEMIDLIST *ppidl)
{
	IFolderView *pFV = NULL;
	int i = GetFolderViewAndItemCount(&pFV, SVGIO_SELECTION);
	if (i == 1) {
		if SUCCEEDED(pFV->GetFocusedItem(&i)) {
			teILFreeClear(ppidl);
			pFV->Item(i, ppidl);
		}
	}
	SafeRelease(&pFV);
}

VOID CteShellBrowser::SaveFocusedItemToHistory()
{
	if (m_uLogIndex < m_ppLog.size()) {
		CteFolderItem *pid1 = NULL;
		teQueryFolderItem(&m_ppLog[m_uLogIndex], &pid1);
		GetFocusedItem(&pid1->m_pidlFocused);
	}
}

VOID CteShellBrowser::FocusItem()
{
	if (!m_pShellView) {
		return;
	}
	CteFolderItem *pid;
	if (m_pFolderItem && !m_dwUnavailable && SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid))) {
		if (pid->m_pidlFocused) {
			teFixListState(m_hwndLV, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS | SVSI_DESELECTOTHERS | SVSI_SELECTIONMARK | SVSI_SELECT);
			SelectItemEx(pid->m_pidlFocused, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS | SVSI_DESELECTOTHERS | SVSI_SELECTIONMARK | SVSI_SELECT);
			teILFreeClear(&pid->m_pidlFocused);
		}
		m_dwUnavailable = pid->m_dwUnavailable;
		pid->Release();
	}
	IFolderView *pFV = NULL;
	if (GetFolderViewAndItemCount(&pFV, SVGIO_ALLVIEW)) {
		int nTop = -1;
		if (SUCCEEDED(pFV->GetFocusedItem(&nTop)) && nTop < 0) {
			pFV->SelectItem(0, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS | SVSI_SELECTIONMARK);
		}
		pFV->Release();
	}
	m_nFocusItem = 0;
}

HRESULT CteShellBrowser::GetAbsPath(FolderItem *pid, UINT wFlags, FolderItems *pFolderItems, FolderItem *pPrevious, CteShellBrowser *pHistSB)
{
	if (wFlags & SBSP_PARENT) {
		if (pPrevious) {
			if (teILGetParent(pPrevious, &m_pFolderItem1)) {
				if (Navigate1(m_pFolderItem1, wFlags & ~SBSP_PARENT, pFolderItems, pPrevious, 0)) {
					return S_FALSE;
				}
				CteFolderItem *pid1;
				teQueryFolderItem(&m_pFolderItem1, &pid1);
				teGetIDListFromObject(pPrevious, &pid1->m_pidlFocused);
				SafeRelease(&pid1);
			}
		} else {
			GetFolderItemFromIDList(&m_pFolderItem1, g_pidls[CSIDL_DESKTOP]);
		}
		if (teILIsEqual(pPrevious, m_pFolderItem1)) {
			return E_FAIL;
		}
		return S_OK;
	}
	if (wFlags & SBSP_NAVIGATEBACK) {
		UINT uLogIndex = pHistSB->m_uLogIndex;
		if (uLogIndex && uLogIndex < m_ppLog.size()) {
			if (m_pFolderItem) {
				if (m_pFolderItem != m_ppLog[uLogIndex]) {
					m_ppLog[uLogIndex]->Release();
					m_pFolderItem->QueryInterface(IID_PPV_ARGS(&m_ppLog[uLogIndex]));
				}
				SaveFocusedItemToHistory();
			}
			if SUCCEEDED(pHistSB->m_ppLog[--uLogIndex]->QueryInterface(IID_PPV_ARGS(&m_pFolderItem1))) {
				if (this == pHistSB) {
					m_uLogIndex = uLogIndex;
				}
				return S_OK;
			}
		}
		return E_FAIL;
	}
	if (wFlags & SBSP_NAVIGATEFORWARD) {
		UINT uLogIndex = pHistSB->m_uLogIndex;
		if (uLogIndex < pHistSB->m_ppLog.size() - 1) {
			if (m_pFolderItem) {
				if (m_pFolderItem != m_ppLog[uLogIndex]) {
					m_ppLog[uLogIndex]->Release();
					m_pFolderItem->QueryInterface(IID_PPV_ARGS(&m_ppLog[uLogIndex]));
				}
				SaveFocusedItemToHistory();
			}
			if SUCCEEDED(pHistSB->m_ppLog[++uLogIndex]->QueryInterface(IID_PPV_ARGS(&m_pFolderItem1))) {
				if (this == pHistSB) {
					m_uLogIndex = uLogIndex;
				}
				return S_OK;
			}
		}
		return E_FAIL;
	}
	LPITEMIDLIST pidl = NULL;
	if (Navigate1(pid, wFlags, pFolderItems, pPrevious, 2)) {
		return S_FALSE;
	}
	if (wFlags & SBSP_RELATIVE) {
		LPITEMIDLIST pidlPrevius = NULL;
		teGetIDListFromObjectForFV(pPrevious, &pidlPrevius);
		if (pidlPrevius && !ILIsEqual(pidlPrevius, g_pidls[CSIDL_RESULTSFOLDER]) && !ILIsEmpty(pidl)) {
			LPITEMIDLIST pidlFull = NULL;
			pidlFull = ILCombine(pidlPrevius, pidl);
			GetFolderItemFromIDList(&m_pFolderItem1, pidlFull);
			teCoTaskMemFree(pidlFull);
		} else if (pPrevious) {
			pPrevious->QueryInterface(IID_PPV_ARGS(&m_pFolderItem1));
		} else if (pidlPrevius) {
			GetFolderItemFromIDList(&m_pFolderItem1, pidlPrevius);
		} else {
			VARIANT v;
			teSetSZ(&v, PATH_BLANK);
			m_pFolderItem1 = new CteFolderItem(&v);
			VariantClear(&v);
		}
		teILFreeClear(&pidlPrevius);
		teILFreeClear(&pidl);
	} else if (pid) {
		pid->QueryInterface(IID_PPV_ARGS(&m_pFolderItem1));
	}
	return m_pFolderItem1 ? S_OK : E_FAIL;
}

VOID CteShellBrowser::Refresh(BOOL bCheck)
{
	m_bRefreshLator = FALSE;
	if (!m_dwUnavailable) {
		SaveFocusedItemToHistory();
		teDoCommand(this, m_hwnd, WM_NULL, 0, 0);//Save folder setings
	}
	SafeRelease(&m_ppDispatch[SB_TotalFileSize]);
	m_ppDispatch[SB_TotalFileSize] = teCreateObj(TE_OBJECT, NULL);
	SafeRelease(&m_ppDispatch[SB_AltSelectedItems]);
	if (bCheck) {
		VARIANT v;
		VariantInit(&v);
		if SUCCEEDED(m_pFolderItem->QueryInterface(IID_PPV_ARGS(&v.pdispVal))) {
			VARIANT vResult;
			v.vt = VT_DISPATCH;
			VariantInit(&vResult);
			teGetDisplayNameOf(&v, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &vResult);
			VariantClear(&v);
			if (vResult.vt == VT_BSTR) {
				if (!m_pShellView || tePathIsNetworkPath(vResult.bstrVal) || m_dwUnavailable) {
					TEInvoke *pInvoke = new TEInvoke[1];
					QueryInterface(IID_PPV_ARGS(&pInvoke->pdisp));
					pInvoke->dispid = TE_METHOD + 0xfb06;//Refresh2Ex
					pInvoke->cRef = 1;
					pInvoke->cDo = 1;
					pInvoke->cArgs = 1;
					pInvoke->pv = GetNewVARIANT(1);
					pInvoke->wErrorHandling = 1;
					pInvoke->wMode = 0;
					pInvoke->pidl = NULL;
					pInvoke->bsPath = ::SysAllocString(vResult.bstrVal);
					VariantClear(&vResult);
					_beginthread(&threadParseDisplayName, 0, pInvoke);
					return;
				}
			}
			VariantClear(&vResult);
		}
	}
	if (m_pShellView) {
		if (m_bVisible) {
			if (m_nUnload == 4) {
				m_nUnload = 0;
			}
			if (m_dwUnavailable || ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
				BrowseObject(NULL, SBSP_RELATIVE | SBSP_SAMEBROWSER);
				return;
			}
			if (GetFolderViewAndItemCount(NULL, SVGIO_ALLVIEW) == 0) {
				BSTR bs;
				if SUCCEEDED(teGetDisplayNameFromIDList(&bs, m_pidl, SHGDN_FORPARSING)) {
					BOOL bNetWork = tePathIsNetworkPath(bs);
					teSysFreeString(&bs);
					if (bNetWork) {
						Suspend(0);
						return;
					}
				}
			}
			m_bRefreshing = TRUE;
			m_bFiltered = FALSE;
			m_pTC->LockUpdate();
			try {
				m_pShellView->Refresh();
				InitFolderSize();
				SetFolderFlags(FALSE);
			} catch (...) {
				g_nException = 0;
#ifdef _DEBUG
				g_strException = L"Refresh";
#endif
			}
			m_pTC->UnlockUpdate();
			ArrangeWindow();
		} else if (m_nUnload == 0) {
			m_nUnload = 4;
		}
	}
}

BOOL CteShellBrowser::SetActive(BOOL bForce)
{
	if (!m_pTC || GetTabIndex() != m_pTC->m_nIndex) {
		return FALSE;
	}
	if (m_hwndAlt) {
		BringWindowToTop(m_hwndAlt);
		SetFocus(m_hwndAlt);
		return TRUE;
	}
	HWND hwnd = GetFocus();
	if (!bForce) {
		if (hwnd) {
			if (IsChild(g_pWebBrowser->m_hwndBrowser, hwnd) || (m_pTV && IsChild(m_pTV->m_hwnd, hwnd))) {
				return FALSE;
			}
			CHAR szClassA[MAX_CLASS_NAME];
			GetClassNameA(hwnd, szClassA, MAX_CLASS_NAME);
			if (lstrcmpA(szClassA, WC_TREEVIEWA) == 0) {
				return FALSE;
			}
		}
	}
	if (m_pShellView) {
		BYTE lpKeyState[256];
		GetKeyboardState(lpKeyState);
		BYTE ksShift = lpKeyState[VK_SHIFT];
		if (ksShift & 0x80) {
			lpKeyState[VK_SHIFT] &= ~0x80;
			SetKeyboardState(lpKeyState);
		}
		if (IsChild(m_hwnd, hwnd)) {
			SetFocus(m_hwnd);
		}
		try {
			m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
		} catch (...) {
			m_pShellView = NULL;
			Suspend(0);
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"SetActive";
#endif
		}
		CheckChangeTabSB(m_hwnd);
		if (ksShift & 0x80 && GetAsyncKeyState(VK_SHIFT) < 0) {
			GetKeyboardState(lpKeyState);
			lpKeyState[VK_SHIFT] = ksShift;
			SetKeyboardState(lpKeyState);
		}
	}
	return TRUE;
}

VOID CteShellBrowser::SetTitle(BSTR szName, int nIndex)
{
	TC_ITEM tcItem;
	BSTR bsText = ::SysAllocStringLen(NULL, MAX_PATH);
	BSTR bsOldText = ::SysAllocStringLen(NULL, MAX_PATH);
	bsOldText[0] = NULL;
	tcItem.pszText = bsOldText;
	tcItem.mask = TCIF_TEXT;
	tcItem.cchTextMax = MAX_PATH;
	int nCount = ::SysStringLen(szName);
	if (nCount >= MAX_PATH) {
		nCount = MAX_PATH - 1;
	}
	int j = 0;
	for (int i = 0; i < nCount; ++i) {
		bsText[j++] = szName[i];
		if (szName[i] == '&') {
			bsText[j++] = '&';
		}
	}
	bsText[j] = NULL;
	TabCtrl_GetItem(m_pTC->m_hwnd, nIndex, &tcItem);
	if (lstrcmpi(bsText, bsOldText)) {
		tcItem.pszText = bsText;
		TabCtrl_SetItem(m_pTC->m_hwnd, nIndex, &tcItem);
		ArrangeWindow();
	}
	SysFreeString(bsOldText);
	SysFreeString(bsText);
}

VOID CteShellBrowser::GetShellFolderView()
{
	teUnadviseAndRelease(m_pdisp, DIID_DShellFolderViewEvents, &m_dwCookie);
	if (m_pShellView && SUCCEEDED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&m_pdisp)))) {
		teAdvise(m_pdisp, DIID_DShellFolderViewEvents, static_cast<IDispatch *>(this), &m_dwCookie);
	} else {
		m_pdisp = NULL;
	}
}

VOID CteShellBrowser::GetFocusedIndex(int *piItem)
{
#ifdef _W2000
	if (g_bIs2000) {
		//Windows 2000
		*piItem = ListView_GetNextItem(m_hwndLV, -1, LVNI_ALL | LVNI_FOCUSED);
		return;
	}
#endif
	*piItem = -1;
	if (!m_pShellView) {
		return;
	}
	IFolderView *pFV;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
		pFV->GetFocusedItem(piItem);
		pFV->Release();
	}
}

VOID CteShellBrowser::SetFolderFlags(BOOL bGetIconSize)
{
	if (!m_pShellView) {
		return;
	}
	IFolderView2 *pFV2;
	if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2)))) {
		DWORD dwMask;
		pFV2->GetCurrentFolderFlags(&dwMask);
		dwMask = (dwMask ^ m_param[SB_FolderFlags]) & (~(FWF_NOENUMREFRESH | FWF_USESEARCHFOLDER));
		if (dwMask) {
			pFV2->SetCurrentFolderFlags(dwMask, m_param[SB_FolderFlags]);
		}
		SafeRelease(&pFV2);
#ifdef _2000XP
	} else {
		ListView_SetExtendedListViewStyleEx(m_hwndLV, LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES, HIWORD(m_param[SB_FolderFlags]));
#endif
	}
	GetViewModeAndIconSize(bGetIconSize);
	if (m_hwndLV) {
		FixColumnEmphasis();
		ListView_SetTextBkColor(m_hwndLV, m_clrTextBk);
		ListView_SetTextColor(m_hwndLV, m_clrText);
		ListView_SetBkColor(m_hwndLV, m_clrBk);
	}
}

VOID CteShellBrowser::InitFolderSize()
{
	m_nFolderSizeIndex = MAXINT;
	m_nLabelIndex = MAXINT;
	m_nSizeIndex = MAXINT;
	m_nLinkTargetIndex = MAXINT;
	if (m_hwndLV) {
		HWND hHeader = ListView_GetHeader(m_hwndLV);
		if (hHeader) {
			UINT nHeader = Header_GetItemCount(hHeader);
			BSTR bsTotalFileSize = tePSGetNameFromPropertyKeyEx(PKEY_TotalFileSize, 0, this);
			BSTR bsLabel = g_pOnFunc[TE_Labels] ? tePSGetNameFromPropertyKeyEx(PKEY_Contact_Label, 0, this) : NULL;
			BSTR bsSize = tePSGetNameFromPropertyKeyEx(PKEY_Size, 0, this);
			BSTR bsLinkTarget = tePSGetNameFromPropertyKeyEx(PKEY_Link_TargetParsingPath, 0, this);
			WCHAR szText[MAX_COLUMN_NAME_LEN];
			HD_ITEM hdi = { HDI_TEXT | HDI_FORMAT, 0, szText, NULL, MAX_COLUMN_NAME_LEN };
			if (!m_pDTColumns.empty()) {
				::ZeroMemory(&m_pDTColumns[0], sizeof(UINT) * m_pDTColumns.size());
			}
			for (size_t n = m_ppColumns.size(); n;) {
				SafeRelease(&m_ppColumns[--n]);
			}
			TEmethodW *pColumns = new TEmethodW[nHeader + 1];
			VARIANT v;
			VariantInit(&v);
			for (int i = nHeader; --i >= 0;) {
				hdi.mask = HDI_TEXT | HDI_FORMAT;
				Header_GetItem(hHeader, i, &hdi);
				int fmt = hdi.fmt;
				pColumns[i].name = ::SysAllocString(szText);
				if (szText[0]) {
					if (m_ppDispatch[SB_ColumnsReplace]) {
						teGetProperty(m_ppDispatch[SB_ColumnsReplace], szText, &v);
					}
					if (v.vt == VT_EMPTY && g_pOnFunc[TE_ColumnsReplace]) {
						teGetProperty(g_pOnFunc[TE_ColumnsReplace], szText, &v);
					}
					if (v.vt != VT_EMPTY) {
						if (m_ppColumns.size() < (size_t)i + 1) {
							m_ppColumns.resize(i + 1, NULL);
						}
						GetDispatch(&v, &m_ppColumns[i]);
						VariantClear(&v);
						if (m_ppColumns[i]) {
							teGetProperty(m_ppColumns[i], L"fmt", &v);
							if (v.vt != VT_EMPTY) {
								int fmt = GetIntFromVariantClear(&v);
								if ((fmt ^ hdi.fmt) & HDF_JUSTIFYMASK) {
									hdi.fmt = (hdi.fmt & ~HDF_JUSTIFYMASK) | fmt;
								}
							}
						}
					}
#ifdef _2000XP
					if (g_bUpperVista) {
#endif
						if (lstrcmpi(szText, bsTotalFileSize) == 0) {
							m_nFolderSizeIndex = i;
							hdi.fmt |= HDF_RIGHT;
						}
						if (lstrcmpi(szText, bsLabel) == 0) {
							m_nLabelIndex = i;
						}
						if (lstrcmpi(szText, bsLinkTarget) == 0) {
							m_nLinkTargetIndex = i;
						}
#ifdef _2000XP
					}
#endif
					if (lstrcmpi(szText, bsSize) == 0) {
						m_nSizeIndex = i;
					}
				}
				if (fmt != hdi.fmt) {
					hdi.mask = HDI_FORMAT;
					Header_SetItem(hHeader, i, &hdi);
				}
			}

			if (g_bsDateTimeFormat && m_pSF2) {
				int *piColumns = SortTEMethodW(pColumns, nHeader);
				SHELLDETAILS sd;
				for (UINT i = 0; m_pSF2->GetDetailsOf(NULL, i, &sd) == S_OK && i < MAX_COLUMNS; ++i) {
					BSTR bs;
					if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs)) {
						int nIndex = teBSearchW(pColumns, nHeader, piColumns, bs);
						if (nIndex >= 0) {
							if (m_pDTColumns.size() < (size_t)nIndex + 1) {
								m_pDTColumns.resize(nIndex + 1, 0);
							}
							m_pDTColumns[nIndex] = i;
						}
						::SysFreeString(bs);
					}
				}
				delete [] piColumns;
			}
			for (int i = nHeader; --i >= 0;) {
				teSysFreeString(&pColumns[i].name);
			}
			delete pColumns;
			teSysFreeString(&bsLinkTarget);
			teSysFreeString(&bsSize);
			teSysFreeString(&bsLabel);
			::SysFreeString(bsTotalFileSize);
			InvalidateRect(m_hwndLV, NULL, FALSE);
		}
	}
}

VOID CteShellBrowser::SetSize(LPCITEMIDLIST pidl, LPWSTR szText, int cch)
{
	if (m_pSF2) {
		VARIANT v;
		VariantInit(&v);
		WIN32_FIND_DATA wfd = { 0 };
		HRESULT hr = teSHGetDataFromIDList(m_pSF2, pidl, SHGDFIL_FINDDATA, &wfd, sizeof(WIN32_FIND_DATA));
		if SUCCEEDED(hr) {
			if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
				if (ILIsEqual(m_pidl, g_pidls[CSIDL_BITBUCKET])) {
					hr = E_FAIL;
				}
			} else {
				ULARGE_INTEGER uli;
				uli.HighPart = wfd.nFileSizeHigh;
				uli.LowPart = wfd.nFileSizeLow;
				teStrFormatSize(GetSizeFormat(), uli.QuadPart, szText, cch);
			}
		}
		if FAILED(hr) {
			LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidl);// for Search folder (search-ms:) and Folder in Recycle bin
			IShellFolder2 *pSF2;
			LPCITEMIDLIST pidlPart;
			if SUCCEEDED(SHBindToParent(pidlFull, IID_PPV_ARGS(&pSF2), &pidlPart)) {
				if (SUCCEEDED(pSF2->GetDetailsEx(pidlPart, &PKEY_Size, &v))) {
					teStrFormatSize(GetSizeFormat(), GetLLFromVariant(&v), szText, cch);
					VariantClear(&v);
				}
				pSF2->Release();
			}
			teCoTaskMemFree(pidlFull);
		}
	}
}

static void threadFolderSize(void *args)
{
	::CoInitialize(NULL);
	try {
		TEFS *pFS;
		while (PopFolderSizeList(&pFS)) {
			IDispatch *pTotalFileSize;
			CoGetInterfaceAndReleaseStream(pFS->pStrmTotalFileSize, IID_PPV_ARGS(&pTotalFileSize));
			if (g_bMessageLoop) {
				VARIANT v;
				teSetLL(&v, teGetFolderSize(pFS->bsPath, MAXINT, pFS->pdwSessionId, pFS->dwSessionId));
				if (*pFS->pdwSessionId == pFS->dwSessionId) {
					tePutProperty(pTotalFileSize, pFS->bsPath, &v);
					if (pFS->hwnd) {
						InvalidateRect(pFS->hwnd, NULL, FALSE);
					}
					SetTimer(g_hwndMain, TET_Status, 100, teTimerProc);
				}
				VariantClear(&v);
			}
			SafeRelease(&pTotalFileSize);
			::SysFreeString(pFS->bsPath);
			delete [] pFS;
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"threadFolderSize";
#endif
	}
	--g_nCountOfThreadFolderSize;
	::CoUninitialize();
	::_endthread();
}

VOID CteShellBrowser::SetFolderSize(IShellFolder2 *pSF2, LPCITEMIDLIST pidl, LPWSTR szText, int cch)
{
	if (pSF2) {
		VARIANT v;
		VariantInit(&v);
		WIN32_FIND_DATA wfd;
		teSHGetDataFromIDList(pSF2, pidl, SHGDFIL_FINDDATA, &wfd, sizeof(WIN32_FIND_DATA));
		if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
			BSTR bsPath;
			teGetDisplayNameBSTR(pSF2, pidl, SHGDN_FORPARSING, &bsPath);
			teGetProperty(m_ppDispatch[SB_TotalFileSize], bsPath, &v);
			if (v.vt == VT_BSTR) {
				if (v.bstrVal) {
					lstrcpyn(szText, v.bstrVal, cch);
				}
			} else if (v.vt != VT_EMPTY) {
				teStrFormatSize(GetSizeFormat(), GetLLFromVariant(&v), szText, cch);
			} else if (m_ppDispatch[SB_TotalFileSize]) {
				v.vt = VT_BSTR;
				v.bstrVal = NULL;
				tePutProperty(m_ppDispatch[SB_TotalFileSize], bsPath, &v);
				TEFS *pFS = new TEFS[1];
				pFS->bsPath = ::SysAllocString(bsPath);
				pFS->pdwSessionId = &m_dwSessionId;
				pFS->dwSessionId = m_dwSessionId;
				pFS->hwnd = m_hwndLV;
				CoMarshalInterThreadInterfaceInStream(IID_IDispatch, m_ppDispatch[SB_TotalFileSize], &pFS->pStrmTotalFileSize);
				PushFolderSizeList(pFS);
				if (g_nCountOfThreadFolderSize < 8) {
					++g_nCountOfThreadFolderSize;
					_beginthread(&threadFolderSize, 0, NULL);
				}
			}
			teSysFreeString(&bsPath);
		} else if (SUCCEEDED(pSF2->GetDetailsEx(pidl, &PKEY_Size, &v))) {
			teStrFormatSize(GetSizeFormat(), GetLLFromVariant(&v), szText, cch);
		}
		VariantClear(&v);
	}
}

VOID CteShellBrowser::SetLabel(LPCITEMIDLIST pidl, LPWSTR szText, int cch)
{
	BSTR bs;
	VARIANT v;
	VariantInit(&v);
	if SUCCEEDED(teGetDisplayNameBSTR(m_pSF2, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &bs)) {
		if (::SysStringLen(bs)) {
			if (g_bLabelsMode) {
				teGetProperty(g_pOnFunc[TE_Labels], bs, &v);
			} else {
				VARIANTARG *pv = GetNewVARIANT(1);
				teSetBSTR(&pv[0], &bs, -1);
				Invoke4(g_pOnFunc[TE_Labels], &v, 1, pv);
			}
			if (v.vt == VT_BSTR) {
				lstrcpyn(szText, v.bstrVal, cch);
			}
			::VariantClear(&v);
		}
		teSysFreeString(&bs);
	}
}

BOOL CteShellBrowser::ReplaceColumns(IDispatch *pdisp, NMLVDISPINFO *lpDispInfo)
{
	BOOL bResult = FALSE;
	IFolderView *pFV;
	LPITEMIDLIST pidl;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
		if SUCCEEDED(pFV->Item(lpDispInfo->item.iItem, &pidl)) {
			VARIANTARG *pv = GetNewVARIANT(3);
			teSetObject(&pv[2], this);
			LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidl);
			teSetIDListRelease(&pv[1], &pidlFull);
			teILFreeClear(&pidl);
			teSetSZ(&pv[0], lpDispInfo->item.pszText);
			VARIANT v;
			VariantInit(&v);
			Invoke4(pdisp, &v, 3, pv);
			if (v.vt == VT_BSTR && ::SysStringLen(v.bstrVal)) {
				lstrcpyn(lpDispInfo->item.pszText, v.bstrVal, lpDispInfo->item.cchTextMax);
				bResult = TRUE;
			}
			VariantClear(&v);
		}
		pFV->Release();
	}
	return bResult;
}

VOID CteShellBrowser::GetViewModeAndIconSize(BOOL bGetIconSize)
{
	if (!m_pShellView || (m_dwUnavailable && !m_bCheckLayout)) {
		return;
	}
	FOLDERSETTINGS fs;
	UINT uViewMode = m_param[SB_ViewMode];
	int iImageSize = m_param[SB_IconSize];
	if SUCCEEDED(m_pShellView->GetCurrentInfo(&fs)) {
		uViewMode = fs.ViewMode;
	}
	if (bGetIconSize || uViewMode != m_param[SB_ViewMode]) {
		IFolderView2 *pFV2;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
			pFV2->GetViewModeAndIconSize((FOLDERVIEWMODE *)&uViewMode, &iImageSize);
			pFV2->Release();
#ifdef _2000XP
		} else {
			if (uViewMode == FVM_ICON) {
				iImageSize = 32;
			} else if (uViewMode == FVM_TILE) {
				iImageSize = 48;
			} else if (uViewMode == FVM_THUMBNAIL) {
				iImageSize = 96;
				HKEY hKey;
				if (RegOpenKeyExA(HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
					DWORD dwSize = sizeof(iImageSize);
					RegQueryValueExA(hKey, "ThumbnailSize", NULL, NULL, (LPBYTE)&iImageSize, &dwSize);
					RegCloseKey(hKey);
				}
			} else {
				iImageSize = 16;
			}
#endif
		}
		m_param[SB_IconSize] = iImageSize;
		m_param[SB_ViewMode] = uViewMode;
		if (m_bCheckLayout && !m_bViewCreated) {
			FOLDERVIEWOPTIONS fvo = teGetFolderViewOptions(m_pidl, uViewMode);
			if (m_pExplorerBrowser) {
				IFolderViewOptions *pOptions;
				if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
					FOLDERVIEWOPTIONS fvo0 = fvo;
					pOptions->GetFolderViewOptions(&fvo0);
					m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
					if ((fvo ^ fvo0 & FVO_VISTALAYOUT) || (m_param[SB_Options] & EBO_SHOWFRAMES) != (m_param[SB_Type] == CTRL_EB ? EBO_SHOWFRAMES : 0)) {
						m_bCheckLayout = FALSE;
					}
					pOptions->Release();
				}
			} else if (m_param[SB_Type] == CTRL_EB || fvo == FVO_DEFAULT) {
				m_bCheckLayout = FALSE;
			}
			if (!m_bCheckLayout) {
				m_nForceViewMode = uViewMode;
				Suspend(4);
			}
			m_bCheckLayout = FALSE;
		}
	}
}

VOID CteShellBrowser::GetVariantPath(FolderItem **ppFolderItem, FolderItems **ppFolderItems, VARIANT *pv)
{
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(ppFolderItems))) {
			VARIANT v;
			GetFolderItemFromFolderItems(ppFolderItem, *ppFolderItems, (Invoke5(*ppFolderItems, DISPID_TE_INDEX, DISPATCH_PROPERTYGET, &v, 0, NULL) == S_OK) ? GetIntFromVariantClear(&v) : 0);
			return;
		}
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(ppFolderItem))) {
			return;
		}
	}
	*ppFolderItem = new CteFolderItem(pv);
}

VOID CteShellBrowser::SetHistory(FolderItems *pFolderItems, UINT wFlags)
{
	if (pFolderItems) {
		UINT uLogIndex = 0;
		VARIANT v;
		if (Invoke5(pFolderItems, DISPID_TE_INDEX, DISPATCH_PROPERTYGET, &v, 0, NULL) == S_OK) {
			uLogIndex = GetIntFromVariantClear(&v);
		}
		while (!m_ppLog.empty()) {
			m_ppLog.back()->Release();
			m_ppLog.pop_back();
		}
		if SUCCEEDED(pFolderItems->get_Count(&v.lVal)) {
			v.vt = VT_I4;
			while (--v.lVal >= 0) {
				FolderItem *pid;
				if SUCCEEDED(pFolderItems->Item(v, &pid)) {
					m_ppLog.push_back(pid);
				}
			}
		}
		m_uLogIndex = UINT(m_ppLog.size()) - 1 - uLogIndex;
		SafeRelease(&pFolderItems);
	} else if ((wFlags & (SBSP_NAVIGATEBACK | SBSP_NAVIGATEFORWARD | SBSP_WRITENOHISTORY)) == 0) {
		if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
			if (!m_pFolderItem) {
				return;
			}
			BOOL bInvalid = TRUE;
			CteFolderItem *pid1 = NULL;
			if SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) {
				if (pid1->m_v.vt == VT_BSTR) {
					bInvalid = FALSE;
				}
				pid1->Release();
			}
			if (bInvalid) {
				return;
			}
		}
		if (m_uLogIndex < m_ppLog.size() && (teILIsEqual(m_pFolderItem, m_ppLog[m_uLogIndex]) || teILIsBlank(m_ppLog[m_uLogIndex]))) {
			if (m_pFolderItem != m_ppLog[m_uLogIndex]) {
				m_ppLog[m_uLogIndex]->Release();
				m_pFolderItem->QueryInterface(IID_PPV_ARGS(&m_ppLog[m_uLogIndex]));
			}
			return;
		}
		while (m_uLogIndex + 1 < m_ppLog.size() && !m_ppLog.empty()) {
			m_ppLog.back()->Release();
			m_ppLog.pop_back();
		}
		while (m_ppLog.size() >= MAX_HISTORY) {
			m_ppLog.front()->Release();
			m_ppLog.erase(m_ppLog.begin());
		}
		if (!m_pFolderItem && m_pidl) {
			GetFolderItemFromIDList(&m_pFolderItem, m_pidl);
		}
		if (m_pFolderItem) {
			FolderItem *pid;
			if SUCCEEDED(m_pFolderItem->QueryInterface(IID_PPV_ARGS(&pid))) {
				if (g_dwTickMount && GetTickCount() - g_dwTickMount < 500) { //#248
					pid->AddRef();
				}
				m_uLogIndex = UINT(m_ppLog.size());
				m_ppLog.push_back(pid);
			}
		}
	}
	g_dwTickMount = 0;
}

STDMETHODIMP CteShellBrowser::GetViewStateStream(DWORD grfMode, IStream **ppStrm)
{
//	return SHCreateStreamOnFileEx(L"", STGM_READWRITE | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, TRUE, NULL, ppStrm);
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::GetControlWindow(UINT id, HWND *lphwnd)
{
	if (lphwnd) {
		*lphwnd = (id == FCW_STATUS) ? m_hwnd : NULL;
		return *lphwnd ? S_OK : E_FAIL;
	}
	return E_POINTER;
}

STDMETHODIMP CteShellBrowser::SendControlMsg(UINT id, UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT *pret)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::QueryActiveShellView(IShellView **ppshv)
{
	return m_pShellView ? m_pShellView->QueryInterface(IID_PPV_ARGS(ppshv)) : E_FAIL;
}

STDMETHODIMP CteShellBrowser::OnViewWindowActive(IShellView *ppshv)
{
	if (m_pTC) {
		CheckChangeTabTC(m_pTC->m_hwnd);
	}
/*/// For check
	BSTR bsPath;
	if SUCCEEDED(teGetDisplayNameFromIDList(&bsPath, m_pidl, SHGDN_FORPARSING)) {
		SetCurrentDirectory(bsPath);
		::SysFreeString(bsPath);
	}
//*/
	return S_OK;
}

STDMETHODIMP CteShellBrowser::SetToolbarItems(LPTBBUTTONSB lpButtons, UINT nButtons, UINT uFlags)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::OnDefaultCommand(IShellView *ppshv)
{
	return DoFunc(TE_OnDefaultCommand, this, S_FALSE);
}

STDMETHODIMP CteShellBrowser::OnStateChange(IShellView *ppshv, ULONG uChange)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::IncludeObject(IShellView *ppshv, LPCITEMIDLIST pidl)
{
	return IncludeObject2(m_pSF2, pidl, NULL);
}

//ICommDlgBrowser2
STDMETHODIMP CteShellBrowser::Notify(IShellView *ppshv, DWORD dwNotifyType)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::GetDefaultMenuText(IShellView *ppshv, LPWSTR pszText, int cchMax)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::GetViewFlags(DWORD *pdwFlags)
{
	*pdwFlags = m_param[SB_ViewFlags] & CDB2GVF_SHOWALLFILES;
	if (teCompareSFClass(m_pSF2, &CLSID_LibraryFolder) || teCompareSFClass(m_pSF2, &CLSID_DBFolder)) {
		*pdwFlags |= CDB2GVF_NOINCLUDEITEM;
	}
	return S_OK;
}
/*/// ICommDlgBrowser3
STDMETHODIMP CteShellBrowser::OnColumnClicked(IShellView *ppshv, int iColumn)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::GetCurrentFilter(LPWSTR pszFileSpec, int cchFileSpec)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::OnPreViewCreated(IShellView *ppshv)
{
	return S_OK;
}
//*/
/*/// IFolderFilter
STDMETHODIMP CteShellBrowser::ShouldShow(IShellFolder *psf, PCIDLIST_ABSOLUTE pidlFolder, PCUITEMID_CHILD pidlItem)
{
	return IncludeObject(NULL, pidlItem);
}

STDMETHODIMP CteShellBrowser::GetEnumFlags(IShellFolder *psf, PCIDLIST_ABSOLUTE pidlFolder, HWND *phwnd, DWORD *pgrfFlags)
{
	return S_OK;
}
//*/

STDMETHODIMP CteShellBrowser::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteShellBrowser::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	HRESULT hr = teGetDispId(methodSB, _countof(methodSB), g_maps[MAP_SB], *rgszNames, rgDispId, TRUE);
	if (hr != S_OK && m_pdisp) {
		hr = m_pdisp->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
		if (hr == S_OK && lstrcmpi(*rgszNames, L"SortColumns") == 0) {
			m_dispidSortColumns = *rgDispId;
		}
	}
	if (hr != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

VOID CteShellBrowser::SetColumnsStr(BSTR bsColumns)
{
	int nAllWidth = 0;
	int nArgs = 0;
	LPTSTR *lplpszArgs = NULL;
	if (bsColumns && bsColumns[0]) {
		lplpszArgs = CommandLineToArgvW(bsColumns, &nArgs);
		nArgs /= 2;
	}
	int nAlloc = nArgs >= (int)m_pDefultColumns.size() ? nArgs + 1 : m_pDefultColumns.size();
	PROPERTYKEY *pPropKey0 = new PROPERTYKEY[nAlloc];
	PROPERTYKEY *pPropKey = &pPropKey0[1];
	TEmethodW *methodArgs0 = new TEmethodW[nAlloc];
	TEmethodW *methodArgs = &methodArgs0[1];
#ifdef _2000XP
	for (int i = nArgs; --i >= 0;) {
		methodArgs[i].name = NULL;
	}
#endif
	BOOL bNoName = TRUE;
	int nCount = 0;
	for (int i = 0; i < nArgs; ++i) {
		if SUCCEEDED(teGetPropertyKeyFromName(m_pSF2, lplpszArgs[i * 2], &pPropKey[nCount])) {
			BOOL bExist = FALSE;
			for (int j = nCount; --j >= 0;) {
				if (IsEqualPropertyKey(pPropKey[nCount], pPropKey[j])) {
					bExist = TRUE;
					break;
				}
			}
			if (bExist) {
				continue;
			}
#ifdef _2000XP
			if (!g_bUpperVista) {
				methodArgs[i].name = tePSGetNameFromPropertyKeyEx(pPropKey[nCount], 0, this);
			}
#endif
			if (IsEqualPropertyKey(pPropKey[nCount], PKEY_ItemNameDisplay)) {
				bNoName = FALSE;
			}
			int n;
			if (StrToIntEx(lplpszArgs[i * 2 + 1], STIF_DEFAULT, &n) && n < 32768) {
				nAllWidth += n;
			} else {
				n = -1;
			}
			if (n < 0) {
				nAllWidth += 100;
			}
			methodArgs[nCount++].id = n;
		}
	}
	if (bNoName) {
		pPropKey = pPropKey0;
		methodArgs = methodArgs0;
		if (nCount) {
			pPropKey[0] = PKEY_ItemNameDisplay;
#ifdef _2000XP
			if (!g_bUpperVista) {
				methodArgs[0].name = tePSGetNameFromPropertyKeyEx(PKEY_ItemNameDisplay, 0, this);
			}
#endif
			methodArgs[0].id = -1;
			++nCount;
			nAllWidth += 100;
		}
	}
	if (m_pShellView && (nCount == 0 || nAllWidth > nCount + 4)) {
		IColumnManager *pColumnManager;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager))) {
			//Default Columns
			if (nCount == 0) {
				pPropKey = pPropKey0;
				methodArgs = methodArgs0;
				nCount = m_pDefultColumns.size();
				for (int i = nCount; i--;) {
					pPropKey[i] = m_pDefultColumns[i];
					methodArgs[i].id = -1;
				}
			}
			if SUCCEEDED(pColumnManager->SetColumns(pPropKey, nCount)) {
				CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_WIDTH };
				for (int i = nCount; i--;) {
					cmci.uWidth = methodArgs[i].id;
					pColumnManager->SetColumnInfo(pPropKey[i], &cmci);
				}
			}
			pColumnManager->Release();
#ifdef _2000XP
		} else if (!g_bUpperVista) {
			UINT uCount = m_nColumns;
			SHELLDETAILS sd;
			int nWidth;
			if (m_pSF2) {
				for (int i = uCount; --i >= 0;) {
					m_pSF2->GetDefaultColumnState(i, &m_pColumns[i].csFlags);
					m_pColumns[i].csFlags &= ~SHCOLSTATE_ONBYDEFAULT;
					m_pColumns[i].nWidth = -3;
				}
				m_pColumns[uCount - 1].csFlags = SHCOLSTATE_TYPE_STR;
				m_pColumns[uCount - 2].csFlags = SHCOLSTATE_TYPE_STR;
				m_pColumns[0].csFlags |= SHCOLSTATE_ONBYDEFAULT;
				int *piArgs = SortTEMethodW(methodArgs, nCount);
				BSTR bs = GetColumnsStr(FALSE);
				if (bs && bs[0]) {
					int nCur;
					LPTSTR *lplpszColumns = CommandLineToArgvW(bs, &nCur);
					::SysFreeString(bs);
					nCur /= 2;
					BOOL bDiff = nCount != nCur || nCur == 0;
					if (!bDiff) {
						TEmethodW *methodColumns = new TEmethodW[nCur];
						for (int i = nCur; --i >= 0;) {
							methodColumns[i].name = lplpszColumns[i * 2];
						}
						int *piCur = SortTEMethodW(methodColumns, nCur);
						for (int i = nCur; --i >= 0;) {
							if (lstrcmpi(methodArgs[piArgs[i]].name, methodColumns[piCur[i]].name)) {
								bDiff = TRUE;
								break;
							}
						}
						delete [] piCur;
						delete [] methodColumns;
					}
					LocalFree(lplpszColumns);

					if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
						if (GetFolderViewAndItemCount(NULL, SVGIO_ALLVIEW)) {
							bDiff = FALSE;
						}
					}
					if (bDiff) {
						int nIndex;
						BOOL bDefault = TRUE;
						if (nCount) {
							for (UINT i = 0; i < m_nColumns && GetDetailsOf(NULL, i, &sd) == S_OK; ++i) {
								BSTR bs;
								if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs)) {
									nIndex = teBSearchW(methodArgs, nCount, piArgs, bs);
									if (nIndex >= 0) {
										m_pColumns[i].csFlags |= SHCOLSTATE_ONBYDEFAULT;
										m_pColumns[i].nWidth = methodArgs[nIndex].id;
										bDefault = false;
									}
									::SysFreeString(bs);
								}
							}
						}
						if (bDefault) {
							for (int i = uCount; --i >= 0;) {
								m_pSF2->GetDefaultColumnState(i, &m_pColumns[i].csFlags);
								m_pColumns[i].nWidth = -3;
							}
						}
						IShellView *pPreviousView = m_pShellView;
						FOLDERSETTINGS fs;
						pPreviousView->GetCurrentInfo(&fs);
						m_param[SB_ViewMode] = fs.ViewMode;
						Show(FALSE, 0);
						ResetPropEx();
						if (CreateViewWindowEx(pPreviousView) == S_OK) {
							pPreviousView->DestroyViewWindow();
							teDelayRelease(&pPreviousView);
							GetShellFolderView();
						}
						Show(TRUE, 0);
						SetPropEx();
					}
					//Columns Order and Width;
					if (m_hwndLV && ListView_GetView(m_hwndLV) == LV_VIEW_DETAILS) {
						HWND hHeader = ListView_GetHeader(m_hwndLV);
						if (hHeader) {
							UINT nHeader = Header_GetItemCount(hHeader);
							if (nHeader == nCount) {
								SetRedraw(FALSE);
								int *piOrderArray = new int[nHeader];
								try {
									for (int i = nHeader; --i >= 0;) {
										piOrderArray[i] = i;
									}
									WCHAR szText[MAX_COLUMN_NAME_LEN];
									HD_ITEM hdi = { HDI_TEXT | HDI_WIDTH, 0, szText, NULL, MAX_COLUMN_NAME_LEN };
									int nIndex;
									for (int i = nHeader; --i >= 0;) {
										Header_GetItem(hHeader, i, &hdi);
										nIndex = teBSearchW(methodArgs, nCount, piArgs, szText);
										if (nIndex >= 0) {
											nWidth = methodArgs[nIndex].id;
											piOrderArray[nIndex] = i;
											if (nWidth != hdi.cxy) {
												if (nWidth == -2) {
													nWidth = LVSCW_AUTOSIZE;
												} else if (nWidth < 0) {
													int cxChar;
													int i = PSGetColumnIndexXP(szText, &cxChar);
													if (i >= 0) {
														nWidth = cxChar * g_nCharWidth;
													}
												}
												ListView_SetColumnWidth(m_hwndLV, i, nWidth);
											}
										}
									}
									ListView_SetColumnOrderArray(m_hwndLV, nHeader, piOrderArray);
								} catch (...) {
									g_nException = 0;
#ifdef _DEBUG
									g_strException = L"SetColumnsStr";
#endif
								}
								delete [] piOrderArray;
								SetRedraw(TRUE);
							}
						}
					}
				}
				delete [] piArgs;
				InitFolderSize();
			}
#endif
		}
	}
	if (lplpszArgs) {
		LocalFree(lplpszArgs);
	}
#ifdef _2000XP
	if (!g_bUpperVista) {
		for (int i = nCount; --i >= 0;) {
			teSysFreeString(&methodArgs[i].name);
		}
	}
#endif
	delete [] methodArgs0;
	delete [] pPropKey0;
}

VOID AddColumnData(LPWSTR pszColumns, LPWSTR pszName, int nWidth)
{
	swprintf_s(&pszColumns[lstrlen(pszColumns)], MAX_COLUMN_NAME_LEN + 20, L" \"%s\" %d", pszName, nWidth);
}

VOID AddColumnDataEx(LPWSTR pszColumns, BSTR bsName, int nWidth)
{
	AddColumnData(pszColumns, bsName, nWidth);
	teSysFreeString(&bsName);
}

BSTR CteShellBrowser::GetColumnsStr(int nFormat)
{
#ifdef _2000XP
	if (g_bUpperVista && m_pDefultColumns.empty()) {
#else
	if (m_pDefultColumns.empty()) {
#endif
		return NULL;
	}
	WCHAR szColumns[SIZE_BUFF];
	szColumns[0] = 0;
	szColumns[1] = 0;

	IColumnManager *pColumnManager;
	if (m_pShellView && SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager)))) {
		UINT uCount;
		if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_VISIBLE, &uCount)) {
			if (uCount) {
				PROPERTYKEY *pPropKey = new PROPERTYKEY[uCount];
				if SUCCEEDED(pColumnManager->GetColumns(CM_ENUM_VISIBLE, pPropKey, uCount)) {
					CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), (DWORD)(CM_MASK_WIDTH | CM_MASK_NAME) };
					for (UINT i = 0; i < uCount; ++i) {
						if SUCCEEDED(pColumnManager->GetColumnInfo(pPropKey[i], &cmci)) {
							AddColumnDataEx(szColumns, tePSGetNameFromPropertyKeyEx(pPropKey[i], nFormat, this), cmci.uWidth);
						}
					}
				}
				delete [] pPropKey;
			}
		}
		pColumnManager->Release();
#ifdef _2000XP
	} else {
		HWND hHeader = ListView_GetHeader(m_hwndLV);
		if (hHeader) {
			UINT nHeader = Header_GetItemCount(hHeader);
			if (nHeader) {
				int *piOrderArray = new int[nHeader];
				ListView_GetColumnOrderArray(m_hwndLV, nHeader, piOrderArray);
				WCHAR szText[MAX_COLUMN_NAME_LEN];
				HD_ITEM hdi = { HDI_TEXT | HDI_WIDTH, 0, szText, NULL, MAX_COLUMN_NAME_LEN };
				for (UINT i = 0; i < nHeader; ++i) {
					Header_GetItem(hHeader, piOrderArray[i], &hdi);
					AddColumnDataXP(szColumns, szText, hdi.cxy, nFormat);
				}
				delete [] piOrderArray;
			}
		} else {
			SHCOLSTATEF csFlags;
			SHELLDETAILS sd;
			BSTR bs;
			for (UINT i = 0; i < m_nColumns && GetDefaultColumnState(i, &csFlags) == S_OK; ++i) {
				if (csFlags & SHCOLSTATE_ONBYDEFAULT) {
					if (GetDetailsOf(NULL, i, &sd) == S_OK) {
						if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs)) {
							int nWidth = m_pColumns[i].nWidth;
							if (nWidth < 0 || nWidth >= 32768) {
								nWidth = sd.cxChar * g_nCharWidth;
							}
							AddColumnDataXP(szColumns, bs, nWidth, nFormat);
							::SysFreeString(bs);
						}
					}
				}
			}
		}
#endif
	}
	return SysAllocString(&szColumns[1]);
}

VOID CteShellBrowser::GetDefaultColumns()
{
	IColumnManager *pColumnManager;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager))) {
		UINT uColumns = 0;
		if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_VISIBLE, &uColumns)) {
			if (m_pDefultColumns.size() != uColumns) {
				m_pDefultColumns.resize(uColumns);
			}
			pColumnManager->GetColumns(CM_ENUM_VISIBLE, &m_pDefultColumns[0], uColumns);
		}
		pColumnManager->Release();
#ifdef _2000XP
	} else {
		m_pDefultColumns.clear();
#endif
	}
}

STDMETHODIMP CteShellBrowser::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		int nIndex;
		int nFormat = 0;
		TC_ITEM tcItem;
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (!m_pFolderItem && dispIdMember >= DISPID_SELECTIONCHANGED && dispIdMember <= DISPID_INITIALENUMERATIONDONE) {
			return  S_OK;
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD + 0xf000 && dispIdMember <= TE_METHOD_MAX) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		if (m_nUnload == 1 || (m_nUnload & 8)) {
			if (dispIdMember >= TE_PROPERTY + 0xf000 && (dispIdMember & 0xfff) >= 0x100) {
				return S_OK;
			}
		}
		switch (dispIdMember) {
		case TE_PROPERTY + 0xf001://Data
			if (nArg >= 0) {
				VariantClear(&m_vData);
				VariantCopy(&m_vData, &pDispParams->rgvarg[nArg]);
			}
			teVariantCopy(pVarResult, &m_vData);
			return S_OK;

		case TE_PROPERTY + 0xf002://hwnd
			teSetPtr(pVarResult, m_hwnd);
			return S_OK;

		case TE_PROPERTY + 0xf003://Type
			if (nArg >= 0) {
				DWORD dwType = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (dwType == CTRL_SB || dwType == CTRL_EB) {
					if (m_param[SB_Type] != dwType) {
						m_param[SB_Type] = dwType;
						m_bCheckLayout = TRUE;
						GetViewModeAndIconSize(TRUE);
					}
				}
			}
			teSetLong(pVarResult, m_param[SB_Type]);
			return S_OK;


		case TE_METHOD + 0xf004://Navigate
		case TE_PROPERTY + 0xf00B://History
			if (nArg >= 0) {
				UINT wFlags = 0;
				FolderItem *pFolderItem = NULL;
				FolderItems *pFolderItems = NULL;
				GetVariantPath(&pFolderItem, &pFolderItems, &pDispParams->rgvarg[nArg]);
				if (nArg >= 1) {
					wFlags = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
				}
				CteShellBrowser *pSB = NULL;

				DWORD param[SB_Count];
				for (int i = SB_Count; i--;) {
					param[i] = m_param[i];
				}
				param[SB_DoFunc] = 1;
				if (pFolderItem || wFlags & (SBSP_PARENT | SBSP_NAVIGATEBACK | SBSP_NAVIGATEFORWARD | SBSP_RELATIVE)) {
					Navigate3(pFolderItem, wFlags, param, &pSB, pFolderItems);
					SafeRelease(&pFolderItem);
				}

				if (pVarResult) {
					//Navigate
					if (dispIdMember == TE_METHOD + 0xf004) {
						teSetObject(pVarResult, pSB ? pSB : this);
					} else if (pSB) {
						pSB->Release();
					}
				}
			}
			//History
			if (dispIdMember == TE_PROPERTY + 0xf00B) {
				if (pVarResult) {
					CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL);
					VARIANT v;
					for (size_t i = 1; i <= m_ppLog.size(); ++i) {
						try {
							teSetObject(&v, m_ppLog[m_ppLog.size() - i]);
							pFolderItems->ItemEx(-1, NULL, &v);
							VariantClear(&v);
						} catch (...) {
							g_nException = 0;
#ifdef _DEBUG
							g_strException = L"History";
#endif
						}
					}
					pFolderItems->m_nIndex = int(m_ppLog.size()) - 1 - m_uLogIndex;
					teSetObjectRelease(pVarResult, pFolderItems);
				}
			}
			return S_OK;

		case TE_METHOD + 0xfb07://Navigate2Ex
			if (nArg >= 3) {
				FolderItem *pFolderItem = NULL;
				IUnknown *punk;
				if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
					punk->QueryInterface(IID_PPV_ARGS(&pFolderItem));
				}
				if (!pFolderItem && !m_pShellView) {
					GetFolderItemFromIDList(&pFolderItem, g_pidls[CSIDL_DESKTOP]);
				}
				if (pFolderItem) {
					DWORD param[SB_Count];
					for (int i = SB_Count; i--;) {
						param[i] = m_param[i];
					}
					FolderItems *pFolderItems = NULL;
					if (FindUnknown(&pDispParams->rgvarg[nArg - 2], &punk)) {
						punk->QueryInterface(IID_PPV_ARGS(&pFolderItems));
					}
					LPITEMIDLIST pidl = NULL;
					if (FindUnknown(&pDispParams->rgvarg[nArg - 3], &punk)) {
						teGetIDListFromObjectForFV(punk, &pidl);
					}
					if (SUCCEEDED(Navigate2(pFolderItem, GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), param, pFolderItems, m_pFolderItem, this)) && m_bVisible) {
						Show(TRUE, 0);
					}
				}
			}
			return S_OK;

		case TE_METHOD + 0xf007://Navigate2
			if (nArg >= 0) {
				DWORD param[SB_Count];
				for (int i = SB_Count; i--;) {
					m_param[i] = g_paramFV[i];
					param[i] = g_paramFV[i];
				}
#ifdef _2000XP
				if (!g_bUpperVista && m_param[SB_IconSize] >= 96 && m_param[SB_ViewMode] == FVM_ICON) {
					m_param[SB_ViewMode] = FVM_THUMBNAIL;
				}
#endif
				FolderItem *pFolderItem = NULL;
				FolderItems *pFolderItems = NULL;
				GetVariantPath(&pFolderItem, &pFolderItems, &pDispParams->rgvarg[nArg]);
				if (nArg >= 1) {
					wFlags = static_cast<WORD>(GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]));
				}
				for (int i = 0; i <= nArg - 2 && i < SB_DoFunc; ++i) {
					int n = GetIntFromVariant(&pDispParams->rgvarg[nArg - i - 2]);
					if (i == SB_TreeAlign && n == 0) {
						break;
					}
					param[i] = n;
				}
				if (param[SB_Type] != CTRL_SB && param[SB_Type] != CTRL_EB) {
					param[SB_Type] = CTRL_SB;
				}
				param[SB_DoFunc] = 0;
				CteShellBrowser *pSB = NULL;
				Navigate3(pFolderItem, wFlags, param, &pSB, pFolderItems);
				SafeRelease(&pFolderItem);
				if (!pSB) {
					pSB = this;
				}
				if (nArg >= SB_Root + 2) {
					pSB->m_pTV->SetRootV(&pDispParams->rgvarg[nArg - 2 - SB_Root]);
				}
				teSetObject(pVarResult, pSB);
			}
			return S_OK;

		case TE_PROPERTY + 0xf008://Index
			if (nArg >= 0) {
				m_pTC->Move(GetTabIndex(), GetIntFromVariant(&pDispParams->rgvarg[nArg]), m_pTC);
			}
			if (pVarResult) {
				teSetLong(pVarResult, GetTabIndex());
			}
			return S_OK;

		case TE_PROPERTY + 0xf009://FolderFlags
			if (nArg >= 0) {
				m_param[SB_FolderFlags] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				SetFolderFlags(FALSE);
			}
			teSetLong(pVarResult, m_param[SB_FolderFlags]);
			return S_OK;

		case TE_PROPERTY + 0xf010://CurrentViewMode
		case TE_METHOD + 0xf010://SetViewMode
			if (nArg >= 0) {
				m_param[SB_ViewMode] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (nArg >= 1) {
					m_param[SB_IconSize] = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
				}
				SetViewModeAndIconSize(nArg >= 1);
			}
			SetFolderFlags(FALSE);
			teSetLong(pVarResult, m_param[SB_ViewMode]);
			return S_OK;

		case TE_PROPERTY + 0xf011://IconSize
			if (m_pShellView && m_bIconSize) {
				GetViewModeAndIconSize(TRUE);
				if (nArg >= 0) {
					int iImageSize = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					IFolderView2 *pFV2;
					if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
						if (iImageSize != m_param[SB_IconSize]) {
							pFV2->SetViewModeAndIconSize(static_cast<FOLDERVIEWMODE>(m_param[SB_ViewMode]), iImageSize);
						}
						pFV2->Release();
						GetViewModeAndIconSize(TRUE);
#ifdef _2000XP
					} else {
						m_param[SB_IconSize] = iImageSize;
#endif
					}
				}
			}
			teSetLong(pVarResult, m_param[SB_IconSize]);
			return S_OK;

		case TE_PROPERTY + 0xf012://Options
			if (nArg >= 0) {
				m_param[SB_Options] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (m_pExplorerBrowser) {
					m_pExplorerBrowser->SetOptions(static_cast<EXPLORER_BROWSER_OPTIONS>((m_param[SB_Options] & ~(EBO_NAVIGATEONCE | EBO_ALWAYSNAVIGATE)) | EBO_NOTRAVELLOG));
					GetViewModeAndIconSize(FALSE);
				}
			}
			if (pVarResult) {
				if (m_pExplorerBrowser) {
					m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
				}
				teSetLong(pVarResult, m_param[SB_Options]);
			}
			return S_OK;

		case TE_PROPERTY + 0xf013://SizeFormat
			if (nArg >= 0) {
				m_nSizeFormat = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (m_hwndLV) {
					InvalidateRect(m_hwndLV, NULL, FALSE);
				}
			}
			teSetLong(pVarResult, m_nSizeFormat);
			return S_OK;

		case TE_PROPERTY + 0xf014://NameFormat //Deprecated
			return S_OK;

		case TE_PROPERTY + 0xf016://ViewFlags
			if (nArg >= 0) {
				m_param[SB_ViewFlags] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (m_param[SB_ViewFlags] & CDB2GVF_NOSELECTVERB) {
					g_param[TE_ColumnEmphasis] = FALSE;
				}
			}
			teSetLong(pVarResult, m_param[SB_ViewFlags]);
			return S_OK;

		case TE_PROPERTY + 0xf017://Id
			teSetLong(pVarResult, m_nSB);
			return S_OK;

		case TE_PROPERTY + 0xf018://FilterView
		case TE_METHOD + 0xf018://Search
			if (nArg >= 0) {
				VARIANT v;
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				if (wFlags & DISPATCH_METHOD) {
					if (m_pdisp) {
						HRESULT hr = S_FALSE;
						if (g_pOnFunc[TE_OnFilterView]) {
							VARIANT vResult;
							VariantInit(&vResult);
							VARIANTARG *pv = GetNewVARIANT(2);
							teSetObject(&pv[1], this);
							VariantCopy(&pv[0], &v);
							Invoke4(g_pOnFunc[TE_OnFilterView], &vResult, 2, pv);
							hr = GetIntFromVariantClear(&vResult);
						}
						if (hr != S_OK) {
							if (Search(v.bstrVal) != S_OK) {
#ifdef _2000XP
								teSysFreeString(&m_bsFilter);
								if (lstrlen(v.bstrVal)) {
									if (StrChr(v.bstrVal, '*')) {
										m_bsFilter = ::SysAllocString(v.bstrVal);
									} else {
										int nLen = ::SysStringLen(v.bstrVal);
										m_bsFilter = ::SysAllocStringLen(NULL, nLen + 2);
										lstrcpy(&m_bsFilter[1], v.bstrVal);
										m_bsFilter[0] = m_bsFilter[nLen + 1] = '*';
										m_bsFilter[nLen + 2] = NULL;
									}
								}
								Refresh(TRUE);
#endif
							}
						}
					}
				} else if (lstrcmpi(m_bsFilter, v.bstrVal)) {
					teSysFreeString(&m_bsFilter);
					if (lstrlen(v.bstrVal)) {
						m_bsFilter = ::SysAllocString(v.bstrVal);
					}
					SafeRelease(&m_ppDispatch[SB_OnIncludeObject]);
					DoFunc(TE_OnFilterChanged, this, S_OK);
				}
				VariantClear(&v);
			}
			teSetSZ(pVarResult, m_bsFilter);
			return S_OK;

		case  TE_PROPERTY + 0xf020://FolderItem
			teSetObject(pVarResult, m_pFolderItem);
			return S_OK;

		case  TE_PROPERTY + 0xf021://TreeView
			teSetObject(pVarResult, m_pTV);
			return S_OK;

		case  TE_PROPERTY + 0xf024://Parent
			IDispatch *pdisp;
			if SUCCEEDED(get_Parent(&pdisp)) {
				teSetObjectRelease(pVarResult, pdisp);
			}
			return S_OK;

		case TE_METHOD + 0xf026://Focus
			SetActive(TRUE);
			CheckChangeTabTC(m_pTC->m_hwnd);
			return S_OK;

		case TE_METHOD + 0xf031://Close
			teSetBool(pVarResult, Close(FALSE));
			if (!m_pTC->m_nLockUpdate) {
				ArrangeWindowEx();
			}
			return S_OK;

		case TE_PROPERTY + 0xf032://Title
			nIndex = GetTabIndex();
			if (nIndex >= 0) {
				if (nArg >= 0) {
					VARIANT vText;
					teVariantChangeType(&vText, &pDispParams->rgvarg[nArg], VT_BSTR);
					SetTitle(vText.bstrVal, nIndex);
					VariantClear(&vText);
				}
				if (pVarResult) {
					BSTR bsText = ::SysAllocStringLen(NULL, MAX_PATH);
					bsText[0] = NULL;
					tcItem.pszText = bsText;
					tcItem.mask = TCIF_TEXT;
					tcItem.cchTextMax = MAX_PATH;
					TabCtrl_GetItem(m_pTC->m_hwnd, nIndex, &tcItem);
					int nCount = lstrlen(tcItem.pszText);
					int j = 0;
					WCHAR c = NULL;
					for (int i = 0; i < nCount; ++i) {
						bsText[j] = bsText[i];
						if (c != '&' || bsText[i] != '&') {
							++j;
							c = bsText[i];
						} else {
							c = NULL;
						}
					}
					bsText[j] = NULL;
					teSetBSTR(pVarResult, &bsText, -1);
				}
			}
			return S_OK;

		case TE_METHOD + 0xf033://Suspend
			Suspend(nArg >= 0 ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : 0);
			return S_OK;

		case TE_METHOD + 0xf040://Items
			FolderItems *pFolderItems;
			if SUCCEEDED(Items(nArg >= 0 ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : SVGIO_ALLVIEW | SVGIO_FLAG_VIEWORDER, &pFolderItems)) {
				teSetObjectRelease(pVarResult, pFolderItems);
			}
			return S_OK;

		case TE_METHOD + 0xf041://SelectedItems
			if SUCCEEDED(SelectedItems(&pFolderItems)) {
				teSetObjectRelease(pVarResult, pFolderItems);
			}
			return S_OK;

		case TE_PROPERTY + 0xf050://ShellFolderView
			teSetObject(pVarResult, m_pdisp);
			return S_OK;

		case TE_PROPERTY + 0xf058://DropTarget
			IDropTarget *pDT;
			pDT = NULL;
			if (!m_pDropTarget2 || FAILED(m_pDropTarget2->QueryInterface(IID_PPV_ARGS(&pDT)))) {
				LPITEMIDLIST pidl;
				if (teGetIDListFromObject(m_pFolderItem, &pidl)) {
					IShellFolder *pSF;
					LPCITEMIDLIST pidlPart;
					if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
						pSF->GetUIObjectOf(g_hwndMain, 1, &pidlPart, IID_IDropTarget, NULL, (LPVOID*)&pDT);
						pSF->Release();
					}
					teCoTaskMemFree(pidl);
				}
			}
			if (pDT) {
				teSetObjectRelease(pVarResult, new CteDropTarget(pDT, NULL));
				pDT->Release();
			}
			return S_OK;

		case TE_PROPERTY + 0xf059://Columns
		case TE_METHOD + 0xf059://GetColumns
			if (m_hwndDV) {
				if (nArg >= 0) {
					if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
						SetColumnsStr(pDispParams->rgvarg[nArg].bstrVal);
					}
					nFormat = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				}
				if (pVarResult) {
					pVarResult->vt = VT_BSTR;
					pVarResult->bstrVal = GetColumnsStr(nFormat);
				}
			}
			return S_OK;
			/*//Searches //test
						case TE_PROPERTY + 0x005A:
							IDispatch *pArray;
							GetNewArray(&pArray);
							IEnumExtraSearch *pees;
							if SUCCEEDED(m_pSF2->EnumSearches(&pees)) {
								EXTRASEARCH es;
								ULONG ulFetched;
								VARIANT v;
								VariantInit(&v);
								WCHAR pszBuff[SIZE_BUFF];
								while (pees->Next(1, &es, &ulFetched) == S_OK) {
									StringFromGUID2(es.guidSearch, pszBuff, 39);
									swprintf_s(&pszBuff[40], SIZE_BUFF, L"\t\s\t\s", es.wszFriendlyName, es.wszUrl);
									teSetSZ(&v, pszBuff);
									teExecMethod(pArray, L"push", NULL, -1, &v);
									VariantClear(&v);
								}
								pees->Release();
							}
							teSetObjectRelease(pVarResult, pArray);
							return S_OK;*/

		case TE_METHOD + 0xf05B://MapColumnToSCID
			if (nArg >= 1 && pVarResult) {
				SHCOLUMNID scid;
#ifdef _2000XP
				if SUCCEEDED(MapColumnToSCID((UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg]), &scid)) {
#else
				if SUCCEEDED(m_pSF2->MapColumnToSCID((UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg]), &scid)) {
#endif
					pVarResult->bstrVal = tePSGetNameFromPropertyKeyEx(scid, GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), this);
					pVarResult->vt = VT_BSTR;
				}
			}
			return S_OK;

		case TE_PROPERTY + 0xf102://hwndList
			teSetPtr(pVarResult, m_hwndLV);
			return S_OK;

		case TE_PROPERTY + 0xf103://hwndView
			teSetPtr(pVarResult, m_hwndDV);
			return S_OK;

		case TE_PROPERTY + 0xf104://SortColumn
		case TE_METHOD + 0xf104://GetSortColumn
			if (nArg >= 0) {
				if (wFlags & DISPATCH_METHOD) {
					nFormat = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				} else if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
					SetSort(pDispParams->rgvarg[nArg].bstrVal);
				}
			}
			if (pVarResult) {
				BSTR bs;
				GetSort(&bs, nFormat);
				teSetBSTR(pVarResult, &bs, -1);
			}
			return S_OK;

		case TE_PROPERTY + 0xf105://GroupBy
			if (nArg >= 0) {
				if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
					SetGroupBy(pDispParams->rgvarg[nArg].bstrVal);
				}
			}
			if (pVarResult) {
				pVarResult->vt = VT_BSTR;
				GetGroupBy(&pVarResult->bstrVal);
			}
			return S_OK;

		case TE_METHOD + 0xf107://HitTest (Screen coordinates)
			if (nArg >= 0 && pVarResult) {
				LVHITTESTINFO info;
				GetPointFormVariant(&info.pt, &pDispParams->rgvarg[nArg]);
				UINT flags = nArg >= 1 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : LVHT_ONITEM;
				teSetLong(pVarResult, (int)DoHitTest(this, info.pt, flags));
				if (pVarResult->lVal < 0) {
					if (m_hwndLV) {
						ScreenToClient(m_hwndLV, &info.pt);
						info.flags = flags;
						ListView_HitTest(m_hwndLV, &info);
						if (info.flags & flags) {
							pVarResult->lVal = info.iItem;
							if (info.flags == LVHT_ONITEMLABEL && m_param[SB_ViewMode] == FVM_DETAILS) {
								RECT rc;
								ListView_GetItemRect(m_hwndLV, info.iItem, &rc, LVIR_ICON);
								if (info.pt.x < rc.left) {
									pVarResult->lVal = -1;
								}
							}
						}
					} else if (g_pAutomation) {
						IUIAutomationElement *pElement, *pElement2 = NULL;
						if SUCCEEDED(g_pAutomation->ElementFromPoint(info.pt, &pElement)) {
							if SUCCEEDED(pElement->GetCurrentPropertyValue(g_PID_ItemIndex, pVarResult)) {
								--pVarResult->lVal;
							}
							if (pVarResult->lVal < 0) {
								IUIAutomationTreeWalker *pWalker = NULL;
								if SUCCEEDED(g_pAutomation->get_ControlViewWalker(&pWalker)) {
									pWalker->GetParentElement(pElement, &pElement2);
									if (pElement2) {
										if SUCCEEDED(pElement2->GetCurrentPropertyValue(g_PID_ItemIndex, pVarResult)) {
											--pVarResult->lVal;
										}
										pElement2->Release();
									}
									pWalker->Release();
								}
							}
							pElement->Release();
						}
					}
				}
				if (nArg == 0) {
					int i = GetIntFromVariantClear(pVarResult);
					IFolderView *pFV;
					if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
						LPITEMIDLIST pidl;
						if SUCCEEDED(pFV->Item(i, &pidl)) {
							LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidl);
							teSetIDListRelease(pVarResult, &pidlFull);
							teCoTaskMemFree(pidl);
						}
						pFV->Release();
					}
				}
			}
			return S_OK;

		case TE_PROPERTY + 0xf108://hwndAlt
			if (nArg >= 0) {
				m_hwndAlt = (HWND)GetPtrFromVariant(&pDispParams->rgvarg[nArg]);
				ArrangeWindow();
			}
			teSetPtr(pVarResult, m_hwndAlt);
			return S_OK;

		case TE_METHOD + 0xf110://ItemCount
			if (pVarResult && m_pShellView) {
				UINT uItem = (nArg >= 0) ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : SVGIO_ALLVIEW;
				if (m_ppDispatch[SB_AltSelectedItems] && uItem == SVGIO_SELECTION) {
					FolderItems *pid;
					if SUCCEEDED(m_ppDispatch[SB_AltSelectedItems]->QueryInterface(IID_PPV_ARGS(&pid))) {
						if SUCCEEDED(pid->get_Count(&pVarResult->lVal)) {
							pVarResult->vt = VT_I4;
						}
						pid->Release();
					}
					return S_OK;
				}
				teSetLong(pVarResult, GetFolderViewAndItemCount(NULL, uItem));
			}
			return S_OK;

		case TE_METHOD + 0xf111://Item
			if (pVarResult && m_pShellView && nArg >= 0) {
				IFolderView *pFV;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
					LPITEMIDLIST pidl;
					if SUCCEEDED(pFV->Item(GetIntFromVariant(&pDispParams->rgvarg[nArg]), &pidl)) {
						LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidl);
						teSetIDListRelease(pVarResult, &pidlFull);
						teCoTaskMemFree(pidl);
					}
					pFV->Release();
				}
			}
			return S_OK;

		case TE_METHOD + 0xf206://Refresh
			if (!m_bVisible && nArg >= 0 && GetBoolFromVariant(&pDispParams->rgvarg[nArg])) {
				m_bRefreshLator = TRUE;
				return S_OK;
			}
			Refresh(TRUE);
			return S_OK;

		case TE_METHOD + 0xfb06://Refresh2Ex
			if (nArg >= 0) {
				FolderItem *pid = NULL;
				GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
				if (pid && m_pFolderItem) {
					m_pFolderItem->Release();
					m_pFolderItem = pid;
					if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
						teILFreeClear(&m_pidl);
						teGetIDListFromObject(m_pFolderItem, &m_pidl);
					}
				}
				Refresh(FALSE);
			}
			return S_OK;

		case TE_METHOD + 0xf207://ViewMenu
			IContextMenu *pCM;
			CteContextMenu *pdispCM;
			pCM = NULL;
			pdispCM = NULL;
			try {
				if SUCCEEDED(m_pShellView->GetItemObject(nArg >= 0 ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : SVGIO_BACKGROUND, IID_PPV_ARGS(&pCM))) {
					CteContextMenu *pdispCM = new CteContextMenu(pCM, NULL, static_cast<IShellBrowser *>(this));
					teSetObject(pVarResult, pdispCM);
				}
			} catch (...) {
				g_nException = 0;
#ifdef _DEBUG
				g_strException = L"ViewMenu";
#endif
			}
			SafeRelease(&pdispCM);
			SafeRelease(&pCM);
			return S_OK;

		case TE_METHOD + 0xf208://TranslateAccelerator
			HRESULT hr;
			hr = E_NOTIMPL;
			if (m_pShellView) {
				if (nArg >= 3) {
					MSG msg;
					msg.hwnd = (HWND)GetPtrFromVariant(&pDispParams->rgvarg[nArg]);
					msg.message = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
					msg.wParam = (WPARAM)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 2]);
					msg.lParam = (LPARAM)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 3]);
					hr = m_pShellView->TranslateAcceleratorW(&msg);
				}
			}
			teSetLong(pVarResult, hr);
			return S_OK;

		case TE_METHOD + 0xf209:// GetItemPosition
			hr = E_NOTIMPL;
			if (m_pShellView) {
				if (nArg >= 1) {
					IFolderView *pFV;
					if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
						LPITEMIDLIST pidl;
						if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
							POINT pt;
							hr = pFV->GetItemPosition(ILFindLastID(pidl), &pt);
							PutPointToVariant(&pt, &pDispParams->rgvarg[nArg - 1]);
							teCoTaskMemFree(pidl);
						}
						pFV->Release();
					}
				}
			}
			teSetLong(pVarResult, hr);
			return S_OK;

		case TE_METHOD + 0xf20A://SelectAndPositionItem
			hr = E_NOTIMPL;
			if (m_pShellView) {
				if (nArg >= 2) {
					VARIANT vMem;
					VariantInit(&vMem);
					LPPOINT ppt = (LPPOINT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 2], &vMem);
					if (ppt) {
						IFolderView *pFV;
						if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
							LPITEMIDLIST pidl = NULL;
							if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
								LPCITEMIDLIST pidlLast = ILFindLastID(pidl);
								hr = pFV->SelectAndPositionItems(1, &pidlLast, ppt, GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]));
								teCoTaskMemFree(pidl);
							}
							pFV->Release();
						}
					}
					VariantClear(&vMem);
				}
			}
			teSetLong(pVarResult, hr);
			return S_OK;

		case TE_METHOD + 0xf280://SelectItem
			if (nArg >= 1) {
				SelectItem(&pDispParams->rgvarg[nArg], GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]));
			}
			return S_OK;

		case TE_PROPERTY + 0xf281://FocusedItem
			FolderItem *pFolderItem;
			if (get_FocusedItem(&pFolderItem) == S_OK) {
				teSetObjectRelease(pVarResult, pFolderItem);
			}
			return S_OK;

		case TE_PROPERTY + 0xf282://GetFocusedItem
			if (pVarResult) {
				GetFocusedIndex(&pVarResult->intVal);
				pVarResult->vt = VT_I4;
			}
			return S_OK;

		case TE_METHOD + 0xf283://GetItemRect
			hr = E_NOTIMPL;
			if (m_pShellView && m_hwndLV) {
				if (nArg >= 1) {
					VARIANT vMem;
					VariantInit(&vMem);
					LPRECT prc = (LPRECT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1], &vMem);
					if (prc) {
						LPITEMIDLIST pidl;
						if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
							LPITEMIDLIST pidlChild = ILFindLastID(pidl);
							IFolderView *pFV = NULL;
							int i = GetFolderViewAndItemCount(&pFV, SVGIO_ALLVIEW);
							while (--i >= 0) {
								LPITEMIDLIST pidlPart;
								if SUCCEEDED(pFV->Item(i, &pidlPart)) {
									if (ILIsEqual(pidlChild, pidlPart)) {
										hr = ListView_GetItemRect(m_hwndLV, i, prc, nArg >= 2 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]) : LVIR_BOUNDS) ? S_OK : E_FAIL;
										i = 0;
									}
									teCoTaskMemFree(pidlPart);
								}
							}
							SafeRelease(&pFV);
							teCoTaskMemFree(pidl);
						}
					}
					teWriteBack(&pDispParams->rgvarg[nArg - 1], &vMem);
				}
			}
			teSetLong(pVarResult, hr);
			return S_OK;

		case TE_METHOD + 0xf300://Notify
			if (nArg >= 2) {
				long lEvent = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				BOOL bTFS = (nArg >= 3 && GetBoolFromVariant(&pDispParams->rgvarg[nArg - 3]));
				for (int i = 1; i < (lEvent & (SHCNE_RENAMEITEM | SHCNE_RENAMEFOLDER) ? 3 : 2); ++i) {
					LPITEMIDLIST pidl;
					if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg - i])) {
						if (!ILIsEmpty(pidl)) {
							IShellFolder2 *pSF2;
							LPCITEMIDLIST pidlLast;
							if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF2), &pidlLast)) {
								if (bTFS) {
									WCHAR szBuf[MAX_COLUMN_NAME_LEN];
									SetFolderSize(pSF2, pidlLast, szBuf, _countof(szBuf));
								} else {
									BSTR bsPath;
									if SUCCEEDED(teGetDisplayNameBSTR(pSF2, pidl, SHGDN_FORPARSING, &bsPath)) {
										if SUCCEEDED(teDelProperty(m_ppDispatch[SB_TotalFileSize], bsPath)) {
											if (m_hwndLV) {
												InvalidateRect(m_hwndLV, NULL, FALSE);
											}
										}
										teSysFreeString(&bsPath);
									}
								}
								pSF2->Release();
							}
						}
						teCoTaskMemFree(pidl);
					}
				}
			}
			return S_OK;

		case TE_METHOD + 0xf400://NavigateComplete
			m_bBeforeNavigate = FALSE;
			return S_OK;

		case TE_METHOD + 0xf500://NavigationComplete
			NavigateComplete(FALSE);
			return S_OK;

		case TE_METHOD + 0xf501://AddItem
			if (nArg >= 0) {
				if (nArg >= 1) {
					DWORD dwSessionId = GetIntFromVariant(&pDispParams->rgvarg[0]);
					if (dwSessionId != m_dwSessionId) {
						return E_ACCESSDENIED;
					}
				}
				LPITEMIDLIST pidl;
				if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
					AddItem(pidl);
				}
			}
			return S_OK;

		case TE_METHOD + 0xf502://RemoveItem
			hr = E_NOTIMPL;
			if (nArg >= 0 && m_pShellView) {
				LPITEMIDLIST pidl;
				if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
					IFolderView *pFV = NULL;
					IResultsFolder *pRF;
					if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) && SUCCEEDED(pFV->GetFolder(IID_PPV_ARGS(&pRF)))) {
						hr = pRF->RemoveIDList(pidl);
						pRF->Release();
#ifdef _2000XP
					} else if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
						IShellFolderView *pSFV;
						if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
							BSTR bs;
							if SUCCEEDED(teGetDisplayNameFromIDList(&bs, pidl, SHGDN_FORPARSING)) {
								teILCloneReplace(&pidl, teILCreateFromPath(bs));
								LPITEMIDLIST pidlChild = teILCreateResultsXP(pidl);
								UINT ui;
								hr = pSFV->RemoveObject(pidlChild, &ui);
								teCoTaskMemFree(pidlChild);
								::SysFreeString(bs);
							}
							pSFV->Release();
						}
#endif
					}
					SafeRelease(&pFV);
					teCoTaskMemFree(pidl);
				}
			}
			teSetLong(pVarResult, hr);
			if (hr == S_OK) {
				SetTimer(g_hwndMain, TET_Status, 500, teTimerProc);
			}
			return S_OK;

		case TE_METHOD + 0xf503://AddItems
			if (nArg >= 0) {
				if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp) || teGetDispatchFromString(&pDispParams->rgvarg[nArg], &pdisp)) {
					if (!LoadString(g_hShell32, 13585, g_szStatus, _countof(g_szStatus))) {
						LoadString(g_hShell32, 6478, g_szStatus, _countof(g_szStatus));
					}
					DoStatusText(this, g_szStatus, 0);
					TEAddItems *pAddItems = new TEAddItems[1];
					CoMarshalInterThreadInterfaceInStream(IID_IDispatch, static_cast<IDispatch *>(this), &pAddItems->pStrmSB);
					CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pdisp, &pAddItems->pStrmArray);
					pdisp->Release();
					pAddItems->pv = GetNewVARIANT(2);
					teSetLong(&pAddItems->pv[0], m_dwSessionId);
					pAddItems->bDeleted = nArg >= 1 && GetBoolFromVariant(&pDispParams->rgvarg[nArg - 1]);
					pAddItems->bNavigateComplete = FALSE;
					pAddItems->pStrmOnCompleted = NULL;
					if (nArg >= 2) {
						IUnknown *punk;
						if (FindUnknown(&pDispParams->rgvarg[nArg - 2], &punk)) {
							CoMarshalInterThreadInterfaceInStream(IID_IDispatch, punk, &pAddItems->pStrmOnCompleted);
						} else {
							pAddItems->bNavigateComplete = GetBoolFromVariant(&pDispParams->rgvarg[nArg - 2]);
						}
					}
					_beginthread(&threadAddItems, 0, pAddItems);
				}
			}
			return S_OK;

		case TE_METHOD + 0xf504://RemoveAll
			teSetLong(pVarResult, RemoveAll());
			return S_OK;

		case TE_PROPERTY + 0xf505://SessionId
			teSetLong(pVarResult, m_dwSessionId);
			return S_OK;
			//
		case DISPID_VALUE:
			teSetObject(pVarResult, this);
		case DISPID_TE_UNDEFIND:
			return S_OK;
			//DIID_DShellFolderViewEvents
/*///
			case DISPID_EXPLORERWINDOWREADY:
				return E_NOTIMPL;
//*///
		case DISPID_SELECTIONCHANGED://XP+
			SetStatusTextSB(NULL);
			return DoFunc(TE_OnSelectionChanged, this, S_OK);

		case DISPID_CONTENTSCHANGED://XP-
			if (m_pShellView) {
				IFolderView2 *pFV2;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
					WCHAR pszFormat[MAX_STATUS];
					pszFormat[0] = NULL;
					LPWSTR lpBuf = pszFormat;
					if (m_dwUnavailable) {
						LoadString(g_hShell32, 4157, pszFormat, MAX_STATUS);
					} else if (!m_bFiltered) {
						lpBuf = NULL;
					}
					pFV2->SetText(FVST_EMPTYTEXT, lpBuf);
					pFV2->Release();
				}
			}
			return DoFunc(TE_OnContentsChanged, this, S_OK);

		case DISPID_FILELISTENUMDONE://XP+
			SetObjectRect();
			if (!m_bNavigateComplete && ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
				return S_OK;
			}
			if (m_bRefreshing) {
				m_bRefreshing = FALSE;
				if (!m_bNavigateComplete && m_bsAltSortColumn) {
					BSTR bs = ::SysAllocString(m_bsAltSortColumn);
					SetSort(bs);
					::SysFreeString(bs);
				}
			}
			OnNavigationComplete2();
			SetListColumnWidth();
			if (!m_hwndLV && m_pExplorerBrowser) {
				SetRedraw(TRUE);
				if (IUnknown_GetWindow(m_pExplorerBrowser, &m_hwnd) == S_OK && m_hwnd) {
					RedrawWindow(m_hwnd, NULL, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
				}
			}
			SetTimer(g_hwndMain, TET_Status, 500, teTimerProc);
			return S_OK;
		case DISPID_VIEWMODECHANGED://XP+
			m_bCheckLayout = TRUE;
			SetFolderFlags(TRUE);
			SetListColumnWidth();
			return DoFunc(TE_OnViewModeChanged, this, S_OK);
		case DISPID_BEGINDRAG://XP+
			DoFunc1(TE_OnBeginDrag, this, pVarResult);
			if (pVarResult->vt != VT_BOOL || pVarResult->boolVal) {
				BOOL bHandled = m_bRegenerateItems || ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER]);
				if (bHandled || g_param[TE_ViewOrder]) {
					FolderItems *pid;
					if SUCCEEDED(SelectedItems(&pid)) {
						if (bHandled) {
							IDataObject *pDataObj = NULL;
							if SUCCEEDED(pid->QueryInterface(IID_PPV_ARGS(&pDataObj))) {
								DWORD dwEffect = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
								teDoDragDrop(g_hwndMain, pDataObj, &dwEffect, FALSE);
								teSetBool(pVarResult, FALSE);
								pDataObj->Release();
							}
						} else if (g_param[TE_ViewOrder]) {
							VARIANT v;
							teSetLong(&v, 0);
							FolderItem *pFolderItem;
							if SUCCEEDED(pid->Item(v, &pFolderItem)) {
								teSetObjectRelease(&v, pFolderItem);
								SelectItem(&v, SVSI_SELECT | SVSI_FOCUSED | SVSI_NOTAKEFOCUS);
								VariantClear(&v);
							}
						}
						pid->Release();
					}
				}
			}
			return S_OK;
		case DISPID_COLUMNSCHANGED://XP-
			InitFolderSize();
			return DoFunc(TE_OnColumnsChanged, this, S_OK);
		case DISPID_ICONSIZECHANGED://XP-
			return DoFunc(TE_OnIconSizeChanged, this, S_OK);
		case DISPID_SORTDONE://XP-
			if (m_nGroupByDelay) {
				m_nGroupByDelay -= 10;
				if (m_bsNextGroup) {
					SetTimer(m_hwnd, (UINT_PTR)&m_bsNextGroup, 250, teTimerProcGroupBy);
				}
			}
			if (m_hwndLV) {
				FixColumnEmphasis();
			}
			m_pTC->RedrawUpdate();
			if (m_bRefreshing) {
				m_bRefreshing = FALSE;
			} else if (m_pShellView && m_bsAltSortColumn) {
				IFolderView2 *pFV2;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
					BSTR bs = NULL;
					GetSort2(&bs, 0);
					if (bs) {
						teSysFreeString(&m_bsAltSortColumn);
						teSysFreeString(&bs);
					}
				}
			}
			if (m_nFocusItem < 0) {
				FocusItem();
			}
			IFolderView *pFV;
			if (GetFolderViewAndItemCount(&pFV,SVGIO_SELECTION) == 0) {
				pFV->SelectItem(0, SVSI_FOCUSED | SVSI_NOTAKEFOCUS);
			}
			SafeRelease(&pFV);
			return DoFunc(TE_OnSort, this, S_OK);
		case DISPID_INITIALENUMERATIONDONE://XP-
			if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
				return S_OK;
			}
			if (m_bViewCreated) {
				return OnViewCreated(NULL);
			}
			SetFolderFlags(FALSE);
			return S_OK;
		default:
			if (dispIdMember >= START_OnFunc && dispIdMember < START_OnFunc + Count_SBFunc) {
				teInvokeObject(&m_ppDispatch[dispIdMember - START_OnFunc], wFlags, pVarResult, nArg, pDispParams->rgvarg);
				return S_OK;
			}
		}
		if (m_pdisp) {
			if (dispIdMember == m_dispidSortColumns && (wFlags & DISPATCH_PROPERTYPUT) && nArg >= 0 && pDispParams->rgvarg[nArg].vt == VT_BSTR) {
				VARIANT v;
				VariantInit(&v);
				Invoke5(m_pdisp, dispIdMember, DISPATCH_PROPERTYGET, &v, 0, NULL);
				int iChanged = v.vt == VT_BSTR ? lstrcmpi(pDispParams->rgvarg[nArg].bstrVal, v.bstrVal) : 1;
				VariantClear(&v);
				if (iChanged) {
					m_nGroupByDelay += 10;
				} else {
					return S_OK;
				}
			}
			return m_pdisp->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
	} catch (...) {
		return teException(pExcepInfo, "FolderView", methodSB, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

STDMETHODIMP CteShellBrowser::get_Application(IDispatch **ppid)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pdisp->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->get_Application(ppid);
		pSFVD->Release();
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::get_Parent(IDispatch **ppid)
{
	return m_pTC->QueryInterface(IID_PPV_ARGS(ppid));
}

STDMETHODIMP CteShellBrowser::get_Folder(Folder **ppid)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pdisp->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->get_Folder(ppid);
		pSFVD->Release();
	}
	return SUCCEEDED(hr) ? hr : GetFolderObjFromIDList(m_pidl, ppid);
}

STDMETHODIMP CteShellBrowser::SelectedItems(FolderItems **ppid)
{
	return Items(g_param[TE_ViewOrder] ? SVGIO_SELECTION | SVGIO_FLAG_VIEWORDER : SVGIO_SELECTION, ppid);
}

HRESULT CteShellBrowser::Items(UINT uItem, FolderItems **ppid)
{
	if (m_ppDispatch[SB_AltSelectedItems] && (uItem & SVGIO_SELECTION)) {
		return m_ppDispatch[SB_AltSelectedItems]->QueryInterface(IID_PPV_ARGS(ppid));
	}
	FolderItems *pItems = NULL;
	IDataObject *pDataObj = NULL;
	if (m_pShellView) {
		int nCount = ListView_GetSelectedCount(m_hwndLV);
		if (nCount || !m_hwndLV || (uItem & SVGIO_TYPE_MASK) != SVGIO_SELECTION) {
#ifdef _2000XP
			if (!g_bUpperVista && nCount > 1) {
				BOOL bResultsFolder = ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER]);
				if (bResultsFolder || (uItem & (SVGIO_TYPE_MASK | SVGIO_FLAG_VIEWORDER)) == (SVGIO_SELECTION | SVGIO_FLAG_VIEWORDER)) {
					CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL);
					IShellFolderView *pSFV;
					if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
						if (uItem & SVGIO_SELECTION) {
							int nIndex = -1, nFocus = -1;
							if ((uItem & SVGIO_FLAG_VIEWORDER) == 0) {
								nFocus = ListView_GetNextItem(m_hwndLV, -1, LVNI_FOCUSED | LVNI_ALL);
								AddPathXP(pFolderItems, pSFV, nFocus, bResultsFolder);
							}
							while ((nIndex = ListView_GetNextItem(m_hwndLV, nIndex, LVNI_SELECTED | LVNI_ALL)) >= 0) {
								if (nIndex != nFocus) {
									AddPathXP(pFolderItems, pSFV, nIndex, bResultsFolder);
								}
							}
						} else {
							UINT uCount = 0;
							pSFV->GetObjectCount(&uCount);
							for (UINT u = 0; u < uCount; ++u) {
								AddPathXP(pFolderItems, pSFV, u, bResultsFolder);
							}
						}
						pSFV->Release();
					}
					*ppid = pFolderItems;
					return S_OK;
				}
			}
#endif
			m_pShellView->GetItemObject(uItem, IID_PPV_ARGS(&pDataObj));
		}
	}
	CteFolderItems *pid = new CteFolderItems(pDataObj, pItems);
	if (m_bRegenerateItems) {
		pid->Regenerate(TRUE);
	}
	*ppid = pid;
	SafeRelease(&pDataObj);
	return S_OK;
}

#ifdef _2000XP
VOID CteShellBrowser::AddPathXP(CteFolderItems *pFolderItems, IShellFolderView *pSFV, int nIndex, BOOL bResultsFolder)
{
	try {
		LPITEMIDLIST pidl;
		if SUCCEEDED(pSFV->GetObjectW(&pidl, nIndex)) {
			VARIANT v;
			VariantInit(&v);
			if (bResultsFolder) {
				if SUCCEEDED(teGetDisplayNameBSTR(m_pSF2, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &v.bstrVal)) {
					v.vt = VT_BSTR;
				}
			} else {
				FolderItem *pid;
				LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidl);
				if (GetFolderItemFromIDList(&pid, pidlFull)) {
					teSetObject(&v, pid);
					pid->Release();
				}
				teCoTaskMemFree(pidlFull);
			}
			if (v.vt != VT_EMPTY) {
				pFolderItems->ItemEx(-1, NULL, &v);
				VariantClear(&v);
			}
			teILFreeClear(&pidl);
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"AddPathXP";
#endif
	}
}

int CteShellBrowser::PSGetColumnIndexXP(LPWSTR pszName, int *pcxChar)
{
	SHELLDETAILS sd;
	BSTR bs;
	for (UINT i = 0; i < m_nColumns && GetDetailsOf(NULL, i, &sd) == S_OK; ++i) {
		if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs)) {
			if (teStrSameIFree(bs, pszName)) {
				if (pcxChar) {
					*pcxChar = sd.cxChar;
				}
				return i;
			}
		}
	}
	return -1;
}

BSTR CteShellBrowser::PSGetNameXP(LPWSTR pszName, int nFormat)
{
	if (nFormat == 0) {
		return ::SysAllocString(pszName);
	}
	int i = PSGetColumnIndexXP(pszName, NULL);
	if (i >= 0) {
		SHCOLUMNID scid;
		if SUCCEEDED(MapColumnToSCID(i, &scid)) {
			return tePSGetNameFromPropertyKeyEx(scid, nFormat, this);
		}
	}
	return NULL;
}

VOID CteShellBrowser::AddColumnDataXP(LPWSTR pszColumns, LPWSTR pszName, int nWidth, int nFormat)
{
	if (nFormat) {
		AddColumnDataEx(pszColumns, PSGetNameXP(pszName, nFormat), nWidth);
		return;
	}
	AddColumnData(pszColumns, pszName, nWidth);
}
#endif

#if defined(_USE_SHELLBROWSER) || defined(_2000XP)
HRESULT CteShellBrowser::NavigateSB(IShellView *pPreviousView, FolderItem *pPrevious)
{
	HRESULT hr = E_FAIL;
	int nCreate = 4;
	while ((nCreate -= 2) > 0) {
		if (GetShellFolder2(&m_pidl) == S_OK) {
#ifdef _2000XP
			if (!g_bUpperVista) {
				m_nFolderName = MAXINT;
				if (m_nColumns) {
					delete [] m_pColumns;
					m_nColumns = 0;
				}
				SHCOLUMNID scid;
				while (m_pSF2->MapColumnToSCID(m_nColumns, &scid) == S_OK && m_nColumns < MAX_COLUMNS) {
					++m_nColumns;
				}
				m_nColumns += 2;
				m_pColumns = new TEColumn[m_nColumns];
				BOOL bIsResultsFolder = ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER]);
				for (int i = m_nColumns; i--;) {
					if FAILED(m_pSF2->GetDefaultColumnState(i, &m_pColumns[i].csFlags)) {
						m_pColumns[i].csFlags = SHCOLSTATE_TYPE_STR | SHCOLSTATE_ONBYDEFAULT;
					}
					m_pColumns[i].nWidth = -1;
					if (bIsResultsFolder && m_pSF2->MapColumnToSCID(i, &scid) == S_OK && IsEqualPropertyKey(scid, PKEY_ItemFolderNameDisplay)) {
						m_nFolderName = i;
					}
				}
				m_pColumns[m_nColumns - 2].csFlags = SHCOLSTATE_TYPE_STR | SHCOLSTATE_SECONDARYUI | SHCOLSTATE_PREFER_VARCMP;
				m_pColumns[m_nColumns - 1].csFlags = SHCOLSTATE_TYPE_STR | SHCOLSTATE_SECONDARYUI | SHCOLSTATE_PREFER_VARCMP;
			}
#endif
			try {
				hr = CreateViewWindowEx(pPreviousView);
			} catch (...) {
				hr = E_FAIL;
				g_nException = 0;
#ifdef _DEBUG
				g_strException = L"NavigateSB";
#endif
			}
		}
		if (hr != S_OK) {
			teILCloneReplace(&m_pidl, g_pidls[CSIDL_RESULTSFOLDER]);
			m_dwUnavailable = GetTickCount();
			++nCreate;
			hr = S_FALSE;
		}
	}
	return hr;
}
#endif

#if defined(_USE_SHELLBROWSER) || defined(_2000XP)
HRESULT CteShellBrowser::CreateViewWindowEx(IShellView *pPreviousView)
{
	HRESULT hr = E_FAIL;
	m_pShellView = NULL;
	RECT rc;
	if (m_pSF2) {
		if SUCCEEDED(CreateViewObject(m_pTC->m_hwndStatic, IID_PPV_ARGS(&m_pShellView)) && m_pShellView) {
#ifdef _2000XP
			if (!g_bUpperVista && m_param[SB_ViewMode] == FVM_ICON && m_param[SB_IconSize] >= 96) {
				m_param[SB_ViewMode] = FVM_THUMBNAIL;
			}
#endif
			try {
				if (::InterlockedIncrement(&m_nCreate) <= 1) {
					m_hwnd = NULL;
					IEnumIDList *peidl = NULL;
					BOOL bResultsFolder = ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER]);
					hr =  bResultsFolder ? S_OK : m_pSF2->EnumObjects(g_hwndMain, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN | SHCONTF_INCLUDESUPERHIDDEN | SHCONTF_NAVIGATION_ENUM, &peidl);
					if (hr == S_OK) {
						if (m_nForceViewMode != FVM_AUTO) {
							m_param[SB_ViewMode] = m_nForceViewMode;
							m_bAutoVM = FALSE;
							m_nForceViewMode = FVM_AUTO;
						}
						FOLDERSETTINGS fs = { m_bAutoVM && !bResultsFolder ? FVM_AUTO : m_param[SB_ViewMode], (m_param[SB_FolderFlags] | FWF_USESEARCHFOLDER) & ~FWF_NOENUMREFRESH };
						hr = m_pShellView->CreateViewWindow(pPreviousView, &fs, static_cast<IShellBrowser *>(this), &rc, &m_hwnd);
					} else {
						hr = E_FAIL;
					}
					SafeRelease(&peidl);
				} else {
					hr = OLE_E_BLANK;
				}
			} catch (...) {
				hr = E_FAIL;
				g_nException = 0;
#ifdef _DEBUG
				g_strException = L"CreateViewWindowEx";
#endif
			}
			InterlockedDecrement(&m_nCreate);
			if (hr == S_FALSE) {
				IFolderView *pFV;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
					IResultsFolder *pRF;
					if SUCCEEDED(pFV->GetFolder(IID_PPV_ARGS(&pRF))) {
						hr = S_OK;
						pRF->Release();
					}
					pFV->Release();
				}
			}
			if (hr == S_OK) {
				GetDefaultColumns();
			}
		}
	}
	return hr;
}
#endif

STDMETHODIMP CteShellBrowser::get_FocusedItem(FolderItem **ppid)
{
	HRESULT hr = E_NOTIMPL;
	*ppid = NULL;
	if (m_ppDispatch[SB_AltSelectedItems]) {
		FolderItems *pItems;
		hr = m_ppDispatch[SB_AltSelectedItems]->QueryInterface(IID_PPV_ARGS(&pItems));
		if SUCCEEDED(hr) {
			VARIANT v;
			VariantInit(&v);
			teExecMethod(m_ppDispatch[SB_AltSelectedItems], L"Index", &v, 0, NULL);
			hr = pItems->Item(v, ppid);
			pItems->Release();
		}
	} else if (m_pdisp) {
		IShellFolderViewDual *pSFVD;
		if SUCCEEDED(m_pdisp->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
			hr = pSFVD->get_FocusedItem(ppid);
			if (hr == S_OK) {
				CteFolderItem *pid1;
				teQueryFolderItem(ppid, &pid1);
				pid1->Release();
			}
			pSFVD->Release();
		}
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::SelectItem(VARIANT *pvfi, int dwFlags)
{
	HRESULT hr = E_FAIL;
	if (m_pShellView) {
		teFixListState(m_hwndLV, dwFlags);
		if (teVarIsNumber(pvfi) && pvfi->vt != VT_BSTR) {
			IFolderView *pFV = NULL;
			int nIndex = GetIntFromVariant(pvfi);
			int nCount = GetFolderViewAndItemCount(&pFV, SVGIO_ALLVIEW);
			if (nIndex < nCount) {
				if (dwFlags & (SVSI_SELECTIONMARK | (SVSI_KEYBOARDSELECT & ~SVSI_SELECT))) {//21.4.7
					pFV->SelectItem(nIndex, dwFlags & (SVSI_SELECTIONMARK | SVSI_KEYBOARDSELECT | SVSI_NOTAKEFOCUS));
				}
				hr = pFV->SelectItem(nIndex, dwFlags & ~(SVSI_SELECTIONMARK | (SVSI_KEYBOARDSELECT & ~SVSI_SELECT)));
			}
			SafeRelease(&pFV);
		} else {
			IDataObject *pDataObj;
			if (GetDataObjFromVariant(&pDataObj, pvfi)) {
				long nCount;
				LPITEMIDLIST *ppidllist = IDListFormDataObj(pDataObj, &nCount);
				if (ppidllist) {
					teCoTaskMemFree(ppidllist[0]);
					for (int i = 1; i <= nCount; ++i) {
						hr = SelectItemEx(ppidllist[i], dwFlags);
						teCoTaskMemFree(ppidllist[i]);
						dwFlags &= ~(SVSI_FOCUSED | SVSI_DESELECTOTHERS | SVSI_SELECTIONMARK);
					}
					delete[] ppidllist;
				}
				SafeRelease(&pDataObj);
			}
			if FAILED(hr) {
				LPITEMIDLIST pidl;
				teGetIDListFromVariant(&pidl, pvfi);
				hr = SelectItemEx(pidl, dwFlags);
				teCoTaskMemFree(pidl);
			}
		}
		CteFolderItem *pid1 = NULL;
		if (m_pFolderItem && SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1))) {
			teILFreeClear(&pid1->m_pidlFocused);
			pid1->Release();
		}
	}
	return hr;
}

HRESULT CteShellBrowser::SelectItemEx(LPITEMIDLIST pidl, int dwFlags)
{
	HRESULT hr = E_FAIL;
	if (m_pShellView) {
		LPITEMIDLIST pidlLast = ILFindLastID(pidl);
		if (dwFlags & (SVSI_SELECTIONMARK | (SVSI_KEYBOARDSELECT & ~SVSI_SELECT))) {//21.4.7
			m_pShellView->SelectItem(pidlLast, dwFlags & (SVSI_SELECTIONMARK | SVSI_KEYBOARDSELECT | SVSI_NOTAKEFOCUS));
		}
		hr = m_pShellView->SelectItem(pidlLast, dwFlags & ~(SVSI_SELECTIONMARK | (SVSI_KEYBOARDSELECT & ~SVSI_SELECT)));
		if (FAILED(hr) && (dwFlags & SVSI_DESELECTOTHERS)) {
			hr = m_pShellView->SelectItem(pidlLast, SVSI_DESELECTOTHERS);
		}
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::PopupItemMenu(FolderItem *pfi, VARIANT vx, VARIANT vy, BSTR *pbs)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pdisp->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->PopupItemMenu(pfi, vx, vy, pbs);
		pSFVD->Release();
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::get_Script(IDispatch **ppDisp)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pdisp->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->get_Script(ppDisp);
		pSFVD->Release();
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::get_ViewOptions(long *plViewOptions)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pdisp->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->get_ViewOptions(plViewOptions);
		pSFVD->Release();
	}
	return hr;
}

int CteShellBrowser::GetTabIndex()
{
	if (m_pTC) {
		int i;
		TC_ITEM tcItem;

		for (i = TabCtrl_GetItemCount(m_pTC->m_hwnd); --i >= 0;) {
			tcItem.mask = TCIF_PARAM;
			TabCtrl_GetItem(m_pTC->m_hwnd, i, &tcItem);
			if (tcItem.lParam == m_nSB) {
				return i;
			}
		}
	}
	return -1;
}

STDMETHODIMP CteShellBrowser::OnNavigationPending(PCIDLIST_ABSOLUTE pidlFolder)
{
	if (ILIsEqual(m_pidl, pidlFolder)) {
		return S_OK;
	}
	if (!m_pFolderItem1 && !m_dwUnavailable && ILIsEqual(pidlFolder, g_pidls[CSIDL_RESULTSFOLDER])) {
		m_nSuspendMode = 4;
		return E_FAIL;
	}
	return OnNavigationPending2((LPITEMIDLIST)pidlFolder);
}

HRESULT CteShellBrowser::OnNavigationPending2(LPITEMIDLIST pidlFolder)
{
	LPITEMIDLIST pidlPrevius = m_pidl;
	m_pidl = ::ILClone(pidlFolder);
	FolderItem *pPrevious = m_pFolderItem;
	if (m_pFolderItem1) {
		m_pFolderItem = m_pFolderItem1;
		m_pFolderItem1 = NULL;
	} else if (!m_dwUnavailable) {
		GetFolderItemFromIDList(&m_pFolderItem, m_pidl);
	}

	HRESULT hr = OnBeforeNavigate(pPrevious, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
	if FAILED(hr) {
		if ((hr == E_ABORT || teILIsBlank(pPrevious)) && Close(FALSE)) {
			return hr;
		}
		m_uLogIndex = m_uPrevLogIndex;
		teCoTaskMemFree(m_pidl);
		m_pidl = pidlPrevius;
		GetShellFolder2(&m_pidl);
//		pidlPrevius = NULL;
		FolderItem *pid;
		m_pFolderItem->QueryInterface(IID_PPV_ARGS(&pid));
		m_pFolderItem = pPrevious;
//		pPrevious = NULL;
		if (hr == E_ACCESSDENIED) {
			HRESULT hr2 = BrowseObject2(pid, SBSP_NEWBROWSER | SBSP_ABSOLUTE);
			if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
				hr = hr2;
			}
		}
		InitFilter();
		InitFolderSize();
		pid->Release();
		m_nSuspendMode = 0;
		return hr;
	}
	if (!ILIsEqual(m_pidl, pidlPrevius)) {
		InitFilter();
	}
	teILFreeClear(&pidlPrevius);
	ResetPropEx();
	//History / Management
	SetHistory(NULL, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
	SafeRelease(&pPrevious);
	if (!m_bAutoVM) {
		SetViewModeAndIconSize(TRUE);
	}
	m_bSetRedraw = FALSE;
	SetRedraw(FALSE);
	return hr;
}

HRESULT CteShellBrowser::IncludeObject2(IShellFolder *pSF, LPCITEMIDLIST pidl, LPITEMIDLIST pidlFull)
{
	HRESULT hr = S_OK;
	if (pSF) {
		if (!(m_param[SB_ViewFlags] & CDB2GVF_SHOWALLFILES)) {
			WIN32_FIND_DATA wfd;
			if SUCCEEDED(SHGetDataFromIDList(pSF, pidl, SHGDFIL_FINDDATA, &wfd, sizeof(WIN32_FIND_DATA))) {
				if (wfd.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) {
					m_bFiltered = TRUE;
					return S_FALSE;
				}
			}
		}
		BSTR bs = NULL;
		BSTR bs2 = NULL;
		teGetDisplayNameBSTR(pSF, pidl, SHGDN_NORMAL, &bs);
		teGetDisplayNameBSTR(pSF, pidl, SHGDN_INFOLDER | SHGDN_FORPARSING, &bs2);
		if (m_ppDispatch[SB_OnIncludeObject]) {
			VARIANT vResult;
			VariantInit(&vResult);
			VARIANTARG *pv = GetNewVARIANT(4);
			teSetObject(&pv[3], this);
			teSetSZ(&pv[2], bs);
			teSetSZ(&pv[1], bs2);
			if (pidlFull) {
				teSetIDList(&pv[0], pidlFull);
			} else if (pidlFull = ILCombine(m_pidl, pidl)) {
				teSetIDListRelease(&pv[0], &pidlFull);
			}
			Invoke4(m_ppDispatch[SB_OnIncludeObject], &vResult, 4, pv);
			hr = GetIntFromVariantClear(&vResult);
		} else if (m_bsFilter) {
			if (!tePathMatchSpec(bs, m_bsFilter)) {
				hr = S_FALSE;
				if (lstrcmpi(bs, bs2) && tePathMatchSpec(bs2, m_bsFilter)) {
					hr = S_OK;
				}
			}
		}
		if (hr == S_OK && g_bsHiddenFilter && g_param[TE_UseHiddenFilter]) {
			if (tePathMatchSpec(bs, g_bsHiddenFilter) || tePathMatchSpec(bs2, g_bsHiddenFilter)) {
				hr = S_FALSE;
			}
		}
		if (hr == S_OK) {
			SetTimer(g_hwndMain, TET_Status, 500, teTimerProc);
		}
		teSysFreeString(&bs2);
		teSysFreeString(&bs);
	}
	if (hr != S_OK) {
		m_bFiltered = TRUE;
	}
	return hr;
}

int CteShellBrowser::GetSizeFormat()
{
	return m_nSizeFormat != -1 ? m_nSizeFormat : g_param[TE_SizeFormat];
}

VOID CteShellBrowser::SetObjectRect()
{
	RECT rc;
	GetWindowRect(m_hwnd, &rc);
	if (m_hwndDV) {
		MoveWindow(m_hwnd, m_rc.left, m_rc.top, m_rc.right - m_rc.left, m_rc.bottom - m_rc.top + 1, TRUE);
	}
	MoveWindow(m_hwnd, m_rc.left, m_rc.top, m_rc.right - m_rc.left, m_rc.bottom - m_rc.top, TRUE);
	if (m_hwndAlt) {
		MoveWindow(m_hwndAlt, 0, 0, m_rc.right - m_rc.left, m_rc.bottom - m_rc.top, TRUE);
	}
	if (m_hwndLV) {
		if (rc.right - rc.left != m_rc.right - m_rc.left || rc.bottom - rc.top != m_rc.bottom - m_rc.top) {
			int nFocused = ListView_GetNextItem(m_hwndLV, -1, LVNI_ALL | LVNI_FOCUSED | LVNI_SELECTED);
			if (nFocused >= 0) {
				PostMessage(m_hwndLV, LVM_ENSUREVISIBLE, nFocused, FALSE);
			}
		}
	}
}

int CteShellBrowser::GetFolderViewAndItemCount(IFolderView **ppFV, UINT uFlags)
{
	int iItems = 0;
	IFolderView *pFV = NULL;
	if (m_pShellView) {
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
			pFV->ItemCount(uFlags, &iItems);
		}
	}
	if (ppFV) {
		*ppFV = pFV;
	} else {
		SafeRelease(&pFV);
	}
	return iItems;
}

HRESULT CteShellBrowser::Search(LPWSTR pszSearch) {
	if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
		return S_OK;
	}
#ifdef _2000XP
	if (!_SHCreateShellItemArrayFromShellItem) {
		return E_NOTIMPL;
	}
#endif
	ISearchFolderItemFactory *psfif = NULL;
	if SUCCEEDED(teCreateInstance(CLSID_SearchFolderItemFactory, NULL, NULL, IID_PPV_ARGS(&psfif))) {
		psfif->Release();
		BSTR bsPath;
		if SUCCEEDED(m_pFolderItem->get_Path(&bsPath)) {
			if (teIsSearchFolder(bsPath)) {
				BSTR bs = NULL;
				teGetSearchArg(&bs, bsPath, L"&crumb=location:");
				if (bs) {
					teSysFreeString(&bsPath);
					bsPath = bs;
				}
			}
		}
		BSTR bsSearch = NULL;
		teCreateSearchPath(&bsSearch, bsPath, pszSearch);
		teSysFreeString(&bsPath);
		LPITEMIDLIST pidl = teILCreateFromPath(bsSearch);
		teSysFreeString(&bsSearch);
		if (pidl) {
			if (!ILIsEqual(m_pidl, pidl)) {
				BrowseObject(pidl, SBSP_SAMEBROWSER);
			}
			teILFreeClear(&pidl);
		}
		return S_OK;
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::OnViewCreated(IShellView *psv)
{
	m_bViewCreated = FALSE;
	if (psv) {
		if (m_pShellView) {
			if (m_pSFVCB) {
				teSetSFVCB(m_pShellView, m_pSFVCB, NULL);
				SafeRelease(&m_pSFVCB);
			}
			SafeRelease(&m_pShellView);
		}
		psv->QueryInterface(IID_PPV_ARGS(&m_pShellView));
		if (teCompareSFClass(m_pSF2, &CLSID_ShellFSFolder)) {
			teSetSFVCB(psv, (IShellFolderViewCB *)this, &m_pSFVCB);
		}
	}
	GetShellFolderView();
	if (m_pExplorerBrowser) {
		IUnknown_GetWindow(m_pExplorerBrowser, &m_hwnd);
	}
	m_bNavigateComplete = TRUE;
	m_dwSessionId = MAKELONG(GetTickCount(), rand());
	SetPropEx();
	if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
		++m_nGroupByDelay;
	}
#ifdef _2000XP
	if (g_bCharWidth) {
		if (m_hwndLV) {
			HWND hHeader = ListView_GetHeader(m_hwndLV);
			if (hHeader) {
				WCHAR szText[MAX_COLUMN_NAME_LEN];
				HD_ITEM hdi = { HDI_TEXT | HDI_WIDTH, 0, szText, NULL, MAX_COLUMN_NAME_LEN };
				Header_GetItem(hHeader, 0, &hdi);
				int cxChar = 0;
				PSGetColumnIndexXP(hdi.pszText, &cxChar);
				if (cxChar) {
					g_nCharWidth = hdi.cxy / cxChar;
				}
				g_bCharWidth = FALSE;
			}
		}
	}
#endif
	SetTabName();
	SetTheme();
	if (m_bVisible && m_hwndDV) { // #290
		RECT rc1, rc2;
		GetWindowRect(m_hwnd, &rc1);// 5ch#6-896
		GetWindowRect(m_hwndDV, &rc2);
		if (rc1.left == rc1.right || rc2.left == rc2.right) {
			ArrangeWindow();
		}
	}
	m_bSetRedraw = TRUE;
	SetRedraw(TRUE);
	if (!m_pTC->m_nLockUpdate) {
		ArrangeWindowEx();
	}
	return S_OK;
}

VOID CteShellBrowser::SetTabName()
{
	//View Tab Name
	int i = GetTabIndex();
	if (i >= 0) {
		SafeRelease(&m_ppDispatch[SB_ColumnsReplace]);
		DoFunc(TE_OnViewCreated, this, E_NOTIMPL);
		InitFolderSize();
	}
	SetStatusTextSB(NULL);
}

STDMETHODIMP CteShellBrowser::OnNavigationComplete(PCIDLIST_ABSOLUTE pidlFolder)
{
	if (!teILIsFileSystemEx(pidlFolder) || teILIsSearchFolder(pidlFolder)) {
		OnNavigationComplete2();
	}
	return S_OK;
}

VOID CteShellBrowser::OnNavigationComplete2()
{
	if (!m_bNavigateComplete) {
		return;
	}
	m_bNavigateComplete = FALSE;
	SetViewModeAndIconSize(TRUE);
	m_bIconSize = TRUE;
	m_bSetRedraw = TRUE;
	SetRedraw(TRUE);
	SetActive(FALSE);
	m_nFocusItem = 1;
	NavigateComplete(TRUE);
	if (m_nFocusItem > 0) {
		FocusItem();
	}
	m_bCheckLayout = TRUE;
	SetFolderFlags(FALSE);
	InitFolderSize();
	if (!m_nGroupByDelay
#ifdef _2000XP
		|| !g_bUpperVista
#endif
	) {
		m_pTC->RedrawUpdate();
	}
}

HRESULT CteShellBrowser::BrowseToObject()
{
	m_nSuspendMode = 2;
	::InterlockedIncrement(&m_pTC->m_lNavigating);
	HRESULT hr = m_pExplorerBrowser->BrowseToObject(m_pSF2, SBSP_ABSOLUTE);
	::InterlockedDecrement(&m_pTC->m_lNavigating);
	if SUCCEEDED(hr) {
		IUnknown_GetWindow(m_pExplorerBrowser, &m_hwnd);
		if (!m_pShellView) {
			m_pExplorerBrowser->GetCurrentView(IID_PPV_ARGS(&m_pShellView));
		}
	}
	return hr;
}

HRESULT CteShellBrowser::GetShellFolder2(LPITEMIDLIST *ppidl)
{
	HRESULT hr = E_FAIL;
	IShellFolder *pSF = NULL;
	GetShellFolder(&pSF, *ppidl);
	if (g_param[TE_LibraryFilter] && pSF && teCompareSFClass(pSF, &CLSID_LibraryFolder)) {//The library cannot be filtered, so convert it to an actual folder.
		BSTR bs;
		teGetDisplayNameFromIDList(&bs, *ppidl, SHGDN_FORPARSING);
		if (teIsFileSystem(bs)) {
			LPITEMIDLIST pidl2 = teILCreateFromPath(bs);
			if (pidl2) {
				IShellFolder *pSF1;
				if (GetShellFolder(&pSF1, pidl2)) {
					teILCloneReplace(ppidl, pidl2);
					SafeRelease(&pSF);
					pSF = pSF1;
				}
				teILFreeClear(&pidl2);
			}
		}
		::SysFreeString(bs);
	}
	if (!pSF) {
		GetShellFolder(&pSF, g_pidls[CSIDL_RESULTSFOLDER]);
		teILCloneReplace(ppidl, g_pidls[CSIDL_RESULTSFOLDER]);
		m_dwUnavailable = GetTickCount();
	}
	if (pSF) {
		SafeRelease(&m_pSF2);
		hr = pSF->QueryInterface(IID_PPV_ARGS(&m_pSF2));
		pSF->Release();
		if FAILED(hr) {
			GetShellFolder(&pSF, g_pidls[CSIDL_RESULTSFOLDER]);
			teILCloneReplace(ppidl, g_pidls[CSIDL_RESULTSFOLDER]);
			m_dwUnavailable = GetTickCount();
			hr = pSF->QueryInterface(IID_PPV_ARGS(&m_pSF2));
		}
	}
	return hr;
}

HRESULT CteShellBrowser::RemoveAll()
{
	HRESULT hr = E_NOTIMPL;
	if (m_pShellView) {
		IFolderView2 *pFV2 = NULL;
		IResultsFolder *pRF;
		m_nGroupByDelay = 1;
		m_dwSessionId = MAKELONG(GetTickCount(), rand());
		if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) && SUCCEEDED(pFV2->GetFolder(IID_PPV_ARGS(&pRF)))) {
			hr = pRF->RemoveAll();
			pRF->Release();
			pFV2->SetGroupBy(PKEY_Null, TRUE);
			pFV2->SetSortColumns(g_pSortColumnNull, _countof(g_pSortColumnNull));
#ifdef _2000XP
		} else {
			IShellFolderView *pSFV;
			if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
				UINT ui;
				hr = pSFV->RemoveObject(NULL, &ui);//the removal of all objects
				pSFV->Release();
			}
#endif
		}
		SafeRelease(&pFV2);
	}
	return hr;
}

VOID CteShellBrowser::SetViewModeAndIconSize(BOOL bSetIconSize)
{
	IFolderView2 *pFV2;
	if (m_pShellView) {
		if (bSetIconSize && SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2)))) {
			FOLDERVIEWMODE uViewMode;
			int iImageSize;
			if (FAILED(pFV2->GetViewModeAndIconSize(&uViewMode, &iImageSize)) || uViewMode != m_param[SB_ViewMode] || iImageSize != m_param[SB_IconSize]) {
				pFV2->SetViewModeAndIconSize(static_cast<FOLDERVIEWMODE>(m_param[SB_ViewMode]), m_param[SB_IconSize]);
			}
			pFV2->Release();
		} else {
			IFolderView *pFV;
			if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
				pFV->SetCurrentViewMode(m_param[SB_ViewMode]);
				pFV->Release();
			}
		}
	}
}

VOID CteShellBrowser::SetListColumnWidth()
{
	if (m_bSetListColumnWidth) {
		if (m_pSF2 && m_hwndLV && m_param[SB_ViewMode] == FVM_LIST) {
			m_bSetListColumnWidth = FALSE;
			int nOrgWidth = ListView_GetColumnWidth(m_hwndLV, 0);
			IFolderView *pFV = NULL;
			int nWidth = nOrgWidth - 28;
			int n = GetFolderViewAndItemCount(&pFV, SVGIO_ALLVIEW);
			while (--n >= 0) {
				LPITEMIDLIST pidl;
				if SUCCEEDED(pFV->Item(n, &pidl)) {
					VARIANT v;
					VariantInit(&v);
					if SUCCEEDED(m_pSF2->GetDetailsEx(pidl, &PKEY_ItemName, &v)) {
						if (v.vt == VT_BSTR) {
							int nWidth1 = ListView_GetStringWidth(m_hwndLV, v.bstrVal);
							if (nWidth1 > nWidth) {
								nWidth = nWidth1;
							}
						}
						VariantClear(&v);
					}
					teCoTaskMemFree(pidl);
				}
			}
			nWidth += 28;
			if (nWidth > nOrgWidth) {
				ListView_SetColumnWidth(m_hwndLV, 0, 1);
				ListView_SetColumnWidth(m_hwndLV, 0, nWidth);
				ListView_EnsureVisible(m_hwndLV, ListView_GetNextItem(m_hwndLV, -1, LVNI_ALL | LVNI_FOCUSED), FALSE);
			}
			SafeRelease(&pFV);
		}
	}
}

VOID CteShellBrowser::AddItem(LPITEMIDLIST pidl)
{
	HRESULT hr = S_OK;
	IFolderView *pFV = NULL;
	IShellFolder *pSF;
    LPCITEMIDLIST pidlPart;
    if (!ILIsEmpty(pidl)) {
		if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
		    LPITEMIDLIST pidlChild = NULL;
			if (IncludeObject2(pSF, pidlPart, pidl) == S_OK) {
				m_bRedraw = TRUE;
				SetRedraw(FALSE);
				try {
					IResultsFolder *pRF;
					if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) && SUCCEEDED(pFV->GetFolder(IID_PPV_ARGS(&pRF)))) {
						BSTR bs2 = NULL;
						if SUCCEEDED(teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_FORPARSING, &bs2)) {
							if (teIsFileSystem(bs2)) {
								SFGAOF sfAttr = SFGAO_FILESYSTEM | SFGAO_FOLDER;
								if FAILED(pSF->GetAttributesOf(1, &pidlPart, &sfAttr)) {
									sfAttr = 0;
								}
								if (!(sfAttr & SFGAO_FILESYSTEM)) {
									WIN32_FIND_DATA wfd;
									teSHGetDataFromIDList(pSF, pidlPart, SHGDFIL_FINDDATA, &wfd, sizeof(WIN32_FIND_DATA));
									teILFreeClear(&pidl);
									pidl = teSHSimpleIDListFromPathEx(bs2, wfd.dwFileAttributes | ((sfAttr & SFGAO_FOLDER) ? FILE_ATTRIBUTE_DIRECTORY : 0), wfd.nFileSizeLow, wfd.nFileSizeHigh, &wfd.ftLastWriteTime);
									m_bRegenerateItems = TRUE;
								}
							}
							teSysFreeString(&bs2);
						}
						pRF->RemoveIDList(pidl);
						pRF->AddIDList(pidl, &pidlChild);
						pRF->Release();
#ifdef _2000XP
					} else if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
						IShellFolderView *pSFV;
						if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
							pidlChild = teILCreateResultsXP(pidl);
							UINT ui;
							pSFV->RemoveObject(pidlChild, &ui);
							pSFV->AddObject(pidlChild, &ui);
							pSFV->Release();
						}
#endif
					}
				} catch (...) {
					g_nException = 0;
#ifdef _DEBUG
					g_strException = L"AddItem";
#endif
				}
				int i = GetFolderViewAndItemCount(NULL, SVGIO_ALLVIEW);
				if ((i % 100) == 0) {
					SetRedraw(TRUE);
				}
			}
			SafeRelease(&pFV);
			teCoTaskMemFree(pidlChild);
		}
		pSF->Release();
    }
	teCoTaskMemFree(pidl);
}

BOOL CteShellBrowser::HasFilter()
{
	return m_bsFilter || m_ppDispatch[SB_OnIncludeObject]|| (g_bsHiddenFilter && g_param[TE_UseHiddenFilter]);
}

VOID CteShellBrowser::NavigateComplete(BOOL bBeginNavigate)
{
	if (!bBeginNavigate || DoFunc(TE_OnBeginNavigate, this, S_OK) != S_FALSE) {
		BSTR bsAltSortColumn = NULL;
		if (m_bRefreshing) {
			m_bRefreshing = FALSE;
			if (m_bsAltSortColumn && !m_bNavigateComplete && ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
				bsAltSortColumn= ::SysAllocString(m_bsAltSortColumn);
			}
		}
		DoFunc(TE_OnNavigateComplete, this, E_NOTIMPL);
		if (bsAltSortColumn) {
			SetSort(bsAltSortColumn);
			::SysFreeString(bsAltSortColumn);
			if (m_nGroupByDelay) {
				m_bRefreshing = TRUE;
			}
		}
		if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
			ArrangeWindow();
		}
	}
}

VOID CteShellBrowser::InitFilter()
{
	teSysFreeString(&m_bsFilter);
	SafeRelease(&m_ppDispatch[SB_OnIncludeObject]);
}

HRESULT CteShellBrowser::SetTheme()
{
	if (_AllowDarkModeForWindow) {
		BOOL bDarkMode = teIsDarkColor(m_clrBk);
		HWND hHeader = ListView_GetHeader(m_hwndLV);
		if (hHeader) {
			_AllowDarkModeForWindow(hHeader, bDarkMode);
			SetWindowTheme(hHeader, bDarkMode ? L"darkmode_itemsview" : L"explorer", NULL);
		}
		_AllowDarkModeForWindow(m_hwndLV, bDarkMode);
	}
	return SetWindowTheme(m_hwndLV, GetThemeName(), NULL);
}

LPWSTR CteShellBrowser::GetThemeName()
{
	BOOL bDarkMode = _AllowDarkModeForWindow && teIsDarkColor(m_clrBk);
	if (g_nWindowTheme == 0) { //Normal style
		return L"explorer";
	}
	if (g_nWindowTheme == 1) { //Classic style
		return bDarkMode ? L"darkmode_explorer" : NULL;
	}
	//Items view style
	return bDarkMode ? L"darkmode_itemsview" : L"itemsview";
}

VOID CteShellBrowser::FixColumnEmphasis()
{
	if (!g_param[TE_ColumnEmphasis] && (int)ListView_GetSelectedColumn(m_hwndLV) >= 0) {
		ListView_SetSelectedColumn(m_hwndLV, -1);
	}
}

STDMETHODIMP CteShellBrowser::OnNavigationFailed(PCIDLIST_ABSOLUTE pidlFolder)
{
	if (m_nSuspendMode) {
		Suspend(m_nSuspendMode);
	}
	return S_OK;
}

//IExplorerPaneVisibility
STDMETHODIMP CteShellBrowser::GetPaneState(REFEXPLORERPANE ep, EXPLORERPANESTATE *peps)
{
	if (g_pOnFunc[TE_OnGetPaneState]) {
		VARIANTARG *pv = GetNewVARIANT(3);
		teSetObject(&pv[2], this);
		pv[1].vt = VT_BSTR;
		pv[1].bstrVal = ::SysAllocStringLen(NULL, 38);
		StringFromGUID2(ep, pv[1].bstrVal, 39);
		teSetObjectRelease(&pv[0], new CteMemory(sizeof(int), peps, 1, L"DWORD"));
		VARIANT vResult;
		VariantInit(&vResult);
		Invoke4(g_pOnFunc[TE_OnGetPaneState], &vResult, 3, pv);
		if (vResult.vt != VT_EMPTY) {
			return GetIntFromVariantClear(&vResult);
		}
	}
	return E_NOTIMPL;
}

#if defined(_USE_SHELLBROWSER) || defined(_2000XP)
//IShellFolder
STDMETHODIMP CteShellBrowser::ParseDisplayName(HWND hwnd, IBindCtx *pbc, LPWSTR pszDisplayName, ULONG *pchEaten, PIDLIST_RELATIVE *ppidl, ULONG *pdwAttributes)
{
	return m_pSF2->ParseDisplayName(hwnd, pbc, pszDisplayName, pchEaten, ppidl, pdwAttributes);
}

STDMETHODIMP CteShellBrowser::EnumObjects(HWND hwndOwner, SHCONTF grfFlags, IEnumIDList **ppenumIDList)
{
	return m_pSF2->EnumObjects(hwndOwner, grfFlags, ppenumIDList);
}

STDMETHODIMP CteShellBrowser::BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppvOut)
{
	return m_pSF2->BindToObject(pidl, pbc, riid, ppvOut);
}

STDMETHODIMP CteShellBrowser::BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppvOut)
{
	return m_pSF2->BindToStorage(pidl, pbc, riid, ppvOut);
}

STDMETHODIMP CteShellBrowser::CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2)
{
	return m_pSF2->CompareIDs(lParam, pidl1, pidl2);
}

STDMETHODIMP CteShellBrowser::CreateViewObject(HWND hwndOwner, REFIID riid, void **ppv)
{
	if (IsEqualIID(riid, IID_IShellView)) {
		SafeRelease(&m_pSFVCB);
		IShellView *pSV;
		if SUCCEEDED(m_pSF2->CreateViewObject(hwndOwner, IID_PPV_ARGS(&pSV))) {
			if (teSetSFVCB(pSV, NULL, &m_pSFVCB) == S_OK) {
				*ppv = pSV;
				return S_OK;
			}
			pSV->Release();
		}
		SFV_CREATE sfvc;
		sfvc.cbSize = sizeof(SFV_CREATE);
		sfvc.pshf = static_cast<IShellFolder *>(this);
		sfvc.psfvcb = static_cast<IShellFolderViewCB *>(this);
		sfvc.psvOuter = NULL;
		return SHCreateShellFolderView(&sfvc, (IShellView **)ppv);
	}
	return m_pSF2->CreateViewObject(hwndOwner, riid, ppv);
}

STDMETHODIMP CteShellBrowser::GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, SFGAOF *rgfInOut)
{
	return m_pSF2->GetAttributesOf(cidl, apidl, rgfInOut);
}

STDMETHODIMP CteShellBrowser::GetUIObjectOf(HWND hwndOwner, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid, UINT *rgfReserved, void **ppv)
{
	return m_pSF2->GetUIObjectOf(hwndOwner, cidl, apidl, riid, rgfReserved, ppv);
}

STDMETHODIMP CteShellBrowser::GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET *pName)
{
	if (m_pSF2) {
		return m_pSF2->GetDisplayNameOf(pidl, uFlags, pName);
	}
	return E_ACCESSDENIED;
}

STDMETHODIMP CteShellBrowser::SetNameOf(HWND hwndOwner, PCUITEMID_CHILD pidl, LPCWSTR pszName, SHGDNF uFlags, PITEMID_CHILD *ppidlOut)
{
	return m_pSF2->SetNameOf(hwndOwner, pidl, pszName, uFlags, ppidlOut);
}

#endif
#ifdef _2000XP

//IShellFolder2
STDMETHODIMP CteShellBrowser::GetDefaultSearchGUID(GUID *pguid)
{
	return m_pSF2->GetDefaultSearchGUID(pguid);
}

STDMETHODIMP CteShellBrowser::EnumSearches(IEnumExtraSearch **ppEnum)
{
	return m_pSF2->EnumSearches(ppEnum);
}

STDMETHODIMP CteShellBrowser::GetDefaultColumn(DWORD dwReserved, ULONG *pSort, ULONG *pDisplay)
{
	return m_pSF2->GetDefaultColumn(dwReserved, pSort, pDisplay);
}

STDMETHODIMP CteShellBrowser::GetDefaultColumnState(UINT iColumn, SHCOLSTATEF *pcsFlags)
{
	if (m_nColumns) {
		if (iColumn < m_nColumns) {
			*pcsFlags = m_pColumns[iColumn].csFlags;
			return S_OK;
		}
	}
	return m_pSF2->GetDefaultColumnState(iColumn, pcsFlags);
}

STDMETHODIMP CteShellBrowser::GetDetailsEx(PCUITEMID_CHILD pidl, const SHCOLUMNID *pscid, VARIANT *pv)
{
	int iColumn = -1;
	if (m_nColumns > 1 && IsEqualPropertyKey(*pscid, PKEY_TotalFileSize)) {
		iColumn = m_nColumns - 1;
	} else if (m_nFolderName != MAXINT && IsEqualPropertyKey(*pscid, PKEY_ItemFolderNameDisplay)) {
		iColumn = m_nFolderName;
	}
	if (iColumn < 0) {
		return m_pSF2->GetDetailsEx(pidl, pscid, pv);
	}
	SHELLDETAILS sd;
	HRESULT hr = GetDetailsOf(pidl, iColumn, &sd);
	if SUCCEEDED(hr) {
		hr = StrRetToBSTR(&sd.str, pidl, &pv->bstrVal);
		if SUCCEEDED(hr) {
			pv->vt = VT_BSTR;
		}
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::GetDetailsOf(PCUITEMID_CHILD pidl, UINT iColumn, SHELLDETAILS *psd)
{
	HRESULT hr;
	if (iColumn < m_nColumns - 2 || iColumn >= m_nColumns) {
		if (pidl && iColumn == m_nFolderName && SUCCEEDED(m_pSF2->GetDisplayNameOf(pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &psd->str))) {
			PathRemoveFileSpec(psd->str.pOleStr);
			hr = S_OK;
		} else {
			hr = m_pSF2->GetDetailsOf(pidl, iColumn, psd);
		}
	} else {
		WCHAR szBuf[MAX_COLUMN_NAME_LEN];
		szBuf[0] = NULL;
		if (pidl) {
			if (iColumn == m_nColumns - 1) {
				m_nFolderSizeIndex = MAXINT - 1;
				SetFolderSize(m_pSF2, pidl, szBuf, _countof(szBuf));
			} else {
				SetLabel(pidl, szBuf, MAX_COLUMN_NAME_LEN);
			}
		} else {
			if (iColumn == m_nColumns - 1) {
				lstrcpy(szBuf, g_szTotalFileSizeXP);
				psd->fmt = LVCFMT_RIGHT;
				psd->cxChar = 16;
			} else {
				lstrcpy(szBuf, g_szLabelXP);
				psd->fmt = LVCFMT_LEFT;
				psd->cxChar = 20;
			}
		}
		psd->str.uType = STRRET_WSTR;
		psd->str.pOleStr = (LPOLESTR)::CoTaskMemAlloc((lstrlen(szBuf) + 1) * sizeof(WCHAR));
		lstrcpy(psd->str.pOleStr, szBuf);
		hr = S_OK;
	}
	if (m_nColumns) {
		if (iColumn < m_nColumns) {
			if (m_pColumns[iColumn].nWidth >= 0 && m_pColumns[iColumn].nWidth < 32768) {
				psd->cxChar = m_pColumns[iColumn].nWidth / g_nCharWidth;
			}
		}
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::MapColumnToSCID(UINT iColumn, SHCOLUMNID *pscid)
{
	if (iColumn == m_nColumns - 1) {
		*pscid = PKEY_TotalFileSize;
		return S_OK;
	}
	if (iColumn == m_nColumns - 2) {
		*pscid = PKEY_Contact_Label;
		return S_OK;
	}
	return m_pSF2->MapColumnToSCID(iColumn, pscid);
}
#endif

//IShellFolderViewCB
STDMETHODIMP CteShellBrowser::MessageSFVCB(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (m_pSF2) {
		for (size_t i = 0; i < g_ppMessageSFVCB.size(); ++i) {
			LPFNMessageSFVCB _MessageSFVCB = (LPFNMessageSFVCB)g_ppMessageSFVCB[i];
			HRESULT hr = _MessageSFVCB(this, m_pSF2, m_pShellView, uMsg, wParam, lParam);
			if (hr != S_OK) {
				return hr;
			}
		}
		if (uMsg == SFVM_FSNOTIFY) {
			if (!g_param[TE_AutoArrange]) {
				if (lParam & (SHCNE_CREATE | SHCNE_MKDIR)) {
					m_dwTickNotify = GetTickCount();
				}
				if (lParam & SHCNE_UPDATEITEM) {
					try {
						if (ILIsEqual(m_pidl, *(LPITEMIDLIST *)wParam)) {
							if (m_dwTickNotify && GetTickCount() - m_dwTickNotify < 500) {
								m_dwTickNotify = 0;
								return S_FALSE;
							}
							int iExists = GetFolderViewAndItemCount(NULL, SVGIO_ALLVIEW);
							int iNew = 0;
							if (iExists > 99999) {
								return S_FALSE;
							}
							SHCONTF grfFlags = SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN;
							HKEY hKey;
							if (RegOpenKeyExA(HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
								DWORD dwData;
								DWORD dwSize = sizeof(DWORD);
								if (RegQueryValueExA(hKey, "ShowSuperHidden", NULL, NULL, (LPBYTE)&dwData, &dwSize) == S_OK) {
									if (dwData) {
										grfFlags |= SHCONTF_INCLUDESUPERHIDDEN;
									}
								}
							}
							IEnumIDList *peidl;
							if SUCCEEDED(m_pSF2->EnumObjects(NULL, grfFlags, &peidl)) {
								LPITEMIDLIST pidlPart = NULL;
								while (iNew <= iExists && peidl->Next(1, &pidlPart, NULL) == S_OK) {
									if (IncludeObject2(m_pSF2, pidlPart, NULL) == S_OK) {
										iNew++;
									}
									teCoTaskMemFree(pidlPart);
								}
								peidl->Release();
							}
							if (iExists == iNew) {
								return S_FALSE;
							}
						}
					} catch (...) {
						g_nException = 0;
#ifdef _DEBUG
						g_strException = L"MessageSFVCB";
#endif
					}
				}
			}
		}
	}
	if (!m_pSFVCB) {
		return E_NOTIMPL;
	}
#ifdef _2000XP
	if (!g_bUpperVista) {
		if (uMsg == 54 || uMsg == 78) {
			return E_NOTIMPL;
		}
	}
#endif
	return m_pSFVCB->MessageSFVCB(uMsg, wParam, lParam);
}

//IPersist
STDMETHODIMP CteShellBrowser::GetClassID(CLSID *pClassID)
{
	IPersist *pPersist;
	HRESULT hr = m_pSF2->QueryInterface(IID_PPV_ARGS(&pPersist));
	if SUCCEEDED(hr) {
		hr = pPersist->GetClassID(pClassID);
		pPersist->Release();
	}
	return hr;
}

//IPersistFolder
STDMETHODIMP CteShellBrowser::Initialize(PCIDLIST_ABSOLUTE pidl)
{
	return BrowseObject(pidl, SBSP_ABSOLUTE | SBSP_SAMEBROWSER);
}

//IPersistFolder2
STDMETHODIMP CteShellBrowser::GetCurFolder(PIDLIST_ABSOLUTE *ppidl)
{
	return teGetIDListFromObject(m_pFolderItem, ppidl) ? S_OK : E_FAIL;
}

//

VOID CteShellBrowser::Suspend(int nMode)
{
	BOOL bVisible = m_bVisible;
	if (nMode == 2) {
		if (m_dwUnavailable || ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
			if (IsWindowEnabled(m_hwnd)) {
				return;
			}
		}
		Show(FALSE, 0);
		m_dwUnavailable = GetTickCount();
		teILCloneReplace(&m_pidl, g_pidls[CSIDL_RESULTSFOLDER]);
	}
	if (!m_dwUnavailable) {
		SaveFocusedItemToHistory();
		teDoCommand(this, m_hwnd, WM_NULL, 0, 0);//Save folder setings
	}
	m_pTC->LockUpdate();
	try {
		DestroyView(0);
		if (nMode == 4) {
			m_nUnload = 8;
		} else {
			m_nUnload = 9;
			if (nMode == 1 && m_pTV) {
				m_pTV->Close();
				m_pTV->m_bSetRoot = TRUE;
			}
			if (nMode != 2) {
				CteFolderItem *pid1 = NULL;
				if SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) {
					if (pid1->m_v.vt == VT_BSTR) {
						teILFreeClear(&m_pidl);
						pid1->Clear();
						pid1->Release();
					}
				}
			}
			if (m_pidl) {
				VARIANT v;
				if (GetVarPathFromFolderItem(m_pFolderItem, &v)) {
					CteFolderItem *pid1;
					teQueryFolderItem(&m_pFolderItem, &pid1);
					if (nMode == 2) {
						pid1->MakeUnavailable();
						m_nUnload = 8;
					} else {
						pid1->Clear();
						VariantClear(&pid1->m_v);
						VariantCopy(&pid1->m_v, &v);
					}
					pid1->Release();
					teILFreeClear(&m_pidl);
					VariantClear(&v);
				}
			}
		}
		Show(bVisible, 0);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"Suspend";
#endif
	}
	m_pTC->UnlockUpdate();
	ArrangeWindow();
}

VOID CteShellBrowser::SetPropEx()
{
	if (m_pShellView && m_pShellView->GetWindow(&m_hwndDV) == S_OK) {
		if (SetWindowLongPtr(m_hwndDV, GWLP_USERDATA, (LONG_PTR)this) != (LONG_PTR)this) {
			SetWindowSubclass(m_hwndDV, TELVProc, 1, 0);
			for (int i = WM_USER + 173; i <= WM_USER + 175; ++i) {
				teChangeWindowMessageFilterEx(m_hwndDV, i, MSGFLT_ALLOW, NULL);
			}
		}
		m_hwndLV = FindWindowExA(m_hwndDV, 0, WC_LISTVIEWA, NULL);
		m_hwndDT = m_hwndLV;
		if (m_param[SB_Type] != CTRL_EB || m_param[SB_FolderFlags] & FWF_NOCLIENTEDGE) {
			teSetExStyleAnd(m_hwnd, ~WS_EX_CLIENTEDGE);
		}
		if (m_param[SB_Type] <= CTRL_SB) {
			if (m_param[SB_FolderFlags] & FWF_NOCLIENTEDGE) {
				SetWindowLong(m_hwnd, GWL_STYLE, GetWindowLong(m_hwnd, GWL_STYLE) & ~WS_BORDER);
			} else {
				SetWindowLong(m_hwnd, GWL_STYLE, GetWindowLong(m_hwnd, GWL_STYLE) | WS_BORDER);
			}
		}
		if (m_hwndLV) {
			SetWindowLong(m_hwndLV, GWL_STYLE, GetWindowLong(m_hwndLV, GWL_STYLE) & ~WS_BORDER);
			if (m_param[SB_Type] <= CTRL_SB || m_param[SB_FolderFlags] & FWF_NOCLIENTEDGE) {
				teSetExStyleAnd(m_hwndLV, ~WS_EX_CLIENTEDGE);
			} else {
				teSetExStyleOr(m_hwndLV, WS_EX_CLIENTEDGE);
			}
			SendMessage(m_hwndLV, WM_CHANGEUISTATE, MAKELONG(UIS_SET, UISF_HIDEFOCUS), 0);
			if (SetWindowLongPtr(m_hwndLV, GWLP_USERDATA, (LONG_PTR)this) != (LONG_PTR)this) {
				SetWindowSubclass(m_hwndLV, TELVProc2, 1, 0);
			}
			FixColumnEmphasis();
		} else {
			m_hwndDT = FindWindowExA(m_hwndDV, NULL, "DirectUIHWND", NULL);
		}
		if (!m_pDropTarget2) {
			m_pDropTarget2 = new CteDropTarget2(m_hwndDT, static_cast<IDispatch *>(this), FALSE);
			teRegisterDragDrop(m_hwndDT, m_pDropTarget2, &m_pDropTarget2->m_pDropTarget);
		}
		m_hwndAlt = NULL;
	}
}

VOID CteShellBrowser::ResetPropEx()
{
	RemoveWindowSubclass(m_hwndLV, TELVProc2, 1);
	SetWindowLongPtr(m_hwndLV, GWLP_USERDATA, 0);
	RemoveWindowSubclass(m_hwndDV, TELVProc, 1);
	SetWindowLongPtr(m_hwndDV, GWLP_USERDATA, 0);
	if (m_pDropTarget2) {
//		SetProp(m_hwndDT, L"OleDropTargetInterface", (HANDLE)m_pDropTarget2->m_pDropTarget);
		RevokeDragDrop(m_hwndDT);
		RegisterDragDrop(m_hwndDT, m_pDropTarget2->m_pDropTarget);
		SafeRelease(&m_pDropTarget2);
	}
}

void CteShellBrowser::Show(BOOL bShow, DWORD dwOptions)
{
	bShow &= 1;
	if ((bShow ^ m_bVisible) || (bShow && m_nUnload)) {
		if (bShow) {
			if (g_nLockUpdate <= 1 && !m_nCreate) {
				if (m_nUnload & 8) {
					if (m_nUnload & 1) {
						CteFolderItem *pid;
						if SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid)) {
							if (pid->m_v.vt == VT_BSTR) {
								pid->Clear();
							}
							pid->Release();
						}
					}
					m_nUnload = 2;
					if SUCCEEDED(BrowseObject(NULL, SBSP_RELATIVE | SBSP_WRITENOHISTORY | SBSP_REDIRECT)) {
						m_nUnload = 0;
					}
				} else if (m_nUnload == 1 || !m_pShellView) {
					m_nUnload = 2;
					if SUCCEEDED(BrowseObject(NULL, SBSP_RELATIVE | SBSP_SAMEBROWSER)) {
						m_nUnload = 0;
					}
				} else if (m_bRefreshLator) {
					Refresh(TRUE);
				}
			}
		}
		if (m_pShellView) {
			m_bVisible = bShow;
			if (bShow) {
				ShowWindow(m_hwnd, SW_SHOWNA);
				if (m_hwndAlt) {
					ShowWindow(m_hwndAlt, SW_SHOWNA);
				} else if (m_hwndLV) {
					ShowWindow(m_hwndLV, SW_SHOWNA);
				}
				SetRedraw(TRUE);
				if (m_param[SB_TreeAlign] & 2) {
					m_pTV->Show();
					ShowWindow(m_pTV->m_hwnd, SW_SHOWNA);
					BringWindowToTop(m_pTV->m_hwnd);
				}
				BringWindowToTop(m_hwnd);
				ArrangeWindow();
				if (m_nUnload == 4) {
					Refresh(FALSE);
				}
			} else if ((dwOptions & 2) == 0) {
				SetRedraw(FALSE);
				m_pShellView->UIActivate(SVUIA_DEACTIVATE);
				ShowWindow(m_hwnd, SW_HIDE);
				if (m_hwndAlt) {
					ShowWindow(m_hwndAlt, SW_HIDE);
				}
				if (m_hwndLV) {
					ShowWindow(m_hwndLV, SW_HIDE);
				}
				if (m_pTV->m_hwnd) {
					ShowWindow(m_pTV->m_hwnd, SW_HIDE);
				}
				if ((dwOptions & 1) && m_dwUnavailable && !m_nCreate && !ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
					Suspend(0);
				}
				ArrangeWindow();
			}
		} else {
			m_bVisible = FALSE;
		}
	}
}

BOOL CteShellBrowser::Close(BOOL bForce)
{
	if ((CanClose(this) || bForce) && !m_pTC->m_lNavigating) {
		int i = GetTabIndex();
		if (i >= 0) {
			TC_ITEM tcItem;
			::ZeroMemory(&tcItem, sizeof(TC_ITEM));
			tcItem.mask = TCIF_PARAM;
			TabCtrl_SetItem(m_pTC->m_hwnd, i, &tcItem);
			SendMessage(m_pTC->m_hwnd, TCM_DELETEITEM, i, 0);
		}
		ShowWindow(m_pTV->m_hwnd, SW_HIDE);
		Clear();
		m_bEmpty = TRUE;
		return TRUE;
	}
	return FALSE;
}

VOID CteShellBrowser::DestroyView(int nFlags)
{
	KillTimer(m_hwnd, (UINT_PTR)this);
	ResetPropEx();
	BOOL bCloseSB = TRUE;
	if ((nFlags & 2) == 0 && m_pExplorerBrowser) {
		try {
			bCloseSB = FALSE;
			if (m_dwEventCookie) {
			  m_pExplorerBrowser->Unadvise(m_dwEventCookie);
			}
			IUnknown_SetSite(m_pExplorerBrowser, NULL);
			Show(FALSE, 0);
			m_pExplorerBrowser->Destroy();
			m_pExplorerBrowser->Release();
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"DestroyView";
#endif
		}
		m_pExplorerBrowser = NULL;
	}
	if (m_pShellView) {
		try {
			if (m_pSFVCB) {
				teSetSFVCB(m_pShellView, m_pSFVCB, NULL);
				SafeRelease(&m_pSFVCB);
			}
			if (bCloseSB && (nFlags & 1) == 0) {
				Show(FALSE, 0);
				m_pShellView->DestroyViewWindow();
			}
			if (nFlags == 0) {
				SafeRelease(&m_pShellView);
			}
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"DestroyView/2";
#endif
		}
	}
}

VOID CteShellBrowser::GetSort(BSTR* pbs, int nFormat)
{
	*pbs = NULL;
	if (!m_pShellView) {
		return;
	}
	if (m_bsAltSortColumn) {
		*pbs = ::SysAllocString(m_bsAltSortColumn);
		return;
	}
	GetSort2(pbs, nFormat);
}

VOID CteShellBrowser::GetSort2(BSTR* pbs, int nFormat)
{
	IFolderView2 *pFV2;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		if (teIsSameSort(pFV2, g_pSortColumnNull, _countof(g_pSortColumnNull))) {
			*pbs = ::SysAllocString(L"System.Null");
		} else {
			SORTCOLUMN pCol[1];
			if SUCCEEDED(pFV2->GetSortColumns(pCol, _countof(pCol))) {
				*pbs = tePSGetNameFromPropertyKeyEx(pCol[0].propkey, nFormat, this);
				if (pCol[0].direction < 0) {
					ToMinus(pbs);
				}
			}
		}
		pFV2->Release();
		return;
	}
#ifdef _2000XP
	LPARAM Sort = 0;
	IShellFolderView *pSFV;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
		if SUCCEEDED(pSFV->GetArrangeParam(&Sort)) {//Sort column
			Sort = LOWORD(Sort);
		}
		pSFV->Release();
	}
	SHCOLUMNID scid;
	if SUCCEEDED(MapColumnToSCID((UINT)Sort, &scid)) {
		*pbs = tePSGetNameFromPropertyKeyEx(scid, nFormat, this);
		HWND hHeader = ListView_GetHeader(m_hwndLV);
		if (hHeader) {
			int nCount = Header_GetItemCount(hHeader);
			for (int i = 0; i < nCount; ++i) {
#ifdef _W2000
				HD_ITEM hdi = { HDI_FORMAT | HDI_BITMAP };
#else
				HD_ITEM hdi = { HDI_FORMAT };
#endif
				Header_GetItem(hHeader, i, &hdi);
				if (hdi.fmt & (HDF_SORTUP | HDF_SORTDOWN | HDF_BITMAP)) {
					if (hdi.fmt & HDF_SORTDOWN) {
						ToMinus(pbs);
					}
#ifdef _W2000
					if (hdi.fmt & HDF_BITMAP) {	//Windows 2000
						if (!(hdi.fmt & (HDF_SORTUP | HDF_SORTDOWN))) {
							HDC hdc = GetDC(hHeader);
							HDC hCompatDC = CreateCompatibleDC(hdc);
							ReleaseDC(hHeader, hdc);
							HGDIOBJ hDefault = SelectObject(hCompatDC, hdi.hbm);
							if (GetPixel(hCompatDC, 3, 4) != GetSysColor(COLOR_BTNFACE)) {
								ToMinus(pbs);
							}
							SelectObject(hCompatDC, hDefault);
							DeleteDC(hCompatDC);
							ReleaseDC(hHeader, hdc);
						}
						DeleteObject(hdi.hbm);
					}
#endif
					break;
				}
			}
		}
	}
#endif
}

VOID CteShellBrowser::GetGroupBy(BSTR* pbs)
{
	*pbs = NULL;
	if (!m_pShellView) {
		return;
	}
	IFolderView2 *pFV2;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		PROPERTYKEY propKey;
		BOOL fAscending;
		if SUCCEEDED(pFV2->GetGroupBy(&propKey, &fAscending)) {
			*pbs = tePSGetNameFromPropertyKeyEx(propKey, 1, this);
			if (!fAscending) {
				ToMinus(pbs);
			}
		}
		pFV2->Release();
		return;
	}
#ifdef _2000XP
	if (ListView_IsGroupViewEnabled(m_hwndLV)) {
		IContextMenu *pCM;
		if SUCCEEDED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&pCM))) {
			HMENU hMenu = CreatePopupMenu();
			pCM->QueryContextMenu(hMenu, 0, 1, 0x7fff, CMF_DEFAULTONLY);
			WCHAR szName[MAX_COLUMN_NAME_LEN];
			MENUITEMINFO mii = { sizeof(MENUITEMINFO), MIIM_SUBMENU };
			GetMenuItemInfo(hMenu, 2, TRUE, &mii);
			int wID = mii.wID;
			HMENU hSubMenu = mii.hSubMenu;
			int nCount = GetMenuItemCount(hSubMenu);
			mii.fMask = MIIM_STATE | MIIM_STRING;
			mii.dwTypeData = szName;
			mii.cch = MAX_COLUMN_NAME_LEN;
			for (int i = 0; i < nCount; ++i) {
				GetMenuItemInfo(hSubMenu, i, TRUE, &mii);
				if (mii.fState & MFS_CHECKED) {
					teStripAmp(szName);
					*pbs = ::SysAllocString(szName);
					break;
				}
			}
			DestroyMenu(hMenu);
			pCM->Release();
		}
		return;
	}
	*pbs = ::SysAllocString(L"System.Null");
#endif
}

FOLDERVIEWOPTIONS CteShellBrowser::teGetFolderViewOptions(LPITEMIDLIST pidl, UINT uViewMode)
{
	if (m_param[SB_FolderFlags] & FWF_CHECKSELECT) {
		return FVO_VISTALAYOUT;
	}
	if (g_bCanLayout) {
		UINT uFlag = 1;
		while (--uViewMode > 0) {
			uFlag *= 2;
		}
		if (uFlag & g_param[TE_Layout]) {
			return FVO_DEFAULT;
		}
	}
	if (g_bsExplorerBrowserFilter) {
		BSTR bsPath;
		if SUCCEEDED(teGetDisplayNameFromIDList(&bsPath, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
			FOLDERVIEWOPTIONS fvo = tePathMatchSpec(bsPath, g_bsExplorerBrowserFilter) ? FVO_DEFAULT : FVO_VISTALAYOUT;
			teSysFreeString(&bsPath);
			return fvo;
		}
	}
	return FVO_VISTALAYOUT;
}

VOID CteShellBrowser::SetSort(BSTR bs)
{
	if (!m_pShellView || lstrlen(bs) == 0) {
		return;
	}
	IFolderView2 *pFV2;
	SORTCOLUMN pCol1[1];
	SORTCOLUMN *pCol = pCol1;
	if (lstrcmpi(bs, L"System.Null") == 0) {
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
			pFV2->SetSortColumns(NULL, 0);
			pFV2->Release();
		}
		return;
	}
	teSysFreeString(&m_bsAltSortColumn);
	if (g_pOnFunc[TE_OnSorting]) {
		VARIANT v;
		VariantInit(&v);
		VARIANTARG *pv = GetNewVARIANT(2);
		teSetObject(&pv[1], this);
		teSetSZ(&pv[0], bs);
		Invoke4(g_pOnFunc[TE_OnSorting], &v, 2, pv);
		if (GetIntFromVariantClear(&v)) {
			m_bsAltSortColumn = ::SysAllocString(bs);
			return;
		}
	}
	int nSortColumns = _countof(pCol1);
	int dir = (bs[0] == '-') ? 1 : 0;
	LPWSTR szNew = &bs[dir];
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		if SUCCEEDED(teGetPropertyKeyFromName(m_pSF2, szNew, &pCol1[0].propkey)) {
			pCol1[0].direction = dir ? SORT_DESCENDING : SORT_ASCENDING;
			if (IsEqualPropertyKey(pCol1[0].propkey, PKEY_Null)) {
				pCol = g_pSortColumnNull;
				nSortColumns = _countof(g_pSortColumnNull);
			}
			if (!teIsSameSort(pFV2, pCol, nSortColumns)) {
				m_nGroupByDelay += 10;
				m_nFocusItem = -1;
				pFV2->SetSortColumns(pCol, nSortColumns);
			}
		}
		pFV2->Release();
		return;
	}
#ifdef _2000XP
	BOOL bGroup = ListView_IsGroupViewEnabled(m_hwndLV);
	if (bGroup) {
		SendMessage(m_hwndDV, WM_COMMAND, CommandID_GROUP, 0);
	}
	BSTR bsName = NULL;
	PROPERTYKEY propKey;
	if SUCCEEDED(_PSPropertyKeyFromStringEx(szNew, &propKey)) {
		bsName = tePSGetNameFromPropertyKeyEx(propKey, 0, this);
		szNew = bsName;
	}
	BSTR bs1;
	GetSort(&bs1, 0);
	int dir1 = 0;
	LPWSTR szOld = NULL;
	if (bs1) {
		if (bs1[0] == '-') {
			dir1 = 1;
		}
		szOld = &bs1[dir1];
	}
	if (dir != dir1 || lstrcmpi(szNew, szOld)) {
		IShellFolderView *pSFV;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
			int i = PSGetColumnIndexXP(szNew, NULL);
			if (i >= 0) {
				pSFV->Rearrange(i);
				if (dir && lstrcmpi(szNew, szOld)) {
					pSFV->Rearrange(i);
				}
			}
			pSFV->Release();
		}
		::SysFreeString(bs1);
		teSysFreeString(&bsName);
	}
	if (bGroup) {
		SendMessage(m_hwndDV, WM_COMMAND, CommandID_GROUP, 0);
	}
#endif
}

VOID CteShellBrowser::SetGroupBy(BSTR bs)
{
	if (!m_pShellView || lstrlen(bs) == 0) {
		return;
	}
	BSTR bs0;
	GetGroupBy(&bs0);
	if (teStrSameIFree(bs0, bs)) {
		return;
	}
	IFolderView2 *pFV2;
	PROPERTYKEY propKey;
	int dir = (bs[0] == '-') ? 1 : 0;
	LPWSTR szNew = &bs[dir];
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		if SUCCEEDED(teGetPropertyKeyFromName(m_pSF2, szNew, &propKey)) {
			teSysFreeString(&m_bsNextGroup);
			if (m_nGroupByDelay > 0) {
				m_bsNextGroup = ::SysAllocString(bs);
				SetTimer(m_hwnd, (UINT_PTR)&m_bsNextGroup, 500, teTimerProcGroupBy);
			} else {
				pFV2->SetGroupBy(propKey, !dir);
				if (m_hwndLV) {
					if (IsEqualPropertyKey(propKey, PKEY_Null)) {
						ListView_EnableGroupView(m_hwndLV, FALSE);
					}
				}
			}
		}
		pFV2->Release();
		return;
	}
#ifdef _2000XP
	if (szNew[0] && lstrcmpi(szNew, L"System.Null")) {
		if (!ListView_IsGroupViewEnabled(m_hwndLV)) {
			SendMessage(m_hwndDV, WM_COMMAND, CommandID_GROUP, 0);
		}
		IContextMenu *pCM;
		if SUCCEEDED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&pCM))) {
			HMENU hMenu = CreatePopupMenu();
			pCM->QueryContextMenu(hMenu, 0, 1, 0x7fff, CMF_DEFAULTONLY);
			MENUITEMINFO mii = { sizeof(MENUITEMINFO), MIIM_SUBMENU | MIIM_ID };
			GetMenuItemInfo(hMenu, 2, TRUE, &mii);
			int wID = mii.wID;
			HMENU hSubMenu = mii.hSubMenu;
			int nCount = GetMenuItemCount(hSubMenu);
			for (int i = 0; i < nCount; ++i) {
				WCHAR szName[MAX_COLUMN_NAME_LEN];
				mii.fMask = MIIM_ID | MIIM_STRING;
				mii.cch = MAX_COLUMN_NAME_LEN;
				mii.dwTypeData = szName;
				GetMenuItemInfo(hSubMenu, i, TRUE, &mii);
				teStripAmp(szName);
				if (lstrcmpi(szName, szNew) == 0) {
					CMINVOKECOMMANDINFO cmi = { sizeof(CMINVOKECOMMANDINFO), 0, NULL, (LPCSTR)(mii.wID - 1) };
					pCM->InvokeCommand(&cmi);
					break;
				}
			}
			DestroyMenu(hMenu);
			pCM->Release();
		}
	} else if (ListView_IsGroupViewEnabled(m_hwndLV)) {
		SendMessage(m_hwndDV, WM_COMMAND, CommandID_GROUP, 0);
	}
#endif
}

VOID CteShellBrowser::SetRedraw(BOOL bRedraw)
{
	SendMessage(m_hwnd, WM_SETREDRAW, m_bSetRedraw && (bRedraw || !m_bRedraw), 0);
	if (m_hwndAlt) {
		BringWindowToTop(m_hwndAlt);
	}
}

//CTE

CTE::CTE(HWND hwnd)
{
	m_cRef = 1;
	m_hwnd = hwnd;
}

CTE::~CTE()
{
}

STDMETHODIMP CTE::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CTE, IDispatch),
		QITABENT(CTE, IDropSource),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CTE::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CTE::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CTE::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CTE::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTE::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodTE, _countof(methodTE), g_maps[MAP_TE], *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CTE::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		IUnknown *punk = NULL;
		IDispatch *pdisp = NULL;

		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		switch(dispIdMember) {
			//Data
			case 1001:
				if (nArg >= 0) {
					VariantClear(&g_vData);
					VariantCopy(&g_vData, &pDispParams->rgvarg[nArg]);
				}
				teVariantCopy(pVarResult, &g_vData);
				return S_OK;
			//hwnd
			case 1002:
				teSetPtr(pVarResult, g_hwndMain);
				return S_OK;
			//About
			case 1004:
				if (nArg >= 0 && pDispParams->rgvarg[nArg].vt == VT_BSTR) {
					StrNCpy(g_szTE, pDispParams->rgvarg[nArg].bstrVal, MAX_LOADSTRING);
				}
				teSetSZ(pVarResult, g_szTE);
				return S_OK;
			//Ctrl
			case TE_METHOD + 1005:
				if (pVarResult) {
					if (nArg >= 0) {
						int nCtrl = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
						switch (nCtrl) {
							case CTRL_FV:
							case CTRL_SB:
							case CTRL_EB:
							case CTRL_TV:
								CteShellBrowser *pSB;
								pSB = NULL;
								if (nArg >= 1) {
									UINT uId = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) - 1;
									if (uId < g_ppSB.size()) {
										pSB = g_ppSB[uId];
									}
								}
								if (!pSB && g_pTC) {
									pSB = g_pTC->GetShellBrowser(g_pTC->m_nIndex);
									if (!pSB) {
										pSB = GetNewShellBrowser(g_pTC);
									}
								}
								if (pSB) {
									if (nCtrl == CTRL_TV) {
										teSetObject(pVarResult, pSB->m_pTV);
									} else {
										teSetObject(pVarResult, pSB);
									}
								}
								break;
							case CTRL_TC:
								CteTabCtrl *pTC;
								pTC = g_pTC;
								if (nArg >= 1) {
									int nId = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
									if (nId) {
										pTC = g_ppTC[nId - 1];
									}
								}
								if (pTC) {
									teSetObject(pVarResult, pTC);
								}
								break;
							case CTRL_WB:
								size_t u;
								if (teFindSubWindow(m_hwnd, &u)) {
									teSetObject(pVarResult, g_pSubWindows[u].pWB);
								} else {
									teSetObject(pVarResult, g_pWebBrowser);
								}
								break;
						}//end_switch
					}
				}
				return S_OK;
			//Ctrls
			case TE_METHOD + 1006:
				if (nArg >= 0) {
					int nCtrl = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					BOOL bAll = nArg >= 1 ? !GetBoolFromVariant(&pDispParams->rgvarg[nArg -  1]) : TRUE;
					IDispatch *pArray = teCreateObj(TE_ARRAY, NULL);
					switch (nCtrl) {
						case CTRL_FV:
						case CTRL_SB:
						case CTRL_EB:
						case CTRL_TV:
							for (size_t i = 0; i < g_ppSB.size(); ++i) {
								CteShellBrowser *pSB = g_ppSB[i];
								if (!pSB->m_bEmpty) {
									if (bAll || pSB->m_pTC->m_bVisible) {
										if (nCtrl == CTRL_TV) {
											if (pSB->m_pTV->m_param[SB_TreeAlign] & 2) {
												teArrayPush(pArray, pSB->m_pTV);
											}
										} else {
											teArrayPush(pArray, pSB);
										}
									}
								}
							}
							if (nCtrl == CTRL_TV) {
								for (size_t i = 0; i < g_ppTV.size(); ++i) {
									CteTreeView *pTV = g_ppTV[i];
									if (!pTV->m_bEmpty && (bAll || pTV->m_param[SB_TreeAlign] & 2)) {
										teArrayPush(pArray, pTV);
									}
								}
							}
							break;
						case CTRL_TC:
							for (size_t i = 0; i < g_ppTC.size(); ++i) {
								CteTabCtrl *pTC = g_ppTC[i];
								if (!pTC->m_bEmpty && (bAll || pTC->m_bVisible)) {
									teArrayPush(pArray, pTC);
								}
							}
							break;
						case CTRL_WB:
							if (g_pWebBrowser) {
								teArrayPush(pArray, g_pWebBrowser);
								for (size_t u = 0; u <  g_pSubWindows.size(); ++u) {
									teArrayPush(pArray, g_pSubWindows[u].pWB);
								}
							}
							break;
					}//end_switch
					if (nArg >= 2 && GetBoolFromVariant(&pDispParams->rgvarg[nArg - 2])) {
						teCreateSafeArrayFromVariantArray(pArray, pVarResult);
						SafeRelease(&pArray);
						return S_OK;
					}
					teSetObjectRelease(pVarResult, teAddLegacySupport(pArray));
				}
				return S_OK;

			//ClearEvents
			case TE_METHOD + 1008:
				ClearEvents();
				teRegister();
				teGetDarkMode();
				g_nReload = 0;
				return S_OK;
			//Reload
			case TE_METHOD + 1009:
				g_nReload = 1;
				if (g_nException-- <= 0 || (nArg >= 0 && GetBoolFromVariant(&pDispParams->rgvarg[nArg]))) {
#ifdef _DEBUG
					if (g_strException) {
						MessageBox(NULL, g_strException, NULL, MB_OK);
					}
#endif
					SetTimer(g_hwndMain, TET_Reload, 100, teTimerProc);
					return S_OK;
				}
				ClearEvents();
				SetTimer(g_hwndMain, TET_Refresh, 100, teTimerProc);
				return S_OK;
			//DIID_DShellWindowsEvents
			case DISPID_WINDOWREGISTERED:
				DoFunc(TE_OnWindowRegistered, g_pTE, S_OK);
				return S_OK;
			//GetObject
			case TE_METHOD + 1020:
			//CreateObject
			case TE_METHOD + 1010:
				if (nArg >= 0) {
					CLSID clsid;
					VARIANT vClass;
					teVariantChangeType(&vClass, &pDispParams->rgvarg[nArg], VT_BSTR);
					HRESULT hr = E_FAIL;
					IUnknown *punk;
					if SUCCEEDED(teCLSIDFromProgID(vClass.bstrVal, &clsid)) {
						if (dispIdMember == TE_METHOD + 1020) {
							hr = GetActiveObject(clsid, NULL, &punk);
						}
						if FAILED(hr) {
							hr = teCreateInstance(clsid, NULL, NULL, IID_PPV_ARGS(&punk));
						}
					} else if (dispIdMember == TE_METHOD + 1020) {
						hr = CoGetObject(vClass.bstrVal, NULL, IID_PPV_ARGS(&punk));
					}
					if SUCCEEDED(hr) {
						teSetObjectRelease(pVarResult, punk);
					}
					VariantInit(&vClass);
				}
				return S_OK;
			//AddEvent
			case TE_METHOD + 1025:
			//RemoveEvent
			case TE_METHOD + 1026:
				if (nArg >= 1) {
					LPWSTR lpMethod = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]);
					LONG_PTR lpProc = GetPtrFromVariant(&pDispParams->rgvarg[nArg - 1]);
					if (lpProc) {
						if (lstrcmpi(lpMethod, L"GetArchive") == 0) {
							teAddRemoveProc(&g_ppGetArchive, lpProc, dispIdMember == TE_METHOD + 1025);
						} else if (lstrcmpi(lpMethod, L"GetImage") == 0) {
							teAddRemoveProc(&g_ppGetImage, lpProc, dispIdMember == TE_METHOD + 1025);
						} else if (lstrcmpi(lpMethod, L"MessageSFVCB") == 0) {
							teAddRemoveProc(&g_ppMessageSFVCB, lpProc, dispIdMember == TE_METHOD + 1025);
						}
					}
				}
				return S_OK;

			case 1030://WindowsAPI
#ifdef _USE_OBJECTAPI
				teSetObject(pVarResult, g_pAPI);
#else
				teSetObjectRelease(pVarResult, new CteWindowsAPI(NULL));
#endif
				return S_OK;

			case 1031://WindowsAPI0
				IDispatch *pAPI;
				pAPI = NULL;
				GetNewObject(&pAPI);
				VARIANT v;
				teSetObjectRelease(&v, new CteWindowsAPI(&dispAPI[0]));
				BSTR bs;
				bs = teMultiByteToWideChar(CP_UTF8, dispAPI[0].name, -1);
				tePutProperty(pAPI, bs, &v);
				::SysFreeString(bs);
				teSetObject(pVarResult, pAPI);
				return S_OK;

			case 1131://CommonDialog
				teSetObjectRelease(pVarResult, new CteCommonDialog());
				return S_OK;

			case 1132://WICBitmap(GdiplusBitmap)
				teSetObjectRelease(pVarResult, new CteWICBitmap());
				return S_OK;

			case TE_METHOD + 1133://FolderItems
				IDataObject *pDataObj;
				pDataObj = NULL;
				if (nArg >= 0) {
					GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg]);
				}
				teSetObjectRelease(pVarResult, new CteFolderItems(pDataObj, NULL));
				SafeRelease(&pDataObj);
				return S_OK;

			case TE_METHOD + 1134://Object
				teSetObjectRelease(pVarResult, teCreateObj(TE_OBJECT, NULL));
				return S_OK;

			case TE_METHOD + 1135://Array
				teSetObjectRelease(pVarResult, teCreateObj(TE_ARRAY, NULL));
				return S_OK;

			case 1136://Arguments
				if (nArg >= 0) {
					VariantClear(&g_vArguments);
					VariantCopy(&g_vArguments, &pDispParams->rgvarg[nArg]);
				} else {
					teVariantCopy(pVarResult, &g_vArguments);
					VariantClear(&g_vArguments);
				}
				return S_OK;

			case 1137://ProgressDialog
				teSetObjectRelease(pVarResult, new CteProgressDialog(NULL));
				return S_OK;

			case 1138://DateTimeFormat
				if (nArg >= 0) {
					teSysFreeString(&g_bsDateTimeFormat);
					if (pDispParams->rgvarg[nArg].vt == VT_BSTR && ::SysStringLen(pDispParams->rgvarg[nArg].bstrVal)) {
						g_bsDateTimeFormat = ::SysAllocString(pDispParams->rgvarg[nArg].bstrVal);
					}
				}
				teSetSZ(pVarResult, g_bsDateTimeFormat);
				return S_OK;

			case 1139://HiddenFilter
				if (nArg >= 0) {
					teSysFreeString(&g_bsHiddenFilter);
					if (pDispParams->rgvarg[nArg].vt == VT_BSTR && ::SysStringLen(pDispParams->rgvarg[nArg].bstrVal)) {
						g_bsHiddenFilter = ::SysAllocString(pDispParams->rgvarg[nArg].bstrVal);
					}
				}
				teSetSZ(pVarResult, g_bsHiddenFilter);
				return S_OK;
			//ThumbnailProvider//Deprecated
/*			case 1150:
				teSetPtr(pVarResult, teThumbnailProvider);
				return S_OK;*/

			case 1160://DragIcon
				if (nArg >= 0) {
					g_bDragIcon = GetBoolFromVariant(&pDispParams->rgvarg[nArg]);
				}
				teSetBool(pVarResult, g_bDragIcon);
				return S_OK;

			case 1180://ExplorerBrowserFilter
				if (nArg >= 0) {
					teSysFreeString(&g_bsExplorerBrowserFilter);
					if (pDispParams->rgvarg[nArg].vt == VT_BSTR && ::SysStringLen(pDispParams->rgvarg[nArg].bstrVal)) {
						g_bsExplorerBrowserFilter = ::SysAllocString(pDispParams->rgvarg[nArg].bstrVal);
					}
				}
				teSetSZ(pVarResult, g_bsExplorerBrowserFilter);
				return S_OK;

#ifdef _USE_TESTOBJECT
			//TestObj
			case 1200:
				teSetObjectRelease(pVarResult, new CteTest());
				return S_OK;
#endif
			//CtrlFromPoint
			case TE_METHOD + 1040:
				if (nArg >= 0) {
					POINT pt;
					GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg]);
					IDispatch *pdisp;
					if SUCCEEDED(ControlFromhwnd(&pdisp, WindowFromPoint(pt))) {
						teSetObjectRelease(pVarResult, pdisp);
					}
				}
				return S_OK;
			//CreateCtrl
			case TE_METHOD + 1050:
				if (nArg >= 0) {
					switch (GetIntFromVariant(&pDispParams->rgvarg[nArg])) {
						case CTRL_TC:
							if (nArg >= 4) {
								CteTabCtrl *pTC;
								pTC = GetNewTC();
								pTC->Init();
								int i = _countof(pTC->m_param) - 1;
								for (i = nArg < i ? nArg : i; i >= 0; --i) {
									pTC->m_param[i] = GetIntFromVariantPP(&pDispParams->rgvarg[nArg - i], i);
								}
								if (pTC->Create()) {
									teSetObject(pVarResult, pTC);
									pTC->Show(TRUE);
								} else {
									pTC->Release();
								}
							}
							break;
						case CTRL_SW:
							if (nArg >= 5) {
								VARIANT v;
								teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
								HWND hwnd1 = (HWND)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 3]);
								if (FindUnknown(&pDispParams->rgvarg[nArg - 3], &punk)) {
									IUnknown_GetWindow(punk, &hwnd1);
								}
								if (!hwnd1) {
									hwnd1 = g_hwndMain;
								}
								HWND hwndParent = hwnd1;
								while (hwnd1 = GetParent(hwndParent)) {
									hwndParent = hwnd1;
								}
								RECT rc, rcWindow;
								GetWindowRect(hwndParent, &rcWindow);
								int w = GetIntFromVariant(&pDispParams->rgvarg[nArg - 4]);
								int h = GetIntFromVariant(&pDispParams->rgvarg[nArg - 5]);
								int x = nArg >= 6 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 6]) : 0;
								int y = nArg >= 7 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 7]) : 0;
								MyRegisterClass(hInst, WINDOW_CLASS2, WndProc2);
								HWND hwnd = CreateWindowEx(WS_EX_LAYERED, WINDOW_CLASS2, g_szTE, WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
									x, y, w, h, hwndParent, NULL, hInst, NULL);
								SetLayeredWindowAttributes(hwnd, 0, 0, LWA_ALPHA);
								teSetDarkMode(hwnd);
								SetClassLongPtr(hwnd, GCLP_HBRBACKGROUND, GetClassLongPtr(g_hwndMain, GCLP_HBRBACKGROUND));
								GetClientRect(hwnd, &rc);
								int a = w - (rc.right - rc.left);
								int b = h - (rc.bottom - rc.top);
								w += a;
								h += b;
								if (!x) {
									x = rcWindow.left + (rcWindow.right - rcWindow.left - w) / 2;
									if (w == rcWindow.right - rcWindow.left) {
										x += a;
									}
								}
								if (!y) {
									y = rcWindow.top + (rcWindow.bottom - rcWindow.top - h) / 2;
									if (h == rcWindow.bottom - rcWindow.top) {
										y += b / 2;
									}
								}
								rc.left = x;
								rc.top = y;
								rc.right = x + w;
								rc.bottom = y + h;
								HMONITOR hMonitor = MonitorFromRect(&rc, MONITOR_DEFAULTTONEAREST);
								MONITORINFO mi;
								mi.cbSize = sizeof(mi);
								if (GetMonitorInfo(hMonitor, &mi)) {
									if (x + w > mi.rcWork.right) {
										x = mi.rcWork.right - w;
									}
									if (x < mi.rcWork.left) {
										x = mi.rcWork.left;
									}
									if (y + h > mi.rcWork.bottom) {
										y = mi.rcWork.bottom - h;
									}
									if (y < mi.rcWork.top) {
										y = mi.rcWork.top;
									}
								}
								MoveWindow(hwnd, x, y, w, h, FALSE);
								ShowWindow(hwnd, SW_SHOWNORMAL);
								TEWebBrowsers TEWB;
								TEWB.hwnd = hwnd;
								TEWB.pWB = new CteWebBrowser(hwnd, v.bstrVal, &pDispParams->rgvarg[nArg - 2]);
								size_t u;
								if (teFindSubWindow(hwnd, &u)) {
									g_pSubWindows[u].pWB->Release();
									g_pSubWindows[u].pWB = TEWB.pWB;
								} else {
									g_pSubWindows.push_back(TEWB);
								}
								teSetObjectRelease(pVarResult, TEWB.pWB);
							}
							break;
						case CTRL_TV:
							CteTreeView *pTV;
							pTV = GetNewTV();
							pTV->Init(NULL, g_hwndMain);
							teSetObject(pVarResult, pTV);
							break;
					}//end_switch
				}
				return S_OK;

			//MainMenu
			case TE_METHOD + 1060:
				if (pVarResult) {
					if (nArg >= 0) {
						if (!g_hMenu) {
							g_hMenu = LoadMenu(g_hShell32, MAKEINTRESOURCE(216));
/*//
							if (!g_hMenu) {
								CteShellBrowser *pSB = GetNewShellBrowser(NULL);
								IShellFolder *pDesktopFolder;
								SHGetDesktopFolder(&pDesktopFolder);
								if (pDesktopFolder->CreateViewObject(g_hwndMain, IID_PPV_ARGS(&pSB->m_pShellView)) == S_OK) {
									FOLDERSETTINGS fs = { FWF_NOWEBVIEW, FVM_LIST };
									RECT rc;
									SetRectEmpty(&rc);
									if SUCCEEDED(pSB->m_pShellView->CreateViewWindow(NULL, &fs, pSB, &rc, &pSB->m_hwnd)) {
										pSB->m_pShellView->UIActivate(SVUIA_ACTIVATE_NOFOCUS);
									}
								}
								pDesktopFolder->Release();
								pSB->Close(TRUE);
							}
*///
						}
						MENUITEMINFO mii = { sizeof(MENUITEMINFO), MIIM_SUBMENU };
						GetMenuItemInfo(g_hMenu, GetIntFromVariant(&pDispParams->rgvarg[nArg]), FALSE, &mii);

						HMENU hMenu = CreatePopupMenu();
						UINT fState = MFS_DISABLED;
						if (g_pTC) {
							CteShellBrowser *pSB = g_pTC->GetShellBrowser(g_pTC->m_nIndex);
							if (pSB) {
								fState = MFS_ENABLED;
							}
						}
						teCopyMenu(hMenu, mii.hSubMenu, fState);
						teSetPtr(pVarResult, hMenu);
					}
				}
				return S_OK;
			//CtrlFromWindow
			case TE_METHOD + 1070:
				if (nArg >= 0) {
					IDispatch *pdisp;
					if SUCCEEDED(ControlFromhwnd(&pdisp, (HWND)GetPtrFromVariant(&pDispParams->rgvarg[nArg]))) {
						teSetObjectRelease(pVarResult, pdisp);
					}
				}
				return S_OK;
			//LockUpdate
			case TE_METHOD + 1080:
				teLockUpdate(nArg < 0 ? 16 :1);
				return S_OK;
			//UnlockUpdate
			case TE_METHOD + 1090:
				teUnlockUpdate(nArg < 0 ? 16 :1);
				if (nArg < 0 && g_nLockUpdate < 16) {
					ArrangeWindowEx();
				}
				return S_OK;

			case TE_METHOD + 1091://ArrangeCB
				if (nArg >= 1) {
					IUnknown *punk;
					if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
						CteTabCtrl *pTC;
						if SUCCEEDED(punk->QueryInterface(g_ClsIdTC, (LPVOID *)&pTC)) {
							ArrangeWindowTC(pTC, (LPRECT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1], NULL));
							pTC->Release();
						}
					}
				}
				return S_OK;

			case DISPID_VALUE:
				teSetObject(pVarResult, this);

			case DISPID_TE_UNDEFIND:
				return S_OK;

			case TE_METHOD + 1100://HookDragDrop//Deprecated
				return S_OK;
			default:
				if (dispIdMember >= START_OnFunc && dispIdMember < START_OnFunc + Count_OnFunc) {
					teInvokeObject(&g_pOnFunc[dispIdMember - START_OnFunc], wFlags, pVarResult, nArg, pDispParams->rgvarg);
					if (nArg >= 0 && dispIdMember == TE_Labels + START_OnFunc && g_pOnFunc[TE_Labels]) {
						g_bLabelsMode = !teHasProperty(g_pOnFunc[TE_Labels], L"call");
					}
					if (dispIdMember == START_OnFunc + TE_OnWindowRegistered && !g_dwCookieSW) {
						teAdvise(g_pSW, DIID_DShellWindowsEvents, static_cast<IDispatch *>(g_pTE), &g_dwCookieSW);
					}
					return S_OK;
				}
				//Type, offsetLeft etc.
				else if (dispIdMember >= TE_OFFSET && dispIdMember < TE_OFFSET + Count_TE_params) {
					if (nArg >= 0 && dispIdMember != TE_OFFSET) {
						int i = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
						if (i != g_param[dispIdMember - TE_OFFSET]) {
							g_param[dispIdMember - TE_OFFSET] = i;
							if (dispIdMember >= TE_OFFSET + TE_Left && dispIdMember <= TE_Bottom + TE_OFFSET) {
								ArrangeWindow();
							}
							if (dispIdMember == TE_OFFSET + TE_Layout) {
								for (size_t i = 0; i < g_ppSB.size(); ++i) {
									CteShellBrowser *pSB = g_ppSB[i];
									if (!pSB->m_bEmpty) {
										pSB->m_bCheckLayout = TRUE;
										pSB->GetViewModeAndIconSize(TRUE);
									}
								}
							}
						}
					}
					teSetLong(pVarResult, g_param[dispIdMember - TE_OFFSET]);
				}
				return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, "external", methodTE, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDropSource
STDMETHODIMP CTE::QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState)
{
	if (fEscapePressed || !g_nDropState || (grfKeyState & (MK_LBUTTON | MK_RBUTTON)) == (MK_LBUTTON | MK_RBUTTON)) {
		g_nDropState = 0;
		return DRAGDROP_S_CANCEL;
	}
	if (grfKeyState & (MK_LBUTTON | MK_RBUTTON)) {
		return S_OK;
	}
	if (GetKeyState(VK_LBUTTON) < 0 || GetKeyState(VK_RBUTTON) < 0) {//21.5.11
		return S_OK;
	}
	VARIANT v;
	VariantInit(&v);
	if (g_nDropState == 2 && g_pOnFunc[TE_OnBeforeGetData]) {
		g_nDropState = 0;
		if (g_pDropTargetHelper) {
			g_pDropTargetHelper->DragLeave();
		}
		VARIANTARG *pv = GetNewVARIANT(3);
		teSetObject(&pv[2], g_pDraggingCtrl);
		teSetObjectRelease(&pv[1], new CteFolderItems(g_pDraggingItems, NULL));
		teSetLong(&pv[0], 4);
		Invoke4(g_pOnFunc[TE_OnBeforeGetData], &v, 3, pv);
	}
	g_nDropState = 0;
	g_bDragging = FALSE;
	return GetIntFromVariantClear(&v) ? DRAGDROP_S_CANCEL : DRAGDROP_S_DROP;
}

STDMETHODIMP CTE::GiveFeedback(DWORD dwEffect)
{
	if (g_pDraggingItems) {
		WPARAM wpEffect[] = { 1, 3, 2, 3, 4, 4, 4, 4 };
		if (teGetDWordFromDataObj(g_pDraggingItems, &IsShowingLayeredFormat, FALSE)) {
			HWND hwnd = (HWND)ULongToHandle(teGetDWordFromDataObj(g_pDraggingItems, &DRAGWINDOWFormat, NULL));
			if (hwnd) {
				SendMessage(hwnd, WM_USER + 2, wpEffect[dwEffect & 7], 0);
				SetCursor(LoadCursor(NULL, IDC_ARROW));
				return S_OK;
			}
		}
	}
	return DRAGDROP_S_USEDEFAULTCURSORS;
}

// CteInternetSecurityManager
CteInternetSecurityManager::CteInternetSecurityManager()
{
	m_cRef = 1;
}

CteInternetSecurityManager::~CteInternetSecurityManager()
{
}

STDMETHODIMP CteInternetSecurityManager::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteInternetSecurityManager, IInternetSecurityManager),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteInternetSecurityManager::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteInternetSecurityManager::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

//IInternetSecurityManager
STDMETHODIMP CteInternetSecurityManager::SetSecuritySite(IInternetSecurityMgrSite *pSite)
{
	return INET_E_DEFAULT_ACTION;
}

STDMETHODIMP CteInternetSecurityManager::GetSecuritySite(IInternetSecurityMgrSite **ppSite)
{
	return INET_E_DEFAULT_ACTION;
}

STDMETHODIMP CteInternetSecurityManager::MapUrlToZone(LPCWSTR pwszUrl, DWORD *pdwZone, DWORD dwFlags)
{
	return INET_E_DEFAULT_ACTION;
//	*pdwZone = URLZONE_LOCAL_MACHINE;
//	return S_OK;
}

STDMETHODIMP CteInternetSecurityManager::GetSecurityId(LPCWSTR pwszUrl, BYTE *pbSecurityId, DWORD *pcbSecurityId, DWORD_PTR dwReserved)
{
	return INET_E_DEFAULT_ACTION;
}

STDMETHODIMP CteInternetSecurityManager::ProcessUrlAction(LPCWSTR pwszUrl, DWORD dwAction, BYTE *pPolicy, DWORD cbPolicy, BYTE *pContext, DWORD cbContext, DWORD dwFlags, DWORD dwReserved)
{
	*pPolicy = teStartsText(L"http:", pwszUrl) ? URLPOLICY_DISALLOW : URLPOLICY_ALLOW;
	return S_OK;
}

STDMETHODIMP CteInternetSecurityManager::QueryCustomPolicy(LPCWSTR pwszUrl, REFGUID guidKey, BYTE **ppPolicy, DWORD *pcbPolicy, BYTE *pContext, DWORD cbContext, DWORD dwReserved)
{
	return INET_E_DEFAULT_ACTION;
}

STDMETHODIMP CteInternetSecurityManager::SetZoneMapping(DWORD dwZone, LPCWSTR lpszPattern, DWORD dwFlags)
{
	return INET_E_DEFAULT_ACTION;
}

STDMETHODIMP CteInternetSecurityManager::GetZoneMappings(DWORD dwZone, IEnumString **ppenumString, DWORD dwFlags)
{
	return INET_E_DEFAULT_ACTION;
}

// CteNewWindowManager
CteNewWindowManager::CteNewWindowManager()
{
	m_cRef = 1;
}

CteNewWindowManager::~CteNewWindowManager()
{
}

STDMETHODIMP CteNewWindowManager::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteNewWindowManager, INewWindowManager),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteNewWindowManager::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteNewWindowManager::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

//INewWindowManager
STDMETHODIMP CteNewWindowManager::EvaluateNewWindow(LPCWSTR pszUrl, LPCWSTR pszName, LPCWSTR pszUrlContext, LPCWSTR pszFeatures, BOOL fReplace, DWORD dwFlags, DWORD dwUserActionTime)
{
	return teStartsText(L"http", pszUrl) ? S_FALSE : S_OK;
}

// CteWebBrowser

CteWebBrowser::CteWebBrowser(HWND hwndParent, WCHAR *szPath, VARIANT *pvArg)
{
	m_cRef = 1;
	m_hwndParent = hwndParent;
	m_dwCookie = 0;
	m_pDragItems = NULL;
	m_pDropTarget = NULL;
	m_nClose = 0;
	VariantInit(&m_vData);
	if (pvArg) {
		VariantClear(&g_vArguments);
		VariantCopy(&g_vArguments, pvArg);
	}
	for (int i = Count_WBFunc; i--;) {
		m_ppDispatch[i] = NULL;
	}
	m_ppDispatch[WB_External] = new CTE(hwndParent);
	MSG        msg;
	RECT       rc;
	IOleObject *pOleObject;
	HRESULT hr = teCreateWebView2(&m_pWebBrowser);
	if FAILED(hr) {
		hr = teCreateInstance(CLSID_WebBrowser, NULL, NULL, IID_PPV_ARGS(&m_pWebBrowser));
	}
	if (SUCCEEDED(hr) && SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleObject)))) {
		pOleObject->SetClientSite(static_cast<IOleClientSite *>(this));
		SetRectEmpty(&rc);
		pOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, &msg, static_cast<IOleClientSite *>(this), 0, hwndParent, &rc);
		pOleObject->Release();
		teAdvise(m_pWebBrowser, DIID_DWebBrowserEvents2, static_cast<IDispatch *>(this), &m_dwCookie);
		m_pWebBrowser->put_Offline(VARIANT_TRUE);
		m_bstrPath = SysAllocString(szPath);
		m_pWebBrowser->Navigate(m_bstrPath, NULL, NULL, NULL, NULL);
		m_pWebBrowser->put_Visible(VARIANT_TRUE);
	}
}

CteWebBrowser::~CteWebBrowser()
{
	teSysFreeString(&m_bstrPath);
	Close();
	VariantClear(&m_vData);
	for (int i = Count_WBFunc; i--;) {
		SafeRelease(&m_ppDispatch[i]);
	}
}

STDMETHODIMP CteWebBrowser::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteWebBrowser, IOleClientSite),
		QITABENT(CteWebBrowser, IOleWindow),
		QITABENT(CteWebBrowser, IOleInPlaceSite),
		QITABENT(CteWebBrowser, IDispatch),
		QITABENT(CteWebBrowser, IDocHostUIHandler),
		QITABENT(CteWebBrowser, IDropTarget),
		QITABENT(CteWebBrowser, IDocHostShowUI),
//		QITABENT(CteWebBrowser, IOleCommandTarget),
		{ 0 },
	};
	*ppvObject = NULL;
	if (IsEqualIID(riid, IID_IServiceProvider)) {
		*ppvObject = static_cast<IServiceProvider *>(new CteServiceProvider(static_cast<IDispatch *>(this), m_pWebBrowser));
		return S_OK;
	} else if (IsEqualIID(riid, IID_IWebBrowser2)) {
		return m_pWebBrowser->QueryInterface(riid, (LPVOID *)ppvObject);
	} else {
		return QISearch(this, qit, riid, ppvObject);
	}
	AddRef();
	return S_OK;
}

STDMETHODIMP_(ULONG) CteWebBrowser::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteWebBrowser::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteWebBrowser::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteWebBrowser::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodWB, 0, NULL, *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteWebBrowser::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD &&  dispIdMember <= TE_METHOD_MAX) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		switch(dispIdMember) {
		case TE_PROPERTY + 1://Data
			if (nArg >= 0) {
				VariantClear(&m_vData);
				VariantCopy(&m_vData, &pDispParams->rgvarg[nArg]);
			}
			teVariantCopy(pVarResult, &m_vData);
			return S_OK;

		case TE_PROPERTY + 2://hwnd
			teSetPtr(pVarResult, get_HWND());
			return S_OK;

		case TE_PROPERTY + 3://Type
			teSetLong(pVarResult, m_hwndParent == g_hwndMain ? CTRL_WB : CTRL_SW);
			return S_OK;

		case TE_METHOD + 4://TranslateAccelerator
			HRESULT hr;
			hr = E_NOTIMPL;
			if (nArg >= 3) {
				IOleInPlaceActiveObject *pActiveObject = NULL;
				if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pActiveObject))) {
					MSG msg;
					msg.hwnd = (HWND)GetPtrFromVariant(&pDispParams->rgvarg[nArg]);
					msg.message = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
					msg.wParam = (WPARAM)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 2]);
					msg.lParam = (LPARAM)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 3]);
					hr = pActiveObject->TranslateAcceleratorW(&msg);
					pActiveObject->Release();
				}
			}
			teSetLong(pVarResult, hr);
			return S_OK;

		case TE_PROPERTY + 5://Application
			IDispatch *pdisp;
			if SUCCEEDED(m_pWebBrowser->get_Application(&pdisp)) {
				teSetObjectRelease(pVarResult, pdisp);
			}
			return S_OK;

		case TE_PROPERTY + 6://Document
			if (g_nBlink == 1 && nArg >= 0) {
				m_pWebBrowser->PutProperty(L"document", pDispParams->rgvarg[nArg]);
			}
			if SUCCEEDED(m_pWebBrowser->get_Document(&pdisp)) {
				teSetObjectRelease(pVarResult, pdisp);
			}
			return S_OK;

		case TE_PROPERTY + 7://Window
			if (teGetHTMLWindow(m_pWebBrowser, IID_PPV_ARGS(&pdisp))) {
				teSetObjectRelease(pVarResult, pdisp);
			}
			return S_OK;

		case TE_METHOD + 8://Focus
			if (g_nBlink == 1) {
				IOleInPlaceObject *pOleInPlaceObject;
				m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleInPlaceObject));
				pOleInPlaceObject->ReactivateAndUndo();
				pOleInPlaceObject->Release();
			} else {
				teSetForegroundWindow(m_hwndParent);
			}
			return S_OK;

		case TE_METHOD + 9://Close
			m_nClose = 1;
			teSetExStyleOr(m_hwndParent, WS_EX_LAYERED);
			SetLayeredWindowAttributes(m_hwndParent, 0, 0, LWA_ALPHA);
			PostMessage(m_hwndParent, WM_CLOSE, 0, 0);
			return S_OK;

		case TE_METHOD + 10://PreventClose
			m_nClose = 0;
			return S_OK;

		case TE_PROPERTY + 11://Visible
			if (m_pWebBrowser) {
				if (nArg >= 0) {
					m_pWebBrowser->put_Visible(GetIntFromVariant(&pDispParams->rgvarg[nArg]));
				}
				if (pVarResult) {
					m_pWebBrowser->get_Visible(&pVarResult->boolVal);
					pVarResult->vt = VT_BOOL;
				}
			}
			return S_OK;

		case TE_PROPERTY + 12://DropMode
			if (m_pWebBrowser && g_nBlink == 1) {
				if (nArg >= 0) {
					m_pWebBrowser->put_RegisterAsDropTarget(GetIntFromVariant(&pDispParams->rgvarg[nArg]));
				}
			}
			return S_OK;

		//DIID_DWebBrowserEvents2
		case DISPID_DOWNLOADCOMPLETE:
			if (g_nReload == 1 && m_hwndParent == g_hwndMain) {
				g_nReload = 2;
				SetTimer(g_hwndMain, TET_Reload, 2000, teTimerProc);
			}
			return S_OK;

		case DISPID_BEFORENAVIGATE2:
			if (nArg >= 6 && m_hwndParent == g_hwndMain && pDispParams->rgvarg[nArg].pdispVal == m_pWebBrowser) {
				if (pDispParams->rgvarg[nArg - 6].vt == (VT_BYREF | VT_BOOL)) {
					VARIANT vURL;
					teVariantChangeType(&vURL, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
					if (teIsFileSystem(vURL.bstrVal) || teStartsText(L"::{", vURL.bstrVal) ||
						(StrChr(vURL.bstrVal, '/') && !tePathMatchSpec(vURL.bstrVal, L"file://*;data:*"))) {
						*pDispParams->rgvarg[nArg - 6].pboolVal = OpenNewWindowV(&vURL);
					}
					VariantClear(&vURL);
				}
			}
			break;
		case DISPID_DOCUMENTCOMPLETE:
			if (m_nClose != 2) {
				IUnknown_GetWindow(m_pWebBrowser, &m_hwndBrowser);
				if (g_bsDocumentWrite) {
					ShowWindow(g_hwndMain, SW_SHOWNORMAL);
					g_pWebBrowser->write(g_bsDocumentWrite);
					teSysFreeString(&g_bsDocumentWrite);
					teSetExStyleAnd(m_hwndParent, ~WS_EX_LAYERED);
					return S_OK;
				}
				if (g_bInit) {
					g_bInit = FALSE;
					SetTimer(g_hwndMain, TET_Create, g_nCreateTimer, teTimerProc);
				}
				if (g_pSW) {
					teRegister();
				}
				if (g_hwndMain != m_hwndParent) {
					SetTimer(m_hwndParent, (UINT_PTR)m_hwndParent, 1000, teTimerProcSW2);
				}
				return S_OK;
			}
			PostMessage(m_hwndParent, WM_CLOSE, 0, 0);
			return S_OK;
			/*/// XP(IE8) does not work.
			case DISPID_NAVIGATECOMPLETE2:
			if (g_hwndMain != m_hwndParent) {
			HRESULT hr = E_NOTIMPL;
			IDispatch *pdisp;
			if (m_pWebBrowser->get_Document(&pdisp) == S_OK) {
			IHTMLDocument2 *pdoc;
			if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pdoc))) {
			IHTMLWindow2 *pwin;
			if SUCCEEDED(pdoc->get_parentWindow(&pwin)) {
			hr = tePutProperty(pwin, L"dialogArguments", &m_vData);
			pwin->Release();
			}
			pdoc->Release();
			}
			pdisp->Release();
			}
			}
			return S_OK;
			//*///
		case DISPID_SECURITYDOMAIN:
			return S_OK;
		case DISPID_AMBIENT_DLCONTROL:
			teSetLong(pVarResult, DLCTL_DLIMAGES | DLCTL_VIDEOS | DLCTL_BGSOUNDS);
			return S_OK;
			/*///
			case DISPID_AMBIENT_USERAGENT:
			teSetSZ(pVarResult, g_szTE);
			return S_OK;
			//*///
		case DISPID_NEWWINDOW3:
			if (nArg >= 4 && pDispParams->rgvarg[nArg - 1].vt == (VT_BYREF | VT_BOOL)) {
				*pDispParams->rgvarg[nArg - 1].pboolVal = VARIANT_TRUE;
				OpenNewWindowV(&pDispParams->rgvarg[nArg - 4]);
			}
			return S_OK;
			/*			case DISPID_QUIT:
			return S_OK;
			case DISPID_NAVIGATEERROR:
			return S_OK;*/
		case DISPID_WINDOWCLOSING:
			if (nArg >= 1 && pDispParams->rgvarg[nArg - 1].vt == (VT_BYREF | VT_BOOL)) {
				pDispParams->rgvarg[nArg - 1].pboolVal[0] = VARIANT_TRUE;
			}
			PostMessage(m_hwndParent, WM_CLOSE, 0, 0);
			return S_FALSE;
		case DISPID_FILEDOWNLOAD:
			if (nArg >= 1 && pDispParams->rgvarg[nArg - 1].vt == (VT_BYREF | VT_BOOL)) {
				pDispParams->rgvarg[nArg - 1].pboolVal[0] = VARIANT_TRUE;
			}
			return S_OK;
		case DISPID_TITLECHANGE:
			if (m_bstrPath && nArg >= 0 && pDispParams->rgvarg[nArg].vt == VT_BSTR) {
				teSetWindowText(m_hwndParent, pDispParams->rgvarg[nArg].bstrVal);
			}
			return S_OK;
		case DISPID_VALUE:
			teSetObject(pVarResult, this);
		case DISPID_TE_UNDEFIND:
			return S_OK;
		default:
			if (dispIdMember >= START_OnFunc && dispIdMember < START_OnFunc + Count_WBFunc) {
				teInvokeObject(&m_ppDispatch[dispIdMember - START_OnFunc], wFlags, pVarResult, nArg, pDispParams->rgvarg);
				return S_OK;
			}
		}
	} catch (...) {
		return teException(pExcepInfo, "WebBrowser", methodWB, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IOleClientSite
STDMETHODIMP CteWebBrowser::SaveObject()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::GetContainer(IOleContainer **ppContainer)
{
	*ppContainer = NULL;

	return E_NOINTERFACE;
}

STDMETHODIMP CteWebBrowser::ShowObject()
{
	return S_OK;
}

STDMETHODIMP CteWebBrowser::OnShowWindow(BOOL fShow)
{
	return S_OK;
}

STDMETHODIMP CteWebBrowser::RequestNewObjectLayout()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::GetWindow(HWND *phwnd)
{
	*phwnd = m_hwndParent;
	return S_OK;
}

STDMETHODIMP CteWebBrowser::ContextSensitiveHelp(BOOL fEnterMode)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::CanInPlaceActivate()
{
	return S_OK;
}

STDMETHODIMP CteWebBrowser::OnInPlaceActivate()
{
	return S_OK;
}

STDMETHODIMP CteWebBrowser::OnUIActivate()
{
	return S_OK;
}

STDMETHODIMP CteWebBrowser::GetWindowContext(IOleInPlaceFrame **ppFrame, IOleInPlaceUIWindow **ppDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	*ppFrame = NULL;
	*ppDoc = NULL;

	GetClientRect(m_hwndParent, lprcPosRect);
	CopyRect(lprcClipRect, lprcPosRect);

	return S_OK;
}

STDMETHODIMP CteWebBrowser::Scroll(SIZE scrollExtant)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::OnUIDeactivate(BOOL fUndoable)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::OnInPlaceDeactivate()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::DiscardUndoState()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::DeactivateAndUndo()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::OnPosRectChange(LPCRECT lprcPosRect)
{
	return S_OK;
}

/*///
//IOleCommandTarget
STDMETHODIMP CteWebBrowser::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD *prgCmds, OLECMDTEXT *pCmdText)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut)
{
	if (nCmdID == OLECMDID_SHOWMESSAGE) {
		return E_NOTIMPL;
	}
	if (nCmdID == OLECMDID_SHOWSCRIPTERROR) {
		Sleep(0);
	}
	return E_NOTIMPL;
}
//*/
//IDocHostUIHandler
STDMETHODIMP CteWebBrowser::ShowContextMenu(DWORD dwID, POINT *ppt, IUnknown *pcmdtReserved, IDispatch *pdispReserved)
{
	if (g_hwndMain == m_hwndParent && g_pOnFunc[TE_OnShowContextMenu]) {
		MSG msg1;
		msg1.hwnd = m_hwndParent;
		msg1.message = WM_CONTEXTMENU;
		msg1.wParam = dwID;
		msg1.pt.x = ppt->x;
		msg1.pt.y = ppt->y;
		return MessageSubPt(TE_OnShowContextMenu, this, &msg1);
	}
	return S_FALSE;
}

STDMETHODIMP CteWebBrowser::GetHostInfo(DOCHOSTUIINFO *pInfo)
{
	pInfo->cbSize        = sizeof(DOCHOSTUIINFO);
	pInfo->dwFlags       = DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_SCROLL_NO | DOCHOSTUIFLAG_IME_ENABLE_RECONVERSION;// DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE
	pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT;
	pInfo->pchHostCss    = NULL;
	pInfo->pchHostNS     = NULL;
	return S_OK;
}

STDMETHODIMP CteWebBrowser::ShowUI(DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc)
{
//On resize window
/*/// For check
	IOleInPlaceUIWindow *pUIWindow;
	if SUCCEEDED(pCommandTarget->QueryInterface(IID_PPV_ARGS(&pUIWindow))) {
	}
//*/
	return S_FALSE;
}

STDMETHODIMP CteWebBrowser::HideUI(VOID)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::UpdateUI(VOID)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::EnableModeless(BOOL fEnable)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::OnDocWindowActivate(BOOL fActivate)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::OnFrameWindowActivate(BOOL fActivate)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fFrameWindow)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::TranslateAccelerator(LPMSG lpMsg, const GUID *pguidCmdGroup, DWORD nCmdID)
{
	return S_FALSE;
}

STDMETHODIMP CteWebBrowser::GetOptionKeyPath(LPOLESTR *pchKey, DWORD dw)
{
	return E_NOTIMPL;
/*/// There is a problem
	WCHAR* szKey = L"Software\\tablacus\\explorer";
    if (pchKey) {
		*pchKey = (LPOLESTR)::CoTaskMemAlloc((lstrlen(szKey) + 1) * sizeof(WCHAR));
		if (*pchKey) {
			lstrcpy(*pchKey, szKey);
			return S_OK;
		}
    }
	return E_INVALIDARG;
*///
}

STDMETHODIMP CteWebBrowser::GetDropTarget(IDropTarget *pDropTarget, IDropTarget **ppDropTarget)
{
	SafeRelease(&m_pDropTarget);
	if (pDropTarget) {
		pDropTarget->QueryInterface(IID_PPV_ARGS(&m_pDropTarget));
	}
	return QueryInterface(IID_PPV_ARGS(ppDropTarget));
}

STDMETHODIMP CteWebBrowser::GetExternal(IDispatch **ppDispatch)
{
	return m_ppDispatch[WB_External]->QueryInterface(IID_PPV_ARGS(ppDispatch));
}

STDMETHODIMP CteWebBrowser::TranslateUrl(DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut)
{
	return S_FALSE;
}

STDMETHODIMP CteWebBrowser::FilterDataObject(IDataObject *pDO, IDataObject **ppDORet)
{
	return E_NOTIMPL;
}

// IDocHostShowUI
STDMETHODIMP CteWebBrowser::ShowMessage(HWND hwnd, LPOLESTR lpstrText, LPOLESTR lpstrCaption, DWORD dwType, LPOLESTR lpstrHelpFile, DWORD dwHelpContext, LRESULT *plResult)
{
	if (dwType == (MB_ICONEXCLAMATION | MB_YESNO)) {// Stop running this script?
		*plResult = IDNO;
		return S_OK;
	}
	*plResult = MessageBox(hwnd, lpstrText, _T(PRODUCTNAME), dwType);
	return S_OK;
}

STDMETHODIMP CteWebBrowser::ShowHelp(HWND hwnd, LPOLESTR pszHelpFile, UINT uCommand, DWORD dwData, POINT ptMouse, IDispatch *pDispatchObjectHit)
{
	return E_NOTIMPL;
}

//IDropTarget
STDMETHODIMP CteWebBrowser::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	g_bDragging = TRUE;
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);
	DWORD dwEffect = *pdwEffect;
	HRESULT	hr = DragSub(TE_OnDragEnter, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		*pdwEffect = dwEffect;
		if (m_pDropTarget->DragEnter(pDataObj, grfKeyState, pt, pdwEffect) == S_OK) {
			hr = S_OK;
		}
		if (g_pDropTargetHelper) {
			if (g_bDragIcon) {
				if (g_nBlink == 1) {
					g_pDropTargetHelper->DragEnter(m_hwndBrowser, pDataObj, (LPPOINT)&pt, *pdwEffect);
				}
			} else {
				g_pDropTargetHelper->DragLeave();
			}
		}
	}
	return hr;
}

STDMETHODIMP CteWebBrowser::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_dwEffect = *pdwEffect;
	HRESULT hr = DragSub(TE_OnDragOver, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (hr != S_OK) {
		*pdwEffect = DROPEFFECT_NONE;
	}
	m_dwEffectTE = *pdwEffect;
	if (m_pDropTarget) {
		if (m_pDropTarget->DragOver(grfKeyState, pt, &m_dwEffect) == S_OK) {
			hr = S_OK;
		} else {
			m_dwEffect = DROPEFFECT_NONE;
		}
		if (g_pDropTargetHelper) {
			if (g_bDragIcon) {
				if (g_nBlink == 1) {
					g_pDropTargetHelper->DragOver((LPPOINT)&pt, *pdwEffect);
				}
			} else {
				g_pDropTargetHelper->DragLeave();
			}
		}
	}
	*pdwEffect |= m_dwEffect;
	m_grfKeyState = grfKeyState;
	return hr;
}

STDMETHODIMP CteWebBrowser::DragLeave()
{
	g_bDragging = FALSE;
	if (m_DragLeave == E_NOT_SET) {
		m_DragLeave = DragLeaveSub(this, &m_pDragItems);
		if (m_pDropTarget) {
			HRESULT hr = m_pDropTarget->DragLeave();
			if (m_DragLeave != S_OK) {
				m_DragLeave = hr;
			}
		}
	}
	if (g_pDropTargetHelper && g_nBlink == 1) {
		g_pDropTargetHelper->DragLeave();
	}
	return m_DragLeave;
}

STDMETHODIMP CteWebBrowser::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	HRESULT hr = E_NOTIMPL;
	DWORD dwEffect = *pdwEffect;
	if (g_pDropTargetHelper) {
		g_pDropTargetHelper->DragLeave();
	}
	if (m_dwEffectTE) {
		GetDragItems(&m_pDragItems, pDataObj);
		hr = DragSub(TE_OnDrop, this, m_pDragItems, &m_grfKeyState, pt, pdwEffect);
	}
	if (m_pDropTarget && m_dwEffect) {
		*pdwEffect = dwEffect;
		if (m_pDropTarget->Drop(pDataObj, m_grfKeyState, pt, pdwEffect) == S_OK) {
			hr = S_OK;
		}
	}
	DragLeave();
	return hr;
}

HWND CteWebBrowser::get_HWND()
{
	HWND hwnd = 0;
	IUnknown_GetWindow(m_pWebBrowser, &hwnd);
	return hwnd;
}

BOOL CteWebBrowser::IsBusy()
{
	VARIANT_BOOL vb;
	m_pWebBrowser->get_Busy(&vb);
	if (vb) {
		return vb;
	}
	READYSTATE rs;
	m_pWebBrowser->get_ReadyState(&rs);
	return rs < READYSTATE_INTERACTIVE;
}

VOID CteWebBrowser::write(LPWSTR pszPath)
{
	ClearEvents();
	if (g_nBlink == 1) {
		VARIANT v;
		teSetLong(&v, 1);
		m_pWebBrowser->Navigate(pszPath, &v, NULL, NULL, NULL);
		return;
	}
	IDispatch *pdisp = NULL;
	IHTMLDocument2 *pDoc = NULL;
	do {
		Sleep(1000);
		if (m_pWebBrowser->get_Document(&pdisp) == S_OK && pdisp) {
			pdisp->QueryInterface(IID_PPV_ARGS(&pDoc));
			pdisp->Release();
		}
	} while (!pDoc);
	SAFEARRAY *psa = SafeArrayCreateVector(VT_VARIANT, 0, 1);
	if (psa) {
		LONG l = 0;
		VARIANT v;
		teSetSZ(&v, pszPath);
		if (::SafeArrayPutElement(psa, &l, &v) == S_OK) {
			pDoc->write(psa);
			VariantClear(&v);
		}
		::SafeArrayDestroy(psa);
	}
}

void CteWebBrowser::Close()
{
	try {
		if (m_pWebBrowser) {
			if (m_hwndParent != g_hwndMain) {
				size_t u;
				if (teFindSubWindow(m_hwndParent, &u)) {
					g_pSubWindows.erase(g_pSubWindows.begin() + u);
				}
			}
			m_pWebBrowser->Quit();
			IOleObject *pOleObject;
			if SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleObject))) {
				RECT rc;
				SetRectEmpty(&rc);
				pOleObject->SetClientSite(NULL);
				pOleObject->DoVerb(OLEIVERB_HIDE, NULL, NULL, 0, m_hwndParent, &rc);
				pOleObject->Close(OLECLOSE_NOSAVE);
				pOleObject->Release();
			}
			teUnadviseAndRelease(m_pWebBrowser, DIID_DWebBrowserEvents2, &m_dwCookie);
			m_pWebBrowser = NULL;
		}
		SafeRelease(&m_pDropTarget);
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"WebBrowser::Close";
#endif
	}
}

// CteTabCtrl

CteTabCtrl::CteTabCtrl()
{
	m_cRef = 1;
	m_nLockUpdate = 0;
	m_bVisible = FALSE;
	::ZeroMemory(&m_si, sizeof(m_si));
	m_si.cbSize = sizeof(m_si);
}

void CteTabCtrl::Init()
{
	m_nIndex = -1;
	m_bEmpty = FALSE;
	VariantInit(&m_vData);
	m_param[TE_Left] = 0;
	m_param[TE_Top] = 0;
	m_param[TE_Width] = 100;
	m_param[TE_Height] = 100;
	m_param[TC_Flags] = TCS_FOCUSNEVER | TCS_HOTTRACK | TCS_MULTILINE | TCS_RAGGEDRIGHT | TCS_SCROLLOPPOSITE | TCS_TOOLTIPS;
	m_param[TC_Align] = 0;
	m_param[TC_TabWidth] = 0;
	m_param[TC_TabHeight] = 0;
	m_nScrollWidth = 0;
	m_bRedraw = FALSE;
	m_lNavigating = 0;
}

VOID CteTabCtrl::GetItem(int i, VARIANT *pVarResult)
{
	CteShellBrowser *pSB = GetShellBrowser(i);
	if (pSB) {
		teSetObject(pVarResult, pSB);
	}
}

BOOL CteTabCtrl::SetDefault()
{
	if (g_pTC != this) {
		if (!g_pTC || !g_pTC->m_bVisible || m_bVisible) {
			g_pTC = this;
			return TRUE;
		}
	}
	return FALSE;
}

VOID CteTabCtrl::Show(BOOL bVisible)
{
	bVisible &= 1;
	if (bVisible) {
		SetDefault();
		ShowWindow(m_hwndStatic, SW_SHOW);
		ArrangeWindow();
	} else {
		RECT rc;
		GetWindowRect(m_hwndStatic, &rc);
		MoveWindow(m_hwndStatic, rc.left - rc.right, rc.top - rc.bottom, rc.right - rc.left, rc.bottom - rc.top, TRUE);
		ShowWindow(m_hwndStatic, SW_HIDE);
	}
	if (bVisible ^ m_bVisible) {
		if (m_bVisible) {
			CteShellBrowser *pSB = GetShellBrowser(m_nIndex);
			if (pSB) {
				pSB->Show(FALSE, 1);
			}
			if (g_pTC == this) {
				if (this == g_pTC) {
					for (size_t i = 0; i < g_ppTC.size(); ++i) {
						CteTabCtrl *pTC = g_ppTC[i];
						if (pTC && !pTC->m_bEmpty && pTC->m_bVisible) {
							g_pTC = pTC;
							break;
						}
					}
				}
			}
		}
		m_bVisible = bVisible;
		if (g_pOnFunc[TE_OnVisibleChanged]) {
			VARIANTARG *pv = GetNewVARIANT(2);
			teSetObject(&pv[1], this);
			teSetBool(&pv[0], bVisible);
			Invoke4(g_pOnFunc[TE_OnVisibleChanged], NULL, 2, pv);
		}
	}
}

BOOL CteTabCtrl::Close(BOOL bForce)
{
	if (m_lNavigating) {
		return FALSE;
	}
	if (CanClose(this) || bForce) {
		Show(FALSE);
		for (int nCount = TabCtrl_GetItemCount(m_hwnd); nCount--;) {
			CteShellBrowser *pSB = GetShellBrowser(nCount);
			if (pSB) {
				pSB->Close(TRUE);
			}
		}
		RevokeDragDrop(m_hwnd);
		RemoveWindowSubclass(m_hwnd, TETCProc, 1);
		DestroyWindow(m_hwnd);
		m_hwnd = NULL;

		RevokeDragDrop(m_hwndButton);
		RemoveWindowSubclass(m_hwndButton, TEBTProc, 1);
		DestroyWindow(m_hwndButton);
		m_hwndButton = NULL;

		RemoveWindowSubclass(m_hwndStatic, TESTProc, 1);
		DestroyWindow(m_hwndStatic);
		m_hwndStatic = NULL;
		m_bEmpty = TRUE;
		return TRUE;
	}
	return FALSE;
}

VOID CteTabCtrl::SetItemSize()
{
	DWORD dwSize = MAKELPARAM(m_param[TC_TabWidth] - m_nScrollWidth, m_param[TC_TabHeight]);
	if (dwSize != m_dwSize) {
		SendMessage(m_hwnd, TCM_SETITEMSIZE, 0, dwSize);
		m_dwSize = dwSize;
	}
}

DWORD CteTabCtrl::GetStyle()
{
	DWORD dwStyle = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | TCS_FOCUSNEVER | m_param[TC_Flags];
	if ((dwStyle & (TCS_SCROLLOPPOSITE | TCS_BUTTONS)) == TCS_SCROLLOPPOSITE) {
		if (m_param[TC_Align] == 4 || m_param[TC_Align] == 5) {
			dwStyle |= TCS_BUTTONS;
		}
	}
	if (dwStyle & TCS_BUTTONS) {
		if (dwStyle & TCS_BOTTOM && m_param[TC_Align] > 1) {
			dwStyle &= ~TCS_BOTTOM;
		}
	}
	if (dwStyle & TCS_SCROLLOPPOSITE) {
		if ((dwStyle & TCS_BUTTONS) || !(dwStyle & TCS_MULTILINE)) {
			dwStyle &= ~TCS_SCROLLOPPOSITE;
		}
	}
	return dwStyle;
}

VOID CteTabCtrl::CreateTC()
{
	m_hwnd = CreateWindowEx(
		WS_EX_TOPMOST | WS_EX_CONTROLPARENT | WS_EX_COMPOSITED, //Extended style
		WC_TABCONTROL, // Class Name
		NULL, // Window Name
		GetStyle(), // Window Style
		CW_USEDEFAULT,    // X coordinate
		CW_USEDEFAULT,    // Y coordinate
		CW_USEDEFAULT,    // Width
		CW_USEDEFAULT,    // Height
		m_hwndButton, // Parent window handle
		(HMENU)0, // Child window identifier
		hInst, // Instance handle
		NULL); // WM_CREATE Parameter
	SetWindowLong(m_hwnd, GWL_STYLE, GetStyle());
	ArrangeWindow();
	SetItemSize();
	SetWindowLongPtr(m_hwnd, GWLP_USERDATA, (LONG_PTR)this);
	SetWindowSubclass(m_hwnd, TETCProc, 1, 0);
	RegisterDragDrop(m_hwnd, m_pDropTarget2);
	BringWindowToTop(m_hwnd);
}

BOOL CteTabCtrl::Create()
{
	m_hwndStatic = CreateWindowEx(
		WS_EX_TOPMOST, WC_STATIC, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | SS_NOTIFY,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, g_hwndMain, (HMENU)0, hInst, NULL);
	SetWindowLongPtr(m_hwndStatic, GWLP_USERDATA, (LONG_PTR)this);
	SetWindowSubclass(m_hwndStatic, TESTProc, 1, 0);

	m_hwndButton = CreateWindowEx(
		WS_EX_TOPMOST, WC_BUTTON, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | BS_OWNERDRAW,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
		m_hwndStatic, (HMENU)0, hInst, NULL);

	SetWindowLongPtr(m_hwndButton, GWLP_USERDATA, (LONG_PTR)this);
	SetWindowSubclass(m_hwndButton, TEBTProc, 1, 0);
	m_pDropTarget2 = new CteDropTarget2(m_hwndButton, static_cast<IDispatch *>(this), TRUE);
	RegisterDragDrop(m_hwndButton, m_pDropTarget2);
	BringWindowToTop(m_hwndButton);

	CreateTC();
	DoFunc(TE_OnCreate, this, E_NOTIMPL);
	DoFunc(TE_OnViewCreated, this, E_NOTIMPL);
	ArrangeWindow();
	return TRUE;
}

CteTabCtrl::~CteTabCtrl()
{
	Close(TRUE);
	VariantClear(&m_vData);
	SafeRelease(&m_pDropTarget2);
}

STDMETHODIMP CteTabCtrl::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteTabCtrl, IDispatch),
		QITABENT(CteTabCtrl, IDispatchEx),
		{ 0 },
	};
	if (IsEqualIID(riid, g_ClsIdTC)) {
		*ppvObject = this;
		AddRef();
		return S_OK;
	}
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteTabCtrl::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteTabCtrl::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteTabCtrl::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteTabCtrl::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTabCtrl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodTC, _countof(methodTC), g_maps[MAP_TC], *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteTabCtrl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		IUnknown *punk = NULL;

		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		switch(dispIdMember) {
		case TE_PROPERTY + 1://Data
			if (nArg >= 0) {
				VariantClear(&m_vData);
				VariantCopy(&m_vData, &pDispParams->rgvarg[nArg]);
			}
			teVariantCopy(pVarResult, &m_vData);
			return S_OK;

		case TE_PROPERTY + 2://hwnd
			teSetPtr(pVarResult, m_hwnd);
			return S_OK;

		case TE_PROPERTY + 3://Type
			teSetLong(pVarResult, CTRL_TC);
			return S_OK;

		case TE_METHOD + 6://HitTest (Screen coordinates)
			if (nArg >= 0 && pVarResult) {
				TCHITTESTINFO info;
				GetPointFormVariant(&info.pt, &pDispParams->rgvarg[nArg]);
				UINT flags = nArg >= 1 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : TCHT_ONITEM;
				teSetLong(pVarResult, static_cast<int>(DoHitTest(this, info.pt, flags)));
				if (g_param[TE_Tab] && pVarResult->lVal < 0) {
					ScreenToClient(m_hwnd, &info.pt);
					info.flags = flags;
					pVarResult->lVal = TabCtrl_HitTest(m_hwnd, &info);
					if (!(info.flags & flags)) {
						pVarResult->lVal = -1;
					}
				}
				if (nArg == 0) {
					GetItem(GetIntFromVariantClear(pVarResult), pVarResult);
				}
			}
			return S_OK;

		case TE_METHOD + 7://Move
			if (nArg >= 1) {
				int nSrc, nDest;
				nDest = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
				CteShellBrowser *pSB;
				if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
					if SUCCEEDED(punk->QueryInterface(g_ClsIdSB, (LPVOID *)&pSB)) {
						pSB->m_pTC->Move(pSB->GetTabIndex(), nDest, this);
						pSB->Release();
						return S_OK;
					}
				}
				nSrc = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (nArg >= 2) {
					if (FindUnknown(&pDispParams->rgvarg[nArg - 2], &punk)) {
						CteTabCtrl *pTC;
						if SUCCEEDED(punk->QueryInterface(g_ClsIdTC, (LPVOID *)&pTC)) {
							Move(nSrc, nDest, pTC);
							pTC->Release();
							return S_OK;
						}
					}
				}
				Move(nSrc, nDest, this);
			}
			return S_OK;

		case TE_PROPERTY + 8://Selected
			if (nArg >= 0) {
				int nIndex = -1;
				if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
					CteShellBrowser *pSB;
					if SUCCEEDED(punk->QueryInterface(g_ClsIdSB, (LPVOID *)&pSB)) {
						nIndex = pSB->GetTabIndex();
						pSB->Release();
					}
				}
				if (nIndex >= 0) {
					TabCtrl_SetCurSel(m_hwnd, nIndex);
				}
			}
			if (pVarResult) {
				CteShellBrowser *pSB = GetShellBrowser(m_nIndex);
				if (!pSB) {
					pSB = GetNewShellBrowser(this);
				}
				teSetObject(pVarResult, pSB);
			}
			return S_OK;

		case TE_METHOD + 9://Close
			Close(FALSE);
			return S_OK;

		case TE_PROPERTY + 10://SelectedIndex
			if (nArg >= 0) {
				TabCtrl_SetCurSel(m_hwnd, GetIntFromVariant(&pDispParams->rgvarg[nArg]));
			}
			if (pVarResult) {
				teSetLong(pVarResult, TabCtrl_GetCurSel(m_hwnd));
			}
			return S_OK;

		case TE_PROPERTY + 11://Visible
			if (nArg >= 0) {
				Show(GetIntFromVariant(&pDispParams->rgvarg[nArg]));
			}
			teSetBool(pVarResult, m_bVisible);
			return S_OK;

		case TE_PROPERTY + 12://Id
			teSetLong(pVarResult, m_nTC);
			return S_OK;

		case TE_METHOD + 13://LockUpdate
			LockUpdate();
			return S_OK;

		case TE_METHOD + 14://UnlockUpdate
			UnlockUpdate();
			return S_OK;

		case TE_METHOD + 15://SetOrder
			if (nArg >= 0) {
				IDispatch *pdisp;
				if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
					int nTabs = teGetObjectLength(pdisp);
					if (nTabs == TabCtrl_GetItemCount(m_hwnd)) {
						TC_ITEM *tabs = new TC_ITEM[nTabs];
						LPWSTR *ppszText = new LPWSTR[nTabs];
						for (int i = nTabs; --i >= 0;) {
							ppszText[i] = new WCHAR[MAX_PATH];
							tabs[i].cchTextMax = MAX_PATH;
							tabs[i].pszText = ppszText[i];
							tabs[i].mask = TCIF_IMAGE | TCIF_PARAM | TCIF_TEXT;
							TabCtrl_GetItem(m_hwnd, i, &tabs[i]);
						}
						int nSel = TabCtrl_GetCurFocus(m_hwnd);
						int nNew = -1;
						VARIANT v;
						VariantInit(&v);
						BOOL bChanged = FALSE;
						for (int i = 0; i < nTabs; ++i) {
							teGetPropertyAt(pdisp, i, &v);
							int j = GetIntFromVariantClear(&v);
							if (i != j) {
								TabCtrl_SetItem(m_hwnd, i, &tabs[j]);
								bChanged = TRUE;
							}
							if (j == nSel) {
								nNew = i;
							}
							delete[] ppszText[j];
						}
						delete[] ppszText;
						delete[] tabs;
						if (bChanged) {
							if (nSel != nNew) {
								TabCtrl_SetCurSel(m_hwnd, nNew);
							}
							ArrangeWindow();
						}
					}
				}
			}
			return S_OK;

		case DISPID_TE_ITEM://Item
			if (nArg >= 0) {
				GetItem(GetIntFromVariant(&pDispParams->rgvarg[nArg]), pVarResult);
			}
			return S_OK;//Count

		case DISPID_TE_COUNT:
			if (pVarResult) {
				teSetLong(pVarResult, TabCtrl_GetItemCount(m_hwnd));
			}
			return S_OK;
			//_NewEnum
		case DISPID_NEWENUM:
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, DISPID_UNKNOWN));
			return S_OK;
			//this
		case DISPID_VALUE:
			teSetObject(pVarResult, this);
		case DISPID_TE_UNDEFIND:
			return S_OK;
		}
		if (dispIdMember >= TE_OFFSET && dispIdMember <= TE_OFFSET + TC_TabHeight) {
			if (nArg >= 0 && dispIdMember != TE_OFFSET) {
				int iValue = GetIntFromVariantPP(&pDispParams->rgvarg[nArg], dispIdMember - TE_OFFSET);
				if (m_param[dispIdMember - TE_OFFSET] != iValue) {
					m_param[dispIdMember - TE_OFFSET] = iValue;
					if (m_hwnd) {
						if (dispIdMember == TE_OFFSET + TC_Flags) {
							DWORD dwStyle = GetWindowLong(m_hwnd, GWL_STYLE);
							DWORD dwPrev = GetStyle();
							if ((dwStyle ^ dwPrev) & 0x7fff) {
								if ((dwStyle ^ dwPrev) & (TCS_SCROLLOPPOSITE | TCS_BUTTONS | TCS_TOOLTIPS)) {
									//Remake
									int nTabs = TabCtrl_GetItemCount(m_hwnd);
									int nSel = TabCtrl_GetCurFocus(m_hwnd);
									TC_ITEM *tabs = new TC_ITEM[nTabs];
									LPWSTR *ppszText = new LPWSTR[nTabs];
									for (int i = nTabs; --i >= 0;) {
										ppszText[i] = new WCHAR[MAX_PATH];
										tabs[i].cchTextMax = MAX_PATH;
										tabs[i].pszText = ppszText[i];
										tabs[i].mask = TCIF_IMAGE | TCIF_PARAM | TCIF_TEXT;
										TabCtrl_GetItem(m_hwnd, i, &tabs[i]);
									}
									HIMAGELIST hImage = TabCtrl_GetImageList(m_hwnd);
									HFONT hFont = (HFONT)SendMessage(m_hwnd, WM_GETFONT, 0, 0);
									RevokeDragDrop(m_hwnd);
									RemoveWindowSubclass(m_hwnd, TETCProc, 1);
									DestroyWindow(m_hwnd);

									CreateTC();
									SendMessage(m_hwnd, WM_SETFONT, (WPARAM)hFont, 0);
									TabCtrl_SetImageList(m_hwnd, hImage);
									for (int i = 0; i < nTabs; ++i) {
										TabCtrl_InsertItem(m_hwnd, i, &tabs[i]);
										delete[] ppszText[i];
									}
									delete[] ppszText;
									delete[] tabs;
									TabCtrl_SetCurSel(m_hwnd, nSel);
									DoFunc(TE_OnViewCreated, this, E_NOTIMPL);
								}
							}
						} else if (dispIdMember == TE_OFFSET + TC_TabWidth || dispIdMember == TE_OFFSET + TC_TabHeight) {
							SetItemSize();
						}
						SetWindowLong(m_hwnd, GWL_STYLE, GetStyle());
						ArrangeWindow();
					}
				}
			}
			if (pVarResult) {
				int i = m_param[dispIdMember - TE_OFFSET];
				if (dispIdMember >= TE_OFFSET + TE_Left && dispIdMember <= TE_OFFSET + TE_Height) {
					if (i & 0x3fff) {
						wchar_t pszSize[20];
						swprintf_s(pszSize, 20, L"%g%%", ((float)(m_param[dispIdMember - TE_OFFSET])) / 100);
						teSetSZ(pVarResult, pszSize);
						return S_OK;
					}
					teSetLong(pVarResult, i / 0x4000);
					return S_OK;
				}
				teSetLong(pVarResult, i);
			}
			return S_OK;
		}
		if (dispIdMember >= DISPID_COLLECTION_MIN && dispIdMember <= DISPID_TE_MAX) {
			GetItem(dispIdMember - DISPID_COLLECTION_MIN, pVarResult);
			return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, "TabControl", methodTC, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteTabCtrl::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	return teGetDispId(methodTC, _countof(methodTC), g_maps[MAP_TC], bstrName, pid, TRUE);
}

STDMETHODIMP CteTabCtrl::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller)
{
	return Invoke(id, IID_NULL, lcid, wFlags, pdp, pvarRes, pei, NULL);
}

STDMETHODIMP CteTabCtrl::DeleteMemberByName(BSTR bstrName, DWORD grfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTabCtrl::DeleteMemberByDispID(DISPID id)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTabCtrl::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTabCtrl::GetMemberName(DISPID id, BSTR *pbstrName)
{
	return teGetMemberName(id, pbstrName);
}

STDMETHODIMP CteTabCtrl::GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid)
{
	*pid = (id < DISPID_COLLECTION_MIN) ? DISPID_COLLECTION_MIN : id + 1;
	return *pid < TabCtrl_GetItemCount(m_hwnd) + DISPID_COLLECTION_MIN ? S_OK : S_FALSE;
}

STDMETHODIMP CteTabCtrl::GetNameSpaceParent(IUnknown **ppunk)
{
	return E_NOTIMPL;
}

VOID CteTabCtrl::Move(int nSrc, int nDest, CteTabCtrl *pDestTab)
{
	if (nDest < 0) {
		nDest = TabCtrl_GetItemCount(m_hwnd) - 1;
		if (nDest < 0) {
			nDest = 0;
		}
	}
	BOOL bOtherTab = m_hwnd != pDestTab->m_hwnd;
	if (nSrc != nDest || bOtherTab) {
		int nIndex = TabCtrl_GetCurSel(m_hwnd);
		WCHAR szText[MAX_PATH];
		TC_ITEM tcItem = { TCIF_TEXT | TCIF_IMAGE | TCIF_PARAM, 0, 0, szText, MAX_PATH };
		TabCtrl_GetItem(m_hwnd, nSrc, &tcItem);
		SendMessage(m_hwnd, TCM_DELETEITEM, nSrc, 1/* Don't delete SB flag */);
		UINT ui = static_cast<UINT>(tcItem.lParam) - 1;
		if (ui < g_ppSB.size()) {
			CteShellBrowser *pSB = g_ppSB[ui];
			teSetParent(pSB->m_hwnd, pDestTab->m_hwndStatic);
			pSB->m_pTC = pDestTab;
			TabCtrl_InsertItem(pDestTab->m_hwnd, nDest, &tcItem);
			if (nSrc == nIndex) {
				m_nIndex = -1;
				if (bOtherTab) {
					if (nSrc > 0) {
						--nSrc;
					}
					TabCtrl_SetCurSel(m_hwnd, nSrc);
				} else {
					TabCtrl_SetCurSel(m_hwnd, nDest);
				}
			}
		}
		if (bOtherTab) {
			pDestTab->m_nIndex = nDest;
			TabCtrl_SetCurSel(pDestTab->m_hwnd, nDest);
			ArrangeWindow();
		}
		TabChanged(TRUE);
	}
}

VOID CteTabCtrl::LockUpdate()
{
	if (InterlockedIncrement(&m_nLockUpdate) == 1) {
		if (g_nBlink == 1) {
			return;
		}
		m_bRedraw = TRUE;
		SendMessage(m_hwndStatic, WM_SETREDRAW, FALSE, 0);
	}
}

VOID CteTabCtrl::UnlockUpdate()
{
	if (::InterlockedDecrement(&m_nLockUpdate) > 0) {
		return;
	}
	m_nLockUpdate = 0;
	RedrawUpdate();
}

VOID CteTabCtrl::RedrawUpdate()
{
	if (!m_bRedraw) {
		return;
	}
	SendMessage(m_hwndStatic, WM_SETREDRAW, TRUE, 0);
	CteShellBrowser *pSB = GetShellBrowser(m_nIndex);
	if (!pSB) {
		return;
	}
	pSB->SetRedraw(TRUE);
	if (g_nLockUpdate || m_nLockUpdate) {
		return;
	}
	RedrawWindow(m_hwndStatic, NULL, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_FRAME | RDW_ALLCHILDREN);
	m_bRedraw = FALSE;
}

VOID CteTabCtrl::TabChanged(BOOL bSameTC)
{
	if (bSameTC) {
		m_nIndex = TabCtrl_GetCurSel(m_hwnd);
	}
	CteShellBrowser *pSB = GetShellBrowser(m_nIndex);
	if (pSB) {
		pSB->SetActive(FALSE);
		pSB->SetStatusTextSB(NULL);
		DoFunc(TE_OnSelectionChanged, pSB->m_pTC, E_NOTIMPL);
		ArrangeWindow();
	}
}

CteShellBrowser* CteTabCtrl::GetShellBrowser(int nPage)
{
	if (nPage < 0 || !m_hwnd) {
		return NULL;
	}
	TC_ITEM tcItem;
	tcItem.mask = TCIF_PARAM;
	TabCtrl_GetItem(m_hwnd, nPage, &tcItem);
	UINT ui = static_cast<UINT>(tcItem.lParam) - 1;
	if (ui < g_ppSB.size()) {
		CteShellBrowser *pSB = g_ppSB[ui];
		if (pSB->m_pTC == this) {
			return pSB;
		}
		SendMessage(m_hwnd, TCM_DELETEITEM, nPage, 1);
	}
	return NULL;
}

// CteFolderItems

CteFolderItems::CteFolderItems(IDataObject *pDataObj, FolderItems *pFolderItems)
{
	m_cRef = 1;
	m_oFolderItems = FALSE;
	if (!pDataObj || FAILED(pDataObj->QueryInterface(IID_PPV_ARGS(&m_pDataObj)))) {
		m_oFolderItems = TRUE;
		m_pDataObj = NULL;
	}
	m_dwEffect = teGetDWordFromDataObj(m_pDataObj, &DROPEFFECTFormat, -1);
	m_pidllist = NULL;
	m_nCount = -1;
	m_bsText = NULL;
	m_bUseText = FALSE;
	m_pFolderItems = pFolderItems;
	m_nIndex = 0;
	m_nUseIDListFormat = -1;
	VariantInit(&m_vData);
}

CteFolderItems::~CteFolderItems()
{
	Clear();
}

VOID CteFolderItems::CreateDataObj()
{
	if (!m_pDataObj) {
		AdjustIDListEx();
		if (m_pidllist && m_nCount > 0) {
			IShellFolder *pSF;
			if (GetShellFolder(&pSF, m_pidllist[0])) {
				pSF->GetUIObjectOf(g_hwndMain, m_nCount, const_cast<LPCITEMIDLIST *>(&m_pidllist[1]), IID_IDataObject, NULL, (LPVOID *)&m_pDataObj);
				pSF->Release();
			}
		}
	}
}

VOID CteFolderItems::Clear()
{
	if (m_pidllist) {
		for (int i = m_nCount; i >= 0; --i) {
			teCoTaskMemFree(m_pidllist[i]);
		}
		delete [] m_pidllist;
		m_pidllist = NULL;
	}
	SafeRelease(&m_pDataObj);
	SafeRelease(&m_pFolderItems);
	teSysFreeString(&m_bsText);
	VariantClear(&m_vData);
	if (m_oFolderItems) {
		while (!m_ovFolderItems.empty()) {
			m_ovFolderItems.back()->Release();
			m_ovFolderItems.pop_back();
		}
		m_oFolderItems = FALSE;
	}
}

VOID CteFolderItems::Regenerate(BOOL bFull)
{
	if (!m_oFolderItems) {
		long nCount;
		if SUCCEEDED(get_Count(&nCount)) {
			VARIANT v;
			v.vt = VT_I4;
			FolderItem *pid;
			for (v.lVal = 0; v.lVal < nCount; ++v.lVal) {
				if (Item(v, &pid) == S_OK) {
					if (bFull) {
						CteFolderItem *pid1;
						teQueryFolderItem(&pid, &pid1);
						if (pid1->m_pidl) {
							if (pid1->m_v.vt != VT_EMPTY || GetVarPathFromFolderItem(pid, &pid1->m_v)) {
								if (pid1->m_v.vt == VT_BSTR) {
									if (teIsFileSystem(pid1->m_v.bstrVal)) {
										LPITEMIDLIST pidl;
										if (teGetIDListFromVariant(&pidl, &pid1->m_v)) {
											if (!ILIsEqual(pidl, pid1->m_pidl)) {
												teILCloneReplace(&pid1->m_pidl, pidl);
											}
											teILFreeClear(&pidl);
										}
									}
								}
							}
						}
						pid1->Release();
					}
					m_ovFolderItems.push_back(pid);
				}
			}
			Clear();
			m_oFolderItems = TRUE;
		}
	}
}

VOID CteFolderItems::ItemEx(long nIndex, VARIANT *pVarResult, VARIANT *pVarNew)
{
	VARIANT v;
	teSetLong(&v, nIndex);
	if (pVarNew) {
		teSysFreeString(&m_bsText);
		Regenerate(FALSE);
		if (m_oFolderItems) {
			IUnknown *punk;
			CteFolderItem *pid = NULL;
			if (FindUnknown(pVarNew, &punk)) {
				punk->QueryInterface(g_ClsIdFI, (LPVOID *)&pid);
			}
			if (pid == NULL) {
				pid = new CteFolderItem(pVarNew);
			}
			if (nIndex >= 0) {
				if (m_ovFolderItems[nIndex]) {
					m_ovFolderItems[nIndex]->Release();
				}
				m_ovFolderItems[nIndex] = pid;
			} else {
				m_ovFolderItems.push_back(pid);
			}
		}
	}
	if (pVarResult) {
		FolderItem *pid = NULL;
		if SUCCEEDED(Item(v, &pid)) {
			teSetObjectRelease(pVarResult, pid);
		}
	}
}

BOOL CteFolderItems::CanIDListFormat()
{
	if (m_nUseIDListFormat < 0) {
		AdjustIDListEx();
		for (int i = m_nCount; i > 0; --i) {
			if (ILGetCount(m_pidllist[i]) != 1) {
				m_nUseIDListFormat = 0;
				return FALSE;
			}
		}
		m_nUseIDListFormat = m_nCount;
	}
	return m_nUseIDListFormat;
}

VOID CteFolderItems::AdjustIDListEx()
{
	if (!m_pidllist) {
		m_nUseIDListFormat = -1;
		if (m_oFolderItems) {
			get_Count(&m_nCount);
			m_pidllist = new LPITEMIDLIST[m_nCount + 1];
			m_pidllist[0] = NULL;
			VARIANT v;
			teSetLong(&v, m_nCount);
			FolderItem *pid = NULL;
			while (--v.lVal >= 0) {
				if (Item(v, &pid) == S_OK) {
					if (!teGetIDListFromObject(pid, &m_pidllist[v.lVal + 1])) {
						m_pidllist[v.lVal + 1] = ::ILClone(g_pidls[CSIDL_DESKTOP]);
						m_nUseIDListFormat = 0;
					}
					pid->Release();
				}
			}
		} else if (m_pDataObj) {
			m_pidllist = IDListFormDataObj(m_pDataObj, &m_nCount);
		}
	}
	AdjustIDList(m_pidllist, m_nCount);
}

STDMETHODIMP CteFolderItems::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENTMULTI(CteFolderItems, IDispatch, IDispatchEx),
		QITABENT(CteFolderItems, IDispatchEx),
		QITABENT(CteFolderItems, FolderItems),
		QITABENT(CteFolderItems, IDataObject),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteFolderItems::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteFolderItems::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}
	return m_cRef;
}

STDMETHODIMP CteFolderItems::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteFolderItems::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	if (m_pFolderItems) {
		return m_pFolderItems->GetTypeInfo(iTInfo, lcid, ppTInfo);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (m_pFolderItems) {
		return m_pFolderItems->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
	}
	if (teGetDispId(methodFIs, _countof(methodFIs), g_maps[MAP_FIs], *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteFolderItems::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (m_pFolderItems) {
			return m_pFolderItems->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
		if (dispIdMember >= DISPID_COLLECTION_MIN && dispIdMember <= DISPID_TE_MAX) {
			ItemEx(dispIdMember - DISPID_COLLECTION_MIN, pVarResult, nArg >= 0 ? &pDispParams->rgvarg[nArg] : NULL);
			return S_OK;
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
			teSetObjectRelease(pVarResult, new CteDispatch(static_cast<FolderItems *>(this), 0, dispIdMember));
			return S_OK;
		}
		switch(dispIdMember) {
		case TE_PROPERTY + 2://Application
			IDispatch *pdisp;
			if SUCCEEDED(get_Application(&pdisp)) {
				teSetObject(pVarResult, pdisp);
				pdisp->Release();
			}
			return S_OK;

		case TE_PROPERTY + 3://Parent
			return S_OK;

		case DISPID_TE_ITEM://Item
			if (nArg >= 0) {
				ItemEx(GetIntFromVariant(&pDispParams->rgvarg[nArg]), pVarResult, nArg >= 1 ? &pDispParams->rgvarg[nArg - 1] : NULL);
			}
			return S_OK;

		case DISPID_TE_COUNT://Count
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				get_Count(&pVarResult->lVal);
			}
			return S_OK;

		case DISPID_NEWENUM://_NewEnum
			IUnknown *punk;
			if SUCCEEDED(_NewEnum(&punk)) {
				teSetObject(pVarResult, punk);
				punk->Release();
			}
			return S_OK;

		case TE_METHOD + 8://AddItem
			if (nArg >= 0) {
				ItemEx(-1, NULL, &pDispParams->rgvarg[nArg]);
			}
			return S_OK;

		case TE_PROPERTY + 9://hDrop
			if (pVarResult) {
				int pi[3] = { 0 };
				for (int i = nArg; i >= 0; --i) {
					if (i < 3) {
						pi[i] = GetIntFromVariant(&pDispParams->rgvarg[nArg - i]);
					}
				}
				teSetPtr(pVarResult, GethDrop(pi[0], pi[1], pi[2]));
			}
			return S_OK;

		case TE_METHOD + 10://GetData
			if (pVarResult) {
				IDataObject *pDataObj;
				if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pDataObj))) {
					STGMEDIUM Medium;
					if (pDataObj->GetData(&UNICODEFormat, &Medium) == S_OK) {
						teSetSZ(pVarResult, static_cast<LPCWSTR>(::GlobalLock(Medium.hGlobal)));
						::GlobalUnlock(Medium.hGlobal);
						::ReleaseStgMedium(&Medium);
					} else if (pDataObj->GetData(&TEXTFormat, &Medium) == S_OK) {
						pVarResult->bstrVal = teMultiByteToWideChar(CP_ACP, static_cast<LPCSTR>(::GlobalLock(Medium.hGlobal)), -1);
						pVarResult->vt = VT_BSTR;
						::GlobalUnlock(Medium.hGlobal);
						::ReleaseStgMedium(&Medium);
					}
					pDataObj->Release();
				}
			}
			return S_OK;

		case TE_METHOD + 11://SetData
			if (nArg >= 0) {
				teSysFreeString(&m_bsText);
				VARIANT v;
				teVariantChangeType(&v, &pDispParams->rgvarg[0], VT_BSTR);
				m_bsText = v.bstrVal;
				v.bstrVal = NULL;
				m_bUseText = TRUE;
			}
			return S_OK;

		case DISPID_TE_INDEX://Index
			if (nArg >= 0) {
				m_nIndex = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
			}
			teSetLong(pVarResult, m_nIndex);
			return S_OK;

		case TE_PROPERTY + 0x1001://dwEffect
			if (nArg >= 0) {
				m_dwEffect = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
			}
			teSetLong(pVarResult, m_dwEffect);
			return S_OK;

		case TE_PROPERTY + 0x1002://pdwEffect
			punk = new CteMemory(sizeof(int), &m_dwEffect, 1, L"DWORD");
			teSetObject(pVarResult, punk);
			punk->Release();
			return S_OK;

		case TE_PROPERTY + 0x1003://Data
			if (nArg >= 0) {
				VariantClear(&m_vData);
				VariantCopy(&m_vData, &pDispParams->rgvarg[nArg]);
			}
			teVariantCopy(pVarResult, &m_vData);
			return S_OK;

		case TE_PROPERTY + 0x1004://UseText
			if (nArg >= 0) {
				m_bUseText = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
			}
			teSetBool(pVarResult, m_bUseText);
			return S_OK;
			//this
		case DISPID_VALUE:
			teSetObject(pVarResult, this);
		case DISPID_TE_UNDEFIND:
			return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, "FolderItems", methodFIs, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

STDMETHODIMP CteFolderItems::get_Count(long *plCount)
{
	if (m_pFolderItems) {
		return m_pFolderItems->get_Count(plCount);
	}
	if (m_oFolderItems) {
		*plCount = m_ovFolderItems.size();
		return S_OK;
	}
	if (m_nCount < 0) {
		if (m_pDataObj) {
			STGMEDIUM Medium;
			if (m_pDataObj->GetData(&IDLISTFormat, &Medium) == S_OK) {
				CIDA *pIDA = (CIDA *)::GlobalLock(Medium.hGlobal);
				m_nCount = pIDA ? pIDA->cidl : 0;
				::GlobalUnlock(Medium.hGlobal);
				::ReleaseStgMedium(&Medium);
			} else if (m_pDataObj->GetData(&HDROPFormat, &Medium) == S_OK) {
				m_nCount = DragQueryFile((HDROP)Medium.hGlobal, (UINT)-1, NULL, 0);
				::ReleaseStgMedium(&Medium);
			}
		} else {
			m_nCount = 0;
		}
	}
	*plCount = m_nCount > 0 ? m_nCount : 0;
	return S_OK;
}

STDMETHODIMP CteFolderItems::get_Application(IDispatch **ppid)
{
	if (m_pFolderItems) {
		return m_pFolderItems->get_Application(ppid);
	}
	if (g_pWebBrowser) {
		return g_pWebBrowser->m_pWebBrowser->get_Application(ppid);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::get_Parent(IDispatch **ppid)
{
	if (m_pFolderItems) {
		return m_pFolderItems->get_Parent(ppid);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::Item(VARIANT index, FolderItem **ppid)
{
	*ppid = NULL;
	int i = GetIntFromVariantClear(&index);
	if (i >= 0) {
		if (m_pFolderItems) {
			return m_pFolderItems->Item(index, ppid);
		}
		if (m_oFolderItems) {
			if (i < m_ovFolderItems.size()) {
				return m_ovFolderItems[i]->QueryInterface(IID_PPV_ARGS(ppid));
			}
			*ppid = NULL;
			return E_FAIL;
		}
	}
	if (!m_pidllist) {
		m_pidllist = IDListFormDataObj(m_pDataObj, &m_nCount);
		if (!m_pidllist) {
			*ppid = NULL;
			return E_FAIL;
		}
	}
	if (i >= 0 && i < m_nCount) {
		if (m_pidllist[0]) {
			LPITEMIDLIST pidl = ILCombine(m_pidllist[0], m_pidllist[i + 1]);
			GetFolderItemFromIDList(ppid, pidl);
			teCoTaskMemFree(pidl);
		} else {
			GetFolderItemFromIDList(ppid, m_pidllist[i + 1]);
		}
	} else if (i == -1) {
		if (m_pidllist[0] == NULL) {
			AdjustIDListEx();
		}
		if (m_pidllist[0]) {
			GetFolderItemFromIDList(ppid, m_pidllist[0]);
		}
	}
	return (*ppid != NULL) ? S_OK : E_FAIL;
}

STDMETHODIMP CteFolderItems::_NewEnum(IUnknown **ppunk)
{
	if (m_pFolderItems) {
		return m_pFolderItems->_NewEnum(ppunk);
	}
	*ppunk = static_cast<IDispatch *>(new CteDispatch(static_cast<FolderItems *>(this), 0, DISPID_UNKNOWN));
	return S_OK;
}

//IDispatchEx
STDMETHODIMP CteFolderItems::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	return teGetDispId(methodFIs, _countof(methodFIs), g_maps[MAP_FIs], bstrName, pid, TRUE);
}

STDMETHODIMP CteFolderItems::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller)
{
	return Invoke(id, IID_NULL, lcid, wFlags, pdp, pvarRes, pei, NULL);
}

STDMETHODIMP CteFolderItems::DeleteMemberByName(BSTR bstrName, DWORD grfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::DeleteMemberByDispID(DISPID id)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::GetMemberName(DISPID id, BSTR *pbstrName)
{
	return teGetMemberName(id, pbstrName);
}

STDMETHODIMP CteFolderItems::GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid)
{
	*pid = (id < DISPID_COLLECTION_MIN) ? DISPID_COLLECTION_MIN : id + 1;
	get_Count(&m_nCount);
	return *pid < m_nCount + DISPID_COLLECTION_MIN ? S_OK : S_FALSE;
}

STDMETHODIMP CteFolderItems::GetNameSpaceParent(IUnknown **ppunk)
{
	return E_NOTIMPL;
}

HDROP CteFolderItems::GethDrop(int x, int y, BOOL fNC)
{
	if (m_pDataObj) {
		STGMEDIUM Medium;
		if (m_pDataObj->GetData(&HDROPFormat, &Medium) == S_OK) {
			HDROP hDrop = (HDROP)Medium.hGlobal;
			if (fNC) {
				LPDROPFILES lpDropFiles = (LPDROPFILES)::GlobalLock(hDrop);
				try {
					lpDropFiles->pt.x = x;
					lpDropFiles->pt.y = y;
					lpDropFiles->fNC = fNC;
				} catch (...) {
					g_nException = 0;
#ifdef _DEBUG
					g_strException = L"lpDropFiles";
#endif
				}
				::GlobalUnlock(hDrop);
			}
			return hDrop;
		}
	}
	if (!m_pidllist) {
		m_pidllist = IDListFormDataObj(m_pDataObj, &m_nCount);
		if (!m_pidllist) {
			AdjustIDListEx();
			if (!m_pidllist) {
				return NULL;
			}
		}
	}
	if (m_nCount == 0) {
		return NULL;
	}
	BSTR *pbslist = new BSTR[m_nCount];
	UINT uSize = sizeof(WCHAR);
	VARIANT v;
	v.vt = VT_I4;
	for (int i = m_nCount; --i >= 0;) {
		pbslist[i] = NULL;
		v.lVal = i;
		FolderItem *pid;
		if SUCCEEDED(Item(v, &pid)) {
			BSTR bsPath;
			if SUCCEEDED(pid->get_Path(&bsPath)) {
				int nLen = ::SysStringByteLen(bsPath);
				if (nLen) {
					pbslist[i] = bsPath;
					uSize += nLen + sizeof(WCHAR);
				} else {
					::SysFreeString(bsPath);
				}
			}
			pid->Release();
		}
	}
	HDROP hDrop = (HDROP)::GlobalAlloc(GHND, sizeof(DROPFILES) + uSize);
	LPDROPFILES lpDropFiles = (LPDROPFILES)::GlobalLock(hDrop);
	try {
		lpDropFiles->pFiles = sizeof(DROPFILES);
		lpDropFiles->pt.x = x;
		lpDropFiles->pt.y = y;
		lpDropFiles->fNC = fNC;
		lpDropFiles->fWide = TRUE;

		LPWSTR lp = (LPWSTR)&lpDropFiles[1];
		for (int i = 0; i < m_nCount; ++i) {
			if (pbslist[i]) {
				lstrcpy(lp, pbslist[i]);
				lp += (SysStringByteLen(pbslist[i]) / 2) + 1;
				::SysFreeString(pbslist[i]);
			}
		}
		*lp = 0;
		delete [] pbslist;
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"GethDrop";
#endif
	}
	::GlobalUnlock(hDrop);
	return hDrop;
}

//IDataObject
STDMETHODIMP CteFolderItems::GetData(FORMATETC *pformatetcIn, STGMEDIUM *pmedium)
{
	if (m_dwEffect != (DWORD)-1 && pformatetcIn->cfFormat == DROPEFFECTFormat.cfFormat) {
		HGLOBAL hGlobal = ::GlobalAlloc(GHND, sizeof(DWORD));
		DWORD *pdwEffect = (DWORD *)::GlobalLock(hGlobal);
		try {
			if (pdwEffect) {
				*pdwEffect = m_dwEffect;
			}
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"pdwEffect";
#endif
		}
		::GlobalUnlock(hGlobal);

		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = hGlobal;
		pmedium->pUnkForRelease = NULL;
		return S_OK;
	}
	if (m_pDataObj) {
		HRESULT hr = m_pDataObj->GetData(pformatetcIn, pmedium);
		if (hr == S_OK) {
			if (g_bsClipRoot) {
				if (pformatetcIn->cfFormat == IDLISTFormat.cfFormat || pformatetcIn->cfFormat == CF_HDROP) {
					BSTR bs = NULL;
					teGetRootFromDataObj(&bs, this);
					if (teStrSameIFree(bs, g_bsClipRoot)) {
						if (g_pOnFunc[TE_OnBeforeGetData]) {
							VARIANTARG *pv = GetNewVARIANT(3);
							teSetObject(&pv[2], g_pTE);
							teSetObjectRelease(&pv[1], this);
							teSetLong(&pv[0], 5);
							Invoke4(g_pOnFunc[TE_OnBeforeGetData], NULL, 3, pv);
						}
					}
					teSysFreeString(&g_bsClipRoot);
				}
			}
			return hr;
		}
	}
	if (pformatetcIn->cfFormat == IDLISTFormat.cfFormat) {
		if (!m_pidllist) {
			m_pidllist = IDListFormDataObj(m_pDataObj, &m_nCount);
			if (!m_pidllist) {
				return DATA_E_FORMATETC;
			}
		}
		AdjustIDListEx();
		UINT uIndex = sizeof(UINT) * (m_nCount + 2);
		UINT uSize = uIndex;
		for (int i = 0; i <= m_nCount; ++i) {
			uSize += ::ILGetSize(m_pidllist[i]);
		}
		HGLOBAL hGlobal = ::GlobalAlloc(GHND, uSize);
		CIDA *pIDA = (CIDA *)::GlobalLock(hGlobal);
		try {
			if (pIDA) {
				pIDA->cidl = m_nCount;
				for (int i = 0; i <= m_nCount; ++i) {
					pIDA->aoffset[i] = uIndex;
					UINT u = ::ILGetSize(m_pidllist[i]);
					char *pc = (char *)pIDA + uIndex;
					::CopyMemory(pc, m_pidllist[i], u);
					uIndex += u;
				}
			}
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"CIDA";
#endif
		}
		::GlobalUnlock(hGlobal);
		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = hGlobal;
		pmedium->pUnkForRelease = NULL;
		return S_OK;
	}
	if (pformatetcIn->cfFormat == CF_HDROP) {
		pmedium->tymed = TYMED_HGLOBAL;
		if (pmedium->hGlobal = (HGLOBAL)GethDrop(0, 0, FALSE)) {
			pmedium->pUnkForRelease = NULL;
			BSTR bs = NULL;
			teGetRootFromDataObj(&bs, this);
			if (teStrSameIFree(bs, g_bsClipRoot)) {
				if (g_pOnFunc[TE_OnBeforeGetData]) {
					VARIANTARG *pv = GetNewVARIANT(3);
					teSetObject(&pv[2], g_pTE);
					teSetObjectRelease(&pv[1], this);
					teSetLong(&pv[0], 5);
					Invoke4(g_pOnFunc[TE_OnBeforeGetData], NULL, 3, pv);
				}
			}
			teSysFreeString(&g_bsClipRoot);
			return S_OK;
		}
	}

	if (pformatetcIn->cfFormat == CF_UNICODETEXT || pformatetcIn->cfFormat == CF_TEXT) {
		if (!m_bsText) {
			try {
				VARIANT vResult, vStr;
				VariantInit(&vResult);
				if SUCCEEDED(DoFunc1(TE_OnClipboardText, this, &vResult)) {
					teVariantChangeType(&vStr, &vResult, VT_BSTR);
					VariantClear(&vResult);
					m_bsText = vStr.bstrVal;
					vStr.bstrVal = NULL;
				}
			} catch (...) {
				g_nException = 0;
#ifdef _DEBUG
				g_strException = L"OnClipboardText";
#endif
			}
		}
		HGLOBAL hMem;
		int nLen = ::SysStringLen(m_bsText) + 1;
		if (pformatetcIn->cfFormat == CF_UNICODETEXT) {
			nLen = ::SysStringByteLen(m_bsText) + sizeof(WCHAR);
			hMem = ::GlobalAlloc(GHND, nLen);
			try {
				LPWSTR lp = static_cast<LPWSTR>(::GlobalLock(hMem));
				if (lp) {
					::CopyMemory(lp, m_bsText, nLen);
				}
			} catch (...) {
				g_nException = 0;
#ifdef _DEBUG
				g_strException = L"CF_UNICODETEXT";
#endif
			}
		} else {
			int nLenA = ::WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)m_bsText, nLen, NULL, 0, NULL, NULL);
			hMem = ::GlobalAlloc(GHND, nLenA);
			try {
				LPSTR lpA = static_cast<LPSTR>(::GlobalLock(hMem));
				if (lpA) {
					::WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)m_bsText, nLen, lpA, nLenA, NULL, NULL);
				}
			} catch (...) {
				g_nException = 0;
#ifdef _DEBUG
				g_strException = L"CF_TEXT";
#endif
			}
		}

		::GlobalUnlock(hMem);
		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = hMem;
		pmedium->pUnkForRelease = NULL;
		return S_OK;
	}
	return DATA_E_FORMATETC;
}

STDMETHODIMP CteFolderItems::GetDataHere(FORMATETC *pformatetc, STGMEDIUM *pmedium)
{
	CreateDataObj();
	if (m_pDataObj) {
		return m_pDataObj->GetDataHere(pformatetc, pmedium);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::QueryGetData(FORMATETC *pformatetc)
{
	if (m_dwEffect != (DWORD)-1 && pformatetc->cfFormat == DROPEFFECTFormat.cfFormat) {
		return S_OK;
	}
	if (QueryGetData2(pformatetc) == S_OK) {
		return S_OK;
	}
	if (m_pDataObj) {
		return m_pDataObj->QueryGetData(pformatetc);
	}
	return DATA_E_FORMATETC;
}

HRESULT CteFolderItems::QueryGetData2(FORMATETC *pformatetc)
{
	if (m_bUseText) {
		if (pformatetc->cfFormat == CF_UNICODETEXT) {
			return S_OK;
		}
		if (pformatetc->cfFormat == CF_TEXT) {
			return S_OK;
		}
	}
	if (m_nCount > 0) {
		if (pformatetc->cfFormat == CF_HDROP) {
			return S_OK;
		}
		if (pformatetc->cfFormat == IDLISTFormat.cfFormat && CanIDListFormat()) {
			return S_OK;
		}
	}
	return DATA_E_FORMATETC;
}

STDMETHODIMP CteFolderItems::GetCanonicalFormatEtc(FORMATETC *pformatectIn, FORMATETC *pformatetcOut)
{
	CreateDataObj();
	if (m_pDataObj) {
		return m_pDataObj->GetCanonicalFormatEtc(pformatectIn, pformatetcOut);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::SetData(FORMATETC *pformatetc, STGMEDIUM *pmedium, BOOL fRelease)
{
	CreateDataObj();
	if (m_pDataObj) {
		return m_pDataObj->SetData(pformatetc, pmedium, fRelease);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenumFormatEtc)
{
	CreateDataObj();
	if (dwDirection == DATADIR_GET) {
		BOOL bIDListFormat = CanIDListFormat();
		BOOL bHDROPFormat = m_nCount > 0;
		BOOL bUNICODEFormat = m_bUseText;
		BOOL bTEXTFormat = m_bUseText;
		BOOL bDROPEFFECTFormat = m_dwEffect != (DWORD)-1;
		std::vector<FORMATETC> formats;
		formats.reserve(m_nCount);
		if (m_pDataObj) {
			IEnumFORMATETC *penumFormatEtc;
			if SUCCEEDED(m_pDataObj->EnumFormatEtc(DATADIR_GET, &penumFormatEtc)) {
				FORMATETC format1;
				while (penumFormatEtc->Next(1, &format1, NULL) == S_OK) {
					if (format1.cfFormat == IDLISTFormat.cfFormat) {
						if (bIDListFormat) {
							bIDListFormat = FALSE;
							formats.push_back(format1);
						}
					} else if (format1.cfFormat == HDROPFormat.cfFormat) {
						if (bHDROPFormat) {
							bHDROPFormat = FALSE;
							formats.push_back(format1);
						}
					} else if (format1.cfFormat == CF_UNICODETEXT) {
						if (bUNICODEFormat) {
							bUNICODEFormat = FALSE;
							formats.push_back(format1);
						}
					} else if (format1.cfFormat == CF_TEXT) {
						if (bTEXTFormat) {
							bTEXTFormat = FALSE;
							formats.push_back(format1);
						}
					} else if (format1.cfFormat == DROPEFFECTFormat.cfFormat) {
						if (bDROPEFFECTFormat) {
							bDROPEFFECTFormat = FALSE;
							formats.push_back(format1);
						}
					} else {
						formats.push_back(format1);
					}
				}
				penumFormatEtc->Release();
			}
		}
		if (bIDListFormat) {
			formats.push_back(IDLISTFormat);
		}
		if (bHDROPFormat) {
			formats.push_back(HDROPFormat);
		}
		if (bUNICODEFormat) {
			formats.push_back(UNICODEFormat);
		}
		if (bTEXTFormat) {
			formats.push_back(TEXTFormat);
		}
		if (bDROPEFFECTFormat) {
			formats.push_back(DROPEFFECTFormat);
		}
		return CreateFormatEnumerator(formats.size(), &formats[0], ppenumFormatEtc);
	}
	if (m_pDataObj) {
		return m_pDataObj->EnumFormatEtc(dwDirection, ppenumFormatEtc);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::DAdvise(FORMATETC *pformatetc, DWORD advf, IAdviseSink *pAdvSink, DWORD *pdwConnection)
{
	CreateDataObj();
	if (m_pDataObj) {
		return m_pDataObj->DAdvise(pformatetc, advf, pAdvSink, pdwConnection);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::DUnadvise(DWORD dwConnection)
{
	if (m_pDataObj) {
		return m_pDataObj->DUnadvise(dwConnection);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::EnumDAdvise(IEnumSTATDATA **ppenumAdvise)
{
	CreateDataObj();
	if (m_pDataObj) {
		return m_pDataObj->EnumDAdvise(ppenumAdvise);
	}
	return E_NOTIMPL;
}

//CteServiceProvider
STDMETHODIMP CteServiceProvider::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteServiceProvider, IServiceProvider),
		{ 0 },
	};
	HRESULT hr = QISearch(this, qit, riid, ppvObject);
	return SUCCEEDED(hr) ? hr : m_pUnk->QueryInterface(riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteServiceProvider::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteServiceProvider::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteServiceProvider::QueryService(REFGUID guidService, REFIID riid, void **ppv)
{
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IServiceProvider)) {
		return E_NOINTERFACE;
	}
	if (IsEqualIID(riid, IID_IShellBrowser) || IsEqualIID(riid, IID_ICommDlgBrowser) || IsEqualIID(riid, IID_ICommDlgBrowser2) || //IsEqualIID(riid, IID_ICommDlgBrowser3)||
		IsEqualIID(riid, IID_IExplorerPaneVisibility)) {
		return m_pUnk->QueryInterface(riid, ppv);
	}
	if (IsEqualIID(riid, IID_IInternetSecurityManager)) {
		*ppv = static_cast<IInternetSecurityManager *>(new CteInternetSecurityManager());
		return S_OK;
	}
	if (IsEqualIID(riid, IID_INewWindowManager)) {
		*ppv = static_cast<INewWindowManager *>(new CteNewWindowManager());
		return S_OK;
	}
	if (m_pUnk2) {
		return m_pUnk2->QueryInterface(riid, ppv);
	}
	return E_NOINTERFACE;
}

CteServiceProvider::CteServiceProvider(IUnknown *punk, IUnknown *punk2)
{
	m_cRef = 1;
	m_pUnk = punk;
	m_pUnk2 = punk2;
}

CteServiceProvider::~CteServiceProvider()
{
}

//CteMemory

CteMemory::CteMemory(int nSize, void *pc, int nCount, LPWSTR lpStruct)
{
	BOOL bSafeArray = FALSE;
	m_cRef = 1;
	m_pc = (char *)pc;
	m_bsStruct = NULL;
	m_nStructIndex = -1;
	if (lstrcmpi(lpStruct, L"SAFEARRAY") == 0) {
		lpStruct = NULL;
		bSafeArray = TRUE;
	}
	if (lpStruct) {
		m_nStructIndex = teBSearchStruct(pTEStructs, _countof(pTEStructs), g_maps[MAP_SS], lpStruct);
		m_bsStruct = ::SysAllocString(lpStruct);
	}
	m_nCount = nCount;
	m_bsAlloc = NULL;
	m_nSize = 0;
	if (nSize > 0) {
		m_nSize = nSize;
		if (pc == NULL || bSafeArray) {
			m_bsAlloc = ::SysAllocStringByteLen(NULL, nSize);
			m_pc = (char *)m_bsAlloc;
			if (m_pc) {
				PVOID pvData;
				if (bSafeArray && pc && ::SafeArrayAccessData((SAFEARRAY *)pc, &pvData) == S_OK) {
					::CopyMemory(m_pc, pvData, nSize);
					::SafeArrayUnaccessData((SAFEARRAY *)pc);
				} else {
					::ZeroMemory(m_pc, nSize);
				}
				if (m_nStructIndex >= 0) {
					if (PathMatchSpecA(pTEStructs[m_nStructIndex].pMethod[0].name, "*Size")) {
						*(int *)m_pc = nSize;
					}
				}
			}
		}
	}
}

CteMemory::~CteMemory()
{
	Free(TRUE);
}

void CteMemory::Free(BOOL bpbs)
{
	teSysFreeString(&m_bsAlloc);
	teSysFreeString(&m_bsStruct);
	while (!m_ppbs.empty()) {
		::SysFreeString(m_ppbs.back());
		m_ppbs.pop_back();
	}
	m_pc = NULL;
}

STDMETHODIMP CteMemory::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteMemory, IDispatch),
		QITABENT(CteMemory, IDispatchEx),
		{ 0 },
	};
	if (IsEqualIID(riid, g_ClsIdStruct)) {
		*ppvObject = this;
		AddRef();
		return S_OK;
	}
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteMemory::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteMemory::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteMemory::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteMemory::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteMemory::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (GetDispID(*rgszNames, 0, rgDispId) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteMemory::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD && dispIdMember <= TE_METHOD_MAX) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		switch(dispIdMember) {
			//P
			case TE_PROPERTY + 0xfff1:
				teSetPtr(pVarResult, m_pc);
				return S_OK;
			//Item
			case DISPID_TE_ITEM:
				if (nArg >= 0) {
					ItemEx(GetIntFromVariant(&pDispParams->rgvarg[nArg]), pVarResult, nArg >= 1 ? &pDispParams->rgvarg[nArg - 1] : NULL);
				}
				return S_OK;
			//Count
			case DISPID_TE_COUNT:
				teSetLong(pVarResult, m_nCount);
				return S_OK;
			//_NewEnum
			case DISPID_NEWENUM:
				teSetObjectRelease(pVarResult, new CteDispatch(this, 0, DISPID_UNKNOWN));
				return S_OK;
			//Read
			case TE_METHOD + 4:
				if (pVarResult) {
					if (nArg >= 1) {
						pVarResult->vt = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
						int nLen = -1;
						if (nArg >= 2) {
							nLen = GetIntFromVariant(&pDispParams->rgvarg[nArg] - 2);
						}
						Read(GetIntFromVariant(&pDispParams->rgvarg[nArg]), nLen, pVarResult);
					}
				}
				return S_OK;
			//Write
			case TE_METHOD + 5:
				if (nArg >= 2) {
					int nLen = -1;
					if (nArg >= 3) {
						nLen = GetIntFromVariant(&pDispParams->rgvarg[nArg] - 3);
					}
					Write(GetIntFromVariant(&pDispParams->rgvarg[nArg]), nLen, GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), &pDispParams->rgvarg[nArg - 2]);
				}
				return S_OK;
			//Size
			case TE_PROPERTY + 0xfff6:
				teSetLong(pVarResult, m_nSize);
				return S_OK;
			//Free
			case TE_METHOD + 7:
				Free(TRUE);
				return S_OK;
			//Clone
			case TE_METHOD + 8:
				if (m_nSize) {
					CteMemory *pMem = new CteMemory(m_nSize, NULL, m_nCount, m_bsStruct);
					::CopyMemory(pMem->m_pc, m_pc, m_nSize);
					teSetObjectRelease(pVarResult, pMem);
				}
				return S_OK;
			//_BLOB
			case TE_PROPERTY + 0xfff9:
				if (m_bsAlloc) {
					if (nArg >= 0 && pDispParams->rgvarg[nArg].vt == VT_BSTR) {
						int nSize = ::SysStringByteLen(pDispParams->rgvarg[nArg].bstrVal);
						if (nSize != m_nSize) {
							m_bsAlloc = ::SysAllocStringByteLen(NULL, nSize);
							m_pc = (char *)m_bsAlloc;
							m_nSize = nSize;
						}
						::CopyMemory(m_pc, pDispParams->rgvarg[nArg].bstrVal, m_nSize);
					}
					if (pVarResult) {
						pVarResult->vt = VT_BSTR;
						pVarResult->bstrVal = ::SysAllocStringByteLen(m_pc, m_nSize);
					}
				}
				return S_OK;

			case DISPID_VALUE:
				teSetObject(pVarResult, this);
			case DISPID_TE_UNDEFIND:
				return S_OK;
			default:
				if (dispIdMember >= DISPID_COLLECTION_MIN && dispIdMember <= DISPID_TE_MAX) {
					ItemEx(dispIdMember - DISPID_COLLECTION_MIN, pVarResult, nArg >= 0 ? &pDispParams->rgvarg[nArg] : NULL);
					return S_OK;
				}
				if (m_pc) {
					VARTYPE vt = dispIdMember >> TE_VT;
					if (vt) {
						if (wFlags & DISPATCH_METHOD) {
							teSetLong(pVarResult, dispIdMember & TE_VI);
							return S_OK;
						}
						if (nArg >= 0) {
							if (vt == VT_BSTR && pDispParams->rgvarg[nArg].vt == VT_BSTR) {
								VARIANT v;
								teSetPtr(&v, AddBSTR(pDispParams->rgvarg[nArg].bstrVal));
								Write(dispIdMember & TE_VI, -1, vt, &v);
								VariantClear(&v);
							} else {
								Write(dispIdMember & TE_VI, -1, vt, &pDispParams->rgvarg[nArg]);
							}
						}
						if (pVarResult) {
							pVarResult->vt = vt;
							Read(dispIdMember & TE_VI, -1, pVarResult);
						}
					}
					return S_OK;
				}
		}
	} catch (...) {
		return teException(pExcepInfo, "Memory", methodMem, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteMemory::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	if (m_nStructIndex >= 0 && teGetDispId(pTEStructs[m_nStructIndex].pMethod, 0, NULL, bstrName, pid, false) == S_OK) {
		return S_OK;
	}
	if (teGetDispId(methodMem, _countof(methodMem), g_maps[MAP_Mem], bstrName, pid, TRUE) == S_OK) {
		return S_OK;
	}
	return teGetDispId(methodMem2, 0, NULL, m_bsStruct, pid, false);
}

STDMETHODIMP CteMemory::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller)
{
	return Invoke(id, IID_NULL, lcid, wFlags, pdp, pvarRes, pei, NULL);
}

STDMETHODIMP CteMemory::DeleteMemberByName(BSTR bstrName, DWORD grfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteMemory::DeleteMemberByDispID(DISPID id)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteMemory::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteMemory::GetMemberName(DISPID id, BSTR *pbstrName)
{
	return teGetMemberName(id, pbstrName);
}

STDMETHODIMP CteMemory::GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid)
{
	*pid = (id < DISPID_COLLECTION_MIN) ? DISPID_COLLECTION_MIN : id + 1;
	return *pid < m_nCount + DISPID_COLLECTION_MIN ? S_OK : S_FALSE;
}

STDMETHODIMP CteMemory::GetNameSpaceParent(IUnknown **ppunk)
{
	return E_NOTIMPL;
}

VOID CteMemory::ItemEx(int i, VARIANT *pVarResult, VARIANT *pVarNew)
{
	if (i >= 0 && i < m_nCount) {
		if (pVarNew) {
			int nSize = m_nSize / m_nCount;
			if (nSize >= 1 && nSize <= 8) {
				LARGE_INTEGER li;
				li.QuadPart = GetLLFromVariant(pVarNew);
				if (li.QuadPart == 0 && pVarNew->vt == VT_BSTR && pVarNew->bstrVal &&
				(lstrcmpi(m_bsStruct, L"BSTR") == 0 || lstrcmpi(m_bsStruct, L"LPWSTR") == 0)) {
					li.QuadPart = (LONGLONG)AddBSTR(pVarNew->bstrVal);
				}
				if (nSize == 1) {//BYTE
					m_pc[i] = LOBYTE(li.LowPart);
				} else if (nSize == 2) {//WORD
					WORD *p = (WORD *)m_pc;
					p[i] = LOWORD(li.LowPart);
				} else if (nSize == 4) {//DWORD
					int *p = (int *)m_pc;
					p[i] = li.LowPart;
				} else if (nSize == 8) {//LONGLONG
					LONGLONG *p = (LONGLONG *)m_pc;
					p[i] = li.QuadPart;
				}
			}
		}
		if (pVarResult) {
			if (m_nCount) {
				int nSize = m_nSize / m_nCount;
				if (nSize) {
					if (nSize == 1) {//BYTE/char
						teSetLong(pVarResult, m_pc[i]);
					} else if (nSize == 2) {//WORD/WCHAR
						WORD *p = (WORD *)m_pc;
						teSetLong(pVarResult, p[i]);
					} else if (nSize == 4) {//DWORD/int
						int *p = (int *)m_pc;
						teSetLong(pVarResult, p[i]);
					} else if (nSize == 8 && lstrcmpi(m_bsStruct, L"POINT")) {//LONGLONG
						LONGLONG *p = (LONGLONG *)m_pc;
						teSetLL(pVarResult, p[i]);
					} else {
						teSetObjectRelease(pVarResult, new CteMemory(nSize, m_pc + (nSize * i), 1, m_bsStruct));
					}
				}
			}
		}
	}
}

BSTR CteMemory::AddBSTR(BSTR bs)
{
	BSTR bsNew = ::SysAllocStringByteLen((char *)bs, ::SysStringByteLen(bs));
	m_ppbs.push_back(bsNew);
	return bsNew;
}

VOID CteMemory::Read(int nIndex, int nLen, VARIANT *pVarResult)
{
	CteMemory *pMem;
	try {
		if (pVarResult->vt & VT_ARRAY) {
			VARTYPE vt = LOBYTE(pVarResult->vt);
			int nSize = SizeOfvt(vt);
			if (nSize) {
				if (nLen < 0) {
					nLen = (m_nSize - nIndex) / nSize;
				}
				if (nLen > 0) {
					SAFEARRAY *psa = SafeArrayCreateVector(vt, 0, nLen);
					if (psa) {
						PVOID pvData;
						if (::SafeArrayAccessData(psa, &pvData) == S_OK) {
							::CopyMemory(pvData, &m_pc[nIndex], nLen * nSize);
							::SafeArrayUnaccessData(psa);
							pVarResult->parray = psa;
						}
					}
				}
			}
			return;
		}
		teParam *pFrom = (teParam *)&m_pc[nIndex];
		switch (pVarResult->vt) {
			case VT_BOOL:
				pVarResult->boolVal = pFrom->boolVal;
				break;
			case VT_I1:
			case VT_UI1:
				pVarResult->bVal = pFrom->bVal;
				break;
			case VT_I2:
			case VT_UI2:
				pVarResult->iVal = pFrom->iVal;
				break;
			case VT_I4:
			case VT_UI4:
				pVarResult->lVal = pFrom->lVal;
				break;
			case VT_I8:
			case VT_UI8:
				teSetLL(pVarResult, pFrom->llVal);
				break;
			case VT_R4:
				pVarResult->fltVal = pFrom->fltVal;
				break;
			case VT_R8:
				pVarResult->dblVal = pFrom->dblVal;
				break;
			case VT_PTR:
			case VT_BSTR:
				teSetPtr(pVarResult, pFrom->uint_ptr);
				break;
			case VT_LPWSTR:
				if (nLen >= 0) {
					if (nLen * (int)sizeof(WCHAR) > m_nSize) {
						nLen = m_nSize / sizeof(WCHAR);
					}
					pVarResult->bstrVal = teSysAllocStringLen((WCHAR *)pFrom, nLen);
					pVarResult->vt = VT_BSTR;
				} else {
					teSetSZ(pVarResult, (WCHAR *)pFrom);
				}
				break;
			case VT_LPSTR:
			case VT_USERDEFINED:
				if (nLen > m_nSize) {
					nLen = m_nSize;
				}
				pVarResult->bstrVal = teMultiByteToWideChar(pVarResult->vt != VT_USERDEFINED ? CP_ACP : CP_UTF8, (LPCSTR)pFrom, nLen);
				pVarResult->vt = VT_BSTR;
				break;
			case VT_FILETIME://UNIX EPOCH
				teSetLL(pVarResult, *(ULONGLONG *)pFrom / 10000 - 11644473600000LL);
//				teFileTimeToVariantTime((LPFILETIME)pFrom, &pVarResult->date);
//				pVarResult->vt = VT_DATE;
				break;
			case VT_CY:
				POINT *ppt;
				ppt = (POINT *)pFrom;
				pMem = new CteMemory(sizeof(POINT), NULL, 1, L"POINT");
				pMem->SetPoint(ppt->x, ppt->y);
				teSetObjectRelease(pVarResult, pMem);
				break;
			case VT_CARRAY:
				pMem = new CteMemory(sizeof(RECT), NULL, 1, L"RECT");
				::CopyMemory(pMem->m_pc, pFrom, sizeof(RECT));
				teSetObjectRelease(pVarResult, pMem);
				break;
			case VT_CLSID:
				LPOLESTR lpsz;
				StringFromCLSID(pFrom->iid, &lpsz);
				teSetSZ(pVarResult, lpsz);
				teCoTaskMemFree(lpsz);
				break;
			case VT_VARIANT:
				teVariantCopy(pVarResult, (VARIANT *)pFrom);
				break;
/*			case VT_RECORD:
				teSetIDList(pVarResult, *(LPITEMIDLIST *)&m_pc[nIndex]);
				break;*/
			default:
				pVarResult->vt = VT_EMPTY;
				break;
		}//end_switch
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"CteMemory::Read";
#endif
	}
}

//Write
VOID CteMemory::Write(int nIndex, int nLen, VARTYPE vt, VARIANT *pv)
{
	try {
		if (pv->vt & VT_ARRAY) {
			int nSize = SizeOfvt(pv->vt);
			if (nSize) {
				if (nLen < 0) {
					nLen = (m_nSize - nIndex) / nSize;
				}
				if (nLen > 0) {
					PVOID pvData;
					if (::SafeArrayAccessData(pv->parray, &pvData) == S_OK) {
						::CopyMemory(&m_pc[nIndex], pvData, nLen * nSize);
						::SafeArrayUnaccessData(pv->parray);
					}
				}
			}
			return;
		}
		teParam d;
		d.llVal = GetLLFromVariant(pv);
		teParam *pDest = (teParam *)&m_pc[nIndex];
		switch (vt) {
			case VT_I1:
			case VT_UI1:
				pDest->bVal = d.bVal;
				break;
			case VT_I2:
			case VT_UI2:
				pDest->iVal = d.iVal;
				break;
			case VT_BOOL:
			case VT_I4:
			case VT_UI4:
				pDest->lVal = d.lVal;
				break;
			case VT_I8:
			case VT_UI8:
				pDest->llVal = d.llVal;
				break;
			case VT_PTR:
			case VT_BSTR:
				pDest->uint_ptr = d.uint_ptr;
				break;
			case VT_LPWSTR:
				VARIANT v;
				teVariantChangeType(&v, pv, VT_BSTR);
				if (v.bstrVal) {
					if (nLen < 0) {
						nLen = (::SysStringLen(v.bstrVal) + 1);
					}
					nLen *= sizeof(WCHAR);
					if (nLen > m_nSize) {
						nLen = m_nSize;
					}
					::CopyMemory(pDest, v.bstrVal, nLen);
				} else {
					*(WCHAR *)pDest = NULL;
				}
				VariantClear(&v);
				break;
			case VT_LPSTR:
			case VT_USERDEFINED:
				teVariantChangeType(&v, pv, VT_BSTR);
				int nLenA;
				nLenA = 0;
				if (v.bstrVal) {
					UINT cp = vt != VT_USERDEFINED ? CP_ACP : CP_UTF8;
					nLenA = ::WideCharToMultiByte(cp, 0, (LPCWSTR)v.bstrVal, nLen, NULL, 0, NULL, NULL) + 1;
					if (nLenA > m_nSize) {
						nLenA = m_nSize;
					}
					::WideCharToMultiByte(cp, 0, (LPCWSTR)v.bstrVal, nLen, (LPSTR)pDest, nLenA - 1, NULL, NULL);
				}
				((LPSTR)pDest)[nLenA] = NULL;
				VariantClear(&v);
				break;
			case VT_FILETIME:
				DATE dt;
				teGetVariantTime(pv, &dt);
				teVariantTimeToFileTime(dt, (LPFILETIME)pDest);
				break;
			case VT_CY:
				GetPointFormVariant((LPPOINT)pDest, pv);
				break;
			case VT_CARRAY:
				char *pc;
				VARIANT vMem;
				VariantInit(&vMem);
				pc = GetpcFromVariant(pv, &vMem);
				if (pc) {
					::CopyMemory(pDest, pc, sizeof(RECT));
				} else {
					::ZeroMemory(pDest, sizeof(RECT));
				}
				VariantClear(&vMem);
				break;
			case VT_CLSID:
				if (pv->vt == VT_BSTR) {
					teCLSIDFromString(pv->bstrVal, (LPCLSID)pDest);
				}
				break;
			case VT_VARIANT:
				teVariantCopy((VARIANT *)pDest, pv);
				break;
/*			case VT_RECORD:
				LPITEMIDLIST *ppidl;
				ppidl = (LPITEMIDLIST *)pDest;
				teCoTaskMemFree(*ppidl);
				teGetIDListFromVariant(ppidl, pv);
				break;*/
		}//end_switch
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"CteMemory::Write";
#endif
	}
}

void CteMemory::SetPoint(int x, int y)
{
	POINT *ppt = (POINT *)m_pc;
	ppt->x = x;
	ppt->y = y;
}

//CteContextMenu

CteContextMenu::CteContextMenu(IUnknown *punk, IDataObject *pDataObj, IUnknown *punkSB)
{
	m_cRef = 1;
	m_pContextMenu = NULL;
	m_pDataObj = pDataObj;
	m_hDll = NULL;
	if (punk) {
		punk->QueryInterface(IID_PPV_ARGS(&m_pContextMenu));
		if (punkSB) {
			IUnknown_SetSite(punk, punkSB);
		}
	}
	::ZeroMemory(m_param, sizeof(m_param));
}

CteContextMenu::~CteContextMenu()
{
	IUnknown_SetSite(m_pContextMenu, NULL);
	SafeRelease(&m_pContextMenu);
	SafeRelease(&m_pDataObj);
	teFreeLibrary(m_hDll, 100);
}

STDMETHODIMP CteContextMenu::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteContextMenu, IDispatch),
//		QITABENT(CteContextMenu, IContextMenu),
		{ 0 },
	};
	if (IsEqualIID(riid, IID_IContextMenu) || IsEqualIID(riid, IID_IContextMenu2) || IsEqualIID(riid, IID_IContextMenu3)) {
		return m_pContextMenu->QueryInterface(riid, ppvObject);
	}
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteContextMenu::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteContextMenu::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}
	return m_cRef;
}

STDMETHODIMP CteContextMenu::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteContextMenu::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteContextMenu::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodCM, 0, NULL, *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteContextMenu::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		HRESULT hr = NULL;
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		switch(dispIdMember) {
		case TE_METHOD + 1://QueryContextMenu
			if (nArg >= 4) {
				if (!m_param[2].uintVal) {
					for (int i = 5; i--;) {
						m_param[i].uint_ptr = (UINT_PTR)GetPtrFromVariant(&pDispParams->rgvarg[nArg - i]);
					}
					teSetLong(pVarResult, m_pContextMenu->QueryContextMenu(m_param[0].hmenu, m_param[1].uintVal, m_param[2].uintVal, m_param[3].uintVal, m_param[4].uintVal));
				}
			}
			return S_OK;

		case TE_METHOD + 2://InvokeCommand
			if (nArg >= 7) {
				if (g_pOnFunc[TE_OnInvokeCommand]) {
					VARIANT vResult;
					VariantInit(&vResult);
					VARIANTARG *pv = GetNewVARIANT(nArg + 2);
					teSetObject(&pv[nArg + 1], this);
					for (int i = nArg; i >= 0; --i) {
						VariantCopy(&pv[i], &pDispParams->rgvarg[i]);
					}
					Invoke4(g_pOnFunc[TE_OnInvokeCommand], &vResult, nArg + 2, pv);
					if (GetIntFromVariantClear(&vResult) == S_OK) {
						teSetLong(pVarResult, S_OK);
						return S_OK;
					}
				}
				CMINVOKECOMMANDINFOEX cmi = { sizeof(cmi) };
				VARIANTARG *pv = GetNewVARIANT(3);
				WCHAR **ppwc = new WCHAR*[3];
				char **ppc = new char*[3];
				BOOL bExec = TRUE;
				for (int i = 0; i <= 2; ++i) {
					if (teVarIsNumber(&pDispParams->rgvarg[nArg - i - 2])) {
						UINT_PTR nVerb = (UINT_PTR)GetPtrFromVariant(&pDispParams->rgvarg[nArg - i - 2]);
						ppwc[i] = (WCHAR *)nVerb;
						ppc[i] = (char *)nVerb;
						if (nVerb > MAXWORD) {
							bExec = FALSE;
						}
					} else {
						teVariantChangeType(&pv[i], &pDispParams->rgvarg[nArg - i - 2], VT_BSTR);
						ppwc[i] = pv[i].bstrVal;
						if (i == 2) {
							if (!ppwc[i]) {
								IShellBrowser *pSB;
								if (GetFolderVew(&pSB)) {
									LPITEMIDLIST pidl;
									if (teGetIDListFromObject(pSB, &pidl)) {
										teGetDisplayNameFromIDList(&ppwc[i], pidl, SHGDN_FORPARSING);
										VariantClear(&pv[i]);
										pv[i].bstrVal = ppwc[i];
										pv[i].vt = VT_BSTR;
										teCoTaskMemFree(pidl);
									}
									pSB->Release();
								}
							}
							if (!teIsFileSystem(ppwc[i]) || tePathIsDirectory(ppwc[i], 0, TRUE) != S_OK) {
								ppwc[i] = NULL;
							}
						}
						if ((ULONG_PTR)ppwc[i] > MAXWORD) {
							ppc[i] = teWideCharToMultiByte(CP_ACP, ppwc[i], -1);
							if (!teStrSameIFree(teMultiByteToWideChar(CP_ACP, ppc[i], -1), ppwc[i])) {
								cmi.fMask = CMIC_MASK_UNICODE;
							}
						} else {
							ppc[i] = (char *)ppwc[i];
						}
					}
				}
				if (bExec) {
					cmi.fMask |= (DWORD)GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					cmi.hwnd  = (HWND)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 1]);
					cmi.lpVerbW = ppwc[0];
					cmi.lpVerb = ppc[0];
					cmi.lpParametersW = ppwc[1];
					cmi.lpParameters = ppc[1];
					cmi.lpDirectoryW = ppwc[2];
					cmi.lpDirectory = ppc[2];
					cmi.nShow = GetIntFromVariant(&pDispParams->rgvarg[nArg - 5]);
					cmi.dwHotKey = (DWORD)GetIntFromVariant(&pDispParams->rgvarg[nArg - 6]);
					cmi.hIcon =(HANDLE)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 7]);
					if (cmi.lpVerb) {//#248
						CHAR szNameA[MAX_PATH];
						szNameA[0] = NULL;
						if ((UINT_PTR)cmi.lpVerb <= MAXWORD) {
							m_pContextMenu->GetCommandString((UINT_PTR)cmi.lpVerb, GCS_VERBA, NULL, szNameA, MAX_PATH);
							if (lstrcmpiA((UINT_PTR)cmi.lpVerb <= MAXWORD ? szNameA : cmi.lpVerb , "mount") == 0) {
								g_dwTickMount = GetTickCount();
							}
						}
					}
					try {
						hr = m_pContextMenu->InvokeCommand((LPCMINVOKECOMMANDINFO)&cmi);
					} catch (...) {
						hr = E_UNEXPECTED;
					}
				}
				teSetLong(pVarResult, hr);
				for (int i = 3; i--;) {
					VariantClear(&pv[i]);
					if ((ULONG_PTR)ppc[i] >= MAXWORD) {
						teSysFreeString((BSTR *)&ppc[i]);
					}
				}
				delete [] pv;
			}
			return S_OK;

		case TE_METHOD + 3://Items
			teSetObjectRelease(pVarResult, new CteFolderItems(m_pDataObj, NULL));
			return S_OK;

		case TE_METHOD + 4://GetCommandString
			if (nArg >= 1) {
				WCHAR szName[MAX_PATH];
				szName[0] = NULL;
				UINT idCmd = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (idCmd <= MAXWORD) {
					m_pContextMenu->GetCommandString(idCmd,
						GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]),
						NULL, (LPSTR)szName, MAX_PATH);
				}
				teSetSZ(pVarResult, szName);
			} else {
				teSetBool(pVarResult, TRUE);
			}
			return S_OK;

		case TE_PROPERTY + 5://FolderView
			IShellBrowser *pSB;
			if (GetFolderVew(&pSB)) {
				teSetObjectRelease(pVarResult, pSB);
			}
			return S_OK;

		case TE_METHOD + 6://HandleMenuMsg
			LRESULT lResult;
			IContextMenu3 *pCM3;
			IContextMenu2 *pCM2;
			lResult = 0;
			if (nArg >= 2) {
				if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pCM3))) {
					pCM3->HandleMenuMsg2(
						(UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg]),
						(WPARAM)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 1]),
						(LPARAM)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 2]), &lResult
					);
					pCM3->Release();
				} else if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pCM2))) {
					pCM2->HandleMenuMsg(
						(UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg]),
						(WPARAM)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 1]),
						(LPARAM)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 2])
					);
					pCM2->Release();
				}
			}
			teSetPtr(pVarResult, lResult);
			return S_OK;

		case DISPID_VALUE:
			teSetObject(pVarResult, this);
		case DISPID_TE_UNDEFIND:
			return S_OK;
		default:
			if (dispIdMember >= TE_PROPERTY + 10 && dispIdMember <= TE_PROPERTY + 14) {
				teSetPtr(pVarResult, m_param[dispIdMember - TE_PROPERTY - 10].long_ptr);
				return S_OK;
			}
			break;
		}
	} catch (...) {
		return teException(pExcepInfo, "ContextMenu", methodCM, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

BOOL CteContextMenu::GetFolderVew(IShellBrowser **ppSB)
{
	BOOL bResult = FALSE;
	IServiceProvider *pSP;
	if SUCCEEDED(IUnknown_GetSite(m_pContextMenu, IID_PPV_ARGS(&pSP))) {
		bResult = SUCCEEDED(pSP->QueryService(SID_SShellBrowser, IID_PPV_ARGS(ppSB)));
		pSP->Release();
	}
	return bResult;
}

/*// IContextMenu
STDMETHODIMP CteContextMenu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
	return m_pContextMenu->QueryContextMenu(hmenu, indexMenu, idCmdFirst, idCmdLast, uFlags);
}

STDMETHODIMP CteContextMenu::GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
	return m_pContextMenu->GetCommandString(idCmd, uFlags, pwReserved, pszName, cchMax);
}

STDMETHODIMP CteContextMenu::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
	return m_pContextMenu->InvokeCommand(pici);
}
//*/

//CteDropTarget

CteDropTarget::CteDropTarget(IDropTarget *pDropTarget, FolderItem *pFolderItem)
{
	m_cRef = 1;
	m_pDropTarget = NULL;
	if (pDropTarget) {
		pDropTarget->QueryInterface(IID_PPV_ARGS(&m_pDropTarget));
	}
	m_pFolderItem = NULL;
	if (pFolderItem) {
		pFolderItem->QueryInterface(IID_PPV_ARGS(&m_pFolderItem));
	}
}

CteDropTarget::~CteDropTarget()
{
	SafeRelease(&m_pFolderItem);
	SafeRelease(&m_pDropTarget);
}

STDMETHODIMP CteDropTarget::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteDropTarget, IDispatch),
		{ 0 },
	};

	if (m_pFolderItem && (IsEqualIID(riid, IID_FolderItem) || IsEqualIID(riid, g_ClsIdFI))) {
		return m_pFolderItem->QueryInterface(riid, ppvObject);
	}
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteDropTarget::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteDropTarget::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}
	return m_cRef;
}

STDMETHODIMP CteDropTarget::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteDropTarget::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDropTarget::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodDT, 0, NULL, *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteDropTarget::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		HRESULT hr;
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		switch(dispIdMember) {
		case TE_METHOD + 1://DragEnter
		case TE_METHOD + 2://DragOver
		case TE_METHOD + 3://Drop
			hr = E_INVALIDARG;
			if (nArg >= 2) {
				BOOL bSingle = FALSE;
				IUnknown *pObj = this;
				IDataObject *pDataObj;
				if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
					try {
						DWORD grfKeyState = (DWORD)GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
						POINTL pt;
						GetPointFormVariant((POINT *)&pt, &pDispParams->rgvarg[nArg - 2]);
						DWORD dwEffect1 = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
						VARIANT v;
						VariantInit(&v);
						IDispatch *pEffect = NULL;
						if (nArg >= 3) {
							if (GetDispatch(&pDispParams->rgvarg[nArg - 3], &pEffect)) {
								teGetPropertyAt(pEffect, 0, &v);
								dwEffect1 = GetIntFromVariantClear(&v);
								if (nArg >= 4) {
									bSingle = GetBoolFromVariant(&pDispParams->rgvarg[nArg - 4]);
									if (nArg >= 5) {
										IUnknown *punk;
										if (FindUnknown(&pDispParams->rgvarg[nArg - 5], &punk)) {
											pObj = punk;
										}
									}
								}
							} else {
								dwEffect1 = GetIntFromVariantClear(&pDispParams->rgvarg[nArg - 3]);
							}
						}
						DWORD dwEffect = dwEffect1;
						CteFolderItems *pDragItems = new CteFolderItems(pDataObj, NULL);
						try {
							POINTL pt0 = {0, 0};
							if (!bSingle || dispIdMember == TE_METHOD + 1) {//DragEnter
								if (m_pDropTarget) {
									try {
										hr = m_pDropTarget->DragEnter(pDataObj, grfKeyState, pt, &dwEffect1);
									} catch (...) {
										hr = E_UNEXPECTED;
									}
								} else {
									dwEffect1 = 0;
								}
								if (m_pFolderItem) {
									DWORD dwEffect2 = dwEffect;
									if (DragSub(TE_OnDragEnter, pObj, pDragItems, &grfKeyState, pt0, &dwEffect2) == S_OK) {
										hr = S_OK;
										dwEffect1 = dwEffect2;
									}
								}
							} else {
								hr = S_OK;
							}
							if (hr == S_OK && dispIdMember >= TE_METHOD + 2) {
								if (!bSingle || dispIdMember == TE_METHOD + 2) {//DragOver
									hr = S_FALSE;
									if (m_pFolderItem) {
										dwEffect1 = dwEffect;
										hr = DragSub(TE_OnDragOver, pObj, pDragItems, &grfKeyState, pt0, &dwEffect1);
									}
									if ((hr != S_OK || dwEffect1 == dwEffect) && m_pDropTarget) {
										dwEffect1 = dwEffect;
										try {
											hr = m_pDropTarget->DragOver(grfKeyState, pt, &dwEffect1);
										} catch (...) {
											hr = E_UNEXPECTED;
										}
									}
								}
								if (hr == S_OK && dispIdMember >= TE_METHOD + 3) {
									if (!bSingle || dispIdMember == TE_METHOD + 3) {//Drop
										hr = S_FALSE;
										if (m_pFolderItem) {
											dwEffect1 = dwEffect;
											hr = DragSub(TE_OnDrop, pObj, pDragItems, &grfKeyState, pt0, &dwEffect1);
										}
										if (hr != S_OK) {
											dwEffect1 = dwEffect;
											if (m_pDropTarget) {
												try {
													hr = m_pDropTarget->Drop(pDataObj, grfKeyState, pt, &dwEffect1);
												} catch (...) {
													hr = E_UNEXPECTED;
												}
											}
										}
									}
								}
							}
						} catch (...) {
							g_nException = 0;
#ifdef _DEBUG
							g_strException = L"DropTarget";
#endif
						}
						pDragItems->Release();
						if (!bSingle) {
							if (m_pDropTarget) {
								m_pDropTarget->DragLeave();
							}
							if (dispIdMember >= TE_METHOD + 3) {
								DragLeaveSub(this, NULL);
							}
						}
						if (pEffect) {
							teSetLong(&v, dwEffect1);
							tePutPropertyAt(pEffect, 0, &v);
							pEffect->Release();
						}
					} catch (...) {
						g_nException = 0;
#ifdef _DEBUG
						g_strException = L"DropTarget2";
#endif
					}
					pDataObj->Release();
				}
			}
			teSetLong(pVarResult, hr);
			return S_OK;

		case TE_METHOD + 4://DragLeave
			if (m_pDropTarget) {
				hr = m_pDropTarget->DragLeave();
			}
			DragLeaveSub(this, NULL);
			teSetLong(pVarResult, hr);
			return S_OK;

		case TE_PROPERTY + 5://Type
			teSetLong(pVarResult, CTRL_DT);
			return S_OK;

		case TE_PROPERTY + 6://FolderItem
			teSetObject(pVarResult, m_pFolderItem);
			return S_OK;

		case DISPID_VALUE:
			teSetObject(pVarResult, this);
		case DISPID_TE_UNDEFIND:
			return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, "DropTarget", methodDT, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteTreeView

CteTreeView::CteTreeView()
{
	m_cRef = 1;
	m_param = NULL;
	m_bMain = TRUE;
	m_pDropTarget2 = NULL;
	m_psiFocus = NULL;
	VariantInit(&m_vData);
	m_bsRoot = ::SysAllocString(L"::{679F85CB-0220-4080-B29B-5540CC05AAB6}\n0");
#ifdef _W2000
	VariantInit(&m_vSelected);
#endif
}

CteTreeView::~CteTreeView()
{
	Close();
	teSysFreeString(&m_bsRoot);
	if (!m_pFV && m_param) {
		delete [] m_param;
		m_param = NULL;
	}
}

VOID CteTreeView::Init(CteShellBrowser *pFV, HWND hwnd)
{
	m_pFV = pFV;
	m_hwndParent = hwnd;
	if (m_pFV) {
		m_param = pFV->m_param;
	} else {
		m_param = new DWORD[SB_Count];
		for (int i = SB_Count; i--;) {
			m_param[i] = g_paramFV[i];
		}
	}
	m_hwnd = NULL;
	m_hwndTV = NULL;
	m_pNameSpaceTreeControl = NULL;
	m_bSetRoot = TRUE;

#ifdef _2000XP
	m_pShellNameSpace = NULL;
#endif
	m_dwCookie = 0;
	m_bEmpty = FALSE;
}

VOID CteTreeView::Close()
{
#ifdef _2000XP
	if (!g_bUpperVista) {
		RemoveWindowSubclass(m_hwnd, TETVProc, 1);
	}
#endif
	RemoveWindowSubclass(m_hwndTV, TETVProc2, 1);
	if (m_pDropTarget2) {
//		SetProp(m_hwndTV, L"OleDropTargetInterface", (HANDLE)m_pDropTarget2->m_pDropTarget);
		RevokeDragDrop(m_hwndTV);
		RegisterDragDrop(m_hwndTV, m_pDropTarget2->m_pDropTarget);
		SafeRelease(&m_pDropTarget2);
	}
	if (m_pNameSpaceTreeControl) {
		m_pNameSpaceTreeControl->RemoveAllRoots();
		m_pNameSpaceTreeControl->TreeUnadvise(m_dwCookie);
		m_dwCookie = 0;

		HWND hwnd;
		if (IUnknown_GetWindow(m_pNameSpaceTreeControl, &hwnd) == S_OK) {
			PostMessage(hwnd, WM_CLOSE, 0, 0);
		}
		SafeRelease(&m_pNameSpaceTreeControl);
	}
#ifdef _2000XP
	if (m_pShellNameSpace) {
		IOleObject *pOleObject;
		if SUCCEEDED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pOleObject))) {
			RECT rc;
			SetRectEmpty(&rc);
			pOleObject->SetClientSite(NULL);
			pOleObject->DoVerb(OLEIVERB_HIDE, NULL, NULL, 0, g_hwndMain, &rc);
			pOleObject->Close(OLECLOSE_NOSAVE);
			pOleObject->Release();
			DestroyWindow(m_hwnd);
		}
		SafeRelease(&m_pShellNameSpace);
	}
#endif
	VariantClear(&m_vData);
	SafeRelease(&m_psiFocus);
#ifdef _W2000
	VariantClear(&m_vSelected);
#endif
}

VOID CteTreeView::Create()
{
	if (_Emulate_XP_ SUCCEEDED(teCreateInstance(CLSID_NamespaceTreeControl, NULL, NULL, IID_PPV_ARGS(&m_pNameSpaceTreeControl)))) {
		RECT rc;
		SetRectEmpty(&rc);
		if SUCCEEDED(m_pNameSpaceTreeControl->Initialize(m_hwndParent, &rc, m_param[SB_TreeFlags])) {
			m_pNameSpaceTreeControl->TreeAdvise(static_cast<INameSpaceTreeControlEvents *>(this), &m_dwCookie);
			if (IUnknown_GetWindow(m_pNameSpaceTreeControl, &m_hwnd) == S_OK) {
				m_hwndTV = FindTreeWindow(m_hwnd);
				if (m_hwndTV) {
					teSetTreeTheme(m_hwndTV, SendMessage(m_hwndTV, TVM_GETBKCOLOR, 0, 0));
					SetWindowLongPtr(m_hwndTV, GWLP_USERDATA, (LONG_PTR)this);
					SetWindowSubclass(m_hwndTV, TETVProc2, 1, 0);
					if (!m_pDropTarget2) {
						m_pDropTarget2 = new CteDropTarget2(m_hwndTV, static_cast<IDispatch *>(this), FALSE);
						teRegisterDragDrop(m_hwndTV, m_pDropTarget2, &m_pDropTarget2->m_pDropTarget);
					}
					if (m_param[SB_TreeFlags] & NSTCS_NOEDITLABELS) {
						SetWindowLongPtr(m_hwndTV, GWL_STYLE, GetWindowLongPtr(m_hwndTV, GWL_STYLE) & ~TVS_EDITLABELS);
					}
					TreeView_SetTextColor(m_hwndTV, GetSysColor(COLOR_WINDOWTEXT));
				}
				BringWindowToTop(m_hwnd);
				ArrangeWindow();
			}
			INameSpaceTreeControl2 *pNSTC2;
			if SUCCEEDED(m_pNameSpaceTreeControl->QueryInterface(IID_PPV_ARGS(&pNSTC2))) {
				pNSTC2->SetControlStyle2(NSTCS2_ALLMASK, NSTCS2_DEFAULT);
				pNSTC2->Release();
			}
			DoFunc(TE_OnCreate, this, E_NOTIMPL);
			return;
		}
	}
#ifdef _2000XP
	m_hwnd = CreateWindowEx(
		WS_EX_CLIENTEDGE, WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | ES_READONLY,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
		m_hwndParent, (HMENU)0, hInst, NULL);
	if FAILED(teCreateInstance(CLSID_ShellShellNameSpace, NULL, NULL, IID_PPV_ARGS(&m_pShellNameSpace))) {
		if FAILED(teCreateInstance(CLSID_ShellNameSpace, NULL, NULL, IID_PPV_ARGS(&m_pShellNameSpace))) {
			return;
		}
	}
	IQuickActivate *pQuickActivate;
	if (FAILED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pQuickActivate)))) {
		return;
	}
	QACONTAINER qaContainer = { sizeof(QACONTAINER), static_cast<IOleClientSite *>(this), NULL, NULL, static_cast<IDispatch *>(this) };
	QACONTROL qaControl = { sizeof(QACONTROL) };
	pQuickActivate->QuickActivate(&qaContainer, &qaControl);
	pQuickActivate->Release();

	IPersistStreamInit  *pPersistStreamInit;
//	IPersistStorage     *pPersistStorage;
//	IPersistMemory      *pPersistMemory;
//	IPersistPropertyBag *pPersistPropertyBag;

	if (m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pPersistStreamInit)) == S_OK) {
		pPersistStreamInit->InitNew();
		pPersistStreamInit->Release();
	}
/*/// Not use in IShellNameSpace
	else if (m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pPersistStorage)) == S_OK) {
		pPersistStorage->InitNew(NULL);
		pPersistStorage->Release();
	} else if (m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pPersistMemory)) == S_OK) {
		pPersistMemory->InitNew();
		pPersistMemory->Release();
	} else if (m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pPersistPropertyBag)) == S_OK) {
		pPersistPropertyBag->InitNew();
		pPersistPropertyBag->Release();
	}
//*/
	else {
		return;
	}
/*/// Not use in IShellNameSpace
	IRunnableObject *pRunnableObject;
	if (qaControl.dwMiscStatus & OLEMISC_ALWAYSRUN) {
		m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pRunnableObject));
		if (pRunnableObject) {
			pRunnableObject->Run(NULL);
			pRunnableObject->Release();
		}
	}
//*/
	RECT rc;
	SetRectEmpty(&rc);

	IOleObject *pOleObject;
	if FAILED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pOleObject))) {
		return;
	}
	pOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, static_cast<IOleClientSite *>(this), 0, m_hwndParent, &rc);
	pOleObject->Release();

	HWND hwnd;
	if (IUnknown_GetWindow(m_pShellNameSpace, &hwnd) == S_OK) {
		m_pShellNameSpace->put_Flags(m_param[SB_TreeFlags] & NSTCS_FAVORITESMODE);
		m_hwndTV = FindTreeWindow(hwnd);
		if (m_hwndTV) {
			HWND hwndParent = GetParent(m_hwndTV);
			SetWindowLongPtr(hwndParent, GWLP_USERDATA, (LONG_PTR)this);
			SetWindowSubclass(hwndParent, TETVProc, 1, 0);
			SetWindowLongPtr(m_hwndTV, GWLP_USERDATA, (LONG_PTR)this);
			SetWindowSubclass(m_hwndTV, TETVProc2, 1, 0);
			if (!m_pDropTarget2) {
				m_pDropTarget2 = new CteDropTarget2(m_hwndTV, static_cast<IDispatch *>(this), FALSE);
				teRegisterDragDrop(m_hwndTV, m_pDropTarget2, &m_pDropTarget2->m_pDropTarget);
			}

			DWORD dwStyle = GetWindowLong(m_hwndTV, GWL_STYLE);
			struct
			{
				DWORD nstc;
				DWORD nsc;
			} style[] = {
				{ NSTCS_HASEXPANDOS, TVS_HASBUTTONS },
				{ NSTCS_HASLINES, TVS_HASLINES },
				{ NSTCS_SINGLECLICKEXPAND, TVS_TRACKSELECT },
				{ NSTCS_FULLROWSELECT, TVS_FULLROWSELECT },
				{ NSTCS_SPRINGEXPAND, TVS_SINGLEEXPAND },
				{ NSTCS_HORIZONTALSCROLL, TVS_NOHSCROLL },//Opposite
				//NSTCS_ROOTHASEXPANDO
				{ NSTCS_SHOWSELECTIONALWAYS, TVS_SHOWSELALWAYS },
				{ NSTCS_NOINFOTIP, TVS_INFOTIP },//Opposite
				{ NSTCS_EVENHEIGHT, TVS_NONEVENHEIGHT },//Opposite
				//NSTCS_NOREPLACEOPEN
				{ NSTCS_DISABLEDRAGDROP, TVS_DISABLEDRAGDROP },
				//NSTCS_NOORDERSTREAM
				{ NSTCS_BORDER, WS_BORDER },
				{ NSTCS_NOEDITLABELS, TVS_EDITLABELS },//Opposite
				//NSTCS_FADEINOUTEXPANDOS
				//NSTCS_EMPTYTEXT
				{ NSTCS_CHECKBOXES, TVS_CHECKBOXES },
				//NSTCS_NOINDENTCHECKS
				//NSTCS_ALLOWJUNCTIONS
				//NSTCS_SHOWTABSBUTTON
				//NSTCS_SHOWDELETEBUTTON
				//NSTCS_SHOWREFRESHBUTTON
			};
			for (int i = _countof(style); i--;) {
				if (m_param[SB_TreeFlags] & style[i].nstc) {
					dwStyle |= style[i].nsc;
				} else {
					dwStyle &= ~style[i].nsc;
				}
			}
			dwStyle ^= (TVS_NOHSCROLL | TVS_INFOTIP | TVS_NONEVENHEIGHT | TVS_EDITLABELS);//Opposite
			SetWindowLongPtr(m_hwndTV, GWL_STYLE, dwStyle & ~WS_BORDER);
			if (dwStyle & WS_BORDER) {
				teSetExStyleOr(m_hwnd, WS_EX_CLIENTEDGE);
			} else {
				teSetExStyleAnd(m_hwnd, ~WS_EX_CLIENTEDGE);
			}
		}
		BringWindowToTop(m_hwnd);
		ArrangeWindow();
	}
	DoFunc(TE_OnCreate, this, E_NOTIMPL);
#endif
}

STDMETHODIMP CteTreeView::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteTreeView, IDispatch),
		QITABENT(CteTreeView, INameSpaceTreeControlEvents),
		QITABENT(CteTreeView, INameSpaceTreeControlCustomDraw),
		QITABENT(CteTreeView, IShellItemFilter),
//		QITABENT(CteTreeView, IContextMenu),
#ifdef _2000XP
		&DIID_DShellNameSpaceEvents, OFFSETOFCLASS(IDispatch, CteTreeView),
		QITABENT(CteTreeView, IOleClientSite),
		QITABENT(CteTreeView, IOleWindow),
		QITABENT(CteTreeView, IOleInPlaceSite),
//		QITABENT(CteTreeView, IOleControlSite),
#endif
		{ 0 },
	};
/*	if (IsEqualIID(riid, IID_IServiceProvider)) {
		*ppvObject = static_cast<IServiceProvider *>(new CteServiceProvider(static_cast<IDispatch *>(this), static_cast<IDispatch *>(m_pFV)));
		return S_OK;
	}*/
	if (m_pNameSpaceTreeControl && IsEqualIID(riid, IID_INameSpaceTreeControl)) {
		return m_pNameSpaceTreeControl->QueryInterface(riid, ppvObject);
	}
#ifdef _2000XP
	if (m_pShellNameSpace) {
		if (IsEqualIID(riid, IID_IOleObject)) {
			return m_pShellNameSpace->QueryInterface(riid, ppvObject);
		}
		if (IsEqualIID(riid, IID_IShellNameSpace)) {
			return m_pShellNameSpace->QueryInterface(riid, ppvObject);
		}
	}
#endif
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteTreeView::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteTreeView::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteTreeView::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteTreeView::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodTV, _countof(methodTV), g_maps[MAP_TV], *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteTreeView::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		if (dispIdMember >= TE_PROPERTY + 0x1101 && dispIdMember < TE_METHOD_MAX) {
			if (m_param[SB_TreeAlign] & 2) {
				if (!m_pNameSpaceTreeControl
#ifdef _2000XP
					&& !m_pShellNameSpace
#endif
				) {
					Create();
				}
			} else if (dispIdMember != TE_PROPERTY + 0x2002 && dispIdMember != TE_METHOD + 0x2003) {
				if (dispIdMember < TE_OFFSET || dispIdMember > TE_OFFSET + 0xff) {
					return S_FALSE;
				}
			}
		}

		switch(dispIdMember) {
		case TE_PROPERTY + 0x1001://Data
			if (nArg >= 0) {
				VariantClear(&m_vData);
				VariantCopy(&m_vData, &pDispParams->rgvarg[nArg]);
			}
			teVariantCopy(pVarResult, &m_vData);
			return S_OK;

		case TE_PROPERTY + 0x1002://Type
			teSetLong(pVarResult, CTRL_TV);
			return S_OK;

		case TE_PROPERTY + 0x1003://hwnd
			teSetPtr(pVarResult, m_hwnd);
			return S_OK;

		case TE_METHOD + 0x1004://Close
			if (m_param) {
				m_param[SB_TreeAlign] = 0;
				ArrangeWindow();
				m_bEmpty = TRUE;
				Close();
			}
			return S_OK;

		case TE_PROPERTY + 0x1005://hwnd
			teSetPtr(pVarResult, m_hwndTV);
			return S_OK;

		case TE_PROPERTY + 0x1007://FolderView
			if (m_pFV) {
				teSetObject(pVarResult, m_pFV);
			} else if (g_pTC) {
				teSetObject(pVarResult, g_pTC->GetShellBrowser(g_pTC->m_nIndex));
			}
			return S_OK;

		case TE_PROPERTY + 0x1008://Align
			if (nArg >= 0) {
				m_param[SB_TreeAlign] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				ArrangeWindow();
				Show();
			}
			teSetLong(pVarResult, m_param[SB_TreeAlign]);
			return S_OK;

		case TE_PROPERTY + 0x1009://Visible
			if (nArg >= 0) {
				m_param[SB_TreeAlign] = GetBoolFromVariant(&pDispParams->rgvarg[nArg]) ? 3 : 1;
				ArrangeWindow();
				Show();
			}
			teSetBool(pVarResult, m_param[SB_TreeAlign] & 2);
			return S_OK;

		case TE_METHOD + 0x1106://Focus
			SetFocus(m_hwndTV);
			return S_OK;

		case TE_METHOD + 0x1107://HitTest (Screen coordinates)
			if (nArg >= 0 && pVarResult) {
				TVHITTESTINFO info;
				HTREEITEM hItem;
				GetPointFormVariant(&info.pt, &pDispParams->rgvarg[nArg]);
				UINT flags = nArg >= 1 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) :
					(m_param[SB_TreeFlags] & NSTCS_FULLROWSELECT) ? TVHT_ONITEM | TVHT_ONITEMRIGHT : TVHT_ONITEM;
				LONG_PTR r = DoHitTest(this, info.pt, flags);
				if (r != -1) {
					hItem = (HTREEITEM)r;
				} else {
					ScreenToClient(m_hwndTV, &info.pt);
					info.flags = flags;
					hItem = TreeView_HitTest(m_hwndTV, &info);
					if (!(info.flags & flags)) {
						hItem = 0;
					}
					if (nArg == 0 && hItem) {
						if (m_pNameSpaceTreeControl) {
							IShellItem *pSI;
							if (m_pNameSpaceTreeControl->HitTest(&info.pt, &pSI) == S_OK) {
								FolderItem *pFI;
								if SUCCEEDED(GetFolderItemFromObject(&pFI, pSI)) {
									teSetObjectRelease(pVarResult, pFI);
								}
								pSI->Release();
							}
							return S_OK;
						}
#ifdef _2000XP
						if (hItem == TreeView_GetSelection(m_hwndTV)) {
							IDispatch *pdisp;
							if (getSelected(&pdisp) == S_OK) {
								teSetObjectRelease(pVarResult, pdisp);
								return S_OK;
							}
						}
#endif
					}
				}
				teSetPtr(pVarResult, hItem);
			}
			return S_OK;

		case TE_METHOD + 0x1206://Refresh
			if (!m_bSetRoot) {
				m_bSetRoot = TRUE;
				SetTimer(m_hwndTV, TET_SetRoot, 100, teTimerProcForTree);
			}
			return S_OK;

		case TE_METHOD + 0x1283://GetItemRect
			if (nArg >= 1) {
				VARIANT vMem;
				VariantInit(&vMem);
				LPRECT prc = (LPRECT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1], &vMem);
				if (prc) {
					LPITEMIDLIST pidl;
					if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
						if (m_pNameSpaceTreeControl) {
							IShellItem *pShellItem;
							if SUCCEEDED(_SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&pShellItem))) {
								teSetLong(pVarResult, m_pNameSpaceTreeControl->GetItemRect(pShellItem, prc));
								pShellItem->Release();
							}
						}
						teCoTaskMemFree(pidl);
					}
				}
				teWriteBack(&pDispParams->rgvarg[nArg - 1], &vMem);
			}
			return S_OK;

		case TE_METHOD + 0x1300://Notify
			if (nArg >= 2 && m_hwndTV && m_pNameSpaceTreeControl) {
				long lEvent = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (lEvent & (SHCNE_MKDIR | SHCNE_MEDIAINSERTED | SHCNE_DRIVEADD | SHCNE_NETSHARE)) {
					LPITEMIDLIST pidl;
					if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg - 1])) {
						IShellItem *psi, *psiParent;
						if SUCCEEDED(_SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&psi))) {
							DWORD dwState;
							if FAILED(m_pNameSpaceTreeControl->GetItemState(psi, NSTCIS_EXPANDED, &dwState)) {
								if SUCCEEDED(psi->GetParent(&psiParent)) {
									if SUCCEEDED(m_pNameSpaceTreeControl->GetItemState(psiParent, NSTCIS_EXPANDED, &dwState)) {
										m_pNameSpaceTreeControl->SetItemState(psi, NSTCIS_EXPANDED, NSTCIS_EXPANDED);
									}
									psiParent->Release();
								}
							}
							psi->Release();
						}
						teCoTaskMemFree(pidl);
					}
				}
				if (nArg >= 4 && lEvent & (SHCNE_DRIVEREMOVED | SHCNE_MEDIAREMOVED | SHCNE_NETUNSHARE | SHCNE_RENAMEFOLDER | SHCNE_RMDIR | SHCNE_SERVERDISCONNECT | SHCNE_UPDATEDIR)) {
					try {
						SendMessage(GetParent(m_hwndTV), WM_USER + 1, GetPtrFromVariant(&pDispParams->rgvarg[nArg - 3]), GetPtrFromVariant(&pDispParams->rgvarg[nArg - 4]));
					} catch (...) {
						g_nException = 0;
#ifdef _DEBUG
						g_strException = L"Notify_Tree";
#endif
					}
				}
			}
			return S_OK;

		case TE_PROPERTY + 0x2000://SelectedItem
			IDispatch *pdisp;
			if (getSelected(&pdisp) == S_OK) {
				teSetObjectRelease(pVarResult, pdisp);
				return S_OK;
			}
			break;

		case TE_METHOD + 0x2001://SelectedItems
			if (getSelected(&pdisp) == S_OK) {
				CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL);
				VARIANT v;
				VariantInit(&v);
				teSetObjectRelease(&v, pdisp);
				pFolderItems->ItemEx(-1, NULL, &v);
				VariantClear(&v);
				teSetObjectRelease(pVarResult, pFolderItems);
				return S_OK;
			}
			break;

		case TE_PROPERTY + 0x2002://Root
			if (pVarResult) {
				teSetSZ(pVarResult, m_bsRoot);
				if (nArg < 0) {
					return S_OK;
				}
			}

		case TE_METHOD + 0x2003://SetRoot
			if (nArg >= 0) {
				m_param[SB_EnumFlags] = g_paramFV[SB_EnumFlags];
				m_param[SB_RootStyle] = g_paramFV[SB_RootStyle];
				if (nArg >= 1) {
					DWORD dwEnumFlags = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
					if (m_param[SB_EnumFlags] != dwEnumFlags) {
						m_param[SB_EnumFlags] = dwEnumFlags;
						m_bSetRoot = TRUE;
					}
					if (nArg >= 2) {
						DWORD dwRootStyle = GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]);
						if (m_param[SB_RootStyle] != dwRootStyle) {
							m_param[SB_RootStyle] = dwRootStyle;
							m_bSetRoot = TRUE;
						}
					}
				}
				SetRootV(&pDispParams->rgvarg[nArg]);
				return S_OK;
			}
			break;

		case TE_METHOD + 0x2004://Expand
			if (nArg >= 1) {
				FolderItem *pid;
				GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
				if (teILIsBlank(pid)) {
					pid->Release();
					return S_OK;
				}
				LPITEMIDLIST pidl;
				teGetIDListFromObject(pid, &pidl);
				pid->Release();
				if (teILIsSearchFolder(pidl)) {
					teCoTaskMemFree(pidl);
					return S_OK;
				}
				if (::ILIsEqual(pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
					::ILRemoveLastID(pidl);
				}
				if (!g_param[TE_ShowInternet]) {
					while (::ILIsEqual(pidl, g_pidls[CSIDL_INTERNET]) || ::ILIsParent(g_pidls[CSIDL_INTERNET], pidl, FALSE)) {
						::ILRemoveLastID(pidl);
					}
				}
				m_dwState = NSTCIS_SELECTED;
				if (GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) != 0) {
					m_dwState |= NSTCIS_EXPANDED;
				}
				SafeRelease(&m_psiFocus);
				_SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&m_psiFocus));
				teCoTaskMemFree(pidl);
				if (m_psiFocus) {
					m_bExpand = TRUE;
					SetTimer(m_hwndTV, TET_Expand, 100, teTimerProcForTree);
					return S_OK;
				}
			}
			break;
			//DIID_DShellNameSpaceEvents
#ifdef _W2000
		case DISPID_FAVSELECTIONCHANGE://(for Windows 2000)
									   //SelectedItem is empty (Windows 2000)
									   //Be sure to set a route in SetRoot (Windows 2000)
			if (nArg >= 3) {
				::VariantCopy(&m_vSelected, &pDispParams->rgvarg[nArg - 3]);
			}
			return E_NOTIMPL;
#endif
#ifdef _2000XP
		case DISPID_SELECTIONCHANGE:
			return DoFunc(TE_OnSelectionChanged, this, E_NOTIMPL);
		case DISPID_DOUBLECLICK:
			return S_OK;
		case DISPID_READYSTATECHANGE:
			return S_OK;
			/*			case DISPID_AMBIENT_BACKCOLOR:
			return S_OK;*/
			/*			case DISPID_INITIALIZED:
			return E_NOTIMPL;*/
		case DISPID_AMBIENT_LOCALEID:
			teSetLong(pVarResult, (SHORT)GetThreadLocale());
			return S_OK;
		case DISPID_AMBIENT_USERMODE:
			teSetBool(pVarResult, TRUE);
			return S_OK;
		case DISPID_AMBIENT_DISPLAYASDEFAULT:
			teSetBool(pVarResult, FALSE);
			return S_OK;
#endif
		case DISPID_VALUE:
			teSetObject(pVarResult, this);
		case DISPID_TE_UNDEFIND:
			return S_OK;
		default:
			if (dispIdMember >= TE_OFFSET && dispIdMember <= TE_OFFSET + SB_RootStyle) {
				if (nArg >= 0) {
					DWORD dw = GetIntFromVariantPP(&pDispParams->rgvarg[nArg], dispIdMember - TE_OFFSET);
					if (dw != m_param[dispIdMember - TE_OFFSET]) {
						m_param[dispIdMember - TE_OFFSET] = dw;
						g_paramFV[dispIdMember - TE_OFFSET] = dw;
						if (dispIdMember >= TE_OFFSET + SB_TreeFlags) {
							Close();
							m_bSetRoot = TRUE;
						}
						ArrangeWindow();
					}
				}
				teSetLong(pVarResult, m_param[dispIdMember - TE_OFFSET]);
				return S_OK;
			}
			break;
		}
		if ((dispIdMember >= TE_PROPERTY && dispIdMember < TE_PROPERTY + TVVERBS) ||
			(dispIdMember >= TE_METHOD && dispIdMember < TE_METHOD + TVVERBS)) {
#ifdef _2000XP
			if (m_pNameSpaceTreeControl) {
				return S_OK;
			}
			if (m_pShellNameSpace) {
				for (int i = _countof(methodTV); i--;) {
					if (methodTV[i].id == dispIdMember) {
						if (m_pShellNameSpace) {
							IDispatch *pDispatch;
							if SUCCEEDED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pDispatch))) {
								DISPID dispid;
								BSTR bs = teMultiByteToWideChar(CP_UTF8, methodTV[i].name, -1);
								if SUCCEEDED(pDispatch->GetIDsOfNames(IID_NULL, &bs, 1, LOCALE_USER_DEFAULT, &dispid)) {
									pDispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, wFlags, pDispParams, pVarResult, NULL, NULL);
								}
								::SysFreeString(bs);
								pDispatch->Release();
							}
						}
						break;
					}
				}
			}
#endif
			return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, "TreeView", methodTV, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//INameSpaceTreeControlEvents
STDMETHODIMP CteTreeView::OnItemClick(IShellItem *psi, NSTCEHITTEST nstceHitTest, NSTCSTYLE nsctsFlags)
{
	if (g_pOnFunc[TE_OnItemClick]) {
		VARIANTARG *pv = GetNewVARIANT(4);
		teSetObject(&pv[3], this);

		if (m_pNameSpaceTreeControl) {
			FolderItem *pid;
			if SUCCEEDED(GetFolderItemFromObject(&pid, psi)) {
				teSetObjectRelease(&pv[2], pid);
			}

		} else {
			IDispatch *pdisp;
			if (getSelected(&pdisp) == S_OK) {
				teSetObjectRelease(&pv[2], pdisp);
			}
		}
		teSetLong(&pv[1], nstceHitTest);
		teSetLong(&pv[0], nsctsFlags);
		VARIANT vResult;
		VariantInit(&vResult);
		Invoke4(g_pOnFunc[TE_OnItemClick], &vResult, 4, pv);
		if (vResult.vt != VT_EMPTY) {
			return GetIntFromVariantClear(&vResult);
		}
	}
	return S_FALSE;
}

STDMETHODIMP CteTreeView::OnPropertyItemCommit(IShellItem *psi)
{
	return S_FALSE;
}

STDMETHODIMP CteTreeView::OnItemStateChanging(IShellItem *psi, NSTCITEMSTATE nstcisMask, NSTCITEMSTATE nstcisState)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::OnItemStateChanged(IShellItem *psi, NSTCITEMSTATE nstcisMask, NSTCITEMSTATE nstcisState)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::OnSelectionChanged(IShellItemArray *psiaSelection)
{
	return DoFunc(TE_OnSelectionChanged, this, S_OK);
}

STDMETHODIMP CteTreeView::OnKeyboardInput(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	return S_FALSE;
}

STDMETHODIMP CteTreeView::OnBeforeExpand(IShellItem *psi)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::OnAfterExpand(IShellItem *psi)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::OnBeginLabelEdit(IShellItem *psi)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::OnEndLabelEdit(IShellItem *psi)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::OnGetToolTip(IShellItem *psi, LPWSTR pszTip, int cchTip)
{
	if (g_pOnFunc[TE_OnToolTip]) {
		VARIANTARG *pv = GetNewVARIANT(2);
		teSetObject(&pv[1], this);
		FolderItem *pid;
		if SUCCEEDED(GetFolderItemFromObject(&pid, psi)) {
			teSetObjectRelease(&pv[0], pid);
		}
		VARIANT vResult;
		VariantInit(&vResult);
		Invoke4(g_pOnFunc[TE_OnToolTip], &vResult, 2, pv);
		if (vResult.vt == VT_BSTR) {
			lstrcpyn(pszTip, vResult.bstrVal, cchTip);
			VariantClear(&vResult);
			return S_OK;
		}
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnBeforeItemDelete(IShellItem *psi)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnItemAdded(IShellItem *psi, BOOL fIsRoot)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnItemDeleted(IShellItem *psi, BOOL fIsRoot)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnBeforeContextMenu(IShellItem *psi, REFIID riid, void **ppv)
{
	*ppv = NULL;
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnAfterContextMenu(IShellItem *psi, IContextMenu *pcmIn, REFIID riid, void **ppv)
{
	MSG msg1;
	msg1.hwnd = m_hwnd;
	msg1.message = WM_CONTEXTMENU;
	msg1.wParam = 0;
	RECT rc;
	m_pNameSpaceTreeControl->GetItemRect(psi, &rc);
	GetCursorPos(&msg1.pt);
	if (!PtInRect(&rc, msg1.pt)) {
		msg1.pt.x = (rc.left + rc.right) / 2;
		msg1.pt.y = (rc.top + rc.bottom) / 2;
	}
	*ppv = NULL;
	if (g_pOnFunc[TE_OnShowContextMenu]) {
		VARIANT v;
		teSetObjectRelease(&v, new CteContextMenu(pcmIn, NULL, NULL));
		if (MessageSubPtV(TE_OnShowContextMenu, this, &msg1, &v) == S_OK) {
			return E_ABORT;
		}
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnBeforeStateImageChange(IShellItem *psi)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::OnGetDefaultIconIndex(IShellItem *psi, int *piDefaultIcon, int *piOpenIcon)
{
	return E_NOTIMPL;
}

//INameSpaceTreeControlCustomDraw
STDMETHODIMP CteTreeView::PrePaint(HDC hdc, RECT *prc, LRESULT *plres)
{
	*plres = CDRF_NOTIFYITEMDRAW;
	return S_OK;
}

STDMETHODIMP CteTreeView::PostPaint(HDC hdc, RECT *prc)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::ItemPrePaint(HDC hdc, RECT *prc, NSTCCUSTOMDRAW *pnstccdItem, COLORREF *pclrText, COLORREF *pclrTextBk, LRESULT *plres)
{
	NMTVCUSTOMDRAW tvcd = { { { 0, 0, NM_CUSTOMDRAW }, CDDS_ITEMPREPAINT, hdc, *prc, 0, pnstccdItem->uItemState }, *pclrText, *pclrTextBk, pnstccdItem->iLevel };
	teCustomDraw(TE_OnItemPrePaint, NULL, this, pnstccdItem->psi, &tvcd.nmcd, &tvcd, plres);
	if (!(pnstccdItem->uItemState & CDIS_SELECTED) || IsAppThemed()) {
		*pclrText = tvcd.clrText;
	}
	*pclrTextBk = tvcd.clrTextBk;
	return S_OK;
}

STDMETHODIMP CteTreeView::ItemPostPaint(HDC hdc, RECT *prc, NSTCCUSTOMDRAW *pnstccdItem)
{
	NMTVCUSTOMDRAW tvcd = { { { 0, 0, NM_CUSTOMDRAW }, CDDS_ITEMPOSTPAINT, hdc, *prc, 0, pnstccdItem->uItemState }, 0, 0, pnstccdItem->iLevel };
	LRESULT lres = CDRF_DODEFAULT;
	teCustomDraw(TE_OnItemPostPaint, NULL, this, pnstccdItem->psi, &tvcd.nmcd, &tvcd, &lres);
	return S_OK;
}

STDMETHODIMP CteTreeView::IncludeItem(IShellItem *psi)
{
	if (g_pOnFunc[TE_OnIncludeItem]) {
		VARIANT v;
		VariantInit(&v);
		VARIANTARG *pv = GetNewVARIANT(2);
		teSetObject(&pv[1], this);
		FolderItem *pid;
		if SUCCEEDED(GetFolderItemFromObject(&pid, psi)) {
			teSetObjectRelease(&pv[0], pid);
		}
		Invoke4(g_pOnFunc[TE_OnIncludeItem], &v, 2, pv);
		if (v.vt != VT_EMPTY) {
			return GetIntFromVariantClear(&v);
		}
	}
	return S_OK;
}

STDMETHODIMP CteTreeView::GetEnumFlagsForItem(IShellItem *psi, SHCONTF *pgrfFlags)
{
	return E_NOTIMPL;
}

//*/
#ifdef _2000XP
//IOleClientSite
STDMETHODIMP CteTreeView::SaveObject()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::GetContainer(IOleContainer **ppContainer)
{
	*ppContainer = NULL;
	return E_NOINTERFACE;
}

STDMETHODIMP CteTreeView::ShowObject()
{
	return S_OK;
}

STDMETHODIMP CteTreeView::OnShowWindow(BOOL fShow)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::RequestNewObjectLayout()
{
	return E_NOTIMPL;
}

//IOleInPlaceSite
STDMETHODIMP CteTreeView::GetWindow(HWND *phwnd)
{
	*phwnd = m_hwnd;
	return S_OK;
}

STDMETHODIMP CteTreeView::ContextSensitiveHelp(BOOL fEnterMode)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::CanInPlaceActivate()
{
	return S_OK;
}

STDMETHODIMP CteTreeView::OnInPlaceActivate()
{
	return S_OK;
}

STDMETHODIMP CteTreeView::OnUIActivate()
{
	return S_OK;
}

STDMETHODIMP CteTreeView::GetWindowContext(IOleInPlaceFrame **ppFrame, IOleInPlaceUIWindow **ppDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	*ppFrame = NULL;
	*ppDoc = NULL;
	GetClientRect(m_hwnd, lprcClipRect);
	CopyRect(lprcPosRect, lprcClipRect);
	return S_OK;
}

STDMETHODIMP CteTreeView::Scroll(SIZE scrollExtant)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnUIDeactivate(BOOL fUndoable)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnInPlaceDeactivate()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::DiscardUndoState()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::DeactivateAndUndo()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnPosRectChange(LPCRECT lprcPosRect)
{
	return S_OK;
}

/*//IOleControlSite
STDMETHODIMP CteTreeView::OnControlInfoChanged()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::LockInPlaceActive(BOOL fLock)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::GetExtendedControl(IDispatch **ppDisp)
{
	*ppDisp = NULL;
		return E_NOINTERFACE;
}

STDMETHODIMP CteTreeView::TransformCoords(POINTL *pPtlHimetric, POINTF *pPtfContainer, DWORD dwFlags)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::TranslateAccelerator(MSG *pMsg, DWORD grfModifiers)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnFocus(BOOL fGotFocus)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::ShowPropertyFrame()
{
	return E_NOTIMPL;
}
//*/
#endif

VOID CteTreeView::SetRootV(VARIANT *pv)
{
	VARIANT v;
	teVariantChangeType(&v, pv, VT_BSTR);
	if (lstrcmpi(v.bstrVal, m_bsRoot)) {
		Close();
		teSysFreeString(&m_bsRoot);
		m_bsRoot = ::SysAllocString(v.bstrVal);
		m_bSetRoot = TRUE;
	}
	VariantClear(&v);
	Show();
}

VOID CteTreeView::SetRoot()
{
	HRESULT hr = E_FAIL;
	if (m_pNameSpaceTreeControl) {
		m_pNameSpaceTreeControl->RemoveAllRoots();
		IShellItem *pShellItem;
		int nRoots = 0;
		BSTR bs = ::SysAllocString(m_bsRoot);
		LPWSTR pszPath = bs;
		for (int i = 0; i <= ::SysStringLen(bs); ++i) {
			if (bs[i] < 0x20) {
				bs[i] = 0;
				PathUnquoteSpaces(pszPath);
				if (lstrlen(pszPath)) {
					LPITEMIDLIST pidl = teILCreateFromPath(pszPath);
					pszPath = &bs[i + 1];
					if (pidl) {
						if SUCCEEDED(_SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&pShellItem))) {
							hr = m_pNameSpaceTreeControl->AppendRoot(pShellItem, m_param[SB_EnumFlags] | SHCONTF_ENABLE_ASYNC, m_param[SB_RootStyle], this);
							pShellItem->Release();
							if SUCCEEDED(hr) {
								++nRoots;
							}
						}
						teCoTaskMemFree(pidl);
					}
				}
			}
		}
		teSysFreeString(&bs);
		if (!nRoots) {
			if SUCCEEDED(_SHCreateItemFromIDList(g_pidls[CSIDL_DESKTOP], IID_PPV_ARGS(&pShellItem))) {
				hr = m_pNameSpaceTreeControl->AppendRoot(pShellItem, m_param[SB_EnumFlags], m_param[SB_RootStyle], this);
				pShellItem->Release();
			}
		}
	}
#ifdef _2000XP
	if (hr != S_OK && m_pShellNameSpace) {
		DWORD dwStyle = GetWindowLong(m_hwndTV, GWL_STYLE);
		if (m_param[SB_RootStyle] & NSTCRS_HIDDEN) {
			dwStyle &= ~TVS_LINESATROOT;
		} else {
			dwStyle |= TVS_LINESATROOT;
		}
		SetWindowLong(m_hwndTV, GWL_STYLE, dwStyle);

		//running to call DISPID_FAVSELECTIONCHANGE SetRoot(Windows 2000)
		m_pShellNameSpace->put_EnumOptions(m_param[SB_EnumFlags]);
		hr = m_pShellNameSpace->SetRoot(m_bsRoot);
		if (hr != S_OK) {
			VARIANT v, v0;
			v0.bstrVal = m_bsRoot;
			v0.vt = VT_I4;
			teVariantChangeType(&v, &v0, VT_I4);
			hr = m_pShellNameSpace->put_Root(v);
			m_pShellNameSpace->SetRoot(NULL);
		}
	}
#endif
	if (hr == S_OK) {
		m_bSetRoot = FALSE;
		DoFunc(TE_OnViewCreated, this, E_NOTIMPL);
	}
}

VOID CteTreeView::Show()
{
	if ((m_pFV && !m_pFV->m_bVisible) || !(m_param[SB_TreeAlign] & 2)) {
		return;
	}
	if (!m_pNameSpaceTreeControl
#ifdef _2000XP
		&& !m_pShellNameSpace
#endif
	) {
		Create();
	}
	if (m_bSetRoot) {
		SetTimer(m_hwndTV, TET_SetRoot, 100, teTimerProcForTree);
	}
}

VOID  CteTreeView::Expand()
{
	if (!m_psiFocus || m_bSetRoot) {
		return;
	}
	if (m_pNameSpaceTreeControl) {
		int iOrder = 1;
		if (m_bExpand) {
			IShellItemArray *psia;
			if SUCCEEDED(m_pNameSpaceTreeControl->GetSelectedItems(&psia)) {
				IShellItem *psi;
				if SUCCEEDED(psia->GetItemAt(0, &psi)) {
					psi->Compare(m_psiFocus, SICHINT_CANONICAL, &iOrder);
					psi->Release();
				}
				psia->Release();
			}
		} else {
			iOrder = 0;
		}
		if (iOrder) {
			m_pNameSpaceTreeControl->SetItemState(m_psiFocus, m_dwState, m_dwState);
			SetTimer(m_hwndTV, TET_Expand, 500, teTimerProcForTree);
		} else {
			m_pNameSpaceTreeControl->EnsureItemVisible(m_psiFocus);
			SafeRelease(&m_psiFocus);
		}
		m_bExpand = FALSE;
#ifdef _2000XP
	} else if (m_pShellNameSpace) {
		LPITEMIDLIST pidl = NULL;
		if (teGetIDListFromObject(m_psiFocus, &pidl)) {
			VARIANT v;
			VariantInit(&v);
			teSetIDList(&v, pidl);
			teILFreeClear(&pidl);
			m_pShellNameSpace->Expand(v, (m_dwState & NSTCIS_EXPANDED) ? 1 : 0);
			VariantClear(&v);
		}
		SafeRelease(&m_psiFocus);
#endif
	}
}

#ifdef _2000XP
VOID CteTreeView::SetObjectRect()
{
	if (m_pShellNameSpace) {
		teSetObjectRects(m_pShellNameSpace, m_hwnd);
	}
}
#endif

HRESULT CteTreeView::getSelected(IDispatch **ppid)
{
	*ppid = NULL;
	HRESULT hr = E_FAIL;
	if (m_pNameSpaceTreeControl) {
		IShellItemArray *psia;
		if SUCCEEDED(m_pNameSpaceTreeControl->GetSelectedItems(&psia)) {
			IShellItem *psi;
			if SUCCEEDED(psia->GetItemAt(0, &psi)) {
				FolderItem *pFI;
				if SUCCEEDED(GetFolderItemFromObject(&pFI, psi)) {
					hr = pFI->QueryInterface(IID_PPV_ARGS(ppid));
					pFI->Release();
				}
				psi->Release();
			}
			psia->Release();
		}
		return hr;
	}
#ifdef _2000XP
	if (m_pShellNameSpace) {
		hr = m_pShellNameSpace->get_SelectedItem(ppid);
#endif
#ifdef _W2000
		if (!*ppid && m_vSelected.vt != VT_EMPTY) {
			*ppid = new CteFolderItem(&m_vSelected);
			hr = S_OK;
		}
#endif
#ifdef _2000XP
	}
#endif
	if (!*ppid) {
		hr = E_FAIL;
	}
	return hr;
}

//IPersist
STDMETHODIMP CteTreeView::GetClassID(CLSID *pClassID)
{
	*pClassID = CLSID_ShellShellNameSpace;
	return S_OK;
}

//IPersistFolder
STDMETHODIMP CteTreeView::Initialize(PCIDLIST_ABSOLUTE pidl)
{
	return E_NOTIMPL;
}

//IPersistFolder2
STDMETHODIMP CteTreeView::GetCurFolder(PIDLIST_ABSOLUTE *ppidl)
{
	IDispatch *pdisp;
	HRESULT hr = getSelected(&pdisp);
	if SUCCEEDED(hr) {
		teGetIDListFromObject(pdisp, ppidl);
		pdisp->Release();
	}
	return S_OK;
}

// CteFolderItem

CteFolderItem::CteFolderItem(VARIANT *pv)
{
	m_cRef = 1;
	m_pidl = NULL;
	m_pidlAlt = NULL;
	m_pFolderItem = NULL;
	m_bStrict = FALSE;
	m_pidlFocused = NULL;
	m_pEnum = NULL;
	m_dwUnavailable = 0;
	m_dwSessionId = g_dwSessionId;
	VariantInit(&m_v);
	if (pv) {
		CteFolderItem *pid;
		if (pv->vt == VT_DISPATCH && SUCCEEDED(pv->pdispVal->QueryInterface(g_ClsIdFI, (LPVOID *)&pid))) {
			m_pidl = ILClone(pid->m_pidl);
			m_pidlAlt = ILClone(pid->m_pidlAlt);
			m_pidlFocused = ILClone(pid->m_pidlFocused);
			if (pid->m_pFolderItem) {
				pid->QueryInterface(IID_PPV_ARGS(&m_pFolderItem));
			}
			m_bStrict = pid->m_bStrict;
			if (pid->m_pEnum) {
				pid->QueryInterface(IID_PPV_ARGS(&m_pEnum));
			}
			m_dwUnavailable = pid->m_dwUnavailable;
			m_dwSessionId = pid->m_dwSessionId;
			VariantCopy(&m_v, &pid->m_v);
			pid->Release();
		} else {
			VariantCopy(&m_v, pv);
		}
	}
}

CteFolderItem::~CteFolderItem()
{
	teCoTaskMemFree(m_pidlAlt);
	teCoTaskMemFree(m_pidl);
	SafeRelease(&m_pFolderItem);
	SafeRelease(&m_pEnum);
	::VariantClear(&m_v);
	teCoTaskMemFree(m_pidlFocused);
}

LPITEMIDLIST CteFolderItem::GetPidl()
{
	if (m_v.vt != VT_EMPTY) {
		if (m_pidl == NULL) {
			if (m_v.vt == VT_BSTR) {
				if (tePathMatchSpec(m_v.bstrVal, L"*\\..")) {
					BSTR bs = ::SysAllocString(m_v.bstrVal);
					PathRemoveFileSpec(bs);
					LPITEMIDLIST pidl = teILCreateFromPathEx(bs);
					::SysFreeString(bs);
					if (pidl) {
						teILCloneReplace(&m_pidlFocused, ILFindLastID(pidl));
						teCoTaskMemFree(pidl);
					}
				}
			}
			if (!teGetIDListFromVariant(&m_pidl, &m_v)) {
				if (m_v.vt == VT_BSTR) {
					m_pidl = teSHSimpleIDListFromPathEx(m_v.bstrVal, FILE_ATTRIBUTE_HIDDEN, -1, -1, NULL);
				}
				if (!m_pidlAlt) {
					m_dwUnavailable = GetTickCount();
				}
			}
		}
	}
	return m_pidl ? m_pidl : m_pidlAlt ? m_pidlAlt : g_pidls[CSIDL_DESKTOP];
}

LPITEMIDLIST CteFolderItem::GetAlt()
{
	if (m_pEnum && m_dwSessionId != g_dwSessionId) {
		m_pEnum = NULL;
		Clear();
	}
	if (m_pidlAlt == NULL) {
		m_dwSessionId = g_dwSessionId;
		if (m_v.vt == VT_BSTR) {
			if (g_pOnFunc[TE_OnTranslatePath]) {
				try {
					if (InterlockedIncrement(&g_nProcFI) < 5) {
						VARIANTARG *pv = GetNewVARIANT(2);
						teSetObject(&pv[1], this);
						VariantCopy(&pv[0], &m_v);
						VARIANT vPath;
						::VariantInit(&vPath);
						Invoke4(g_pOnFunc[TE_OnTranslatePath], &vPath, 2, pv);
						if (vPath.vt != VT_EMPTY && vPath.vt != VT_NULL) {
							if (teGetIDListFromVariant(&m_pidlAlt, &vPath)) {
								m_dwUnavailable = 0;
							}
						}
						VariantClear(&vPath);
					}
				} catch(...) {
					g_nException = 0;
#ifdef _DEBUG
					g_strException = L"TranslatePath";
#endif
				}
				::InterlockedDecrement(&g_nProcFI);
			}
		}
	}
	if (m_pidlAlt == NULL) {
		m_pidlAlt = ILClone(m_pidl);
	}
	return m_pidlAlt;
}

BOOL CteFolderItem::GetFolderItem()
{
	if (m_pFolderItem) {
		return TRUE;
	}
	return GetFolderItemFromIDList2(&m_pFolderItem, GetPidl());
}

VOID CteFolderItem::Clear()
{
	teILFreeClear(&m_pidl);
	teILFreeClear(&m_pidlAlt);
	m_dwUnavailable = 0;
	m_dispExtendedProperty = -1;
}

BSTR CteFolderItem::GetStrPath()
{
	if (m_v.vt == VT_BSTR) {
		if (m_pidl == NULL || (m_pidlAlt && !ILIsEqual(m_pidl, m_pidlAlt)) || teIsSearchFolder(m_v.bstrVal)) {
			return m_v.bstrVal;
		}
	}
	return NULL;
}

VOID CteFolderItem::MakeUnavailable()
{
	teILCloneReplace(&m_pidlAlt, g_pidls[CSIDL_RESULTSFOLDER]);
	m_dwUnavailable = GetTickCount();
}

STDMETHODIMP CteFolderItem::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteFolderItem, IDispatch),
		QITABENT(CteFolderItem, FolderItem),
		QITABENT(CteFolderItem, IPersist),
		QITABENT(CteFolderItem, IPersistFolder),
		QITABENT(CteFolderItem, IPersistFolder2),
		QITABENT(CteFolderItem, IParentAndItem),
		{ 0 },
	};
	if (IsEqualIID(riid, g_ClsIdFI)) {
		*ppvObject = this;
		AddRef();
		return S_OK;
	}
	HRESULT hr = QISearch(this, qit, riid, ppvObject);
	if FAILED(hr) {
		if (IsEqualIID(riid, IID_FolderItem2) || IsEqualIID(riid, CLSID_ShellFolderItem)) {
			GetFolderItem();
		}
		if (m_pFolderItem) {
			hr = m_pFolderItem->QueryInterface(riid, ppvObject);
		}
	}
	return hr;
}

STDMETHODIMP_(ULONG) CteFolderItem::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteFolderItem::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteFolderItem::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteFolderItem::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

//IPersist
STDMETHODIMP CteFolderItem::GetClassID(CLSID *pClassID)
{
	*pClassID = CLSID_FolderItem;
	return S_OK;
}

//IPersistFolder
STDMETHODIMP CteFolderItem::Initialize(PCIDLIST_ABSOLUTE pidl)
{
	teILCloneReplace(&m_pidl, pidl);
	VariantClear(&m_v);
	teILFreeClear(&m_pidlAlt);
	m_dwUnavailable = 0;
	return S_OK;
}

//IPersistFolder2
STDMETHODIMP CteFolderItem::GetCurFolder(PIDLIST_ABSOLUTE *ppidl)
{
	*ppidl = ::ILClone(GetPidl());
	return S_OK;
}

//IParentAndItem
STDMETHODIMP CteFolderItem::GetParentAndItem(PIDLIST_ABSOLUTE *ppidlParent, IShellFolder **ppsf, PITEMID_CHILD *ppidlChild)
{
	HRESULT hr = E_NOTIMPL;
	LPITEMIDLIST pidl = GetPidl();
	if (pidl) {
		if (ppsf) {
			LPCITEMIDLIST pidlChild;
			SHBindToParent(const_cast<LPCITEMIDLIST>(pidl), IID_PPV_ARGS(ppsf), &pidlChild);
		}
		if (ppidlParent) {
			*ppidlParent = ILClone(pidl);
			ILRemoveLastID(*ppidlParent);
		}
		*ppidlChild = ILClone(ILFindLastID(pidl));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP CteFolderItem::SetParentAndItem(PCIDLIST_ABSOLUTE pidlParent, IShellFolder *psf, PCUITEMID_CHILD pidlChild)
{
	return E_NOTIMPL;
}

//IDispatch
STDMETHODIMP CteFolderItem::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	HRESULT hr = teGetDispId(methodFI, 0, NULL, *rgszNames, rgDispId, TRUE);
	if SUCCEEDED(hr) {
		return S_OK;
	}
	if (GetFolderItem()) {
		hr = m_pFolderItem->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
		if (SUCCEEDED(hr) && *rgDispId > 0) {
			if (lstrcmpi(*rgszNames, L"ExtendedProperty") == 0) {
				m_dispExtendedProperty = *rgDispId;
			}
			*rgDispId += 10;
		}
	}
#ifdef _DEBUG
	if FAILED(hr) {
		Sleep(0);
	}
#endif
	if (hr != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteFolderItem::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	int i;
	try {
		HRESULT hr = E_FAIL;
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		switch(dispIdMember) {
		case 1://Name
			if (nArg >= 0) {
				hr = put_Name(pDispParams->rgvarg[nArg].bstrVal);
			}
			if (pVarResult) {
				hr = get_Name(&pVarResult->bstrVal);
				if SUCCEEDED(hr) {
					pVarResult->vt = VT_BSTR;
				}
			}
			return hr;

		case 2://Path
			if (pVarResult) {
				if SUCCEEDED(get_Path(&pVarResult->bstrVal)) {
					pVarResult->vt = VT_BSTR;
				}
			}
			return S_OK;

		case 3://Alt
			if (nArg >= 0) {
				teILFreeClear(&m_pidlAlt);
				if (teGetIDListFromVariant(&m_pidlAlt, &pDispParams->rgvarg[nArg])) {
					m_dwUnavailable = 0;
				}
			}
			if (pVarResult) {
				teSetIDList(pVarResult, GetAlt());
			}
			return S_OK;

		case 5://Unavailable
			i = 0;
			if (nArg >= 0) {
				m_dwUnavailable = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (m_dwUnavailable) {
					m_dwUnavailable += GetTickCount();
				}
			}
			if (m_dwUnavailable) {
				i = GetTickCount() - m_dwUnavailable;
				if (i == 0) {
					i = 1;
				}
				if (i < 0) {
					i = MAXINT;
				}
			}
			teSetLong(pVarResult, i);
			return S_OK;

		case 6://Enum
			if (pVarResult || wFlags & DISPATCH_METHOD) {
				GetAlt();
			}
			teInvokeObject(&m_pEnum, wFlags, pVarResult, nArg, pDispParams->rgvarg);
			return S_OK;

		case 9://_BLOB ** To be necessary
			return S_OK;
			/*/
			case 8://FocusedItem //Reserved future
			if (nArg >= 0) {
			teILFreeClear(&m_pidlFocused);
			teGetIDListFromVariant(&m_pidlFocused, &pDispParams->rgvarg[nArg]);
			}
			if (pVarResult) {
			teSetIDList(pVarResult, m_pidlFocused);
			}
			return S_OK;
			//*/
		case DISPID_VALUE:
			teSetObject(pVarResult, this);
		case DISPID_TE_UNDEFIND:
			return S_OK;
		}
		if (GetFolderItem()) {
			if (dispIdMember > 10) {
				dispIdMember -= 10;
			}
			hr = m_pFolderItem->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
			if (dispIdMember == m_dispExtendedProperty && pVarResult && nArg >= 0 && !pVarResult->bstrVal && pDispParams->rgvarg[nArg].vt == VT_BSTR) {
				PROPERTYKEY propKey;
				if SUCCEEDED(_PSPropertyKeyFromStringEx(pDispParams->rgvarg[nArg].bstrVal, &propKey)) {
					if (IsEqualPropertyKey(propKey, PKEY_Link_TargetParsingPath)) {
						BSTR bs;
						if SUCCEEDED(get_Path(&bs)) {
							teGetJunctionLinkTarget(bs, &pVarResult->bstrVal, -1);
							if (pVarResult->bstrVal) {
								pVarResult->vt = VT_BSTR;
							}
							teSysFreeString(&bs);
						}
					}
				}
			}
			return hr != DISP_E_EXCEPTION ? hr : S_FALSE;
		}
	} catch (...) {
		return teException(pExcepInfo, "FolderItem", methodFI, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

STDMETHODIMP CteFolderItem::get_Application(IDispatch **ppid)
{
	return GetFolderItem() ? m_pFolderItem->get_Application(ppid) : E_FAIL;
}

STDMETHODIMP CteFolderItem::get_Parent(IDispatch **ppid)
{
	return GetFolderItem() ? m_pFolderItem->get_Parent(ppid) : E_FAIL;
}

STDMETHODIMP CteFolderItem::get_Name(BSTR *pbs)
{
	if SUCCEEDED(tePathGetFileName(pbs, GetStrPath())) {
		return S_OK;
	}
	return teGetDisplayNameFromIDList(pbs, GetPidl(), SHGDN_INFOLDER);
}

STDMETHODIMP CteFolderItem::put_Name(BSTR bs)
{
	HRESULT hr = E_FAIL;
	if (g_pOnFunc[TE_OnSetName]) {
		VARIANT v;
		VariantInit(&v);
		VARIANTARG *pv = GetNewVARIANT(2);
		teSetObject(&pv[1], this);
		teSetSZ(&pv[0], bs);
		Invoke4(g_pOnFunc[TE_OnSetName], &v, 2, pv);
		if (v.vt != VT_EMPTY) {
			return GetIntFromVariantClear(&v);
		}
	}
	if (GetPidl() && m_pidl) {
		IShellFolder *pSF;
		LPCITEMIDLIST pidlPart;
		if SUCCEEDED(SHBindToParent(m_pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
			BSTR bs2 = bs;
			if (bs[0] == '.' && !StrChr(&bs[1], '.')) {
				int i = ::SysStringLen(bs);
				bs2 = ::SysAllocStringLen(NULL, i);
				lstrcpy(bs2, bs);
				bs2[i] = '.';
				bs2[i + 1] = NULL;
			}
			hr = pSF->SetNameOf(g_hwndMain, pidlPart, bs2, SHGDN_NORMAL, NULL);
			if (bs != bs2) {
				::SysFreeString(bs2);
			}
			pSF->Release();
		}
	}
	return hr;
}

STDMETHODIMP CteFolderItem::get_Path(BSTR *pbs)
{
	BSTR bs = GetStrPath();
	if (bs) {
		*pbs = ::SysAllocString(bs);
		return S_OK;
	}
	return teGetDisplayNameFromIDList(pbs, GetPidl(), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
}

STDMETHODIMP CteFolderItem::get_GetLink(IDispatch **ppid)
{
	return GetFolderItem() ? m_pFolderItem->get_GetLink(ppid) : E_FAIL;
}

STDMETHODIMP CteFolderItem::get_GetFolder(IDispatch **ppid)
{
	return GetFolderObjFromIDList(GetPidl(), (Folder **)ppid);
}

STDMETHODIMP CteFolderItem::get_IsLink(VARIANT_BOOL *pb)
{
	return GetFolderItem() ? m_pFolderItem->get_IsLink(pb) : E_FAIL;
}

STDMETHODIMP CteFolderItem::get_IsFolder(VARIANT_BOOL *pb)
{
	return GetFolderItem() ? m_pFolderItem->get_IsFolder(pb) : E_FAIL;
}

STDMETHODIMP CteFolderItem::get_IsFileSystem(VARIANT_BOOL *pb)
{
	return GetFolderItem() ? m_pFolderItem->get_IsFileSystem(pb) : E_FAIL;
}

STDMETHODIMP CteFolderItem::get_IsBrowsable(VARIANT_BOOL *pb)
{
	return GetFolderItem() ? m_pFolderItem->get_IsBrowsable(pb) : E_FAIL;
}

STDMETHODIMP CteFolderItem::get_ModifyDate(DATE *pdt)
{
	return GetFolderItem() ? m_pFolderItem->get_ModifyDate(pdt) : E_FAIL;
}

STDMETHODIMP CteFolderItem::put_ModifyDate(DATE dt)
{
	return GetFolderItem() ? m_pFolderItem->put_ModifyDate(dt) : E_FAIL;
}

STDMETHODIMP CteFolderItem::get_Size(LONG *pul)
{
	return GetFolderItem() ? m_pFolderItem->get_Size(pul) : E_FAIL;
}

STDMETHODIMP CteFolderItem::get_Type(BSTR *pbs)
{
	return GetFolderItem() ? m_pFolderItem->get_Type(pbs) : E_FAIL;
}

STDMETHODIMP CteFolderItem::Verbs(FolderItemVerbs **ppfic)
{
	return GetFolderItem() ? m_pFolderItem->Verbs(ppfic) : E_FAIL;
}

STDMETHODIMP CteFolderItem::InvokeVerb(VARIANT vVerb)
{
	return GetFolderItem() ? m_pFolderItem->InvokeVerb(vVerb) : E_FAIL;
}

//CteCommonDialog

CteCommonDialog::CteCommonDialog()
{
	m_cRef = 1;

	::ZeroMemory(&m_ofn, sizeof(OPENFILENAME));
	m_ofn.lStructSize = sizeof(OPENFILENAME);
	m_ofn.hwndOwner = GetForegroundWindow();
//	m_ofn.hInstance = hInst;
	m_ofn.nMaxFile = MAX_PATH;
}

CteCommonDialog::~CteCommonDialog()
{

	teSysFreeString(&m_ofn.lpstrFile);
	teSysFreeString(const_cast<BSTR *>(&m_ofn.lpstrInitialDir));
	teSysFreeString(const_cast<BSTR *>(&m_ofn.lpstrFilter));
	teSysFreeString(const_cast<BSTR *>(&m_ofn.lpstrDefExt));
	teSysFreeString(const_cast<BSTR *>(&m_ofn.lpstrTitle));
}

STDMETHODIMP CteCommonDialog::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteCommonDialog, IDispatch),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteCommonDialog::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteCommonDialog::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteCommonDialog::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteCommonDialog::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteCommonDialog::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodCD, _countof(methodCD), g_maps[MAP_CD], *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteCommonDialog::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		LPCWSTR *pbs = NULL;
		DWORD *pdw = NULL;
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		switch(dispIdMember) {
			case TE_PROPERTY + 10://FileName
				if (nArg >= 0) {
					VARIANT vFile;
					teVariantChangeType(&vFile, &pDispParams->rgvarg[nArg], VT_BSTR);
					if (!m_ofn.lpstrFile) {
						m_ofn.lpstrFile = teSysAllocStringLenEx(vFile.bstrVal, m_ofn.nMaxFile);
					} else {
						lstrcpyn(m_ofn.lpstrFile, vFile.bstrVal, m_ofn.nMaxFile);
					}
					VariantClear(&vFile);
				}
				teSetSZ(pVarResult, m_ofn.lpstrFile);
				return S_OK;

			case TE_PROPERTY + 13://Filter
				if (nArg >= 0) {
					teSysFreeString(const_cast<BSTR *>(&m_ofn.lpstrFilter));
					if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
						teSysFreeString(const_cast<BSTR *>(&m_ofn.lpstrFilter));
						int i = ::SysStringLen(pDispParams->rgvarg[nArg].bstrVal);
						BSTR bs = teSysAllocStringLenEx(pDispParams->rgvarg[nArg].bstrVal, i + 1);
						while (i >= 0) {
							if (bs[i] == '|') {
								bs[i] = NULL;
							}
							--i;
						}
						m_ofn.lpstrFilter = bs;
					}
				}
				teSetSZ(pVarResult, m_ofn.lpstrInitialDir);
				return S_OK;
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
			case DISPID_TE_UNDEFIND:
				return S_OK;
		}
		if (dispIdMember >= TE_METHOD + 40) {
			if (!m_ofn.lpstrFile) {
				m_ofn.lpstrFile = ::SysAllocStringLen(NULL, m_ofn.nMaxFile);
				m_ofn.lpstrFile[0] = NULL;
			}
			BOOL bResult = FALSE;
			switch (dispIdMember) {
				//ShowOpen
				case TE_METHOD + 40:
					if (m_ofn.Flags & OFN_ENABLEHOOK) {
						m_ofn.lpfnHook = OFNHookProc;
					}
					g_bDialogOk = FALSE;
					bResult = GetOpenFileName(&m_ofn) || g_bDialogOk;
					break;
				//ShowSave
				case TE_METHOD + 41:
					bResult = GetSaveFileName(&m_ofn);
					break;
/*/// IFileOpenDialog
				case 42:
					{
						IFileOpenDialog *pFileOpenDialog;
						if (_SHCreateItemFromIDList && SUCCEEDED(teCreateInstance(CLSID_FileOpenDialog, NULL, NULL, IID_PPV_ARGS(&pFileOpenDialog)))) {
							IShellItem *psi;
							LPITEMIDLIST pidl = teILCreateFromPath(const_cast<LPWSTR>(m_ofn.lpstrInitialDir));
							if (pidl) {
								if SUCCEEDED(_SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&psi))) {
									pFileOpenDialog->SetFolder(psi);
									psi->Release();
								}
								teCoTaskMemFree(pidl);
							}
							pFileOpenDialog->SetFileName(m_ofn.lpstrFile);
							DWORD dwOptions;
							pFileOpenDialog->GetOptions(&dwOptions);
							pFileOpenDialog->SetOptions(dwOptions | FOS_PICKFOLDERS);
							if SUCCEEDED(pFileOpenDialog->Show(g_hwndMain)) {
								if SUCCEEDED(pFileOpenDialog->GetResult(&psi)) {
									if (GetIDListFromObject(psi, &pidl)) {
										BSTR bs;
										teGetDisplayNameFromIDList(&bs, pidl, SHGDN_FORPARSING);
										teCoTaskMemFree(pidl);
										lstrcpyn(m_ofn.lpstrFile, bs, m_ofn.nMaxFile);
										bResult = TRUE;
									}
									psi->Release();
								}
							}
							pFileOpenDialog->Release();
							break;
						}
#ifdef _2000XP
						m_ofn.Flags |= OFN_ENABLEHOOK;
						m_ofn.lpfnHook = OFNHookProc;
						g_bDialogOk = FALSE;
						bResult = GetOpenFileName(&m_ofn) || g_bDialogOk;
						break;
						BROWSEINFO bi;
						bi.hwndOwner = g_hwndMain;
						bi.pidlRoot = NULL;
						bi.pszDisplayName =	m_ofn.lpstrFile;
						bi.lpszTitle = m_ofn.lpstrFileTitle;
						bi.ulFlags = BIF_EDITBOX;
						bi.lpfn = NULL;
						LPITEMIDLIST pidl =	SHBrowseForFolder(&bi);
						if (pidl) {
							BSTR bs;
							teGetDisplayNameFromIDList(&bs, pidl, SHGDN_FORPARSING);
							teCoTaskMemFree(pidl);
							lstrcpyn(m_ofn.lpstrFile, bs, m_ofn.nMaxFile);
							::SysFreeString(bs);
							bResult = TRUE;
						}
#endif
					}
					break;
//*/
			}//end_switch
			teSetBool(pVarResult, bResult);
			return S_OK;
		}
		if (dispIdMember >= TE_PROPERTY + 30) {
			switch (dispIdMember) {
				case TE_PROPERTY + 30:
					pdw = &m_ofn.nMaxFile;
					break;
				case TE_PROPERTY + 31:
					pdw = &m_ofn.Flags;
					break;
				case TE_PROPERTY + 32:
					pdw = &m_ofn.nFilterIndex;
					break;
				case TE_PROPERTY + 33:
					pdw = &m_ofn.FlagsEx;
					break;
			}//end_switch
			if (nArg >= 0) {
				*pdw = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
			}
			teSetLong(pVarResult, *pdw);
			return S_OK;
		}
		if (dispIdMember >= TE_PROPERTY + 20) {
			switch (dispIdMember) {
				case TE_PROPERTY + 20:
					pbs = &m_ofn.lpstrInitialDir;
					break;
				case TE_PROPERTY + 21:
					pbs = &m_ofn.lpstrDefExt;
					break;
				case TE_PROPERTY + 22:
					pbs = &m_ofn.lpstrTitle;
					break;
			}//end_switch
			if (nArg >= 0) {
				if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
					SysReAllocString(const_cast<BSTR *>(pbs), pDispParams->rgvarg[nArg].bstrVal);
				}
			}
			teSetSZ(pVarResult, *pbs);
			return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, "CommonDialog", methodCD, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteWICBitmap

CteWICBitmap::CteWICBitmap()
{
	m_cRef = 1;
	m_uFrame = 0;
	m_uFrameCount = 1;
	m_guidSrc = GUID_NULL;
	m_pImage = NULL;
	m_pStream = NULL;
	m_ppMetadataQueryReader[0] = NULL;
	m_ppMetadataQueryReader[1] = NULL;
	if FAILED(teCreateInstance(CLSID_WICImagingFactory, NULL, NULL, IID_PPV_ARGS(&m_pWICFactory))) {
		m_pWICFactory = NULL;
	}
	for (int i = Count_WICFunc; i--;) {
		m_ppDispatch[i] = NULL;
	}
	if (g_dwMainThreadId == GetCurrentThreadId()) {
		for (int i = Count_WICFunc; i--;) {
			if (g_pOnFunc[TE_OnGetAlt + i]) {
				g_pOnFunc[TE_OnGetAlt + i]->QueryInterface(IID_PPV_ARGS(&m_ppDispatch[i]));
			}
		}
	}
}

CteWICBitmap::~CteWICBitmap()
{
	ClearImage(TRUE);
	for (int i = Count_WICFunc; i--;) {
		SafeRelease(&m_ppDispatch[i]);
	}
	SafeRelease(&m_pWICFactory);
}

VOID CteWICBitmap::ClearImage(BOOL bAll)
{
	if (bAll) {
		m_uFrame = 0;
		m_uFrameCount = 1;
		SafeRelease(&m_pStream);
		SafeRelease(&m_ppMetadataQueryReader[0]);
		SafeRelease(&m_ppMetadataQueryReader[1]);
	}
	SafeRelease(&m_pImage);
}

STDMETHODIMP CteWICBitmap::QueryInterface(REFIID riid, void **ppvObject)
{
	if (m_pImage && IsEqualIID(riid, IID_IWICBitmap)) {
		return m_pImage->QueryInterface(riid, ppvObject);
	}
	if (m_pStream && IsEqualIID(riid, IID_IStream)) {
		return m_pStream->QueryInterface(riid, ppvObject);
	}
	static const QITAB qit[] =
	{
		QITABENT(CteWICBitmap, IDispatch),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteWICBitmap::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteWICBitmap::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}
	return m_cRef;
}

STDMETHODIMP CteWICBitmap::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteWICBitmap::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWICBitmap::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodGB, _countof(methodGB), g_maps[MAP_GB], *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

HBITMAP CteWICBitmap::GetHBITMAP(COLORREF clBk)
{
	WICPixelFormatGUID guidPF;
	if (clBk == -4) {
#ifdef _WIN64
		clBk = -2;
#else
		clBk = g_bUpperVista ? -2 : GetSysColor(COLOR_MENU);
#endif
	}
	m_pImage->GetPixelFormat(&guidPF);
	BOOL bAlphaBlend = (int)clBk >= 0 && !IsEqualGUID(guidPF, GUID_WICPixelFormat24bppBGR);
	Get((int)clBk != -2 ? GUID_WICPixelFormat32bppBGRA : GUID_WICPixelFormat32bppPBGRA);

	HBITMAP hBM = NULL;
	UINT w = 0, h = 0;
	if (SUCCEEDED(m_pImage->GetSize(&w, &h)) && w && h) {
		BITMAPINFO bmi;
		UINT cbSize = 0;
		RGBQUAD *pcl = NULL;
		::ZeroMemory(&bmi.bmiHeader, sizeof(BITMAPINFOHEADER));
		bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bmi.bmiHeader.biWidth = w;
		bmi.bmiHeader.biHeight = -(LONG)h;
		bmi.bmiHeader.biPlanes = 1;
		bmi.bmiHeader.biBitCount = 32;
		hBM = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, (void **)&pcl, NULL, 0);
		if (hBM && pcl) {
			if SUCCEEDED(m_pImage->CopyPixels(NULL, w * 4, w * h * 4, (LPBYTE)pcl)) {
				if (bAlphaBlend) {
					int r0 = GetRValue(clBk);
					int g0 = GetGValue(clBk);
					int b0 = GetBValue(clBk);
					for (int i = w * h; --i >= 0; ++pcl) {
						int a = pcl->rgbReserved;
						pcl->rgbBlue = (pcl->rgbBlue - b0) * a / 0xff + b0;
						pcl->rgbGreen = (pcl->rgbGreen - g0) * a / 0xff + g0;
						pcl->rgbRed = (pcl->rgbRed - r0) * a / 0xff + r0;
						pcl->rgbReserved = 0xff;
					}
				}
			}
		}
	}
	return hBM;
}

BOOL CteWICBitmap::Get(WICPixelFormatGUID guidNewPF)
{
	WICPixelFormatGUID guidPF;
	m_pImage->GetPixelFormat(&guidPF);
	if (IsEqualGUID(guidPF, guidNewPF)) {
		return TRUE;
	}
	BOOL b = FALSE;
	IWICFormatConverter *pConverter;
	if SUCCEEDED(m_pWICFactory->CreateFormatConverter(&pConverter)) {
		if SUCCEEDED(pConverter->Initialize(m_pImage, guidNewPF, WICBitmapDitherTypeNone, NULL, 0.0f, WICBitmapPaletteTypeMedianCut)) {
			pConverter->GetPixelFormat(&guidPF);
			if (IsEqualGUID(guidPF, guidNewPF)) {
				IWICBitmap *pIBitmap;
				b = SUCCEEDED(m_pWICFactory->CreateBitmapFromSource(pConverter, WICBitmapCacheOnDemand, &pIBitmap));
				if (b) {
					ClearImage(FALSE);
					m_pImage = pIBitmap;
				}
			}
			SafeRelease(&pConverter);
		}
	}
	return b;
}

BOOL CteWICBitmap::GetEncoderClsid(LPWSTR pszName, CLSID* pClsid, LPWSTR pszMimeType)
{
	IEnumUnknown *pEnum;
	WCHAR pszType[MAX_PATH];
	WCHAR pszExt[MAX_PATH];
	LPWSTR pszExt0 = PathFindExtension(pszName);
	HRESULT hr = m_pWICFactory->CreateComponentEnumerator(WICEncoder, WICComponentEnumerateDefault, &pEnum);
	if SUCCEEDED(hr) {
		ULONG cbActual = 0;
		IUnknown *punk;
		hr = E_FAIL;
		while (pEnum->Next(1, &punk, &cbActual) == S_OK && hr != S_OK) {
			IWICBitmapCodecInfo *pCodecInfo;
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCodecInfo))) {
				UINT uActual = 0;
				if (pszMimeType) {
					pszType[0] = NULL;
					pCodecInfo->GetMimeTypes(MAX_PATH, pszType, &uActual);
					if (m_pStream && lstrlen(pszName) == 0) {
						if SUCCEEDED(pCodecInfo->GetContainerFormat(pClsid)) {
							if (IsEqualGUID(*pClsid, m_guidSrc)) {
								hr = S_OK;
							}
						}
					}
					if (hr != S_OK && teFindType(pszType, pszName)) {
						hr = pCodecInfo->GetContainerFormat(pClsid);
					}
				}
				if (hr != S_OK) {
					pszExt[0] = NULL;
					pCodecInfo->GetFileExtensions(MAX_PATH, pszExt, &uActual);
					if (teFindType(pszExt, pszExt0)) {
						hr = pCodecInfo->GetContainerFormat(pClsid);
					}
				}
				pCodecInfo->Release();
			}
			punk->Release();
		}
		pEnum->Release();
	}
	if SUCCEEDED(hr) {
		if (pszMimeType) {
			LPWSTR pszExt2 = StrChr(pszType, ',');
			int i = (int)(pszExt2 - pszType + 1);
			if (i <= 0 || i > 15) {
				i = 15;
			}
			lstrcpyn(pszMimeType, pszType, i);
		}
		return TRUE;
	}
	return FALSE;
}

// nAlpha 0:Use alpha  1:Use premultiplied alpha  2:Ignore alpha  3:Unknown
HRESULT CteWICBitmap::CreateBitmapFromHBITMAP(HBITMAP hBM, HPALETTE hPalette, int nAlpha)
{
	if (nAlpha == 3) {
		nAlpha = 2;
		BITMAP bm;
		if (GetObject(hBM, sizeof(BITMAP), &bm) && bm.bmBitsPixel == 32) {
			RGBQUAD *pcl = (RGBQUAD *)bm.bmBits;
			if (pcl) {
				try {
					for (int i = bm.bmWidth * bm.bmHeight; i--; ++pcl) {
						if (pcl->rgbReserved) {
							nAlpha = 0;
							break;
						}
					}
				} catch (...) {}
			}
		}
	}
	IWICBitmap *pIBitmap;
	HRESULT hr = m_pWICFactory->CreateBitmapFromHBITMAP(hBM, hPalette, (WICBitmapAlphaChannelOption)nAlpha, &pIBitmap);
	if SUCCEEDED(hr) {
		ClearImage(FALSE);
		m_pImage = pIBitmap;
	}
	return hr;
}

HRESULT CteWICBitmap::CreateStream(IStream *pStream, CLSID encoderClsid, LONG lQuality)
{
	HRESULT hr = E_FAIL;
	if (pStream) {
		IWICBitmapEncoder *pEncoder;
		if SUCCEEDED(m_pWICFactory->CreateEncoder(encoderClsid, NULL, &pEncoder)) {
			if (lQuality == 0 && m_pStream && IsEqualGUID(encoderClsid, m_guidSrc)) {
				teCopyStream(m_pStream, pStream);
				hr = S_OK;
			} else {
				UINT w = 0, h = 0;
				m_pImage->GetSize(&w, &h);
				WICPixelFormatGUID guidPF;
				m_pImage->GetPixelFormat(&guidPF);
				pEncoder->Initialize(pStream, WICBitmapEncoderNoCache);
				IWICBitmapFrameEncode *pFrameEncode;
				IPropertyBag2 *pPropBag;
				hr = pEncoder->CreateNewFrame(&pFrameEncode, &pPropBag);
				if SUCCEEDED(hr) {
					if (lQuality) {
						PROPBAG2 option = { 0 };
						option.pstrName = L"ImageQuality";
						VARIANT v;
						v.vt = VT_R4;
						v.fltVal = FLOAT(lQuality) / 100;
						pPropBag->Write(1, &option, &v);
					}
					hr = pFrameEncode->Initialize(pPropBag);
					if SUCCEEDED(hr) {
						pFrameEncode->SetSize(w, h);
						pFrameEncode->SetPixelFormat(&guidPF);
						IWICPalette *pPalette = NULL;
						if SUCCEEDED(m_pWICFactory->CreatePalette(&pPalette)) {
							if SUCCEEDED(m_pImage->CopyPalette(pPalette)) {
								pFrameEncode->SetPalette(pPalette);
							}
							SafeRelease(&pPalette);
						}
						hr = pFrameEncode->WriteSource(m_pImage, NULL);
						if SUCCEEDED(hr) {
							hr = pFrameEncode->Commit();
						}
						hr = pEncoder->Commit();
					}
					SafeRelease(&pPropBag);
					SafeRelease(&pFrameEncode);
				}
			}
			SafeRelease(&pEncoder);
		}
	}
	return hr;
}
/*
HRESULT CteWICBitmap::CreateBMPStream(IStream *pStream, LPWSTR szMime)
{
	HRESULT hr = E_FAIL;
	if (pStream) {
		HBITMAP hBM = GetHBITMAP(GetSysColor(COLOR_BTNFACE));
		if (hBM) {
			BITMAPFILEHEADER bf;
			BITMAPINFO bi;
			DIBSECTION ds;
			GetObject(hBM, sizeof(DIBSECTION), &ds);
			CopyMemory(&bi.bmiHeader, &ds.dsBmih, sizeof(BITMAPINFOHEADER));
			bf.bfType = 0x4d42;
			bf.bfReserved1 = 0;
			bf.bfReserved2 = 0;
			DWORD cb = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);
			bf.bfOffBits = cb;
			bf.bfSize = cb + bi.bmiHeader.biSizeImage;
			bf.bfType = 0x4d42;
			bf.bfReserved1 = 0;
			bf.bfReserved2 = 0;

			LPVOID pBits = HeapAlloc(GetProcessHeap(), 0, bi.bmiHeader.biSizeImage);
			if (pBits) {
				HDC hdc = GetDC(NULL);
				if (hdc) {
					if (GetDIBits(hdc, hBM, 0, ds.dsBm.bmHeight, pBits, (BITMAPINFO*)&bi, DIB_RGB_COLORS)) {
						LARGE_INTEGER liOffset;
						liOffset.QuadPart = 0;
						pStream->Seek(liOffset, STREAM_SEEK_SET, NULL);
						hr = pStream->Write(&bf, sizeof(BITMAPFILEHEADER), NULL);
						if SUCCEEDED(hr) {
							hr = pStream->Write(&bi, sizeof(BITMAPINFOHEADER), NULL);
						}
						if SUCCEEDED(hr) {
							hr = pStream->Write(pBits, bi.bmiHeader.biSizeImage, NULL);
						}
						if SUCCEEDED(hr) {
							lstrcpy(szMime, L"image/bmp");
							pStream->Seek(liOffset, STREAM_SEEK_SET, NULL);
						}
					}
					ReleaseDC(NULL, hdc);
				}
				HeapFree(GetProcessHeap(), 0, pBits);
			}
			DeleteObject(hBM);
		}
	}
	return hr;
}
*/

BOOL CteWICBitmap::HasImage()
{
	if (m_pImage) {
		UINT w = 0, h = 0;
		m_pImage->GetSize(&w, &h);
		return w != 0;
	}
	return FALSE;
}

CteWICBitmap* CteWICBitmap::GetBitmapObj()
{
	return HasImage() ? this : NULL;
}

VOID CteWICBitmap::GetFrameFromStream(IStream *pStream, UINT uFrame, BOOL bInit)
{
	m_uFrame = uFrame;
	LARGE_INTEGER liOffset;
	liOffset.QuadPart = 0;
	pStream->Seek(liOffset, SEEK_SET, NULL);
	IWICBitmapDecoder *pDecoder;
	if SUCCEEDED(m_pWICFactory->CreateDecoderFromStream(pStream, NULL, WICDecodeMetadataCacheOnDemand, &pDecoder)) {
		SafeRelease(&m_ppMetadataQueryReader[0]);
		pDecoder->GetMetadataQueryReader(&m_ppMetadataQueryReader[0]);
		IWICBitmapFrameDecode *pFrameDecode;
		pDecoder->GetFrameCount(&m_uFrameCount);
		if (m_uFrameCount) {
			if SUCCEEDED(pDecoder->GetFrame(uFrame, &pFrameDecode)) {
				IWICBitmap *pIBitmap;
				if SUCCEEDED(m_pWICFactory->CreateBitmapFromSource(pFrameDecode, WICBitmapCacheOnDemand, &pIBitmap)) {
					ClearImage(FALSE);
					m_pImage = pIBitmap;
				}
				SafeRelease(&m_ppMetadataQueryReader[1]);
				pFrameDecode->GetMetadataQueryReader(&m_ppMetadataQueryReader[1]);
				SafeRelease(&pFrameDecode);
			}
			if (m_uFrameCount > 1 && bInit) {
				if (m_pStream = SHCreateMemStream(NULL, NULL)) {
					teCopyStream(pStream, m_pStream);
					pDecoder->GetContainerFormat(&m_guidSrc);
				}
			}
		}
		SafeRelease(&pDecoder);
	}
}

VOID CteWICBitmap::FromStreamRelease(IStream *pStream, LPWSTR lpfn, BOOL bExtend, int cx)
{
	SafeRelease(&m_pStream);
	m_uFrameCount = 1;
	LARGE_INTEGER liOffset;
	liOffset.QuadPart = 0;
	GetFrameFromStream(pStream, 0, TRUE);
	if (!HasImage()) {
		HRESULT hr = E_FAIL;
		try {
			for (size_t i = 0; i < g_ppGetImage.size(); ++i) {
				pStream->Seek(liOffset, STREAM_SEEK_SET, NULL);
				HBITMAP hBM = NULL;
				int nAlpha = 3;
				LPFNGetImage _GetImage = (LPFNGetImage)g_ppGetImage[i];
				hr = _GetImage(pStream, lpfn, cx, &hBM, &nAlpha);
				if (hr == S_OK) {
					CreateBitmapFromHBITMAP(hBM, 0, nAlpha);
					::DeleteObject(hBM);
					break;
				}
			}
		} catch(...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"GetImage";
#endif
		}
	}
	SafeRelease(&pStream);
}

HRESULT CteWICBitmap::GetArchive(LPWSTR lpfn, int cx)
{
	HRESULT hr = E_FAIL;
	BSTR bsArcPath = NULL;
	BSTR bsItem = NULL;
	if (teGetArchiveSpec(lpfn, &bsArcPath, &bsItem)) {
		try {
			for (size_t i = 0; i < g_ppGetArchive.size(); ++i) {
				IStream *pStreamOut = NULL;
				LPFNGetArchive lpfnGetArchive = (LPFNGetArchive)g_ppGetArchive[i];
				hr = lpfnGetArchive(bsArcPath, bsItem, &pStreamOut, NULL);
				if (hr == S_OK) {
					FromStreamRelease(pStreamOut, bsItem, NULL, cx);
					break;
				}
			}
		}
		catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"GetArchive";
#endif
		}
	}
	teSysFreeString(&bsArcPath);
	teSysFreeString(&bsItem);
	return hr;
}

STDMETHODIMP CteWICBitmap::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		HRESULT hr = E_FAIL;
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (!m_pWICFactory) {
			return S_OK;
		}
		DISPID dispid = dispIdMember & TE_METHOD_MASK;
		if (dispid < 100) {
			ClearImage(TRUE);
		} else if (!m_pImage && dispid < 900) {
			return S_OK;
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		switch(dispIdMember) {
			//FromHBITMAP
			case TE_METHOD + 1:
			//FromStream
			case TE_METHOD + 5:
			//FromSource
			case TE_METHOD + 9:
				if (nArg >= 0) {
					IUnknown *punk;
					if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
						IStream *pStream;
						if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pStream))) {
							LPWSTR lpfn = NULL;
							int cx = 0;
							if (dispIdMember == TE_METHOD + 5 && nArg >= 1) {
								lpfn = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]);
								cx = (nArg >= 2) ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]) : 0;
							}
							FromStreamRelease(pStream, lpfn, TRUE, cx);
							teSetObject(pVarResult, GetBitmapObj());
							return S_OK;
						}
						IWICBitmap *pIBitmap;
						if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pIBitmap))) {
							int nOpt = (nArg >= 1) ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : 2;
							m_pWICFactory->CreateBitmapFromSource(pIBitmap, (WICBitmapCreateCacheOption)nOpt, &m_pImage);
							pIBitmap->Release();
							teSetObject(pVarResult, GetBitmapObj());
							return S_OK;
						}
					}
					HPALETTE pal = (nArg >= 1) ? (HPALETTE)GetPtrFromVariant(&pDispParams->rgvarg[nArg - 1]) : 0;
					int nAlpha = (nArg >= 2) ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]) : 3;
					CreateBitmapFromHBITMAP((HBITMAP)GetPtrFromVariant(&pDispParams->rgvarg[nArg]), pal, nAlpha);
					teSetObject(pVarResult, GetBitmapObj());
				}
				return S_OK;
			//FromHICON
			case TE_METHOD + 2:
				if (nArg >= 0) {
					m_pWICFactory->CreateBitmapFromHICON((HICON)GetPtrFromVariant(&pDispParams->rgvarg[nArg]), &m_pImage);
					teSetObject(pVarResult, GetBitmapObj());
				}
				return S_OK;
			//FromResource
			case TE_METHOD + 3:
				if (nArg >= 5) {
					HBITMAP hBM = (HBITMAP)LoadImage((HINSTANCE)GetPtrFromVariant(&pDispParams->rgvarg[nArg]), GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 3]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 4]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 5]));
					if (hBM) {
						CreateBitmapFromHBITMAP(hBM, 0, 3);
						::DeleteObject(hBM);
					}
					teSetObject(pVarResult, GetBitmapObj());
				}
				return S_OK;
			//FromFile
			case TE_METHOD + 4:
				if (nArg >= 0) {
					int cx = (nArg >= 1) ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : 0;
					LPWSTR lpfn = NULL;
					LPWSTR bsfn = NULL;
					LPITEMIDLIST pidl = NULL;
					if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
						lpfn = pDispParams->rgvarg[nArg].bstrVal;
					} else if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
						if SUCCEEDED(teGetDisplayNameFromIDList(&bsfn, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
							lpfn = bsfn;
						}
					}
					if (teStartsText(L"data:", lpfn)) {
						LPWSTR lpBase64 = StrChr(lpfn, ',');
						if (lpBase64) {
							DWORD dwData;
							DWORD dwLen = lstrlen(++lpBase64);
							if (CryptStringToBinary(lpBase64, dwLen, CRYPT_STRING_BASE64, NULL, &dwData, NULL, NULL) && dwData > 0) {
								BSTR bs = ::SysAllocStringByteLen(NULL, dwData);
								CryptStringToBinary(lpBase64, dwLen, CRYPT_STRING_BASE64, (BYTE *)bs, &dwData, NULL, NULL);
								IStream *pStream = SHCreateMemStream((BYTE *)bs, dwData);
								::SysFreeString(bs);
								FromStreamRelease(pStream, lpfn, TRUE, cx);
							}
						}
					}
					if (!HasImage()) {
						IStream *pStream = NULL;
						if (pidl) {
							IShellFolder *pSF;
							LPCITEMIDLIST pidlPart;
							if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
								pSF->BindToStorage(pidlPart, NULL, IID_PPV_ARGS(&pStream));
								pSF->Release();
							}
						}
						if (!pStream) {
							if FAILED(SHCreateStreamOnFileEx(lpfn, STGM_READ | STGM_SHARE_DENY_NONE, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &pStream)) {
								pStream = SHCreateMemStream(NULL, NULL);
							}
						}
						FromStreamRelease(pStream, lpfn, FALSE, cx);
						if (!HasImage()) {
							VARIANT vAlt;
							vAlt.bstrVal = NULL;
							if (m_ppDispatch[WIC_OnGetAlt]) {
								WCHAR pszPath[MAX_PATH + 10];
								GetTempPath(MAX_PATH, pszPath);
								PathAppend(pszPath, L"tablacus\\");
								int nLen = lstrlen(pszPath);
								if (StrCmpNI(lpfn, pszPath, nLen) == 0) {
									DWORD dwSessionId = 0;
									swscanf_s(&lpfn[nLen], L"%lx", &dwSessionId);
									if (dwSessionId) {
										VARIANTARG *pv = GetNewVARIANT(2);
										teSetLong(&pv[1], dwSessionId);
										teSetSZ(&pv[0], &lpfn[nLen]);
										Invoke4(m_ppDispatch[WIC_OnGetAlt], &vAlt, 2, pv);
										if (vAlt.vt != VT_BSTR) {
											VariantClear(&vAlt);
											vAlt.bstrVal = NULL;
										}
									}
								}
							}
							if (g_ppGetArchive.size()) {
								hr = GetArchive(lpfn, cx);
								if (hr != S_OK && vAlt.bstrVal) {
									hr = GetArchive(vAlt.bstrVal, cx);
								}
							}
							VariantClear(&vAlt);
						}
					}
					teSysFreeString(&bsfn);
					teILFreeClear(&pidl);
					teSetObject(pVarResult, GetBitmapObj());
				}
				return S_OK;
			//FromArchive
			case TE_METHOD + 6:
				if (nArg >= 3) {
					IStorage *pStorage = NULL;
					HMODULE hDll;
					HRESULT hr = teInitStorage(&pDispParams->rgvarg[nArg], &pDispParams->rgvarg[nArg - 1], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 2]), &hDll, &pStorage);
					if SUCCEEDED(hr) {
						VARIANT v;
						teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 3], VT_BSTR);
						LPWSTR lpPath;
						LPWSTR lpPath1 = v.bstrVal;
						while (lpPath = StrChr(lpPath1, '\\')) {
							lpPath[0] = NULL;
							IStorage *pStorageNew;
							pStorage->OpenStorage(lpPath1, NULL, STGM_READ, 0, 0, &pStorageNew);
							if (pStorageNew) {
								pStorage->Release();
								pStorage = pStorageNew;
							} else {
								break;
							}
							lpPath1 = lpPath + 1;
						}
						IStream *pStream;
						hr = pStorage->OpenStream(lpPath1, NULL, STGM_READ, NULL, &pStream);
						if SUCCEEDED(hr) {
							int cx = (nArg >= 4) ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 4]) : 0;
							FromStreamRelease(pStream, lpPath1, TRUE, cx);
						}
						VariantClear(&v);
					}
					SafeRelease(&pStorage);
					teFreeLibrary(hDll, 0);
					teSetObject(pVarResult, GetBitmapObj());
				}
				return S_OK;
			//FromItem//Deprecated
			case TE_METHOD + 7:
				return S_OK;
			//FromClipboard
			case TE_METHOD + 8:
				if (::IsClipboardFormatAvailable(CF_BITMAP)) {
					::OpenClipboard(NULL);
					HBITMAP hBM = (HBITMAP)::GetClipboardData(CF_BITMAP);
					CreateBitmapFromHBITMAP(hBM, 0, nArg >= 0 ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : 2);
					::CloseClipboard();
				}
				teSetObject(pVarResult, GetBitmapObj());
				return S_OK;
			//Create
			case TE_METHOD + 90:
				if (nArg >= 1) {
					IWICBitmap *pIBitmap;
					HRESULT hr = m_pWICFactory->CreateBitmap(GetIntFromVariant(&pDispParams->rgvarg[nArg]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]),
						GUID_WICPixelFormat32bppBGRA, WICBitmapCacheOnLoad, &pIBitmap);
					if SUCCEEDED(hr) {
						ClearImage(FALSE);
						m_pImage = pIBitmap;
					}
					teSetObject(pVarResult, GetBitmapObj());
				}
				return S_OK;
			//Free
			case TE_METHOD + 99:
				return S_OK;
			//Save
			case TE_METHOD + 100:
				if (nArg >= 0) {
					VARIANT vText;
					teVariantChangeType(&vText, &pDispParams->rgvarg[nArg], VT_BSTR);
					CLSID encoderClsid;
					if (GetEncoderClsid(vText.bstrVal, &encoderClsid, NULL)) {
						LONG lQuality = nArg ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : 0;
						IStream *pStream;
						hr = SHCreateStreamOnFileEx(vText.bstrVal, STGM_WRITE | STGM_CREATE | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, TRUE, NULL, &pStream);
						if SUCCEEDED(hr) {
							hr = CreateStream(pStream, encoderClsid, lQuality);
							SafeRelease(&pStream);
						}
					}
					VariantClear(&vText);
				}
				teSetLong(pVarResult, hr);
				return S_OK;
			//Base64
			case TE_METHOD + 101:
			//DataURI
			case TE_METHOD + 102:
			//GetStream
			case TE_METHOD + 103:
				VARIANT vText;
				vText.bstrVal = NULL;
				vText.vt = VT_BSTR;
				if (nArg >= 0) {
					teVariantChangeType(&vText, &pDispParams->rgvarg[nArg], VT_BSTR);
				}
				CLSID encoderClsid;
				if (lstrlen(vText.bstrVal) == 0 && !m_pStream) {
					teSetSZ(&vText, L"image/png");
				}
				IStream *pStream;
				pStream = SHCreateMemStream(NULL, NULL);
				WCHAR szMime[16];
				if (GetEncoderClsid(vText.bstrVal, &encoderClsid, szMime)) {
					hr = CreateStream(pStream, encoderClsid, 0);
//					hr = CreateBMPStream(pStream, szMime);
				}
				if SUCCEEDED(hr) {
					if (dispIdMember == TE_METHOD + 103) { //GetStream
						teSetObjectRelease(pVarResult, new CteObject(pStream));
					} else {
						DWORD dwSize = teGetStreamSize(pStream);
						BSTR bsBuff = SysAllocStringByteLen(NULL, dwSize);
						if (bsBuff) {
							ULONG ulBytesRead;
							if SUCCEEDED(pStream->Read(bsBuff, dwSize, &ulBytesRead)) {
								wchar_t szHead[32];
								szHead[0] = NULL;
								if (dispIdMember == TE_METHOD + 102) {
									swprintf_s(szHead, 32, L"data:%s;base64,", szMime);
								}
								int nDest = lstrlen(szHead);
								DWORD dwSize;
								CryptBinaryToString((BYTE *)bsBuff, ulBytesRead, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, NULL, &dwSize);
								if (dwSize > 0) {
									pVarResult->vt = VT_BSTR;
									pVarResult->bstrVal = ::SysAllocStringLen(NULL, nDest + dwSize - 1);
									lstrcpy(pVarResult->bstrVal, szHead);
									CryptBinaryToString((BYTE *)bsBuff, ulBytesRead, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, &pVarResult->bstrVal[nDest], &dwSize);
								}
							}
							::SysFreeString(bsBuff);
						}
					}
					SafeRelease(&pStream);
				}
				VariantClear(&vText);
				return S_OK;
			//GetWidth
			case TE_METHOD + 110:
				UINT w, h;
				w = 0;
				m_pImage->GetSize(&w, &h);
				teSetLong(pVarResult, w);
				return S_OK;
			//GetHeight
			case TE_METHOD + 111:
				h = 0;
				m_pImage->GetSize(&w, &h);
				teSetLong(pVarResult, h);
				return S_OK;
			//GetPixel
			case TE_METHOD + 112:
				if (nArg >= 1 && Get(GUID_WICPixelFormat32bppBGRA)) {
					UINT w = 0, h = 0;
					m_pImage->GetSize(&w, &h);
					IWICBitmapLock *pBitmapLock;
					WICRect rc = {0, 0, (INT)w, (INT)h};
					if SUCCEEDED(m_pImage->Lock(&rc, WICBitmapLockRead, &pBitmapLock)) {
						UINT cbSize = 0;
						LONG *pbData = NULL;
						if SUCCEEDED(pBitmapLock->GetDataPointer(&cbSize, (PBYTE *)&pbData)) {
							UINT x = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
							if (x < w) {
								UINT y = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
								if (y < h) {
									teSetLong(pVarResult, pbData[x + y * w]);
								}
							}
						}
						pBitmapLock->Release();
					}
				}
				return S_OK;
			//SetPixel
			case TE_METHOD + 113:
				if (nArg >= 2) {
					if (Get(GUID_WICPixelFormat32bppBGRA)) {
						UINT w = 0, h = 0;
						m_pImage->GetSize(&w, &h);
						IWICBitmapLock *pBitmapLock;
						WICRect rc = {0, 0, (INT)w, (INT)h};
						if SUCCEEDED(m_pImage->Lock(&rc, WICBitmapLockWrite, &pBitmapLock)) {
							UINT cbSize = 0;
							LONG *pbData = NULL;
							if SUCCEEDED(pBitmapLock->GetDataPointer(&cbSize, (PBYTE *)&pbData)) {
								UINT x = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
								if (x < w) {
									UINT y = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
									if (y < h) {
										pbData[x + y * w] = GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]);
									}
								}
							}
							SafeRelease(&pBitmapLock);
						}
					}
				}
				return S_OK;
			//GetPixelFormat
			case TE_METHOD + 114:
				WICPixelFormatGUID guidPF;
				m_pImage->GetPixelFormat(&guidPF);
				WCHAR pszBuff[40];
				StringFromGUID2(guidPF, pszBuff, 39);
				teSetSZ(pVarResult, pszBuff);
				return S_OK;
			//FillRect
			case TE_METHOD + 115:
				if (nArg >= 4) {
					if (Get(GUID_WICPixelFormat32bppBGRA)) {
						int w = 0, h = 0;
						m_pImage->GetSize((UINT *)&w, (UINT *)&h);
						IWICBitmapLock *pBitmapLock;
						WICRect rc = {0, 0, w, h};
						if SUCCEEDED(m_pImage->Lock(&rc, WICBitmapLockWrite, &pBitmapLock)) {
							UINT cbSize = 0;
							LONG *pbData = NULL;
							if SUCCEEDED(pBitmapLock->GetDataPointer(&cbSize, (PBYTE *)&pbData)) {
								int w0 = GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]);
								int h0 = GetIntFromVariant(&pDispParams->rgvarg[nArg - 3]);
								int x0 = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
								if (x0 < 0) {
									w0 -= x0;
									x0 = 0;
								}
								int y0 = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
								if (y0 < 0) {
									h0 -= y0;
									y0 = 0;
								}
								LONG cl = GetIntFromVariant(&pDispParams->rgvarg[nArg - 4]);
								LONG clBk = nArg >= 5 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 5]) : 0;
								for (int j = h0, y = y0; --j >= 0; ++y) {
									if (y < h) {
										for (int i = w0, x = x0; --i >= 0; ++x) {
											if (x < w) {
												LONG *pcl = &pbData[x + y * w];
												if (nArg < 5 || *pcl == clBk) {
													*pcl = cl;
												}
											} else {
												break;
											}
										}
									} else {
										break;
									}
								}
							}
							SafeRelease(&pBitmapLock);
						}
					}
				}
				return S_OK;

			case TE_METHOD + 116: //Mask
				if (nArg >= 1) {
					DWORD cl = GetIntFromVariant(&pDispParams->rgvarg[nArg]) & 0xffffff;
					DWORD clBk = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) & 0xffffff;
					if (Get(GUID_WICPixelFormat32bppBGRA)) {
						int w = 0, h = 0;
						m_pImage->GetSize((UINT *)&w, (UINT *)&h);
						IWICBitmapLock *pBitmapLock;
						WICRect rc = {0, 0, w, h};
						if SUCCEEDED(m_pImage->Lock(&rc, WICBitmapLockWrite, &pBitmapLock)) {
							UINT cbSize = 0;
							DWORD *pdwData = NULL;
							if SUCCEEDED(pBitmapLock->GetDataPointer(&cbSize, (PBYTE *)&pdwData)) {
								LONG clBk = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
								for (int y = h; --y >= 0;) {
									for (int x = w; --x >= 0;) {
										DWORD *pcl = &pdwData[x + y * w];
										DWORD a = *pcl << 24;
										*pcl = a ? a | cl : clBk;
									}
								}
							}
							SafeRelease(&pBitmapLock);
						}
					}
				}
				return S_OK;

			case TE_METHOD + 117://AlphaBlend
				if (nArg >= 1) {
					if (Get(GUID_WICPixelFormat32bppBGRA)) {
						int w = 0, h = 0;
						m_pImage->GetSize((UINT *)&w, (UINT *)&h);
						IWICBitmapLock *pBitmapLock;
						WICRect rc = {0, 0, w, h};
						if SUCCEEDED(m_pImage->Lock(&rc, WICBitmapLockWrite, &pBitmapLock)) {
							UINT cbSize = 0;
							RGBQUAD *pbData = NULL;
							if SUCCEEDED(pBitmapLock->GetDataPointer(&cbSize, (PBYTE *)&pbData)) {
								RECT rcFill = { 0, 0, w, h };
								LPRECT prc = (LPRECT)GetpcFromVariant(&pDispParams->rgvarg[nArg], NULL);
								if (prc) {
									CopyRect(&rcFill, prc);
								}
								LONG mix = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
								LONG max = nArg >= 2 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]) : 100;
								for (int y = rcFill.top; y < rcFill.bottom; ++y) {
									for (int x = rcFill.left; x < rcFill.right; ++x) {
										RGBQUAD *pcl = &pbData[x + y * w];
										pcl->rgbReserved = pcl->rgbReserved * mix / max;
									}
								}
							}
							SafeRelease(&pBitmapLock);
						}
					}
				}
				return S_OK;

			case TE_METHOD + 120: //GetThumbnailImage
				if (nArg >= 1) {
					//WIC
					UINT w = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					UINT h = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
					IWICBitmapScaler *pScaler;
					if SUCCEEDED(m_pWICFactory->CreateBitmapScaler(&pScaler)) {
						int mode = nArg >= 2 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]) : 4;
						if (!g_bUpper10 && mode > 3) {
							mode = 3;
						}
						if SUCCEEDED(pScaler->Initialize(m_pImage, w, h, WICBitmapInterpolationMode(mode))) {
							CteWICBitmap *pWB = new CteWICBitmap();
							m_pWICFactory->CreateBitmapFromSource(pScaler, WICBitmapCacheOnDemand, &pWB->m_pImage);
							SafeRelease(&pScaler);
							teSetObjectRelease(pVarResult, pWB);
						}
					}
				}
				return S_OK;
			//RotateFlip
			case TE_METHOD + 130:
				if (nArg >= 0) {
					IWICBitmapFlipRotator *pFlipRotator;
					if SUCCEEDED(m_pWICFactory->CreateBitmapFlipRotator(&pFlipRotator)) {
						if SUCCEEDED(pFlipRotator->Initialize(m_pImage, static_cast<WICBitmapTransformOptions>(GetIntFromVariant(&pDispParams->rgvarg[nArg])))) {
							IWICBitmap *pIBitmap;
							if SUCCEEDED(m_pWICFactory->CreateBitmapFromSource(pFlipRotator, WICBitmapCacheOnDemand, &pIBitmap)) {
								ClearImage(TRUE);
								m_pImage = pIBitmap;
							}
						}
						SafeRelease(&pFlipRotator);
					}
				}
				return S_OK;
			//GetFrameCount
			case TE_METHOD + 140:
				teSetLong(pVarResult, HasImage() ? m_uFrameCount : 1);
				return S_OK;
			//Frame
			case TE_PROPERTY + 150:
				if (nArg >= 0 && m_pStream) {
					UINT uFrame = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					if (m_uFrame != uFrame) {
						GetFrameFromStream(m_pStream, uFrame, FALSE);
					}
				}
				if (pVarResult && HasImage()) {
					teSetLong(pVarResult, m_uFrame);
				}
				return S_OK;
			//GetMetadata
			case TE_METHOD + 160:
			//GetFrameMetadata
			case TE_METHOD + 161:
				if (nArg >= 0 && pVarResult) {
					if (m_ppMetadataQueryReader[dispIdMember - TE_METHOD - 160]) {
						PROPVARIANT propVar;
						PropVariantInit(&propVar);
						if SUCCEEDED(m_ppMetadataQueryReader[dispIdMember - TE_METHOD - 160]->GetMetadataByName(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), &propVar)) {
							_PropVariantToVariant(&propVar, pVarResult);
							PropVariantClear(&propVar);
						}
					}
				}
				return S_OK;
			//GetHBITMAP
			case TE_METHOD + 210:
				teSetPtr(pVarResult, GetHBITMAP(nArg >= 0 ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : -1));
				return S_OK;
			//GetHICON
			case TE_METHOD + 211:
			//DrawEx
			case TE_METHOD + 212:
				HBITMAP hBM;
				if (hBM = GetHBITMAP(-1)) {
					BITMAP bm;
					GetObject(hBM, sizeof(BITMAP), &bm);
					HIMAGELIST himl = ImageList_Create(bm.bmWidth, bm.bmHeight, ILC_COLOR32, 0, 0);
					ImageList_Add(himl, hBM, NULL);
					DeleteObject(hBM);
					if (dispIdMember == TE_METHOD + 211) {
						teSetPtr(pVarResult, ImageList_GetIcon(himl, 0, nArg >= 0 ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : ILD_NORMAL));
					} else if (dispIdMember == TE_METHOD + 212 && nArg >= 7) {
						teSetBool(pVarResult, ImageList_DrawEx(himl, 0, (HDC)GetPtrFromVariant(&pDispParams->rgvarg[nArg]),
							GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]),
							GetIntFromVariant(&pDispParams->rgvarg[nArg - 3]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 4]),
							GetIntFromVariant(&pDispParams->rgvarg[nArg - 5]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 6]),
							GetIntFromVariant(&pDispParams->rgvarg[nArg - 7])
						));
					}
					ImageList_Destroy(himl);
				}
				return S_OK;
			//GetCodecInfo
			case TE_METHOD + 900:
				IDispatch *pdisp;
				if (nArg >= 2 && GetDispatch(&pDispParams->rgvarg[nArg - 2], &pdisp)) {
					IEnumUnknown *pEnum;
					HRESULT hr = m_pWICFactory->CreateComponentEnumerator(WICComponentType(GetIntFromVariant(&pDispParams->rgvarg[nArg])), WICComponentEnumerateDefault, &pEnum);
					if SUCCEEDED(hr) {
						int nMode = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
						ULONG cbActual = 0;
						IUnknown *punk;
						hr = E_FAIL;
						while(pEnum->Next(1, &punk, &cbActual) == S_OK && hr != S_OK) {
							IWICBitmapCodecInfo *pCodecInfo;
							if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCodecInfo))) {
								UINT uActual = 0;
								teGetCodecInfo(pCodecInfo, nMode, 0, NULL, &uActual);
								if (uActual > 1) {
									VARIANT *pv = GetNewVARIANT(1);
									pv->vt = VT_BSTR;
									pv->bstrVal = ::SysAllocStringLen(NULL, uActual - 1);
									teGetCodecInfo(pCodecInfo, nMode, uActual, pv->bstrVal, &uActual);
									teExecMethod(pdisp, L"push", NULL, 1, pv);
								}
								pCodecInfo->Release();
							}
							punk->Release();
						}
						pEnum->Release();
					}
					pdisp->Release();
				}
				return S_OK;
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
			case DISPID_TE_UNDEFIND:
				return S_OK;
			default:
				if (dispIdMember >= START_OnFunc && dispIdMember < START_OnFunc + Count_WICFunc) {
					teInvokeObject(&m_ppDispatch[dispIdMember - START_OnFunc], wFlags, pVarResult, nArg, pDispParams->rgvarg);
					return S_OK;
				}
		}
	} catch (...) {
		return teException(pExcepInfo, "WICBitmap", methodGB, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

#ifdef _USE_BSEARCHAPI
//CteWindowsAPI
CteWindowsAPI::CteWindowsAPI(TEDispatchApi *pApi)
{
	m_cRef = 1;
	m_pApi = pApi;
}

CteWindowsAPI::~CteWindowsAPI()
{
}

STDMETHODIMP CteWindowsAPI::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteWindowsAPI, IDispatch),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteWindowsAPI::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteWindowsAPI::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteWindowsAPI::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteWindowsAPI::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWindowsAPI::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	int nIndex = teBSearchApi(dispAPI, _countof(dispAPI), g_maps[MAP_API], *rgszNames);
	if (nIndex >= 0) {
		*rgDispId = nIndex + TE_METHOD;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"ADBSTRM") == 0) {
		*rgDispId = TE_PROPERTY;
		return S_OK;
	}
	*rgDispId = DISPID_TE_UNDEFIND;
	return S_OK;
}

STDMETHODIMP CteWindowsAPI::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (wFlags & DISPATCH_METHOD) {
		TEDispatchApi *pApi = (dispIdMember >= TE_METHOD && dispIdMember < _countof(dispAPI) + TE_METHOD) ? &dispAPI[dispIdMember - TE_METHOD] : m_pApi;
		if (pApi) {
			try {
				return teInvokeAPI(pApi, pDispParams, pVarResult, pExcepInfo);
			} catch (...) {
				return teExceptionEx(pExcepInfo, "WindowsAPI", pApi->name);
			}
		}
	} else if (dispIdMember == DISPID_VALUE) {
		teSetObject(pVarResult, this);
		return S_OK;
	} else if (dispIdMember >= TE_METHOD && dispIdMember < _countof(dispAPI) + TE_METHOD) {
		teSetObjectRelease(pVarResult, new CteWindowsAPI(&dispAPI[dispIdMember - TE_METHOD]));
		return S_OK;
	} else if (dispIdMember == TE_PROPERTY) {
		teSetSZ(pVarResult, L"ADODB.Stream");
		return S_OK;
	}
	return S_OK;
}
#endif
//CteDispatch

CteDispatch::CteDispatch(IDispatch *pDispatch, int nMode, DISPID dispId)
{
	m_cRef = 1;
	pDispatch->QueryInterface(IID_PPV_ARGS(&m_pDispatch));
	m_pDispatch2 = NULL;
	m_nMode = nMode;
	m_dispIdMember = dispId;
	m_nIndex = 0;
	m_pActiveScript = NULL;
	m_hDll = NULL;
}

CteDispatch::~CteDispatch()
{
	Clear();
}

VOID CteDispatch::Clear()
{
	SafeRelease(&m_pDispatch2);
	SafeRelease(&m_pDispatch);
	if (m_pActiveScript) {
		m_pActiveScript->SetScriptState(SCRIPTSTATE_CLOSED);
		m_pActiveScript->Close();
		SafeRelease(&m_pActiveScript);
	}
	teFreeLibrary(m_hDll, 100);
	m_hDll = NULL;
}


STDMETHODIMP CteDispatch::QueryInterface(REFIID riid, void **ppvObject)
{
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
		AddRef();
		return S_OK;
	}
	if (m_dispIdMember == DISPID_UNKNOWN && IsEqualIID(riid, IID_IEnumVARIANT)) {
		*ppvObject = static_cast<IEnumVARIANT *>(this);
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CteDispatch::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteDispatch::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteDispatch::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteDispatch::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDispatch::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (m_pDispatch) {
		if (m_nMode) {
			return m_pDispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
		} else {
			if (!m_pDispatch2) {
				VARIANT v;
				VariantInit(&v);
				Invoke5(m_pDispatch, m_dispIdMember, DISPATCH_METHOD, &v, 0, NULL);
				if (v.vt == VT_DISPATCH) {
					m_pDispatch2 = v.pdispVal;
				} else {
					VariantClear(&v);
				}
			}
			if (m_pDispatch2) {
				return m_pDispatch2->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
			}
		}
	}
	return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP CteDispatch::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (m_nMode & 3) {
			if (dispIdMember == DISPID_NEWENUM) {
				teSetObject(pVarResult, this);
				return S_OK;
			}
			if (!m_pDispatch) {
				return E_UNEXPECTED;
			}
			if (dispIdMember == DISPID_VALUE) {
				int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
				if (nArg >= 0) {
					VARIANTARG *pv = GetNewVARIANT(1);
					teVariantChangeType(pv, &pDispParams->rgvarg[nArg], VT_BSTR);
					m_pDispatch->GetIDsOfNames(IID_NULL, &pv[0].bstrVal, 1, LOCALE_USER_DEFAULT, &dispIdMember);
					VariantClear(&pv[0]);
					if (nArg >= 1) {
						VariantCopy(&pv[0], &pDispParams->rgvarg[nArg - 1]);
						Invoke5(m_pDispatch, dispIdMember, DISPATCH_PROPERTYPUT, NULL, 1, pv);
					} else {
						delete [] pv;
					}
					if (pVarResult) {
						Invoke5(m_pDispatch, dispIdMember, DISPATCH_PROPERTYGET, pVarResult, 0, NULL);
					}
					return S_OK;
				}
				teSetObject(pVarResult, this);
				return S_OK;
			}
			return m_pDispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		} else if (dispIdMember != DISPID_VALUE && m_pDispatch2) {
			return m_pDispatch2->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
		if (wFlags & DISPATCH_PROPERTYGET && dispIdMember == DISPID_VALUE) {
			teSetObject(pVarResult, this);
			return S_OK;
		}
		return m_pDispatch->Invoke((dispIdMember != DISPID_VALUE || m_nMode) ? dispIdMember : m_dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	} catch (...) {
		return teException(pExcepInfo, "Dispatch", NULL, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

STDMETHODIMP CteDispatch::Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched)
{
	if (!m_pDispatch) {
		return E_UNEXPECTED;
	}
	if (rgVar) {
		if (m_nMode) {
			IDispatchEx *pdex = NULL;
			if SUCCEEDED(m_pDispatch->QueryInterface(IID_PPV_ARGS(&pdex))) {
				HRESULT hr;
				hr = pdex->GetNextDispID(0, m_dispIdMember, &m_dispIdMember);
				if (hr == S_OK) {
					hr = pdex->GetMemberName(m_dispIdMember, &rgVar->bstrVal);
					rgVar->vt = VT_BSTR;
				}
				pdex->Release();
				return hr;
			}
		}
		int nCount = 0;
		VARIANT v;
		if (Invoke5(m_pDispatch, DISPID_TE_COUNT, DISPATCH_PROPERTYGET, &v, 0, NULL) == S_OK) {
			nCount = GetIntFromVariantClear(&v);
		}
		if (m_nIndex < nCount) {
			if (pCeltFetched) {
				*pCeltFetched = 1;
			}
			VARIANTARG *pv = GetNewVARIANT(1);
			teSetLong(&pv[0], m_nIndex++);
			return Invoke5(m_pDispatch, DISPID_TE_ITEM, DISPATCH_METHOD, rgVar, 1, pv);
		}
	}
	return S_FALSE;
}

STDMETHODIMP CteDispatch::Skip(ULONG celt)
{
	m_nIndex += celt;
	return S_OK;
}

STDMETHODIMP CteDispatch::Reset(void)
{
	m_nIndex = 0;
	return S_OK;
}

STDMETHODIMP CteDispatch::Clone(IEnumVARIANT **ppEnum)
{
	if (ppEnum && m_pDispatch) {
		CteDispatch *pdisp = new CteDispatch(m_pDispatch, m_nMode, m_dispIdMember);
		*ppEnum = static_cast<IEnumVARIANT *>(pdisp);
		return S_OK;
	}
	return E_POINTER;
}

//CteActiveScriptSite
CteActiveScriptSite::CteActiveScriptSite(IUnknown *punk, EXCEPINFO *pExcepInfo, HRESULT *phr)
{
	m_cRef = 1;
	m_phr = phr;
	m_pDispatchEx = NULL;
	if (punk) {
		punk->QueryInterface(IID_PPV_ARGS(&m_pDispatchEx));
	}
	m_pExcepInfo = pExcepInfo;
	if (!pExcepInfo) {
		m_pExcepInfo = &g_ExcepInfo;
	}
}

CteActiveScriptSite::~CteActiveScriptSite()
{
	SafeRelease(&m_pDispatchEx);
}

STDMETHODIMP CteActiveScriptSite::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteActiveScriptSite, IActiveScriptSite),
		QITABENT(CteActiveScriptSite, IActiveScriptSiteWindow),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteActiveScriptSite::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteActiveScriptSite::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteActiveScriptSite::GetLCID(LCID *plcid)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteActiveScriptSite::GetItemInfo(LPCOLESTR pstrName,
                   DWORD dwReturnMask,
                   IUnknown **ppiunkItem,
                   ITypeInfo **ppti)
{
	HRESULT hr = TYPE_E_ELEMENTNOTFOUND;
	if (dwReturnMask & SCRIPTINFO_IUNKNOWN) {
		if (m_pDispatchEx) {
			VARIANT v;
			VariantInit(&v);
			teGetProperty(m_pDispatchEx, const_cast<LPOLESTR>(pstrName), &v);
			if (FindUnknown(&v, ppiunkItem)) {
				return S_OK;
			}
			VariantClear(&v);
		}
		if (g_dwMainThreadId == GetCurrentThreadId()) {
			if (g_pWebBrowser && lstrcmpi(pstrName, L"window") == 0) {
				teGetHTMLWindow(g_pWebBrowser->m_pWebBrowser, IID_PPV_ARGS(ppiunkItem));
				return S_OK;
			}
#ifdef _USE_BSEARCHAPI
		} else if (lstrcmpi(pstrName, L"api") == 0) {
			*ppiunkItem = new CteWindowsAPI(NULL);
			return S_OK;
#endif
		}
	}
	return hr;
}

STDMETHODIMP CteActiveScriptSite::GetDocVersionString(BSTR *pbstrVersion)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteActiveScriptSite::OnScriptError(IActiveScriptError *pscripterror)
{
	if (!pscripterror) {
		return E_POINTER;
	}
	if (m_pExcepInfo && (SUCCEEDED(pscripterror->GetExceptionInfo(m_pExcepInfo)))) {
		BSTR bsSrcText = NULL;
		pscripterror->GetSourceLineText(&bsSrcText);
		DWORD dw1 = 0;
		ULONG ulLine = 0;
		LONG ulColumn = 0;
		pscripterror->GetSourcePosition(&dw1, &ulLine, &ulColumn);
		WCHAR pszLine[32];
		swprintf_s(pszLine, 32, L"\nLine: %d\n", ulLine);
		BSTR bs = ::SysAllocStringLen(NULL, ::SysStringLen(m_pExcepInfo->bstrDescription) + ::SysStringLen(bsSrcText) + lstrlen(pszLine));
		lstrcpy(bs, m_pExcepInfo->bstrDescription);
		lstrcat(bs, pszLine);
		if (bsSrcText) {
			lstrcat(bs, bsSrcText);
			::SysFreeString(bsSrcText);
		}
		teSysFreeString(&m_pExcepInfo->bstrDescription);
		m_pExcepInfo->bstrDescription = bs;
		if (m_pExcepInfo == &g_ExcepInfo) {
			MessageBox(NULL, bs, _T(PRODUCTNAME), MB_OK | MB_ICONERROR);
			teSysFreeString(&m_pExcepInfo->bstrDescription);
		}
#ifdef _DEBUG
		::OutputDebugString(bs);
		if (g_nBlink == 1) {
			MessageBox(NULL, bs, L"Tablacus Explorer", MB_OK);
		}
#endif
		*m_phr = m_pExcepInfo->scode;
	}
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::OnStateChange(SCRIPTSTATE ssScriptState)
{
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::OnScriptTerminate(const VARIANT *pvarResult,const EXCEPINFO *pexcepinfo)
{
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::OnEnterScript(void)
{
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::OnLeaveScript(void)
{
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::GetWindow(HWND *phwnd)
{
	*phwnd = g_hwndMain;
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::EnableModeless(BOOL fEnable)
{
	return E_NOTIMPL;
}

#ifdef _USE_OBJECTAPI
//CteAPI
CteAPI::CteAPI(TEDispatchApi *pApi)
{
	m_cRef = 1;
	m_pApi = pApi;
}

CteAPI::~CteAPI()
{
}

STDMETHODIMP CteAPI::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteAPI, IDispatch),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteAPI::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteAPI::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteAPI::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteAPI::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteAPI::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteAPI::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		if (wFlags & DISPATCH_METHOD) {
			return teInvokeAPI(m_pApi, pDispParams, pVarResult, pExcepInfo);
		}
		if (dispIdMember == DISPID_VALUE) {
			teSetObject(pVarResult, this);
			return S_OK;
		}
	} catch (...) {
		return teExceptionEx(pExcepInfo, "WindowsAPI", m_pApi->name);
	}
	return DISP_E_MEMBERNOTFOUND;
}
#endif

///CteDispatchEx

CteDispatchEx::CteDispatchEx(IUnknown *punk, BOOL bLegacy)
{
	m_cRef = 1;
	m_bDispathEx = FALSE;
	punk->QueryInterface(IID_PPV_ARGS(&m_pdex));
	m_bLegacy = bLegacy;
}

CteDispatchEx::~CteDispatchEx()
{
	SafeRelease(&m_pdex);
}

STDMETHODIMP CteDispatchEx::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteDispatchEx, IDispatch),
		QITABENT(CteDispatchEx, IDispatchEx),
		{ 0 },
	};
	if (IsEqualIID(riid, IID_IDispatchEx)) {
		m_bDispathEx = TRUE;
	}
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteDispatchEx::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteDispatchEx::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteDispatchEx::GetTypeInfoCount(UINT *pctinfo)
{
	return m_pdex->GetTypeInfoCount(pctinfo);
}

STDMETHODIMP CteDispatchEx::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return m_pdex->GetTypeInfo(iTInfo, lcid, ppTInfo);
}

STDMETHODIMP CteDispatchEx::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	HRESULT hr = m_bDispathEx ?
		m_pdex->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId) :
		m_pdex->GetDispID(*rgszNames, fdexNameEnsure | fdexNameCaseSensitive, rgDispId);
	if FAILED(hr) {
		GetLegacyDispId(*rgszNames, rgDispId, !m_bDispathEx, &hr);
	}
	return hr;
}

STDMETHODIMP CteDispatchEx::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (wFlags == DISPATCH_PROPERTYPUT) {
		if (pDispParams->cArgs > 0) {
			if (pDispParams->rgvarg[pDispParams->cArgs - 1].vt == VT_DISPATCH) {
				wFlags = DISPATCH_PROPERTYPUTREF;
			}
		}
	}
	HRESULT hr = m_bDispathEx ?
		m_pdex->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr) :
		m_pdex->InvokeEx(dispIdMember, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, NULL);
	if FAILED(hr) {
		InvokeLegacy(dispIdMember, wFlags, pDispParams, pVarResult, &hr);
	}
	if (wFlags == DISPATCH_PROPERTYPUT) {
		VARIANT v;
		VariantInit(&v);
		if (Invoke5(m_pdex, dispIdMember, DISPATCH_PROPERTYGET, &v, 0, NULL) == S_OK) {
			VariantClear(&v);
		}
	}
	return hr;
}

//IDispatchEx
STDMETHODIMP CteDispatchEx::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	HRESULT hr = m_pdex->GetDispID(bstrName, grfdex, pid);
	if FAILED(hr) {
		GetLegacyDispId(bstrName, pid, TRUE, &hr);
	}
	return hr;
}

STDMETHODIMP CteDispatchEx::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller)
{
	HRESULT hr = m_pdex->InvokeEx(id, lcid, wFlags, pdp, pvarRes, pei, pspCaller);
	if FAILED(hr) {
		InvokeLegacy(id, wFlags, pdp, pvarRes, &hr);
	}
	return hr;
}

STDMETHODIMP CteDispatchEx::DeleteMemberByName(BSTR bstrName, DWORD grfdex)
{
	m_pdex->DeleteMemberByName(bstrName, grfdex);
	return S_OK;
}

STDMETHODIMP CteDispatchEx::DeleteMemberByDispID(DISPID id)
{
	return m_pdex->DeleteMemberByDispID(id);
}

STDMETHODIMP CteDispatchEx::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex)
{
	return m_pdex->GetMemberProperties(id, grfdexFetch, pgrfdex);
}

STDMETHODIMP CteDispatchEx::GetMemberName(DISPID id, BSTR *pbstrName)
{
	return m_pdex->GetMemberName(id, pbstrName);
}

STDMETHODIMP CteDispatchEx::GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid)
{
	return m_pdex->GetNextDispID(grfdex, id, pid);
}

STDMETHODIMP CteDispatchEx::GetNameSpaceParent(IUnknown **ppunk)
{
	return m_pdex->GetNameSpaceParent(ppunk);
}

VOID CteDispatchEx::GetLegacyDispId(BSTR bstrName, DISPID *pid, BOOL bDispatchEx, HRESULT *phr)
{
	if (!m_bLegacy) {
		return;
	}
	if (lstrcmpi(bstrName, L"Count") == 0) {
		BSTR bs = ::SysAllocString(L"length");
		if (bDispatchEx) {
			*phr = m_pdex->GetDispID(bs, fdexNameCaseSensitive, pid);
		} else {
			*phr = m_pdex->GetIDsOfNames(IID_NULL, &bs, 1, LOCALE_USER_DEFAULT, pid);
		}
		::SysFreeString(bs);
	} else if (lstrcmp(bstrName, L"Item") == 0) {
		*pid = DISPID_COLLECT;
		*phr = S_OK;
	}
}

VOID CteDispatchEx::InvokeLegacy(DISPID id, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, HRESULT *phr)
{
	if (!m_bLegacy || id != DISPID_COLLECT) {
		return;
	}
	if ((wFlags & DISPATCH_METHOD) && pdp && pdp->cArgs > 0) {
		*phr = teGetPropertyAt(m_pdex, GetIntFromVariant(&pdp->rgvarg[pdp->cArgs - 1]), pvarRes);
		return;
	}
	if (wFlags == DISPATCH_PROPERTYGET) {
		teSetObjectRelease(pvarRes, new CteDispatch(this, 0, DISPID_COLLECT));
		*phr = S_OK;
	}
}

//*///

CteDropTarget2::CteDropTarget2(HWND hwnd, IUnknown *punk, BOOL bUseHelper)
{
	m_cRef = 1;
	m_pDragItems = NULL;
	m_pDropTarget = NULL;
	m_bUseHelper = bUseHelper;

	m_hwnd = hwnd;
	m_punk = punk;
}

CteDropTarget2::~CteDropTarget2()
{
	SafeRelease(&m_pDropTarget);
}

STDMETHODIMP CteDropTarget2::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDropTarget)) {
		*ppvObject = static_cast<IDropTarget *>(this);
		AddRef();
		return S_OK;
	}
	if (m_pDropTarget) {
		return m_pDropTarget->QueryInterface(riid, ppvObject);
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CteDropTarget2::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteDropTarget2::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

//IDropTarget
STDMETHODIMP CteDropTarget2::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	g_bDragging = TRUE;
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);
	DWORD dwEffect = *pdwEffect;
	HRESULT hr = DragSub(TE_OnDragEnter, m_punk, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (m_pDropTarget && dwEffect == *pdwEffect) {
		hr = m_pDropTarget->DragEnter(pDataObj, grfKeyState, pt, pdwEffect);
	} else if (g_bDragIcon && g_pDropTargetHelper) {
		m_bUseHelper = TRUE;
		hr = g_pDropTargetHelper->DragEnter(m_hwnd, pDataObj, (LPPOINT)&pt, *pdwEffect);
	}
	if (!g_bDragIcon && g_pDropTargetHelper) {
		g_pDropTargetHelper->DragLeave();
	}
	return hr;
}

STDMETHODIMP CteDropTarget2::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	DWORD dwEffect = *pdwEffect;
	HRESULT hr = DragSub(TE_OnDragOver, m_punk, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (m_pDropTarget && dwEffect == *pdwEffect) {
		hr = m_pDropTarget->DragOver(grfKeyState, pt, pdwEffect);
	} else if (g_bDragIcon && g_pDropTargetHelper) {
		m_bUseHelper = TRUE;
		g_pDropTargetHelper->DragOver((LPPOINT)&pt, *pdwEffect);
	}
	if (!g_bDragIcon && g_pDropTargetHelper) {
		g_pDropTargetHelper->DragLeave();
	}
	m_grfKeyState = grfKeyState;
	return hr;
}

STDMETHODIMP CteDropTarget2::DragLeave()
{
	g_bDragging = FALSE;
	if (m_DragLeave == E_NOT_SET) {
		if (m_bUseHelper && g_pDropTargetHelper) {
			g_pDropTargetHelper->DragLeave();
		}
		m_DragLeave = DragLeaveSub(m_punk, &m_pDragItems);
		if (m_pDropTarget) {
			HRESULT hr = m_pDropTarget->DragLeave();
			if (m_DragLeave != S_OK) {
				m_DragLeave = hr;
			}
		}
	}
	return m_DragLeave;
}

STDMETHODIMP CteDropTarget2::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	GetDragItems(&m_pDragItems, pDataObj);
	DWORD dwEffect = *pdwEffect;
	if (m_bUseHelper && g_pDropTargetHelper) {
		g_pDropTargetHelper->DragLeave();
	}
	HRESULT hr = DragSub(TE_OnDrop, m_punk, m_pDragItems, &m_grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		if FAILED(hr) {
			*pdwEffect = dwEffect;
			hr = m_pDropTarget->Drop(pDataObj, m_grfKeyState, pt, pdwEffect);
		}
	}
	if (m_bUseHelper && g_pDropTargetHelper) {
		g_pDropTargetHelper->Drop(pDataObj, (LPPOINT)&pt, *pdwEffect);
	}
	DragLeave();
	return hr;
}

#ifdef _USE_HTMLDOC
// CteDocHostUIHandler
CteDocHostUIHandler::CteDocHostUIHandler()
{
	m_cRef = 1;
}

CteDocHostUIHandler::~CteDocHostUIHandler()
{
}

STDMETHODIMP CteDocHostUIHandler::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDocHostUIHandler)) {
		*ppvObject = static_cast<IDocHostUIHandler *>(this);
	} else if (IsEqualIID(riid, IID_IServiceProvider)) {
		*ppvObject = static_cast<IServiceProvider *>(new CteServiceProvider(NULL, NULL));
		return S_OK;
	} else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
}

STDMETHODIMP_(ULONG) CteDocHostUIHandler::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteDocHostUIHandler::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}
//for IHTMLDocument2
//IDocHostUIHandler
STDMETHODIMP CteDocHostUIHandler::ShowContextMenu(DWORD dwID, POINT *ppt, IUnknown *pcmdtReserved, IDispatch *pdispReserved)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDocHostUIHandler::GetHostInfo(DOCHOSTUIINFO *pInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDocHostUIHandler::ShowUI(DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc)
{
	return S_FALSE;
}

STDMETHODIMP CteDocHostUIHandler::HideUI(VOID)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDocHostUIHandler::UpdateUI(VOID)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDocHostUIHandler::EnableModeless(BOOL fEnable)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDocHostUIHandler::OnDocWindowActivate(BOOL fActivate)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDocHostUIHandler::OnFrameWindowActivate(BOOL fActivate)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDocHostUIHandler::ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fFrameWindow)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDocHostUIHandler::TranslateAccelerator(LPMSG lpMsg, const GUID *pguidCmdGroup, DWORD nCmdID)
{
	return S_FALSE;
}

STDMETHODIMP CteDocHostUIHandler::GetOptionKeyPath(LPOLESTR *pchKey, DWORD dw)
{
	return S_FALSE;
}

STDMETHODIMP CteDocHostUIHandler::GetDropTarget(IDropTarget *pDropTarget, IDropTarget **ppDropTarget)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDocHostUIHandler::GetExternal(IDispatch **ppDispatch)
{
	return g_pTE->QueryInterface(IID_PPV_ARGS(ppDispatch));
}

STDMETHODIMP CteDocHostUIHandler::TranslateUrl(DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut)
{
	return S_FALSE;
}

STDMETHODIMP CteDocHostUIHandler::FilterDataObject(IDataObject *pDO, IDataObject **ppDORet)
{
	return E_NOTIMPL;
}
#endif
#ifdef _USE_TESTOBJECT
//CteTest
CteTest::CteTest()
{
	m_cRef = 1;
}

CteTest::~CteTest()
{
	Sleep(1);
}

STDMETHODIMP CteTest::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	} else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
}

STDMETHODIMP_(ULONG) CteTest::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteTest::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteTest::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteTest::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTest::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTest::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	switch (dispIdMember) {
		//this
		case DISPID_VALUE:
			teSetObject(pVarResult, this);
			return S_OK;
	}//end_switch
	return DISP_E_MEMBERNOTFOUND;
}
#endif

//CteProgressDialog

CteProgressDialog::CteProgressDialog(IProgressDialog *ppd)
{
	m_cRef = 1;
	m_ppd = NULL;
	if (ppd) {
		ppd->QueryInterface(IID_PPV_ARGS(&m_ppd));
	}
	if (!m_ppd) {
		teCreateInstance(CLSID_ProgressDialog, NULL, NULL, IID_PPV_ARGS(&m_ppd));
	}
}

CteProgressDialog::~CteProgressDialog()
{
	SafeRelease(&m_ppd);
}

STDMETHODIMP CteProgressDialog::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteProgressDialog, IDispatch),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteProgressDialog::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteProgressDialog::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteProgressDialog::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteProgressDialog::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteProgressDialog::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodPD, 0, NULL, *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteProgressDialog::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		LPCWSTR *pbs = NULL;
		DWORD *pdw = NULL;
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		VARIANT v;
		switch(dispIdMember) {
			case TE_METHOD + 1: //HasUserCancelled
				teSetBool(pVarResult, m_ppd->HasUserCancelled());

			case TE_METHOD + 4: //SetProgress
				if (nArg >= 1) {
					teSetProgress(m_ppd, GetLLFromVariant(&pDispParams->rgvarg[nArg]), GetLLFromVariant(&pDispParams->rgvarg[nArg - 1]), nArg >= 2 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]) : 0);
				}
				return S_OK;

			case TE_METHOD + 2: //SetCancelMsg
				if (nArg >= 0) {
					teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
					teSetLong(pVarResult, m_ppd->SetCancelMsg(v.bstrVal, NULL));
					VariantClear(&v);
				}
				return S_OK;

			case TE_METHOD + 3: //SetLine
				if (nArg >= 2) {
					teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
					teSetLong(pVarResult, m_ppd->SetLine(GetIntFromVariant(&pDispParams->rgvarg[nArg]), v.bstrVal, GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]), NULL));
					VariantClear(&v);
				}
				return S_OK;

			case TE_METHOD + 5: //SetTitle
				if (nArg >= 0) {
					teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
					teSetLong(pVarResult, m_ppd->SetTitle(v.bstrVal));
					VariantClear(&v);
				}
				return S_OK;

			case TE_METHOD + 6: //StartProgressDialog
				if (nArg >= 2) {
					IUnknown *punk = NULL;
					FindUnknown(&pDispParams->rgvarg[nArg - 1], &punk);
					teSetLong(pVarResult, m_ppd->StartProgressDialog((HWND)GetPtrFromVariant(&pDispParams->rgvarg[nArg]), punk, GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]), NULL));
				}
				return S_OK;

			case TE_METHOD + 7: //StopProgressDialog
				m_ppd->SetLine(2, L"", TRUE, NULL);
				teSetLong(pVarResult, m_ppd->StopProgressDialog());
				return S_OK;

			case TE_METHOD + 8: //Timer
				if (nArg >= 0) {
					teSetLong(pVarResult, m_ppd->Timer(GetIntFromVariant(&pDispParams->rgvarg[nArg]), NULL));
				}
				return S_OK;

			case TE_METHOD + 9: //SetAnimation
#ifdef _2000XP
				if (nArg >= 1) {
					teSetLong(pVarResult, m_ppd->SetAnimation((HINSTANCE)GetPtrFromVariant(&pDispParams->rgvarg[nArg]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 1])));
				}
#endif
				return S_OK;
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
			case DISPID_TE_UNDEFIND:
				return S_OK;
		}//end_switch
	} catch (...) {
		return teException(pExcepInfo, "ProgressDialog", methodPD, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteObject
CteObject::CteObject(PVOID pObj)
{
	m_punk = NULL;
	m_liOffset.QuadPart = 0;
	if (pObj) {
		try {
			IUnknown *punk = static_cast<IUnknown *>(pObj);
			punk->QueryInterface(IID_PPV_ARGS(&m_punk));
			IStream *pStream;
			if SUCCEEDED(m_punk->QueryInterface(IID_PPV_ARGS(&pStream))) {
				m_liOffset.QuadPart = teGetStreamPos(pStream);
				pStream->Release();
			}
		} catch (...) {
			g_nException = 0;
#ifdef _DEBUG
			g_strException = L"CteObject";
#endif
		}
	}
	m_cRef = 1;
}

CteObject::~CteObject()
{
	Clear();
}

VOID CteObject::Clear()
{
	SafeRelease(&m_punk);
}

STDMETHODIMP CteObject::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
		AddRef();
		return S_OK;
	}
	if (m_punk) {
		if (IsEqualIID(riid, IID_IStream)) {
			IStream *pStream;
			if SUCCEEDED(m_punk->QueryInterface(IID_PPV_ARGS(&pStream))) {
				pStream->Seek(m_liOffset, STREAM_SEEK_SET, NULL);
				pStream->Release();
			}
		}
		return m_punk->QueryInterface(riid, ppvObject);
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CteObject::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteObject::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteObject::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteObject::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteObject::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodCO, 0, NULL, *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteObject::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
		teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
		return S_OK;
	}
	switch (dispIdMember) {
		case TE_METHOD + 1:
			Clear();
			return S_OK;
		case DISPID_VALUE:
			teSetObject(pVarResult, this);
		case DISPID_TE_UNDEFIND:
			return S_OK;
	}//end_switch
	return DISP_E_MEMBERNOTFOUND;
}

//CteFileSystemBindData
CteFileSystemBindData::CteFileSystemBindData()
{
	m_cRef = 1;
}

CteFileSystemBindData::~CteFileSystemBindData()
{
}

STDMETHODIMP CteFileSystemBindData::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteFileSystemBindData, IFileSystemBindData),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteFileSystemBindData::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteFileSystemBindData::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteFileSystemBindData::SetFindData(const WIN32_FIND_DATAW *pfd)
{
	m_wfd = *pfd;
	return S_OK;
}

STDMETHODIMP CteFileSystemBindData::GetFindData(WIN32_FIND_DATAW *pfd)
{
	*pfd = m_wfd;
	return S_OK;
}

//CteEnumerator
CteEnumerator::CteEnumerator(VARIANT *pvObject)
{
	m_cRef = 1;
	VariantInit(&m_vItem);
	m_pEnumVARIANT = NULL;
	m_pdex = NULL;
	m_hr = E_NOTIMPL;
	IDispatch *pdisp;
	m_dispIdMember = DISPID_UNKNOWN;
	if (GetDispatch(pvObject, &pdisp)) {
		pdisp->QueryInterface(IID_PPV_ARGS(&m_pdex));
		VARIANT v;
		VariantInit(&v);
		IShellWindows *pSW = NULL;
		FolderItems *pid = NULL;
		IUnknown *punkEnum = NULL;
		if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pSW))) {
			pSW->_NewEnum(&punkEnum);
		} else if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pid))) {
			pid->_NewEnum(&punkEnum);
		} else if SUCCEEDED(Invoke5(pdisp, DISPID_NEWENUM, DISPATCH_PROPERTYGET, &v, 0, NULL)) {
			FindUnknown(&v, &punkEnum);
		}
		if (punkEnum) {
			punkEnum->QueryInterface(IID_PPV_ARGS(&m_pEnumVARIANT));
		}
		VariantClear(&v);
		SafeRelease(&pid);
		SafeRelease(&pSW);
		pdisp->Release();
	}
	GetItem();
}

CteEnumerator::~CteEnumerator()
{
	VariantClear(&m_vItem);
	SafeRelease(&m_pEnumVARIANT);
	SafeRelease(&m_pdex);
}

STDMETHODIMP CteEnumerator::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CTE, IDispatch),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteEnumerator::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteEnumerator::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteEnumerator::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteEnumerator::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteEnumerator::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (teGetDispId(methodEN, 0, NULL, *rgszNames, rgDispId, TRUE) != S_OK) {
		*rgDispId = DISPID_TE_UNDEFIND;
	}
	return S_OK;
}

STDMETHODIMP CteEnumerator::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
		teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
		return S_OK;
	}
	switch (dispIdMember) {
	case TE_METHOD + 1://atEnd
		teSetBool(pVarResult, m_hr);
		return S_OK;

	case TE_METHOD + 2://item
		teVariantCopy(pVarResult, &m_vItem);
		return S_OK;

	case TE_METHOD + 3://moveFirst
		if (m_pEnumVARIANT) {
			m_pEnumVARIANT->Reset();
		}
		m_dispIdMember = DISPID_UNKNOWN;
		GetItem();
		return S_OK;

	case TE_METHOD + 4://moveNext
		GetItem();
		return S_OK;

	case DISPID_VALUE:
		teSetObject(pVarResult, this);
	case DISPID_TE_UNDEFIND:
		return S_OK;
	}//end_switch
	return DISP_E_MEMBERNOTFOUND;
}

VOID CteEnumerator::GetItem()
{
	if (m_pEnumVARIANT) {
		ULONG uFetched;
		m_hr = m_pEnumVARIANT->Next(1, &m_vItem, &uFetched);
		return;
	}

	if (m_pdex) {
		m_hr = m_pdex->GetNextDispID(0, m_dispIdMember, &m_dispIdMember);
		if (m_hr == S_OK) {
			m_hr = m_pdex->GetMemberName(m_dispIdMember, &m_vItem.bstrVal);
			if (m_hr == S_OK) {
				m_vItem.vt = VT_BSTR;
			}
		}
	}
}
