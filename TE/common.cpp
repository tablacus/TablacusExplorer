#include "stdafx.h"
#if defined(_WINDLL) || defined(_DEBUG)

#include "common.h"

extern HWND g_hwndMain;
extern int g_nException;
extern int g_nBlink;
extern IDispatch *g_pMBText;
extern LPITEMIDLIST g_pidls[MAX_CSIDL2];
extern BSTR	g_bsAnyText;
extern BSTR	g_bsName;
extern BSTR	g_bsAnyTextId;
extern BOOL	g_bUpperVista;
extern int		g_param[Count_TE_params];
extern BSTR		 g_bsPidls[MAX_CSIDL2];
extern IQueryParser *g_pqp;
extern DWORD	g_dwMainThreadId;
extern std::unordered_map<std::string, LONG> g_maps[MAP_LENGTH];
extern DWORD	g_dwFreeLibrary;
extern std::vector<HMODULE>	g_pFreeLibrary;
extern WCHAR	g_szTE[MAX_LOADSTRING];
extern FORMATETC HDROPFormat;
extern FORMATETC IDLISTFormat;
extern LPFNChangeWindowMessageFilterEx _ChangeWindowMessageFilterEx;
extern int		g_nDropState;
extern IDropTargetHelper *g_pDropTargetHelper;
extern IDataObject	*g_pDraggingItems;
extern BSTR	g_bsCmdLine;
extern BSTR	g_bsDateTimeFormat;
extern BSTR	g_bsTitle;
extern TEStruct pTEStructs[];
extern IDispatch	*g_pJS;
extern DWORD   g_dwCookieJS;
extern IDispatch	*g_pOnFunc[Count_OnFunc];
extern GUID		g_ClsIdFI;

#ifdef _DEBUG
extern LPWSTR	g_strException;
#endif
#ifdef _WINDLL
extern HINSTANCE g_hinstDll;
#endif
#ifdef _2000XP
extern LPFNSetDllDirectoryW _SetDllDirectoryW;
extern LPFNIsWow64Process _IsWow64Process;
extern LPFNPSPropertyKeyFromString _PSPropertyKeyFromString;
extern LPFNPSGetPropertyKeyFromName _PSGetPropertyKeyFromName;
extern LPFNPSPropertyKeyFromString _PSPropertyKeyFromStringEx;
extern LPFNPSGetPropertyDescription _PSGetPropertyDescription;
extern LPFNPSStringFromPropertyKey _PSStringFromPropertyKey;
extern LPFNPropVariantToVariant _PropVariantToVariant;
extern LPFNVariantToPropVariant _VariantToPropVariant;
extern LPFNSHCreateItemFromIDList _SHCreateItemFromIDList;
extern LPFNSHGetIDListFromObject _SHGetIDListFromObject;
extern LPFNChangeWindowMessageFilter _ChangeWindowMessageFilter;
extern LPFNAddClipboardFormatListener _AddClipboardFormatListener;
extern LPFNRemoveClipboardFormatListener _RemoveClipboardFormatListener;
extern LPFNSHCreateShellItemArrayFromShellItem _SHCreateShellItemArrayFromShellItem;
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

extern char* GetpcFromVariant(VARIANT *pv, VARIANT *pvMem);
extern IDropSource* teFindDropSource();
extern IDispatch* teCreateObj(LONG lId, VARIANT *pvArg);
extern VOID CALLBACK teTimerProcParse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
extern VOID CALLBACK teTimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
extern BOOL GetDataObjFromObject(IDataObject **ppDataObj, IUnknown *punk);

UINT	g_uCrcTable[256];
long	g_nProcDrag    = 0;

VOID teInitCommon()
{
	for (UINT i = 0; i < 256; ++i) {
		UINT c = i;
		for (int j = 8; j--;) {
			c = (c & 1) ? (0xedb88320 ^ (c >> 1)) : (c >> 1);
		}
		g_uCrcTable[i] = c;
	}

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

BOOL teStartsText(LPWSTR pszSub, LPCWSTR pszFile)
{
	BOOL bResult = pszFile ? TRUE : FALSE;
	WCHAR wc;
	while (bResult && (wc = *pszSub++)) {
		bResult = towlower(wc) == towlower(*pszFile++);
	}
	return bResult;
}

BOOL teVarIsNumber(VARIANT *pv) {
	return pv->vt == VT_I4 || pv->vt == VT_R8 || pv->vt == (VT_ARRAY | VT_I4) || (pv->vt == VT_BSTR && ::SysStringLen(pv->bstrVal) == 18 && teStartsText(L"0x", pv->bstrVal));
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

int GetIntFromVariantClear(VARIANT *pv)
{
	int i = GetIntFromVariant(pv);
	VariantClear(pv);
	return i;
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

HRESULT teGetPropertyAt(IDispatch *pdisp, int i, VARIANT *pv)
{
	wchar_t pszName[8];
	swprintf_s(pszName, 8, L"%d", i);
	return teGetProperty(pdisp, pszName, pv);
}

VOID Invoke4(IDispatch *pdisp, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs)
{
	Invoke5(pdisp, DISPID_VALUE, DISPATCH_METHOD, pvResult, nArgs, pvArgs);
}

HRESULT Invoke5(IDispatch *pdisp, DISPID dispid, WORD wFlags, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs)
{
	HRESULT hr = E_UNEXPECTED;
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
		if (pdisp) {
			hr = pdisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, wFlags, &dispParams, pvResult, NULL, NULL);
		}
#ifdef _DEBUG
		else {
			Sleep(1);
		}
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
		hr = E_UNEXPECTED;
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

/*
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
*/
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

VOID teCreateSearchPath(BSTR *pbs, LPWSTR pszPath, LPWSTR pszSearch)
{
	LPWSTR pszHeader = L"search-ms:crumb=";
	LPWSTR pszLocation = L"&crumb=location:";
	DWORD dw = 2048;
	BSTR bsSearch3 = teSysAllocStringLen(L"", dw);
	UrlEscape(pszSearch, bsSearch3, &dw, URL_ESCAPE_SEGMENT_ONLY);
	dw = 2048;
	BSTR bsPath3 = ::SysAllocStringLen(L"", dw);
	UrlEscape(pszPath, bsPath3, &dw, URL_ESCAPE_SEGMENT_ONLY);
	teSysFreeString(pbs);
	*pbs = ::SysAllocStringLen(NULL, lstrlen(pszHeader) + lstrlen(bsSearch3) + lstrlen(pszLocation) + lstrlen(bsPath3));
	lstrcpy(*pbs, pszHeader);
	lstrcat(*pbs, bsSearch3);
	lstrcat(*pbs, pszLocation);
	lstrcat(*pbs, bsPath3);
	teSysFreeString(&bsPath3);
	teSysFreeString(&bsSearch3);
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
			if (!teStrCmpI(bsSearch1, &bsSearch2[nPos >= 0 && StrChr(bsSearch2, '~') ? nPos : 0]) || StrStrI(bsSearch2, g_bsAnyText)|| StrStrI(bsSearch2, g_bsAnyTextId)) {
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
												swprintf_s(&bsSearch[i], ::SysStringLen(bsSearch) - i, L"(%c..%c OR ~<%c)", pszSearch[0], pszSearch[4], pszSearch[4]);
											} else {
												if (pszComma) {
													StrNCpy(&bsSearch[i], pszSearch, pszComma - pszSearch + 1);
												} else {
													lstrcpy(&bsSearch[i], pszSearch);
												}
												for (LPWSTR pszIn = &bsSearch[i], pszOut = pszIn;;) {
													WCHAR wc = *pszIn++;
													if (wc == ' ') {
														continue;
													}
													if (!wc || wc == '(') {
														*pszOut = NULL;
														break;
													}
													*pszOut++ = tolower(wc);
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
	if (pSF->EnumObjects(hwnd, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN | SHCONTF_INCLUDESUPERHIDDEN | SHCONTF_NAVIGATION_ENUM, &peidl) == S_OK) {
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
					if (teStrCmpI(lpfn, pszPath) == 0) {
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

VOID tePathAppend(BSTR *pbsPath, LPCWSTR pszPath, LPWSTR pszFile)
{
	BSTR bsPath = teSysAllocStringLen(pszPath, lstrlen(pszPath) + lstrlen(pszFile) + 1);
	PathAppend(bsPath, pszFile);
	*pbsPath = ::SysAllocString(bsPath);
	teSysFreeString(&bsPath);
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
					if (teStrCmpI(bsSearchX, pszNull) == 0) {
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
							LPITEMIDLIST pidl2 = ::SHSimpleIDListFromPath(*pbs);
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

BOOL GetCSIDLFromPath(int *i, LPWSTR pszPath)
{
	int n = lstrlen(pszPath);
	if (n < 3 && pszPath[0] >= '0' && pszPath[0] <= '9') {
		swscanf_s(pszPath, L"%d", i);
		return (*i <= 9 && n == 1) || (*i >= 10 && *i < MAX_CSIDL2);
	}
	return FALSE;
}

int ILGetCount(LPCITEMIDLIST pidl)
{
	if (ILIsEmpty(pidl)) {
		return 0;
	}
	return ILGetCount(::ILGetNext(pidl)) + 1;
}

BOOL teGetIDListFromObject(IUnknown *punk, LPITEMIDLIST *ppidl)
{
	*ppidl = NULL;
	if (!punk) {
		return FALSE;
	}
#ifdef CHECK_HANDLELEAK
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

int teStrCmpI(LPCWSTR lpStringW, LPCWSTR lpString2) {
	int wc1 = lpStringW ? tolower(*lpStringW) : NULL;
	int wc2 = lpString2 ? tolower(*lpString2) : NULL;
	int result = wc1 - wc2;
	if (result || wc1 == NULL || wc2 == NULL) {
		return result;
	}
	for (int i = 1;;++i) {
		wc1 = tolower(lpStringW[i]);
		wc2 = tolower(lpString2[i]);
		result = wc1 - wc2;
		if (result || wc1 == NULL || wc2 == NULL) {
			break;
		}
	};
	return result;
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

VOID teSysFreeString(BSTR *pbs)
{
	if (*pbs) {
		::SysFreeString(*pbs);
		*pbs = NULL;
	}
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

LPITEMIDLIST teILCreateFromPathEx(LPWSTR pszPath)
{
	LPITEMIDLIST pidl = NULL;
	IBindCtx *pbc = NULL;
	ULONG chEaten;
	ULONG dwAttributes;
	IShellFolder *pSF;
	if SUCCEEDED(SHGetDesktopFolder(&pSF)) {
		if SUCCEEDED(CreateBindCtx(0, &pbc)) {
			pbc->RegisterObjectParam(STR_PARSE_PREFER_FOLDER_BROWSING, teFindDropSource());
		}
		try {
			pSF->ParseDisplayName(NULL, pbc, pszPath, &chEaten, &pidl, &dwAttributes);
		} catch (...) {
			pidl = NULL;
		}
		SafeRelease(&pbc);
		pSF->Release();
	}
#ifdef _DEBUG
	if (g_dwMainThreadId == GetCurrentThreadId()) {
		if (teStartsText(L"\\\\", pszPath)) {
			Sleep(0);
		}
	}
#endif
	return pidl;
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

BOOL teIsFileSystem(LPOLESTR pszPath)
{
	return tePathMatchSpec(pszPath, L"?:\\*;\\\\*\\*") && !teStartsText(L"\\\\\\", pszPath);
}

BOOL teIsSearchFolder(LPWSTR lpszPath)
{
	return teStartsText(L"search-ms:", lpszPath);
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
#ifdef USE_TESTPATHMATCHSPEC
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

BOOL teStrSameIFree(BSTR bs, LPWSTR lpstr2)
{
	BOOL b = teStrCmpI(bs, lpstr2) == 0;
	::SysFreeString(bs);
	return b;
}

int teStrCmpIWA(LPCWSTR lpStringW, LPCSTR lpStringA) {
	int wc1 = lpStringW ? tolower(*lpStringW) : NULL;
	int wc2 = lpStringA ? tolower(*lpStringA) : NULL;
	int result = wc1 - wc2;
	if (result || wc1 == NULL || wc2 == NULL) {
		return result;
	}
	for (int i = 1;;++i) {
		wc1 = tolower(lpStringW[i]);
		wc2 = tolower(lpStringA[i]);
		result = wc1 - wc2;
		if (result || wc1 == NULL || wc2 == NULL) {
			break;
		}
	};
	return result;
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

int teGetObjectLength(IDispatch *pdisp)
{
	VARIANT v;
	VariantInit(&v);
	teGetProperty(pdisp, L"length", &v);
	return GetIntFromVariantClear(&v);
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

VOID teReleaseExists(TEExists *pExists)
{
	if (::InterlockedDecrement(&pExists->cRef) == 0) {
		CloseHandle(pExists->hEvent);
		delete [] pExists;
	}
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
			if (hr == E_INVALID_PASSWORD || hr == E_CANCELLED || E_LOGON_FAILURE) {
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
#ifdef CHECK_HANDLELEAK
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
			if SUCCEEDED(teCreateInstance(CLSID_SearchFolderItemFactory, NULL, NULL, IID_PPV_ARGS(&psfif))) {
				IShellItem *psi = NULL;
				if (teCreateItemFromPath(bsPath3, &psi)) {
					teSysFreeString(&bsPath3);
					IShellItemArray *psia;
					if (
#ifdef _2000XP
						_SHCreateShellItemArrayFromShellItem &&
#endif
						SUCCEEDED(_SHCreateShellItemArrayFromShellItem(psi, IID_PPV_ARGS(&psia)))) {
						psfif->SetScope(psia);
						psia->Release();
						teGetSearchArg(&bsPath3, pszPath, L"crumb=");
						psfif->SetDisplayName(bsPath3);
						if (!g_pqp) {
							IQueryParserManager *pqpm = NULL;
							if SUCCEEDED(CoCreateInstance(__uuidof(QueryParserManager), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pqpm))) {
								if SUCCEEDED(pqpm->CreateLoadedParser(L"SystemIndex", MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL), IID_PPV_ARGS(&g_pqp))) {
									BOOL fUnderstandNQS = FALSE;
									BOOL fAutoWildCard = TRUE;
									HKEY hKey;
									if (RegOpenKeyExA(HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Search\\Preferences", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
										DWORD dwSize = sizeof(BOOL);
										RegQueryValueExA(hKey, "EnableNaturalQuerySyntax", NULL, NULL, (LPBYTE)&fUnderstandNQS, &dwSize);
										RegQueryValueExA(hKey, "AutoWildCard", NULL, NULL, (LPBYTE)&fAutoWildCard, &dwSize);
										RegCloseKey(hKey);
									}
									pqpm->InitializeOptions(fUnderstandNQS, fAutoWildCard, g_pqp);
									pqpm->SetOption(QPMO_PRELOCALIZED_SCHEMA_BINARY_PATH, FALSE);
									for (int i = 0; i < ARRAYSIZE(g_rgGenericProperties); ++i) {
										PROPVARIANT propvar;
										propvar.pwszVal = const_cast<LPWSTR>(g_rgGenericProperties[i].pszPropertyName);
										propvar.vt = VT_LPWSTR;
										g_pqp->SetMultiOption(SQMO_DEFAULT_PROPERTY, g_rgGenericProperties[i].pszSemanticType, &propvar);
									}
								}
								pqpm->Release();
							}
						}
						if (g_pqp) {
							IQuerySolution *pqs = NULL;
							ICondition *pc, *pc1;
							if (SUCCEEDED(g_pqp->Parse(bsPath3, NULL, &pqs)) && SUCCEEDED(pqs->GetQuery(&pc, NULL))) {
								SYSTEMTIME st;
								GetLocalTime(&st);
								IConditionFactory2* pcf2;
								if SUCCEEDED(pqs->QueryInterface(IID_PPV_ARGS(&pcf2))) {
									ICondition2* pc2;
									if SUCCEEDED(pcf2->ResolveCondition(pc, SQRO_DEFAULT, &st, IID_PPV_ARGS(&pc2))) {
										psfif->SetCondition(pc2);
										pc2->Release();
									}
									pcf2->Release();																	   
								} else if SUCCEEDED(pqs->Resolve(pc, SQRO_DONT_SPLIT_WORDS, &st, &pc1)) {
									psfif->SetCondition(pc1);
									pc1->Release();
								} else {
									psfif->SetCondition(pc);
								}
								pc->Release();
								psfif->GetIDList(&pidl);
							}
							teSysFreeString(&bsPath3);
							SafeRelease(&pqs);
						}
					}
					SafeRelease(&psi);
				}
			}
			SafeRelease(&psfif);
			teSysFreeString(&bsPath3);
			teSysFreeString(&bsPath2);
			return pidl;
		}
		if (tePathMatchSpec(pszPath, L"*\\..\\*;*\\..;*\\.\\*;*\\.;*%*%*")) {
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
					if (tePathIsNetworkPath(pszPath) && PathIsRoot(pszPath)) {
						HRESULT hr = tePathIsDirectory(pszPath, 0, 3);
						if (FAILED(hr) && hr != E_ACCESSDENIED) {
							teILFreeClear(&pidl);
						}
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
						if (!teStrCmpI(pszPath, g_bsPidls[i])) {
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

VOID teReleaseILCreate(TEILCreate *pILC)
{
	if (::InterlockedDecrement(&pILC->cRef) == 0) {
		teCoTaskMemFree(pILC->pidlResult);
		CloseHandle(pILC->hEvent);
		delete [] pILC;
	}
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
	DWORD dwms = teFindDropSource() ? g_param[TE_NetworkTimeout] : 2000;
	if (bForceLimit && (!dwms || dwms > 500)) {
		dwms = 500;
	}
#ifdef CHECK_HANDLELEAK
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
				if (pidl = pILC->pidlResult) {
					pILC->pidlResult = NULL;
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
/*
LPITEMIDLIST teSimpleILCreateFromPath(LPWSTR pszPath)
{
if (teIsFileSystem(pszPath)) {
return teSHSimpleIDListFromPathEx(pszPath, FILE_ATTRIBUTE_DIRECTORY, 0, 0, NULL);
}
return teILCreateFromPath(pszPath);
}
*/

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

int teUMSearch(int nMap, LPOLESTR bs)
{
	CHAR pszName[32];
	for (int j = 0; pszName[j] = (bs && j < _countof(pszName) - 1) ? tolower(bs[j]) : NULL; ++j);
	auto itr = g_maps[nMap].find(pszName);
	if (itr != g_maps[nMap].end()) {
		return itr->second;
	}
	return -1;
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

HRESULT tePutProperty(IUnknown *punk, LPOLESTR sz, VARIANT *pv)
{
	HRESULT hr = tePutProperty0(punk, sz, pv, fdexNameEnsure);
	VariantClear(pv);
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

LONGLONG teGetU(LONGLONG ll)
{
	if ((ll & (~MAXINT)) == (~MAXINT)) {
		return ll & MAXUINT;
	}
	return ll & MAXINT64;
}

int CalcCrc32(BYTE *pc, int nLen, UINT c)
{
	c ^= MAXUINT;
	for (int i = 0; i < nLen; ++i) {
		c = g_uCrcTable[LOBYTE(c ^ pc[i])] ^ (c >> 8);
	}
	return c ^ MAXUINT;
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
		hr = DoDragDrop(pDataObj, teFindDropSource(), *pdwEffect, pdwEffect);
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

HRESULT tePutPropertyAt(PVOID pObj, int i, VARIANT *pv)
{
	wchar_t pszName[8];
	swprintf_s(pszName, 8, L"%d", i);
	return tePutProperty((IUnknown *)pObj, pszName, pv);
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

HRESULT teCLSIDFromProgID(__in LPCOLESTR lpszProgID, __out LPCLSID lpclsid)
{
	if (teStrCmpIWA(lpszProgID, "ads") == 0) {
		*lpclsid = CLSID_ADODBStream;
		return S_OK;
	}
	if (teStrCmpIWA(lpszProgID, "fso") == 0) {
		*lpclsid = CLSID_ScriptingFileSystemObject;
		return S_OK;
	}
	if (teStrCmpIWA(lpszProgID, "sha") == 0) {
		*lpclsid = CLSID_Shell;
		return S_OK;
	}
	if (teStrCmpIWA(lpszProgID, "wsh") == 0) {
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
					HANDLE hFile = CreateFile(bsPath, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, NULL);
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

HRESULT teStopProgressDialog(IProgressDialog *ppd)
{
	HWND hwnd;
	if (IUnknown_GetWindow(ppd, &hwnd) == S_OK) {
		ShowWindow(hwnd, SW_HIDE);
	}
	return ppd->StopProgressDialog();
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
		if (teGetIDListFromVariant(&pidl, pv, TRUE)) {
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

void GetVarPathFromIDList(VARIANT *pVarResult, LPITEMIDLIST pidl, int uFlags)
{
	int i;

	if (uFlags & SHGDN_FORPARSINGEX) {
		for (i = 0; i < MAX_CSIDL; ++i) {
			if (g_pidls[i] && ::ILIsEqual(pidl, g_pidls[i])) {
				LPITEMIDLIST pidl2 = ::SHSimpleIDListFromPath(g_bsPidls[i]);
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
#ifdef _WINDLL
	if (!hModule) {
		teGetModuleFileName(g_hinstDll, pbsPath);
		::PathRemoveFileSpec(*pbsPath);
		::PathRemoveFileSpec(*pbsPath);
#ifdef _WIN64
#ifdef _DEBUG
		PathAppend(*pbsPath, L"TEd64.exe");
#else
		::PathAppend(*pbsPath, L"TE64.exe");
#endif
#else
#ifdef _DEBUG
		PathAppend(*pbsPath, L"TEd32.exe");
#else
		PathAppend(*pbsPath, L"TE32.exe");
#endif
#endif
		i = lstrlen(*pbsPath);
	}
#endif
	return i;
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
	if (nSizeFormat && (IsEqualPropertyKey(*ppropKey, PKEY_Size) || IsEqualPropertyKey(*ppropKey, PKEY_TotalFileSize)) && teVarIsNumber(pv)) {
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

VOID teSetExStyleOr(HWND hwnd, LONG l)
{
	LONG exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
	if (exStyle != (exStyle | l)) {
		SetWindowLong(hwnd, GWL_EXSTYLE, exStyle | l);
	}
}

VOID teSetExStyleAnd(HWND hwnd, LONG l)
{
	LONG exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
	if (exStyle != (exStyle & l)) {
		SetWindowLongPtr(hwnd, GWL_EXSTYLE, exStyle & l);
	}
}

BOOL teGetIDListFromVariant(LPITEMIDLIST *ppidl, VARIANT *pv, BOOL bForceLimit)
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

int GetSizeOfStruct(LPOLESTR bs)
{
	int nIndex = teUMSearch(MAP_SS, bs);
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

VOID teArrayPush(IDispatch *pdisp, PVOID pObj)
{
	VARIANT *pv = GetNewVARIANT(1);
	teSetObject(pv, pObj);
	teExecMethod(pdisp, L"push", NULL, 1, pv);
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

VOID teSetUM(int nMap, LPCSTR name, LONG lData)
{
	CHAR pszName[32];
#ifdef _DEBUG
	if (lstrlenA(name) >= _countof(pszName) - 1) {
		MessageBoxA(NULL, name, "Too long >=32", MB_OK);
	}
#endif
	for (int j = 0; pszName[j] = tolower(name[j]); ++j);
	g_maps[nMap][pszName] = lData;
}

VOID teInitUM(int nMap, TEmethod *method, int nCount)
{
	for (int i = 0; i < nCount; ++i) {
		teSetUM(nMap, method[i].name, method[i].id);
	}
}

int teBSearch(TEmethod *method, int nSize, LPOLESTR bs)
{
	int nMin = 0;
	int nMax = nSize - 1;
	int nIndex, nCC;

	while (nMin <= nMax) {
		nIndex = (nMin + nMax) / 2;
		nCC = teStrCmpIWA(bs, method[nIndex].name);
		if (nCC < 0) {
			nMax = nIndex - 1;
			continue;
		}
		if (nCC > 0) {
			nMin = nIndex + 1;
			continue;
		}
		return nIndex;
	}
	return -1;
}

#ifdef _DEBUG
int teLSearch(TEmethod *method, int nSize, LPOLESTR bs)
{
	WCHAR pszPath[32];
	CHAR pszPath1[32], pszPath2[32];
	//	for (int j = 0; pszPath1[j] = (bs && j < _countof(pszPath1) - 1) ? tolower(bs[j]) : NULL; ++j);
	for (int j = 0; pszPath1[j] = (bs && j < _countof(pszPath1) - 1) ? bs[j] : NULL; ++j);

	for (int i = 0; i < nSize; ++i) {
		::MultiByteToWideChar(CP_UTF8, 0, method[i].name, -1, pszPath, 32);
		//		for (int j = 0; pszPath2[j] = (j < _countof(pszPath2) - 1) ? tolower(pszPath[j]) : NULL; ++j);
		//		if (lstrcmpi(bs, pszPath) == 0) {
		//		if (teStrCmpI(bs, pszPath) == 0) {
		//		if (StrCmpI(bs, pszPath) == 0) {
		//		if (teStrCmpI(bs, pszPath) == 0) {
		if (strcmp(pszPath1, method[i].name) == 0) {
			return i;
		}
	}
	return -1;
}
#endif

VOID teCreateSafeArrayFromVariantArray(IDispatch *pdisp, VARIANT *pVarResult)
{
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

#ifdef _DEBUG
VOID teCheckMethod(LPSTR name, TEmethod *method, int nCount)
{
	for (int i = 0; i < nCount - 1; ++i) {
		if (lstrcmpiA(method[i].name, method[i + 1].name) >= 0) {
			MessageBoxA(NULL, method[i].name, name, MB_OK);
		}
	}
}
#endif

HRESULT teGetDispId(TEmethod *method, int n, LPOLESTR bs, DISPID *rgDispId, BOOL bNum)
{
	if (method) {
		int nIndex = teBSearch(method, n, bs);
		if (nIndex >= 0) {
			*rgDispId = method[nIndex].id;
			return S_OK;
		}
	} else {
		CHAR pszName[32];
		for (int j = 0; pszName[j] = (bs && j < _countof(pszName) - 1) ? tolower(bs[j]) : NULL; ++j);
		auto itr = g_maps[n].find(pszName);
		if (itr != g_maps[n].end()) {
			*rgDispId = itr->second;
			return S_OK;
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
	*rgDispId = DISPID_TE_UNDEFIND;
	return DISP_E_UNKNOWNNAME;
}

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

HRESULT teException(EXCEPINFO *pExcepInfo, LPCSTR pszObjA, TEmethod* pMethod, int nCount, DISPID dispIdMember)
{
	LPSTR pszNameA = NULL;
	if (pMethod) {
		for (int i = 0; i < nCount; ++i) {
			if (pMethod[i].id == dispIdMember) {
				pszNameA = pMethod[i].name;
			}
		}
	}
	return teExceptionEx(pExcepInfo, pszObjA, pszNameA);
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

LPITEMIDLIST teSHSimpleIDListFromPathEx(LPWSTR lpstr, DWORD dwAttr, DWORD nSizeLow, DWORD nSizeHigh, FILETIME *pft)
{
	if (!teIsFileSystem(lpstr)) {
		return teIsSearchFolder(lpstr) ? NULL : ::SHSimpleIDListFromPath(lpstr);
	}
	LPITEMIDLIST pidl = NULL;
	IBindCtx *pbc = NULL;
	ULONG chEaten;
	ULONG dwAttributes;
	IShellFolder *pSF;
	if SUCCEEDED(SHGetDesktopFolder(&pSF)) {
		if SUCCEEDED(CreateBindCtx(0, &pbc)) {
			pbc->RegisterObjectParam(STR_PARSE_PREFER_FOLDER_BROWSING, teFindDropSource());
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

VOID GetFolderItemFromVariant(FolderItem **ppid, VARIANT *pv)
{
	IUnknown *punk;
	if (!FindUnknown(pv, &punk) || FAILED(punk->QueryInterface(IID_PPV_ARGS(ppid)))) {
		*ppid = new CteFolderItem(pv);
	}
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

BOOL teILIsEqual(IUnknown *punk1, IUnknown *punk2)
{
	BOOL bResult = FALSE;
	if (punk1 && punk2) {
		BSTR bs1, bs2;
		teGetPath(&bs1, punk1);
		teGetPath(&bs2, punk2);
		bResult = teStrCmpI(bs1, bs2) == 0;
		if (bResult) {
			DWORD dwUnavailable = 0;
			CteFolderItem *pid;
			if SUCCEEDED(punk1->QueryInterface(g_ClsIdFI, (LPVOID *)&pid)) {
				dwUnavailable = pid->m_dwUnavailable;
				pid->Release();
			}
			if SUCCEEDED(punk2->QueryInterface(g_ClsIdFI, (LPVOID *)&pid)) {
				dwUnavailable |= pid->m_dwUnavailable;
				pid->Release();
			}
			if (!dwUnavailable) {
				LPITEMIDLIST pidl1 = NULL;
				if (teGetIDListFromObject(punk1, &pidl1)) {
					LPITEMIDLIST pidl2 = NULL;
					if (teGetIDListFromObject(punk2, &pidl2)) {
						bResult = ::ILIsEqual(pidl1, pidl2);
					}
					teILFreeClear(&pidl2);
				}
				teILFreeClear(&pidl1);
			}
		}
		teSysFreeString(&bs2);
		teSysFreeString(&bs1);
	}
	return bResult;
}

VOID teGetPath(BSTR *pbs, IUnknown *punk)
{
	*pbs = NULL;
	FolderItem *pid;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pid))) {
		pid->get_Path(pbs);
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

VOID teQueryFolderItemReplace(FolderItem **ppfi, CteFolderItem **ppid)
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

VOID GetDragItems(CteFolderItems **ppDragItems, IDataObject *pDataObj)
{
	SafeRelease(ppDragItems);
	*ppDragItems = new CteFolderItems(pDataObj, NULL);
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

BOOL teIsClan(HWND hwndRoot, HWND hwnd)
{
	while (hwnd != hwndRoot) {
		hwnd = GetParent(hwnd);
		if (!hwnd) {
			return FALSE;
		}
	}
	return TRUE;
}

#endif
