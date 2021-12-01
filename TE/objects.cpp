#include "stdafx.h"
#include "common.h"
#if defined(_WINDLL) || defined(_EXEONLY)

#include "objects.h"

#ifndef IID_IWICBitmap
const IID IID_IWICBitmap                    = {0x00000121, 0xa8f2, 0x4877, { 0xba, 0x0a, 0xfd, 0x2b, 0x66, 0x45, 0xfb, 0x94}};
#endif

extern HWND g_hwndMain;
extern int g_nException;
extern BOOL	g_bUpperVista;
extern BOOL	g_bUpper10;
extern DWORD	g_dwMainThreadId;
extern IDispatch	*g_pOnFunc[Count_OnFunc];
extern std::vector<LONG_PTR> g_ppGetArchive;
extern std::vector<LONG_PTR> g_ppGetImage;
extern TEStruct pTEStructs[];
extern LPITEMIDLIST g_pidls[MAX_CSIDL2];
extern std::vector<CteFolderItem *> g_ppENum;
extern GUID		g_ClsIdFI;
extern GUID		g_ClsIdStruct;
extern BSTR	g_bsClipRoot;
extern BOOL	g_bDragging;
extern BOOL	g_bDragIcon;
extern IDropTargetHelper *g_pDropTargetHelper;

extern FORMATETC HDROPFormat;
extern FORMATETC IDLISTFormat;
extern FORMATETC DROPEFFECTFormat;
extern FORMATETC UNICODEFormat;
extern FORMATETC TEXTFormat;
extern FORMATETC DRAGWINDOWFormat;
extern FORMATETC IsShowingLayeredFormat;

#ifdef _DEBUG
extern LPWSTR	g_strException;
#endif
#ifdef _2000XP
extern LPFNPSPropertyKeyFromString _PSPropertyKeyFromString;
extern LPFNPSGetPropertyKeyFromName _PSGetPropertyKeyFromName;
extern LPFNPSPropertyKeyFromString _PSPropertyKeyFromStringEx;
extern LPFNPSGetPropertyDescription _PSGetPropertyDescription;
extern LPFNPSStringFromPropertyKey _PSStringFromPropertyKey;
extern LPFNPropVariantToVariant _PropVariantToVariant;
extern LPFNVariantToPropVariant _VariantToPropVariant;
extern LPFNSHCreateItemFromIDList _SHCreateItemFromIDList;
extern LPFNSHGetIDListFromObject _SHGetIDListFromObject;
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
extern IDropSource* teFindDropSource();

OPENFILENAME *g_pofn = NULL;
BOOL	g_bDialogOk = FALSE;
long	g_nProcFI      = 0;

TEmethod methodGB[] = {
	{ TE_METHOD + 1, "FromHBITMAP" },
	{ TE_METHOD + 2, "FromHICON" },
	{ TE_METHOD + 3, "FromResource" },
	{ TE_METHOD + 4, "FromFile" },
	{ TE_METHOD + 5, "FromStream" },
	{ TE_METHOD + 6, "FromArchive" },
	{ TE_METHOD + 7, "FromItem" },//Deprecated
	{ TE_METHOD + 8, "FromClipboard" },
	{ TE_METHOD + 9, "FromSource" },
	{ TE_METHOD + 90, "Create" },
	{ TE_METHOD + 99, "Free" },

	{ TE_METHOD + 100, "Save" },
	{ TE_METHOD + 101, "Base64" },
	{ TE_METHOD + 102, "DataURI" },
	{ TE_METHOD + 103, "GetStream" },
	{ TE_METHOD + 110, "GetWidth" },
	{ TE_METHOD + 111, "GetHeight" },
	{ TE_METHOD + 112, "GetPixel" },
	{ TE_METHOD + 113, "SetPixel" },
	{ TE_METHOD + 114, "GetPixelFormat" },
	{ TE_METHOD + 115, "FillRect" },
	{ TE_METHOD + 116, "Mask" },
	{ TE_METHOD + 117, "AlphaBlend" },
	{ TE_METHOD + 120, "GetThumbnailImage" },
	{ TE_METHOD + 130, "RotateFlip" },
	{ TE_METHOD + 140, "GetFrameCount" },
	{ TE_PROPERTY + 150, "Frame" },
	{ TE_METHOD + 160, "GetMetadata" },
	{ TE_METHOD + 161, "GetFrameMetadata" },

	{ TE_METHOD + 210, "GetHBITMAP" },
	{ TE_METHOD + 211, "GetHICON" },
	{ TE_METHOD + 212, "DrawEx" },

	{ TE_METHOD + 900, "GetCodecInfo" },
	{ START_OnFunc + WIC_OnGetAlt, "OnGetAlt" },
};

//vvv Sort by name vvv
TEmethod methodFI[] = {
	{ 9, "_BLOB" },	//To be necessary
	{ 3, "Alt" },
	{ 6, "Enum" },
//	{ 4, "FocusedItem" },
	{ 9, "FolderItem" }, // Reserved future
	{ 7, "Id" },
	{ 1, "Name" },
	{ 2, "Path" },
	{ 5, "Unavailable" },
};

TEmethod methodFIs[] = {
	{ DISPID_NEWENUM, "_NewEnum" },
	{ TE_METHOD + 8, "AddItem" },
	{ TE_PROPERTY + 2, "Application" },
	{ DISPID_TE_COUNT, "Count" },
	{ TE_PROPERTY + 0x1003, "Data" },
	{ TE_PROPERTY + 0x1001, "dwEffect" },
	{ TE_METHOD + 10, "GetData" },
	{ TE_PROPERTY + 9, "hDrop" },
	{ DISPID_TE_INDEX, "Index" },
	{ DISPID_TE_ITEM, "Item" },
	{ DISPID_TE_COUNT, "length" },
	{ TE_PROPERTY + 3, "Parent" },
	{ TE_PROPERTY + 0x1002, "pdwEffect" },
	{ TE_METHOD + 11, "SetData" },
	{ TE_PROPERTY + 0x1004, "UseText" },
};

TEmethod methodMem[] = {
	{ TE_PROPERTY + 0xfff9, "_BLOB" },
	{ DISPID_NEWENUM, "_NewEnum" },
	{ TE_METHOD + 8, "Clone" },
	{ DISPID_TE_COUNT, "Count" },
	{ TE_METHOD + 7, "Free" },
	{ DISPID_TE_ITEM, "Item" },
	{ DISPID_TE_COUNT, "length" },
	{ TE_PROPERTY + 0xfff1, "P" },
	{ TE_METHOD + 4, "Read" },
	{ TE_PROPERTY + 0xfff6, "Size" },
	{ TE_METHOD + 5, "Write" },
};

TEmethod methodMem2[] = {
	{ VT_UI1 << TE_VT, "BYTE" },
	{ VT_UI4 << TE_VT, "DWORD" },
	{ VT_PTR << TE_VT, "HANDLE" },
	{ VT_I4  << TE_VT, "int" },
	{ VT_PTR << TE_VT, "LPWSTR" },
	{ VT_UI2 << TE_VT, "WCHAR" },
	{ VT_UI2 << TE_VT, "WORD" },
};

TEmethod methodDT[] = {
	{ TE_METHOD + 1, "DragEnter" },
	{ TE_METHOD + 4, "DragLeave" },
	{ TE_METHOD + 2, "DragOver" },
	{ TE_METHOD + 3, "Drop" },
	{ TE_PROPERTY + 6, "FolderItem" },
	{ TE_PROPERTY + 5, "Type" },
};

TEmethod methodCD[] = {
	{ TE_PROPERTY + 21, "DefExt" },
	{ TE_PROPERTY + 10, "FileName" },
	{ TE_PROPERTY + 13, "Filter" },
	{ TE_PROPERTY + 32, "FilterIndex" },
	{ TE_PROPERTY + 31, "Flags" },
	{ TE_PROPERTY + 31, "FlagsEx" },
	{ TE_PROPERTY + 20, "InitDir" },
	{ TE_PROPERTY + 30, "MaxFileSize" },
	{ TE_METHOD + 40, "ShowOpen" },
	{ TE_METHOD + 41, "ShowSave" },
	{ TE_PROPERTY + 22, "Title" },
};

TEmethod methodPD[] = {
	{ TE_METHOD + 1, "HasUserCancelled" },
	{ TE_METHOD + 9, "SetAnimation" },
	{ TE_METHOD + 2, "SetCancelMsg" },
	{ TE_METHOD + 3, "SetLine" },
	{ TE_METHOD + 4, "SetProgress" },
	{ TE_METHOD + 5, "SetTitle" },
	{ TE_METHOD + 6, "StartProgressDialog" },
	{ TE_METHOD + 7, "StopProgressDialog" },
	{ TE_METHOD + 8, "Timer" },
};

TEmethod methodCO[] = {
	{ TE_METHOD + 1, "Free" },
};

TEmethod methodEN[] = {
	{ TE_METHOD + 1, "atEnd" },
	{ TE_METHOD + 2, "item" },
	{ TE_METHOD + 3, "moveFirst" },
	{ TE_METHOD + 4, "moveNext" },
};


VOID InitObjects()
{
#ifdef _DEBUG
	//Check orders
	teCheckMethod("methodFI", methodFI, _countof(methodFI));
	teCheckMethod("methodFIs", methodFIs, _countof(methodFIs));
	teCheckMethod("methodMem", methodMem, _countof(methodMem));
	teCheckMethod("methodMem2", methodMem2, _countof(methodMem2));
	teCheckMethod("methodDT", methodDT, _countof(methodDT));
	teCheckMethod("methodCD", methodCD, _countof(methodCD));
	teCheckMethod("methodPD", methodPD, _countof(methodPD));
	teCheckMethod("methodCO", methodCO, _countof(methodCO));
	teCheckMethod("methodEN", methodEN, _countof(methodEN));
#endif
	teInitUM(MAP_GB, methodGB, _countof(methodGB));
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
		return teStrCmpI(pszText, pszFind) == 0;
	}
	return FALSE;
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
	m_dwSessionId = MAKELONG(GetTickCount(), rand());
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
	for (size_t i = 0; i < g_ppENum.size(); ++i) {
		if (g_ppENum[i] == this) {
			g_ppENum.erase(g_ppENum.begin() + i);
			break;
		}
	}
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
			if (m_dwUnavailable || !teGetIDListFromVariant(&m_pidl, &m_v, TRUE)) {
				if (m_v.vt == VT_BSTR) {
					m_pidl = teSHSimpleIDListFromPathEx(m_v.bstrVal, FILE_ATTRIBUTE_HIDDEN, -1, -1, NULL);
				}
				if (!m_pidlAlt && !m_dwUnavailable) {
					m_dwUnavailable = GetTickCount();
				}
			}
		}
	}
	return m_pidl ? m_pidl : m_pidlAlt ? m_pidlAlt : g_pidls[CSIDL_DESKTOP];
}

LPITEMIDLIST CteFolderItem::GetAlt()
{
	if (m_pidlAlt == NULL) {
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
	SafeRelease(&m_pEnum);
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	HRESULT hr = teGetDispId(methodFI, _countof(methodFI), *rgszNames, rgDispId, FALSE);
	if SUCCEEDED(hr) {
		return S_OK;
	}
	if (GetFolderItem()) {
		hr = m_pFolderItem->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
		if SUCCEEDED(hr) {
			if (*rgDispId > 0) {
				if (teStrCmpIWA(*rgszNames, "ExtendedProperty") == 0) {
					m_dispExtendedProperty = *rgDispId;
				}
				*rgDispId += 10;
			}
		} else {
			*rgDispId = DISPID_TE_UNDEFIND;
		}
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
			if (wFlags & DISPATCH_PROPERTYPUT) {
				for (size_t i = 0; i < g_ppENum.size(); ++i) {
					if (g_ppENum[i] == this) {
						return S_OK;
					}
				}
				g_ppENum.push_back(this);
			}
			return S_OK;

		case 7://Id
			teSetLong(pVarResult, m_dwSessionId);
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

		case 9://_BLOB ** To be necessary
			return S_OK;

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
		return teException(pExcepInfo, "FolderItem", methodFI, _countof(methodFI), dispIdMember);
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
						teQueryFolderItemReplace(&pid, &pid1);
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	teGetDispId(methodFIs, _countof(methodFIs), *rgszNames, rgDispId, TRUE);
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
		return teException(pExcepInfo, "FolderItems", methodFIs, _countof(methodFIs), dispIdMember);
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
/*	if (g_pWebBrowser) {
		return g_pWebBrowser->m_pWebBrowser->get_Application(ppid);
	}*/////
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
	return teGetDispId(methodFIs, _countof(methodFIs), bstrName, pid, TRUE);
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
							teSetObject(&pv[2], teFindDropSource());
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
					teSetObject(&pv[2], teFindDropSource());
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

//CteMemory
CteMemory::CteMemory(int nSize, void *pc, int nCount, LPWSTR lpStruct)
{
	BOOL bSafeArray = FALSE;
	m_cRef = 1;
	m_pc = (char *)pc;
	m_bsStruct = NULL;
	m_nStructIndex = -1;
	if (teStrCmpIWA(lpStruct, "SAFEARRAY") == 0) {
		lpStruct = NULL;
		bSafeArray = TRUE;
	}
	if (lpStruct) {
		m_nStructIndex = teUMSearch(MAP_SS, lpStruct);
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
					if (pTEStructs[m_nStructIndex].bSize) {
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
		return teException(pExcepInfo, "Memory", methodMem, _countof(methodMem), dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteMemory::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	if (m_nStructIndex >= 0 && teGetDispId(pTEStructs[m_nStructIndex].pMethod, pTEStructs[m_nStructIndex].nCount, bstrName, pid, FALSE) == S_OK) {
		return S_OK;
	}
	if (teGetDispId(methodMem, _countof(methodMem), bstrName, pid, TRUE) == S_OK) {
		return S_OK;
	}
	return teGetDispId(methodMem2, _countof(methodMem2), m_bsStruct, pid, FALSE);
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
					(teStrCmpIWA(m_bsStruct, "BSTR") == 0 || teStrCmpIWA(m_bsStruct, "LPWSTR") == 0)) {
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
					} else if (nSize == 8 && teStrCmpIWA(m_bsStruct, "POINT")) {//LONGLONG
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
#ifdef _DEBUG
	//::OutputDebugStringA("WICBitmap: ");
	//::OutputDebugString(*rgszNames);
	//::OutputDebugStringA("\n");
#endif
	teGetDispId(NULL, MAP_GB, *rgszNames, rgDispId, FALSE);
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
			
		case TE_METHOD + 3://FromResource
			if (nArg >= 5) {
				HBITMAP hBM = (HBITMAP)LoadImage((HINSTANCE)GetPtrFromVariant(&pDispParams->rgvarg[nArg]), GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 3]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 4]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 5]));
				if (hBM) {
					CreateBitmapFromHBITMAP(hBM, 0, 3);
					::DeleteObject(hBM);
				}
				teSetObject(pVarResult, GetBitmapObj());
			}
			return S_OK;
			
		case TE_METHOD + 4://FromFile
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
#ifdef _DEBUG
					::OutputDebugStringA("WICBitmap.FromFile: ");
					::OutputDebugString(lpfn);
					::OutputDebugStringA("\n");
#endif
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
		return teException(pExcepInfo, "WICBitmap", methodGB, _countof(methodGB), dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

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
		return teException(pExcepInfo, "Dispatch", NULL, 0, dispIdMember);
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	if (teStrCmpIWA(bstrName, "Count") == 0) {
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;

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
	teGetDispId(methodDT, _countof(methodDT), *rgszNames, rgDispId, FALSE);
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
		return teException(pExcepInfo, "DropTarget", methodDT, _countof(methodDT), dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteDropTarget2
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
	if (g_pDropTargetHelper) {
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	teGetDispId(methodCO, _countof(methodCO), *rgszNames, rgDispId, FALSE);
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	teGetDispId(methodCD, _countof(methodCD), *rgszNames, rgDispId, FALSE);
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

			case TE_METHOD + 40://ShowOpen
				if (m_ofn.Flags & OFN_ENABLEHOOK) {
					g_pofn = &m_ofn;
					m_ofn.Flags &= ~OFN_ENABLEHOOK;
					/*
					IFileOpenDialog *pFileOpenDialog;
					if (FALSE && _SHCreateItemFromIDList && SUCCEEDED(teCreateInstance(CLSID_FileOpenDialog, NULL, NULL, IID_PPV_ARGS(&pFileOpenDialog)))) {
					IShellItem *psi;
					if (teCreateItemFromPath(const_cast<LPWSTR>(m_ofn.lpstrInitialDir), &psi)) {
					pFileOpenDialog->SetFolder(psi);
					psi->Release();
					}
					pFileOpenDialog->SetFileName(m_ofn.lpstrFile);
					pFileOpenDialog->SetOptions(m_ofn.Flags & (FOS_OVERWRITEPROMPT	| FOS_NOCHANGEDIR | FOS_PICKFOLDERS | FOS_NOVALIDATE | FOS_ALLOWMULTISELECT | FOS_PATHMUSTEXIST | FOS_FILEMUSTEXIST | FOS_CREATEPROMPT | FOS_SHAREAWARE | FOS_NOREADONLYRETURN));
					if SUCCEEDED(pFileOpenDialog->Show(g_hwndMain)) {
					if SUCCEEDED(pFileOpenDialog->GetResult(&psi)) {
					LPWSTR pszPath;
					if SUCCEEDED(psi->GetDisplayName(SIGDN_DESKTOPABSOLUTEEDITING, &pszPath)) {
					lstrcpyn(m_ofn.lpstrFile, pszPath, m_ofn.nMaxFile);
					::CoTaskMemFree(pszPath);
					bResult = TRUE;
					}
					psi->Release();
					}
					}
					pFileOpenDialog->Release();
					break;
					}
					m_ofn.lpfnHook = OFNHookProc;*/
				}
				g_bDialogOk = FALSE;
				bResult = GetOpenFileName(&m_ofn) || g_bDialogOk;
				g_pofn = NULL;
				break;

			case TE_METHOD + 41://ShowSave
				bResult = GetSaveFileName(&m_ofn);
				break;
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
		return teException(pExcepInfo, "CommonDialog", methodCD, _countof(methodCD), dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteProgressDialog

CteProgressDialog::CteProgressDialog(IProgressDialog *ppd)
{
	m_cRef = 1;
	m_ppd = NULL;
	//	m_hwnd = NULL;
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	teGetDispId(methodPD, _countof(methodPD), *rgszNames, rgDispId, FALSE);
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
		return teException(pExcepInfo, "ProgressDialog", methodPD, _countof(methodPD), dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteServiceProvider
STDMETHODIMP CteServiceProvider::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteServiceProvider, IServiceProvider),
	{ 0 },
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
		QITABENT(CteEnumerator, IDispatch),
	{ 0 },
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	teGetDispId(methodEN, _countof(methodEN), *rgszNames, rgDispId, FALSE);
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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

#endif
