#pragma once

#include <Mshtml.h>
#include <mshtmhst.h>
#include "mshtmdid.h"
#include <commdlg.h>
#include <commctrl.h>
#include <Shlobj.h>
#include <Shellapi.h>
#include <shobjidl.h>
#include <Shlwapi.h>
#include <exdispid.h>
#include <shdispid.h>
#include <Commdlg.h>
#include <Dbt.h>
#include <propkey.h>
#include <process.h>
#include <Wincrypt.h>
#include <ActivScp.h>
#include <windowsx.h>
#include <CommonControls.h>
#include <UIAutomationClient.h>
#include <UIAutomationCore.h>
#include <Uxtheme.h>
#include <wincodec.h>
#include <wincodecsdk.h>
#include <Thumbcache.h>
#include <mmsystem.h>
#include <WinIoCtl.h>
#include <tlhelp32.h>
#include <structuredquery.h>
//#include <Vssym32.h>
#include <vector>
#include <unordered_map>
#define Count_WICFunc			1

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
	DWORD			m_dwSessionId;
private:
	LONG			m_cRef;
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
	VOID CreateDataObj();
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
	//HWND	m_hwnd;
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


#ifdef USE_HTMLDOC
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
#ifdef USE_TESTOBJECT
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
