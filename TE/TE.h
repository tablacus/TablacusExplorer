#pragma once

//#define _USE_BSEARCHAPI
//#define _USE_APIHOOK
//#define _USE_HTMLDOC
//#define _USE_TESTOBJECT
//#define _USE_TESTPATHMATCHSPEC
#define _Emulate_XP_	//FALSE &&
#ifndef _WIN64
#define _2000XP
//#define _W2000
#endif
//bug fix 1 for Windows 10 Insider Preview 14986-
#define _FIXWIN10IPBUG1

#include "resource.h"
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
#ifndef _2000XP
#include <Propsys.h>
#endif
#ifdef _USE_APIHOOK
#include <imagehlp.h>
#endif

#import <shdocvw.dll> exclude("OLECMDID", "OLECMDF", "OLECMDEXECOPT", "tagREADYSTATE") auto_rename
//#import <mshtml.tlb>
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "urlmon.lib")
#pragma comment(lib, "imm32.lib")
#pragma comment(lib, "crypt32.lib")
#pragma comment(lib, "UxTheme.lib")
#pragma comment(lib, "Msimg32.lib")
#ifndef _2000XP
#pragma comment(lib, "Propsys.lib")
#endif
#ifdef _USE_APIHOOK
#pragma comment(lib, "imagehlp.lib")
#endif
#ifdef _WIN64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#if _MSC_VER >=1600
#pragma execution_character_set("utf-8")
#define CP_TE CP_UTF8
#else
#define CP_TE CP_ACP
#endif

class CteShellBrowser;
class CteTreeView;

union teParam
{
    LONGLONG llVal;
    LONG lVal;
    ULONG ulVal;
    INT intVal;
    UINT uintVal;
	BYTE bVal;
	SHORT iVal;
	BSTR bstrVal;
    VARIANT_BOOL boolVal;
	FLOAT fltVal;
	DOUBLE dblVal;

	BYTE *pbVal;
    CHAR *pcVal;
    INT *pintVal;

	WORD word;
	DWORD dword;
	WPARAM wparam;
	LPARAM lparam;
	COLORREF colorref;
	LONG_PTR long_ptr;
	UINT_PTR uint_ptr;
	ULONG_PTR ulong_ptr;
	LARGE_INTEGER li;
	ULARGE_INTEGER uli;

	HANDLE handle;
	HWND hwnd;
	HKEY hkeyVal;
	HICON hicon;
	HBITMAP hbitmap;
	HDROP hdrop;
	HDC hdc;
	HMENU hmenu;
	HMODULE hmodule;
	HIMAGELIST himagelist;
	HGDIOBJ hgdiobj;
	HBRUSH hbrush;
	HCURSOR hcursor;
	HIMC himc;
	HRGN hrgn;
	HINSTANCE hinstance;
	HMONITOR hmonitor;
	HICON *phicon;

	LPWSTR lpwstr;
	LPCWSTR lpcwstr;
	LPOLESTR lpolestr;
	LPCOLESTR lpcolestr;
	LPSTR lpstr;
	LPCSTR lpcstr;

	OLECMDID olecmdid;
	OLECMDEXECOPT olecmdexecopt;

	PLOGFONT plogfont;
	PICONINFO piconinfo;
	PNOTIFYICONDATA pnotifyicondata;
	PCHANGEFILTERSTRUCT pchangefilterstruct;
	WNDPROC wndproc;

	LPVOID lpvoid;
	LPBOOL lpbool;
	LPDWORD lpdword;
	LPHANDLE lphandle;
	LPSHFILEOPSTRUCT lpshfileopstruct;
	LPSHELLEXECUTEINFO lpshellexecuteinfo;
	LPMONITORINFO lpmonitorinfo;
	SHFILEINFO *lpshfileinfo;
	LPPOINT lppoint;
	LPMSG lpmsg;
	LPSIZE lpsize;
	LPRECT lprect;
	LPCRECT lpcrect;
	LPPAINTSTRUCT lppaintstruct;
	LPFINDREPLACE lpfindreplace;
	LPWIN32_FIND_DATA lpwin32_find_data;
	LPCHOOSECOLOR lpchoosecolor;
	LPCHOOSEFONT lpchoosefont;
	LPINPUT lpinput;
	LPTPMPARAMS lptpmparams;
	LPITEMIDLIST lpitemidlist;
	LPMENUINFO lpmenuinfo;
	LPCMENUINFO lpcmenuinfo;
	LPMENUITEMINFO lpmenuiteminfo;
	LPCMENUITEMINFO lpcmenuiteminfo;
	LPOSVERSIONINFO lposversioninfo;
	PRTL_OSVERSIONINFOEXW prtl_osversioninfoexw;
	PGUITHREADINFO pgui;

	ATOM atom;
	ASSOCF assocf;
	ASSOCSTR assocstr;
	LCID lcid;
	LCTYPE lctype;
	PROPDESC_FORMAT_FLAGS propdesc_format_flags;
	PROPERTYUI_FORMAT_FLAGS propertyui_format_flags;
	FILEOP_FLAGS fileop_flags;

	IImageList *pImageList;
	IID iid;
	VARIANT variant;
	EXCEPINFO *pExcepInfo;
};

//Unnamed function
typedef VOID (WINAPI * LPFNSHRunDialog)(HWND hwnd, HICON hIcon, LPWSTR pszPath, LPWSTR pszTitle, LPWSTR pszPrompt, DWORD dwFlags);

//Closed function
typedef BOOL (WINAPI * LPFNRegenerateUserEnvironment)(LPVOID *lpEnvironment, BOOL bUpdate);

#ifdef _2000XP

//XP SP1 or higher.
typedef BOOL (WINAPI* LPFNSetDllDirectoryW)(__in_opt LPCWSTR lpPathName);

//XP SP2 or higher.
typedef BOOL (WINAPI* LPFNIsWow64Process)(HANDLE hProcess, PBOOL Wow64Process);

//XP SP2 or higher with Windows Desktop Search.
typedef HRESULT (STDAPICALLTYPE* LPFNPSPropertyKeyFromString)(__in LPCWSTR pszString,  __out PROPERTYKEY *pkey);
typedef HRESULT (STDAPICALLTYPE* LPFNPSGetPropertyKeyFromName)(__in PCWSTR pszName, __out PROPERTYKEY *ppropkey);
typedef HRESULT (STDAPICALLTYPE* LPFNPSGetPropertyDescription)(__in REFPROPERTYKEY propkey, __in REFIID riid,  __deref_out void **ppv);
typedef HRESULT (STDAPICALLTYPE* LPFNPSStringFromPropertyKey)(__in REFPROPERTYKEY pkey, __out_ecount(cch) LPWSTR psz, __in UINT cch);

//Vista or higher.
typedef HRESULT (STDAPICALLTYPE* LPFNSHCreateItemFromIDList)(__in PCIDLIST_ABSOLUTE pidl, __in REFIID riid, __deref_out void **ppv);
typedef HRESULT (STDAPICALLTYPE* LPFNSHGetIDListFromObject)(__in IUnknown *punk, __deref_out PIDLIST_ABSOLUTE *ppidl);
typedef BOOL (WINAPI* LPFNChangeWindowMessageFilter)(__in UINT message, __in DWORD dwFlag);
typedef BOOL (WINAPI* LPFNAddClipboardFormatListener)(__in HWND hwnd);
typedef BOOL (WINAPI* LPFNRemoveClipboardFormatListener)(__in HWND hwnd);
//typedef HRESULT (STDAPICALLTYPE * LPFNPSFormatForDisplayAlloc)(__in REFPROPERTYKEY key, __in REFPROPVARIANT propvar, __in PROPDESC_FORMAT_FLAGS pdff, __deref_out PWSTR *ppszDisplay);
//typedef BOOL (WINAPI * LPFNChangeWindowMessageFilter)(UINT message, DWORD dwFlag);

#endif

//7 or higher
typedef BOOL (WINAPI* LPFNChangeWindowMessageFilterEx)(__in HWND hwnd, __in UINT message, __in DWORD action, __inout_opt PCHANGEFILTERSTRUCT pChangeFilterStruct);

//8 or higher
typedef BOOL (WINAPI* LPFNSetDefaultDllDirectories)(__in DWORD DirectoryFlags);

//RTL
typedef NTSTATUS (WINAPI* LPFNRtlGetVersion)(PRTL_OSVERSIONINFOEXW lpVersionInformation);
//DLL
typedef HRESULT (STDAPICALLTYPE * LPFNDllGetClassObject)(REFCLSID rclsid, REFIID riid, LPVOID* ppv);
typedef HRESULT (STDAPICALLTYPE * LPFNDllCanUnloadNow)(void);

#ifdef _USE_APIHOOK
//API Hook
typedef LSTATUS (APIENTRY* LPFNRegQueryValueExW)(HKEY hKey, LPCWSTR lpValueName, LPDWORD lpReserved, LPDWORD lpType, LPBYTE lpData, LPDWORD lpcbData);
#endif

#ifndef LOAD_LIBRARY_SEARCH_SYSTEM32
#define LOAD_LIBRARY_SEARCH_SYSTEM32 0x00000800
#endif
#ifndef LOAD_LIBRARY_SEARCH_USER_DIRS
#define LOAD_LIBRARY_SEARCH_USER_DIRS 0x00000400
#endif

	
//Tablacus DLL Add-ons
typedef VOID (WINAPI * LPFNGetProcObjectW)(VARIANT *pVarResult);

//Dispatch Windows API
typedef VOID (__cdecl * LPFNDispatchAPI)(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult);

#define DISPID_AMBIENT_OFFLINEIFNOTCONNECTED -5501
#define E_CANCELLED         HRESULT_FROM_WIN32(ERROR_CANCELLED)
#define E_FILE_NOT_FOUND    HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)
#define E_PATH_NOT_FOUND    HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND)
#define E_NOT_READY         HRESULT_FROM_WIN32(ERROR_NOT_READY)
#define E_BAD_NETPATH       HRESULT_FROM_WIN32(ERROR_BAD_NETPATH)
#define E_INVALID_PASSWORD  HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD)

#define WINDOW_CLASS			L"TablacusExplorer"
#define WINDOW_CLASS2			L"TablacusExplorer2"
#define MAX_LOADSTRING			100
#define MAX_HISTORY				32
#define MAX_CSIDL				256
#define MAX_FORMATS				1024
#define MAX_TC					128
#define MAX_TV					128
#define MAX_FV					1024
#define MAX_PATHEX				32768
#define MAX_PROP				4096
#define MAX_STATUS				1024
#define MAX_COLUMNS				8192
#define SIZE_BUFF				32768
#define TET_Create				0x1fa1
#define TET_Reload				0x1fa2
#define TET_Size				0x1fa3
#define TET_Redraw				0x1fa4
#define TET_Status				0x1fa5
#define TET_Unload				0x1fa6
#define TET_Title				0x1fa7
#define TET_FreeLibrary			0x1fa8
#define TET_Refresh				0x1fa9
#define TET_EndThread			0x1faa
#define SHGDN_FORPARSINGEX	0x80000000
#define START_OnFunc			5000
#define TE_Labels				0
#define TE_OnMouseMessage		1
#define TE_OnViewCreated		2
#define TE_OnDefaultCommand		3
#define TE_OnCreate				4
#define TE_OnDragEnter			5
#define TE_OnDragOver			6
#define TE_OnDrop				7
#define TE_OnDragLeave			8
#define TE_OnBeforeNavigate		9
#define TE_OnItemClick			10
#define TE_OnGetPaneState		11
#define TE_OnMenuMessage		12
#define TE_OnSystemMessage		13
#define TE_OnShowContextMenu	14
#define TE_OnSelectionChanged	15
#define TE_OnClose				16
#define TE_OnVisibleChanged		17
#define TE_OnAppMessage			18
#define TE_OnStatusText			19
#define TE_OnToolTip			20
#define TE_OnNewWindow			21
#define	TE_OnWindowRegistered	22
#define TE_OnSelectionChanging	23
#define TE_OnClipboardText		24
#define TE_OnCommand			25
#define TE_OnInvokeCommand		26
#define TE_OnArrange			27
#define TE_OnHitTest			28
#define TE_OnTranslatePath		29
#define TE_OnNavigateComplete	30
#define TE_OnILGetParent		31
#define TE_OnViewModeChanged	32
#define TE_OnColumnsChanged		33
#define TE_OnKeyMessage			34
#define TE_OnItemPrePaint		35
#define TE_OnColumnClick		36
#define TE_OnBeginDrag			37
#define TE_OnBeforeGetData		38
#define TE_OnIconSizeChanged	39
#define TE_OnFilterChanged		40
#define TE_OnBeginLabelEdit		41
#define TE_OnEndLabelEdit		42
#define TE_OnReplacePath		43
#define TE_OnBeginNavigate		44
#define TE_OnSort				45
#define TE_OnFromFile			46
#define TE_OnFromStream			47
#define TE_OnEndThread			48
#define TE_OnItemPostPaint		49
#define TE_OnHandleIcon			50
#define Count_OnFunc			51
#define SB_TotalFileSize		0
#define SB_OnIncludeObject		1
#define SB_AltSelectedItems		2
#define Count_SBFunc			3

#define CTRL_FV          0
#define CTRL_SB          1
#define CTRL_EB          2
#define CTRL_TE    0x10000
#define CTRL_WB    0x20000
#define CTRL_SW    0x20001
#define CTRL_AR    0x2ffff
#define CTRL_TC    0x30000
#define CTRL_TV    0x40000
#define CTRL_DT    0x80000

#define TE_VT 24
#define TE_VI 0xffffff
#define TE_METHOD		0x40000000
#define TE_OFFSET		0x30000000
#define DISPID_TE_ITEM  0x3fffffff
#define DISPID_TE_COUNT 0x3ffffffe
#define DISPID_TE_INDEX 0x3ffffffd
#define DISPID_TE_MAX TE_VI

#define TE_Type		0
#define TE_Left		1
#define TE_Top		2
#define TE_Width	3
#define TE_Right	3
#define TE_Height	4
#define TE_Bottom	4
#define TE_Flags	5
#define TE_Tab		5
#define TE_CmdShow	6
#define TE_Layout	7
#define TE_NetworkTimeout	8
#define Count_TE_params	9

#define TC_Align	6
#define TC_TabWidth		7
#define TC_TabHeight	8

#define SB_Type			0
#define SB_ViewMode		1
#define SB_FolderFlags	2
#define	SB_Options		3
#define	SB_ViewFlags	4
#define SB_IconSize		5

#define SB_TreeAlign	6
#define SB_TreeWidth	7
#define SB_TreeFlags	8
#define SB_EnumFlags	9
#define SB_RootStyle	10
#define SB_Root			11

#define	SB_SizeFormat	12

#define	SB_DoFunc	13
#define	SB_Count	14

#define	TVVERBS 16

#define MAP_TE	0
#define MAP_SB	1
#define MAP_TC	2
#define MAP_TV	3
#define MAP_CD	4
#define MAP_SS	5
#define MAP_Mem	6
#define MAP_GB	7
#define MAP_FIs	8
#ifdef _USE_BSEARCHAPI
#define MAP_API	9
#define MAP_LENGTH	10
#else
#define MAP_LENGTH	9

#define TE_API_RESULT	12
#define TE_EXCEPINFO	13

#endif
#ifndef offsetof
#define offsetof(type, member)	(DWORD)(UINT_PTR)&((type *)0)->member
#endif

#ifdef _WIN64
#define teSetPtr(pVar, nData)	teSetLL(pVar, (LONGLONG)nData)
#define GetPtrFromVariant(pv)	GetLLFromVariant(pv)
#define GetPtrFromVariantClear(pv)	GetLLFromVariantClear(pv)
#else
#define teSetPtr(pVar, nData)	teSetLong(pVar, (LONG)nData)
#define GetPtrFromVariant(pv)	GetIntFromVariant(pv)
#define GetPtrFromVariantClear(pv)	GetIntFromVariantClear(pv)
#endif

#ifdef _2000XP
#define CommandID_GROUP 0x7601
#endif

struct TEmethod
{
	LONG   id;
	LPSTR name;
};

struct TEStruct
{
	LONG   lSize;
	LPSTR name;
	TEmethod *pMethod;
};

struct TEColumn
{
	SHCOLSTATEF	csFlags;
	int		nWidth;
};

struct TEmethodW
{
	LONG   id;
	LPWSTR name;
};

struct TEInvoke
{
	VARIANT *pv;
	IDispatch *pdisp;
	LPITEMIDLIST	pidl;
	DISPID dispid;
	int	cArgs;
	HRESULT hr;
	LONG cRef;
	WORD wErrorHandling;
	WORD wMode;
};

struct TEDispatchApi
{
	char nArgs;
	char nStr1;
	char nStr2;
	char nStr3;
	LPSTR name;
	LPFNDispatchAPI fn;
};

struct TEFS
{
	BSTR bsName;
	BSTR bsPath;
	IStream *pStrmTotalFileSize;
	LPITEMIDLIST *ppidl;
	LPITEMIDLIST pidl;
	HWND hwnd;
};

struct TEExists
{
	LPWSTR pszPath;
	HANDLE hEvent;
	LONG cRef;
	int iUseFS;
	HRESULT hr;
};

struct TEILCreate
{
	LPWSTR pszPath;
	HANDLE hEvent;
	LPITEMIDLIST pidlResult;
	LONG cRef;
};

struct TEAddItems
{
	VARIANTARG *pv;
	IStream *pStrmSB;
	IStream *pStrmArray;
	BOOL	bDeleted;
	BOOL	bNavigateComplete;
};

const CLSID CLSID_ShellShellNameSpace = {0x2F2F1F96, 0x2BC1, 0x4b1c, { 0xBE, 0x28, 0xEA, 0x37, 0x74, 0xF4, 0x67, 0x6A}};
const CLSID CLSID_JScriptChakra       = {0x16d51579, 0xa30b, 0x4c8b, { 0xa2, 0x76, 0x0f, 0xf4, 0xdc, 0x41, 0xe7, 0x55}};

class CteDll : public IUnknown
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	CteDll(HMODULE hDll);
	~CteDll();
public:
	LONG		m_cRef;
	HMODULE		m_hDll;
};

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
	BOOL GetFolderItem();

	VARIANT			m_v;
	LPITEMIDLIST	m_pidl;
	FolderItem		*m_pFolderItem;
	LPITEMIDLIST	m_pidlFocused;
	int				m_nSelected;
	BOOL			m_bStrict;
	DWORD			m_dwUnavailable;
private:
	LONG			m_cRef;
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

	CteFolderItems(IDataObject *pDataObj, FolderItems *pFolderItems, BOOL bEditable);
	~CteFolderItems();

	HDROP GethDrop(int x, int y, BOOL fNC);
	VOID Regenerate();
	VOID ItemEx(long nIndex, VARIANT *pVarResult, VARIANT *pVarNew);
	VOID AdjustIDListEx();
	VOID Clear();
	HRESULT QueryGetData2(FORMATETC *pformatetc);
public:
	int				m_nIndex;
	DWORD			m_dwEffect;
	LPITEMIDLIST	*m_pidllist;

private:
	IDataObject		*m_pDataObj;
	FolderItems		*m_pFolderItems;
	IDispatch		*m_oFolderItems;
	BSTR			m_bsText;

	LONG			m_cRef;
	LONG			m_nCount;
	BOOL			m_bUseILF;
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

	CteDropTarget2(HWND hwnd, IUnknown *punk);
	~CteDropTarget2();
public:
	CteFolderItems *m_pDragItems;
	IDropTarget *m_pDropTarget;
	HWND        m_hwnd;
	IUnknown	*m_punk;

	DWORD	m_grfKeyState;
	HRESULT m_DragLeave;
	LONG	m_cRef;
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

	CTE(int nCmdShow);
	~CTE();
public:
	VARIANT m_vData;
	CteDropTarget2 *m_pDropTarget2;
private:
	LONG	m_cRef;
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
	BSTR	m_bstrPath;
	HWND	m_hwndParent;
	HWND	m_hwndBrowser;

	HRESULT m_DragLeave;
	BOOL	m_bRedraw;
private:
	CteFolderItems *m_pDragItems;
	IDropTarget *m_pDropTarget;
	IDispatch		*m_pExternal;
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
	VOID UnLockUpdate(BOOL bDirect);
	VOID RedrawUpdate();
	VOID Show(BOOL bVisible, BOOL bMain);
	VOID GetItem(int i, VARIANT *pVarResult);
	DWORD GetStyle();
	VOID SetItemSize();
public:
	SCROLLINFO m_si;
	HWND	m_hwnd;
	HWND	m_hwndStatic;
	HWND	m_hwndButton;

	WNDPROC	m_DefTCProc;
	WNDPROC	m_DefBTProc;
	WNDPROC	m_DefSTProc;

	DWORD	m_dwSize;
	LONG	m_nLockUpdate;
	int		m_nIndex;
	int		m_param[9];
	int		m_nTC;
	int		m_nScrollWidth;
	BOOL	m_bEmpty;
	BOOL	m_bVisible;
	BOOL	m_bRedraw;
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
#ifdef _2000XP
	public IShellFolder2, public IShellFolderViewCB,
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
#ifdef _2000XP
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
	//IShellFolderViewCB
	STDMETHODIMP MessageSFVCB(UINT uMsg, WPARAM wParam, LPARAM lParam);
#endif
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
	BOOL Navigate1(FolderItem *pFolderItem, UINT wFlags, FolderItems *pFolderItems, FolderItem *pPrevious, LPITEMIDLIST *ppidl, int nErrorHandling);
	VOID Navigate1Ex(LPOLESTR pstr, FolderItems *pFolderItems, UINT wFlags, FolderItem *pPrevious, int nErrorHandleing);
	HRESULT Navigate2(FolderItem *pFolderItem, UINT wFlags, DWORD *param, FolderItems *pFolderItems, FolderItem *pPrevious, CteShellBrowser *pHistSB);
	HRESULT Navigate3(FolderItem *pFolderItem, UINT wFlags, DWORD *param, CteShellBrowser **ppSB, FolderItems *pFolderItems);
	HRESULT NavigateEB(DWORD dwFrame);
	HRESULT OnBeforeNavigate(FolderItem *pPrevious, UINT wFlags);
//	void InitializeMenuItem(HMENU hmenu, LPTSTR lpszItemName, int nId, HMENU hmenuSub);
	VOID GetSort(BSTR* pbs, int nFormat);
	VOID SetSort(BSTR bs);
	VOID GetGroupBy(BSTR* pbs);
	VOID SetGroupBy(BSTR bs);
	HRESULT SetRedraw(BOOL bRedraw);
	VOID SetColumnsStr(BSTR bsColumns);
	BSTR GetColumnsStr(int nFormat);
	VOID GetDefaultColumns();
	VOID SaveFocusedItemToHistory();
	VOID FocusItem(BOOL bFree);
	HRESULT GetAbsPidl(LPITEMIDLIST *pidlOut, FolderItem **ppid, FolderItem *pid, UINT wFlags, FolderItems *pFolderItems, FolderItem *pPrevious, CteShellBrowser *pHistSB);
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
	VOID SetFolderSize(LPCITEMIDLIST pidl, LPWSTR szText, int cch);
	VOID SetLabel(LPCITEMIDLIST pidl, LPWSTR szText, int cch);
	HRESULT PropertyKeyFromName(BSTR bs, PROPERTYKEY *pkey);
	FOLDERVIEWOPTIONS teGetFolderViewOptions(LPITEMIDLIST pidl, UINT uViewMode);
	VOID OnNavigationComplete2();
	HRESULT BrowseToObject();
	HRESULT GetShellFolder2(LPITEMIDLIST pidl);
	VOID SetLVSettings();
	HRESULT NavigateSB(IShellView *pPreviousView, FolderItem *pPrevious);
	HRESULT CreateViewWindowEx(IShellView *pPreviousView);
	VOID SetTabName();
	HRESULT RemoveAll();
	VOID SetViewModeAndIconSize(BOOL bSetIconSize);
	VOID SetListColumnWidth();
#ifdef _2000XP
	VOID AddPathXP(CteFolderItems *pFolderItems, IShellFolderView *pSFV, int nIndex, BOOL bResultsFolder);
	int PSGetColumnIndexXP(LPWSTR pszName, int *pcxChar);
	BSTR PSGetNameXP(LPWSTR pszName, int nFormat);
	VOID AddColumnDataXP(LPWSTR pszColumns, LPWSTR pszName, int nWidth, int nFormat);
#endif
#ifdef _FIXWIN10IPBUG1
	VOID FixWin10IPBug1();
#endif
public:
	HWND		m_hwnd;
	HWND		m_hwndDV;
	HWND		m_hwndLV;
	HWND		m_hwndDT;
	CteTabCtrl		*m_pTC;
	CteTreeView	*m_pTV;
	LONG_PTR	m_DefProc;
	IShellView  *m_pShellView;
	IDispatch	*m_ppDispatch[Count_SBFunc];
	FolderItem *m_pFolderItem;
	FolderItem *m_pFolderItem1;
	IExplorerBrowser *m_pExplorerBrowser;
	LPITEMIDLIST m_pidl;
	IShellFolder2 *m_pSF2;
	SORTCOLUMN m_col;
	DWORD		m_dwNameFormat;
	BOOL	m_bSetListColumnWidth;
#ifdef _2000XP
	TEColumn	*m_pColumns;
	UINT		m_nColumns;
#endif
	DWORD		m_param[SB_Count];
	int			m_nFolderSizeIndex;
	int			m_nLabelIndex;
	int			m_nSizeIndex;
	int			m_nSB;
	int			m_nUnload;
	int			m_nFocusItem;
	DWORD		m_nOpenedType;
	DWORD		m_dwCookie;
	DWORD		m_dwSessionId;
	COLORREF	m_clrBk, m_clrTextBk, m_clrText;
	BOOL		m_bEmpty;
	BOOL		m_bInit;
	BOOL		m_bVisible;
	BOOL		m_bCheckLayout;
	BOOL		m_bRefreshLator;
	BOOL		m_bRefreshing;
	BOOL		m_bAutoVM;
private:
	VARIANT		m_vData;
	FolderItem	**m_ppLog;
	FolderItem	**m_ppFocus;
	IDispatch	*m_pDSFV;
	PROPERTYKEY *m_pDefultColumns;
	CteDropTarget2 *m_pDropTarget2;
	BSTR		m_bsFilter;

	LONG		m_cRef;
	DWORD		m_dwEventCookie;
	int			m_nLogCount;
	int			m_nLogIndex;
	int			m_nPrevLogIndex;
	UINT		m_nDefultColumns;
	LONG		m_nCreate;
	BOOL		m_bIconSize;
	LONG		m_dwUnavailable;
	BOOL		m_bNavigateComplete;
	BOOL		m_bEnableSuspend;
#ifdef _2000XP
	IShellFolderViewCB	*m_pSFVCB;
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
	BSTR	*m_ppbs;
	BSTR	m_bsStruct;
	BSTR	m_bsAlloc;

	LONG	m_cRef;
	int		m_nbs;
	int		m_nStructIndex;
};

class CteContextMenu : public IDispatch
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

	CteContextMenu(IUnknown *punk, IDataObject *pDataObj, IUnknown *pSB);
	~CteContextMenu();

	BOOL GetFolderVew(IShellBrowser **ppSB);
public:
	IContextMenu *m_pContextMenu;
	CteDll *m_pDll;
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
	public INameSpaceTreeControlEvents, public INameSpaceTreeControlCustomDraw
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
	BOOL Create(BOOL bIfVisible);
	HRESULT getSelected(IDispatch **pItem);
	VOID SetRoot();
#ifdef _2000XP
	VOID SetObjectRect();
#endif
public:
	VARIANT		m_vRoot;
	HWND        m_hwnd;
	HWND        m_hwndTV;
	INameSpaceTreeControl	*m_pNameSpaceTreeControl;
	CteShellBrowser	*m_pFV;
	BOOL	m_bEmpty;
#ifdef _2000XP
	IShellNameSpace *m_pShellNameSpace;
#endif
	WNDPROC		m_DefProc;
	WNDPROC		m_DefProc2;

	BOOL		m_bMain;
	BOOL		m_bSetRoot;
private:
	VARIANT m_vData;
	LPWSTR	lplpVerbs;
	CteDropTarget2 *m_pDropTarget2;
	HWND	m_hwndParent;
	DWORD	*m_param;
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
	VOID FromStreamRelease(IStream *pStream, LPWSTR lpfn, BOOL bExtend);
	VOID GetFrameFromStream(IStream *pStream, UINT uFrame, BOOL bInit);
	BOOL HasImage();
	CteWICBitmap* GetBitmapObj();
	VOID ClearImage(BOOL bAll);
	HBITMAP GetHBITMAP(COLORREF clBk);
	BOOL Get(WICPixelFormatGUID guidNewPF);
	HRESULT CreateStream(IStream *pStream, ULARGE_INTEGER *puliSize, CLSID encoderClsid, LONG lQuality);
//	HRESULT CreateBMPStream(IStream *pStream, ULARGE_INTEGER *puliSize, LPWSTR szMime);
private:
	IWICBitmap *m_pImage;
	IStream *m_pStream;
	CLSID m_guidSrc;
	IWICMetadataQueryReader *m_ppMetadataQueryReader[2];
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
#else
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

	CteDispatch(IDispatch *pDispatch, int nMode);
	~CteDispatch();

	VOID Clear();
public:
	IActiveScript *m_pActiveScript;
	CteDll		*m_pDll;
	DISPID		m_dispIdMember;
	int			m_nIndex;
private:
	IDispatch	*m_pDispatch;

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

	CteDispatchEx(IUnknown *punk);
	~CteDispatchEx();
private:
	IDispatchEx *m_pdex;
	LONG	m_cRef;
	DISPID	m_dispItem;
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

	CteProgressDialog();
	~CteProgressDialog();
private:
	IProgressDialog *m_ppd;
	LONG	m_cRef;
};
