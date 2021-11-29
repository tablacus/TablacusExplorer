#pragma once

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
#include <mmsystem.h>
#include <WinIoCtl.h>
#include <tlhelp32.h>
#include <structuredquery.h>
//#include <Vssym32.h>
#include <vector>
#include <unordered_map>
#ifdef _USE_TEOBJ
#include <map>
#endif
#ifndef _2000XP
#include <Propsys.h>
#include <Propvarutil.h>
#endif
#ifdef _USE_APIHOOK
#include <imagehlp.h>
#endif

//#import <shdocvw.dll> exclude("OLECMDID", "OLECMDF", "OLECMDEXECOPT", "tagREADYSTATE") auto_rename
//#import <mshtml.tlb>
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "urlmon.lib")
#pragma comment(lib, "imm32.lib")
#pragma comment(lib, "crypt32.lib")
#pragma comment(lib, "UxTheme.lib")
#pragma comment(lib, "Msimg32.lib")
#pragma comment(lib, "Winmm.lib")
#ifndef _2000XP
#pragma comment(lib, "Propsys.lib")
#endif
#ifdef _USE_APIHOOK
#pragma comment(lib, "imagehlp.lib")
#endif
#ifdef _WINDLL
#ifdef __EDG__
#define DLLEXPORT
#else
#define DLLEXPORT __pragma(comment(linker, "/EXPORT:" __FUNCTION__ "=" __FUNCDNAME__))
#endif

#ifdef _WIN64
//#pragma comment(linker, "/DEF:\"te64.def\"")
#else
//#pragma comment(linker, "/DEF:\"te32.def\"")
#endif
#endif
#if !defined(_WINDLL) || defined(_EXEONLY)
#ifdef _WIN64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#ifndef _EXEONLY
#pragma comment(linker, "/entry:\"wWinMain\"")
#endif
#endif

#if _MSC_VER >=1600
#pragma execution_character_set("utf-8")
#define CP_TE CP_UTF8
#else
#define CP_TE CP_ACP
#endif
#ifndef WM_DPICHANGED
#define WM_DPICHANGED       0x02E0
#endif

class CteShellBrowser;
class CteTreeView;

#ifndef MONITOR_DPI_TYPE
enum MONITOR_DPI_TYPE {
	MDT_EFFECTIVE_DPI,
	MDT_ANGULAR_DPI,
	MDT_RAW_DPI,
	MDT_DEFAULT
};
#endif

#ifndef WINCOMPATTRDATA
struct WINCOMPATTRDATA
{
	DWORD dwAttrib;
	PVOID pvData;
	SIZE_T cbData;
};
#endif
//#define FVO_TABLACUS (FVO_VISTALAYOUT | FVO_CUSTOMPOSITION | FVO_CUSTOMORDERING)

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
	LPSECURITY_ATTRIBUTES lpSecurityAttributes;
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
	MONITOR_DPI_TYPE MonitorDpiType;
	RGBQUAD rgbquad;
};


enum PreferredAppMode
{
	APPMODE_DEFAULT = 0,
	APPMODE_ALLOWDARK = 1,
	APPMODE_FORCEDARK = 2,
	APPMODE_FORCELIGHT = 3,
	APPMODE_MAX = 4
};

typedef int (WINAPI* LPFNMessageBoxW)(HWND hWnd, LPCWSTR lpText, LPCWSTR lpCaption, UINT uType);

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
typedef HRESULT (STDAPICALLTYPE* LPFNPropVariantToVariant)(__in const PROPVARIANT *pPropVar, __out VARIANT *pVar);
typedef HRESULT (STDAPICALLTYPE* LPFNVariantToPropVariant)(__in const VARIANT* pVar, __out PROPVARIANT* pPropVar);

//Vista or higher.
typedef HRESULT (STDAPICALLTYPE* LPFNSHCreateItemFromIDList)(__in PCIDLIST_ABSOLUTE pidl, __in REFIID riid, __deref_out void **ppv);
typedef HRESULT (STDAPICALLTYPE* LPFNSHGetIDListFromObject)(__in IUnknown *punk, __deref_out PIDLIST_ABSOLUTE *ppidl);
typedef BOOL (WINAPI* LPFNChangeWindowMessageFilter)(__in UINT message, __in DWORD dwFlag);
typedef BOOL (WINAPI* LPFNAddClipboardFormatListener)(__in HWND hwnd);
typedef BOOL (WINAPI* LPFNRemoveClipboardFormatListener)(__in HWND hwnd);
//typedef HRESULT (STDAPICALLTYPE * LPFNPSFormatForDisplayAlloc)(__in REFPROPERTYKEY key, __in REFPROPVARIANT propvar, __in PROPDESC_FORMAT_FLAGS pdff, __deref_out PWSTR *ppszDisplay);
//typedef BOOL (WINAPI * LPFNChangeWindowMessageFilter)(UINT message, DWORD dwFlag);
#endif
typedef BOOL(WINAPI* LPFNSetWindowCompositionAttribute)(HWND hWnd, WINCOMPATTRDATA*);
typedef HRESULT (STDAPICALLTYPE * LPFNDwmSetWindowAttribute)(HWND hwnd, DWORD dwAttribute, LPCVOID pvAttribute, DWORD   cbAttribute);
typedef HRESULT (STDAPICALLTYPE * LPFNSHCreateShellItemArrayFromShellItem)(IShellItem *psi, REFIID riid, void **ppv);

//7 or higher
typedef BOOL (WINAPI* LPFNChangeWindowMessageFilterEx)(__in HWND hwnd, __in UINT message, __in DWORD action, __inout_opt PCHANGEFILTERSTRUCT pChangeFilterStruct);

//8 or higher
typedef BOOL (WINAPI* LPFNSetDefaultDllDirectories)(__in DWORD DirectoryFlags);

//8.1 or higher
typedef HRESULT (WINAPI* LPFNGetDpiForMonitor)(HMONITOR hmonitor, MONITOR_DPI_TYPE dpiType, UINT *dpiX, UINT *dpiY);

//10 Dark mode
typedef bool (WINAPI* LPFNAllowDarkModeForApp)(BOOL allow); //Deprecated since 18334
typedef PreferredAppMode (WINAPI* LPFNSetPreferredAppMode)(PreferredAppMode appMode);
typedef bool (WINAPI* LPFNAllowDarkModeForWindow)(HWND hwnd, BOOL allow);
typedef bool (WINAPI* LPFNShouldAppsUseDarkMode)();
typedef void (WINAPI* LPFNRefreshImmersiveColorPolicyState)();
typedef bool (WINAPI* LPFNIsDarkModeAllowedForWindow)(HWND hwnd);
//Scroll bar
typedef HTHEME (WINAPI *LPFNOpenNcThemeData)(HWND hWnd, LPCWSTR pszClassList);

//RTL
typedef NTSTATUS (WINAPI* LPFNRtlGetVersion)(PRTL_OSVERSIONINFOEXW lpVersionInformation);

//DLL
typedef HRESULT (STDAPICALLTYPE* LPFNDllGetClassObject)(REFCLSID rclsid, REFIID riid, LPVOID* ppv);
typedef HRESULT (STDAPICALLTYPE* LPFNDllCanUnloadNow)(void);

//EntoryPoint
typedef void (__stdcall* LPFNEntryPointW)(HWND hwnd, HINSTANCE hinst, LPWSTR lpszCmdLine, int nCmdShow);
//Plug in(Archive)
typedef HRESULT(WINAPI* LPFNGetArchive)(LPCWSTR lpArcPath, LPCWSTR lpItem, IStream **ppStream, LPVOID lpReserved);
//Plug in(Image)
typedef HRESULT (WINAPI* LPFNGetImage)(IStream *pStream, LPCWSTR lpPath, int cx, HBITMAP *phBM, int *pnAlpha);
//Plug in (MessageSFVCB)
typedef HRESULT (WINAPI* LPFNMessageSFVCB)(ICommDlgBrowser2 *pSB, IShellFolder2 *pSF2, IShellView *pSV, UINT uMsg, WPARAM wParam, LPARAM lParam);

#ifdef _USE_APIHOOK
//API Hook
typedef LSTATUS (APIENTRY* LPFNRegQueryValueExW)(HKEY hKey, LPCWSTR lpValueName, LPDWORD lpReserved, LPDWORD lpType, LPBYTE lpData, LPDWORD lpcbData);
typedef LSTATUS(APIENTRY* LPFNRegQueryValueW)(HKEY hKey, LPCWSTR lpSubKey, LPWSTR lpData, PLONG lpcbData);
typedef DWORD (WINAPI* LPFNGetSysColor)(int nIndex);
#endif

#ifndef LOAD_LIBRARY_SEARCH_SYSTEM32
#define LOAD_LIBRARY_SEARCH_SYSTEM32 0x00000800
#endif
#ifndef LOAD_LIBRARY_SEARCH_USER_DIRS
#define LOAD_LIBRARY_SEARCH_USER_DIRS 0x00000400
#endif

#ifndef IsReparseTagMicrosoft
#define IsReparseTagMicrosoft(_tag)             (((_tag) & 0x80000000))
#endif

#ifndef IO_REPARSE_TAG_MOUNT_POINT
#define IO_REPARSE_TAG_MOUNT_POINT              (0xA0000003L)
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
#define PATH_BLANK				L"about:blank"
#define MAX_LOADSTRING			100
#define MAX_HISTORY				32
#define MAX_CSIDL				0x3e
#define CSIDL_LIBRARY			0x3e
#define CSIDL_USER				0x3f
#define CSIDL_RESULTSFOLDER		0x40
#define MAX_CSIDL2				0x41

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
#define TET_EndThread			0x1fa6
#define TET_Title				0x1fa7
#define TET_FreeLibrary			0x1fa8
#define TET_Refresh				0x1fa9
#define TET_Expand				0x1faa
#define TET_SetRoot				0x1fab

#define TWM_CLIPBOARDUPDATE		WM_APP
#define SHGDN_ORIGINAL		0x40000000
#define SHGDN_FORPARSINGEX	0x80000000
#define TE_Labels				0
#define TE_ColumnsReplace		1
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
#define TE_OnMouseMessage		46
#define TE_OnEndThread			47
#define	TE_OnGetAlt				48
#define TE_OnItemPostPaint		49
#define TE_OnHandleIcon			50
#define TE_OnSorting			51
#define TE_OnSetName			52
#define TE_OnIncludeItem		53
#define TE_OnContentsChanged	54
#define TE_OnFilterView			55
#ifdef _USE_SYNC
#define TE_FN					56
#endif
#define Count_OnFunc			56
#define SB_TotalFileSize		0
#define SB_ColumnsReplace		1
#define SB_AltSelectedItems		2
#define SB_VirtualName			3
#define SB_OnIncludeObject		4
#define Count_SBFunc			5
#define WIC_OnGetAlt			0
#define Count_WICFunc			1
#define WB_External				0
#define WB_OnClose				1
#define Count_WBFunc			2

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
#define CTRL_FI    0xF0000

#define TE_VT 24
#define TE_VI 0xffffff
#define TE_METHOD		0x60010000
#define TE_METHOD_MAX	0x6001ffff
#define TE_METHOD_MASK	0x0000ffff
#define TE_PROPERTY		0x40010000
#define START_OnFunc	0x4001fc00
#define TE_OFFSET		0x4001ff00
#define DISPID_TE_ITEM  0x6001ffff
#define DISPID_TE_COUNT 0x4001ffff
#define DISPID_TE_INDEX 0x4001fffe
#define DISPID_TE_UNDEFIND 0x4001fffd
#define DISPID_TE_MAX TE_VI

#define DISPID_CB_ARANGE 0x4001fb00
#define TE_OBJECT	1
#define TE_ARRAY	2

#define TE_Type		0
#define TE_Left		1
#define TE_Top		2
#define TE_Width	3
#define TE_Right	3
#define TE_Height	4
#define TE_Bottom	4
#define TE_Tab		5
#define TE_CmdShow	6
#define TE_Layout	7
#define TE_NetworkTimeout	8
#define TE_SizeFormat		9
#define TE_Version			10
#define TE_UseHiddenFilter	11
#define TE_ColumnEmphasis	12
#define TE_ViewOrder		13
#define TE_LibraryFilter	14
#define TE_AutoArrange		15
#define TE_ShowInternet		16
#define Count_TE_params		17

#define	TE_AutoViewMode	0x10000
#define TECL_DARKTEXT 0xffffff
#define TECL_DARKTEXT2 0xe0e0e0
#define TECL_DARKBG 0x202020
#define TECL_DARKTAB 0x606060

#define TC_Flags	5
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

#define	SB_DoFunc	12
#define	SB_Count	13

#define	TVVERBS 16

#define MAP_TE	0
#define MAP_SB	1
#define MAP_TC	2
#define MAP_TV	3
#define MAP_GB	4
#define MAP_SS	5
#define MAP_API	6
#define MAP_LENGTH	7
#define TE_API_RESULT	12
#define TE_EXCEPINFO	13
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
	LONG lSize;
	WORD bSize;
	WORD nCount;
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
	BSTR	bsPath;
	LPITEMIDLIST	pidl;
	DISPID dispid;
	int	cArgs;
	HRESULT hr;
	LONG cRef;
	LONG cDo;
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

struct TEExecScript
{
	BSTR bsScript;
	BSTR bsLang;
	IStream *pStream;
};

struct TEFS
{
	BSTR bsPath;
	IStream *pStrmTotalFileSize;
	PDWORD pdwSessionId;
	DWORD dwSessionId;
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
	IStream *pStrmOnCompleted;
	BOOL	bDeleted;
	BOOL	bNavigateComplete;
};

struct TESortColumns
{
	SORTCOLUMN *pSC;
	int nCount;
	DWORD dwSessionId;
};

struct
{
	PCWSTR pszPropertyName;
	PCWSTR pszSemanticType;
}
const g_rgGenericProperties[] =
{
	{ L"System.Generic.String",          L"System.StructuredQueryType.String" },
	{ L"System.Generic.Integer",         L"System.StructuredQueryType.Integer" },
	{ L"System.Generic.DateTime",        L"System.StructuredQueryType.DateTime" },
	{ L"System.Generic.Boolean",         L"System.StructuredQueryType.Boolean" },
	{ L"System.Generic.FloatingPoint",   L"System.StructuredQueryType.FloatingPoint" }
};

typedef struct _REPARSE_DATA_BUFFER {
	ULONG  ReparseTag;
	USHORT ReparseDataLength;
	USHORT Reserved;
	union {
		struct {
			USHORT SubstituteNameOffset;
			USHORT SubstituteNameLength;
			USHORT PrintNameOffset;
			USHORT PrintNameLength;
			ULONG  Flags;
			WCHAR  PathBuffer[1];
		} SymbolicLinkReparseBuffer;
		struct {
			USHORT SubstituteNameOffset;
			USHORT SubstituteNameLength;
			USHORT PrintNameOffset;
			USHORT PrintNameLength;
			WCHAR  PathBuffer[1];
		} MountPointReparseBuffer;
		struct {
			UCHAR DataBuffer[1];
		} GenericReparseBuffer;
	} DUMMYUNIONNAME;
} REPARSE_DATA_BUFFER, *PREPARSE_DATA_BUFFER;

const CLSID CLSID_ShellShellNameSpace       = {0x2F2F1F96, 0x2BC1, 0x4b1c, { 0xBE, 0x28, 0xEA, 0x37, 0x74, 0xF4, 0x67, 0x6A}};
const CLSID CLSID_JScript9Chakra            = {0x16d51579, 0xa30b, 0x4c8b, { 0xa2, 0x76, 0x0f, 0xf4, 0xdc, 0x41, 0xe7, 0x55}};
const CLSID CLSID_JScriptEdgeChakra         = {0x1B7CD997, 0xE5FF, 0x4932, { 0xA7, 0xA6, 0x2A, 0x9E, 0x63, 0x6D, 0xA3, 0x85}};
const CLSID CLSID_ADODBStream               = {0x00000566, 0x0000, 0x0010, { 0x80, 0x00, 0x00, 0xAA, 0x00, 0x6D, 0x2E, 0xA4}};
const CLSID CLSID_ScriptingFileSystemObject = {0x0D43FE01, 0xF093, 0x11CF, { 0x89, 0x40, 0x00, 0xA0, 0xC9, 0x05, 0x42, 0x28}};
const CLSID CLSID_WScriptShell              = {0x72C24DD5, 0xD70A, 0x438B, { 0x8A, 0x42, 0x98, 0x42, 0x4B, 0x88, 0xAF, 0xB8}};
const CLSID CLSID_LibraryFolder             = {0xa5a3563a, 0x5755, 0x4a6f, { 0x85, 0x4e, 0xaf, 0xa3, 0x23, 0x0b, 0x19, 0x9f}};
const CLSID CLSID_DBFolder                  = {0xB2952B16, 0x0E07, 0x4E5A, { 0xB9, 0x93, 0x58, 0xC5, 0x2C, 0xB9, 0x4C, 0xAE}};
const CLSID CLSID_ResultsFolder             = {0x2965E715, 0xEB66, 0x4719, { 0xB5, 0x3F, 0x16, 0x72, 0x67, 0x3B, 0xBE, 0xFA}};

//Tablacus Explorer (Edge)
const CLSID CLSID_WebBrowserExt             = {0x55bbf1b8, 0x0d30, 0x4908, { 0xbe, 0x0c, 0xd5, 0x76, 0x61, 0x2a, 0x0f, 0x48}};
// {BD34E79B-963F-4AFB-B03E-C5BD289B5080}
const IID SID_TablacusObject                = {0xbd34e79b, 0x963f, 0x4afb, { 0xb0, 0x3e, 0xc5, 0xbd, 0x28, 0x9b, 0x50, 0x80}};
// {A7A52B88-B449-47BB-BD92-ABCCD8A6FED7}
const IID SID_TablacusArray                 = {0xa7a52b88, 0xb449, 0x47bb, { 0xbd, 0x92, 0xab, 0xcc, 0xd8, 0xa6, 0xfe, 0xd7 }};

#define teVariantCopy(d, s)     if (d) { VariantCopy(d, s); }
//Common functions
#ifdef _DEBUG
VOID teCheckMethod(LPSTR name, TEmethod *method, int nCount);
#endif
VOID teInitCommon();
VOID SafeRelease(PVOID ppObj);
VOID teCoTaskMemFree(LPVOID pv);
VARIANTARG* GetNewVARIANT(int n);
BOOL teSetObject(VARIANT *pv, PVOID pObj);
VOID teSetLong(VARIANT *pv, LONG i);
VOID teSetBool(VARIANT *pv, BOOL b);
VOID teSetLL(VARIANT *pv, LONGLONG ll);
BOOL teVarIsNumber(VARIANT *pv);
BOOL teStartsText(LPWSTR pszSub, LPCWSTR pszFile);
int GetIntFromVariantClear(VARIANT *pv);
int GetIntFromVariant(VARIANT *pv);
HRESULT teGetPropertyAt(IDispatch *pdisp, int i, VARIANT *pv);
HRESULT teGetProperty(IDispatch *pdisp, LPOLESTR sz, VARIANT *pv);
VOID Invoke4(IDispatch *pdisp, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs);
HRESULT Invoke5(IDispatch *pdisp, DISPID dispid, WORD wFlags, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs);
VOID teClearVariantArgs(int nArgs, VARIANTARG *pvArgs);
HRESULT teGetDisplayNameFromIDList(BSTR *pbs, LPITEMIDLIST pidl, SHGDNF uFlags);
HRESULT teGetDisplayNameBSTR(IShellFolder *pSF, PCUITEMID_CHILD pidl, SHGDNF uFlags, BSTR *pbs);
BOOL GetCSIDLFromPath(int *i, LPWSTR pszPath);
int ILGetCount(LPCITEMIDLIST pidl);
VOID teCreateSearchPath(BSTR *pbs, LPWSTR pszPath, LPWSTR pszSearch);
BOOL teGetIDListFromObject(IUnknown *punk, LPITEMIDLIST *ppidl);
VOID teGetSearchArg(BSTR *pbs, LPWSTR pszPath, LPWSTR pszArg);
int teStrCmpI(LPCWSTR lpStringW, LPCWSTR lpString2);
BOOL teILIsBlank(FolderItem *pFolderItem);
VOID teILCloneReplace(LPITEMIDLIST *ppidl, LPCITEMIDLIST pidl);
VOID teILFreeClear(LPITEMIDLIST *ppidl);
LPITEMIDLIST teILCreateFromPath2(LPITEMIDLIST pidlParent, LPWSTR pszPath, HWND hwnd);
VOID teSysFreeString(BSTR *pbs);
BOOL GetShellFolder(IShellFolder **ppSF, LPCITEMIDLIST pidl);
LPITEMIDLIST teILCreateFromPathEx(LPWSTR pszPath);
BOOL teILIsSearchFolder(LPCITEMIDLIST pidl);
BOOL teIsFileSystem(LPOLESTR pszPath);
BOOL teIsSearchFolder(LPWSTR lpszPath);
BSTR teSysAllocStringLen(const OLECHAR *strIn, UINT uSize);
VOID tePathAppend(BSTR *pbsPath, LPCWSTR pszPath, LPWSTR pszFile);
BOOL tePathMatchSpec1(LPCWSTR pszFile, LPWSTR pszSpec, WCHAR wSpecEnd);
BOOL tePathMatchSpec(LPCWSTR pszFile, LPWSTR pszSpec);
BOOL teStrSameIFree(BSTR bs, LPWSTR lpstr2);
int teStrCmpIWA(LPCWSTR lpStringW, LPCSTR lpStringA);
UINT GetpDataFromVariant(UCHAR **ppc, VARIANT *pv, VARIANT *pvMem);
BOOL GetDispatch(VARIANT *pv, IDispatch **ppdisp);
BOOL FindUnknown(VARIANT *pv, IUnknown **ppunk);
BSTR GetLPWSTRFromVariant(VARIANT *pv);
LONGLONG GetLLFromVariant(VARIANT *pv);
LONGLONG GetLLFromVariantClear(VARIANT *pv);
BOOL GetLLFromVariant2(LONGLONG *pll, VARIANT *pv);
char* GetpcFromVariant(VARIANT *pv, VARIANT *pvMem);
int SizeOfvt(VARTYPE vt);
DWORD teGetDWordFromDataObj(IDataObject *pDataObj, FORMATETC *pformatetcIn, DWORD dw);
HRESULT teDelProperty(IUnknown *punk, LPOLESTR sz);
VOID teAddRemoveProc(std::vector<LONG_PTR> *pppProc, LONG_PTR lpProc, BOOL bAdd);
int teGetObjectLength(IDispatch *pdisp);
LPITEMIDLIST teILCreateFromPath(LPWSTR pszPath);
LPITEMIDLIST teILCreateFromPath0(LPWSTR pszPath, BOOL bForceLimit);
BOOL teCreateItemFromPath(LPWSTR pszPath, IShellItem **ppSI);
LPITEMIDLIST teILCreateFromPath1(LPWSTR pszPath);
BOOL tePathIsNetworkPath(LPCWSTR pszPath);
HRESULT teCreateInstance(CLSID clsid, LPWSTR lpszDllFile, HMODULE *phDll, REFIID riid, PVOID *ppvObj);
BSTR teMultiByteToWideChar(UINT CodePage, LPCSTR lpA, int nLenA);
LPSTR teWideCharToMultiByte(UINT CodePage, LPCWSTR lpW, int nLenW);
HRESULT tePathIsDirectory(LPWSTR pszPath, int dwms, int iUseFS);
VOID CALLBACK teTimerProcParse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
VOID teAsyncInvoke(WORD wMode, int nArg, DISPPARAMS *pDispParams, VARIANT *pVarResult);
BOOL teSetObjectRelease(VARIANT *pv, PVOID pObj);
VOID teReleaseInvoke(TEInvoke *pInvoke);
HRESULT teILFolderExists(LPITEMIDLIST pidl);
int teUMSearch(int nMap, LPOLESTR bs);
VOID teWriteBack(VARIANT *pvArg, VARIANT *pvDat);
HRESULT tePutProperty(IUnknown *punk, LPOLESTR sz, VARIANT *pv);
HRESULT tePutProperty0(IUnknown *punk, LPOLESTR sz, VARIANT *pv, DWORD grfdex);
BOOL teVariantTimeToFileTime(DOUBLE dt, LPFILETIME pft);
VOID teGetRootFromDataObj(BSTR *pbs, IDataObject *pDataObj);
BSTR teSysAllocStringLenEx(const BSTR strIn, UINT uSize);
BSTR teSysAllocStringByteLen(LPCSTR psz, UINT len, UINT org);
LONGLONG teGetU(LONGLONG ll);
int CalcCrc32(BYTE *pc, int nLen, UINT c);
BOOL teFreeLibrary(HMODULE hDll, UINT uElpase);
VOID teVariantChangeType(VARIANTARG * pvargDest, const VARIANTARG * pvarSrc, VARTYPE vt);
BOOL teGetVariantTime(VARIANT *pv, DATE *pdt);
BOOL teGetSystemTime(VARIANT *pv, SYSTEMTIME *pst);
BOOL teVariantTimeToSystemTime(DATE dt, SYSTEMTIME *pst);
HRESULT teExecMethod(IDispatch *pdisp, LPOLESTR sz, VARIANT *pvResult, int nArg, VARIANTARG *pvArgs);
BOOL teFileTimeToVariantTime(LPFILETIME pft, DOUBLE *pdt);
VOID teExtraLongPath(BSTR *pbs);
LONGLONG GetParamFromVariant(VARIANT *pv, VARIANT *pvMem);
HRESULT teExceptionEx(EXCEPINFO *pExcepInfo, LPCSTR pszObjA, LPCSTR pszNameA);
VOID teSetULL(VARIANT *pv, LONGLONG ll);
VOID teSetSZ(VARIANT *pv, LPCWSTR lpstr);
BOOL GetDataObjFromVariant(IDataObject **ppDataObj, VARIANT *pv);
BOOL GetDataObjFromVariant2(IDataObject **ppDataObj, VARIANT *pv);
VOID AdjustIDList(LPITEMIDLIST *ppidllist, int nCount);
LPITEMIDLIST* IDListFormDataObj(IDataObject *pDataObj, long *pnCount);
int teDragQueryFile(HDROP hDrop, UINT iFile, BSTR *pbsPath);
VOID teSetBSTR(VARIANT *pv, BSTR *pbs, int nLen);
BOOL teChangeWindowMessageFilterEx(HWND hwnd, UINT message, DWORD action, PCHANGEFILTERSTRUCT pChangeFilterStruct);
VOID GetPointFormVariant(POINT *ppt, VARIANT *pv);
VOID PutPointToVariant(POINT *ppt, VARIANT *pv);
VOID teSzToArgv(IDispatch *pArray, LPTSTR lpsz);
HRESULT teDoDragDrop(HWND hwnd, IDataObject *pDataObj, DWORD *pdwEffect, BOOL bDropState);
HRESULT tePutPropertyAt(PVOID pObj, int i, VARIANT *pv);
HRESULT teInitStorage(LPVARIANTARG pvDllFile, LPVARIANTARG pvClass, LPWSTR lpwstr, HMODULE *phDll, IStorage **ppStorage);
HMODULE teCreateInstanceV(LPVARIANTARG pvDllFile, LPVARIANTARG pvClass, REFIID riid, PVOID *ppvObj);
HRESULT teCLSIDFromProgID(__in LPCOLESTR lpszProgID, __out LPCLSID lpclsid);
HRESULT teCLSIDFromString(__in LPCOLESTR lpsz, __out LPCLSID lpclsid);
HRESULT teExtract(IStorage *pStorage, LPWSTR lpszFolderPath, IProgressDialog *ppd, int *pnItems, int nCount, int nBase);
VOID teSetProgress(IProgressDialog *ppd, ULONGLONG ullCurrent, ULONGLONG ullTotal, int nMode);
BOOL teSetProgressEx(IProgressDialog *ppd, ULONGLONG ullCurrent, ULONGLONG ullTotal, int nMode);
VOID teCommaSize(LPWSTR pszIn, LPWSTR pszOut, UINT cchBuf, int nDigits);
VOID teStrFormatSize(DWORD dwFormat, LONGLONG qdw, LPWSTR pszBuf, UINT cchBuf);
LPWSTR teGetCommandLine();
HWND FindTreeWindow(HWND hwnd);
HWND teFindChildByClassA(HWND hwnd, LPCSTR lpClassA);
VOID teGetDisplayNameOf(VARIANT *pv, int uFlags, VARIANT *pVarResult);
void GetVarPathFromIDList(VARIANT *pVarResult, LPITEMIDLIST pidl, int uFlags);
HRESULT tePathGetFileName(BSTR *pbs, LPWSTR pszPath);
BOOL teLocalizePath(LPWSTR pszPath, BSTR *pbsPath);
int teGetMenuString(BSTR *pbs, HMENU hMenu, UINT uIDItem, BOOL fByPosition);
VOID teMenuText(LPWSTR sz);
int teGetModuleFileName(HMODULE hModule, BSTR *pbsPath);
int teGetthreadCount(DWORD dwProcessId);
BOOL teCreateSafeArray(VARIANT *pv, PVOID pSrc, DWORD dwSize, BOOL bBSTR);
BOOL GetVarArrayFromIDList(VARIANT *pv, LPITEMIDLIST pidl);
HRESULT GetFolderItemFromObject(FolderItem **ppid, IUnknown *pid);
VOID teSetIDList(VARIANT *pv, LPITEMIDLIST pidl);
VOID teSetIDListRelease(VARIANT *pv, LPITEMIDLIST *ppidl);
BOOL GetFolderItemFromIDList(FolderItem **ppid, LPITEMIDLIST pidl);
HRESULT teGetPropertyI(IDispatch *pdisp, LPOLESTR sz, VARIANT *pv);
HRESULT tePSFormatForDisplay(PROPERTYKEY *ppropKey, VARIANT *pv, DWORD pdfFlags, LPWSTR *ppszDisplay, int nSizeFormat);
VOID teGetFileTimeFromItem(IShellFolder2 *pSF2, LPCITEMIDLIST pidl, const SHCOLUMNID *pscid, FILETIME *pft);
BOOL teSetForegroundWindow(HWND hwnd);
VOID teSetExStyleOr(HWND hwnd, LONG l);
VOID teSetExStyleAnd(HWND hwnd, LONG l);
BOOL teGetIDListFromVariant(LPITEMIDLIST *ppidl, VARIANT *pv, BOOL bForceLimit = FALSE);
BOOL teSetWindowText(HWND hwnd, LPCWSTR lpcwstr);
LONGLONG teGetStreamPos(IStream *pStream);
DWORD teGetStreamSize(IStream *pStream);
VOID teCopyStream(IStream *pSrc, IStream *pDst);
HRESULT teSHGetDataFromIDList(IShellFolder *pSF, LPCITEMIDLIST pidlPart, int nFormat, void *pv, int cb);
int GetSizeOfStruct(LPOLESTR bs);
int GetSizeOf(VARIANT *pv);
VOID GetNewObject1(LPOLESTR sz, IDispatch **ppdisp);
VOID GetNewArray(IDispatch **ppArray);
VOID GetNewObject(IDispatch **ppObj);
VOID teArrayPush(IDispatch *pdisp, PVOID pObj);
HRESULT STDAPICALLTYPE teGetDpiForMonitor(HMONITOR hmonitor, MONITOR_DPI_TYPE dpiType, UINT *dpiX, UINT *dpiY);
HRESULT STDAPICALLTYPE tePSPropertyKeyFromStringEx(__in LPCWSTR pszString, __out PROPERTYKEY *pkey);
BOOL MessageSub(int nFunc, PVOID pObj, MSG *pMsg, HRESULT *phr);
VOID teSetUM(int nMap, LPCSTR name, LONG lData);
VOID teInitUM(int nMap, TEmethod *method, int nCount);
int teBSearch(TEmethod *method, int nSize, LPOLESTR bs);
VOID teCreateSafeArrayFromVariantArray(IDispatch *pdisp, VARIANT *pVarResult);
