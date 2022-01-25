// TE.cpp
// Tablacus Explorer (C)2011 Gaku
// MIT Lisence
// Visual Studio Express 2017 for Windows Desktop
// Windows SDK v7.1 - x64 exe
// Visual Studio 2017 (v141) - x64 dll
// Visual Studio 2015 - Windows XP (v140_xp) - x86 exe dll

// https://tablacus.github.io/

#include "stdafx.h"
#include "common.h"
#if defined(_WINDLL) || defined(_EXEONLY)
#include "objects.h"
#include "api.h"
#include "darkmode.h"
#include "TE.h"

// Global Variables:
HINSTANCE hInst;								// current instance
#ifdef _WINDLL
HINSTANCE g_hinstDll = NULL;
#endif
WCHAR	g_szTE[MAX_LOADSTRING];
WCHAR	g_szText[MAX_STATUS];
WCHAR	g_szStatus[MAX_STATUS];
BSTR	g_bsTitle = NULL;
BSTR	g_bsName = NULL;
BSTR	g_bsAnyText = NULL;
BSTR	g_bsAnyTextId = NULL;
HWND	g_hwndMain;
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
LPFNGetDpiForMonitor _GetDpiForMonitor = NULL;
LPFND2D1CreateFactory _D2D1CreateFactory = NULL;
LPFNDWriteCreateFactory _DWriteCreateFactory = NULL;
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
#ifdef USE_APIHOOK
LPFNRegQueryValueExW _RegQueryValueExW = NULL;
LPFNRegQueryValueW _RegQueryValueW = NULL;
LPFNGetSysColor _GetSysColor = NULL;
LPFNOpenNcThemeData _OpenNcThemeData = NULL;
#endif
extern LPFNSetPreferredAppMode _SetPreferredAppMode;
extern LPFNAllowDarkModeForWindow _AllowDarkModeForWindow;
extern LPFNSetWindowCompositionAttribute _SetWindowCompositionAttribute;
extern LPFNDwmSetWindowAttribute _DwmSetWindowAttribute;
extern LPFNShouldAppsUseDarkMode _ShouldAppsUseDarkMode;
extern LPFNRefreshImmersiveColorPolicyState _RefreshImmersiveColorPolicyState;
extern LRESULT CALLBACK TabCtrlProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);

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

CTE			*g_pTE;
std::vector<CteShellBrowser *> g_ppSB;
std::vector<CteTreeView *> g_ppTV;
std::vector<LONG_PTR> g_ppGetArchive;
std::vector<LONG_PTR> g_ppGetImage;
std::vector<LONG_PTR> g_ppMessageSFVCB;
std::vector<CteFolderItem *> g_ppENum;
CteWebBrowser *g_pWebBrowser = NULL;
#ifdef USE_OBJECTAPI
IDispatch	*g_pAPI = NULL;
#endif

LPITEMIDLIST g_pidls[MAX_CSIDL2];
BSTR		 g_bsPidls[MAX_CSIDL2];

IDispatch	*g_pOnFunc[Count_OnFunc];
std::vector<HMODULE>	g_pFreeLibrary;
std::vector<HMODULE>	g_phModule;
IDispatch	*g_pJS = NULL;
std::unordered_map<HWND, CteWebBrowser*> g_umSubWindows;
std::vector<TEFS *>	g_pFolderSize;
CRITICAL_SECTION g_csFolderSize;
IDropTargetHelper *g_pDropTargetHelper = NULL;
IUnknown	*g_pDraggingCtrl = NULL;
IDataObject	*g_pDraggingItems = NULL;
IDispatchEx *g_pCM = NULL;
IServiceProvider *g_pSP = NULL;
IQueryParser *g_pqp = NULL;
IDispatch *g_pMBText = NULL;
ULONG_PTR g_Token;
HHOOK	g_hHook;
HHOOK	g_hMouseHook;
HHOOK	g_hMessageHook;
HHOOK	g_hMenuKeyHook;
std::unordered_map<DWORD, HHOOK> g_umCBTHook;
HBRUSH	g_hbrDarkBackground;
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
POINT	g_ptDouble;
WPARAM	g_wParam2;

LONG	g_nLockUpdate = 0;
LONG    g_nCountOfThreadFolderSize = 0;
UINT	g_dwDoubleTime;
DWORD	g_paramFV[SB_Count];
DWORD	g_dwMainThreadId;
DWORD	g_dwCookieSW = 0;
DWORD   g_dwCookieJS = 0;
DWORD	g_dwTickKey;
DWORD	g_dwFreeLibrary = 0;
DWORD	g_dwTickMount = 0;
long	g_nProcKey	   = 0;
long	g_nProcMouse   = 0;
long	g_nProcFV      = 0;
long	g_nProcTV      = 0;
long	g_nThreads	   = 0;
long	g_lSWCookie = 0;
int		g_nReload = 0;
int		g_nException = 256;
std::unordered_map<std::string, LONG> g_maps[MAP_LENGTH];
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
BOOL	g_bInit = FALSE;
BOOL	g_bShowParseError = TRUE;
BOOL	g_bDragging = FALSE;
BOOL	g_bCanLayout = FALSE;
BOOL	g_bUpper10;
extern BOOL	g_bDarkMode;
extern std::unordered_map<HWND, HWND> g_umDlgProc;
BOOL	g_bDragIcon = TRUE;
COLORREF g_clrBackground = GetSysColor(COLOR_WINDOW);
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
//HHOOK	g_hMenuGMHook;
#endif
#ifdef CHECK_HANDLELEAK
int g_nLeak = 0;
#endif

TEmethod methodTE[] = {
	{ 1001, "Data" },
	{ 1002, "hwnd" },
	{ 1004, "About" },
	{ TE_METHOD + 1005, "Ctrl" },
	{ TE_METHOD + 1006, "Ctrls" },
	{ TE_METHOD + 1008, "ClearEvents" },
	{ TE_METHOD + 1009, "Reload" },
	{ TE_METHOD + 1010, "CreateObject" },
	{ TE_METHOD + 1020, "GetObject" },
	{ TE_METHOD + 1025, "AddEvent" },
	{ TE_METHOD + 1026, "RemoveEvent" },
	{ 1030, "WindowsAPI" },
	{ 1031, "WindowsAPI0" },
	{ 1131, "CommonDialog" },
	{ 1132, "WICBitmap" },
	{ 1132, "GdiplusBitmap" },//Deprecated
	{ 1136, "Arguments" },
	{ 1137, "ProgressDialog" },
	{ 1138, "DateTimeFormat" },
	{ 1139, "HiddenFilter" },
//	{ 1140, "Background" },//Deprecated
//	{ 1150, "ThumbnailProvider" },//Deprecated
	{ 1160, "DragIcon" },
	{ 1180, "ExplorerBrowserFilter" },
	{ TE_METHOD + 1133, "FolderItems" },
	{ TE_METHOD + 1134, "Object" },
	{ TE_METHOD + 1135, "Array" },
//	{ TE_METHOD + 1136, "Collection" },//Deprecated
	{ TE_METHOD + 1050, "CreateCtrl" },
	{ TE_METHOD + 1040, "CtrlFromPoint" },
	{ TE_METHOD + 1060, "MainMenu" },
	{ TE_METHOD + 1070, "CtrlFromWindow" },
	{ TE_METHOD + 1080, "LockUpdate" },
	{ TE_METHOD + 1090, "UnlockUpdate" },
	{ TE_METHOD + 1091, "ArrangeCB" },
	{ TE_METHOD + 1100, "HookDragDrop" },//Deprecated
#ifdef USE_TESTOBJECT
	{ 1200, "TestObj" },
#endif
	{ TE_OFFSET + TE_Type   , "Type" },
	{ TE_OFFSET + TE_Left   , "offsetLeft" },
	{ TE_OFFSET + TE_Top    , "offsetTop" },
	{ TE_OFFSET + TE_Right  , "offsetRight" },
	{ TE_OFFSET + TE_Bottom , "offsetBottom" },
	{ TE_OFFSET + TE_Tab, "Tab" },
	{ TE_OFFSET + TE_CmdShow, "CmdShow" },
	{ TE_OFFSET + TE_Layout, "Layout" },
	{ TE_OFFSET + TE_NetworkTimeout, "NetworkTimeout" },
	{ TE_OFFSET + TE_SizeFormat, "SizeFormat" },
	{ TE_OFFSET + TE_Version, "Version" },
	{ TE_OFFSET + TE_UseHiddenFilter, "UseHiddenFilter" },
	{ TE_OFFSET + TE_ColumnEmphasis, "ColumnEmphasis" },
	{ TE_OFFSET + TE_ViewOrder, "ViewOrder" },
	{ TE_OFFSET + TE_LibraryFilter, "LibraryFilter" },
	{ TE_OFFSET + TE_AutoArrange, "AutoArrange" },
	{ TE_OFFSET + TE_ShowInternet, "ShowInternet" },

	{ START_OnFunc + TE_Labels, "Labels" },
	{ START_OnFunc + TE_ColumnsReplace, "ColumnsReplace" },
	{ START_OnFunc + TE_OnBeforeNavigate, "OnBeforeNavigate" },
	{ START_OnFunc + TE_OnViewCreated, "OnViewCreated" },
	{ START_OnFunc + TE_OnKeyMessage, "OnKeyMessage" },
	{ START_OnFunc + TE_OnMouseMessage, "OnMouseMessage" },
	{ START_OnFunc + TE_OnCreate, "OnCreate" },
	{ START_OnFunc + TE_OnDefaultCommand, "OnDefaultCommand" },
	{ START_OnFunc + TE_OnItemClick, "OnItemClick" },
	{ START_OnFunc + TE_OnGetPaneState, "OnGetPaneState" },
	{ START_OnFunc + TE_OnMenuMessage, "OnMenuMessage" },
	{ START_OnFunc + TE_OnSystemMessage, "OnSystemMessage" },
	{ START_OnFunc + TE_OnShowContextMenu, "OnShowContextMenu" },
	{ START_OnFunc + TE_OnSelectionChanged, "OnSelectionChanged" },
	{ START_OnFunc + TE_OnClose, "OnClose" },
	{ START_OnFunc + TE_OnDragEnter, "OnDragEnter" },
	{ START_OnFunc + TE_OnDragOver, "OnDragOver" },
	{ START_OnFunc + TE_OnDrop, "OnDrop" },
	{ START_OnFunc + TE_OnDragLeave, "OnDragLeave" },
	{ START_OnFunc + TE_OnAppMessage, "OnAppMessage" },
	{ START_OnFunc + TE_OnStatusText, "OnStatusText" },
	{ START_OnFunc + TE_OnToolTip, "OnToolTip" },
	{ START_OnFunc + TE_OnNewWindow, "OnNewWindow" },
	{ START_OnFunc + TE_OnWindowRegistered, "OnWindowRegistered" },
	{ START_OnFunc + TE_OnSelectionChanging, "OnSelectionChanging" },
	{ START_OnFunc + TE_OnClipboardText, "OnClipboardText" },
	{ START_OnFunc + TE_OnCommand, "OnCommand" },
	{ START_OnFunc + TE_OnInvokeCommand, "OnInvokeCommand" },
	{ START_OnFunc + TE_OnArrange, "OnArrange" },
	{ START_OnFunc + TE_OnHitTest, "OnHitTest" },
	{ START_OnFunc + TE_OnVisibleChanged, "OnVisibleChanged" },
	{ START_OnFunc + TE_OnTranslatePath, "OnTranslatePath" },
	{ START_OnFunc + TE_OnNavigateComplete, "OnNavigateComplete" },
	{ START_OnFunc + TE_OnILGetParent, "OnILGetParent" },
	{ START_OnFunc + TE_OnViewModeChanged, "OnViewModeChanged" },
	{ START_OnFunc + TE_OnColumnsChanged, "OnColumnsChanged" },
	{ START_OnFunc + TE_OnItemPrePaint, "OnItemPrePaint" },
	{ START_OnFunc + TE_OnColumnClick, "OnColumnClick" },
	{ START_OnFunc + TE_OnBeginDrag, "OnBeginDrag" },
	{ START_OnFunc + TE_OnBeforeGetData, "OnBeforeGetData" },
	{ START_OnFunc + TE_OnIconSizeChanged, "OnIconSizeChanged" },
	{ START_OnFunc + TE_OnFilterChanged, "OnFilterChanged" },
	{ START_OnFunc + TE_OnBeginLabelEdit, "OnBeginLabelEdit" },
	{ START_OnFunc + TE_OnEndLabelEdit, "OnEndLabelEdit" },
	{ START_OnFunc + TE_OnReplacePath, "OnReplacePath" },
	{ START_OnFunc + TE_OnBeginNavigate, "OnBeginNavigate" },
	{ START_OnFunc + TE_OnSort, "OnSort" },
	{ START_OnFunc + TE_OnGetAlt, "OnGetAlt" },
	{ START_OnFunc + TE_OnEndThread, "OnEndThread" },
	{ START_OnFunc + TE_OnItemPostPaint, "OnItemPostPaint" },
	{ START_OnFunc + TE_OnHandleIcon, "OnHandleIcon" },
	{ START_OnFunc + TE_OnSorting, "OnSorting" },
	{ START_OnFunc + TE_OnSetName, "OnSetName" },
	{ START_OnFunc + TE_OnIncludeItem, "OnIncludeItem" },
	{ START_OnFunc + TE_OnContentsChanged, "OnContentsChanged" },
	{ START_OnFunc + TE_OnFilterView, "OnFilterView" },
#ifdef _USE_SYNC
	{ START_OnFunc + TE_FN, "fn" },
#endif
};

TEmethod methodSB[] = {
	{ TE_PROPERTY + 0xf001, "Data" },
	{ TE_PROPERTY + 0xf002, "hwnd" },
	{ TE_PROPERTY + 0xf003, "Type" },
	{ TE_METHOD + 0xf004, "Navigate" },
	{ TE_METHOD + 0xf007, "Navigate2" },
	{ TE_PROPERTY + 0xf008, "Index" },
	{ TE_PROPERTY + 0xf009, "FolderFlags" },
	{ TE_PROPERTY + 0xf00B, "History" },
	{ TE_PROPERTY + 0xf010, "CurrentViewMode" },
	{ TE_METHOD + 0xf010, "SetViewMode" },
	{ TE_PROPERTY + 0xf011, "IconSize" },
	{ TE_PROPERTY + 0xf012, "Options" },
	{ TE_PROPERTY + 0xf013, "SizeFormat" },
	{ TE_PROPERTY + 0xf014, "NameFormat" }, //Deprecated
	{ TE_PROPERTY + 0xf016, "ViewFlags" },
	{ TE_PROPERTY + 0xf017, "Id" },
	{ TE_PROPERTY + 0xf018, "FilterView" },
	{ TE_METHOD + 0xf018, "Search" },
	{ TE_PROPERTY + 0xf020, "FolderItem" },
	{ TE_PROPERTY + 0xf021, "TreeView" },
	{ TE_PROPERTY + 0xf024, "Parent" },
	{ TE_METHOD + 0xf026, "Focus" },
	{ TE_METHOD + 0xf031, "Close" },
	{ TE_PROPERTY + 0xf032, "Title" },
	{ TE_METHOD + 0xf033, "Suspend" },
	{ TE_METHOD + 0xf040, "Items" },
	{ TE_METHOD + 0xf041, "SelectedItems" },
	{ TE_PROPERTY + 0xf050, "ShellFolderView" },
	{ TE_PROPERTY + 0xf058, "Droptarget" },
	{ TE_PROPERTY + 0xf059, "Columns" },
	{ TE_METHOD + 0xf059, "GetColumns" },
//	{ TE_PROPERTY + 0xf05A, "Searches" },
	{ TE_METHOD + 0xf05B, "MapColumnToSCID" },
	{ TE_PROPERTY + 0xf102, "hwndList" },
	{ TE_PROPERTY + 0xf103, "hwndView" },
	{ TE_PROPERTY + 0xf104, "SortColumn" },
	{ TE_METHOD + 0xf104, "GetSortColumn" },
	{ TE_PROPERTY + 0xf105, "GroupBy" },
	{ TE_METHOD + 0xf107, "HitTest" },
	{ TE_PROPERTY + 0xf108, "hwndAlt" },
	{ TE_METHOD + 0xf110, "ItemCount" },
	{ TE_METHOD + 0xf111, "Item" },
	{ TE_METHOD + 0xf206, "Refresh" },
	{ TE_METHOD + 0xf207, "ViewMenu" },
	{ TE_METHOD + 0xf208, "TranslateAccelerator" },
	{ TE_METHOD + 0xf209, "GetItemPosition" },
	{ TE_METHOD + 0xf20A, "SelectAndPositionItem" },
	{ TE_METHOD + 0xf280, "SelectItem" },
	{ TE_PROPERTY + 0xf281, "FocusedItem" },
	{ TE_PROPERTY + 0xf282, "GetFocusedItem" },
	{ TE_METHOD + 0xf283, "GetItemRect" },
//	{ TE_METHOD + 0xf284, "ViewProperty" },
	{ TE_METHOD + 0xf300, "Notify" },
	{ TE_METHOD + 0xf400, "NavigateComplete" },
	{ TE_METHOD + 0xf501, "AddItem" },
	{ TE_METHOD + 0xf502, "RemoveItem" },
	{ TE_METHOD + 0xf503, "AddItems" },
	{ TE_METHOD + 0xf504, "RemoveAll" },
	{ TE_PROPERTY + 0xf505, "SessionId" },
	{ START_OnFunc + SB_TotalFileSize, "TotalFileSize" },
	{ START_OnFunc + SB_ColumnsReplace, "ColumnsReplace" },
	{ START_OnFunc + SB_OnIncludeObject, "OnIncludeObject" },
	{ START_OnFunc + SB_AltSelectedItems, "AltSelectedItems" },
	{ START_OnFunc + SB_VirtualName, "VirtualName" }, //Deprecated
};

TEmethod methodTC[] = {
	{ TE_PROPERTY + 1, "Data" },
	{ TE_PROPERTY + 2, "hwnd" },
	{ TE_PROPERTY + 3, "Type" },
	{ TE_METHOD + 6, "HitTest" },
	{ TE_METHOD + 7, "Move" },
	{ TE_PROPERTY + 8, "Selected" },
	{ TE_METHOD + 9, "Close" },
	{ TE_PROPERTY + 10, "SelectedIndex" },
	{ TE_PROPERTY + 11, "Visible" },
	{ TE_PROPERTY + 12, "Id" },
	{ TE_METHOD + 13, "LockUpdate" },
	{ TE_METHOD + 14, "UnlockUpdate" },
	{ TE_METHOD + 15, "SetOrder" },
	{ DISPID_NEWENUM, "_NewEnum" },
	{ DISPID_TE_ITEM, "Item" },
	{ DISPID_TE_COUNT, "Count" },
	{ DISPID_TE_COUNT, "length" },
	{ TE_OFFSET + TE_Left, "Left" },
	{ TE_OFFSET + TE_Top, "Top" },
	{ TE_OFFSET + TE_Width, "Width" },
	{ TE_OFFSET + TE_Height, "Height" },
	{ TE_OFFSET + TC_Flags, "Style" },
	{ TE_OFFSET + TC_Align, "Align" },
	{ TE_OFFSET + TC_TabWidth, "TabWidth" },
	{ TE_OFFSET + TC_TabHeight, "TabHeight" },
};

TEmethod methodTV[] = {
	{ TE_PROPERTY + 0x1001, "Data" },
	{ TE_PROPERTY + 0x1002, "Type" },
	{ TE_PROPERTY + 0x1003, "hwnd" },
	{ TE_METHOD + 0x1004, "Close" },
	{ TE_PROPERTY + 0x1005, "hwndTree" },
	{ TE_PROPERTY + 0x1007, "FolderView" },
	{ TE_PROPERTY + 0x1008, "Align" },
	{ TE_PROPERTY + 0x1009, "Visible" },
	{ TE_METHOD + 0x1106, "Focus" },
	{ TE_METHOD + 0x1107, "HitTest" },
	{ TE_METHOD + 0x1206, "Refresh" },
	{ TE_METHOD + 0x1283, "GetItemRect" },
	{ TE_METHOD + 0x1300, "Notify" },
	{ TE_OFFSET + SB_TreeWidth, "Width" },
	{ TE_OFFSET + SB_TreeFlags, "Style" },
	{ TE_OFFSET + SB_EnumFlags, "EnumFlags" },
	{ TE_OFFSET + SB_RootStyle, "RootStyle" },
	{ TE_PROPERTY + 0x2000, "SelectedItem" },
	{ TE_METHOD + 0x2001, "SelectedItems" },
	{ TE_PROPERTY + 0x2002, "Root" },
	{ TE_METHOD + 0x2003, "SetRoot" },
	{ TE_METHOD + 0x2004, "Expand" },
	{ TE_PROPERTY + 0x2005, "Columns" },
	{ TE_PROPERTY + 0x2006, "CountViewTypes" },
	{ TE_PROPERTY + 0x2007, "Depth" },
	{ TE_PROPERTY + 0x2008, "EnumOptions" },
	{ TE_PROPERTY + 0x2009, "Export" },
	{ TE_PROPERTY + 0x200a, "Flags" },
	{ TE_PROPERTY + 0x200b, "Import" },
	{ TE_PROPERTY + 0x200c, "Mode" },
	{ TE_PROPERTY + 0x200d, "ResetSort" },
	{ TE_PROPERTY + 0x200e, "SetViewType" },
	{ TE_PROPERTY + 0x200f, "Synchronize" },
	{ TE_PROPERTY + 0x2010, "TVFlags" },
};

//vvv Sort by name vvv
TEmethod methodWB[] = {
	{ TE_PROPERTY + 5, "Application" },
	{ TE_METHOD + 9, "Close" },
	{ TE_PROPERTY + 1, "Data" },
	{ TE_PROPERTY + 6, "Document" },
	{ TE_PROPERTY + 12, "DropMode" },
	{ TE_METHOD + 8, "Focus" },
	{ TE_PROPERTY + 2, "hwnd" },
	{ START_OnFunc + WB_OnClose, "OnClose" },
	{ TE_METHOD + 10, "PreventClose" },
	{ TE_METHOD + 4, "TranslateAccelerator" },
	{ TE_PROPERTY + 3, "Type" },
	{ TE_PROPERTY + 11, "Visible" },
	{ TE_PROPERTY + 7, "Window" },
};

TEmethod methodCM[] = {
	{ TE_PROPERTY + 5, "FolderView" },
	{ TE_METHOD + 4, "GetCommandString" },
	{ TE_METHOD + 6, "HandleMenuMsg" },
	{ TE_PROPERTY + 10, "hmenu" },
	{ TE_PROPERTY + 12, "idCmdFirst" },
	{ TE_PROPERTY + 13, "idCmdLast" },
	{ TE_PROPERTY + 11, "indexMenu" },
	{ TE_METHOD + 2, "InvokeCommand" },
	{ TE_METHOD + 3, "Items" },
	{ TE_METHOD + 1, "QueryContextMenu" },
	{ TE_PROPERTY + 14, "uFlags" },
};

// Forward declarations of functions included in this code module:
ATOM MyRegisterClass(HINSTANCE hInstance, LPWSTR szClassName, WNDPROC lpfnWndProc);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
IDispatch* teCreateObj(LONG lId, VARIANT *pvArg);

VOID ArrangeWindow();
VOID CALLBACK teTimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
VOID CALLBACK teTimerProcParse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
VOID CALLBACK teTimerProcForTree(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
BOOL teCreateItemFromPath(LPWSTR pszPath, IShellItem **ppSI);
int teGetModuleFileName(HMODULE hModule, BSTR *pbsPath);
BOOL GetVarArrayFromIDList(VARIANT *pv, LPITEMIDLIST pidl);
HRESULT STDAPICALLTYPE tePSPropertyKeyFromStringEx(__in LPCWSTR pszString, __out PROPERTYKEY *pkey);
VOID teArrayPush(IDispatch *pdisp, PVOID pObj);
static void threadParseDisplayName(void *args);

extern TEDispatchApi dispAPI[];
extern VOID InitObjects();
extern VOID InitAPI();

//Unit

IDropSource* teFindDropSource()
{
	return static_cast<IDropSource *>(g_pTE);
}

VOID teFixListState(HWND hwnd, int dwFlags) {
	if (hwnd && (dwFlags & SVSI_FOCUSED)) {// bug fix
#ifdef _2000XP
		if (g_bUpperVista) {
#endif
			int nFocused = ListView_GetNextItem(hwnd, -1, LVNI_FOCUSED);
			if (nFocused >= 0) {
				ListView_SetItemState(hwnd, nFocused, 0, LVIS_FOCUSED);
			}
#ifdef _2000XP
		}
#endif
	}
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
			HRESULT hr = pActiveObject->TranslateAcceleratorW(pMsg);
			if (hr == S_OK || (SUCCEEDED(hr) && GetKeyState(VK_MENU) < 0)) {
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
	if (!g_bMessageLoop) {
		return FALSE;
	}
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

HRESULT teQueryFolderItem(IUnknown *punk, CteFolderItem **ppid)
{
	HRESULT hr = punk->QueryInterface(g_ClsIdFI, (LPVOID *)ppid);
	if FAILED(hr) {
		*ppid = new CteFolderItem(NULL);
		LPITEMIDLIST pidl;
		if (teGetIDListFromObject(punk, &pidl)) {
			(*ppid)->Initialize(pidl);
			::CoTaskMemFree(pidl);
			hr = S_OK;
		}
	}
	return hr;
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

BSTR teLoadFromFile(BSTR bsFile)
{
	BSTR bsResult = NULL;
	HANDLE hFile = CreateFile(bsFile, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, NULL);
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
				LPITEMIDLIST pidl = ::SHSimpleIDListFromPath(bsPath);
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

#ifdef USE_TESTPATHMATCHSPEC
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

BOOL teHasProperty(IDispatch *pdisp, LPOLESTR sz)
{
	DISPID dispid;
	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &dispid);
	return hr == S_OK;
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

BOOL teSetRect(HWND hwnd, int left, int top, int right, int bottom)
{
	RECT rc;
	GetWindowRect(hwnd, &rc);
	MoveWindow(hwnd, left, top, right - left, bottom - top, TRUE);
	return (rc.right - rc.left != right - left || rc.bottom - rc.top != bottom - top);
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
		if (teIsClan(pSB->m_hwnd, hwnd)) {
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
		if (teIsClan(pTV->m_hwnd, hwnd)) {
			return pTV->m_bMain ? pTV : NULL;
		}
	}
	return NULL;
}

CteTreeView* TVfromhwnd(HWND hwnd)
{
	for (size_t i = 0; i < g_ppSB.size(); ++i) {
		CteShellBrowser *pSB = g_ppSB[i];
		if (teIsClan(pSB->m_pTV->m_hwnd, hwnd)) {
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

#ifdef USE_APIHOOK

HTHEME WINAPI teOpenNcThemeData(HWND hWnd, LPCWSTR pszClassList)
{
	if (teStrCmpIWA(pszClassList, "ScrollBar") == 0) {
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
	if (teStrCmpI(pszString, g_szTotalFileSizeCodeXP) == 0) {
		*pkey = PKEY_TotalFileSize;
		return S_OK;
	}
	if (teStrCmpI(pszString, g_szLabelCodeXP) == 0) {
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

BOOL teCreateItemFromVariant(VARIANT *pv, IShellItem **ppSI)
{
	BOOL Result = FALSE;
	LPITEMIDLIST pidl = NULL;
	if (teGetIDListFromVariant(&pidl, pv, TRUE)) {
		Result = SUCCEEDED(_SHCreateItemFromIDList(pidl, IID_PPV_ARGS(ppSI)));
	}
	teILFreeClear(&pidl);
	return Result;
}

BOOL GetFolderItemFromFolderItems(FolderItem **ppFolderItem, FolderItems *pFolderItems, int nIndex)
{
	VARIANT v;
	teSetLong(&v, nIndex);
	if (pFolderItems->Item(v, ppFolderItem) != S_OK) {
		teSetSZ(&v, PATH_BLANK);
		*ppFolderItem = new CteFolderItem(&v);
		VariantClear(&v);
	}
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
			if (teStrCmpI(bs, method[pi[nIndex]].name) < 0) {
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
		nCC = teStrCmpI(bs, method[pMap[nIndex]].name);
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

#ifdef USE_LOG
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

HRESULT ParseScript(LPOLESTR lpScript, LPOLESTR lpLang, VARIANT *pv, IDispatch **ppdisp, EXCEPINFO *pExcepInfo)
{
	HRESULT hr = E_FAIL;
	CLSID clsid;
	IActiveScript *pas = NULL;
	HRESULT *phr = &hr;

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
		CteActiveScriptSite *pass = new CteActiveScriptSite(punk, pExcepInfo);
		pas->SetScriptSite(pass);
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
			pass->m_pExcepInfo = NULL;
			if (pass->m_hr != S_OK) {
				hr = pass->m_hr;
			}
			pass->Release();
		}
		if (!ppdisp || !*ppdisp) {
			pas->SetScriptState(SCRIPTSTATE_CLOSED);
			pas->Close();
		}
		pas->Release();
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
								bAdd = !teGetUnavailable(pItem);
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
		teStopProgressDialog(ppd);
		SafeRelease(&ppd);
	}
	::CoUninitialize();
	::_endthread();
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
	for (int j = Count_OnFunc; j-- > 1;) {
		SafeRelease(&g_pOnFunc[j]);
	}

	for (size_t i = 0; i < g_ppSB.size(); ++i) {
		CteShellBrowser *pSB = g_ppSB[i];
		for (int j = Count_SBFunc; j-- > 1;) {
			SafeRelease(&pSB->m_ppDispatch[j]);
		}
	}

	while (!g_ppENum.empty()) {
		g_ppENum.back()->Clear();
		g_ppENum.pop_back();
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
		if (teIsClan(g_pWebBrowser->get_HWND(), hwnd)) {
			return g_pWebBrowser->QueryInterface(IID_PPV_ARGS(ppdisp));
		}
	}
	return E_FAIL;
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

#ifdef _DEBUG
/*
LRESULT CALLBACK MenuGMProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = 1;
	if (nCode >= 0) {
		if (nCode == HC_ACTION) {
			WCHAR pszNum[99];
			swprintf_s(pszNum, 99, L"%x\n", wParam);
			::OutputDebugString(pszNum);
		}
	}
	return lResult ? CallNextHookEx(g_hMenuGMHook, nCode, wParam, lParam) : TRUE;
}*/
#endif
LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
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
								if (InterlockedIncrement(&g_nProcMouse) < 5 || wParam != g_wParam2) {
									g_wParam2 = wParam;
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
									msg.wParam = pMHS->mouseData;
									msg.pt.x = pMHS->pt.x;
									msg.pt.y = pMHS->pt.y;
									CHAR szClassA[MAX_CLASS_NAME];
									GetClassNameA(msg.hwnd, szClassA, MAX_CLASS_NAME);
									if (lstrcmpA(szClassA, WC_LISTVIEWA) == 0) {
										if (msg.message == WM_LBUTTONDOWN) {
											msg.hwnd = WindowFromPoint(pMHS->pt);
										}
									} else if (lstrcmpA(szClassA, "DirectUIHWND") == 0) {
										if (msg.message == WM_LBUTTONDOWN) {
											if (GetTickCount() - g_dwDoubleTime < GetDoubleClickTime() &&
												abs(g_ptDouble.x - msg.pt.x) < GetSystemMetrics(SM_CXDRAG) &&
												abs(g_ptDouble.y - msg.pt.y) < GetSystemMetrics(SM_CXDRAG)) {
												msg.message = WM_LBUTTONDBLCLK;
											} else {
												g_ptDouble.x = msg.pt.x;
												g_ptDouble.y = msg.pt.y;
											}
											msg.hwnd = WindowFromPoint(pMHS->pt);
										} else if (msg.message == WM_LBUTTONUP) {
											g_dwDoubleTime = GetTickCount();
										}
									}
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
			if (((LPNMHDR)lParam)->code == LVN_GETDISPINFO && !(pSB->m_dwRedraw & 1)) {
				NMLVDISPINFO *lpDispInfo = (NMLVDISPINFO *)lParam;
				if (lpDispInfo->item.mask & LVIF_TEXT) {
					if (pSB->m_param[SB_ViewMode] == FVM_DETAILS || lpDispInfo->item.iSubItem == 0) {
						IDispatch *pdisp = (size_t)lpDispInfo->item.iSubItem < pSB->m_ppColumns.size() ? pSB->m_ppColumns[lpDispInfo->item.iSubItem] : NULL;
						if (pdisp) {
							bDoCallProc = FALSE;
							lResult = DefSubclassProc(hwnd, msg, wParam, lParam);
							pSB->ReplaceColumnsEx(pdisp, &lpDispInfo->item, NULL);
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
								pSB->m_pDTColumns[lpDispInfo->item.iSubItem] = 0;
							}
						}
					}
				}
				if ((lpDispInfo->item.mask & LVIF_IMAGE) && g_pOnFunc[TE_OnHandleIcon] && !(pSB->m_dwRedraw & 1)) {
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
							pSB->SetColumnsProperties(pnmv->iSubItem);
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
			} else if (((LPNMHDR)lParam)->code == LVN_BEGINSCROLL) {
				if (abs(((LPNMLVSCROLL)lParam)->dy) > (pSB->m_param[SB_ViewMode] == FVM_DETAILS ? 16 : 256) || pSB->m_param[SB_ViewMode] == FVM_LIST) {
					pSB->m_dwRedraw |= 2;
					SendMessage(pSB->m_hwnd, WM_SETREDRAW, FALSE, 0);
				}
			} else if (((LPNMHDR)lParam)->code == LVN_ENDSCROLL) {
				if (abs(((LPNMLVSCROLL)lParam)->dy) > (pSB->m_param[SB_ViewMode] == FVM_DETAILS ? 16 : 256) || pSB->m_param[SB_ViewMode] == FVM_LIST) {
					pSB->RedrawUpdate();
				}
			}

/// Custom Draw
			if (pSB->m_pShellView) {
				LPNMLVCUSTOMDRAW lplvcd = (LPNMLVCUSTOMDRAW)lParam;
				if (lplvcd->nmcd.hdr.code == NM_CUSTOMDRAW) {
					if (g_nWindowTheme != 2 && teIsDarkColor(pSB->m_clrBk)) { //Fix groups in dark background
						teFixGroup(lplvcd, pSB->m_clrBk);
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
										int nSelected = ListView_GetNextItem(hwndLV, -1, LVNI_SELECTED);
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
									if (GetKeyState(VK_SHIFT) < 0 && ListView_GetNextItem(hwndLV, -1, LVNI_SELECTED) >= 0) {
										int nFocused = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED);
										if (!ListView_GetItemState(hwndLV, nFocused, LVIS_SELECTED)) {
											ListView_SetItemState(hwndLV, ListView_GetNextItem(hwndLV, -1, LVNI_SELECTED), LVIS_FOCUSED, LVIS_FOCUSED);
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
		case LVM_SETSELECTEDCOLUMN:
			if (!g_param[TE_ColumnEmphasis]) {
				wParam = -1;
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
						SetTextColor(pnmcd->hdc, TECL_DARKTEXT);
						return CDRF_DODEFAULT;
					}
					if (pnmcd->dwDrawStage == CDDS_ITEMPOSTPAINT) {
						return CDRF_DODEFAULT;
					}
				}
			} else if (((LPNMHDR)lParam)->code == HDN_ITEMCHANGED) {
				pSB->FixColumnEmphasis();
			} else if (((LPNMHDR)lParam)->code == HDN_DROPDOWN) {
				if (pSB->SetColumnsProperties(((LPNMHEADER)lParam)->iItem)) {
					return 0;
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
		g_x = GET_X_LPARAM(lParam);
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
	if (g_pOnFunc[TE_OnArrange]) {
		VARIANTARG *pv = GetNewVARIANT(3);
		teSetObject(&pv[2], g_pTE);
		teSetLong(&pv[0], CTRL_TE);
		Invoke4(g_pOnFunc[TE_OnArrange], NULL, 3, pv);
	}
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
						teDoCommand(pSB->m_pTV, hwnd, WM_SIZE, 0, 0);//Resize
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
			} else if (msg != TCM_GETITEM) {
				bCancelPaint = FALSE;
			}
		}

		if (Result) {
			Result = TETCProc2(pTC, hwnd, msg, wParam, lParam);
		}
		return Result ? TabCtrlProc(hwnd, msg, wParam, lParam, uIdSubclass, dwRefData) : 0;
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
								CteShellBrowser *pSB1 = NULL;
								if (pMsg->message == WM_KEYDOWN) {
									if SUCCEEDED(pSB->QueryInterface(g_ClsIdSB, (LPVOID *)&pSB1)) {
										if (pSB1->m_hwndLV) {
											if ((pMsg->wParam == VK_PRIOR || pMsg->wParam == VK_NEXT)) {
												pSB1->m_dwRedraw |= 2;
												SendMessage(pSB1->m_hwnd, WM_SETREDRAW, FALSE, 0);
											}
										}
									}
								}
								if (pSV->TranslateAcceleratorW(pMsg) == S_OK) {
									hrResult = S_OK;
								}
								if (pSB1) {
									pSB1->RedrawUpdate();
									pSB1->Release();
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
				auto itr = g_umSubWindows.find(GetParent((HWND)GetWindowLongPtr(GetParent(pMsg->hwnd), GWLP_HWNDPARENT)));
				if (itr != g_umSubWindows.end()) {
					teTranslateAccelerator(itr->second, pMsg, &hrResult);
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

HMODULE teLoadLibrary(LPWSTR lpszName)
{
	BSTR bsPath = NULL;
	BSTR bsSystem;
	teGetDisplayNameFromIDList(&bsSystem, g_pidls[CSIDL_SYSTEM], SHGDN_FORPARSING);
	tePathAppend(&bsPath, bsSystem, lpszName);
	HMODULE hModule = NULL;
	hModule = GetModuleHandle(bsPath);
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

//WIC GetImage
#ifdef _WINDLL
EXTERN_C HRESULT WINAPI GetImage(IStream *pStream, LPWSTR lpfn, int cx, HBITMAP *phBM, int *pnAlpha)
{
	DLLEXPORT;
#else
HRESULT WINAPI GetImage(IStream *pStream, LPWSTR lpfn, int cx, HBITMAP *phBM, int *pnAlpha)
{
#endif
	HRESULT hr = E_FAIL;
	CteWICBitmap *pWB = new CteWICBitmap();
	pStream->AddRef();
	pWB->FromStreamRelease(&pStream, lpfn, FALSE, cx);
	if (pWB->HasImage()) {
		*phBM = pWB->GetHBITMAP(-1);
		*pnAlpha = 0;
		hr = S_OK;
	}
	pWB->Release();
	return hr;
}

//Dispatch API

VOID teApiGetProcAddress(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	CHAR szProcNameA[MAX_LOADSTRING];
	LPSTR lpProcNameA = szProcNameA;
	if (pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
		::WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)pDispParams->rgvarg[nArg - 1].bstrVal, -1, lpProcNameA, MAX_LOADSTRING, NULL, NULL);
		if (!param[0].hmodule) {
			if (!lstrcmpiA(lpProcNameA, "GetImage")) {
				teSetPtr(pVarResult, GetImage);
				return;
			}
		}
	} else {
		lpProcNameA = MAKEINTRESOURCEA(param[1].word);
	}
	teSetPtr(pVarResult, GetProcAddress(param[0].hmodule, lpProcNameA));
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

VOID Initlize()
{
#ifdef _DEBUG
	//Check orders
	teCheckMethod("methodCM", methodCM, _countof(methodCM));
	teCheckMethod("methodWB", methodWB, _countof(methodWB));
#endif
	teInitUM(MAP_TE, methodTE, _countof(methodTE));
	teInitUM(MAP_SB, methodSB, _countof(methodSB));
	teInitUM(MAP_TC, methodTC, _countof(methodTC));
	teInitUM(MAP_TV, methodTV, _countof(methodTV));
	InitObjects();
	InitAPI();
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
#ifdef _WINDLL
	wcex.hIcon			= LoadIcon(g_hinstDll, MAKEINTRESOURCE(IDI_TE));
	wcex.hIconSm		= LoadIcon(g_hinstDll, MAKEINTRESOURCE(IDI_TE));
#else
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_TE));
	wcex.hIconSm		= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_TE));
#endif
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground  = (HBRUSH)(COLOR_BTNFACE + 1);
	wcex.lpszMenuName	= 0;
	wcex.lpszClassName	= szClassName;

	return RegisterClassEx(&wcex);
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
				MultiByteToWideChar(CP_UTF8, 0, teStrCmpIWA(g_pWebBrowser->m_bstrPath, "TE32.exe") ? PathFileExists(g_pWebBrowser->m_bstrPath) ?
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
					pSB->RedrawUpdate();
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
							if (g_pWebBrowser && !teIsClan(g_pWebBrowser->get_HWND(), pcwp->hwnd)) {
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

#ifdef USE_APIHOOK
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
	BSTR bsTEWV2;
	teGetModuleFileName(NULL, &bsTEWV2);
	PathRemoveFileSpec(bsTEWV2);
#ifdef _WIN64
	PathAppend(bsTEWV2, L"lib\\tewv64.dll");
#else
	PathAppend(bsTEWV2, L"lib\\tewv32.dll");
#endif
#ifdef _DEBUG
	if (!PathFileExists(bsTEWV2)) {
#ifdef _WIN64
		lstrcpy(bsTEWV2, L"C:\\cpp\\tewv\\x64\\Debug\\tewv64d.dll");
#else
		lstrcpy(bsTEWV2, L"C:\\cpp\\tewv\\Debug\\tewv32d.dll");
#endif
	}
#endif
	hr = teCreateInstance(CLSID_WebBrowserExt, bsTEWV2, &g_hTEWV, IID_PPV_ARGS(ppWebBrowser));
	teSysFreeString(&bsTEWV2);
	if (hr == S_OK && !g_pSP) {
		VARIANT v;
		VariantInit(&v);
		(*ppWebBrowser)->GetProperty(L"version", &v);
		if (GetIntFromVariantClear(&v) >= 21070900) {
			(*ppWebBrowser)->QueryInterface(IID_PPV_ARGS(&g_pSP));
		} else {
			(*ppWebBrowser)->Release();
			hr = E_FAIL;
		}
	}
	g_nBlink = hr == S_OK ? 1 : 2;
	return hr;
}

#ifdef _WINDLL
BOOL WINAPI DllMain(HINSTANCE hinstDll, DWORD dwReason, LPVOID lpReserved)
{
	DWORD dwThreadId = GetCurrentThreadId();
	switch (dwReason) {
	case DLL_PROCESS_ATTACH:
		g_hinstDll = hinstDll;
	case DLL_THREAD_ATTACH:
		g_umCBTHook.try_emplace(dwThreadId, SetWindowsHookEx(WH_CBT, (HOOKPROC)CBTProc, NULL, dwThreadId));
		break;
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		{
			auto itr = g_umCBTHook.find(dwThreadId);
			if (itr != g_umCBTHook.end()) {
				UnhookWindowsHookEx(itr->second);
				g_umCBTHook.erase(itr);
			}
		}
		break;
	}
	return TRUE;
}
#endif

#ifdef _WINDLL
EXTERN_C void CALLBACK RunDLLW(HWND hWnd, HINSTANCE hInstance, LPWSTR lpCmdLine, int nCmdShow)
{
	DLLEXPORT;
#else
int APIENTRY _tWinMain(HINSTANCE hInstance,
	HINSTANCE hPrevInstance,
	LPTSTR    lpCmdLine,
	int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);
#endif
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
		g_pidls[CSIDL_UNAVAILABLE] = ILClone(g_pidls[CSIDL_RESULTSFOLDER]);
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
#ifdef USE_APIHOOK
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
//			*(FARPROC *)&_IsDarkModeAllowedForWindow = GetProcAddress(hDll, MAKEINTRESOURCEA(137));
#ifdef USE_APIHOOK
			*(FARPROC *)&_OpenNcThemeData = GetProcAddress(hDll, MAKEINTRESOURCEA(49));
			teAPIHook(L"comctl32.dll", (LPVOID)_OpenNcThemeData, &teOpenNcThemeData);
#endif
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
	if (hDll = teLoadLibrary(L"D2d1.dll")) {
		*(FARPROC *)&_D2D1CreateFactory = GetProcAddress(hDll, "D2D1CreateFactory");
	}
	if (hDll = teLoadLibrary(L"Dwrite.dll")) {
		*(FARPROC *)&_DWriteCreateFactory = GetProcAddress(hDll, "DWriteCreateFactory");
	}
	teLoadLibrary(L"mscoree.dll");//for .NET Shell exetension
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
	teInitCommon();

	g_dwMainThreadId = GetCurrentThreadId();
	VariantInit(&g_vArguments);
	BSTR bsPath = NULL;
	teGetModuleFileName(NULL, &bsPath);//Executable Path
	if (!g_pidls[CSIDL_UNAVAILABLE]) {
		ISearchFolderItemFactory *psfif = NULL;
		if SUCCEEDED(teCreateInstance(CLSID_SearchFolderItemFactory, NULL, NULL, IID_PPV_ARGS(&psfif))) {
			IShellItem *psi = NULL;
			if (teCreateItemFromPath(bsPath, &psi)) {
				IShellItemArray *psia;
				if (
#ifdef _2000XP
					_SHCreateShellItemArrayFromShellItem &&
#endif
					SUCCEEDED(_SHCreateShellItemArrayFromShellItem(psi, IID_PPV_ARGS(&psia)))) {
					psfif->SetScope(psia);
					psia->Release();
				}
			}
			psfif->GetIDList(&g_pidls[CSIDL_UNAVAILABLE]);
			psfif->Release();
		}
	}
	if (!g_pidls[CSIDL_UNAVAILABLE]) {
		g_pidls[CSIDL_UNAVAILABLE] = ILClone(g_pidls[CSIDL_RESULTSFOLDER]);
	}
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
#ifdef _WINDLL
					return;
#else
					return FALSE;
#endif
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
	if (!g_hwndMain) {
		Finalize();
#ifdef _WINDLL
		return;
#else
		return FALSE;
#endif
	}
	// ClipboardFormat
	IDLISTFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_SHELLIDLIST);
	DROPEFFECTFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
	DRAGWINDOWFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(L"DragWindow");
	IsShowingLayeredFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(L"IsShowingLayered");
	// Hook
	g_hHook = SetWindowsHookEx(WH_CALLWNDPROC, (HOOKPROC)HookProc, NULL, g_dwMainThreadId);
	g_hMouseHook = SetWindowsHookEx(WH_MOUSE, (HOOKPROC)MouseProc, NULL, g_dwMainThreadId);
#ifdef _DEBUG
	DWORD dwThreadId = GetCurrentThreadId();
	auto itr = g_umCBTHook.find(dwThreadId);
	if (itr == g_umCBTHook.end()) {
		g_umCBTHook[dwThreadId] = SetWindowsHookEx(WH_CBT, (HOOKPROC)CBTProc, NULL, dwThreadId);
	}
#endif
	//Brush
	g_hbrDarkBackground = CreateSolidBrush(TECL_DARKBG);
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
#ifdef _WINDLL
		return;
#else
		return FALSE;
#endif
	}
	teSysFreeString(&bsScript);
	IGlobalInterfaceTable *pGlobalInterfaceTable;
	CoCreateInstance(CLSID_StdGlobalInterfaceTable, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pGlobalInterfaceTable));
	pGlobalInterfaceTable->RegisterInterfaceInGlobal(g_pJS, IID_IDispatch, &g_dwCookieJS);
	InitializeCriticalSection(&g_csFolderSize);
	//WindowsAPI
#ifdef USE_OBJECTAPI
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
#ifdef USE_HTMLDOC
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
#ifdef USE_HTMLDOC
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
		bVisible = (vResult.vt == VT_BSTR) && (teStrCmpIWA(vResult.bstrVal, "wait") == 0);
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
#ifdef USE_HTMLDOC
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
		for (auto itr = g_umSubWindows.begin(); itr != g_umSubWindows.end(); ++itr) {
			itr->second->Release();
		}
		g_umSubWindows.clear();
#ifdef USE_OBJECTAPI
		SafeRelease(&g_pAPI);
#endif
		pGlobalInterfaceTable->RevokeInterfaceFromGlobal(g_dwCookieJS);
		pGlobalInterfaceTable->Release();
		SafeRelease(&g_pqp);
		SafeRelease(&g_pJS);
		SafeRelease(&g_pAutomation);
		SafeRelease(&g_pDropTargetHelper);
#ifdef _DEBUG
		auto itr = g_umCBTHook.find(GetCurrentThreadId());
		if (itr != g_umCBTHook.end()) {
			UnhookWindowsHookEx(itr->second);
			g_umCBTHook.erase(itr);
	}
#endif
		UnhookWindowsHookEx(g_hMouseHook);
		UnhookWindowsHookEx(g_hHook);
		DeleteObject(g_hbrDarkBackground);
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
	DeleteCriticalSection(&g_csFolderSize);
	SafeRelease(&g_pTE);
	VariantClear(&g_vData);
#ifndef _WINDLL
	return (int) msg.wParam;
#endif
}

VOID ArrangeWindow()
{
	if (g_bArrange) {
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

BOOL teResolveLink(LPITEMIDLIST *ppidl)
{
	BOOL bResult = FALSE;
	IShellFolder2 *pSF2;
	LPCITEMIDLIST pidlPart;
	if SUCCEEDED(SHBindToParent(*ppidl, IID_PPV_ARGS(&pSF2), &pidlPart)) {
		SFGAOF sfAttr = SFGAO_LINK | SFGAO_FOLDER;
		if SUCCEEDED(pSF2->GetAttributesOf(1, &pidlPart, &sfAttr)) {
			if ((sfAttr & (SFGAO_LINK | SFGAO_FOLDER)) == SFGAO_LINK) {
				IShellLink *pSL;
				if SUCCEEDED(pSF2->GetUIObjectOf(NULL, 1, &pidlPart, IID_IShellLink, NULL, (LPVOID *)&pSL)) {
					if (pSL->Resolve(NULL, SLR_NO_UI) == S_OK) {
						LPITEMIDLIST pidl;
						if (pSL->GetIDList(&pidl) == S_OK) {
							teCoTaskMemFree(*ppidl);
							*ppidl = pidl;
							bResult = TRUE;
						}
					}
					pSL->Release();
				}
			}
		}
		pSF2->Release();
	}
	return bResult;
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
			if (!pInvoke->pidl && teIsFileSystem(pInvoke->bsPath)) {
				pInvoke->pidl = ::SHSimpleIDListFromPath(pInvoke->bsPath);
			}
			pInvoke->hr = E_PATH_NOT_FOUND;
			if (pInvoke->pidl) {
				teResolveLink(&pInvoke->pidl);
				IShellFolder *pSF;
				if (GetShellFolder(&pSF, pInvoke->pidl)) {
					IEnumIDList *peidl = NULL;
					pInvoke->hr = pSF->EnumObjects(NULL, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS, &peidl);
					if (pInvoke->hr == S_OK) {
						peidl->Release();
					}
					pSF->Release();
				}
				if (pInvoke->hr == E_PATH_NOT_FOUND || pInvoke->hr == WININET_E_CANNOT_CONNECT ||
					pInvoke->hr == E_BAD_NET_NAME ||pInvoke->hr == E_NOT_READY || pInvoke->hr == E_BAD_NETPATH) {
					teILFreeClear(&pInvoke->pidl);
				}
			}
		}
		if (pInvoke->wMode) {
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

VOID CALLBACK teTimerProcParse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);
	TEInvoke *pInvoke = (TEInvoke *)idEvent;
	try {
		if (pInvoke->dispid == TE_METHOD + 0xfb07) {/*Navigate2Ex*/
			CteShellBrowser *pSB = NULL;
			if SUCCEEDED(pInvoke->pdisp->QueryInterface(g_ClsIdSB, (LPVOID *)&pSB)) {
				LONG lNavigating = pSB->m_pTC->m_lNavigating;
				pSB->Release();
				if (lNavigating) {
					SetTimer(g_hwndMain, (UINT_PTR)pInvoke, 100 * lNavigating, teTimerProcParse);
					return;
				}
			}
		}
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
						WCHAR pszError[SIZE_BUFF];
						if (FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
							NULL, pInvoke->hr, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&lpBuf, SIZE_BUFF, NULL)) {
							swprintf_s(pszError, MAX_COLUMNS, L"%s (0x%x)\n\n%s", lpBuf, pInvoke->hr, pInvoke->bsPath);
							int r = MessageBox(g_hwndMain, pszError, _T(PRODUCTNAME), MB_ABORTRETRYIGNORE);
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
						pSB->SetEmptyText();
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
			case TET_Redraw:
				SendMessage(pTV->m_hwnd, WM_SETREDRAW, TRUE, 0);
				RedrawWindow(pTV->m_hwnd, NULL, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
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
				SafeRelease(&g_pqp);
				if (g_pSW) {
					teRegister();
				}
				teGetDarkMode();
				teSetDarkMode(hWnd);
				CHAR pszClassA[MAX_CLASS_NAME];
				for (auto itr = g_umDlgProc.begin(); itr != g_umDlgProc.end(); ++itr) {
					GetClassNameA(itr->second, pszClassA, MAX_CLASS_NAME);
					if (::PathMatchSpecA(pszClassA, TOOLTIPS_CLASSA)) {
						SetWindowTheme(itr->second, g_bDarkMode ? L"darkmode_explorer" : L"explorer", NULL);
					}
				}
				if (_RegenerateUserEnvironment) {
					try {
						if (teStrCmpIWA((LPCWSTR)lParam, "Environment") == 0) {
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
			case CCM_SETBKCOLOR:
				g_clrBackground = lParam;
				break;
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
		switch (message)
		{
			case WM_CLOSE:
				try {
					auto itr = g_umSubWindows.find(hWnd);
					if (itr != g_umSubWindows.end()) {
						CteWebBrowser *pWB = itr->second;
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
				{
					auto itr = g_umSubWindows.find(hWnd);
					if (itr != g_umSubWindows.end()) {
						teSetObjectRects(itr->second->m_pWebBrowser, hWnd);
					}
				}
				break;
			case WM_ACTIVATE:
				if (LOWORD(wParam)) {
					auto itr = g_umSubWindows.find(hWnd);
					if (itr != g_umSubWindows.end()) {
						SetTimer(itr->second->m_hwndParent, (UINT_PTR)itr->second, 100, teTimerProcSW);
					}
				}
				break;
			case WM_ERASEBKGND:
				return 1;
			case WM_SETTINGCHANGE:
				teGetDarkMode();
				teSetDarkMode(hWnd);
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
	m_dwRedraw &= ~3;
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
			if (m_pFolderItem->m_v.vt == VT_BSTR) {
				m_pFolderItem->QueryInterface(IID_PPV_ARGS(&pid));
				wFlags &= ~(SBSP_RELATIVE | SBSP_REDIRECT);
			} else if (m_pFolderItem->m_dwUnavailable) {
				if (m_pFolderItem->m_v.vt == VT_EMPTY && m_pFolderItem->m_pidl) {
					GetVarPathFromFolderItem(m_pFolderItem, &m_pFolderItem->m_v);
					teILFreeClear(&m_pFolderItem->m_pidl);
					m_pFolderItem->QueryInterface(IID_PPV_ARGS(&pid));
					wFlags &= ~(SBSP_RELATIVE | SBSP_REDIRECT);
				}
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
			if ((!pid->m_pidl && pid->m_v.vt == VT_BSTR && !pid->GetAlt()) ||
				(pid->m_dwUnavailable && (GetTickCount() - pid->m_dwUnavailable > 500))) {
				if ((m_pShellView && pid->m_dwUnavailable) ||
					tePathMatchSpec(pid->m_v.bstrVal, L"*.lnk") ||
					(tePathIsNetworkPath(pid->m_v.bstrVal) && tePathIsDirectory(pid->m_v.bstrVal, 100, 3) != S_OK)) {
					if (m_nUnload || g_nLockUpdate <= 1) {
						Navigate1Ex(pid->m_v.bstrVal, pFolderItems, wFlags, pPrevious, nErrorHandling);
					} else {
						m_nUnload = 9;
					}
					bResult = TRUE;
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

HRESULT CteShellBrowser::Navigate2(FolderItem *pFolderItem, UINT wFlags, DWORD *param, FolderItems *pFolderItems, CteFolderItem *pPrevious, CteShellBrowser *pHistSB)
{
	RECT rc;
	LPITEMIDLIST pidl = NULL;
	HRESULT hr;

	SafeRelease(&m_pFolderItem1);
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
		if (!teGetIDListFromObjectForFV(static_cast<FolderItem *>(m_pFolderItem1), &pidl)) {
			hr = E_FAIL;
		}
		if (teILIsBlank(m_pFolderItem1)) {
			teILCloneReplace(&pidl, g_pidls[CSIDL_RESULTSFOLDER]);
		}
	}
	if (m_pFolderItem1 && m_pFolderItem1->m_dwUnavailable) {
		if (m_pShellView) {
			hr = S_FALSE;
		} else if (ILIsEmpty(pidl)) {
			teILCloneReplace(&pidl, g_pidls[CSIDL_UNAVAILABLE]);
		}
	}
	if (hr != S_OK) {
		if (hr == S_FALSE && !m_pFolderItem) {
			if SUCCEEDED(teQueryFolderItem(pFolderItem, &m_pFolderItem)) {
				teILCloneReplace(&m_pFolderItem->m_pidlAlt, g_pidls[CSIDL_RESULTSFOLDER]);
			}
		}
		teCoTaskMemFree(pidl);
		return m_pidl ? S_OK : hr;
	}
	hr = S_FALSE;
	if (!pidl) {
		return S_OK;
	}
	if (teResolveLink(&pidl)) {
		if (m_pFolderItem1) {
			m_pFolderItem1->Initialize(pidl);
		}
	}
	DWORD dwFrame = (m_param[SB_Type] == CTRL_EB
#ifdef _2000XP
		&& g_bUpperVista
#endif
	) ? EBO_SHOWFRAMES : 0;
	UINT uViewMode = m_param[SB_ViewMode];
	m_nGroupByDelay = 0;
	//ExplorerBrowser
	if (m_pExplorerBrowser && !pFolderItems
#ifdef _2000XP
		&& g_bUpperVista
#endif
		) {
		m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
		if (!ILIsEqual(pidl, g_pidls[CSIDL_RESULTSFOLDER]) || !ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
			if (GetShellFolder2(&pidl) == S_OK) {
				BrowseToObject(dwFrame, teGetFolderViewOptions(pidl, m_param[SB_ViewMode]));
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
				teGetIDListFromObject(static_cast<FolderItem *>(pPrevious), &m_pidl);
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
	if (!teILIsEqual(static_cast<FolderItem *>(m_pFolderItem), static_cast<FolderItem *>(pPrevious))) {
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
#ifdef USE_SHELLBROWSER
		hr = E_FAIL;
		if (dwFrame) {
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
		if (hr == E_ACCESSDENIED) {
			hr = E_FAIL;
		}
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
		teGetIDListFromObject(static_cast<FolderItem *>(pPrevious), &m_pidl);
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

VOID CteShellBrowser::GetInitFS(FOLDERSETTINGS *pfs)
{
	if (m_nForceViewMode != FVM_AUTO) {
		m_param[SB_ViewMode] = m_nForceViewMode;
		m_param[SB_IconSize] = m_nForceIconSize;
		m_bAutoVM = FALSE;
	}
	pfs->ViewMode = !m_bAutoVM || ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER]) ? m_param[SB_ViewMode] : FVM_AUTO;
	pfs->fFlags = (m_param[SB_FolderFlags] | FWF_USESEARCHFOLDER | FWF_SNAPTOGRID) & ~FWF_NOENUMREFRESH;
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
				FOLDERSETTINGS fs;
				GetInitFS(&fs);
				if (SUCCEEDED(m_pExplorerBrowser->Initialize(m_pTC->m_hwndStatic, &rc, &fs))) {
					hr = GetShellFolder2(&m_pidl);
					if (hr == S_OK) {
						m_pExplorerBrowser->Advise(static_cast<IExplorerBrowserEvents *>(this), &m_dwEventCookie);
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
						m_hwnd = NULL;
						hr = BrowseToObject(dwFrame, teGetFolderViewOptions(m_pidl, m_param[SB_ViewMode]));
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
		if (!teILIsEqual(static_cast<FolderItem *>(m_pFolderItem), pPrevious)) {
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
		if (teILIsEqual(static_cast<FolderItem *>(m_pFolderItem), static_cast<FolderItem *>(pPrevious))) {
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
		GetFocusedItem(&m_ppLog[m_uLogIndex]->m_pidlFocused);
	}
}

VOID CteShellBrowser::FocusItem()
{
	if (!m_pShellView) {
		return;
	}
	if (m_pFolderItem && !m_pFolderItem->m_dwUnavailable) {
		if (m_pFolderItem->m_pidlFocused) {
			teFixListState(m_hwndLV, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS | SVSI_DESELECTOTHERS | SVSI_SELECTIONMARK | SVSI_SELECT);
			SelectItemEx(m_pFolderItem->m_pidlFocused, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS | SVSI_DESELECTOTHERS | SVSI_SELECTIONMARK | SVSI_SELECT);
			teILFreeClear(&m_pFolderItem->m_pidlFocused);
		}
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
			FolderItem *pid1 = NULL;
			if (teILGetParent(pPrevious, &pid1)) {
				teQueryFolderItem(pid1, &m_pFolderItem1);
				pid1->Release();
				if (Navigate1(m_pFolderItem1, wFlags & ~SBSP_PARENT, pFolderItems, pPrevious, 0)) {
					return S_FALSE;
				}
				teGetIDListFromObject(pPrevious, &m_pFolderItem1->m_pidlFocused);
			}
		} else {
			if (!m_pFolderItem1) {
				m_pFolderItem1 = new CteFolderItem(NULL);
			}
			m_pFolderItem1->Initialize(g_pidls[CSIDL_DESKTOP]);
		}
		if (teILIsEqual(static_cast<FolderItem *>(pPrevious), static_cast<FolderItem *>(m_pFolderItem1))) {
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
					m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&m_ppLog[uLogIndex]);
				}
				SaveFocusedItemToHistory();
			}
			if SUCCEEDED(pHistSB->m_ppLog[--uLogIndex]->QueryInterface(g_ClsIdFI, (LPVOID *)&m_pFolderItem1)) {
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
					m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&m_ppLog[uLogIndex]);
				}
				SaveFocusedItemToHistory();
			}
			if SUCCEEDED(pHistSB->m_ppLog[++uLogIndex]->QueryInterface(g_ClsIdFI, (LPVOID *)&m_pFolderItem1)) {
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
		if (m_pShellView || g_nLockUpdate > 1) {
			return S_FALSE;
		}
		teQueryFolderItem(pid, &m_pFolderItem1);
		m_pFolderItem1->MakeUnavailable();
		return S_OK;
	}
	if (wFlags & SBSP_RELATIVE) {
		LPITEMIDLIST pidlPrevius = NULL;
		teGetIDListFromObjectForFV(pPrevious, &pidlPrevius);
		if (pidlPrevius && !ILIsEqual(pidlPrevius, g_pidls[CSIDL_RESULTSFOLDER]) && !ILIsEmpty(pidl)) {
			LPITEMIDLIST pidlFull = NULL;
			pidlFull = ILCombine(pidlPrevius, pidl);
			m_pFolderItem1 = new CteFolderItem(NULL);
			m_pFolderItem1->Initialize(pidlFull);
			teCoTaskMemFree(pidlFull);
		} else if (pPrevious) {
			pPrevious->QueryInterface(g_ClsIdFI, (LPVOID *)&m_pFolderItem1);
		} else if (pidlPrevius) {
			m_pFolderItem1 = new CteFolderItem(NULL);
			m_pFolderItem1->Initialize(pidlPrevius);
		} else {
			VARIANT v;
			teSetSZ(&v, PATH_BLANK);
			m_pFolderItem1 = new CteFolderItem(&v);
			VariantClear(&v);
		}
		teILFreeClear(&pidlPrevius);
		teILFreeClear(&pidl);
	} else if (pid) {
		teQueryFolderItem(pid, &m_pFolderItem1);
	}
	return m_pFolderItem1 ? S_OK : E_FAIL;
}

VOID CteShellBrowser::Refresh(BOOL bCheck)
{
	m_bRefreshLator = FALSE;
	if (!m_pFolderItem->m_dwUnavailable) {
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
				if (!m_pShellView || tePathIsNetworkPath(vResult.bstrVal) || m_pFolderItem->m_dwUnavailable) {
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
			if (m_pFolderItem->m_dwUnavailable || ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
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
			if (teIsClan(g_pWebBrowser->m_hwndBrowser, hwnd) || (m_pTV && teIsClan(m_pTV->m_hwnd, hwnd))) {
				return FALSE;
			}
			CHAR szClassA[MAX_CLASS_NAME];
			GetClassNameA(hwnd, szClassA, MAX_CLASS_NAME);
			if (::PathMatchSpecA(szClassA, WC_TREEVIEWA)) {
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
		if (teIsClan(m_hwnd, hwnd)) {
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
	if (teStrCmpI(bsText, bsOldText)) {
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
		*piItem = ListView_GetNextItem(m_hwndLV, -1, LVNI_FOCUSED);
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
		dwMask = (dwMask ^ m_param[SB_FolderFlags]) & (~(FWF_NOENUMREFRESH | FWF_USESEARCHFOLDER | FWF_SNAPTOGRID));
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
						if (teStrCmpI(szText, bsTotalFileSize) == 0) {
							m_nFolderSizeIndex = i;
							hdi.fmt |= HDF_RIGHT;
						}
						if (teStrCmpI(szText, bsLabel) == 0) {
							m_nLabelIndex = i;
						}
						if (teStrCmpI(szText, bsLinkTarget) == 0) {
							m_nLinkTargetIndex = i;
						}
#ifdef _2000XP
					}
#endif
					if (teStrCmpI(szText, bsSize) == 0) {
						m_nSizeIndex = i;
					}
				}
				if (fmt != hdi.fmt) {
					hdi.mask = HDI_FORMAT;
					Header_SetItem(hHeader, i, &hdi);
				}
				LVCOLUMN col = { LVCF_MINWIDTH };
				ListView_SetColumn(m_hwndLV, i, &col);
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

HRESULT CteShellBrowser::GetPropertyKey(int iItem, PROPERTYKEY *pPropKey)
{
	UINT uCount = 0;
	IColumnManager *pColumnManager;
	if (m_pShellView && SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager)))) {
		if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_VISIBLE, &uCount)) {
			if (iItem < uCount) {
				PROPERTYKEY *pPropKeyList = new PROPERTYKEY[uCount];
				if SUCCEEDED(pColumnManager->GetColumns(CM_ENUM_VISIBLE, pPropKeyList, uCount)) {
					*pPropKey = pPropKeyList[iItem];
				}
				delete[] pPropKeyList;
			}
		}
		pColumnManager->Release();
	}
	if (iItem < uCount) {
		return S_OK;
	}
	HWND hHeader = ListView_GetHeader(m_hwndLV);
	if (hHeader) {
		UINT nHeader = Header_GetItemCount(hHeader);
		if (nHeader) {
			int *piOrderArray = new int[nHeader];
			ListView_GetColumnOrderArray(m_hwndLV, nHeader, piOrderArray);
			iItem = piOrderArray[iItem];
			delete[] piOrderArray;
		}
	}
	HDITEM hdi;
	WCHAR pszText[MAX_COLUMN_NAME_LEN];
	hdi.cchTextMax = MAX_COLUMN_NAME_LEN;
	hdi.pszText = pszText;
	hdi.mask = HDI_TEXT;
	Header_GetItem(ListView_GetHeader(m_hwndLV), iItem, &hdi);
	return teGetPropertyKeyFromName(m_pSF2, pszText, pPropKey);
}

BOOL CteShellBrowser::SetColumnsProperties(int iItem)
{
	if (iItem == m_nLabelIndex) {
		SetLabelProperties();
	} else if (iItem == m_nFolderSizeIndex) {
		if (!SetFolderSizeProperties()) {
			return TRUE;
		}
	} else {
		IDispatch *pdisp = (size_t)iItem < m_ppColumns.size() ? m_ppColumns[iItem] : NULL;
		if (pdisp) {
			PROPERTYKEY propKey;
			if SUCCEEDED(GetPropertyKey(iItem, &propKey)) {
				SetReplaceColumnsProperties(pdisp, propKey);
			}
		}
	}
	return FALSE;
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
				if (v.bstrVal && cch >= 0) {
					lstrcpyn(szText, v.bstrVal, cch);
				}
			} else if (v.vt != VT_EMPTY) {
				if (cch >= 0) {
					teStrFormatSize(GetSizeFormat(), GetLLFromVariant(&v), szText, cch);
				}
			} else if (m_ppDispatch[SB_TotalFileSize]) {
				v.vt = VT_BSTR;
				v.bstrVal = NULL;
				tePutProperty(m_ppDispatch[SB_TotalFileSize], bsPath, &v);
				TEFS *pFS = new TEFS[1];
				pFS->bsPath = ::SysAllocString(bsPath);
				pFS->pdwSessionId = &m_pFolderItem->m_dwSessionId;
				pFS->dwSessionId = m_pFolderItem->m_dwSessionId;
				pFS->hwnd = m_hwndLV;
				CoMarshalInterThreadInterfaceInStream(IID_IDispatch, m_ppDispatch[SB_TotalFileSize], &pFS->pStrmTotalFileSize);
				PushFolderSizeList(pFS);
				if (g_nCountOfThreadFolderSize < 8) {
					++g_nCountOfThreadFolderSize;
					_beginthread(&threadFolderSize, 0, NULL);
				}
			}
			teSysFreeString(&bsPath);
		} else if SUCCEEDED(pSF2->GetDetailsEx(pidl, &PKEY_Size, &v)) {
			if (cch >= 0) {
				teStrFormatSize(GetSizeFormat(), GetLLFromVariant(&v), szText, cch);
			}
		}
		if (cch < 0) {
			_VariantToPropVariant(&v, (PROPVARIANT *)szText);
		}
		VariantClear(&v);
	}
}

BOOL CteShellBrowser::SetFolderSizeProperties()
{
	BOOL bAllOK = TRUE;
	IFolderView2 *pFV2;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		PROPVARIANT propVar;
		PropVariantInit(&propVar);
		int iItems = 0;
		pFV2->ItemCount(SVGIO_ALLVIEW, &iItems);
		for (int i = 0; i < iItems; ++i) {
			LPITEMIDLIST pidl;
			pFV2->Item(i, &pidl);
			SetFolderSize(m_pSF2, pidl, (LPWSTR)&propVar, -1);
			if (propVar.vt != VT_EMPTY && propVar.vt != VT_BSTR) {
				pFV2->SetViewProperty(pidl, PKEY_TotalFileSize, propVar);
			} else {
				bAllOK = FALSE;
			}
			PropVariantClear(&propVar);
			teCoTaskMemFree(pidl);
		}
		pFV2->Release();
	}
	return bAllOK;
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
				if (cch >= 0) {
					lstrcpyn(szText, v.bstrVal, cch);
				} else {
					_VariantToPropVariant(&v, (PROPVARIANT *)szText);
				}
			}
			::VariantClear(&v);
		}
		teSysFreeString(&bs);
	}
}

VOID CteShellBrowser::SetLabelProperties()
{
	IFolderView2 *pFV2;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		PROPVARIANT propVar;
		PropVariantInit(&propVar);
		int iItems = 0;
		pFV2->ItemCount(SVGIO_ALLVIEW, &iItems);
		for (int i = 0; i < iItems; ++i) {
			LPITEMIDLIST pidl;
			pFV2->Item(i, &pidl);
			SetLabel(pidl, (LPWSTR)&propVar, -1);
			pFV2->SetViewProperty(pidl, PKEY_Contact_Label, propVar);
			PropVariantClear(&propVar);
			teCoTaskMemFree(pidl);
		}
		pFV2->Release();
	}
}

BOOL CteShellBrowser::ReplaceColumns(IDispatch *pdisp, LVITEM *plvi, LPITEMIDLIST pidl)
{
	BOOL bResult = FALSE;
	VARIANTARG *pv = GetNewVARIANT(3);
	teSetObject(&pv[2], this);
	LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidl);
	teSetIDListRelease(&pv[1], &pidlFull);
	if (plvi->cchTextMax >= 0) {
		teSetSZ(&pv[0], plvi->pszText);
	}
	VARIANT v;
	VariantInit(&v);
	Invoke4(pdisp, &v, 3, pv);
	if (v.vt == VT_BSTR && ::SysStringLen(v.bstrVal)) {
		if (plvi->cchTextMax >= 0) {
			lstrcpyn(plvi->pszText, v.bstrVal, plvi->cchTextMax);
		} else {
			_VariantToPropVariant(&v, (PROPVARIANT *)plvi->pszText);
		}
		bResult = TRUE;
	}
	VariantClear(&v);
	return bResult;
}

VOID CteShellBrowser::ReplaceColumnsEx(IDispatch *pdisp, LVITEM *plvi, LPITEMIDLIST *ppidl)
{
	IFolderView *pFV;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
		LPITEMIDLIST pidl;
		if SUCCEEDED(pFV->Item(plvi->iItem, &pidl)) {
			BOOL bResult = FALSE;
			if (teHasProperty(pdisp, L"push")) {
				VARIANT v;
				VariantInit(&v);
				for (int nLen = teGetObjectLength(pdisp); --nLen >= 0;) {
					if SUCCEEDED(teGetPropertyAt(pdisp, nLen, &v)) {
						IDispatch *pdisp2;
						if (GetDispatch(&v, &pdisp2)) {
							if (bResult = ReplaceColumns(pdisp2, plvi, pidl)) {
								nLen = 0;
							}
							pdisp2->Release();
						}
						VariantClear(&v);
					}
				}
			} else {
				bResult = ReplaceColumns(pdisp, plvi, pidl);
			}
			if (ppidl) {
				*ppidl = pidl;
			} else {
				if (!bResult && plvi->cchTextMax >= 0) {
					PROPERTYKEY propKey;
					if SUCCEEDED(GetPropertyKey(plvi->iSubItem, &propKey)) {
						VARIANT v;
						VariantInit(&v);
						m_pSF2->GetDetailsEx(pidl, &propKey, &v);
						if (v.vt == VT_BSTR) {
							lstrcpyn(plvi->pszText, v.bstrVal, plvi->cchTextMax);
						}
						VariantClear(&v);
					}
				}
				teILFreeClear(&pidl);
			}
		}
		SafeRelease(&pFV);
	}
}

VOID CteShellBrowser::SetReplaceColumnsProperties(IDispatch *pdisp, PROPERTYKEY propKey)
{
	PROPVARIANT propVar;
	LPWSTR pszDisplay = NULL;
	propVar.vt = VT_BSTR;
	propVar.bstrVal = ::SysAllocString(L"OK");
#ifdef _2000XP
	if (_PSGetPropertyDescription) {
#endif
		IPropertyDescription *pdesc;
		if SUCCEEDED(_PSGetPropertyDescription(propKey, IID_PPV_ARGS(&pdesc))) {
			pdesc->FormatForDisplay(propVar, PDFF_DEFAULT, &pszDisplay);
			pdesc->Release();
		}
#ifdef _2000XP
	}
#endif
	if (!pszDisplay) {
		PropVariantClear(&propVar);
		return;
	}
	if (lstrcmp(propVar.bstrVal, pszDisplay)) {
		teCoTaskMemFree(pszDisplay);
		PropVariantClear(&propVar);
		return;
	}
	teCoTaskMemFree(pszDisplay);
	LVITEM lvi;
	lvi.cchTextMax = -1;
	IFolderView2 *pFV2;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		PropVariantInit(&propVar);
		lvi.pszText = (LPWSTR)&propVar;
		int iItems = 0;
		pFV2->ItemCount(SVGIO_ALLVIEW, &iItems);
		for (lvi.iItem = 0; lvi.iItem < iItems; ++lvi.iItem) {
			LPITEMIDLIST pidl = NULL;
			ReplaceColumnsEx(pdisp, &lvi, &pidl);
			if (propVar.vt != VT_EMPTY) {
				pFV2->SetViewProperty(pidl, propKey, propVar);
				PropVariantClear(&propVar);
			} else {
				pFV2->GetViewProperty(pidl, propKey, &propVar);
				if (propVar.vt == VT_BSTR) {
					VARIANT v;
					VariantInit(&v);
					m_pSF2->GetDetailsEx(pidl, &propKey, &v);
					if (v.vt == VT_BSTR) {
						if (lstrcmp(propVar.bstrVal, v.bstrVal)) {
							teSysFreeString(&propVar.bstrVal);
							propVar.bstrVal = v.bstrVal;
							pFV2->SetViewProperty(pidl, propKey, propVar);
							v.vt = VT_EMPTY;
						}
					}
					VariantClear(&v);
				}
				PropVariantClear(&propVar);
			}
			teILFreeClear(&pidl);
		}
		pFV2->Release();
	}
}

VOID CteShellBrowser::GetViewModeAndIconSize(BOOL bGetIconSize)
{
	if (!m_pShellView || (m_pFolderItem->m_dwUnavailable && !m_bCheckLayout)) {
		return;
	}
	FOLDERSETTINGS fs;
	fs.ViewMode = m_param[SB_ViewMode];
	int iImageSize = m_param[SB_IconSize];
	m_pShellView->GetCurrentInfo(&fs);
	if (bGetIconSize || fs.ViewMode != m_param[SB_ViewMode]) {
		IFolderView2 *pFV2;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
			pFV2->GetViewModeAndIconSize((FOLDERVIEWMODE *)&fs.ViewMode, &iImageSize);
			pFV2->Release();
#ifdef _2000XP
		} else {
			if (fs.ViewMode == FVM_ICON) {
				iImageSize = 32;
			} else if (fs.ViewMode == FVM_TILE) {
				iImageSize = 48;
			} else if (fs.ViewMode == FVM_THUMBNAIL) {
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
		m_param[SB_ViewMode] = fs.ViewMode;
		if (m_bCheckLayout && !m_bViewCreated) {
			FOLDERVIEWOPTIONS fvo = teGetFolderViewOptions(m_pidl, fs.ViewMode);
			if (m_pExplorerBrowser) {
				IFolderViewOptions *pOptions;
				if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
					FOLDERVIEWOPTIONS fvo0 = fvo;
					pOptions->GetFolderViewOptions(&fvo0);
					m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
					if (fvo ^ fvo0 & FVO_VISTALAYOUT) {
						m_bCheckLayout = FALSE;
					}
					if ((m_param[SB_Options] & EBO_SHOWFRAMES) != (m_param[SB_Type] == CTRL_EB ? EBO_SHOWFRAMES : 0)) {
						m_bCheckLayout = FALSE;
						m_nForceViewMode = FVM_AUTO;
					}
					pOptions->Release();
				}
			} else if (m_param[SB_Type] == CTRL_EB || fvo == FVO_DEFAULT) {
				m_bCheckLayout = FALSE;
			}
			if (m_nForceViewMode != FVM_AUTO) {
				m_param[SB_ViewMode] = m_nForceViewMode;
				m_param[SB_IconSize] = m_nForceIconSize;
				m_nForceViewMode = FVM_AUTO;
				SetViewModeAndIconSize(TRUE);
			} else if (!m_bCheckLayout) {
				m_nForceViewMode = fs.ViewMode;
				m_nForceIconSize = m_param[SB_IconSize];
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
					CteFolderItem *pid1;
					teQueryFolderItem(pid, &pid1);
					m_ppLog.push_back(pid1);
					pid1->Release();
				}
			}
		}
		m_uLogIndex = UINT(m_ppLog.size()) - 1 - uLogIndex;
		SafeRelease(&pFolderItems);
	} else if ((wFlags & (SBSP_NAVIGATEBACK | SBSP_NAVIGATEFORWARD | SBSP_WRITENOHISTORY)) == 0) {
		if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
			if (!m_pFolderItem || m_pFolderItem->m_v.vt != VT_BSTR) {
				return;
			}
		}
		if (m_uLogIndex < m_ppLog.size() && (teILIsEqual(static_cast<FolderItem *>(m_pFolderItem), static_cast<FolderItem *>(m_ppLog[m_uLogIndex])) || teILIsBlank(m_ppLog[m_uLogIndex]))) {
			if (m_pFolderItem != m_ppLog[m_uLogIndex]) {
				m_ppLog[m_uLogIndex]->Release();
				m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&m_ppLog[m_uLogIndex]);
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
			m_pFolderItem = new CteFolderItem(NULL);
			m_pFolderItem->Initialize(m_pidl);
		}
		if (m_pFolderItem) {
			if (g_dwTickMount && GetTickCount() - g_dwTickMount < 500) { //#248
				m_pFolderItem->AddRef();
			}
			m_pFolderItem->AddRef();
			m_uLogIndex = UINT(m_ppLog.size());
			m_ppLog.push_back(m_pFolderItem);
		}
	}
	g_dwTickMount = 0;
}

STDMETHODIMP CteShellBrowser::GetViewStateStream(DWORD grfMode, IStream **ppStrm)
{
//	return SHCreateStreamOnFileEx(L"", STGM_READWRITE | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_ARCHIVE, TRUE, NULL, ppStrm);
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
	HRESULT hr = teGetDispId(NULL, MAP_SB, *rgszNames, rgDispId, FALSE);
	if (hr != S_OK && m_pdisp) {
		hr = m_pdisp->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
		if (hr == S_OK) {
			if (teStrCmpIWA(*rgszNames, "SortColumns") == 0) {
				m_dispidSortColumns = *rgDispId;
			}
		} else {
			*rgDispId = DISPID_TE_UNDEFIND;
		}
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
							if (teStrCmpI(methodArgs[piArgs[i]].name, methodColumns[piCur[i]].name)) {
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
					CteFolderItem *pid;
					if SUCCEEDED(pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid)) {
						if (pid->m_dwUnavailable && pid->m_v.vt == VT_BSTR) {
							pid->Clear();
							pid->m_dwUnavailable = NULL;
						}
						pid->Release();
					}
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
					if (m_uLogIndex < m_ppLog.size()) {
						if (m_pFolderItem != m_ppLog[m_uLogIndex]) {
							m_ppLog[m_uLogIndex]->Release();
							m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&m_ppLog[m_uLogIndex]);
						}
					}
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
				} else if (teStrCmpI(m_bsFilter, v.bstrVal)) {
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
				if (nArg >= 1 && GetBoolFromVariant(&pDispParams->rgvarg[nArg - 1])) {
					teCreateSafeArrayFromVariantArray(pFolderItems, pVarResult);
					SafeRelease(&pFolderItems);
					return S_OK;
				}
				teSetObjectRelease(pVarResult, pFolderItems);
			}
			return S_OK;

		case TE_METHOD + 0xf041://SelectedItems
			if SUCCEEDED(SelectedItems(&pFolderItems)) {
				if (nArg >= 0 && GetBoolFromVariant(&pDispParams->rgvarg[nArg])) {
					teCreateSafeArrayFromVariantArray(pFolderItems, pVarResult);
					SafeRelease(&pFolderItems);
					return S_OK;
				}
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
				if (teGetIDListFromObject(static_cast<FolderItem *>(m_pFolderItem), &pidl)) {
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
				if (uItem == SVGIO_SELECTION) {
					if (m_ppDispatch[SB_AltSelectedItems]) {
						FolderItems *pid;
						if SUCCEEDED(m_ppDispatch[SB_AltSelectedItems]->QueryInterface(IID_PPV_ARGS(&pid))) {
							if SUCCEEDED(pid->get_Count(&pVarResult->lVal)) {
								pVarResult->vt = VT_I4;
							}
							pid->Release();
						}
						return S_OK;
					}
					if (m_pFolderItem->m_pidlFocused && m_pFolderItem->m_pidl) {
						teSetLong(pVarResult, 1);
						return S_OK;
					}
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
				if (pid) {
					DWORD dwUnavailable = m_pFolderItem->m_dwUnavailable;
					m_pFolderItem->m_dwUnavailable = 0;
					if (dwUnavailable) {
						BrowseObject2(pid, SBSP_SAMEBROWSER);
						pid->Release();
					} else if (m_pFolderItem) {
						m_pFolderItem->Release();
						teQueryFolderItem(pid, &m_pFolderItem);
						SafeRelease(&pid);
						if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
							teILFreeClear(&m_pidl);
							teGetIDListFromObject(static_cast<FolderItem *>(m_pFolderItem), &m_pidl);
						}
						Refresh(FALSE);
					}
				}
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
#ifdef USE_VIEWPROPERTY
		case TE_METHOD + 0xf284://ViewProperty
			if (m_pShellView && nArg >= 1 && pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
				PROPERTYKEY propKey;
				if SUCCEEDED(teGetPropertyKeyFromName(m_pSF2, pDispParams->rgvarg[nArg - 1].bstrVal, &propKey)) {
					IFolderView2 *pFV2;
					if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
						LPITEMIDLIST pidl = NULL;
						if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
							LPCITEMIDLIST pidlLast = ILFindLastID(pidl);
							PROPVARIANT propVar;
							PropVariantInit(&propVar);
							if (nArg >= 2) {
								_VariantToPropVariant(&pDispParams->rgvarg[nArg - 2], &propVar);
								pFV2->SetViewProperty(pidlLast, propKey, propVar);
								PropVariantClear(&propVar);
							}
							if (pVarResult) {
								pFV2->GetViewProperty(pidlLast, propKey, &propVar);
								_PropVariantToVariant(&propVar, pVarResult);
								PropVariantClear(&propVar);
							}
							teILFreeClear(&pidl);
						}
						pFV2->Release();
					}
				}
			}
			return S_OK;
#endif
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
			if (m_bVisible && !IsWindowVisible(m_hwnd)) {
				ShowWindow(m_hwnd, SW_SHOWNA);
			}
			return S_OK;

		case TE_METHOD + 0xf500://NavigationComplete
			NavigateComplete(FALSE);
			return S_OK;

		case TE_METHOD + 0xf501://AddItem
			if (nArg >= 0) {
				if (nArg >= 1) {
					if (!m_pFolderItem || m_pFolderItem->m_dwSessionId != GetIntFromVariant(&pDispParams->rgvarg[0])) {
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
					teSetLong(&pAddItems->pv[0], m_pFolderItem->m_dwSessionId);
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
			if (m_pFolderItem) {
				teSetLong(pVarResult, m_pFolderItem->m_dwSessionId);
			}
			return S_OK;

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
			if (m_pFolderItem) {
				teILFreeClear(&m_pFolderItem->m_pidlFocused);
			}
			return DoFunc(TE_OnSelectionChanged, this, S_OK);

		case DISPID_CONTENTSCHANGED://XP-
			SetEmptyText();
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
				int iChanged = v.vt == VT_BSTR ? teStrCmpI(pDispParams->rgvarg[nArg].bstrVal, v.bstrVal) : 1;
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
		return teException(pExcepInfo, "FolderView", methodSB, _countof(methodSB), dispIdMember);
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

VOID CteShellBrowser::ExItems(CteFolderItems *pFolderItems, IFolderView *pFV, int nCount, UINT uItem)
{
	BOOL bResultsFolder = ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER]);
	if (uItem & SVGIO_SELECTION) {
		if (m_hwndLV) {
			int nIndex = -1, nFocus = -1;
			if ((uItem & SVGIO_FLAG_VIEWORDER) == 0) {
				nFocus = ListView_GetNextItem(m_hwndLV, -1, LVNI_FOCUSED);
				if (ListView_GetItemState(m_hwndLV, nFocus, LVIS_SELECTED) & LVIS_SELECTED) {
					AddPathEx(pFolderItems, pFV, nFocus, bResultsFolder);
				}
			}
			while ((nIndex = ListView_GetNextItem(m_hwndLV, nIndex, LVNI_SELECTED)) >= 0) {
				if (nIndex != nFocus) {
					AddPathEx(pFolderItems, pFV, nIndex, bResultsFolder);
				}
			}
		} else {
			int i;
			if SUCCEEDED(pFV->GetFocusedItem(&i)) {
				AddPathEx(pFolderItems, pFV, i, bResultsFolder);
			}
		}
	} else {
		for (int i = 0; i < nCount; ++i) {
			AddPathEx(pFolderItems, pFV, i, bResultsFolder);
		}
	}
}

HRESULT CteShellBrowser::Items(UINT uItem, FolderItems **ppid)
{
	if (uItem & SVGIO_SELECTION) {
		if (m_ppDispatch[SB_AltSelectedItems]) {
			return m_ppDispatch[SB_AltSelectedItems]->QueryInterface(IID_PPV_ARGS(ppid));
		}
		if (m_pFolderItem->m_pidlFocused && m_pFolderItem->m_pidl) {
			CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL);
			VARIANT v;
			VariantInit(&v);
			LPITEMIDLIST pidl = ILCombine(m_pFolderItem->m_pidl, m_pFolderItem->m_pidlFocused);
			teSetIDList(&v, pidl);
			teILFreeClear(&pidl);
			pFolderItems->ItemEx(-1, NULL, &v);
			VariantClear(&v);
			*ppid = pFolderItems;
			return S_OK;
		}
	}
	FolderItems *pItems = NULL;
	IDataObject *pDataObj = NULL;
	IFolderView *pFV = NULL;
	int nCount = GetFolderViewAndItemCount(&pFV, uItem);
	if (pFV) {
		if (nCount || !m_hwndLV || (uItem & SVGIO_TYPE_MASK) != SVGIO_SELECTION) {
#ifdef _2000XP
			if (!g_bUpperVista && nCount > 1) {
				BOOL bResultsFolder = ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER]);
				if (bResultsFolder || (uItem & (SVGIO_TYPE_MASK | SVGIO_FLAG_VIEWORDER)) == (SVGIO_SELECTION | SVGIO_FLAG_VIEWORDER)) {
					CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL);
					ExItems(pFolderItems, pFV, nCount, uItem);
					pFV->Release();
					return S_OK;
				}
			}
#endif
			m_pShellView->GetItemObject(uItem, IID_PPV_ARGS(&pDataObj));
		}
	}
	CteFolderItems *pid = new CteFolderItems(pDataObj, pItems);
	if (pFV) {
		if (nCount) {
			long lCount = 0;
			pid->get_Count(&lCount);
			if (lCount != nCount) {
				pid->Clear();
				ExItems(pid, pFV, nCount, uItem);
			}
		}
		pFV->Release();
	}
	if (m_bRegenerateItems) {
		pid->Regenerate(TRUE);
	}
	*ppid = pid;
	SafeRelease(&pDataObj);
	return S_OK;
}

VOID CteShellBrowser::AddPathEx(CteFolderItems *pFolderItems, IFolderView *pFV, int nIndex, BOOL bResultsFolder)
{
	try {
		LPITEMIDLIST pidl;
		if SUCCEEDED(pFV->Item(nIndex, &pidl)) {
			VARIANT v;
			VariantInit(&v);
			LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidl);
			if (bResultsFolder) {
				if SUCCEEDED(teGetDisplayNameFromIDList(&v.bstrVal, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
					v.vt = VT_BSTR;
				}
			} else {
				FolderItem *pid;
				if (GetFolderItemFromIDList(&pid, pidlFull)) {
					teSetObject(&v, pid);
					pid->Release();
				}
			}
			teCoTaskMemFree(pidlFull);
			if (v.vt != VT_EMPTY) {
				pFolderItems->ItemEx(-1, NULL, &v);
				VariantClear(&v);
			}
			teILFreeClear(&pidl);
		}
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"AddPathEx";
#endif
	}
}

#ifdef _2000XP
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

#if defined(USE_SHELLBROWSER) || defined(_2000XP)
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
			m_pFolderItem->MakeUnavailable();
			++nCreate;
			hr = S_FALSE;
		}
	}
	return hr;
}
#endif

#if defined(USE_SHELLBROWSER) || defined(_2000XP)
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
					hr =  ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER]) ? S_OK : m_pSF2->EnumObjects(g_hwndMain, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN | SHCONTF_INCLUDESUPERHIDDEN | SHCONTF_NAVIGATION_ENUM, &peidl);
					if (hr == S_OK) {
						FOLDERSETTINGS fs;
						GetInitFS(&fs);
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
		return hr;
	}
	if (m_pFolderItem->m_pidlFocused && m_pFolderItem->m_pidl) {
		CteFolderItem *pid1 = new CteFolderItem(NULL);
		LPITEMIDLIST pidl = ILCombine(m_pFolderItem->m_pidl, m_pFolderItem->m_pidlFocused);
		hr = pid1->Initialize(m_pidl);
		teILFreeClear(&pidl);
		*ppid = pid1;
		return hr;
	}
	if (m_pdisp) {
		IShellFolderViewDual *pSFVD;
		if SUCCEEDED(m_pdisp->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
			hr = pSFVD->get_FocusedItem(ppid);
			if (hr == S_OK) {
				CteFolderItem *pid1;
				teQueryFolderItemReplace(ppid, &pid1);
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
				if (m_hwndLV && !(dwFlags & (SVSI_SELECT | SVSI_DESELECTOTHERS))) {//21.12.15
					if (nIndex == ListView_GetNextItem(m_hwndLV, -1, LVNI_SELECTED)) {
						if (ListView_GetNextItem(m_hwndLV, nIndex, LVNI_SELECTED) < 0) {
							dwFlags |= SVSI_DESELECTOTHERS;
						}
					}
				}
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
		if (m_pFolderItem) {
			teILFreeClear(&m_pFolderItem->m_pidlFocused);
		}
	}
	return hr;
}

HRESULT CteShellBrowser::SelectItemEx(LPITEMIDLIST pidl, int dwFlags)
{
	HRESULT hr = E_FAIL;
	if (m_pShellView) {
		LPITEMIDLIST pidlLast = ILFindLastID(pidl);

		if (m_hwndLV && !(dwFlags & (SVSI_SELECT | SVSI_DESELECTOTHERS))) {//21.12.15
			int nSel = ListView_GetNextItem(m_hwndLV, -1, LVNI_SELECTED);
			if (nSel >= 0) {
				IFolderView *pFV = NULL;
				if (GetFolderViewAndItemCount(&pFV, SVGIO_SELECTION) == 1) {
					LPITEMIDLIST pidlSel = NULL;
					if SUCCEEDED(pFV->Item(nSel, &pidlSel)) {
						if (ILIsEqual(pidlLast, pidlSel)) {
							dwFlags |= SVSI_DESELECTOTHERS;
						}
						teILFreeClear(&pidlSel);
					}
				}
				SafeRelease(&pFV);
			}
		}
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
	if (!m_pFolderItem1 && !m_pFolderItem->m_dwUnavailable && ILIsEqual(pidlFolder, g_pidls[CSIDL_RESULTSFOLDER])) {
		m_nSuspendMode = 4;
		return E_FAIL;
	}
	if (ILIsEqual(pidlFolder, g_pidls[CSIDL_UNAVAILABLE])) {
		if (teCompareSFClass(m_pSF2, &CLSID_ResultsFolder)) {
			return S_OK;
		}
	}
	return OnNavigationPending2((LPITEMIDLIST)pidlFolder);
}

HRESULT CteShellBrowser::OnNavigationPending2(LPITEMIDLIST pidlFolder)
{
	LPITEMIDLIST pidlPrevius = m_pidl;
	m_pidl = ::ILClone(pidlFolder);
	CteFolderItem *pPrevious = m_pFolderItem;
	if (m_pFolderItem1) {
		m_pFolderItem = m_pFolderItem1;
		m_pFolderItem1 = NULL;
	} else if (!m_pFolderItem->m_dwUnavailable) {
		m_pFolderItem = new CteFolderItem(NULL);
		m_pFolderItem->Initialize(m_pidl);
	}

	HRESULT hr = OnBeforeNavigate(pPrevious, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
	if FAILED(hr) {
		m_uLogIndex = m_uPrevLogIndex;
		teCoTaskMemFree(m_pidl);
		m_pidl = pidlPrevius;
		GetShellFolder2(&m_pidl);
		FolderItem *pid = m_pFolderItem;
		m_pFolderItem = pPrevious;
		if (hr == E_ABORT || teILIsBlank(m_pFolderItem)) {
			if (Close(FALSE)) {
				return hr;
			}
		}
		if (hr == E_ACCESSDENIED) {
			HRESULT hr2 = BrowseObject2(pid, SBSP_NEWBROWSER | SBSP_ABSOLUTE);
			if (ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
				hr = hr2;
			}
		}
		InitFilter();
		InitFolderSize();
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
	if (m_pExplorerBrowser) {
		FOLDERSETTINGS fs;
		GetInitFS(&fs);
		m_pExplorerBrowser->SetFolderSettings(&fs);
	}
	if (m_bAutoVM) {
		IFolderView2 *pFV2;
		if (m_pShellView && SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2)))) {
			pFV2->GetViewModeAndIconSize((FOLDERVIEWMODE *)&m_param[SB_ViewMode], (int *)&m_param[SB_IconSize]);
			pFV2->Release();
			if (m_param[SB_ViewMode] == FVM_ICON && m_param[SB_IconSize] < 96) {
				m_param[SB_IconSize] = 96;
			}
		}
	}
	SetViewModeAndIconSize(TRUE);
	m_dwRedraw |= 4;
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
				if (teStrCmpI(bs, bs2) && tePathMatchSpec(bs2, m_bsFilter)) {
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
			int nFocused = ListView_GetNextItem(m_hwndLV, -1, LVNI_FOCUSED | LVNI_SELECTED);
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
		CteFolderItem *pid = new CteFolderItem(NULL);
		pid->m_v.vt = VT_BSTR;
		pid->m_v.bstrVal = NULL;
		teCreateSearchPath(&pid->m_v.bstrVal, bsPath, pszSearch);
		teSysFreeString(&bsPath);
		BrowseObject2(pid, SBSP_SAMEBROWSER);
		pid->Release();
		return S_OK;
	}
	return E_NOTIMPL;
}
/*
	teSysFreeString(&bsPath);
	if (m_pdisp) {
		IShellFolderViewDual3 *pSFVD3;
		if SUCCEEDED(m_pdisp->QueryInterface(IID_PPV_ARGS(&pSFVD3))) {
			pSFVD3->FilterView(pszSearch);
			pSFVD3->Release();
		}
	}
	return S_OK;
*/

VOID CteShellBrowser::RedrawUpdate()
{
	if (m_dwRedraw & 3) {
		m_dwRedraw &= ~3;
		SetRedraw(TRUE);
		RedrawWindow(m_hwnd, NULL, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
	}
}

VOID CteShellBrowser::SetEmptyText()
{
	if (m_pShellView) {
		IFolderView2 *pFV2;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
			WCHAR pszFormat[MAX_STATUS];
			pszFormat[0] = NULL;
			LPWSTR lpBuf = pszFormat;
			if (m_pFolderItem->m_dwUnavailable) {
				if (GetTickCount() - m_pFolderItem->m_dwUnavailable > 500) {
					LoadString(g_hShell32, 4157, pszFormat, MAX_STATUS);
				}
			} else if (!m_bFiltered) {
				lpBuf = NULL;
			}
			pFV2->SetText(FVST_EMPTYTEXT, lpBuf);
			pFV2->Release();
		}
	}
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
		/*else if (m_pExplorerBrowser && teCompareSFClass(m_pSF2, &CLSID_ResultsFolder)) {
			if (g_pAutomation) {
				m_hwndDT = FindWindowExA(m_hwndDV, NULL, "DirectUIHWND", NULL);
				IUIAutomationElement *pElement;
				if SUCCEEDED(g_pAutomation->ElementFromHandle(m_hwndDT, &pElement)) {
					pElement->Release();
				}
			}
		}*/
	}
	GetShellFolderView();
	if (m_pExplorerBrowser) {
		IUnknown_GetWindow(m_pExplorerBrowser, &m_hwnd);
	}
	m_bNavigateComplete = TRUE;
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
	m_dwRedraw &= ~4;
	SetRedraw(TRUE);
	if (!m_pTC->m_nLockUpdate) {
		ArrangeWindowEx();
	}
	if (m_pFolderItem->m_dwUnavailable) {
		SetEmptyText();
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
	if (m_bVisible && !IsWindowVisible(m_hwnd)) {
		ShowWindow(m_hwnd, SW_SHOWNA);
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
	m_dwRedraw &= ~4;
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

HRESULT CteShellBrowser::BrowseToObject(DWORD dwFrame, FOLDERVIEWOPTIONS fvoFlags)
{
	HRESULT hr;
	m_nSuspendMode = 2;
	::InterlockedIncrement(&m_pTC->m_lNavigating);
	try {
		IFolderViewOptions *pOptions;
		if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
			pOptions->SetFolderViewOptions(FVO_VISTALAYOUT, fvoFlags);
			pOptions->Release();
		}
		if (!m_hwnd) {
			IEnumIDList *peidl = NULL;
			hr = m_pSF2->EnumObjects(NULL, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS, &peidl);
			if (hr == S_OK) {
				peidl->Release();
			} else if (hr != S_FALSE) {
				m_pFolderItem->MakeUnavailable();
				teILCloneReplace(&m_pidl, g_pidls[CSIDL_UNAVAILABLE]);
				GetShellFolder2(&m_pidl);
			}
			if (dwFrame && g_bDarkMode) {
				if (teCompareSFClass(m_pSF2, &CLSID_ResultsFolder)) {
					m_pExplorerBrowser->BrowseToIDList(g_pidls[CSIDL_UNAVAILABLE], SBSP_ABSOLUTE);
				}
			}
		}
		hr = m_pExplorerBrowser->BrowseToObject(m_pSF2, SBSP_ABSOLUTE);
	} catch (...) {
		hr = E_FAIL;
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"BrowseToObject";
#endif
	}
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
	SafeRelease(&m_pSF2);
	if (pSF) {
		hr = pSF->QueryInterface(IID_PPV_ARGS(&m_pSF2));
		pSF->Release();
	}
	if (!m_pSF2) {
		GetShellFolder(&pSF, g_pidls[CSIDL_UNAVAILABLE]);
		teILCloneReplace(ppidl, g_pidls[CSIDL_UNAVAILABLE]);
		m_pFolderItem->MakeUnavailable();
		hr = pSF->QueryInterface(IID_PPV_ARGS(&m_pSF2));
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
		m_pFolderItem->m_dwSessionId = MAKELONG(GetTickCount(), rand());
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
				if (iImageSize != m_param[SB_IconSize] && uViewMode == m_param[SB_ViewMode] && !m_pFolderItem->m_pidlFocused) {
					GetFocusedItem(&m_pFolderItem->m_pidlFocused);
				}
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
				ListView_EnsureVisible(m_hwndLV, ListView_GetNextItem(m_hwndLV, -1, LVNI_FOCUSED), FALSE);
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
				m_dwRedraw |= 2;
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
					} else if (!g_bUpperVista && ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
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

#if defined(USE_SHELLBROWSER) || defined(_2000XP)
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
		WCHAR szBuf[MAX_PATH];
		szBuf[0] = NULL;
		if (pidl) {
			if (iColumn == m_nColumns - 1) {
				m_nFolderSizeIndex = MAXINT - 1;
				SetFolderSize(m_pSF2, pidl, szBuf, _countof(szBuf));
			} else {
				SetLabel(pidl, szBuf, _countof(szBuf));
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
								RegCloseKey(hKey);
							}
							IEnumIDList *peidl;
							if (m_pSF2->EnumObjects(NULL, grfFlags, &peidl) == S_OK) {
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
	return teGetIDListFromObject(static_cast<FolderItem *>(m_pFolderItem), ppidl) ? S_OK : E_FAIL;
}

//

VOID CteShellBrowser::Suspend(int nMode)
{
	BOOL bVisible = m_bVisible;
	if (nMode == 2) {
		if (m_pFolderItem->m_dwUnavailable || ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
			if (IsWindowEnabled(m_hwnd)) {
				return;
			}
		}
		Show(FALSE, 0);
		m_pFolderItem->MakeUnavailable();
		teILCloneReplace(&m_pidl, g_pidls[CSIDL_UNAVAILABLE]);
	}
	if (!m_pFolderItem->m_dwUnavailable) {
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
				if (m_pFolderItem->m_v.vt == VT_BSTR) {
					teILFreeClear(&m_pidl);
					m_pFolderItem->Clear();
				}
			}
			if (m_pidl) {
				VARIANT v;
				if (GetVarPathFromFolderItem(m_pFolderItem, &v)) {
					if (nMode == 2) {
						m_pFolderItem->MakeUnavailable();
						m_nUnload = 8;
					} else {
						m_pFolderItem->Clear();
						VariantClear(&m_pFolderItem->m_v);
						VariantCopy(&m_pFolderItem->m_v, &v);
					}
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
			SetWindowSubclass(m_hwndDV, TELVProc, (UINT_PTR)TELVProc, 0);
			for (int i = WM_USER + 173; i <= WM_USER + 175; ++i) {
				teChangeWindowMessageFilterEx(m_hwndDV, i, MSGFLT_ALLOW, NULL);
			}
/*			HWND hwndDUI = FindWindowExA(GetParent(m_hwndDV), NULL, "DUIViewWndClassName", NULL);
			if (hwndDUI) {
				if (hwndDUI = FindWindowExA(hwndDUI, NULL, "DirectUIHWND", NULL)) {
					SendMessage(hwndDUI, WM_SETTINGCHANGE, 0, (LPARAM)L"ShellState");
				}
			}*/
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
				SetWindowSubclass(m_hwndLV, TELVProc2, (UINT_PTR)TELVProc2, 0);
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
	RemoveWindowSubclass(m_hwndLV, TELVProc2, (UINT_PTR)TELVProc2);
	SetWindowLongPtr(m_hwndLV, GWLP_USERDATA, 0);
	RemoveWindowSubclass(m_hwndDV, TELVProc, (UINT_PTR)TELVProc);
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
						if (m_pFolderItem->m_v.vt == VT_BSTR) {
							m_pFolderItem->Clear();
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
				if ((dwOptions & 1) && m_pFolderItem->m_dwUnavailable && !m_nCreate && !ILIsEqual(m_pidl, g_pidls[CSIDL_RESULTSFOLDER])) {
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
	if (m_pFolderItem->m_dwUnavailable) {
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
	if (teStrCmpIWA(bs, "System.Null") == 0) {
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
			} else if (IsEqualPropertyKey(pCol1[0].propkey, PKEY_Contact_Label)) {
				SetLabelProperties();
			} else if (IsEqualPropertyKey(pCol1[0].propkey, PKEY_TotalFileSize)) {
				SetFolderSizeProperties();
			} else {
				BSTR bs = tePSGetNameFromPropertyKeyEx(pCol1[0].propkey, 0, NULL);
				if (bs) {
					VARIANT v;
					VariantInit(&v);
					if (m_ppDispatch[SB_ColumnsReplace]) {
						teGetProperty(m_ppDispatch[SB_ColumnsReplace], bs, &v);
					}
					if (v.vt == VT_EMPTY && g_pOnFunc[TE_ColumnsReplace]) {
						teGetProperty(g_pOnFunc[TE_ColumnsReplace], bs, &v);
					}
					teSysFreeString(&bs);
					IDispatch *pdisp;
					if (GetDispatch(&v, &pdisp)) {
						SetReplaceColumnsProperties(pdisp, pCol1[0].propkey);
					}
					VariantClear(&v);
				}
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
	if (dir != dir1 || teStrCmpI(szNew, szOld)) {
		IShellFolderView *pSFV;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
			int i = PSGetColumnIndexXP(szNew, NULL);
			if (i >= 0) {
				pSFV->Rearrange(i);
				if (dir && teStrCmpI(szNew, szOld)) {
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
	if (szNew[0] && teStrCmpIWA(szNew, "System.Null")) {
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
				if (teStrCmpI(szName, szNew) == 0) {
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
	BOOL b = !(m_dwRedraw & 4) && (bRedraw || !(m_dwRedraw & 3));
	SendMessage(m_hwnd, WM_SETREDRAW, b, 0);
	if (b) {
		m_dwRedraw &= ~1;
		if (m_hwndAlt) {
			BringWindowToTop(m_hwndAlt);
		}
	} else {
		m_dwRedraw |= 1;
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	teGetDispId(NULL, MAP_TE, *rgszNames, rgDispId, FALSE);
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
								{
									auto itr = g_umSubWindows.find(m_hwnd);
									teSetObject(pVarResult, itr != g_umSubWindows.end() ? itr->second : g_pWebBrowser);
								}
								break;
							case CTRL_FI:
								if (nArg >= 1) {
									DWORD dwSessionId = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
									for (size_t i = 0; i < g_ppENum.size(); ++i) {
										if (dwSessionId == g_ppENum[i]->m_dwSessionId) {
											teSetObject(pVarResult, g_ppENum[i]);
											break;
										}
									}
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
								for (auto itr = g_umSubWindows.begin(); itr != g_umSubWindows.end(); ++itr) {
									teArrayPush(pArray, itr->second);
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
				teSetDarkMode(g_hwndMain);
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
						if (teStrCmpIWA(lpMethod, "GetArchive") == 0) {
							teAddRemoveProc(&g_ppGetArchive, lpProc, dispIdMember == TE_METHOD + 1025);
						} else if (teStrCmpIWA(lpMethod, "GetImage") == 0) {
							teAddRemoveProc(&g_ppGetImage, lpProc, dispIdMember == TE_METHOD + 1025);
						} else if (teStrCmpIWA(lpMethod, "MessageSFVCB") == 0) {
							teAddRemoveProc(&g_ppMessageSFVCB, lpProc, dispIdMember == TE_METHOD + 1025);
						}
					}
				}
				return S_OK;

			case 1030://WindowsAPI
#ifdef USE_OBJECTAPI
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
				BSTR bs;
				bs = ::SysAllocString(L"CreateObject");
				teSetObjectRelease(&v, new CteWindowsAPI(&dispAPI[teUMSearch(MAP_API, bs)]));
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

#ifdef USE_TESTOBJECT
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
								auto itr = g_umSubWindows.find(hwnd);
								if (itr != g_umSubWindows.end()) {
									itr->second->Release();
								}
								CteWebBrowser *pWB = new CteWebBrowser(hwnd, v.bstrVal, &pDispParams->rgvarg[nArg - 2]);
								g_umSubWindows[hwnd] = pWB;
								teSetObjectRelease(pVarResult, pWB);
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
		return teException(pExcepInfo, "external", methodTE, _countof(methodTE), dispIdMember);
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	teGetDispId(methodWB, _countof(methodWB), *rgszNames, rgDispId, FALSE);
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
		return teException(pExcepInfo, "WebBrowser", methodWB, _countof(methodWB), dispIdMember);
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
				auto itr = g_umSubWindows.find(m_hwndParent);
				if (itr != g_umSubWindows.end()) {
					g_umSubWindows.erase(itr);
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
		DoFunc(TE_OnVisibleChanged, this, E_NOTIMPL);
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
		RemoveWindowSubclass(m_hwnd, TETCProc, (UINT_PTR)TETCProc);
		DestroyWindow(m_hwnd);
		m_hwnd = NULL;

		RevokeDragDrop(m_hwndButton);
		RemoveWindowSubclass(m_hwndButton, TEBTProc, (UINT_PTR)TEBTProc);
		DestroyWindow(m_hwndButton);
		m_hwndButton = NULL;

		RemoveWindowSubclass(m_hwndStatic, TESTProc, (UINT_PTR)TESTProc);
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
	SetWindowSubclass(m_hwnd, TETCProc, (UINT_PTR)TETCProc, 0);
	RegisterDragDrop(m_hwnd, m_pDropTarget2);
	BringWindowToTop(m_hwnd);
}

BOOL CteTabCtrl::Create()
{
	m_hwndStatic = CreateWindowEx(
		WS_EX_TOPMOST, WC_STATIC, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | SS_NOTIFY,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, g_hwndMain, (HMENU)0, hInst, NULL);
	SetWindowLongPtr(m_hwndStatic, GWLP_USERDATA, (LONG_PTR)this);
	SetWindowSubclass(m_hwndStatic, TESTProc, (UINT_PTR)TESTProc, 0);

	m_hwndButton = CreateWindowEx(
		WS_EX_TOPMOST, WC_BUTTON, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | BS_OWNERDRAW,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
		m_hwndStatic, (HMENU)0, hInst, NULL);

	SetWindowLongPtr(m_hwndButton, GWLP_USERDATA, (LONG_PTR)this);
	SetWindowSubclass(m_hwndButton, TEBTProc, (UINT_PTR)TEBTProc, 0);
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	teGetDispId(NULL, MAP_TC, *rgszNames, rgDispId, TRUE);
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
									RemoveWindowSubclass(m_hwnd, TETCProc, (UINT_PTR)TETCProc);
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
		return teException(pExcepInfo, "TabControl", methodTC, _countof(methodTC), dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteTabCtrl::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	return teGetDispId(NULL, MAP_TC, bstrName, pid, TRUE);
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	teGetDispId(methodCM, _countof(methodCM), *rgszNames, rgDispId, FALSE);
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
							if (lstrcmpiA(szNameA, "mount") == 0) {
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
		return teException(pExcepInfo, "ContextMenu", methodCM, _countof(methodCM), dispIdMember);
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
		RemoveWindowSubclass(m_hwnd, TETVProc, (UINT_PTR)TETVProc);
	}
#endif
	RemoveWindowSubclass(m_hwndTV, TETVProc2, (UINT_PTR)TETVProc2);
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
	if (EMULATE_XP SUCCEEDED(teCreateInstance(CLSID_NamespaceTreeControl, NULL, NULL, IID_PPV_ARGS(&m_pNameSpaceTreeControl)))) {
		RECT rc;
		SetRectEmpty(&rc);
		if SUCCEEDED(m_pNameSpaceTreeControl->Initialize(m_hwndParent, &rc, m_param[SB_TreeFlags] & ~(NSTCS_CHECKBOXES | NSTCS_PARTIALCHECKBOXES | NSTCS_EXCLUSIONCHECKBOXES | NSTCS_DIMMEDCHECKBOXES | NSTCS_NOINDENTCHECKS | NSTCS_NOREPLACEOPEN))) {
			m_pNameSpaceTreeControl->TreeAdvise(static_cast<INameSpaceTreeControlEvents *>(this), &m_dwCookie);
			if (IUnknown_GetWindow(m_pNameSpaceTreeControl, &m_hwnd) == S_OK) {
				m_hwndTV = FindTreeWindow(m_hwnd);
				if (m_hwndTV) {
					teSetTreeTheme(m_hwndTV, SendMessage(m_hwndTV, TVM_GETBKCOLOR, 0, 0));
					SetWindowLongPtr(m_hwndTV, GWLP_USERDATA, (LONG_PTR)this);
					SetWindowSubclass(m_hwndTV, TETVProc2, (UINT_PTR)TETVProc2, 0);
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
			SetWindowSubclass(hwndParent, TETVProc, (UINT_PTR)TETVProc, 0);
			SetWindowLongPtr(m_hwndTV, GWLP_USERDATA, (LONG_PTR)this);
			SetWindowSubclass(m_hwndTV, TETVProc2, (UINT_PTR)TETVProc2, 0);
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	teGetDispId(NULL, MAP_TV, *rgszNames, rgDispId, FALSE);
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
				int n = GetBoolFromVariant(&pDispParams->rgvarg[nArg]) ? 3 : 1;
				BOOL bChanged = m_param[SB_TreeAlign] != n;
				m_param[SB_TreeAlign] = n;
				ArrangeWindow();
				Show();
				if (bChanged) {
					DoFunc(TE_OnVisibleChanged, this, E_NOTIMPL);
				}
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
					if (m_pNameSpaceTreeControl) {
						IShellItem *pShellItem;
						if (teCreateItemFromVariant(&pDispParams->rgvarg[nArg], &pShellItem)) {
							teSetLong(pVarResult, m_pNameSpaceTreeControl->GetItemRect(pShellItem, prc));
							pShellItem->Release();
						}
					}
				}
				teWriteBack(&pDispParams->rgvarg[nArg - 1], &vMem);
			}
			return S_OK;

		case TE_METHOD + 0x1300://Notify
			if (nArg >= 2 && m_hwndTV && m_pNameSpaceTreeControl) {
				long lEvent = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (lEvent & (SHCNE_MKDIR | SHCNE_MEDIAINSERTED | SHCNE_DRIVEADD | SHCNE_NETSHARE)) {
					IShellItem *psi, *psiParent;
					if (teCreateItemFromVariant(&pDispParams->rgvarg[nArg - 1], &psi)) {
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
		return teException(pExcepInfo, "TreeView", methodTV, _countof(methodTV), dispIdMember);
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
	SendMessage(m_hwnd, WM_SETREDRAW, FALSE, 0);
	SetTimer(m_hwndTV, TET_Redraw, 100, teTimerProcForTree);
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
	SendMessage(m_hwnd, WM_SETREDRAW, FALSE, 0);
	SetTimer(m_hwndTV, TET_Redraw, 100, teTimerProcForTree);
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::OnItemAdded(IShellItem *psi, BOOL fIsRoot)
{
	SendMessage(m_hwnd, WM_SETREDRAW, FALSE, 0);
	SetTimer(m_hwndTV, TET_Redraw, 100, teTimerProcForTree);
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
	if (teStrCmpI(v.bstrVal, m_bsRoot)) {
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
					if (teCreateItemFromPath(pszPath, &pShellItem)) {
						hr = m_pNameSpaceTreeControl->AppendRoot(pShellItem, m_param[SB_EnumFlags] | SHCONTF_ENABLE_ASYNC, m_param[SB_RootStyle], this);
						pShellItem->Release();
						if SUCCEEDED(hr) {
							++nRoots;
						}
					}
					pszPath = &bs[i + 1];
				}
			}
		}
		teSysFreeString(&bs);
		if (!nRoots) {
			if SUCCEEDED(_SHCreateItemFromIDList(g_pidls[CSIDL_DESKTOP], IID_PPV_ARGS(&pShellItem))) {
				hr = m_pNameSpaceTreeControl->AppendRoot(pShellItem, m_param[SB_EnumFlags] | SHCONTF_ENABLE_ASYNC, m_param[SB_RootStyle], this);
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

//CteActiveScriptSite
CteActiveScriptSite::CteActiveScriptSite(IUnknown *punk, EXCEPINFO *pExcepInfo)
{
	m_cRef = 1;
	m_pDispatchEx = NULL;
	if (punk) {
		punk->QueryInterface(IID_PPV_ARGS(&m_pDispatchEx));
	}
	m_pExcepInfo = pExcepInfo;
	m_hr = S_OK;
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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
			if (g_pWebBrowser && teStrCmpIWA(pstrName, "window") == 0) {
				teGetHTMLWindow(g_pWebBrowser->m_pWebBrowser, IID_PPV_ARGS(ppiunkItem));
				return S_OK;
			}
		} else if (teStrCmpIWA(pstrName, "api") == 0) {
			*ppiunkItem = new CteWindowsAPI(NULL);
			return S_OK;
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
	EXCEPINFO ExcepInfo;
	if (!pscripterror) {
		return E_POINTER;
	}
	EXCEPINFO *pExcepInfo = m_pExcepInfo ? m_pExcepInfo : &ExcepInfo;
	if SUCCEEDED(pscripterror->GetExceptionInfo(pExcepInfo)) {
		m_hr = pExcepInfo->scode;
		BSTR bsSrcText = NULL;
		pscripterror->GetSourceLineText(&bsSrcText);
		DWORD dw1 = 0;
		ULONG ulLine = 0;
		LONG ulColumn = 0;
		pscripterror->GetSourcePosition(&dw1, &ulLine, &ulColumn);
		WCHAR pszLine[64];
		swprintf_s(pszLine, _countof(pszLine), L"\n(0x%08x) Line: %d\n", m_hr, ulLine);
		BSTR bs = ::SysAllocStringLen(NULL, ::SysStringLen(pExcepInfo->bstrDescription) + ::SysStringLen(bsSrcText) + lstrlen(pszLine));
		lstrcpy(bs, pExcepInfo->bstrDescription);
		lstrcat(bs, pszLine);
		if (bsSrcText) {
			lstrcat(bs, bsSrcText);
			::SysFreeString(bsSrcText);
		}
		teSysFreeString(&pExcepInfo->bstrDescription);
		pExcepInfo->bstrDescription = bs;
#ifdef _DEBUG
		::OutputDebugString(bs);
#endif
		if (!m_pExcepInfo) {
			MessageBox(NULL, bs, _T(PRODUCTNAME), MB_OK | MB_ICONERROR);
			teSysFreeString(&pExcepInfo->bstrDescription);
			teSysFreeString(&pExcepInfo->bstrHelpFile);
			teSysFreeString(&pExcepInfo->bstrSource);
		}
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

#endif
