// TE.cpp
// Tablacus Explorer (C)2011- Gaku
// MIT Lisence
// Visual C++ 2008 Express Edition
// Windows SDK v7.0 
// http://www.eonet.ne.jp/~gakana/tablacus/

#include "stdafx.h"
#include <windows.h>
#include "TE.h"

// Global Variables:
HINSTANCE hInst;								// current instance
WCHAR g_szTE[MAX_LOADSTRING];
HWND g_hwndMain = NULL;
HWND g_hwndBrowser = NULL;
LONG g_nSize = MAXWORD;
LONG g_nLockUpdate = 0;
CteTabs *g_pTabs = NULL;
CteTabs *g_pTC[MAX_TC];
IShellFolder *g_pDesktopFolder = NULL;
HINSTANCE	g_hShell32 = NULL;
HINSTANCE	g_hKernel32 = NULL;
HINSTANCE	g_hCrypt32 = NULL;
LPFNSHCreateItemFromIDList lpfnSHCreateItemFromIDList = NULL;
LPFNSetDllDirectoryW lpfnSetDllDirectoryW = NULL;
LPFNSHParseDisplayName lpfnSHParseDisplayName = NULL;
LPFNSHRunDialog lpfnSHRunDialog = NULL;
LPFNCryptBinaryToStringW lpfnCryptBinaryToStringW = NULL;
DWORD	g_dragdrop = 0;
UINT	*g_pCrcTable = NULL;
LPITEMIDLIST g_pidls[MAX_CSIDL];
DWORD		g_paramFV[SB_Count];

int	g_nTCCount = 0;
int	g_nTCIndex = 0;
LONG nUpdate = 0;
HWND			g_hDefLV = 0;
WNDPROC			g_DefTCProc = NULL;
WNDPROC			g_DefBTProc = NULL;
CTE *g_pTE;
CteWebBrowser *g_pWebBrowser;
CteShellBrowser *g_pSB[MAX_FV];
GUID		g_ClsIdStruct;
GUID		g_ClsIdContextMenu;
GUID		g_ClsIdSB;
GUID		g_ClsIdTV;
GUID		g_ClsIdTC;
GUID		g_ClsIdFI;

IDispatch			*g_pOnFunc[Count_OnFunc];
IShellDispatch		*g_psha = NULL;
IShellWindows		*g_pSW = NULL;
//IHTMLDocument2		*g_pHtmlDoc = NULL;
IDispatch			*g_pJS = NULL;
IDispatch			*g_pObject  = NULL;
IDispatch			*g_pArray   = NULL;

FORMATETC HDROPFormat  = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC IDLISTFormat = {0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC DROPEFFECTFormat  = {0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC UNICODEFormat  = {CF_UNICODETEXT, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};

long		g_nProcKey	   = 0;
long		g_nProcMouse   = 0;
long		g_nProcDrag    = 0;
long		g_nProcFV      = 0;
long		g_nProcTV      = 0;
long		g_nThreads	   = 0;
WPARAM		g_LastMsg;
HWND		g_hwndLast;
BOOL		g_bInit = true;
CteContextMenu *g_pCCM = NULL;
ULONG_PTR g_Token;
Gdiplus::GdiplusStartupInput g_StartupInput;
HHOOK g_hHook;
HHOOK g_hMouseHook;
HHOOK g_hMenuKeyHook = NULL;
HMENU g_hMenu = NULL;
int g_nCmdShow = SW_SHOWNORMAL;
HANDLE g_hMutex;
WCHAR g_szText[MAX_PATH];
int			g_nAvgWidth = 7;
BOOL		g_bAvgWidth = true;
OSVERSIONINFO osInfo;
BOOL	g_bReload = false;
BSTR	g_bsCmdLine = NULL;

int *g_maps[MAP_LENGTH];

method structSize[] = {
	{ sizeof(BITMAP), L"BITMAP" },
	{ sizeof(BSTR), L"BSTR" },
	{ sizeof(BYTE), L"BYTE" },
	{ sizeof(char), L"char" },
	{ sizeof(CHOOSECOLOR), L"CHOOSECOLOR" },
	{ sizeof(CHOOSEFONT), L"CHOOSEFONT" },
	{ sizeof(COPYDATASTRUCT), L"COPYDATASTRUCT" },
	{ sizeof(DWORD), L"DWORD" },
	{ sizeof(DIBSECTION), L"DIBSECTION" },
	{ sizeof(EncoderParameter), L"EncoderParameter" },
	{ sizeof(EncoderParameter), L"EncoderParameters" },
	{ sizeof(EXCEPINFO), L"EXCEPINFO" },
	{ sizeof(FINDREPLACE), L"FINDREPLACE" },
	{ sizeof(FOLDERSETTINGS), L"FOLDERSETTINGS" },
	{ sizeof(GUID), L"GUID" },
	{ sizeof(HANDLE), L"HANDLE" },
	{ sizeof(ICONINFO), L"ICONINFO" },
	{ sizeof(ICONMETRICS), L"ICONMETRICS" },
	{ sizeof(int), L"int" },
	{ sizeof(KEYBDINPUT) + sizeof(DWORD), L"KEYBDINPUT" },
	{ 256, L"KEYSTATE" },
	{ sizeof(LOGFONT), L"LOGFONT" },
	{ sizeof(LPWSTR), L"LPWSTR" },
	{ sizeof(LVBKIMAGE), L"LVBKIMAGE" },
	{ sizeof(LVFINDINFO), L"LVFINDINFO" },
	{ sizeof(LVGROUP), L"LVGROUP" },
	{ sizeof(LVHITTESTINFO), L"LVHITTESTINFO" },
	{ sizeof(LVITEM), L"LVITEM" },
	{ sizeof(MENUITEMINFO), L"MENUITEMINFO" },
	{ sizeof(MOUSEINPUT) + sizeof(DWORD), L"MOUSEINPUT" },
	{ sizeof(MSG), L"MSG" },
	{ sizeof(NONCLIENTMETRICS), L"NONCLIENTMETRICS" },
	{ sizeof(NOTIFYICONDATA), L"NOTIFYICONDATA" },
	{ sizeof(NMHDR), L"NMHDR" },
	{ sizeof(OSVERSIONINFO), L"OSVERSIONINFO" },
	{ sizeof(OSVERSIONINFOEX), L"OSVERSIONINFOEX" },
	{ sizeof(PAINTSTRUCT), L"PAINTSTRUCT" },
	{ sizeof(POINT), L"POINT" },
	{ sizeof(RECT), L"RECT" },
	{ sizeof(SHELLEXECUTEINFO), L"SHELLEXECUTEINFO" },
	{ sizeof(SHFILEINFO), L"SHFILEINFO" },
	{ sizeof(SHFILEOPSTRUCT), L"SHFILEOPSTRUCT" },
	{ sizeof(SIZE), L"SIZE" },
	{ sizeof(TCHITTESTINFO), L"TCHITTESTINFO" },
	{ sizeof(TCITEM), L"TCITEM" },
	{ sizeof(TVHITTESTINFO), L"TVHITTESTINFO" },
	{ sizeof(TVITEM), L"TVITEM" },
	{ sizeof(WIN32_FIND_DATA), L"WIN32_FIND_DATA" },
	{ sizeof(VARIANT), L"VARIANT" },
	{ sizeof(WCHAR), L"WCHAR" },
	{ sizeof(WORD), L"WORD" },
};

method methodTE[] = {
	{ 1001, L"Data" },
	{ 1002, L"hwnd" },
	{ 1004, L"About" },
	{ TE_METHOD + 1005, L"Ctrl" },
	{ TE_METHOD + 1006, L"Ctrls" },
	{ 1007, L"Version" },
	{ TE_METHOD + 1008, L"ClearEvents" },
	{ TE_METHOD + 1009, L"Reload" },
	{ TE_METHOD + 1010, L"CreateObject" },
	{ TE_METHOD + 1020, L"GetObject" },
	{ 1030, L"WindowsAPI" },
	{ 1131, L"CommonDialog" },
	{ 1132, L"GdiplusBitmap" },
	{ TE_METHOD + 1133, L"FolderItems" },
	{ TE_METHOD + 1134, L"Object" },
	{ TE_METHOD + 1135, L"Array" },
	{ TE_METHOD + 1136, L"Collection" },
	{ TE_METHOD + 1050, L"CreateCtrl" },
	{ TE_METHOD + 1040, L"CtrlFromPoint" },
	{ TE_METHOD + 1060, L"MainMenu" },
	{ TE_METHOD + 1070, L"CtrlFromWindow" },
	{ TE_METHOD + 1080, L"LockUpdate" },
	{ TE_METHOD + 1090, L"UnlockUpdate" },
	{ TE_METHOD + 1100, L"HookDragDrop" },
#ifdef _USE_TESTOBJECT
	{ 1200, L"TestObj" },
#endif
	{ TE_OFFSET + TE_Type   , L"Type" },
	{ TE_OFFSET + TE_Left   , L"offsetLeft" },
	{ TE_OFFSET + TE_Top    , L"offsetTop" },
	{ TE_OFFSET + TE_Right  , L"offsetRight" },
	{ TE_OFFSET + TE_Bottom , L"offsetBottom" },
	{ TE_OFFSET + TE_Tab, L"Tab" },
	{ TE_OFFSET + TE_CmdShow, L"CmdShow" },
	{ START_OnFunc + TE_OnBeforeNavigate, L"OnBeforeNavigate" },
	{ START_OnFunc + TE_OnViewCreated, L"OnViewCreated" },
	{ START_OnFunc + TE_OnKeyMessage, L"OnKeyMessage" },
	{ START_OnFunc + TE_OnMouseMessage, L"OnMouseMessage" },
	{ START_OnFunc + TE_OnCreate, L"OnCreate" },
	{ START_OnFunc + TE_OnDefaultCommand, L"OnDefaultCommand" },
	{ START_OnFunc + TE_OnItemClick, L"OnItemClick" },
	{ START_OnFunc + TE_OnGetPaneState, L"OnGetPaneState" },
	{ START_OnFunc + TE_OnMenuMessage, L"OnMenuMessage" },
	{ START_OnFunc + TE_OnSystemMessage, L"OnSystemMessage" },
	{ START_OnFunc + TE_OnShowContextMenu, L"OnShowContextMenu" },
	{ START_OnFunc + TE_OnSelectionChanged, L"OnSelectionChanged" },
	{ START_OnFunc + TE_OnClose, L"OnClose" },
	{ START_OnFunc + TE_OnDragEnter, L"OnDragEnter" },
	{ START_OnFunc + TE_OnDragOver, L"OnDragOver" },
	{ START_OnFunc + TE_OnDrop, L"OnDrop" },
	{ START_OnFunc + TE_OnDragLeave, L"OnDragLeave" },
	{ START_OnFunc + TE_OnAppMessage, L"OnAppMessage" },
	{ START_OnFunc + TE_OnStatusText, L"OnStatusText" },
	{ START_OnFunc + TE_OnToolTip, L"OnToolTip" },
	{ START_OnFunc + TE_OnNewWindow, L"OnNewWindow" },
	{ START_OnFunc + TE_OnWindowRegistered, L"OnWindowRegistered" },
	{ START_OnFunc + TE_OnSelectionChanging, L"OnSelectionChanging" },
	{ START_OnFunc + TE_OnClipboardText, L"OnClipboardText" },
	{ START_OnFunc + TE_OnCommand, L"OnCommand" },
	{ START_OnFunc + TE_OnInvokeCommand, L"OnInvokeCommand" },
	{ START_OnFunc + TE_OnArrange, L"OnArrange" },
	{ START_OnFunc + TE_OnHitTest, L"OnHitTest" },
	{ START_OnFunc + TE_OnVisibleChanged, L"OnVisibleChanged" },
};

method methodAPI[] = {
	{ 1040, L"Memory" },
	{ 1061, L"ContextMenu" },
	{ 1070, L"DropTarget" },
	{ 1080, L"OleGetClipboard" },
	{ 1082, L"OleSetClipboard" },
	{ 1122, L"sizeof" },
	{ 1132, L"LowPart" },
	{ 1142, L"HighPart" },
	{ 1154, L"QuadPart" },
	{ 1163, L"pvData" },
	{ 1170, L"ExecMethod" },
	{ 2063, L"FindWindow" },
	{ 2073, L"FindWindowEx" },
	{ 2111, L"PathMatchSpec" },
	{ 2230, L"CommandLineToArgv" },
	{ 2431, L"ILIsEqual" },
	{ 2440, L"ILClone" },
	{ 2441, L"ILIsParent" },
	{ 2450, L"ILRemoveLastID" },
	{ 2460, L"ILFindLastID" },
	{ 2461, L"ILIsEmpty" },
	{ 2763, L"FindFirstFile" },
	{ 2773, L"WindowFromPoint" },
//	{ 2782, L"GetDoubleClickTime" },//Deprecated
	{ 2792, L"GetThreadCount" },
	{ 2801, L"DoEvents" },
	{ 2804, L"sscanf" },
	{ 2811, L"SetFileTime" },
	//(string) etc
	{ 6000, L"ILCreateFromPath" },
	{ 6010, L"GetProcObject" },
	//(string) bool
	{ 6001, L"SetCurrentDirectory" },//use wsh.CurrentDirectory
	{ 6011, L"SetDllDirectory" },
	{ 6021, L"PathIsNetworkPath" },
	//(string)int
	{ 6002, L"RegisterWindowMessage"},
	{ 6012, L"StrCmpI"},
	{ 6022, L"StrLen"},
	{ 6032, L"StrCmpNI"},
	{ 6042, L"StrCmpLogical"},
	//(string)string
	{ 6005, L"PathQuoteSpaces"},
	{ 6015, L"PathUnquoteSpaces"},
	{ 6025, L"GetShortPathName"},
	{ 6035, L"PathCreateFromUrl"},
	//(mem) bool
	{ 7001, L"GetCursorPos" },
	{ 7011, L"GetKeyboardState" },
	{ 7021, L"SetKeyboardState" },
	{ 7031, L"GetVersionEx" },
	{ 7041, L"ChooseFont" },
	{ 7051, L"ChooseColor" },
	{ 7061, L"TranslateMessage" },
	{ 7071, L"ShellExecuteEx" },
	//(mem)int
	//(mem) handle
	{ 9003, L"CreateFontIndirect" },
	{ 9013, L"CreateIconIndirect" },
	{ 9023, L"FindText" },
	{ 9033, L"FindReplace" },
	{ 9043, L"DispatchMessage" },
	//No return value
	{ 10000, L"PostMessage" },
	{ 10010, L"SHFreeNameMappings" },
	{ 10020, L"CoTaskMemFree" },
	{ 10030, L"Sleep" },
	{ 10040, L"ShRunDialog" },
	{ 10050, L"DragAcceptFiles" },
	{ 10060, L"DragFinish" },
	{ 10070, L"mouse_event" },
	//BOOL
	{ 20001, L"DrawIcon" },
	{ 20011, L"DestroyMenu" },
	{ 20021, L"FindClose" },
	{ 20031, L"FreeLibrary" },
	{ 20041, L"ImageList_Destroy" },
	{ 20051, L"DeleteObject" },
	{ 20061, L"DestroyIcon" },
	{ 20071, L"IsWindow" },
	{ 20081, L"IsWindowVisible" },
	{ 20091, L"IsZoomed" },
	{ 20101, L"IsIconic" },
	{ 20111, L"OpenIcon" },
	{ 20121, L"SetForegroundWindow" },
	{ 20131, L"BringWindowToTop" },
	{ 20141, L"DeleteDC" },
	{ 20151, L"CloseHandle" },
	{ 20161, L"IsMenu" },
	{ 20171, L"MoveWindow" },
	{ 20181, L"SetMenuItemBitmaps" },
	{ 20191, L"ShowWindow" },
	{ 20201, L"DeleteMenu" },
	{ 20211, L"RemoveMenu" },
	{ 20231, L"DrawIconEx" },
	{ 20241, L"EnableMenuItem" },
	{ 20251, L"ImageList_Remove" },
	{ 20261, L"ImageList_SetIconSize" },
	{ 20271, L"ImageList_Draw" },
	{ 20281, L"ImageList_DrawEx" },
	{ 20291, L"ImageList_SetImageCount" },
	{ 20301, L"ImageList_SetOverlayImage" },
	{ 20321, L"ImageList_Copy" },
	{ 20331, L"DestroyWindow" },
	{ 20341, L"LineTo" },
	{ 20351, L"ReleaseCapture" },
	{ 20361, L"SetCursorPos" },
	{ 20371, L"DestroyCursor" },
	{ 20381, L"SHFreeShared" },
	//+other
	{ 25001, L"InsertMenu" },
	{ 25011, L"SetWindowText" },
	{ 25021, L"RedrawWindow" },
	{ 25031, L"MoveToEx" },
	{ 25041, L"InvalidateRect" },
	{ 25051, L"SendNotifyMessage"},
	//+mem(0)
	{ 29001, L"PeekMessage" },
	{ 29002, L"SHFileOperation" },
	{ 29011, L"GetMessage" },
	//+mem(1)
	{ 29101, L"GetWindowRect" },
	{ 29102, L"GetWindowThreadProcessId" },
	{ 29111, L"GetClientRect" },
	{ 29112, L"SendInput" },
	{ 29121, L"ScreenToClient" },
	{ 29122, L"MsgWaitForMultipleObjectsEx" },
	{ 29131, L"ClientToScreen" },
	{ 29141, L"GetIconInfo" },
	{ 29151, L"FindNextFile" },
	{ 29161, L"FillRect" },
	{ 29171, L"Shell_NotifyIcon" },
	{ 29191, L"EndPaint" },
	//+mem(2)
	{ 29201, L"SystemParametersInfo" },
	{ 29211, L"ImageList_GetIconSize" },
	{ 29221, L"GetTextExtentPoint32" },
	{ 29232, L"SHGetDataFromIDList" },
	//+mem(3)
	{ 29301, L"InsertMenuItem" },
	{ 29311, L"GetMenuItemInfo" },
	{ 29321, L"SetMenuItemInfo" },
	//int
	{ 30002, L"ImageList_SetBkColor" },
	{ 30012, L"ImageList_AddIcon" },
	{ 30022, L"ImageList_Add" },
	{ 30032, L"ImageList_AddMasked" },
	{ 30042, L"GetKeyState" },
	{ 30052, L"GetSystemMetrics" },
	{ 30062, L"GetSysColor" },
	{ 30082, L"GetMenuItemCount" },
	{ 30092, L"ImageList_GetImageCount" },
	{ 30102, L"ReleaseDC" },
	{ 30112, L"GetMenuItemID" },
	{ 30122, L"ImageList_Replace" },
	{ 30132, L"ImageList_ReplaceIcon" },
	{ 30142, L"SetBkMode" },
	{ 30152, L"SetBkColor" },
	{ 30162, L"SetTextColor" },
	{ 30172, L"MapVirtualKey" },
	{ 30182, L"WaitForInputIdle" },
	{ 30192, L"WaitForSingleObject" },
	{ 30202, L"GetMenuDefaultItem" },
	{ 30212, L"CRC32" },
	//+other
	{ 35002, L"TrackPopupMenuEx" },
	{ 35012, L"ExtractIconEx" },
	{ 35022, L"GetObject" },
	{ 35032, L"MultiByteToWideChar" },
	{ 35042, L"WideCharToMultiByte" },
	{ 35052, L"GetAttributesOf" },
	{ 35062, L"DoDragDrop" },
	{ 35072, L"CompareIDs" },
	{ 35082, L"ExecScript" },
	{ 35080, L"GetScriptDispatch" },
	{ 35090, L"GetDispatch" },
	//Handle
	{ 40003, L"ImageList_GetIcon" },
	{ 40013, L"ImageList_Create" },
	{ 40023, L"GetWindowLongPtr" },
	{ 40033, L"GetClassLongPtr" },
	{ 40043, L"GetSubMenu" },
	{ 40053, L"SelectObject" },
	{ 40063, L"GetStockObject" },
	{ 40073, L"GetSysColorBrush" },
	{ 40083, L"SetFocus" },
	{ 40093, L"GetDC" },
	{ 40103, L"CreateCompatibleDC" },
	{ 40113, L"CreatePopupMenu" },
	{ 40123, L"CreateMenu" },
	{ 40133, L"CreateCompatibleBitmap" },
	{ 40143, L"SetWindowLongPtr" },
	{ 40153, L"SetClassLongPtr" },
	{ 40163, L"ImageList_Duplicate" },
	{ 40173, L"SendMessage" },
	{ 40183, L"GetSystemMenu" },
	{ 40193, L"GetWindowDC" },
	{ 40203, L"CreatePen" },
	{ 40213, L"SetCapture" },
	{ 40223, L"SetCursor" },
	{ 40233, L"CallWindowProc" },
	{ 40243, L"GetWindow" },
	{ 40253, L"GetTopWindow" },
	{ 40263, L"OpenProcess" },
	//+other
	{ 45003, L"LoadMenu" },
	{ 45013, L"LoadIcon" },
	{ 45023, L"LoadLibraryEx" },
	{ 45033, L"LoadImage" },
	{ 45043, L"ImageList_LoadImage" },
	{ 45053, L"SHGetFileInfo" },
	{ 45083, L"CreateWindowEx" },
	{ 45093, L"ShellExecute" },
	{ 45103, L"BeginPaint" },
	{ 45113, L"LoadCursor" },
	{ 45123, L"LoadCursorFromFile" },
	{ 45133, L"SHParseDisplayName" },
	//string
	{ 50005, L"GetWindowText" },
	{ 50015, L"GetClassName" },
	{ 50025, L"GetModuleFileName" },
	{ 50035, L"GetCommandLine" },
	{ 50045, L"GetCurrentDirectory" },//use wsh.CurrentDirectory
	{ 50055, L"GetMenuString" },
	{ 50065, L"GetDisplayNameOf" },
	{ 50075, L"GetKeyNameText"},
	{ 50085, L"DragQueryFile"},
	{ 50095, L"SysAllocString"},
	{ 50105, L"SysAllocStringLen"},
	{ 50115, L"SysAllocStringByteLen"},
	{ 50125, L"sprintf"},
	{ 50135, L"base64_encode"},
	{ 50145, L"LoadString"},
};

method methodSB[] = {
	{ 0x10000001, L"Data" },
	{ 0x10000002, L"hwnd" },
	{ 0x10000003, L"Type" },
	{ 0x10000004, L"Navigate" },
	{ 0x10000007, L"Navigate2" },
	{ 0x10000008, L"Index" },
	{ 0x10000009, L"FolderFlags" },
	{ 0x1000000B, L"History" },
	{ 0x10000010, L"CurrentViewMode" },
	{ 0x10000011, L"IconSize" },
	{ 0x10000012, L"Options" },
	{ 0x10000016, L"ViewFlags" },
	{ 0x10000017, L"Id" },
	{ 0x10000018, L"FilterView" },
	{ 0x10000020, L"FolderItem" },
	{ 0x10000024, L"Parent" },
	{ 0x10000031, L"Close" },
	{ 0x10000032, L"Title" },
	{ 0x10000050, L"ShellFolderView" },
	{ 0x10000103, L"hwndView" },
	{ 0x10000104, L"SortColumn" },
	{ 0x10000106, L"Focus" },
	{ 0x10000107, L"HitTest" },
	{ 0x10000108, L"Droptarget" },
	{ 0x10000109, L"Columns"},
	{ 0x10000202, L"Items" },
	{ 0x10000203, L"SelectedItems" },
	{ 0x10000206, L"Refresh" },
	{ 0x10000207, L"ViewMenu" },
	{ 0x10000208, L"TranslateAccelerator" },
	{ 0x10000209, L"GetItemPosition" },
	{ 0x1000020A, L"SelectAndPositionItem" },
	{ 0x10000210, L"TreeView" },
	{ 0x10000280, L"SelectItem" },
	{ 0x10000281, L"FocusedItem" },
	{ 0x10000282, L"GetFocusedItem" },
	{ START_OnFunc + SB_OnIncludeObject, L"OnIncludeObject" },
};

method methodWB[] = {
	{ 0x10000001, L"Data" },
	{ 0x10000002, L"hwnd" },
	{ 0x10000003, L"Type" },
	{ 0x10000004, L"TranslateAccelerator" },
	{ 0x10000005, L"Application" },
	{ 0x10000006, L"Document" },
};

method methodTC[] = {
	{ 1, L"Data" },
	{ 2, L"hwnd" },
	{ 3, L"Type" },
	{ 6, L"HitTest" },
	{ 7, L"Move" },
	{ 8, L"Selected" },
	{ 9, L"Close" },
	{ 10, L"SelectedIndex" },
	{ 11, L"Visible" },
	{ 12, L"Id" },
	{ DISPID_NEWENUM, L"_NewEnum" },
	{ DISPID_TE_ITEM, L"Item" },
	{ DISPID_TE_COUNT, L"Count" },
	{ DISPID_TE_COUNT, L"length" },
	{ TE_OFFSET + TE_Left, L"Left" },
	{ TE_OFFSET + TE_Top, L"Top" },
	{ TE_OFFSET + TE_Width, L"Width" },
	{ TE_OFFSET + TE_Height, L"Height" },
	{ TE_OFFSET + TE_Flags, L"Style" },
	{ TE_OFFSET + TE_Align, L"Align" },
	{ TE_OFFSET + TC_TabWidth, L"TabWidth" },
	{ TE_OFFSET + TC_TabHeight, L"TabHeight" },
};

method methodFIs[] = {
	{ 2, L"Application" },
	{ 3, L"Parent" },
	{ 8, L"AddItem" },
	{ 9, L"hDrop" },
	{ DISPID_NEWENUM, L"_NewEnum" },
	{ DISPID_TE_ITEM, L"Item" },
	{ DISPID_TE_COUNT, L"Count" },
	{ DISPID_TE_COUNT, L"length" },
	{ DISPID_TE_INDEX, L"Index" },
	{ 0x10000001, L"dwEffect" },
	{ 0x10000002, L"pdwEffect" },
};

method methodDT[] = {
	{ 1, L"DragEnter" },
	{ 2, L"DragOver" },
	{ 3, L"Drop" },
	{ 4, L"DragLeave" },
	{ 5, L"Type" },
	{ 6, L"FolderItem" },
};

method methodTV[] = {
	{ 0x10000001, L"Data" },
	{ 0x10000002, L"Type" },
	{ 0x10000003, L"hwnd" },
	{ 0x10000004, L"Close" },
	{ 0x10000007, L"FolderView" },
	{ 0x10000008, L"Align" },
	{ 0x10000105, L"Focus" },
	{ 0x10000106, L"HitTest" },
	{ TE_OFFSET + SB_TreeWidth, L"Width" },
	{ TE_OFFSET + SB_TreeFlags, L"Style" },
	{ TE_OFFSET + SB_EnumFlags, L"EnumFlags" },
	{ TE_OFFSET + SB_RootStyle, L"RootStyle" },
	{ 0x20000000, L"SelectedItem" },
	{ 0x20000001, L"SelectedItems" },
	{ 0x20000002, L"Root" },
	{ 0x20000003, L"SetRoot" },
	{ 0x20000004, L"Expand" },
	{ 0x20000005, L"Columns" },
	{ 0x20000006, L"CountViewTypes" },
	{ 0x20000007, L"Depth" },
	{ 0x20000008, L"EnumOptions" },
	{ 0x20000009, L"Export" },
	{ 0x2000000a, L"Flags" },
	{ 0x2000000b, L"Import" },
	{ 0x2000000c, L"Mode" },
	{ 0x2000000d, L"ResetSort" },
	{ 0x2000000e, L"SetViewType" },
	{ 0x2000000f, L"Synchronize" },
	{ 0x20000010, L"TVFlags" },
};

method methodFI[] = {
	{ 1, L"Application" },
	{ 2, L"Parent" },
	{ 3, L"Name" },
	{ 4, L"Path" },
	{ 5, L"GetLink" },
	{ 6, L"GetFolder" },
	{ 7, L"IsLink" },
	{ 8, L"IsFolder" },
	{ 9, L"IsFileSystem" },
	{ 10, L"IsBrowsable" },
	{ 11, L"ModifyDate" },
	{ 12, L"Size" },
	{ 13, L"Type" },
	{ 14, L"Verbs" },
	{ 15, L"InvokeVerb" },
	{ 99, L"Alt" },
};

method methodMem[] = {
	{ 1, L"P" },
	{ 4, L"Read" },
	{ 5, L"Write" },
	{ 6, L"Size" },
	{ 7, L"Free" },
	{ 8, L"Clone" },
	{ DISPID_NEWENUM, L"_NewEnum" },
	{ DISPID_TE_ITEM, L"Item" },
	{ DISPID_TE_COUNT, L"Count" },
	{ DISPID_TE_COUNT, L"length" },
};

method methodMem2[] = {
	{ VT_I4  << TE_VT, L"int" },
	{ VT_UI4 << TE_VT, L"DWORD" },
	{ VT_UI1 << TE_VT, L"BYTE" },
	{ VT_UI2 << TE_VT, L"WORD" },
	{ VT_UI2 << TE_VT, L"WCHAR" },
	{ VT_PTR << TE_VT, L"HANDLE" },
	{ VT_PTR << TE_VT, L"LPWSTR" },
};

method methodCM[] = {
	{ 1, L"QueryContextMenu" },
	{ 2, L"InvokeCommand" },
	{ 3, L"Items" },
	{ 4, L"GetCommandString" },
};

method methodCD[] = {
	{ 40, L"ShowOpen" },
	{ 41, L"ShowSave" },
	{ 10, L"FileName" },
	{ 13, L"Filter" },
	{ 20, L"InitDir" },
	{ 21, L"DefExt" },
	{ 22, L"Title" },
	{ 30, L"MaxFileSize" },
	{ 31, L"Flags" },
	{ 32, L"FilterIndex" },
	{ 31, L"FlagsEx" },
};

method methodGB[] = {
	{ 1, L"FromHBITMAP" },
	{ 2, L"FromHICON" },
	{ 3, L"FromResource" },
	{ 4, L"FromFile" },
	{ 99, L"Free" },

	{ 100, L"Save" },
	{ 101, L"Base64" },
	{ 102, L"DataURI" },
	{ 110, L"GetWidth" },
	{ 111, L"GetHeight" },
	{ 120, L"GetThumbnailImage" },
	{ 130, L"RotateFlip" },

	{ 210, L"GetHBITMAP" },
	{ 211, L"GetHICON" },
};

// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
VOID ArrangeWindow();

BOOL GetIDListFromObject(IUnknown *punk, LPITEMIDLIST *ppidl);

//Unit

int CalcCrc32(BYTE *pc, int nLen, UINT c)
{
	c ^= MAXUINT;
	for (int i = 0; i < nLen; i++) {
		c = g_pCrcTable[(c ^ pc[i]) & 0xff] ^ (c >> 8);
	}
	return c ^ MAXUINT;
}

int teGetModuleFileName(HMODULE hModule, LPTSTR *ppszPath)
{
	int i = 0;
	for (int nSize = MAX_PATH; nSize < MAX_PATHEX; nSize += MAX_PATH) {
		*ppszPath = new WCHAR[nSize];
		i = GetModuleFileName(hModule, *ppszPath, nSize);
		if (i + 1 < nSize) {
			(*ppszPath)[i] = NULL;
			break;
		}
		delete [] *ppszPath;
	}
	return i;
}

//i >= 0
VOID teDragQueryFile(HDROP hDrop, UINT iFile, LPTSTR *ppszPath)
{
	for (UINT nSize = MAX_PATH; nSize < MAX_PATHEX; nSize += MAX_PATH) {
		*ppszPath = new WCHAR[nSize];
		UINT i = DragQueryFile(hDrop, iFile, *ppszPath, nSize);
		if (i + 1 < nSize) {
			(*ppszPath)[i] = NULL;
			break;
		}
		delete [] *ppszPath;
	}
}

VOID tePathAppend(LPWSTR *ppszPath, LPWSTR pszPath, LPWSTR pszFile)
{
	*ppszPath = new WCHAR[lstrlen(pszPath) + lstrlen(pszFile) + 2];
	lstrcpy(*ppszPath, pszPath);
	PathAppend(*ppszPath, pszFile);
}

BOOL teSetObject(VARIANT *pv, PVOID pObj)
{
	if (pObj) {
		IUnknown *punk = static_cast<IUnknown *>(pObj);
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->pdispVal))) {
			pv->vt = VT_DISPATCH;
			return true;
		}
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->punkVal))) {
			pv->vt = VT_UNKNOWN;
			return true;
		}
	}
	return false;
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
	if (*ppDragItems) {
		(*ppDragItems)->Release();
	}
	*ppDragItems = new CteFolderItems(pDataObj, NULL, false);
}

BOOL teIsFileSystem(LPOLESTR bs)
{
	return lstrlen(bs) >= 3 && ((bs[0] == '\\' && bs[1] == '\\') || (bs[1] == ':' && bs[2] == '\\'));
}

VARIANTARG* GetNewVARIANT(int n)
{
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

BOOL GetDispatch(VARIANT *pv, IDispatch **ppdisp)
{
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		return SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(ppdisp)));
	}
	return false;
}

VOID ClearEvents()
{
	for (int j = Count_OnFunc - 1; j >= 0; j--) {
		if (g_pOnFunc[j]) {
			g_pOnFunc[j]->Release();
			g_pOnFunc[j] = NULL;
		}
	}

	int i = MAX_FV;
	while (--i >= 0) {
		if (g_pSB[i]) {
			for (int j = Count_SBFunc - 1; j >= 0; j--) {
				if (g_pSB[i]->m_pOnFunc[j]) {
					g_pSB[i]->m_pOnFunc[j]->Release();
					g_pSB[i]->m_pOnFunc[j] = NULL;
				}
			}
		}
		else {
			break;
		}
	}
	g_pTE->m_param[TE_Tab] = TRUE;
}

VOID teRegisterDragDrop(HWND hwnd, IDropTarget *pDropTarget, IDropTarget **ppDropTarget)
{
	*ppDropTarget = static_cast<IDropTarget *>(GetProp(hwnd, L"OleDropTargetInterface"));
	SetProp(hwnd, L"OleDropTargetInterface", (HANDLE)pDropTarget);
}

VOID teSetParent(HWND hwnd, HWND hwndParent)
{
	if (GetParent(hwnd) != hwndParent) {
		SetParent(hwnd, hwndParent);
	}
}

BSTR SysAllocStringLenEx(const OLECHAR *strIn, UINT uSize, UINT uOrg)
{
	BSTR bs = SysAllocStringLen(NULL, uSize);
	lstrcpyn(bs, strIn, uSize < uOrg ? uSize : uOrg);
	return bs;
}

HRESULT GetDisplayNameFromPidl(BSTR *pbs, LPITEMIDLIST pidl, SHGDNF uFlags)
{
	HRESULT hr = E_FAIL;
	IShellFolder *pSF;
	LPCITEMIDLIST ItemID;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &ItemID)) {
		STRRET strret;
		if SUCCEEDED(pSF->GetDisplayNameOf(ItemID, uFlags, &strret)) {
			if SUCCEEDED(StrRetToBSTR(&strret, ItemID, pbs)) {
				hr = S_OK;
			}
		}
		pSF->Release();
	}
	return hr;
}

void GetVarPathFromPidl(VARIANT *pVarResult, LPITEMIDLIST pidl, int uFlags)
{
	LPITEMIDLIST pidl2;

	int i;
	BOOL bSpecial;
	bSpecial = false;

	if (uFlags & SHGDN_FORPARSINGEX) {
		for (i = 0; i < MAX_CSIDL; i ++) {
			if (pidl2 = g_pidls[i]) {
				if (::ILIsEqual(pidl, pidl2)) {
					pidl2 = NULL;
					BSTR bsPath;
					if SUCCEEDED(GetDisplayNameFromPidl(&bsPath, pidl, SHGDN_FORPARSING)) {
						pidl2 = SHSimpleIDListFromPath(bsPath);
						::SysFreeString(bsPath);
					}
					if (!::ILIsEqual(pidl, pidl2)) {
						pVarResult->lVal = i;
						pVarResult->vt = VT_I4;
					}
					if (pidl2) {
						::CoTaskMemFree(pidl2);
					}
					break;
				}
			}
		}
	}
	if (pVarResult->vt != VT_I4) {
		if SUCCEEDED(GetDisplayNameFromPidl(&pVarResult->bstrVal, pidl, uFlags & (~SHGDN_FORPARSINGEX))) {
			pVarResult->vt = VT_BSTR;
		}
	}
}

BOOL GetMemFormVariant(VARIANT *pv, CteMemory **ppMem)
{
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		return SUCCEEDED(punk->QueryInterface(g_ClsIdStruct, (LPVOID *)ppMem));
	}
	return FALSE;
}

LONGLONG GetLLFromVariant(VARIANT *pv)
{
	if (pv) {
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return GetLLFromVariant(pv->pvarVal);
		}
		if (pv->vt == VT_I4) {
			return pv->lVal;
		}
		if (pv->vt == VT_R8) {
			return (LONGLONG)pv->dblVal;
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
		if (pv->vt == VT_I8) {
			return pv->llVal;
		}
		if (pv->vt == VT_UI4) {
			return pv->ulVal;
		}
		if (pv->vt == VT_UI8) {
			return pv->ullVal;
		}
		CteMemory *pMem;
		if (GetMemFormVariant(pv, &pMem)) {
			LONGLONG ll = (LONGLONG)pMem->m_pc;
			pMem->Release();
			return ll;
		}
		VARIANT vo;
		VariantInit(&vo);
		if (SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8))) {
			return vo.llVal;
		}
		if (SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_UI8))) {
			return vo.ullVal;
		}
		if (SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I4))) {
			return vo.lVal;
		}
		if (SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_UI4))) {
			return vo.ulVal;
		}
	}
	return 0;
}

VOID teVariantChangeType(__out VARIANTARG * pvargDest,
				__in const VARIANTARG * pvarSrc, __in VARTYPE vt)
{
	VariantInit(pvargDest);
	if (pvarSrc->vt == (VT_ARRAY | VT_I4)) {
		VARIANT v;
		v.llVal = GetLLFromVariant((VARIANT *)pvarSrc);
		v.vt = VT_I8;
		if FAILED(VariantChangeType(pvargDest, &v, 0, vt)) {
			pvargDest->llVal = 0;
		}
	}
	else if FAILED(VariantChangeType(pvargDest, pvarSrc, 0, vt)) {
		pvargDest->llVal = 0;
	}
}

LPWSTR teGetCommandLine()
{
	if (!g_bsCmdLine) {
		LPWSTR strCmdLine = GetCommandLine();
		int nSize = lstrlen(strCmdLine) + MAX_PATH;
		g_bsCmdLine = SysAllocStringLen(NULL, nSize);
		int j = 0;
		int i = 0;
		while (i < nSize) {
			if (StrCmpNI(&strCmdLine[j], L"/idlist,:", 9) == 0) {
				LONGLONG hData;
				DWORD dwProcessID;
				WCHAR sz[MAX_PATH + 2];
				swscanf_s(&strCmdLine[j + 9], L"%32[1234567890]s", sz, _countof(sz));
				swscanf_s(sz, L"%lld", &hData);
				int k = lstrlen(sz);
				swscanf_s(&strCmdLine[j + k + 10], L"%32[1234567890]s", sz, _countof(sz));
				swscanf_s(sz, L"%d", &dwProcessID);
				j += k + lstrlen(sz) + 11;
#ifdef _2000XP
				LPITEMIDLIST pidl = (LPITEMIDLIST)SHLockShared((HANDLE)hData, dwProcessID);
				if (pidl) {
					VARIANT v, v2;
					GetVarPathFromPidl(&v, pidl, SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
					SHUnlockShared(pidl);
					SHFreeShared((HANDLE)hData, dwProcessID);
					teVariantChangeType(&v2, &v, VT_BSTR);
					lstrcpy(sz, v2.bstrVal);
					PathQuoteSpaces(sz);
					lstrcpy(&g_bsCmdLine[i], sz);
					i += lstrlen(sz);
					VariantClear(&v);
					VariantClear(&v2);
					while (strCmdLine[j] && strCmdLine[j] != _T(',')) {
						j++;
					}
				}
#endif
			}
			else {
				g_bsCmdLine[i++] = strCmdLine[j++];
			}
		}
	}
	return g_bsCmdLine;
}

HRESULT teCLSIDFromProgID(__in LPCOLESTR lpszProgID, __out LPCLSID lpclsid)
{
	if (StrCmpNI(lpszProgID, L"new:", 4) == 0) {
		return CLSIDFromString(&lpszProgID[4], lpclsid);
	}
	return CLSIDFromProgID(lpszProgID, lpclsid);
}

HWND FindTreeWindow(HWND hwnd)
{
	HWND hwnd1 = FindWindowEx(hwnd, 0, WC_TREEVIEW, NULL);
	return hwnd1 ? hwnd1 : FindWindowEx(FindWindowEx(hwnd, 0, L"NamespaceTreeControl", NULL), 0, WC_TREEVIEW, NULL);
}

static void threadFileOperation(void *args)
{
	::InterlockedIncrement(&g_nThreads);
	LPSHFILEOPSTRUCT pFO = (LPSHFILEOPSTRUCT)args; 
	try {
		::SHFileOperation(pFO);
	}
	catch (...) {
	}
	::SysFreeString((BSTR)pFO->pTo);
	::SysFreeString((BSTR)pFO->pFrom);
	delete [] pFO;
	::InterlockedDecrement(&g_nThreads);
	::_endthread();
}

BOOL teSetRect(HWND hwnd, int left, int top, int right, int bottom)
{
	RECT rc;
	GetWindowRect(hwnd, &rc);
	MoveWindow(hwnd, left, top, right - left, bottom - top, true);
	return (rc.right - rc.left != right - left || rc.bottom - rc.top != bottom - top);
}

BOOL GetShellFolder(IShellFolder **ppSF, LPCITEMIDLIST pidl)
{
	if (ILIsEmpty(pidl)) {
		SHGetDesktopFolder(ppSF);
		return true;
	}
	IShellFolder *pSF;
	LPCITEMIDLIST ItemID;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &ItemID)) {
		pSF->BindToObject(ItemID, NULL, IID_PPV_ARGS(ppSF));
		pSF->Release();
	}
//	g_pDesktopFolder->BindToObject(pidl, NULL, IID_PPV_ARGS(ppSF));
	if (*ppSF == NULL) {
		SHGetDesktopFolder(ppSF);	
	}
	return true;
}

LPITEMIDLIST teILCreateFromPath2(IShellFolder *pSF, LPWSTR pszPath, SHGDNF uFlags)
{
	LPITEMIDLIST pidl = NULL;
	IEnumIDList *peidl = NULL;
	BSTR bstr = NULL;
	int ashgdn[] = {SHGDN_FORPARSING, SHGDN_FORADDRESSBAR};
	int nLen;
	if SUCCEEDED(pSF->EnumObjects(NULL, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN, &peidl)) {
		LPITEMIDLIST pidlPart;
		ULONG celtFetched;
		while (peidl->Next(1, &pidlPart, &celtFetched) == S_OK) {
			STRRET strret;
			for (int j = 0; j < 2; j++) {
				SHGDNF uFlag = uFlags & ashgdn[j];
				if (uFlag && SUCCEEDED(pSF->GetDisplayNameOf(pidlPart, uFlag, &strret))) {
					if SUCCEEDED(StrRetToBSTR(&strret, pidlPart, &bstr)) {
						if (lstrcmpi(const_cast<LPCWSTR>(pszPath), (LPCWSTR)bstr) == 0) {
							::SysFreeString(bstr);
							LPITEMIDLIST pidlParent;
							if (GetIDListFromObject(pSF, &pidlParent)) {
								pidl = ILCombine(pidlParent, pidlPart);
								::CoTaskMemFree(pidlParent);
								::CoTaskMemFree(pidlPart);
								peidl->Release();
								return pidl;
							}
							return NULL;
						}
						nLen = lstrlen(bstr);
						if (nLen && StrCmpNI(const_cast<LPCWSTR>(pszPath), (LPCWSTR)bstr, nLen) == 0) {
							IShellFolder *pSF1;
							if SUCCEEDED(pSF->BindToObject(pidlPart, NULL, IID_PPV_ARGS(&pSF1))) {
								pidl = teILCreateFromPath2(pSF1, pszPath, uFlag);
								pSF1->Release();
								if (pidl) {
									return pidl;
								}
							}
						}
						::SysFreeString(bstr);
					}
				}
			}
			::CoTaskMemFree(pidlPart);
		}
		peidl->Release();
	}
	return NULL;
}

LPITEMIDLIST teILCreateFromPath(LPWSTR pszPath)
{
	LPITEMIDLIST pidl = NULL;
	BSTR bstr = NULL;
	if (pszPath) {
		LPWSTR pszPath2 = NULL;
		if (pszPath[0] == _T('"')) {
			pszPath2 = new WCHAR[lstrlen(pszPath) + 1];
			lstrcpy(pszPath2, pszPath);
			PathUnquoteSpaces(pszPath2);
			pszPath = pszPath2;
		}
		int n = PathGetDriveNumber(pszPath);
		if (n >= 0 && DriveType(n) == DRIVE_NO_ROOT_DIR && lstrlen(pszPath) > 3) {
			WCHAR szDrive[4];
			lstrcpyn(szDrive, pszPath, 4);
			if (lpfnSHParseDisplayName) {
				lpfnSHParseDisplayName(szDrive, NULL, &pidl, 0, NULL);
			}
			else {
				pidl = ILCreateFromPath(szDrive);
			}
			if (!pidl) {
				return NULL;
			}
			IShellFolder *pSF = NULL;
			if (GetShellFolder(&pSF, pidl)) {
				IEnumIDList *peidl = NULL;
				if SUCCEEDED(pSF->EnumObjects(g_hwndMain, SHCONTF_FOLDERS, &peidl)) {
					peidl->Release();
				}
				pSF->Release();
			}
			CoTaskMemFree(pidl);
		}
		if (lpfnSHParseDisplayName) {
			lpfnSHParseDisplayName(pszPath, NULL, &pidl, 0, NULL);
		}
		else {
			pidl = ILCreateFromPath(pszPath);
		}
		if (pidl == NULL) {
			for (int i = 0; i < MAX_CSIDL; i++) {
				if (pidl = g_pidls[i]) {
					if SUCCEEDED(GetDisplayNameFromPidl(&bstr, pidl, SHGDN_FORADDRESSBAR)) {
						n = lstrcmpi(const_cast<LPCWSTR>(pszPath), (LPCWSTR)bstr);
						::SysFreeString(bstr);
						if (n == 0) {
							pidl = ILClone(pidl);
							break;
						}
					}
				}
			}
			if (pidl == NULL) {
				IShellFolder *pDesktopFolder;
				if (SHGetDesktopFolder(&pDesktopFolder) == S_OK) {
					pidl = teILCreateFromPath2(pDesktopFolder, pszPath, SHGDN_FORPARSING | SHGDN_FORADDRESSBAR);
					pDesktopFolder->Release();
				}
			}
		}
		if (pszPath2) {
			delete [] pszPath2;
		}
	}
	return pidl;
}

int ILGetCount(LPITEMIDLIST pidl)
{
	if (ILIsEmpty(pidl)) {
		return 0;
	}
	return ILGetCount(ILGetNext(pidl)) + 1;
}

LPWSTR teGetMenuString(HMENU hMenu, UINT uIDItem, BOOL fByPosition)
{
	MENUITEMINFO mii;
	::ZeroMemory(&mii, sizeof(MENUITEMINFO));
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask  = MIIM_STRING;
	GetMenuItemInfo(hMenu, uIDItem, fByPosition, &mii);
	if (mii.cch) {
		LPWSTR dwTypeData = new WCHAR[++mii.cch + 1];
		mii.dwTypeData = dwTypeData;
		GetMenuItemInfo(hMenu, uIDItem, fByPosition, &mii);
		dwTypeData[mii.cch] = NULL;
		return dwTypeData;
	}
	else {
		return NULL;
	}
}

VOID teMenuText(LPWSTR sz)
{
	int i = 0, j = 0;
	while (sz[i]) {
		if (sz[i] == '&') {
			i++;
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
		i++;
	}
	return -1;
}

VOID ToMinus(BSTR *pbs)
{
	int nLen = SysStringByteLen(*pbs) + sizeof(WCHAR);
	LPWSTR szCol = (LPWSTR)new CHAR[nLen + sizeof(WCHAR)];
	szCol[0] = '-';
	::CopyMemory(&szCol[1], *pbs, nLen);
	::SysReAllocString(pbs, szCol);
	delete [] szCol;
}

int SizeOfvt(VARTYPE vt)
{
	switch (vt & 0xff) {
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
	}
	return 0;
}

char* GetpcFormVariant(VARIANT *pv)
{
	CteMemory *pMem;
	if (GetMemFormVariant(pv, &pMem)) {
		char *pc = pMem->m_pc; 
		pMem->Release();
		return pc;
	}
	return NULL;
}

BOOL CALLBACK EnumChildProc(HWND hwnd, LPARAM lParam)
{
	WCHAR szClass[MAX_CLASS_NAME + 1];
	if (GetClassName(hwnd, szClass, MAX_CLASS_NAME)) {
		if (lstrcmpi(szClass, L"SHELLDLL_DefView") == 0) {
			*(HWND *)lParam = hwnd;
			return FALSE;
		}
	}
	return TRUE;
}

void teCopyMenu(HMENU hDest, HMENU hSrc, UINT fState)
{
	int n = GetMenuItemCount(hSrc);
	while (--n >= 0) {
		MENUITEMINFO mii;
		::ZeroMemory(&mii, sizeof(MENUITEMINFO));
		mii.cbSize = sizeof(MENUITEMINFO);
		mii.fMask  = MIIM_STRING;
		GetMenuItemInfo(hSrc, n, TRUE, &mii);
		LPWSTR dwTypeData = new WCHAR[++mii.cch + 1];
		mii.dwTypeData = dwTypeData;
		mii.fMask  = MIIM_ID | MIIM_TYPE | MIIM_SUBMENU | MIIM_STATE;
		GetMenuItemInfo(hSrc, n, TRUE, &mii);
		HMENU hSubMenu = mii.hSubMenu; 
		if (hSubMenu) {
			mii.hSubMenu = CreatePopupMenu();
		}
		mii.fState = fState;
		InsertMenuItem(hDest, 0, TRUE, &mii);
		if (hSubMenu) {
			teCopyMenu(mii.hSubMenu, hSubMenu, fState);
		}
		delete [] dwTypeData;
	}
}

HRESULT GetDataFromPidl(LPITEMIDLIST pidl, int nFormat, void *pv, int cb)
{
	HRESULT hr = E_FAIL;
	IShellFolder *pSF;
	LPCITEMIDLIST ItemID;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &ItemID)) {
		hr = SHGetDataFromIDList(pSF, ItemID, nFormat, pv, cb);
		pSF->Release();
	}
	return hr;
}

int SBfromhwnd(HWND hwnd)
{
	int i = MAX_FV;
	while (--i >= 0) {
		if (g_pSB[i]) {
			if (g_pSB[i]->m_hwnd == hwnd || IsChild(g_pSB[i]->m_hwnd, hwnd)) {
				return i;
			}
		}
		else {
			return -1;
		}
	}
	return -1;
}

int TCfromhwnd(HWND hwnd)
{
	int i = MAX_TC;
	while (--i >= 0) {
		if (g_pTC[i]) {
			if (g_pTC[i]->m_hwnd == hwnd || g_pTC[i]->m_hwndStatic == hwnd || g_pTC[i]->m_hwndButton == hwnd) {
				return i;
			}
		}
		else {
			return -1;
		}
	}
	return -1;
}

int TVfromhwnd(HWND hwnd, BOOL bAll)
{
	int i = MAX_FV;
	while (--i >= 0) {
		if (g_pSB[i]) {
			HWND hTV = g_pSB[i]->m_pTV->m_hwnd;
			if (hTV == hwnd || IsChild(hTV, hwnd)) {
				if (bAll || g_pSB[i]->m_pTV->m_bMain) {
					return i;
				}
				else {
					return -1;
				}
			}
		}
		else {
			return -1;
		}
	}
	return -1;
}

void CheckChangeTabSB(HWND hwnd)
{
	if (g_pTabs) {
		int i = SBfromhwnd(hwnd);
		if (i >= 0) {
			if (g_pTabs->m_hwnd != g_pSB[i]->m_pTabs->m_hwnd) {
				g_pTabs = g_pSB[i]->m_pTabs;
				g_pSB[i]->m_pTabs->TabChanged(false);
			}
		}
	}
}

void CheckChangeTabTC(HWND hwnd, BOOL bFocusSB)
{
	int i = TCfromhwnd(hwnd);
	if (i >= 0) {
		if (g_pTabs->m_hwnd != g_pTC[i]->m_hwnd) {
			g_pTabs = g_pTC[i];
			g_pTC[i]->TabChanged(false);
			if (bFocusSB) {
				CteShellBrowser *pSB;
				pSB = g_pTC[i]->GetShellBrowser(g_pTC[i]->m_nIndex);
				if (pSB) {
					pSB->m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
				}
			}
		}
	}
}

BOOL AdjustIDList(LPITEMIDLIST *ppidllist, int nCount)
{
	if (!ppidllist || ppidllist[0] || nCount <= 0) {
		return false;
	}
	ppidllist[0] = ::ILClone(ppidllist[1]);
	ILRemoveLastID(ppidllist[0]);
	int nCommon = ILGetCount(ppidllist[0]);
	int nBase = nCommon;

	for (int i = nCount; i > 1 && nCommon; i--) {
		LPITEMIDLIST pidl = ::ILClone(ppidllist[i]);
		ILRemoveLastID(pidl);
		int nLevel = ILGetCount(pidl);
		while (nLevel > nCommon) {
			ILRemoveLastID(pidl);
			nLevel--;
		}
		while (nLevel < nCommon) {
			ILRemoveLastID(ppidllist[0]);
			nCommon--;
		}
		while (!ILIsEqual(pidl, ppidllist[0])) {
			ILRemoveLastID(pidl);
			ILRemoveLastID(ppidllist[0]);
			nCommon--;
		}
		::CoTaskMemFree(pidl);
	}

	if (nCommon) {
		for (int i = nCount; i > 0; i--) {
			LPITEMIDLIST pidl = ppidllist[i];
			for (int j = nCommon; j > 0; j--) {
				pidl = ILGetNext(pidl);
			}
			pidl = ::ILClone(pidl);
			::CoTaskMemFree(ppidllist[i]);
			ppidllist[i] = pidl;
		}
	}
	return nCommon == nBase;
}

LPITEMIDLIST* IDListFormDataObj(IDataObject *pDataObj, long *pnCount)
{
	LPITEMIDLIST *ppidllist = NULL;
	*pnCount = 0;
	if (pDataObj) {
		STGMEDIUM Medium;
		if (pDataObj->GetData(&IDLISTFormat, &Medium) == S_OK) {
			CIDA *pIDA = (CIDA *)GlobalLock(Medium.hGlobal);
			try {
				if (pIDA) {
					*pnCount = pIDA->cidl;
					ppidllist = new LPITEMIDLIST[*pnCount + 1];
					for (int i = *pnCount; i >= 0; i--) {
						char *pc = (char *)pIDA + pIDA->aoffset[i];
						ppidllist[i] = ::ILClone((LPCITEMIDLIST)pc);
					}
				}
			} catch(...) {
				*pnCount = 0;
			}
			GlobalUnlock(Medium.hGlobal);
			ReleaseStgMedium(&Medium);
		}
		else if (pDataObj->GetData(&HDROPFormat, &Medium) == S_OK) {
			*pnCount = DragQueryFile((HDROP)Medium.hGlobal, (UINT)-1, NULL, 0);
			ppidllist = new LPITEMIDLIST[*pnCount + 1];
			ppidllist[0] = NULL;
			for (int i = *pnCount; i-- >= 0;) {
				LPWSTR pszPath;
				teDragQueryFile((HDROP)Medium.hGlobal, i, &pszPath);
				ppidllist[i + 1] = teILCreateFromPath(pszPath);
				delete [] pszPath;
			}
			ReleaseStgMedium(&Medium);
		}
	}
	return ppidllist;
}

VOID GetPidlFromSV(LPITEMIDLIST *ppidl, IShellView *pSV)
{
	IFolderView *pFV;
	if SUCCEEDED(pSV->QueryInterface(IID_PPV_ARGS(&pFV))) {
		IShellFolder *pSF;
		if SUCCEEDED(pFV->GetFolder(IID_PPV_ARGS(&pSF))) {
			GetIDListFromObject(pSF, ppidl);
			pSF->Release();
		}
		pFV->Release();
		return;
	}
#ifdef _W2000
	//Windows 2000
	IDataObject *pDataObj;
	if SUCCEEDED(pSV->GetItemObject(SVGIO_ALLVIEW, IID_PPV_ARGS(&pDataObj))) {
		long nCount;
		LPITEMIDLIST *pidllist = IDListFormDataObj(pDataObj, &nCount);
		*ppidl = pidllist[0];
		while (--nCount >= 1) {
			if (pidllist) {
				::CoTaskMemFree(pidllist[nCount]);
			}
		}
		if (pidllist) {
			delete [] pidllist;
		}
		if (pDataObj) {
			pDataObj->Release();
		}
	}
#endif
}

VOID HandleMenuMessage(MSG *pMsg)
{
	LRESULT lResult = 0;
	if (g_pCCM) {
		if (g_pCCM->m_pContextMenu3) {
			g_pCCM->m_pContextMenu3->HandleMenuMsg2(pMsg->message, pMsg->wParam, pMsg->lParam, &lResult);
		}
		else if (g_pCCM->m_pContextMenu2) {
			g_pCCM->m_pContextMenu2->HandleMenuMsg(pMsg->message, pMsg->wParam, pMsg->lParam);
		}
	}
}

BOOL GetIDListFromObject(IUnknown *punk, LPITEMIDLIST *ppidl)
{
	*ppidl = NULL;
	if (!punk) {
		return FALSE;
	}
	IPersistFolder2 *pPF2;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pPF2))) {
		pPF2->GetCurFolder(ppidl);
		pPF2->Release();
		if (*ppidl != NULL) {
			return true;
		}
	}
	IPersistIDList *pPI;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pPI))) {
		pPI->GetIDList(ppidl);
		pPI->Release();
		if (*ppidl != NULL) {
			return true;
		}
	}
	FolderItem *pFI;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pFI))) {
		BSTR bstr;
		if SUCCEEDED(pFI->get_Path(&bstr)) {
			*ppidl = teILCreateFromPath(bstr);
			::SysFreeString(bstr);
		}
		pFI->Release();
		if (*ppidl != NULL) {
			return true;
		}
	}
	IServiceProvider *pSP;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pSP))) {
		IShellBrowser *pSB;
		if SUCCEEDED(pSP->QueryService(SID_SShellBrowser, IID_PPV_ARGS(&pSB))) {
			IShellView *pSV;
			if SUCCEEDED(pSB->QueryActiveShellView(&pSV)) {
				GetPidlFromSV(ppidl, pSV);
				pSV->Release();
			}
			pSB->Release();
		}
		pSP->Release();
		if (*ppidl != NULL) {
			return true;
		}
	}
	CteTreeView *pTV;
	if SUCCEEDED(punk->QueryInterface(g_ClsIdTV, (LPVOID *)&pTV)) {
		IDispatch * pdisp;
		if SUCCEEDED(pTV->getSelected(&pdisp)) {
			GetIDListFromObject(pdisp, ppidl);
			pdisp->Release();
		}
		pTV->Release();
		if (*ppidl != NULL) {
			return true;
		}
	}
	return false;
}

BOOL GetIDListFromObjectEx(IUnknown *punk, LPITEMIDLIST *ppidl)
{
	if (!GetIDListFromObject(punk, ppidl)) {
		*ppidl = ::ILClone(g_pidls[CSIDL_DESKTOP]);
	}
	return true;
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
		if (pv->vt == VT_UI4) {
			return pv->ulVal;
		}
		if (pv->vt == VT_R8) {
			return (int)(LONGLONG)pv->dblVal;
		}
		VARIANT vo;
		VariantInit(&vo);
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I4)) {
			return vo.lVal;
		}
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_UI4)) {
			return vo.ulVal;
		}
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8)) {
			return (int)vo.llVal;
		}
	}
	return 0;
}

BOOL GetPidlFromVariant(LPITEMIDLIST *ppidl, VARIANT *pv)
{
	*ppidl = NULL;
	switch(pv->vt) {
		case VT_UNKNOWN:
		case VT_DISPATCH:
			GetIDListFromObject(pv->punkVal, ppidl);
			break;
		case VT_UNKNOWN | VT_BYREF:
		case VT_DISPATCH | VT_BYREF:
			GetIDListFromObject(*pv->ppunkVal, ppidl);
			break;
		case VT_ARRAY | VT_I1:
		case VT_ARRAY | VT_UI1:
		case VT_ARRAY | VT_I1 | VT_BYREF:
		case VT_ARRAY | VT_UI1 | VT_BYREF:
			LONG lUBound, lLBound, nSize;
			SAFEARRAY *parray;
			parray = (pv->vt & VT_BYREF) ? pv->pparray[0] : pv->parray;
			SafeArrayGetUBound(parray, 1, &lUBound);
			SafeArrayGetLBound(parray, 1, &lLBound);
			nSize = lUBound - lLBound + 1;
			*ppidl = (LPITEMIDLIST)CoTaskMemAlloc(nSize);
			::CopyMemory(*ppidl, parray, nSize);
			break;
		case VT_NULL:
			break;
		case VT_BSTR:
		case VT_LPWSTR:
			*ppidl = teILCreateFromPath(pv->bstrVal);
			break;
		case VT_BSTR | VT_BYREF:
		case VT_LPWSTR | VT_BYREF:
			*ppidl = teILCreateFromPath(*pv->pbstrVal);
			break;
		case VT_VARIANT | VT_BYREF:
			return GetPidlFromVariant(ppidl, pv->pvarVal);
	}
	if (*ppidl == NULL) {
		VARIANT v;
		VariantInit(&v);
		if (SUCCEEDED(VariantChangeType(&v, pv, 0, VT_I4))) {
			if (v.lVal < MAX_CSIDL) {
				*ppidl = ::ILClone(g_pidls[v.lVal]);
			}
		}
	}
	return (*ppidl != NULL);
}

BOOL GetVarArrayFromPidl(VARIANT *vaPidl, LPITEMIDLIST pidl)
{
	SAFEARRAY *psa;
	ULONG cbData;

	VariantInit(vaPidl);
	cbData = ILGetSize(pidl);
	psa = SafeArrayCreateVector(VT_UI1, 0, cbData);
	if (psa) {
		PVOID pvData;
		if (::SafeArrayAccessData(psa, &pvData) == S_OK) {
			::CopyMemory(pvData, pidl, cbData);
			::SafeArrayUnaccessData(psa);
			vaPidl->vt = VT_ARRAY | VT_UI1;
			vaPidl->parray = psa;
			return true;
		}
	}
	return false;
}

HRESULT GetFolderObjFromPidl(LPITEMIDLIST pidl, Folder** ppsdf)
{
	VARIANT v;
	GetVarArrayFromPidl(&v, pidl);
	HRESULT hr = g_psha->NameSpace(v, ppsdf);
	VariantClear(&v);
	return hr;
}

BOOL GetFolderItemFromPidl(FolderItem **ppid, LPITEMIDLIST pidl)
{
	BOOL Result;
	Result = false;
	*ppid = NULL;

	IPersistFolder *pPF = NULL;
	if FAILED(CoCreateInstance(CLSID_FolderItem, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pPF))) {
		Folder *pF = NULL;
		if (GetFolderObjFromPidl(pidl, &pF) == S_OK) {
			Folder3 *pF3 = NULL;
			if SUCCEEDED(pF->QueryInterface(IID_PPV_ARGS(&pF3))) {
				pF3->get_Self(ppid);
				pF3->Release();
			}
			pF->Release();
			if (*ppid) {
				return S_OK;
			}
		}
		pPF = new CteFolderItem(NULL);
	}
	if SUCCEEDED(pPF->Initialize(pidl)) {
		Result = SUCCEEDED(pPF->QueryInterface(IID_PPV_ARGS(ppid)));
	}
	pPF->Release();
	return Result;
}

BOOL GetFolderItemFromFolderItems(FolderItem **ppFolderItem, FolderItems *pFolderItems, int nIndex)
{
	VARIANT v;
	v.vt = VT_I4;
	v.lVal = nIndex;
	if (pFolderItems->Item(v, ppFolderItem) == S_OK) {
		return true;
	}
	GetFolderItemFromPidl(ppFolderItem, g_pidls[CSIDL_DESKTOP]);
	return true;
}

BOOL GetDataObjFromVariant(IDataObject **ppDataObj, VARIANT *pv)
{
	*ppDataObj = NULL;
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(ppDataObj))) {
			return true;
		}
		FolderItems *pItems;
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pItems))) {
			long lCount;
			pItems->get_Count(&lCount);
			VARIANT v;
			VariantInit(&v);
			v.vt = VT_I4;
			IShellFolder *pSF = NULL;
			LPCITEMIDLIST *pidlList = new LPCITEMIDLIST[lCount];
			for (v.lVal = 0; v.lVal < lCount; v.lVal++) {
				FolderItem *pItem;
				if SUCCEEDED(pItems->Item(v, &pItem)) {
					LPITEMIDLIST pidl = NULL;
					HRESULT hr = E_FAIL;

					IPersistFolder2 *pPF2;
					hr = pItem->QueryInterface(IID_PPV_ARGS(&pPF2));
					if SUCCEEDED(hr) {
						hr = pPF2->GetCurFolder(&pidl);
						pPF2->Release();
					}
					if (hr != S_OK) {
						IPersistIDList *pPI;
						hr = pItem->QueryInterface(IID_PPV_ARGS(&pPI));
						if SUCCEEDED(hr) {
							hr = pPI->GetIDList(&pidl);
							pPI->Release();
						}
					}
					if (hr != S_OK) {
						BSTR bstr;
						hr = pItem->get_Path(&bstr);
						if SUCCEEDED(hr) {
							pidl = teILCreateFromPath(bstr);
							::SysFreeString(bstr);
						}
					}
					if (pSF) {
						pidlList[v.lVal] = ::ILClone(ILFindLastID(pidl));
					}
					else {
						SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlList[v.lVal]);
						pidlList[v.lVal] = ::ILClone(pidlList[v.lVal]);
					}
					::CoTaskMemFree(pidl);
				}
			}
			if (pSF) {
				pSF->GetUIObjectOf(g_hwndMain, lCount, pidlList, IID_IDataObject, NULL, (LPVOID*)ppDataObj);
				pSF->Release();
			}
			while (--lCount >= 0) {
				::CoTaskMemFree((LPVOID)pidlList[lCount]);
			}
			pItems->Release();
			if (*ppDataObj) {
				return true;
			}
		}
	}
	LPITEMIDLIST pidl;
	pidl = NULL;
	if (GetPidlFromVariant(&pidl, pv)) {
		IShellFolder *pSF;
		LPCITEMIDLIST ItemID;
		if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &ItemID)) {
			pSF->GetUIObjectOf(g_hwndMain, 1, &ItemID, IID_IDataObject, NULL, (LPVOID*)ppDataObj);
			pSF->Release();
		}
		::CoTaskMemFree(pidl);
	}
	return *ppDataObj != NULL;
}

HRESULT GetFolderItemFromShellItem(FolderItem **ppid, IShellItem *psi)
{
	HRESULT hr = E_FAIL;
	if (psi) {
		IShellFolder *pSF;
		if SUCCEEDED(psi->BindToHandler(NULL, BHID_SFObject, IID_PPV_ARGS(&pSF))) {
			LPITEMIDLIST pidl;
			if (GetIDListFromObject(pSF, &pidl)) {
				hr = GetFolderItemFromPidl(ppid, pidl);
				::CoTaskMemFree(pidl);
			}
			pSF->Release();
		}
	}
	return hr;
}

BOOL GetVariantFromShellItem(VARIANT *pv, IShellItem *psi)
{
	FolderItem *pid;
	if SUCCEEDED(GetFolderItemFromShellItem(&pid, psi)) {
		return teSetObject(pv, pid);
	}
	return false;
}

VOID SetVariantLL(VARIANT *pv, LONGLONG ll)
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
		SAFEARRAY *psa;
		psa = SafeArrayCreateVector(VT_I4, 0, sizeof(LONGLONG) / sizeof(int));
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

int GetIntFromVariantClear(VARIANT *pv)
{
	int i = GetIntFromVariant(pv);
	VariantClear(pv);
	return i;
}

VOID SetReturnValue(VARIANT *pVarResult, DISPID dispIdMember, LONGLONG *pResult)
{
	if (pVarResult) {
		switch (dispIdMember % 10) {
			case 1:
				pVarResult->boolVal = (*(BOOL *)pResult) ? VARIANT_TRUE : VARIANT_FALSE;
				pVarResult->vt = VT_BOOL;
				break;
			case 2:
				pVarResult->lVal = *(LONG *)pResult;
				pVarResult->vt = VT_I4;
				break;
			case 3:
				SetVariantLL(pVarResult, (LONGLONG)*(HANDLE *)pResult);
				break;
			case 4:
				SetVariantLL(pVarResult, *pResult);
				break;
		}
	}
}

int* SortMethod(method *method, int nSize)
{
	nSize /= sizeof(method[0]);
	int *pi = new int[nSize];

	for (int j = 0; j < nSize; j++) {
		BSTR bs = method[j].name;
		int nMin = 0;
		int nMax = j - 1;
		int nIndex;
		while (nMin <= nMax) {
			nIndex = (nMin + nMax) / 2;
			if (lstrcmpi(bs, method[pi[nIndex]].name) < 0) {
				nMax = nIndex - 1;
			}
			else {
				nMin = nIndex + 1;
			}
		}
		for (int i = j; i > nMin; i--) {
			pi[i] = pi[i - 1];
		}
		pi[nMin] = j;
	}
	return pi;
}

int teBSearch(method *method, int nSize, int* pMap, LPOLESTR bs)
{
	int nMin = 0;
	int nMax = nSize - 1;
	int nIndex, nCC;

	while (nMin <= nMax) {
		nIndex = (nMin + nMax) / 2;
		nCC = lstrcmpi(bs, method[pMap[nIndex]].name);
		if (nCC < 0) {
			nMax = nIndex - 1;
		}
		else if (nCC > 0) {
			nMin = nIndex + 1;
		}
		else {
			return nIndex;
		}
	}
	return -1;
}

HRESULT teGetDispId(method *method, int nSize, int* pMap, LPOLESTR bs, DISPID *rgDispId, BOOL bNum)
{
	int nIndex = teBSearch(method, nSize / sizeof(method[0]), pMap, bs);
	if (nIndex >= 0) {
		*rgDispId = method[pMap[nIndex]].id;
		return S_OK;
	}
	
	if (bNum) {
		VARIANT v, vo;
		v.bstrVal = bs;
		v.vt = VT_BSTR;
		VariantInit(&vo);
		if (SUCCEEDED(VariantChangeType(&vo, &v, 0, VT_I4))) {
			*rgDispId = vo.lVal + DISPID_COLLECTION_MIN;
			return *rgDispId < DISPID_TE_MAX ? S_OK : S_FALSE;
		}
	}
	return DISP_E_UNKNOWNNAME;
}

HRESULT teGetMemberName(DISPID id, BSTR *pbstrName)
{
	if (id >= DISPID_COLLECTION_MIN && id <= DISPID_TE_MAX) {
		wchar_t szName[8];
		swprintf_s((wchar_t *)&szName, 8, L"%d", id - DISPID_COLLECTION_MIN);
		*pbstrName = ::SysAllocString(szName);
		return S_OK;
	}
	return E_NOTIMPL;
}

int GetSizeOfStruct(LPOLESTR bs)
{
	DISPID nSize = 0;
	teGetDispId(structSize, sizeof(structSize), g_maps[MAP_SS], bs, &nSize, false); 
	return nSize;
}

LPWSTR GetSubStruct(BSTR bs)
{
	struct
	{
		LPWSTR name;
		LPWSTR sub;
	} s[] = {
		{ L"EncoderParameters", L"EncoderParameter" },
	};
	for (int i = _countof(s); i--;) {
		if (lstrcmpi(bs, s[i].name) == 0) {
			return s[i].sub;
		}
	}
	return bs;
}

int GetVTOfStruct(BSTR bs)
{
	return lstrcmpi(bs, L"VARIANT") ? VT_NULL : VT_VARIANT;
}

int GetOffsetOfStruct(BSTR bs)
{
	return lstrcmpi(bs, L"EncoderParameters") ? 0 : sizeof(UINT);
}
/*
	struct
	{
		LONG   size;
		LPWSTR name;
	} s[] = {
		{ sizeof(UINT), L"EncoderParameters" },
	};
	for (int i = _countof(s); i--;) {
		if (lstrcmpi(bs, s[i].name) == 0) {
			return s[i].size;
		}
	}
	return 0;
}*/


int GetSizeOf(VARIANT *pv)
{
	if (pv->vt == VT_BSTR) {
		return GetSizeOfStruct(pv->bstrVal) + GetOffsetOfStruct(pv->bstrVal);
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
			int n = SysStringLen(pv->bstrVal);
			if (n > 1 && pv->bstrVal[n - 1] == _T('%')) {
				pv->bstrVal[n - 1] = NULL;
				VARIANT v;
				VariantInit(&v);
				if SUCCEEDED(VariantChangeType(&v, pv, 0, VT_R4)) {
					i = (int)(v.fltVal * 100);
					if (i > 10000) {
						i = 10000;	//Not exceed 100%
					}
				}
				VariantClear(&v);
				pv->bstrVal[n - 1] = _T('%');
			}
		}
		if (i == 0) {
			return GetIntFromVariant(pv) * 0x4000;
		}
		return i;
	}
	else {
		return GetIntFromVariant(pv);
	}
}

VOID GetPointFormVariant(POINT *pt, VARIANT *pv)
{
	CteMemory *pMem;
	if (GetMemFormVariant(pv, &pMem)) {
		pt->x = pMem->GetLong(0);
		pt->y = pMem->GetLong(1);
		pMem->Release();
		return;
	}
	int nPt;
	nPt = GetIntFromVariant(pv);
	pt->x = LOWORD(nPt);
	pt->y = HIWORD(nPt);
}


//There is no need for SysFreeString
BSTR GetLPWSTRFromVariant(VARIANT *pv)
{
	if (pv->vt == (VT_VARIANT | VT_BYREF)) {
		return GetLPWSTRFromVariant(pv->pvarVal);
	}
	switch (pv->vt) {
		case VT_BSTR:
		case VT_LPWSTR:
			return pv->bstrVal;
		default:
			return (BSTR)GetLLFromVariant(pv);
	}
}

VOID GetpDataFromVariant(UCHAR **ppc, int *pnLen, VARIANT *pv)
{
	if (pv->vt == (VT_VARIANT | VT_BYREF)) {
		return GetpDataFromVariant(ppc, pnLen, pv->pvarVal);
	}
	if (pv->vt & VT_ARRAY) {
		*ppc = (UCHAR *)pv->parray->pvData;
		*pnLen = pv->parray->rgsabound[0].cElements * SizeOfvt(pv->vt);
		return;
	}
	BSTR bs = GetLPWSTRFromVariant(pv);
	if (bs) {
		*pnLen = ::SysStringByteLen(bs);
		*ppc = (UCHAR *)bs;
		return;
	}
	CteMemory *pMem;
	if (GetMemFormVariant(pv, &pMem)) {
		*ppc = (UCHAR *)pMem->m_pc; 
		*pnLen = pMem->m_nSize;
		pMem->Release();
		return;
	}
	*ppc = NULL;
	*pnLen = 0;
}

HRESULT Invoke5(IDispatch *pdisp, DISPID dispid, WORD wFlags, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs)
{
	HRESULT hr;

	int i;
	// DISPPARAMS
	DISPPARAMS dispParams;
	dispParams.rgvarg = pvArgs;
	dispParams.cArgs = abs(nArgs);
	DISPID dispidName = DISPID_PROPERTYPUT;
	if (wFlags & DISPATCH_PROPERTYPUT) {
		dispParams.cNamedArgs = 1;
		dispParams.rgdispidNamedArgs = &dispidName;
	}
	else {
		dispParams.rgdispidNamedArgs = NULL;
		dispParams.cNamedArgs = 0;
	}
	try {
		hr = pdisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
			wFlags, &dispParams, pvResult, NULL, NULL);
	} catch (...) {
		hr = E_FAIL;
	}
	// VARIANT Clean-up of an array
	if (pvArgs && nArgs >= 0) {
		for (i = nArgs - 1; i >=  0; i--){
			VariantClear(&pvArgs[i]);
		}
		delete[] pvArgs;
		pvArgs = NULL;
	}
	return hr;
}

HRESULT Invoke4(IDispatch *pdisp, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs)
{
	return Invoke5(pdisp, DISPID_VALUE, DISPATCH_METHOD, pvResult, nArgs, pvArgs);
}

HRESULT ParseScript(LPOLESTR lpScript, LPOLESTR lpLang, VARIANT *pv, IUnknown *pOnError, IDispatch **ppdisp)
{
	HRESULT hr = E_FAIL;
	CLSID clsid;
	IActiveScript *pas = NULL;
	if (PathMatchSpec(lpLang, L"J*Script")) {
		CoCreateInstance(CLSID_JScriptChakra, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pas));
	}
	if (pas == NULL && teCLSIDFromProgID(lpLang, &clsid) == S_OK) {
		CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pas));
	}
	if (pas) {
		IActiveScriptProperty *paspr;
		if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&paspr))) {
			VARIANT v;
			VariantInit(&v);
			v.vt = VT_I4;
			v.lVal = 0;
			while (++v.lVal <= 256 && paspr->SetProperty(SCRIPTPROP_INVOKEVERSIONING, NULL, &v) == S_OK) {
			}
		}
		IUnknown *punk = NULL;
		FindUnknown(pv, &punk);
		CteActiveScriptSite *pass = new CteActiveScriptSite(punk, pOnError);
		pas->SetScriptSite(pass);
		pass->Release();
		IActiveScriptParse *pasp;
		if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&pasp))) {
			hr = pasp->InitNew();
			if (punk) {
				IDispatchEx *pdex;
				if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pdex))) {
					DISPID dispid;
					HRESULT hr = pdex->GetNextDispID(fdexEnumAll, DISPID_STARTENUM, &dispid);
					while (hr == NOERROR) {
						BSTR bs;
						if (pdex->GetMemberName(dispid, &bs) == S_OK) {
							pas->AddNamedItem(bs, SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE | SCRIPTITEM_GLOBALMEMBERS);
							::SysFreeString(bs);
						}
						hr = pdex->GetNextDispID(fdexEnumAll, dispid, &dispid);
					}
					pdex->Release();
				}
			}
			else if (pv && pv->boolVal) {
				pas->AddNamedItem(L"window", SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE | SCRIPTITEM_GLOBALMEMBERS);
			}
			VARIANT v;
			VariantInit(&v);
			hr = pasp->ParseScriptText(lpScript, NULL, NULL, NULL, 0, 1, SCRIPTTEXT_ISPERSISTENT | SCRIPTTEXT_ISVISIBLE, &v, NULL);
			if (hr == S_OK) {
				pas->SetScriptState(SCRIPTSTATE_CONNECTED);
				if (ppdisp) {
					IDispatch *pdisp;
					if (pas->GetScriptDispatch(NULL, &pdisp) == S_OK) {
						CteDispatch *odisp = new CteDispatch(pdisp, 2);
						pdisp->Release();
						if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&odisp->m_pActiveScript))) {
							odisp->QueryInterface(IID_PPV_ARGS(ppdisp));
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

HRESULT tePropertyGet(IDispatch *pdisp, LPOLESTR sz, VARIANT *pv)
{
	DISPID dispid;
	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &dispid);
	if (hr == S_OK) {
		hr = Invoke5(pdisp, dispid, DISPATCH_PROPERTYGET, pv, 0, NULL);
	}
	return hr;
}

HRESULT DoFunc(int nFunc, PVOID pObj, HRESULT hr)
{
	if (g_pOnFunc[nFunc]) {
		VARIANT vResult;
		VariantInit(&vResult);
		VARIANTARG *pv = GetNewVARIANT(1);
		teSetObject(&pv[0], pObj);
		Invoke4(g_pOnFunc[nFunc], &vResult, 1, pv);
		return GetIntFromVariant(&vResult);
	}
	return hr;
}

VOID GetNewArray(IDispatch **ppArray)
{
	VARIANT v;
	VariantInit(&v);
	if (Invoke4(g_pArray, &v, 0, NULL) == S_OK) {
		GetDispatch(&v, ppArray);
		VariantClear(&v);
	}
}

HRESULT DoStatusText(PVOID pObj, LPCWSTR lpszStatusText, int iPart)
{
	HRESULT hr = E_NOTIMPL;
	if (g_pOnFunc[TE_OnStatusText]) {
		VARIANTARG *pv;
		VARIANT vResult;
		VariantInit(&vResult);
		pv = GetNewVARIANT(3);
		teSetObject(&pv[2], pObj);
		pv[1].vt = VT_BSTR;
		pv[1].bstrVal = SysAllocString(lpszStatusText);
		pv[0].vt = VT_I4;
		pv[0].lVal = 0;
		Invoke4(g_pOnFunc[TE_OnStatusText], &vResult, 3, pv);
		hr = GetIntFromVariant(&vResult);
	}
	return hr;
}

LONGLONG DoHitTest(PVOID pObj, POINT pt, UINT flags)
{
	if (g_pOnFunc[TE_OnHitTest]) {
		VARIANT vResult;
		VariantInit(&vResult);
		VARIANTARG *pv = GetNewVARIANT(3);
		teSetObject(&pv[2], pObj);
		CteMemory *pstPt = new CteMemory(2 * sizeof(int), NULL, VT_NULL, 1, L"POINT");
		pstPt->SetLong(0, pt.x);
		pstPt->SetLong(1, pt.y);
		teSetObject(&pv[1], pstPt);
		pstPt->Release();
		pv[0].lVal = flags;
		pv[0].vt = VT_I4;
		Invoke4(g_pOnFunc[TE_OnHitTest], &vResult, 3, pv);
		return GetLLFromVariant(&vResult);
	}
	return -1;
}

HRESULT DragSub(int nFunc, PVOID pObj, CteFolderItems *pDragItems, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	HRESULT hr = E_FAIL;
	VARIANT vResult;
	VARIANTARG *pv;
	CteMemory *pstPt;
	CteMemory *pstEffect;

	if (g_pOnFunc[nFunc]) {
		try {
			if (InterlockedIncrement(&g_nProcDrag) < 10) {
				pv = GetNewVARIANT(5);
				teSetObject(&pv[4], pObj);
				teSetObject(&pv[3], pDragItems);

				pv[2].lVal = grfKeyState;
				pv[2].vt = VT_I4;

				pstPt = new CteMemory(2 * sizeof(int), NULL, VT_NULL, 1, L"POINT");
				pstPt->SetLong(0, pt.x);
				pstPt->SetLong(1, pt.y);
				teSetObject(&pv[1], pstPt);
				pstPt->Release();

				pstEffect = new CteMemory(sizeof(int), (char *)pdwEffect, VT_NULL, 1, L"DWORD");
				teSetObject(&pv[0], pstEffect);
				pstEffect->Release();

				if SUCCEEDED(Invoke4(g_pOnFunc[nFunc], &vResult, 5, pv)) {
					hr = GetIntFromVariant(&vResult);
				}
			}
		}
		catch(...) {}
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

HRESULT MessageSub(int nFunc, PVOID pObj, MSG *pMsg, HRESULT hrDefault)
{
	VARIANT vResult;
	VARIANTARG *pv;

	if (g_pOnFunc[nFunc]) {
		pv = GetNewVARIANT(5);
		teSetObject(&pv[4], pObj);
		SetVariantLL(&pv[3], (LONGLONG)pMsg->hwnd);
		pv[2].lVal = (long)pMsg->message;
		pv[2].vt = VT_I4;
		SetVariantLL(&pv[1], pMsg->wParam);
		SetVariantLL(&pv[0], pMsg->lParam);
		if SUCCEEDED(Invoke4(g_pOnFunc[nFunc], &vResult, 5, pv)) {
			return GetIntFromVariantClear(&vResult);
		}
	}
	return hrDefault;
}

HRESULT MessageSubPt(int nFunc, PVOID pObj, MSG *pMsg)
{
	VARIANT vResult;

	if (g_pOnFunc[nFunc]) {
		VARIANTARG *pv = GetNewVARIANT(5);
		teSetObject(&pv[4], pObj);
		SetVariantLL(&pv[3], (LONGLONG)pMsg->hwnd); 
		pv[2].lVal = (long)pMsg->message;
		pv[2].vt = VT_I4;
		SetVariantLL(&pv[1], pMsg->wParam);

		CteMemory *pMem = new CteMemory(2 * sizeof(int), NULL, VT_NULL, 1, L"POINT");
		pMem->SetLong(0, pMsg->pt.x);
		pMem->SetLong(1, pMsg->pt.y);
		teSetObject(&pv[0], pMem);
		pMem->Release();

		if SUCCEEDED(Invoke4(g_pOnFunc[nFunc], &vResult, 5, pv)) {
			return GetIntFromVariantClear(&vResult);
		}
	}
	return S_FALSE;
}

int GetShellBrowser2(CteShellBrowser *pSB)
{
	int i = MAX_FV;
	while (--i >= 0) {
		if (g_pSB[i]) {
			if (!g_pSB[i]->m_bEmpty) {
				if (g_pSB[i] == pSB) {
					return i;
				}
			}
		}
	}
	return 0;
}

CteShellBrowser* GetNewShellBrowser(CteTabs *pTabs)
{
	CteShellBrowser *pSB = NULL;
	int i = MAX_FV;
	while (--i >= 0) {
		pSB = g_pSB[i]; 
		if (pSB) {
			if (pSB->m_bEmpty) {
				pSB->Init(pTabs, false);
				break;
			}
		}
		else {
			pSB = new CteShellBrowser(pTabs);
			pSB->m_nSB = i;
			g_pSB[i] = pSB;
			break;
		}
	}
	if (pTabs && i >= 0) {
		TC_ITEM tcItem;
		tcItem.mask = TCIF_PARAM;
		tcItem.lParam = i;
		if (pTabs) {
			TabCtrl_InsertItem(pTabs->m_hwnd, pTabs->m_nIndex + 1, &tcItem);
			if (pTabs->m_nIndex < 0) {
				pTabs->m_nIndex = 0;
			}
		}
		return pSB;
	}
	return NULL;
}

CteTabs* GetNewTC()
{
	int i = MAX_TC;
	while (--i >= 0) {
		if (g_pTC[i]) {
			if (g_pTC[i]->m_bEmpty) {
				return g_pTC[i];
			}
		}
		else {
			g_pTC[i] = new CteTabs(); 
			g_pTC[i]->m_nTC = i;
			return g_pTC[i];
		}
	}
	return NULL;
}

HRESULT ControlFromhwnd(IDispatch **ppdisp, HWND hwnd)
{
	*ppdisp = NULL;
	int i;
	i = SBfromhwnd(hwnd);
	if (i >= 0) {
		return g_pSB[i]->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	i = TCfromhwnd(hwnd);
	if (i >= 0) {
		return g_pTC[i]->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	i = TVfromhwnd(hwnd, false);
	if (i >= 0) {
		return g_pSB[i]->m_pTV->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	if (g_hwndMain == hwnd && g_pTE) {
		return g_pTE->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	if (IsChild(g_pWebBrowser->get_HWND(), hwnd)) {
		return g_pWebBrowser->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	return E_FAIL;
}

LRESULT CALLBACK MenuKeyProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = 1;
	if (nCode >= 0) {
		if (nCode == HC_ACTION) {
			MSG msg;
			msg.hwnd = GetFocus();
			msg.message = (lParam >= 0) ? WM_KEYDOWN : WM_KEYUP;
			msg.wParam = wParam;
			msg.lParam = lParam;
			try {
				if (InterlockedIncrement(&g_nProcKey) == 1) {
					if (MessageSub(TE_OnKeyMessage, g_pTE, &msg, S_FALSE) == S_OK) {
						lResult = 0;
					}
				}
			}
			catch(...) {}
			::InterlockedDecrement(&g_nProcKey);
		}
	}
	return lResult ? CallNextHookEx(g_hMenuKeyHook, nCode, wParam, lParam) : lResult;
}

LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	HRESULT hrResult = S_FALSE;
	IDispatch *pdisp = NULL;
	MOUSEHOOKSTRUCTEX *pmse;
	if (nCode >= 0) {
		if (nCode == HC_ACTION) {
			if (g_pOnFunc[TE_OnMouseMessage]) {
				pmse = (MOUSEHOOKSTRUCTEX *)lParam;

				if (ControlFromhwnd(&pdisp, pmse->hwnd) == S_OK) {
					try {
						if (InterlockedIncrement(&g_nProcMouse) == 1) {
							int nTV = TVfromhwnd(pmse->hwnd, false);
							TVHITTESTINFO ht;
							if (nTV >= 0) {
								ht.pt.x = pmse->pt.x;
								ht.pt.y = pmse->pt.y;
								ht.flags = TVHT_ONITEM;
								ScreenToClient(pmse->hwnd, &ht.pt);
								TreeView_HitTest(pmse->hwnd, &ht);
								if (wParam == WM_LBUTTONUP || wParam == WM_LBUTTONDBLCLK || 
									wParam == WM_RBUTTONUP || wParam == WM_RBUTTONDBLCLK ||							
									wParam == WM_MBUTTONUP || wParam == WM_MBUTTONDBLCLK ||
									wParam == WM_XBUTTONUP || wParam == WM_XBUTTONDBLCLK) {
									if (ht.hItem && ht.flags & TVHT_ONITEM) {
										TreeView_SelectItem(pmse->hwnd, ht.hItem);
									}
								}
							}
							VARIANT vResult;
							VARIANTARG *pv = GetNewVARIANT(7);
							teSetObject(&pv[6], pdisp);
							SetVariantLL(&pv[5], (LONGLONG)pmse->hwnd);
							g_hwndLast = pmse->hwnd;
							SetVariantLL(&pv[4], wParam);
							g_LastMsg = wParam;
							pv[3].lVal = pmse->mouseData;
							pv[3].vt = VT_I4;
							CteMemory *pMem;
							pMem = new CteMemory(2 * sizeof(int), NULL, VT_NULL, 1, L"POINT");
							pMem->SetLong(0, pmse->pt.x);
							pMem->SetLong(1, pmse->pt.y);
							teSetObject(&pv[2], pMem);
							pMem->Release();
							pv[1].lVal = pmse->wHitTestCode;
							pv[1].vt = VT_I4;
							SetVariantLL(&pv[0], pmse->dwExtraInfo);
							try {
								if SUCCEEDED(Invoke4(g_pOnFunc[TE_OnMouseMessage], &vResult, 7, pv)) {
									hrResult = GetIntFromVariantClear(&vResult);
								}
							} catch(...) {}
#ifdef _2000XP
							if (hrResult != S_OK && nTV >= 0) {
								if (wParam == WM_LBUTTONUP || wParam == WM_LBUTTONDBLCLK || wParam == WM_MBUTTONUP) {
									if (g_pSB[nTV]->m_pTV->m_pShellNameSpace && ht.hItem) {
										NSTCSTYLE nStyle = NSTCECT_LBUTTON;
										if (wParam == WM_LBUTTONDBLCLK) {
											nStyle = NSTCECT_LBUTTON | NSTCECT_DBLCLICK;
										}
										else if (wParam == WM_MBUTTONUP) {
											nStyle = NSTCECT_MBUTTON;
										}
										hrResult = g_pSB[nTV]->m_pTV->OnItemClick(NULL, ht.flags, nStyle);
									}
								}
							}
#endif
						}
						else if (wParam == g_LastMsg && g_hwndLast == pmse->hwnd) {
							hrResult = S_OK;
						}
					}
					catch(...) {}
					::InterlockedDecrement(&g_nProcMouse);
					pdisp->Release();
				}
			}
		}
	}
	return (hrResult != S_OK) ? CallNextHookEx(g_hMouseHook, nCode, wParam, lParam) : TRUE;
}

#ifdef _2000XP
LRESULT CALLBACK TETVProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	HRESULT Result = S_FALSE;
	int nTV = TVfromhwnd(hwnd, true);
	if (nTV < 0) {
		return 0;
	}
	if (msg == WM_NOTIFY) {
		LPNMHDR pnmhdr = (LPNMHDR)lParam;
		if (pnmhdr->code == NM_RCLICK) {
			try {
				if (InterlockedIncrement(&g_nProcTV) == 1) {
					MSG msg1;
					msg1.hwnd = hwnd;
					msg1.message = WM_CONTEXTMENU;
					msg1.wParam = (WPARAM)hwnd;
					GetCursorPos(&msg1.pt);
					Result = MessageSubPt(TE_OnShowContextMenu, g_pSB[nTV]->m_pTV, &msg1);
				}
			}
			catch(...) {}
			::InterlockedDecrement(&g_nProcTV);
		}
	}
	if (msg == WM_CONTEXTMENU) {	//Double-clicking is the default menu.
		Result = S_OK;				//Treated with NM_RCLICK
	}
	return Result ? CallWindowProc(g_pSB[nTV]->m_pTV->m_DefProc, hwnd, msg, wParam, lParam) : 0;
}
#endif

LRESULT CALLBACK TETVProc2(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	int nTV = TVfromhwnd(hwnd, true);
	if (nTV < 0) {
		return 0;
	}
	if (g_pSB[nTV]->m_pTV->m_bMain) {
		if (msg == WM_ENTERMENULOOP) {
			g_pSB[nTV]->m_pTV->m_bMain = false;
		}
	}
	else {
		if (msg == WM_EXITMENULOOP) {
			g_pSB[nTV]->m_pTV->m_bMain = true;
		}
		else if (msg >= TV_FIRST && msg < TV_FIRST + 0x100) {
			return 0;
		}
	}
	if (msg == WM_COMMAND) {
		LRESULT Result = S_FALSE;
		try {
			if (InterlockedIncrement(&g_nProcTV) == 1) {
				if (g_pOnFunc[TE_OnCommand]) {
					MSG msg1;
					msg1.hwnd = hwnd;
					msg1.message = msg;
					msg1.wParam = wParam;
					msg1.lParam = lParam;
					Result = MessageSub(TE_OnCommand, g_pSB[nTV]->m_pTV, &msg1, S_FALSE);
				}
			}
		}
		catch(...) {}
		::InterlockedDecrement(&g_nProcTV);
		if (Result == 0) {
			return 0;
		}
	}
	return CallWindowProc(g_pSB[nTV]->m_pTV->m_DefProc2, hwnd, msg, wParam, lParam);
}

LRESULT TELVProc0(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, int nSB)
{
	LRESULT Result = S_FALSE;

	if (msg == WM_CONTEXTMENU || msg == SB_SETTEXTA || msg == SB_SETTEXT || msg == WM_COMMAND) {
		try {
			if (InterlockedIncrement(&g_nProcFV) == 1) {
				switch (msg) {
					case WM_CONTEXTMENU:
						IDispatch *pdisp;
						if (ControlFromhwnd(&pdisp, hwnd) == S_OK) {
							MSG msg1;
							msg1.hwnd = hwnd;
							msg1.message = msg;
							msg1.wParam = wParam;
							msg1.pt.x = LOWORD(lParam);
							msg1.pt.y = HIWORD(lParam);
							Result = MessageSubPt(TE_OnShowContextMenu, pdisp, &msg1);
						}
						break;
					case SB_SETTEXTA:
						int nLen;
						nLen = MultiByteToWideChar(CP_ACP, 0, (LPCSTR)lParam, -1, NULL, NULL);
						LPWSTR lpszW;
						lpszW = new WCHAR[nLen];
						MultiByteToWideChar(CP_ACP, 0, (LPCSTR)lParam, -1, lpszW, nLen);
						Result = DoStatusText(g_pSB[nSB], const_cast<LPCWSTR>(lpszW), static_cast<int>(wParam));
						delete []lpszW;
						break;
					case SB_SETTEXT:
						Result = DoStatusText(g_pSB[nSB], (LPCWSTR)lParam, static_cast<int>(wParam));
						break;
					case WM_COMMAND:
						if (g_pOnFunc[TE_OnCommand]) {
							MSG msg1;
							msg1.hwnd = hwnd;
							msg1.message = msg;
							msg1.wParam = wParam;
							msg1.lParam = lParam;
							Result = MessageSub(TE_OnCommand, g_pSB[nSB], &msg1, S_FALSE);
						}
						break;
				}//end_switch
			}
		}
		catch(...) {}
		::InterlockedDecrement(&g_nProcFV);
	}
	if (g_pSB[nSB]->m_bNoRowSelect) {
		g_pSB[nSB]->m_bNoRowSelect = false;
		HWND hList = FindWindowEx(hwnd, 0, WC_LISTVIEW, NULL);
		if (hList) {
			PostMessage(hList, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, SendMessage(hList, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & (~LVS_EX_FULLROWSELECT));
		}
	}
	return Result;
}

LRESULT CALLBACK TELVProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	int nSB = SBfromhwnd(hwnd);
	if (nSB < 0) {
		return 0;
	}
	LRESULT Result = TELVProc0(hwnd, msg, wParam, lParam, nSB);

	if (msg == WM_COMMAND) {
		WORD wID;
		wID = LOWORD(wParam);
		if (wID >= 28713 && wID <= 28759) {
			CallWindowProc((WNDPROC)g_pSB[nSB]->m_DefProc[0], hwnd, msg, wParam, lParam); 
			g_pSB[nSB]->m_bNoRowSelect = !(g_pSB[nSB]->m_param[SB_FolderFlags] & FWF_FULLROWSELECT);
			Result = S_OK;
		}
	}
	return Result ? CallWindowProc((WNDPROC)g_pSB[nSB]->m_DefProc[0], hwnd, msg, wParam, lParam) : 0;
}

LRESULT CALLBACK TELVProc2(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	int nSB = SBfromhwnd(hwnd);
	if (nSB < 0) {
		return 0;
	}
	return TELVProc0(hwnd, msg, wParam, lParam, nSB) ? CallWindowProc((WNDPROC)g_pSB[nSB]->m_DefProc[1], hwnd, msg, wParam, lParam) : 0;
}

LRESULT CALLBACK TEBTProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT Result = 1;
	int nTC;

	switch (msg) {
		case WM_NOTIFY:
			switch (((LPNMHDR)lParam)->code) {
				case TCN_SELCHANGE:
					nTC = TCfromhwnd(hwnd);
					if (nTC >= 0) {
						if (g_pTabs->m_hwnd != g_pTC[nTC]->m_hwnd && g_pTC[nTC]->m_bVisible) {
							g_pTabs = g_pTC[nTC];
						}
						g_pTC[nTC]->TabChanged(true);
						Result = 0;
					}
					break;
				case TCN_SELCHANGING:
					nTC = TCfromhwnd(hwnd);
					if (nTC >= 0) {
						Result = DoFunc(TE_OnSelectionChanging, g_pTC[nTC], 0);
					}
					break;
				case TTN_GETDISPINFO:
					if (g_pOnFunc[TE_OnToolTip]) {
						nTC = TCfromhwnd(hwnd);
						if (nTC >= 0) {
							VARIANT vResult;
							VariantInit(&vResult);
							VARIANTARG *pv = GetNewVARIANT(2);
							VariantInit(&pv[1]);
							teSetObject(&pv[1], g_pTC[nTC]);
							VariantInit(&pv[0]);
							SetVariantLL(&pv[0], ((LPNMHDR)lParam)->idFrom);
							Invoke4(g_pOnFunc[TE_OnToolTip], &vResult, 2, pv);
							VARIANT vText;
							teVariantChangeType(&vText, &vResult, VT_BSTR);
							::ZeroMemory(g_szText, sizeof(g_szText));
							lstrcpyn(g_szText, vText.bstrVal, sizeof(g_szText));
							((LPTOOLTIPTEXT)lParam)->lpszText = g_szText;
							VariantClear(&vResult);
							VariantClear(&vText);
						}
					}
					break;
			}
			break;
		case WM_SETFOCUS:
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
			CheckChangeTabTC(hwnd, true);
			break;
		case WM_CONTEXTMENU:
			nTC = TCfromhwnd(hwnd);
			if (nTC >= 0) {
				MSG msg1;
				msg1.hwnd = g_pTC[nTC]->m_hwnd;
				msg1.message = msg;
				msg1.wParam = wParam;
				msg1.pt.x = LOWORD(lParam);
				msg1.pt.y = HIWORD(lParam);
				MessageSubPt(TE_OnShowContextMenu, g_pTC[nTC], &msg1);
				Result = 0;
			}
			break;
		case WM_VSCROLL:
			nTC = TCfromhwnd(hwnd);
			if (nTC >= 0) {
				CteTabs *pTC = g_pTC[nTC];
				switch(LOWORD(wParam)){
					case SB_LINEUP:		pTC->m_si.nPos -= 16; break;
					case SB_LINEDOWN:	pTC->m_si.nPos += 16; break;
					case SB_PAGEUP:		pTC->m_si.nPos -= pTC->m_si.nTrackPos / 2; break;
					case SB_PAGEDOWN:	pTC->m_si.nPos += pTC->m_si.nTrackPos / 2; break;
					case SB_THUMBPOSITION: 
					case SB_THUMBTRACK: pTC->m_si.nPos = HIWORD(wParam); break;
				}
				if (pTC->m_si.nPos > pTC->m_si.nMax) {
					pTC->m_si.nPos = pTC->m_si.nMax;
				}
				if (pTC->m_si.nPos < 0) {
					pTC->m_si.nPos = 0;
				}
				SetScrollInfo(pTC->m_hwndButton, SB_VERT, &pTC->m_si, TRUE);
				ArrangeWindow();
			}
			break;
	}
	return Result ? CallWindowProc(g_DefBTProc, hwnd, msg, wParam, lParam) : 0;
}


LRESULT CALLBACK TETCProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT Result = 1;
	int nIndex, nCount, nTC;

	switch (msg) {
		case TCM_DELETEITEM:
			if (lParam) {
				lParam = 0;
			}
			else {
				nTC = TCfromhwnd(hwnd);
				if (nTC >= 0) {
					CteShellBrowser *pSB = NULL;
					pSB = g_pTC[nTC]->GetShellBrowser((int)wParam);
					CallWindowProc(g_DefTCProc, hwnd, msg, wParam, lParam);
					Result = 0;
					if (pSB) {
						pSB->Close(true);
					}
					nIndex = TabCtrl_GetCurSel(hwnd);
					if (nIndex >= 0) {
						g_pTC[nTC]->m_nIndex = nIndex;
					}
					else {
						nCount = TabCtrl_GetItemCount(hwnd);
						nIndex = g_pTC[nTC]->m_nIndex;
						if ((int)wParam > nIndex) {
							nIndex--;
						}
						if (nIndex >= nCount) {
							nIndex = nCount - 1;
						}
						if (nIndex < 0) {
							nIndex = 0;
						}
						g_pTC[nTC]->m_nIndex = -1;
						//TabCtrl_SetCurSel(hwnd, nIndex);
						CallWindowProc(g_DefTCProc, hwnd, TCM_SETCURSEL, nIndex, 0);
					}
					g_pTC[nTC]->TabChanged(true);
					ArrangeWindow();
				}
			}
			break;
		case TCM_SETCURSEL:
			if (wParam != (UINT_PTR)TabCtrl_GetCurSel(hwnd)) { 
				nTC = TCfromhwnd(hwnd);
				if (nTC >= 0) {
					CallWindowProc(g_DefTCProc, hwnd, msg, wParam, lParam);
					Result = 0;
					g_pTC[nTC]->TabChanged(true);
					if (g_pTE->m_param[TE_Tab] && g_pTC[nTC]->m_param[TE_Align] == 1) {
						if (g_pTC[nTC]->m_param[TE_Flags] & TCS_SCROLLOPPOSITE) {
							ArrangeWindow();
						}
					}
				}
			}
			break;
		case TCM_SETITEM:
			nTC = TCfromhwnd(hwnd);
			if (nTC >= 0) {
				if (g_pTC[nTC]->m_param[TE_Flags] & TCS_FIXEDWIDTH) {
					CallWindowProc(g_DefTCProc, hwnd, msg, wParam, lParam);
					Result = 0;
					CallWindowProc(g_DefTCProc, g_pTC[nTC]->m_hwnd, TCM_SETITEMSIZE, 0, MAKELPARAM(g_pTC[nTC]->m_param[TC_TabWidth], g_pTC[nTC]->m_param[TC_TabHeight] + 1));
					CallWindowProc(g_DefTCProc, g_pTC[nTC]->m_hwnd, TCM_SETITEMSIZE, 0, MAKELPARAM(g_pTC[nTC]->m_param[TC_TabWidth], g_pTC[nTC]->m_param[TC_TabHeight]));
				}
			}
			break;
		case WM_SETFOCUS:
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
			CheckChangeTabTC(hwnd, true);
			break;
	}
	return Result ? CallWindowProc(g_DefTCProc, hwnd, msg, wParam, lParam) : 0;
}

HRESULT MessageProc(MSG *pMsg)
{
	HRESULT hrResult = S_FALSE;
	IDispatch *pdisp = NULL;
	IShellBrowser *pSB = NULL;
	IShellView *pSV = NULL;
	IWebBrowser2 *pWB = NULL;
	IOleInPlaceActiveObject *pActiveObject = NULL;

	if (pMsg->message >= WM_KEYFIRST && pMsg->message <= WM_KEYLAST) {
		try {
			if (InterlockedIncrement(&g_nProcKey) == 1) {
				if (ControlFromhwnd(&pdisp, pMsg->hwnd) == S_OK) {
					if (MessageSub(TE_OnKeyMessage, pdisp, pMsg, S_FALSE) == S_OK) {
						hrResult = S_OK;
					}
					else if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pSB))) {
						if SUCCEEDED(pSB->QueryActiveShellView(&pSV)) {
							if (pSV->TranslateAcceleratorW(pMsg) == S_OK) {
								hrResult = S_OK;
							}
							pSV->Release();
						}
						pSB->Release();
					}
					else if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pWB))) {
						if SUCCEEDED(pWB->QueryInterface(IID_PPV_ARGS(&pActiveObject))) {
							if (pActiveObject->TranslateAcceleratorW(pMsg) == S_OK) {
								hrResult = S_OK;
							}
							pActiveObject->Release();
						}
						pWB->Release();
					}
					pdisp->Release();
				}
			}
		}
		catch(...) {}
		::InterlockedDecrement(&g_nProcKey);
	}
	return hrResult;
}

VOID Initlize()
{
	g_maps[MAP_TE] = SortMethod(methodTE, sizeof(methodTE));
	g_maps[MAP_SB] = SortMethod(methodSB, sizeof(methodSB));
	g_maps[MAP_TC] = SortMethod(methodTC, sizeof(methodTC));
	g_maps[MAP_TV] = SortMethod(methodTV, sizeof(methodTV));
	g_maps[MAP_WB] = SortMethod(methodWB, sizeof(methodWB));
	g_maps[MAP_API] = SortMethod(methodAPI, sizeof(methodAPI));
	g_maps[MAP_Mem] = SortMethod(methodMem, sizeof(methodMem));
	g_maps[MAP_M2] = SortMethod(methodMem2, sizeof(methodMem2));
	g_maps[MAP_GB] = SortMethod(methodGB, sizeof(methodGB));
	g_maps[MAP_FIs] = SortMethod(methodFIs, sizeof(methodFIs));
	g_maps[MAP_FI] = SortMethod(methodFI, sizeof(methodFI));
	g_maps[MAP_DT] = SortMethod(methodDT, sizeof(methodDT));
	g_maps[MAP_CM] = SortMethod(methodCM, sizeof(methodCM));
	g_maps[MAP_CD] = SortMethod(methodCD, sizeof(methodCD));
	g_maps[MAP_SS] = SortMethod(structSize, sizeof(structSize));
}

VOID Finalize()
{
	try {
		for (int i = _countof(g_maps); i--;) {
			delete [] g_maps[i];
		}
		lpfnSHCreateItemFromIDList = NULL;
		for (int i = MAX_CSIDL; i--;) {
			::CoTaskMemFree(g_pidls[i]);
			g_pidls[i] = NULL;
		}
		if (g_hShell32) {
			FreeLibrary(g_hShell32);
		}
		if (g_hKernel32) {
			FreeLibrary(g_hKernel32);
		}
		if (g_hCrypt32) {
			FreeLibrary(g_hCrypt32);
		}
		if (g_bsCmdLine) {
			SysFreeString(g_bsCmdLine);
		}
		if (g_pCrcTable) {
			delete [] g_pCrcTable;
		}
	}
	catch (...) {
	}
	try {
		Gdiplus::GdiplusShutdown(g_Token);
	} catch (...) {}
	try {
		g_pArray && g_pArray->Release();
	} catch (...) {}
	try {
		g_pObject && g_pObject->Release();
	} catch (...) {}
	try {
		g_pJS && g_pJS->Release();
//		g_pHtmlDoc && g_pHtmlDoc->Release();
	} catch (...) {}
	try {
		g_pSW && g_pSW->Release();
	} catch (...) {}
	try {
		g_psha && g_psha->Release();
	} catch (...) {}
	try {
		g_pDesktopFolder && g_pDesktopFolder->Release();
	} catch (...) {}
	g_pDesktopFolder = NULL;
	try {
		::OleUninitialize();
	}
	catch (...) {
	}
	ReleaseMutex(g_hMutex);
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
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_TE));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_BTNFACE + 1);
	wcex.lpszMenuName	= 0;
	wcex.lpszClassName	= WINDOW_CLASS;
	wcex.hIconSm		= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_16));

	return RegisterClassEx(&wcex);
}

int GetEncoderClsid(const WCHAR* pszName, CLSID* pClsid, LPWSTR pszMimeType)
{
	UINT num = 0;			// number of image encoders
	UINT size = 0;			// size of the image encoder array in bytes

	ImageCodecInfo* pImageCodecInfo = NULL;

	GetImageEncodersSize(&num, &size);
	if(size == 0) {
		return -1;  // Failure
	}
	pImageCodecInfo = (ImageCodecInfo*) new char[size];
	if(pImageCodecInfo == NULL) {
		return -1;  // Failure
	}
	GetImageEncoders(num, size, pImageCodecInfo);

	for (UINT j = 0; j < num; j++)
	{
		if (PathMatchSpec(pszName, pImageCodecInfo[j].FilenameExtension) ||
				lstrcmp(pszName, pImageCodecInfo[j].MimeType) == 0) {
			if (pClsid) {
				*pClsid = pImageCodecInfo[j].Clsid;
			}
			if (pszMimeType) {
				lstrcpyn(pszMimeType, pImageCodecInfo[j].MimeType, 16);
			}
			delete [] pImageCodecInfo;
			return j;  // Success
		}
	}
	delete [] pImageCodecInfo;
	return -1;  // Failure
}

void teCalcClientRect(int *param, LPRECT rc, LPRECT rcClient) 
{
	if (param[TE_Left] & 0x3fff) {
		rc->left = (param[TE_Left] * (rcClient->right - rcClient->left)) / 10000 + rcClient->left;
	}
	else {
		rc->left = param[TE_Left] / 0x4000 + (param[TE_Left] >= 0 ? rcClient->left : rcClient->right);
	}
	if (param[TE_Top] & 0x3fff) {
		rc->top = (param[TE_Top] * (rcClient->bottom - rcClient->top)) / 10000 + rcClient->top;
	}
	else {
		rc->top = param[TE_Top] / 0x4000 + (param[TE_Top] >= 0 ? rcClient->top : rcClient->bottom);
	}
	if (param[TE_Width] & 0x3fff) {
		rc->right = param[TE_Width] * (rcClient->right - rcClient->left) / 10000 + rc->left;
	}
	else {
		rc->right = param[TE_Width] / 0x4000 + (param[TE_Width] >= 0 ? rc->left : rcClient->right);
	}

	if (rc->right > rcClient->right) {
		rc->right = rcClient->right;
	}
	if (param[TE_Height] & 0x3fff) {
		rc->bottom = param[TE_Height] * (rcClient->bottom - rcClient->top) / 10000 + rc->top;
	}
	else {
		rc->bottom = param[TE_Height] / 0x4000 + (param[TE_Height] >= 0 ? rc->top : rcClient->bottom);
	}
	if (rc->bottom > rcClient->bottom) {
		rc->bottom = rcClient->bottom;
	}
}

BOOL CanClose(PVOID pObj)
{
	return DoFunc(TE_OnClose, pObj, S_OK) != S_FALSE;
}

VOID ArrangeTree(CteShellBrowser *pSB, LPRECT rc)
{
	RECT rcTree = *rc;
	rcTree.right = rc->left + pSB->m_param[SB_TreeWidth];
	CteTreeView *pTV = pSB->m_pTV;
	if (pTV) {
		CopyRect(&pSB->m_pTV->m_rc, &rcTree);
		if (pTV->m_pShellNameSpace) {
			IOleInPlaceObject *pOleInPlaceObject;
			pTV->m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pOleInPlaceObject));
			pOleInPlaceObject->SetObjectRects(&rcTree, &rcTree);
			pOleInPlaceObject->Release();
		}
		else {
			MoveWindow(pTV->m_hwnd, rcTree.left, rcTree.top,
				rcTree.right - rcTree.left, rcTree.bottom - rcTree.top,	true);
		}	
		rc->left += pSB->m_param[SB_TreeWidth];
	}
}

VOID CALLBACK teTimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);
	switch (idEvent) {
		case TET_Create:
			if (g_pOnFunc[TE_OnCreate]) {
				DoFunc(TE_OnCreate, g_pTE, E_NOTIMPL);
				DoFunc(TE_OnCreate, g_pWebBrowser, E_NOTIMPL);
			}
			ShowWindow(hwnd, g_nCmdShow);
			break;
		case TET_Reload:
			if (g_bReload) {
				PostMessage(hwnd, WM_CLOSE, 0, 0);
				LPWSTR pszPath, pszQuoted;
				int i = teGetModuleFileName(NULL, &pszPath);
				pszQuoted = new WCHAR[i + 3];
				lstrcpy(pszQuoted, pszPath);
				delete [] pszPath;
				PathQuoteSpaces(pszQuoted);
				ShellExecute(hwnd, NULL, pszQuoted, NULL, NULL, SW_SHOWNORMAL);
				delete [] pszQuoted;
			}
			break;
		case TET_Size2:
			if (g_nSize > 0) {
				SetTimer(g_hwndMain, TET_Size2, 500, teTimerProc);
				break;
			}
			g_nSize++;
		case TET_Size:
			RECT rc, rcTab, rcClient;

			GetClientRect(hwnd, &rc);
			CopyRect(&rcClient, &rc);
			rcClient.left += g_pTE->m_param[TE_Left];
			rcClient.top += g_pTE->m_param[TE_Top];
			rcClient.right -= g_pTE->m_param[TE_Width];
			rcClient.bottom -= g_pTE->m_param[TE_Height];
			CteShellBrowser *pSB;

			if (g_pWebBrowser) {
				IOleInPlaceObject *pOleInPlaceObject;
				g_pWebBrowser->m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleInPlaceObject));
				pOleInPlaceObject->SetObjectRects(&rc, &rc);
				pOleInPlaceObject->Release();
			}

			int i;
			i = MAX_TC;
			while (--i >= 0) {
				CteTabs *pTC = g_pTC[i];
				if (pTC) {
					pTC->LockUpdate();
					try {
						pSB = pTC->GetShellBrowser(pTC->m_nIndex);
						teCalcClientRect(pTC->m_param, &rc, &rcClient);
						if (g_pOnFunc[TE_OnArrange]) {
							VARIANTARG *pv;
							pv = GetNewVARIANT(2);
							teSetObject(&pv[1], pTC);

							CteMemory *pstR;
							pstR = new CteMemory(sizeof(RECT), (char *)&rc, VT_NULL, 1, L"RECT");
							teSetObject(&pv[0], pstR);
							pstR->Release();

							Invoke4(g_pOnFunc[TE_OnArrange], NULL, 2, pv);
						}
						if (!pTC->m_bEmpty && pTC->m_bVisible) {
							int nAlign = g_pTE->m_param[TE_Tab] ? pTC->m_param[TE_Align] : 0;
							if (pSB && (pSB->m_param[SB_TreeAlign] == (TE_TV_Use | TE_TV_View | TE_TV_Left) ||
							   (nAlign == 0 && pSB->m_param[SB_TreeAlign] == (TE_TV_Use | TE_TV_View)))) {
								ArrangeTree(pSB, &rc);
							}
							MoveWindow(pTC->m_hwndStatic, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, TRUE);
							int left = rc.left;
							int top = rc.top;

							OffsetRect(&rc, -left, -top);
							CopyRect(&rcTab, &rc);
							int nCount = TabCtrl_GetItemCount(pTC->m_hwnd);
							int h = rc.bottom;
							if (nCount) {
								if (nAlign == 4 || nAlign == 5) {
									rc.right = pTC->m_param[TC_TabWidth];
								}
								DWORD dwStyle = pTC->GetStyle();
								if ((dwStyle & (TCS_BOTTOM | TCS_BUTTONS)) == (TCS_BOTTOM | TCS_BUTTONS)) {
									SetWindowLong(pTC->m_hwnd, GWL_STYLE, dwStyle & ~TCS_BOTTOM);
									TabCtrl_AdjustRect(pTC->m_hwnd, FALSE, &rc);
									SetWindowLong(pTC->m_hwnd, GWL_STYLE, dwStyle);
									int i = rcTab.bottom - rc.bottom;
									rc.bottom -= rc.top - i;
									rc.top = i;
								}
								else {
									TabCtrl_AdjustRect(pTC->m_hwnd, FALSE, &rc);
								}
								h = rcTab.bottom - (rc.bottom - rc.top) - 4;
								switch (nAlign) {
									case 0:						//none
										CopyRect(&rc, &rcTab);
										SetRect(&rcTab, 0, 0, 0, 0);
										break;
									case 2:						//top
										CopyRect(&rc, &rcTab);
										rcTab.bottom = h;
										rc.top += h;
										break;
									case 3:						//bottom
										CopyRect(&rc, &rcTab);
										rcTab.top = rcTab.bottom - h;
										rc.bottom -= h;
										break;
									case 4:						//left
										SetRect(&rc, pTC->m_param[TC_TabWidth], 0, rcTab.right, rcTab.bottom);
										rcTab.right = rc.left;
										break;
									case 5:						//Right
										SetRect(&rc, 0, 0, rcTab.right - pTC->m_param[TC_TabWidth], rcTab.bottom);
										rcTab.left = rc.right;
										break;
								}
							}
							MoveWindow(pTC->m_hwndButton, rcTab.left, rcTab.top,
								rcTab.right - rcTab.left, rcTab.bottom - rcTab.top, true);
							if (nCount && rcTab.bottom) {
								if (nAlign == 4 || nAlign == 5) {
									if (h <= rcTab.bottom) {
										rcTab.bottom = h;
										ShowScrollBar(pTC->m_hwndButton, SB_VERT, FALSE);
										if (pTC->m_si.nTrackPos && pTC->m_param[TE_Flags] & TCS_FIXEDWIDTH) {
											SendMessage(pTC->m_hwnd, TCM_SETITEMSIZE, 0, MAKELPARAM(pTC->m_param[TC_TabWidth], pTC->m_param[TC_TabHeight]));
										}
										pTC->m_si.nTrackPos = 0;
										pTC->m_si.nPos = 0;
									}
									else {
										int h2 = h - rcTab.bottom; 
										if (h2 < pTC->m_si.nPos) {
											pTC->m_si.nPos = h2;
										}
										ShowScrollBar(pTC->m_hwndButton, SB_VERT, TRUE);
										pTC->m_si.fMask = SIF_RANGE | SIF_POS | SIF_PAGE;
										pTC->m_si.nMax = h2;
										pTC->m_si.nTrackPos = rcTab.bottom;
										pTC->m_si.nPage = h2 * rc.bottom / (rc.bottom + h);
										SetScrollInfo(pTC->m_hwndButton, SB_VERT, &pTC->m_si, TRUE);
										if (pTC->m_param[TE_Flags] & TCS_FIXEDWIDTH) {
											SendMessage(pTC->m_hwnd, TCM_SETITEMSIZE, 0, MAKELPARAM(pTC->m_param[TC_TabWidth] - GetSystemMetrics(SM_CXHSCROLL), pTC->m_param[TC_TabHeight]));
										}
									}
								}
								else {
									ShowScrollBar(pTC->m_hwndButton, SB_VERT, FALSE);
									if (pTC->m_si.nTrackPos && pTC->m_param[TE_Flags] & TCS_FIXEDWIDTH) {
										SendMessage(pTC->m_hwnd, TCM_SETITEMSIZE, 0, MAKELPARAM(pTC->m_param[TC_TabWidth], pTC->m_param[TC_TabHeight]));
									}
									pTC->m_si.nTrackPos = 0;
									pTC->m_si.nPos = 0;
								}
								if (teSetRect(pTC->m_hwnd, 0, -pTC->m_si.nPos, rcTab.right, (rcTab.bottom - rcTab.top) + pTC->m_si.nPos)) {
									ArrangeWindow();
								}
								if (pSB) {
									if (pSB->m_param[SB_TreeAlign] == (TE_TV_Use | TE_TV_View)) {
										ArrangeTree(pSB, &rc);//, left, top);TODO
									}
									if (pSB->m_bVisible) {
										teSetParent(pSB->m_pTV->m_hwnd, pSB->m_pTV->GetParentWindow());
										ShowWindow(pSB->m_pTV->m_hwnd, ((pSB->m_param[SB_TreeAlign] & TE_TV_View) ? SW_SHOW : SW_HIDE));
									}
									else {
										pSB->Show(TRUE);
									}
									MoveWindow(pSB->m_hwnd, rc.left, rc.top, rc.right - rc.left, 
										rc.bottom - rc.top, true);
								}
							}
							else {
								MoveWindow(pTC->m_hwndButton, 0, 0, 1, 1, true);
								pSB = pTC->GetShellBrowser(pTC->m_nIndex);
								if (pSB) {
									MoveWindow(pSB->m_hwnd, rc.left, rc.top, rc.right - rc.left, 
										rc.bottom - rc.top, true);
									if (!pSB->m_bVisible) {
										pSB->Show(true);
									}
									if (pSB->m_bVisible) {
										teSetParent(pSB->m_pTV->m_hwnd, g_hwndMain);
										ShowWindow(pSB->m_pTV->m_hwnd, ((pSB->m_param[SB_TreeAlign] & TE_TV_View) ? SW_SHOW : SW_HIDE));
									}
								}
							}
						}
					} catch (...) {}
					pTC->UnLockUpdate();
					pTC->RedrawUpdate();
				}
				else{
					break;
				}
			}
			g_nSize--;
			break;
		case TET_Redraw:
			i = MAX_TC;
			while (--i >= 0) {
				if (g_pTC[i]) {
					g_pTC[i]->RedrawUpdate();
				}
				else {
					break;
				}
			}
			break;
		case TET_Show:
			i = MAX_TC;
			while (--i >= 0) {
				CteTabs *pTC = g_pTC[i];
				if (pTC) {
					if (pTC->m_bVisible) {
						CteShellBrowser *pSB = pTC->GetShellBrowser(pTC->m_nIndex);
						if (pSB && !pSB->m_bEmpty && pSB->m_bVisible) {
							HWND hwndDV = pSB->m_hwnd;
#ifdef _VISTA7
							if (pSB->m_pExplorerBrowser) {
								EnumChildWindows(hwndDV, EnumChildProc, (LPARAM)&hwndDV);
							}
#endif
							HWND hList = FindWindowEx(hwndDV, 0, WC_LISTVIEW, NULL);
							if (hList) {
								ShowWindow(hList, SW_SHOW);
							}
						}
					}
				}
				else {
					break;
				}
			}
			break;
	}
}

LRESULT CALLBACK HookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	CWPSTRUCT *pcwp;

	if (nCode >= 0) {
		if (nCode == HC_ACTION) {
			if (wParam == NULL) {
				pcwp = (CWPSTRUCT *)lParam;
				switch (pcwp->message)
				{
					case WM_SETFOCUS:
						CheckChangeTabSB(pcwp->hwnd);
						break;
					case WM_SHOWWINDOW:
						if (!pcwp->wParam) {
							int i = MAX_TC;
							while (--i >= 0) {
								CteTabs *pTC = g_pTC[i];
								if (pTC) {
									if (pTC->m_bVisible) {
										CteShellBrowser *pSB = pTC->GetShellBrowser(pTC->m_nIndex);
										if (pSB && !pSB->m_bEmpty && pSB->m_bVisible) {
											HWND hwndDV = pSB->m_hwnd;
#ifdef _VISTA7
											if (pSB->m_pExplorerBrowser) {
												EnumChildWindows(hwndDV, EnumChildProc, (LPARAM)&hwndDV);
											}
#endif
											HWND hList = FindWindowEx(hwndDV, 0, WC_LISTVIEW, NULL);
											if (hList == pcwp->hwnd) {
												KillTimer(g_hwndMain, TET_Show);
												SetTimer(g_hwndMain, TET_Show, 500, teTimerProc);
											}
										}
									}
								}
								else {
									break;
								}
							}
						}
					case WM_ACTIVATE:
					case WM_ACTIVATEAPP:
					case WM_KILLFOCUS:
					case WM_NOTIFY:
						if (g_pOnFunc[TE_OnSystemMessage]) {
							IDispatch *pdisp;
							if (ControlFromhwnd(&pdisp, pcwp->hwnd) == S_OK) {
								MSG msg1;
								msg1.hwnd = pcwp->hwnd;
								msg1.message = pcwp->message;
								msg1.wParam = pcwp->wParam;
								msg1.lParam = pcwp->lParam;
								MessageSub(TE_OnSystemMessage, pdisp, &msg1, S_FALSE);
							}
						}
						break;
				}
			}
		}
	}
	return CallNextHookEx(g_hHook, nCode, wParam, lParam); 
}

int APIENTRY _tWinMain(HINSTANCE hInstance,
					HINSTANCE hPrevInstance,
					LPTSTR    lpCmdLine,
					int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	g_nCmdShow = nCmdShow;
	LPWSTR pszPath, pszIndex;
	BSTR bsPath;
	hInst = hInstance;

	::OleInitialize(NULL);
	//Multiple Launch
	g_pCrcTable = new UINT[256];
	for (UINT i = 0; i < 256; i++) {
		UINT c = i;
		for (int j = 0; j < 8; j++) {
			c = (c & 1) ? (0xedb88320 ^ (c >> 1)) : (c >> 1);
		}
		g_pCrcTable[i] = c;
	}
	teGetModuleFileName(NULL, &pszPath);//Executable Path

	for (int i = 0; pszPath[i]; i++) {
		if (pszPath[i] == '\\') {
			pszPath[i] = '/';
		}
		pszPath[i] = towupper(pszPath[i]);
	}
	int nCrc32 = CalcCrc32((BYTE *)pszPath, lstrlen(pszPath) * sizeof(WCHAR), 0);
	g_hMutex = CreateMutex(NULL, FALSE, pszPath);
	delete [] pszPath;

	if (GetLastError() == ERROR_ALREADY_EXISTS) {
		teGetModuleFileName(NULL, &pszPath);//Executable Path
		HWND hwndTE = NULL;
		while (hwndTE = FindWindowEx(NULL, hwndTE, WINDOW_CLASS, NULL)) {
			if (GetWindowLongPtr(hwndTE, GWLP_USERDATA) == nCrc32) {
				BSTR bs = teGetCommandLine();
				COPYDATASTRUCT cd;
				cd.dwData = 0;
				cd.lpData = bs;
				cd.cbData = ::SysStringByteLen(bs) + sizeof(WCHAR);
				DWORD_PTR dwResult;
				LRESULT lResult = SendMessageTimeout(hwndTE, WM_COPYDATA, nCmdShow, LPARAM(&cd), SMTO_ABORTIFHUNG, 1000 * 30, &dwResult);
				::SysFreeString(bs);
				if (lResult && dwResult == S_OK) {
					::CloseHandle(g_hMutex);
					SetForegroundWindow(hwndTE);
					Finalize();
					return FALSE;
				}
			}
		}
	}

	for (int i = MAX_CSIDL; i--;) {
		SHGetFolderLocation(NULL, i, NULL, NULL, &g_pidls[i]);
	}
	//Late Binding
	GetDisplayNameFromPidl(&bsPath, g_pidls[CSIDL_SYSTEM], SHGDN_FORPARSING);

	tePathAppend(&pszPath, bsPath, L"kernel32.dll");
	g_hKernel32 = LoadLibrary(pszPath);
	if (g_hKernel32) {
		lpfnSetDllDirectoryW = (LPFNSetDllDirectoryW)GetProcAddress(g_hKernel32, "SetDllDirectoryW");
		if (lpfnSetDllDirectoryW) {
			lpfnSetDllDirectoryW(L"");
		}
	}
	delete [] pszPath;

	tePathAppend(&pszPath, bsPath, L"shell32.dll");
	g_hShell32 = LoadLibrary(pszPath);
	if (g_hShell32) {
		lpfnSHCreateItemFromIDList = (LPFNSHCreateItemFromIDList)GetProcAddress(g_hShell32, "SHCreateItemFromIDList");
		lpfnSHParseDisplayName = (LPFNSHParseDisplayName)GetProcAddress(g_hShell32, "SHParseDisplayName");
		lpfnSHRunDialog = (LPFNSHRunDialog)GetProcAddress(g_hShell32, MAKEINTRESOURCEA(61));
	}
	delete [] pszPath;

	tePathAppend(&pszPath, bsPath, L"crypt32.dll");
	g_hCrypt32 = LoadLibrary(pszPath);
	if (g_hCrypt32) {
		lpfnCryptBinaryToStringW = (LPFNCryptBinaryToStringW)GetProcAddress(g_hCrypt32, "CryptBinaryToStringW");
	}
	delete [] pszPath;
	::SysFreeString(bsPath);

	Initlize();
	osInfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFO); 
	GetVersionEx(&osInfo);

	InitCommonControls();
	SHGetDesktopFolder(&g_pDesktopFolder);
	MSG msg;
	CoCreateInstance(CLSID_Shell, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_psha));

	/*Get JScript Object
	if SUCCEEDED(CoCreateInstance(CLSID_HTMLDocument, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_pHtmlDoc))) {
		SAFEARRAY *psa = SafeArrayCreateVector(VT_VARIANT, 0, 1);
		if (psa) {
			VARIANT *pv;
			SafeArrayAccessData(psa, (LPVOID *)&pv);
			pv->vt = VT_BSTR;
			pv->bstrVal = ::SysAllocString(L"<body><script type='text/javascript'>document.body.onclick=function(){return {}};document.body.onkeyup=function(){return []}</script></body>");
			SafeArrayUnaccessData(psa);
			g_pHtmlDoc->write(psa);
			VariantClear(pv);
			g_pHtmlDoc->close();
			IHTMLElement *pEle = NULL;
			g_pHtmlDoc->get_body(&pEle);
			if (pEle) {
				if SUCCEEDED(pEle->get_onclick(pv)) {
					if (GetDispatch(pv, &g_pObject)) {
						pv->vt =VT_EMPTY;
					}
					VariantClear(pv);
				}
				if SUCCEEDED(pEle->get_onkeyup(pv)) {
					if (GetDispatch(pv, &g_pArray)) {
						pv->vt =VT_EMPTY;
					}
					VariantClear(pv);
				}
				pEle->Release();
			}
			SafeArrayDestroy(psa);
		}
	}*/

	//Get JScript Object
	LPOLESTR lp = L"function o(){return {}};function a(){return []};";
	if (ParseScript(lp, L"JScript", NULL, NULL, &g_pJS) == S_OK) {
		lp = L"o";
		VARIANT v;
		VariantInit(&v);
		DISPID dispid;
		if (g_pJS->GetIDsOfNames(IID_NULL, &lp, 1, LOCALE_USER_DEFAULT, &dispid) == S_OK) {
			if (Invoke5(g_pJS, dispid, DISPATCH_PROPERTYGET, &v, 0, NULL) == S_OK) {
				g_pObject = v.pdispVal;
				v.vt = VT_EMPTY;
			}
		}
		lp = L"a";
		if (g_pJS->GetIDsOfNames(IID_NULL, &lp, 1, LOCALE_USER_DEFAULT, &dispid) == S_OK) {
			if (Invoke5(g_pJS, dispid, DISPATCH_PROPERTYGET, &v, 0, NULL) == S_OK) {
				g_pArray = v.pdispVal;
				v.vt = VT_EMPTY;
			}
		}
	}

	//Initialize FolderView & TreeView settings
	g_paramFV[SB_Type] = 1;
	g_paramFV[SB_ViewMode] = FVM_DETAILS;
	g_paramFV[SB_FolderFlags] = FWF_SHOWSELALWAYS | FWF_NOWEBVIEW;
	g_paramFV[SB_IconSize] = 0;
	g_paramFV[SB_Options] = EBO_SHOWFRAMES | EBO_ALWAYSNAVIGATE;
	g_paramFV[SB_ViewFlags] = 0;

	g_paramFV[SB_TreeAlign] = TE_TV_Use;
	g_paramFV[SB_TreeWidth] = 200;
	g_paramFV[SB_TreeFlags] = NSTCS_HASEXPANDOS | NSTCS_SHOWSELECTIONALWAYS | NSTCS_BORDER | NSTCS_HASLINES | NSTCS_HORIZONTALSCROLL;
	g_paramFV[SB_EnumFlags] = SHCONTF_FOLDERS;
	g_paramFV[SB_RootStyle] = NSTCRS_VISIBLE | NSTCRS_EXPANDED;

	// Initialize GDI+
	Gdiplus::GdiplusStartup(&g_Token, &g_StartupInput, NULL);
	MyRegisterClass(hInstance);
	// Title & Version
	lstrcpy(g_szTE, L"Tablacus Explorer " _T(STRING(VER_Y)) L"." _T(STRING(VER_M)) L"." _T(STRING(VER_D)) L" Gaku");

	g_hwndMain = CreateWindow(WINDOW_CLASS, g_szTE, WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
	  CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);
	if (!g_hwndMain)
	{
		Finalize();
		return FALSE;
	}
	SetWindowLongPtr(g_hwndMain, GWLP_USERDATA, nCrc32);
	// ClipboardFormat
	IDLISTFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_SHELLIDLIST);
	DROPEFFECTFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
	// Hook
	g_hHook = SetWindowsHookEx(WH_CALLWNDPROC, (HOOKPROC)HookProc, hInst, GetCurrentThreadId());
	g_hMouseHook = SetWindowsHookEx(WH_MOUSE, (HOOKPROC)MouseProc, hInst, GetCurrentThreadId());
	// Create own class
	CoCreateGuid(&g_ClsIdSB);
	CoCreateGuid(&g_ClsIdTV);
	CoCreateGuid(&g_ClsIdTC);
	CoCreateGuid(&g_ClsIdStruct);
	CoCreateGuid(&g_ClsIdContextMenu);
	CoCreateGuid(&g_ClsIdFI);

	// CTE
	g_pTE = new CTE();
	g_pTE->m_param[TE_CmdShow] = nCmdShow;
	teGetModuleFileName(NULL, &pszPath);//Executable Path
	PathRemoveFileSpec(pszPath);
	tePathAppend(&pszIndex, pszPath, L"script\\index.html");
	delete [] pszPath;
	g_pWebBrowser = new CteWebBrowser(g_hwndMain, pszIndex);
	delete [] pszIndex;

	// Init ShellWindows
	IConnectionPoint *pCP = NULL;
	DWORD dwSWCookie;
	long lSWCookie = 0;
	if (SUCCEEDED(CoCreateInstance(CLSID_ShellWindows,
		NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&g_pSW)))) {
		g_pSW->Register(g_pWebBrowser->m_pWebBrowser, (LONG)g_hwndMain, SWC_3RDPARTY, &lSWCookie);
		IConnectionPointContainer *pCPC;
		if SUCCEEDED(g_pSW->QueryInterface(IID_PPV_ARGS(&pCPC))) {
			if SUCCEEDED(pCPC->FindConnectionPoint(DIID_DShellWindowsEvents, &pCP)) {
				IUnknown *punk;
				if SUCCEEDED(g_pTE->QueryInterface(IID_PPV_ARGS(&punk))) {
					pCP->Advise(punk, &dwSWCookie);
					punk->Release();
				}
			}
			pCPC->Release();
		}
	}

	// Main message loop:
	for (;;) {
		try {
			if (!GetMessage(&msg, NULL, 0, 0)) {
				break;
			}
			if (MessageProc(&msg) != S_OK) {
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}
		catch (...) {
		}
	}
	//At the end of processing
	try {
		if (lSWCookie) {
			g_pSW->Revoke(lSWCookie);
		}
	} catch (...) {}
	try {
		if (pCP) {
			pCP->Unadvise(dwSWCookie);
			pCP->Release();
		}
	} catch (...) {}
	try {
		UnhookWindowsHookEx(g_hMouseHook);
	} catch (...) {}
	try {
		UnhookWindowsHookEx(g_hHook);
	} catch (...) {}
	try {
		g_hMenu && DestroyMenu(g_hMenu);
	} catch (...) {}
	Finalize();
	return (int) msg.wParam;
}

VOID ArrangeWindow()
{
	if (g_nSize <= 0 && g_nLockUpdate == 0) {
		g_nSize = 1;
		SetTimer(g_hwndMain, TET_Size, 1, teTimerProc);
		return;
	}
	KillTimer(g_hwndMain, TET_Size2);
	SetTimer(g_hwndMain, TET_Size2, 500, teTimerProc);
}

VOID ParseInvoke(TEInvoke *pInvoke)
{
	LPITEMIDLIST pidl = (LPITEMIDLIST)pInvoke->pResult;
	VariantClear(&pInvoke->pv[pInvoke->cArgs - 1]);
	if (pidl) {
		try {
			FolderItem *pid;
			if (GetFolderItemFromPidl(&pid, pidl)) {
				teSetObject(&pInvoke->pv[pInvoke->cArgs - 1], pid);
				pid->Release();
			}
			::CoTaskMemFree(pidl);
		} catch (...) {}
	}
	Invoke5(pInvoke->pdisp, pInvoke->dispid, DISPATCH_METHOD, NULL, pInvoke->cArgs, pInvoke->pv);
	delete [] pInvoke;
}

VOID CALLBACK teTimerProcParse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);
	ParseInvoke((TEInvoke *)idEvent);
}

VOID CALLBACK teTimerProcBrowse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);
	TEBrowse *pBrowse = (TEBrowse *)idEvent;
	g_pSB[pBrowse->nSB]->BrowseObject(pBrowse->pidl, SBSP_ABSOLUTE | SBSP_SAMEBROWSER);
	::CoTaskMemFree(pBrowse->pidl);
	delete [] pBrowse;
}

static void threadParseDisplayName(void *args)
{
	::OleInitialize(NULL);
	TEInvoke *pInvoke = (TEInvoke *)args;
	pInvoke->pResult = teILCreateFromPath(pInvoke->pv[pInvoke->cArgs - 1].bstrVal);
	SetTimer(g_hwndMain, (UINT_PTR)args, 10, teTimerProcParse);
	::OleUninitialize();
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
	int i;

	switch (message)
	{
		case WM_CLOSE:
			if (CanClose(g_pTE)) {
				DestroyWindow(hWnd);
			}
			return 0;
		case WM_SIZE:
			ArrangeWindow();
			break;
		case WM_DESTROY:
			if (g_pOnFunc[TE_OnSystemMessage]) {
				msg1.hwnd = hWnd;
				msg1.message = message;
				msg1.wParam = wParam;
				msg1.lParam = lParam;
				MessageSub(TE_OnSystemMessage, g_pTE, &msg1, S_FALSE);
			}
			//Close Tab
			i = MAX_TC;
			while (--i >= 0) {
				if (g_pTC[i]) {
					g_pTC[i]->Close(true);
				}
				else {
					break;
				}
			}
			//Close Browser
			g_pWebBrowser->Close();

			PostQuitMessage(0);
			break;
		case WM_CONTEXTMENU:	//System Menu
			MSG msg1;
			msg1.hwnd = hWnd;
			msg1.message = message;
			msg1.wParam = wParam;
			msg1.pt.x = LOWORD(lParam);
			msg1.pt.y = HIWORD(lParam);
			MessageSubPt(TE_OnShowContextMenu, g_pTE, &msg1);
			return DefWindowProc(hWnd, message, wParam, lParam);
			break;
		case TEM_BrowseObject:
			LPITEMIDLIST pidl;
			pidl = (LPITEMIDLIST)lParam;
			g_pSB[wParam]->BrowseObject(pidl, SBSP_ABSOLUTE | SBSP_SAMEBROWSER);
			::CoTaskMemFree(pidl);
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
			msg1.hwnd = hWnd;
			msg1.message = message;
			msg1.wParam = wParam;
			msg1.lParam = lParam;
			BOOL bDone;
			bDone = true;

			if (g_pOnFunc[TE_OnMenuMessage]) {
				HRESULT hr = MessageSub(TE_OnMenuMessage, g_pTE, &msg1, S_FALSE);
				if (hr == S_OK) {
					bDone = false;
				}
				if (message == WM_MENUCHAR && hr & 0xffff0000) {
					return hr;
				}
			}
			if (bDone) {
				HandleMenuMessage(&msg1);
			}
			break;
		//System
		case WM_POWERBROADCAST:
		case WM_SETFOCUS:
		case WM_KILLFOCUS:
		case WM_DROPFILES:
		case WM_DEVICECHANGE:
		case WM_FONTCHANGE:
		case WM_ENABLE:
		case WM_NOTIFY:
		case WM_SETCURSOR:
		case WM_SETTINGCHANGE:
		case WM_SYSCOLORCHANGE:
		case WM_SYSCOMMAND:
		case WM_THEMECHANGED:
		case WM_USERCHANGED:
		case WM_COPYDATA:
			if (g_pOnFunc[TE_OnSystemMessage]) {
				msg1.hwnd = hWnd;
				msg1.message = message;
				msg1.wParam = wParam;
				msg1.lParam = lParam;
				HRESULT hr;
				hr = MessageSub(TE_OnSystemMessage, g_pTE, &msg1, S_OK);
				if (hr != S_OK) {
					return hr;
				}
			}
			return DefWindowProc(hWnd, message, wParam, lParam);
		case WM_COMMAND:
			if (g_pOnFunc[TE_OnCommand]) {
				msg1.hwnd = hWnd;
				msg1.message = message;
				msg1.wParam = wParam;
				msg1.lParam = lParam;
				HRESULT hr;
				hr = MessageSub(TE_OnCommand, g_pTE, &msg1, S_FALSE);
				if (hr == S_OK) {
					return 1;
				}
			}
			return DefWindowProc(hWnd, message, wParam, lParam);
	default:
		if ((message >= WM_APP && message <= 0xffff)) {		
			if (g_pOnFunc[TE_OnAppMessage]) {
				msg1.hwnd = hWnd;
				msg1.message = message;
				msg1.wParam = wParam;
				msg1.lParam = lParam;
				HRESULT hr = MessageSub(TE_OnAppMessage, g_pTE, &msg1, S_FALSE);
				if (hr == S_OK) {
					return hr;
				}
			}
		}
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

// CteShellBrowser

CteShellBrowser::CteShellBrowser(CteTabs *pTabs)
{
	m_cRef = 1;
	m_nCreate = 0;
	m_bEmpty = true;
	m_bInit = false;
	m_bsFilter = NULL;
	VariantInit(&m_vRoot);
	m_vRoot.vt = VT_I4;
	m_vRoot.lVal = 0;
	m_pTV = new CteTreeView();
	VariantInit(&m_Data);
	Init(pTabs, true);
}

void CteShellBrowser::Init(CteTabs *pTabs, BOOL bNew)
{
	m_hwnd = NULL;
	m_pShellView = NULL;
	m_pExplorerBrowser = NULL;
	m_dwEventCookie = 0;
	m_pidl = NULL;
	m_pFolderItem = NULL;
	m_nUnload = 0;
	m_DefProc[0] = NULL;
	m_DefProc[1] = NULL;
	m_pDropTarget = NULL;
	m_pDragItems = NULL;
	m_pDSFV = NULL;
	m_pSF2 = NULL;
	m_nColumns = 0;
	m_nDefultColumns = 0;
	VariantClear(&m_Data);

	for (int i = SB_Count - 1; i >= 0; i--) {
		m_param[i] = g_paramFV[i];
	}

	m_pTabs = pTabs;
	if (bNew) {
		m_ppLog = new FolderItem*[MAX_HISTORY];
	}
	else {
		while (--m_nLogCount >= 0) {
			m_ppLog[m_nLogCount]->Release();
		}
	}
	for (int i = Count_SBFunc; i--;) {
		if (!bNew && m_pOnFunc[i]) {
			m_pOnFunc[i]->Release();
		}
		m_pOnFunc[i] = NULL;
	}
	m_bVisible = false;
	m_nLogCount = 0;
	m_nLogIndex = -1;

	m_pTV->m_pFV = this;
	m_pTV->Init();
	DoFunc(TE_OnCreate, this, E_NOTIMPL);
}

CteShellBrowser::~CteShellBrowser()
{
	for (int i = _countof(m_pOnFunc); i--;) {
		if (m_pOnFunc[i]) {
			try {
				m_pOnFunc[i]->Release();
				m_pOnFunc[i] = NULL;
			}	
			catch (...) {}
		}
	}
	DestroyView(0);
	while (--m_nLogCount >= 0) {
		m_ppLog[m_nLogCount]->Release();
	}
	if (m_bsFilter) {
		::SysFreeString(m_bsFilter);
		m_bsFilter = NULL;
	}
	if (m_pDSFV) {
		m_pDSFV->Release();
	}
	if (m_pSF2) {
		m_pSF2->Release();
	}
	if (m_nColumns) {
		delete [] m_pColumns;
	}
	if (m_nDefultColumns) {
		delete [] m_pDefultColumns;
	}
	if (m_pidl) {
		::CoTaskMemFree(m_pidl);
	}
	if (m_pFolderItem) {
		m_pFolderItem->Release();
	}
	VariantClear(&m_Data);
}

STDMETHODIMP CteShellBrowser::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IOleWindow) || IsEqualIID(riid, IID_IShellBrowser)) {
		*ppvObject = static_cast<IShellBrowser *>(this);
	}
	else if (IsEqualIID(riid, IID_ICommDlgBrowser)) {
		*ppvObject = static_cast<ICommDlgBrowser *>(this);
	}
	else if (IsEqualIID(riid, IID_ICommDlgBrowser2)) {
		*ppvObject = static_cast<ICommDlgBrowser2 *>(this);
	}
	else if (IsEqualIID(riid, IID_IServiceProvider)) {
		*ppvObject = static_cast<IServiceProvider *>(this);
	}
	else if (IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else if (IsEqualIID(riid, IID_IShellFolderViewDual)) {
		*ppvObject = static_cast<IShellFolderViewDual *>(this);
	}
#ifdef _VISTA7
	else if (IsEqualIID(riid, IID_IExplorerBrowserEvents)) {
		*ppvObject = static_cast<IExplorerBrowserEvents *>(this);
	}
	else if (IsEqualIID(riid, IID_IExplorerPaneVisibility)) {
		*ppvObject = static_cast<IExplorerPaneVisibility *>(this);
	}
#endif
#ifdef _2000XP
	else if (IsEqualIID(riid, IID_IShellFolderViewCB)) {
		*ppvObject = static_cast<IShellFolderViewCB *>(this);
	}
#endif
	else if (IsEqualIID(riid, g_ClsIdSB)) {
		*ppvObject = this;
	}
	else if (IsEqualIID(riid, IID_IDropTarget)) {
		*ppvObject = static_cast<IDropTarget *>(this);
	}
	else if (IsEqualIID(riid, IID_IPersist)) {
		*ppvObject = static_cast<IPersist *>(this);
	}
	else if (IsEqualIID(riid, IID_IPersistFolder)) {
		*ppvObject = static_cast<IPersistFolder *>(this);
	}
	else if (IsEqualIID(riid, IID_IPersistFolder2)) {
		*ppvObject = static_cast<IPersistFolder2 *>(this);
	}
	else if (IsEqualIID(riid, IID_IShellFolder)) {
		if (m_pSF2) {
#ifdef _2000XP
			*ppvObject = static_cast<IShellFolder *>(this);
#else
			return m_pSF2->QueryInterface(riid, ppvObject);
#endif
		}
		else {
			HRESULT hr = E_NOINTERFACE;
#ifdef _VISTA7
			if (m_pShellView) {
				IFolderView2 *pFV2;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
					hr = pFV2->GetFolder(riid, ppvObject);
					pFV2->Release();
				}
			}
#endif
			return hr;
		}
	}
	else if (IsEqualIID(riid, IID_IShellFolder2)) {
		if (m_pSF2) {
#ifdef _2000XP
			*ppvObject = static_cast<IShellFolder2 *>(this);
#else
			return m_pSF2->QueryInterface(riid, ppvObject);
#endif
		}
		else {
			HRESULT hr = E_NOINTERFACE;
#ifdef _VISTA7
			if (m_pShellView) {
				IFolderView2 *pFV2;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
					hr = pFV2->GetFolder(riid, ppvObject);
					pFV2->Release();
				}
			}
#endif
			return hr;
		}
	}
	else if (m_pShellView && IsEqualIID(riid, IID_IShellView)) {
		return m_pShellView->QueryInterface(riid, ppvObject);
	}
	else {
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

STDMETHODIMP CteShellBrowser::GetWindow(HWND *phwnd)
{
	*phwnd = m_pTabs->m_hwndStatic;
	return S_OK;
}

STDMETHODIMP CteShellBrowser::ContextSensitiveHelp(BOOL fEnterMode)
{
	return E_NOTIMPL;
}

void CteShellBrowser::InitializeMenuItem(HMENU hmenu, LPTSTR lpszItemName, int nId, HMENU hmenuSub)
{
	MENUITEMINFO mii;
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask  = MIIM_ID | MIIM_STRING | MIIM_SUBMENU;
	mii.wID    = nId;
	mii.dwTypeData = lpszItemName;
	mii.hSubMenu = hmenuSub;
	InsertMenuItem(hmenu, nId, FALSE, &mii);
}

STDMETHODIMP CteShellBrowser::InsertMenusSB(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths)
{
	if (!g_hMenu) {
		g_hMenu = LoadMenu(g_hShell32, MAKEINTRESOURCE(216));
		if (!g_hMenu) {
			method pull_downs[] = {
				{ FCIDM_MENU_FILE, L"&File" },
				{ FCIDM_MENU_EDIT, L"&Edit" },
				{ FCIDM_MENU_VIEW, L"&View" },
				{ FCIDM_MENU_TOOLS, L"&Tools" },
				{ FCIDM_MENU_HELP, L"&Help" },
			};
			for (size_t i = 0; i < 5; i++)
			{
				InitializeMenuItem(hmenuShared, pull_downs[i].name, pull_downs[i].id, CreatePopupMenu());
				lpMenuWidths->width[i] = 0;
			}
			lpMenuWidths->width[0] = 2; // FILE: File, Edit
			lpMenuWidths->width[2] = 2; // VIEW: View, Tools
			lpMenuWidths->width[4] = 1; // WINDOW: Help
		}
	}
	return S_OK;
}

STDMETHODIMP CteShellBrowser::SetMenuSB(HMENU hmenuShared, HOLEMENU holemenuRes, HWND hwndActiveObject)
{
	if (!g_hMenu) {
		if (hwndActiveObject == m_hwnd) {
			g_hMenu = CreatePopupMenu();
			teCopyMenu(g_hMenu, hmenuShared, MF_ENABLED);
			return S_OK;
		}
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::RemoveMenusSB(HMENU hmenuShared)
{
	int nCount = GetMenuItemCount(hmenuShared);
	while (--nCount >= 0) {
		DeleteMenu(hmenuShared, nCount, MF_BYPOSITION);
	}
	return S_OK;
}

STDMETHODIMP CteShellBrowser::SetStatusTextSB(LPCWSTR lpszStatusText)
{
	return DoStatusText(this, lpszStatusText, 0);
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
	CteShellBrowser *pSB = NULL;
	DWORD param[SB_Count];
	for (int i = SB_Count - 1; i >= 0; i--) {
		param[i] = m_param[i];
	}
	param[SB_DoFunc] = 1;
	FolderItem *pid;
	GetFolderItemFromPidl(&pid, const_cast<LPITEMIDLIST>(pidl));
	return Navigate3(pid, wFlags, param, &pSB, NULL);
}

HRESULT CteShellBrowser::Navigate3(FolderItem *pFolderItem, UINT wFlags, DWORD *param, CteShellBrowser **ppSB, FolderItems *pFolderItems)
{
	HRESULT hr = E_FAIL;
	m_pTabs->LockUpdate();
	try {
		if (!(wFlags & SBSP_NEWBROWSER || m_bEmpty)) {
			hr = Navigate2(pFolderItem, wFlags, param, pFolderItems, m_pidl);
			if (hr == E_ACCESSDENIED) {
				wFlags |= SBSP_NEWBROWSER;
			}
		}
		if (wFlags & SBSP_NEWBROWSER || m_bEmpty) {
			CteShellBrowser *pSB = this;
			LPITEMIDLIST pidl = m_pidl;
			BOOL bNew = !m_bEmpty && (wFlags & SBSP_NEWBROWSER);
			if (bNew) {
				pSB = GetNewShellBrowser(m_pTabs);
				*ppSB = pSB;
				VariantCopy(&pSB->m_vRoot, &m_vRoot);
				pidl = ::ILClone(m_pidl);
			}
			for (int i = SB_Count; i--;) {
				pSB->m_param[i] = param[i];
			}
			pSB->m_bEmpty = false;
			hr = pSB->Navigate2(pFolderItem, wFlags, param, pFolderItems, pidl);
			if (bNew && hr == S_OK && (wFlags & SBSP_ACTIVATE_NOFOCUS) == 0) {
				TabCtrl_SetCurSel(m_pTabs->m_hwnd, m_pTabs->m_nIndex + 1);
				m_pTabs->m_nIndex = TabCtrl_GetCurSel(m_pTabs->m_hwnd);
				if (this != pSB) {
					Show(true);
					pSB->Show(false);
				}
			}
			if (hr == E_ACCESSDENIED) {
				pSB->Close(true);
			}
		}
	} catch (...) {
		hr = E_FAIL;
	}
	m_pTabs->UnLockUpdate();
	pFolderItem && pFolderItem->Release();
	return hr;
}

BOOL CteShellBrowser::Navigate1(FolderItem *pFolderItem, UINT wFlags, FolderItems *pFolderItems, LPITEMIDLIST pidlPrevius)
{
	CteFolderItem *pid = NULL;
	if (pFolderItem) {
		if SUCCEEDED(pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid)) {
			if (!pid->m_pidl && pid->m_v.vt == VT_BSTR) {
				if (PathIsNetworkPath(pid->m_v.bstrVal)) {
					TEInvoke *pInvoke = new TEInvoke[1];
					QueryInterface(IID_PPV_ARGS(&pInvoke->pdisp));
					pInvoke->dispid = 0x20000000;//Navigate2Ex
					pInvoke->cArgs = 4;
					pInvoke->pv = GetNewVARIANT(4);
					::VariantCopy(&pInvoke->pv[3], &pid->m_v); 
					pInvoke->pv[2].lVal = wFlags;
					pInvoke->pv[2].vt = VT_I4;
					teSetObject(&pInvoke->pv[1], pFolderItems);
					FolderItem *pFI;
					GetFolderItemFromPidl(&pFI, pidlPrevius);
					teSetObject(&pInvoke->pv[0], pFI);
					pFI && pFI->Release();
					_beginthread(&threadParseDisplayName, 0, pInvoke);
					pid->Release();
					return TRUE;
				}
			}
			pid->Release();
		}
	}
	return FALSE;
}

HRESULT CteShellBrowser::Navigate2(FolderItem *pFolderItem, UINT wFlags, DWORD *param, FolderItems *pFolderItems, LPITEMIDLIST pidlPrevius)
{
	RECT rc;
	LPITEMIDLIST pidl;
	IFolderView2 *pFV2;
	FolderItem *pFolderItem1 = NULL;
	LPITEMIDLIST pidlFocus = NULL;

	HRESULT hr = S_FALSE;
	SetRedraw(FALSE);
	if (m_bInit) {
		return E_FAIL;
	}
	for (int i = SB_Count - 1; i >= 0; i--) {
		m_param[i] = param[i];
		g_paramFV[i] = param[i];
	}
	if (!GetAbsPidl(&pidl, &pFolderItem1, pFolderItem, wFlags, pFolderItems, pidlPrevius)) {
		if (pidlPrevius && m_pidl != pidlPrevius) {
			::CoTaskMemFree(pidlPrevius);
		}
		return S_OK;
	}
	if (pidlPrevius && wFlags & SBSP_PARENT) {
		pidlFocus = ::ILFindLastID(pidlPrevius);
	}
	if (m_pidl && m_pidl != pidlPrevius) {
		::CoTaskMemFree(m_pidl);	
	}
	m_pidl = pidl;
	FolderItem *pPrevius = m_pFolderItem;
	m_pFolderItem = pFolderItem1;
	if (g_nLockUpdate || !m_pTabs->m_bVisible) {
		m_nUnload = 1;
		//History / Management
		SetHistory(pFolderItems, wFlags);
		if (pidlPrevius) {
			::CoTaskMemFree(pidlPrevius);
		}
		pPrevius && pPrevius->Release();
		return S_OK;
	}

	//Previous display
	IShellView *pPreviusView = m_pShellView;
	if (param[SB_DoFunc]) {
		if (pPreviusView) {
#ifdef _VISTA7
			if SUCCEEDED(pPreviusView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
				pFV2->GetViewModeAndIconSize((FOLDERVIEWMODE *)&m_param[SB_ViewMode], (int *)&m_param[SB_IconSize]);
				pFV2->Release();
			}
			else
#endif
			{
				FOLDERSETTINGS fs;
				pPreviusView->GetCurrentInfo(&fs);
				m_param[SB_ViewMode] = fs.ViewMode;
			}
#ifdef _VISTA7
			if (m_pExplorerBrowser) {
				m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
			}
#endif
		}
	}
	EXPLORER_BROWSER_OPTIONS dwflag;
	dwflag = static_cast<EXPLORER_BROWSER_OPTIONS>(m_param[SB_Options]);

	m_bIconSize = FALSE;
	if (param[SB_DoFunc]) {
		if (g_pOnFunc[TE_OnBeforeNavigate]) {
			VARIANT vResult;
			VariantInit(&vResult);

			VARIANTARG *pv;
			CteMemory *pMem;
			pv = GetNewVARIANT(4);
			teSetObject(&pv[3], this);

			pMem = new CteMemory(4 * sizeof(int), (char *)&m_param[SB_ViewMode], VT_NULL, 1, L"FOLDERSETTINGS");
			pMem->QueryInterface(IID_PPV_ARGS(&pv[2].punkVal));
			teSetObject(&pv[2], pMem);
			pMem->Release();

			pv[1].vt = VT_I4;
			pv[1].lVal = wFlags;
			if (pPreviusView && pPrevius) {
				teSetObject(&pv[0], pPrevius);
			}
			m_bInit = true;
			try {
				Invoke4(g_pOnFunc[TE_OnBeforeNavigate], &vResult, 4, pv);
			}
			catch (...) {}
			m_bInit = false;
			HRESULT hr = GetIntFromVariantClear(&vResult);
			if (hr != S_OK && m_nUnload != 2) {
				::CoTaskMemFree(m_pidl);
				m_pidl = pidlPrevius;
				m_pFolderItem && m_pFolderItem->Release();
				m_pFolderItem = pPrevius;
				return hr;
			}
		}
	}

	SetRectEmpty(&rc);

	//History / Management
	SetHistory(pFolderItems, wFlags);
	DestroyView(m_param[SB_Type]);
	Show(false);
	BOOL bExplorerBrowser = m_param[SB_Type] == 2;
#ifdef _2000XP
	if (bExplorerBrowser && osInfo.dwMajorVersion <= 5) {
		m_param[SB_Type] = 1;
	}
#endif
	//ShellBrowser
	if (m_param[SB_Type] == 1) {
		IShellFolder *pShellFolder = NULL;
		if (m_nColumns) {
			delete [] m_pColumns;
			m_nColumns = 0;
		}
		int nDog = 5;
		do {
			if (GetShellFolder(&pShellFolder, m_pidl)) {
				if SUCCEEDED(pShellFolder->QueryInterface(IID_PPV_ARGS(&m_pSF2))) {
					hr = CreateViewWindowEx(pPreviusView);
				}
				pShellFolder->Release();
			}
			if (hr != S_OK) {
#ifdef _VISTA7
				if (hr == S_FALSE) {
					if (osInfo.dwMajorVersion >= 6) {
						bExplorerBrowser = true;//Use ExplorerBrowser
						m_param[SB_Options] |= EBO_SHOWFRAMES;
						break;
					}
				}
#endif
				Error(&nDog);
			}
		} while (hr != S_OK && nDog--);
	}
#ifdef _VISTA7
	//ExplorerBrowser
	if (bExplorerBrowser) {
		if (m_pExplorerBrowser) {
			DestroyView(CTRL_SB);
		}
		if (SUCCEEDED(CoCreateInstance(CLSID_ExplorerBrowser,
		NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pExplorerBrowser)))) {
			if (SUCCEEDED(m_pExplorerBrowser->Initialize(g_pTabs->m_hwndStatic, &rc, (FOLDERSETTINGS *) &m_param[SB_ViewMode]))) {
				m_pExplorerBrowser->SetOptions(static_cast<EXPLORER_BROWSER_OPTIONS>(m_param[SB_Options]));
				m_pExplorerBrowser->Advise(this, &m_dwEventCookie);
				IFolderViewOptions *pOptions;
				if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
					pOptions->SetFolderViewOptions(FVO_VISTALAYOUT, FVO_VISTALAYOUT);
					pOptions->Release();
				}
				int nDog = 5;
				do {
					hr = m_pExplorerBrowser->BrowseToIDList(m_pidl, SBSP_ABSOLUTE);
					if (hr == S_OK) {
						m_pExplorerBrowser->GetCurrentView(IID_PPV_ARGS(&m_pShellView));
						if (m_pShellView) {
							GetDefaultColumns();
						}
						IOleWindow *pOleWindow;
						if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
							pOleWindow->GetWindow(&m_hwnd);
							pOleWindow->Release();
						}
						if (m_param[SB_FolderFlags] & FWF_NOCLIENTEDGE) {
							SetWindowLong(m_hwnd, GWL_STYLE, GetWindowLong(m_hwnd, GWL_STYLE) & ~WS_BORDER);
						}
						DoStatusText(this, NULL, 0);
					}
					else {
						Error(&nDog);
					}
				} while (hr != S_OK && nDog--);
			}
		}
	}
#endif
	if (hr != S_OK) {
		if (m_pShellView) {
			m_pShellView->Release();
		}
		m_pShellView = pPreviusView;
		pPreviusView = NULL;

		::CoTaskMemFree(m_pidl);
		m_pidl = pidlPrevius;
		pidlPrevius = NULL;
		m_pFolderItem = pPrevius;
		m_pFolderItem && m_pFolderItem->Release();
		pPrevius = NULL;
	}

	if (pPreviusView) {
		pPreviusView->DestroyViewWindow();
		pPreviusView->Release();
		pPreviusView = NULL;
	}
	else {
		IFolderView2 *pFV2;
		if (m_pShellView) {
#ifdef _VISTA7
			if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
				IShellFolder2 *pSF2;
				if SUCCEEDED(pFV2->GetFolder(IID_PPV_ARGS(&pSF2))) {
					SORTCOLUMN col;
					if SUCCEEDED(pSF2->MapColumnToSCID(0, &col.propkey)) {
						col.direction = 1;
						pFV2->SetSortColumns(&col, 1);
					}
					pSF2->Release();
				}
				pFV2->Release();
			}
			else
#endif
			{
#ifdef _2000XP
				LPARAM Sort;
				IShellFolderView *pSFV;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
					if SUCCEEDED(pSFV->GetArrangeParam(&Sort)) {
						if (Sort) {
							pSFV->Rearrange(0);
						}
					}
					pSFV->Release();
				}
#endif
			}
		}
	}
	if (pidlFocus) {
		m_pShellView->SelectItem(pidlFocus, SVSI_FOCUSED | SVSI_ENSUREVISIBLE);
	}
	if (pidlPrevius) {
		::CoTaskMemFree(pidlPrevius);
	}
	pPrevius && pPrevius->Release();

	m_nOpenedType = m_param[SB_Type];
	if (m_pTabs && GetTabIndex() == m_pTabs->m_nIndex) {
		Show(true);
	}
	if (m_pExplorerBrowser) {
		IUnknown_SetSite(m_pExplorerBrowser, static_cast<IServiceProvider *>(this));
		m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
	}
	else {
		if (g_bAvgWidth) {
			HWND hList = FindWindowEx(m_hwnd, 0, WC_LISTVIEW, NULL);
			if (hList) {
				HWND hHeader = ListView_GetHeader(hList);
				if (hHeader) {
					HD_ITEM hdi;
					::ZeroMemory(&hdi, sizeof(HD_ITEM));
					hdi.mask = HDI_TEXT | HDI_WIDTH;
					hdi.cchTextMax = MAX_COLUMN_NAME_LEN;
					WCHAR szText[MAX_COLUMN_NAME_LEN];
					hdi.pszText = (LPWSTR)&szText;
					Header_GetItem(hHeader, 0, &hdi);
					IShellFolder2 *pSF2;
					if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pSF2))) {								
						SHELLDETAILS sd;
						for (UINT i = 0; pSF2->GetDetailsOf(NULL, i, &sd) == S_OK; i++) {
							BSTR bs;
							if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs)) {
								BOOL bSame = lstrcmpi(bs, hdi.pszText) == 0;
								SysFreeString(bs);
								if (bSame) {
									if (sd.cxChar) {
										g_nAvgWidth = hdi.cxy / sd.cxChar;
									}
									break;
								}
							}
						}
						g_bAvgWidth = false;
						pSF2->Release();
					}
				}
			}
		}
		OnViewCreated(NULL);
		ArrangeWindow();
	}
	return S_OK;
}

BOOL CteShellBrowser::GetAbsPidl(LPITEMIDLIST *pidlOut, FolderItem **ppid, FolderItem *pid, UINT wFlags, FolderItems *pFolderItems, LPITEMIDLIST pidlPrevius)
{
	if (wFlags & SBSP_PARENT) {
		if (!pidlPrevius || ILIsEmpty(pidlPrevius)) {
			*pidlOut = ::ILClone(g_pidls[CSIDL_DESKTOP]);
		}
		else {
			*pidlOut = ::ILClone(pidlPrevius);
			::ILRemoveLastID(*pidlOut);
		}
		GetFolderItemFromPidl(ppid, *pidlOut);
		return TRUE;
	}
	if (wFlags & SBSP_NAVIGATEBACK) {
		if (m_nLogIndex > 0 && GetIDListFromObjectEx(m_ppLog[--m_nLogIndex], pidlOut)) {
			m_ppLog[m_nLogIndex]->QueryInterface(IID_PPV_ARGS(ppid));
			if (Navigate1(*ppid, wFlags, pFolderItems, pidlPrevius)) {
				(*ppid)->Release();
				return FALSE;
			}
			return TRUE;
		}
		return FALSE;
	}
	if (wFlags & SBSP_NAVIGATEFORWARD) {
		if (m_nLogIndex < m_nLogCount - 1 && GetIDListFromObjectEx(m_ppLog[++m_nLogIndex], pidlOut)) {
			m_ppLog[m_nLogIndex]->QueryInterface(IID_PPV_ARGS(ppid));
			if (Navigate1(*ppid, wFlags, pFolderItems, pidlPrevius)) {
				(*ppid)->Release();
				return FALSE;
			}
			return TRUE;
		}
		return FALSE;
	}
	if (Navigate1(pid, wFlags, pFolderItems, pidlPrevius)) {
		return FALSE;
	}
	LPITEMIDLIST pidl = NULL;
	if (pid) {
		GetIDListFromObject(pid, &pidl);
	}
	if (wFlags & SBSP_RELATIVE) {
		if (pidlPrevius && pidl) {
			*pidlOut = ILCombine(pidlPrevius, pidl);
		}
		else {
			*pidlOut = ILClone(pidlPrevius);
		}
		GetFolderItemFromPidl(ppid, *pidlOut);
		::CoTaskMemFree(pidl);
		return TRUE;
	}
	if (pid) {
		pid->QueryInterface(IID_PPV_ARGS(ppid));
		if (pidl) {
			*pidlOut = pidl;
			return TRUE;
		}
	}
	if (*ppid) {
		GetFolderItemFromPidl(ppid, *pidlOut);
		return TRUE;
	}
	return FALSE;
}

VOID CteShellBrowser::HookDragDrop(int nMode)
{
	if (m_hwnd && nMode & 1) {
#ifdef _VISTA7
		HWND hwndDV = m_hwnd;
		if (m_pExplorerBrowser) {
			EnumChildWindows(m_hwnd, EnumChildProc, (LPARAM)&hwndDV);
		}
#endif
		if (!m_pDropTarget) {
			HWND hwnd = FindWindowEx(hwndDV, 0, WC_LISTVIEW, NULL);
			teRegisterDragDrop(hwnd, this, &m_pDropTarget);
		}
	}
	if (nMode & 2 && m_pTV && m_pTV->m_hwnd && !m_pTV->m_pDropTarget) {
		HWND hwnd = FindTreeWindow(m_pTV->m_hwnd);
		if (hwnd) {
			teRegisterDragDrop(hwnd, static_cast<IDropTarget *>(m_pTV), &m_pTV->m_pDropTarget);
		}
	}
}

VOID CteShellBrowser::Error(int *pnDog)
{
	ILRemoveLastID(m_pidl);
	WCHAR szPath[MAX_PATH];//OK
	if (SHGetPathFromIDList(m_pidl, szPath)) {
		if (PathIsNetworkPath(szPath)) {
			*pnDog = 1;
		}
	}
	else {
		*pnDog = 1;
	}
	if (*pnDog == 1) {
		::CoTaskMemFree(m_pidl);
		m_pidl = ::ILClone(g_pidls[CSIDL_DESKTOP]);
	}
	m_pFolderItem && m_pFolderItem->Release();
	GetFolderItemFromPidl(&m_pFolderItem, m_pidl);
}

VOID CteShellBrowser::Refresh()
{
	if (m_pShellView) {
		if (m_bVisible) {
#ifdef _VISTA7
			if (m_pExplorerBrowser) {
				m_pShellView->UIActivate(SVUIA_ACTIVATE_NOFOCUS);
			}
#endif
			m_pShellView->Refresh();
#ifdef _VISTA7
			if (m_pExplorerBrowser) {
				SetRedraw(TRUE);
				m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
			}
#endif
			if (m_nUnload == 3) {
				m_nUnload = 0;
			}
		}
		else if (m_nUnload == 0) {
			m_nUnload = 3;
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
		int nLogIndex = 0;
		VARIANT v;
		if (Invoke5(pFolderItems, DISPID_TE_INDEX, DISPATCH_PROPERTYGET, &v, 0, NULL) == S_OK) {
			nLogIndex = GetIntFromVariantClear(&v);
		}

		while (--m_nLogCount >= 0) {
			m_ppLog[m_nLogCount]->Release();
		}
		m_nLogCount = 0;
		long nCount = 0;
		pFolderItems->get_Count(&nCount);
		if (nCount > MAX_HISTORY) {
			nCount = MAX_HISTORY;
		}
		VariantInit(&v);
		v.vt = VT_I4;
		while (--nCount >= 0) {
			v.lVal = nCount;
			pFolderItems->Item(v, &m_ppLog[m_nLogCount++]);
		}
		m_nLogIndex = m_nLogCount - nLogIndex - 1;
		pFolderItems->Release();
		pFolderItems = NULL;
	}
	else if ((wFlags & (SBSP_NAVIGATEBACK | SBSP_NAVIGATEFORWARD | SBSP_WRITENOHISTORY)) == 0) {
		LPITEMIDLIST pidl = NULL;
		if (!(m_nLogCount > 0 && GetIDListFromObjectEx(m_ppLog[m_nLogIndex], &pidl) && ::ILIsEqual(m_pidl, pidl))) {
			while (m_nLogIndex < m_nLogCount - 1 && m_nLogCount > 0) {
				m_ppLog[--m_nLogCount]->Release();
			}
			if (m_nLogCount >= MAX_HISTORY) {
				m_nLogCount = MAX_HISTORY - 1;
				m_ppLog[0]->Release();
				for (int i = 0; i < MAX_HISTORY - 1; i++) {
					m_ppLog[i] = m_ppLog[i + 1];
				}
			}
			m_nLogIndex = m_nLogCount;
			if (!m_pFolderItem && m_pidl) {
				GetFolderItemFromPidl(&m_pFolderItem, m_pidl);
			}
			m_pFolderItem && m_pFolderItem->QueryInterface(IID_PPV_ARGS(&m_ppLog[m_nLogCount++]));
		}
		if (pidl) {
			::CoTaskMemFree(pidl);
		}
	}
}

STDMETHODIMP CteShellBrowser::GetViewStateStream(DWORD grfMode, IStream **ppStrm)
{
#ifdef _VISTA7
	if (m_pExplorerBrowser) {
		if (m_DefProc[1]) {
			HWND hwnd = 0;
			EnumChildWindows(m_hwnd, EnumChildProc, (LPARAM)&hwnd);
			if (hwnd) {
				LONG_PTR wp = GetWindowLongPtr(hwnd, GWLP_WNDPROC);
				if (wp == (LONG_PTR)TELVProc2) {
					SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LONG_PTR)m_DefProc[1]);
				}
			}
			m_DefProc[1] = NULL;
		}
	}
#endif
	return E_NOTIMPL;
}


STDMETHODIMP CteShellBrowser::GetControlWindow(UINT id, HWND *lphwnd)
{
	if (id == FCW_STATUS) {
		if (lphwnd) {
			HWND hwnd = m_hwnd;
#ifdef _VISTA7
			if (m_pExplorerBrowser) {
				EnumChildWindows(m_hwnd, EnumChildProc, (LPARAM)&hwnd);
			}
#endif
			if (hwnd) {
				*lphwnd = hwnd;
				return S_OK;
			}
		}
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::SendControlMsg(UINT id, UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT *pret)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::QueryActiveShellView(IShellView **ppshv)
{
	if (m_pShellView) {
		return m_pShellView->QueryInterface(IID_PPV_ARGS(ppshv));
	}
	return E_FAIL;
}

STDMETHODIMP CteShellBrowser::OnViewWindowActive(IShellView *ppshv)
{
	if (m_pTabs) {
		CheckChangeTabTC(m_pTabs->m_hwnd, false);
	}
	BSTR bsPath;
	if SUCCEEDED(GetDisplayNameFromPidl(&bsPath, m_pidl, SHGDN_FORPARSING)) {
		SetCurrentDirectory(bsPath);
		::SysFreeString(bsPath);
	}
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
	if (uChange == CDBOSC_SELCHANGE) {
		return DoFunc(TE_OnSelectionChanged, this, S_OK);
	}
	return S_OK;
}

STDMETHODIMP CteShellBrowser::IncludeObject(IShellView *ppshv, LPCITEMIDLIST pidl)
{
	if (m_pOnFunc[SB_OnIncludeObject]) {
		VARIANTARG *pv;
		VARIANT vResult;
		VariantInit(&vResult);
		pv = GetNewVARIANT(2);
		teSetObject(&pv[1], this);
		LPITEMIDLIST FullID;
		FullID = ILCombine(m_pidl, pidl);
		FolderItem *pFolderItem;
		if (GetFolderItemFromPidl(&pFolderItem, FullID)) {
			teSetObject(&pv[0], pFolderItem);
			pFolderItem->Release();
		}
		::CoTaskMemFree(FullID);
		Invoke4(m_pOnFunc[SB_OnIncludeObject], &vResult, 2, pv);
		return GetIntFromVariantClear(&vResult);
	}
	if (m_bsFilter) {
		HRESULT hr = S_OK;
		IShellFolder *pSF;
		if (GetShellFolder(&pSF, m_pidl)) {
			STRRET strret;
			if SUCCEEDED(pSF->GetDisplayNameOf(pidl, SHGDN_NORMAL, &strret)) {
				BSTR bsFile;
				if SUCCEEDED(StrRetToBSTR(&strret, pidl, &bsFile)) {
					if (!PathMatchSpec(bsFile, m_bsFilter)) {
						hr = S_FALSE;
					}
					::SysFreeString(bsFile);
				}
				pSF->Release();
			}
		}
		return hr;
	}
	return S_OK;
}

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
	*pdwFlags = m_param[SB_ViewFlags] & (~CDB2GVF_NOINCLUDEITEM);
	return S_OK;
}
//IServiceProvider
STDMETHODIMP CteShellBrowser::QueryService(REFGUID guidService, REFIID riid, void **ppv)
{
	return QueryInterface(riid, ppv);
}

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
	HRESULT hr = teGetDispId(methodSB, sizeof(methodSB), g_maps[MAP_SB], *rgszNames, rgDispId, true); 
	if (hr != S_OK && m_pDSFV) {
		return m_pDSFV->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
	}
	return hr;
}

VOID AddColumnData(LPWSTR pszColumns, LPWSTR pszName, int nWidth)
{
	lstrcat(pszColumns, L" ");
	WCHAR szName[MAX_COLUMN_NAME_LEN + 2];
	lstrcpyn(szName, pszName, MAX_COLUMN_NAME_LEN);
	PathQuoteSpaces(szName);
	lstrcat(pszColumns, szName);
	lstrcat(pszColumns, L" ");
	VARIANT v, vs;
	vs.vt = VT_I4;
	vs.lVal = nWidth;
	teVariantChangeType(&v, &vs, VT_BSTR);
	lstrcat(pszColumns, v.bstrVal);
	VariantClear(&v);
}

BSTR CteShellBrowser::GetColumnsStr()
{
	WCHAR szColumns[8192];
	szColumns[0] = 0;
	szColumns[1] = 0;
#ifdef _VISTA7
	IColumnManager *pColumnManager;
	if (m_pShellView && SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager)))) {
		UINT uCount;
		if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_VISIBLE, &uCount)) {
			if (uCount) {
				PROPERTYKEY *propKey = new PROPERTYKEY[uCount];
				if SUCCEEDED(pColumnManager->GetColumns(CM_ENUM_VISIBLE, propKey, uCount)) {
					CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_WIDTH | CM_MASK_NAME };
					for (UINT i = 0; i < uCount; i++) {
						if SUCCEEDED(pColumnManager->GetColumnInfo(propKey[i], &cmci)) {
							AddColumnData(szColumns, cmci.wszName, cmci.uWidth);
						}
					}
				}
				delete [] propKey;
			}
		}
		pColumnManager->Release();
	}
	else
#endif
	{
#ifdef _2000XP
		HWND hList = FindWindowEx(m_hwnd, 0, WC_LISTVIEW, NULL);
		if (hList) {
			HWND hHeader = ListView_GetHeader(hList);
			if (hHeader) {
				UINT nHeader = Header_GetItemCount(hHeader);
				if (nHeader) {
					int *piOrderArray = new int[nHeader];
					ListView_GetColumnOrderArray(hList, nHeader, piOrderArray);
					HD_ITEM hdi;
					::ZeroMemory(&hdi, sizeof(HD_ITEM));
					hdi.mask = HDI_TEXT | HDI_WIDTH;
					hdi.cchTextMax = MAX_COLUMN_NAME_LEN;
					WCHAR szText[MAX_COLUMN_NAME_LEN];
					hdi.pszText = (LPWSTR)&szText;
					for (UINT i = 0; i < nHeader; i++) {
						Header_GetItem(hHeader, piOrderArray[i], &hdi);
						AddColumnData(szColumns, szText, hdi.cxy);
					}
					delete [] piOrderArray;
				}
			}
			else {
				IShellFolder2 *pSF2;
				if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pSF2))) {								
					SHELLDETAILS sd;
					SHCOLSTATEF csFlags;
					for (UINT i = 0; pSF2->GetDetailsOf(NULL, i, &sd) == S_OK && i < 8192; i++) {
						if (pSF2->GetDefaultColumnState(i, &csFlags) == S_OK) {
							if (csFlags & SHCOLSTATE_ONBYDEFAULT) {
								BSTR bs;
								if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs)) {
									AddColumnData(szColumns, bs, sd.cxChar * g_nAvgWidth);
									::SysFreeString(bs);
								}
							}
						}
					}
					pSF2->Release();
				}
			}
		}
#endif
	}
	return SysAllocString(&szColumns[1]);
}

VOID CteShellBrowser::GetDefaultColumns()
{
#ifdef _VISTA7
	if (m_nDefultColumns) {
		m_nDefultColumns = 0;
		delete [] m_pDefultColumns;
	}
	IColumnManager *pColumnManager;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager))) {
		if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_VISIBLE, &m_nDefultColumns)) {
			m_pDefultColumns = new PROPERTYKEY[m_nDefultColumns];
			pColumnManager->GetColumns(CM_ENUM_VISIBLE, m_pDefultColumns, m_nDefultColumns);
		}
		pColumnManager->Release();
	}
#endif
}

STDMETHODIMP CteShellBrowser::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	int nIndex;
	TC_ITEM tcItem;

	if ((m_bEmpty || m_nUnload == 1) && dispIdMember >= 0x10000100 && dispIdMember) {
		return S_OK;
	}

	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	switch(dispIdMember) {
		//Data
		case 0x10000001:	
			if (nArg >= 0) {
				VariantClear(&m_Data);
				VariantCopy(&m_Data, &pDispParams->rgvarg[nArg]);
			}
			if (pVarResult) {
				VariantCopy(pVarResult, &m_Data);
			}
			return S_OK;
		//hwnd
		case 0x10000002:
			SetVariantLL(pVarResult, (LONGLONG)m_hwnd);
			return S_OK;
		//Type
		case 0x10000003:
			if (nArg >= 0) {
				DWORD dwType = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (dwType == CTRL_SB || dwType == CTRL_EB) {
					if (m_param[SB_Type] != dwType) {
						m_param[SB_Type] = dwType;
						if (!m_bEmpty) {
							BrowseObject(NULL, SBSP_RELATIVE | SBSP_SAMEBROWSER);
						}
					}
				}
			}
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = m_param[SB_Type];
			}
			return S_OK;

		//Navigate
		case 0x10000004:
		//History
		case 0x1000000B:
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
				for (int i = SB_Count - 1; i >= 0; i--) {
					param[i] = m_param[i];
				}
				param[SB_DoFunc] = 1;
				if (pFolderItem || wFlags & (SBSP_PARENT | SBSP_NAVIGATEBACK | SBSP_NAVIGATEFORWARD | SBSP_RELATIVE)) {
					Navigate3(pFolderItem, wFlags, param, &pSB, pFolderItems);
				}

				if (pVarResult) {
					//Navigate
					if (dispIdMember == 0x10000004) {
						teSetObject(pVarResult, pSB ? pSB : this);
					}
					else if (pSB) {
						pSB->Release();
					}
				}
			}
			//History
			if (dispIdMember == 0x1000000B) {
				if (pVarResult) {
					CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL, true);
					VARIANT v;
					for (int i = 1; i <= m_nLogCount; i++) {
						teSetObject(&v, m_ppLog[m_nLogCount - i]);
						pFolderItems->ItemEx(-1, NULL, &v);
						VariantClear(&v);
					}
					pFolderItems->m_nIndex = m_nLogCount - 1 - m_nLogIndex;
					teSetObject(pVarResult, pFolderItems);
					pFolderItems->Release();
				}
			}
			return S_OK;	
		//Navigate2Ex
		case 0x20000000:
			if (nArg >= 3) {
				FolderItem *pFolderItem = NULL;
				IUnknown *punk;
				if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
					punk->QueryInterface(IID_PPV_ARGS(&pFolderItem));
				}
				if (!pFolderItem && !m_pidl) {
					GetFolderItemFromPidl(&pFolderItem, g_pidls[CSIDL_DESKTOP]);
				}
				if (pFolderItem) {
					DWORD param[SB_Count];
					for (int i = SB_Count; i--;) {
						param[i] = m_param[i];
					}
					FolderItems *pFolderItems = NULL;
					LPITEMIDLIST pidl = NULL;
					wFlags = static_cast<WORD>(GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]));
					if (FindUnknown(&pDispParams->rgvarg[nArg - 2], &punk)) {
						punk->QueryInterface(IID_PPV_ARGS(&pFolderItems));
					}
					if (FindUnknown(&pDispParams->rgvarg[nArg - 3], &punk)) {
						GetIDListFromObject(punk, &pidl);
					}
					Navigate2(pFolderItem, wFlags, param, pFolderItems, pidl);
				}
			}
			return S_OK;	
		//Navigate2
		case 0x10000007:
			if (nArg >= 0) {
				DWORD param[SB_Count];
				for (int i = SB_Count; i--;) {
					m_param[i] = g_paramFV[i];
					param[i] = g_paramFV[i];
				}
				FolderItem *pFolderItem = NULL;
				FolderItems *pFolderItems = NULL;
				GetVariantPath(&pFolderItem, &pFolderItems, &pDispParams->rgvarg[nArg]);
				if (nArg >= 1) {
					wFlags = static_cast<WORD>(GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]));
				}
				if (nArg >= SB_Root + 2) {
					VariantCopy(&m_vRoot, &pDispParams->rgvarg[nArg - 2 - SB_Root]);
				}
				for (int i = 0; i <= nArg - 2 && i < SB_DoFunc; i++) {
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
				if (pVarResult) {
					teSetObject(pVarResult, pSB ? pSB : this);
				}
			}
			return S_OK;	
		//Index
		case 0x10000008:
			if (nArg >= 0) {
				m_pTabs->Move(GetTabIndex(), GetIntFromVariant(&pDispParams->rgvarg[nArg]), m_pTabs); 
			}
			if (pVarResult) {
				pVarResult->lVal = GetTabIndex();
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//FolderFlags
		case 0x10000009:
			if (nArg >= 0) {
				m_param[SB_FolderFlags] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
#ifdef _VISTA7
				if (m_pShellView) {
					IFolderView2 *pFV2;
					if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
						pFV2->SetCurrentFolderFlags(MAXDWORD32, m_param[SB_FolderFlags]);
						pFV2->Release();
					}
				}
#endif
			}
			if (pVarResult) {
				pVarResult->lVal = m_param[SB_FolderFlags];
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//CurrentViewMode
		case 0x10000010:
#ifdef _VISTA7
			IFolderView2 *pFV2;
			if (nArg >= 1 && m_pShellView && SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2)))) {
				m_param[SB_ViewMode] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				m_param[SB_IconSize] = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
				pFV2->SetViewModeAndIconSize(static_cast<FOLDERVIEWMODE>(m_param[SB_ViewMode]), m_param[SB_IconSize]);			
				pFV2->Release();
			}
			else
#endif
			if (nArg >= 0) {
				m_param[SB_ViewMode] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (m_pShellView) {
					IFolderView *pFV;
					if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
						pFV->SetCurrentViewMode(m_param[SB_ViewMode]);
						pFV->Release();
					}
				}
			}
			if (pVarResult) {
				if (m_pShellView) {
					FOLDERSETTINGS fs;
					if SUCCEEDED(m_pShellView->GetCurrentInfo(&fs)) {
						m_param[SB_ViewMode] = fs.ViewMode;
					}
				}
				pVarResult->lVal = m_param[SB_ViewMode];
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//IconSize
		case 0x10000011:
			if (nArg >= 0) {
				m_param[SB_IconSize] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
			}
			if (m_pShellView && m_bIconSize) {
#ifdef _VISTA7
				IFolderView2 *pFV2;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
					FOLDERVIEWMODE uViewMode;
					int iImageSize;
					pFV2->GetViewModeAndIconSize(&uViewMode, &iImageSize);
					if (nArg >= 0) {
						if (iImageSize != m_param[SB_IconSize]) {
							pFV2->SetViewModeAndIconSize(uViewMode, m_param[SB_IconSize]);
						}
					}
					else {
						m_param[SB_IconSize] = iImageSize;
					}
					pFV2->Release();
				}
				else
#endif
				{
#ifdef _2000XP
					if (m_param[SB_ViewMode] == FVM_ICON && m_param[SB_IconSize] >= 96) {
						m_param[SB_ViewMode] = FVM_THUMBNAIL;
						if (nArg >= 0) {
							if (m_pShellView) {
								IFolderView *pFV;
								if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
									pFV->SetCurrentViewMode(m_param[SB_ViewMode]);
									pFV->Release();
								}
							}
						}
					}
					if (m_param[SB_ViewMode] == FVM_ICON) {
						m_param[SB_IconSize] = GetSystemMetrics(SM_CXICON);
					}
					else if (m_param[SB_ViewMode] == FVM_THUMBNAIL) {
						HKEY hKey;
						RegOpenKeyEx(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer", 0, KEY_READ, &hKey);
						DWORD dwSize = sizeof(m_param[SB_IconSize]);
						RegQueryValueEx(hKey, L"ThumbnailSize", NULL, NULL, (LPBYTE)&m_param[SB_IconSize], &dwSize);
						RegCloseKey(hKey);
					}
					else {
						m_param[SB_IconSize] = GetSystemMetrics(SM_CXSMICON);
					}
#endif
				}
			}
			if (pVarResult) {
				pVarResult->lVal = m_param[SB_IconSize];
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//Options
		case 0x10000012:
			if (nArg >= 0) {
				m_param[SB_Options] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
#ifdef _VISTA7
				if (m_pExplorerBrowser) {
					m_pExplorerBrowser->SetOptions(static_cast<EXPLORER_BROWSER_OPTIONS>(m_param[SB_Options]));
				}
#endif
			}
			if (pVarResult) {
#ifdef _VISTA7
				if (m_pExplorerBrowser) {
					m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
				}
#endif
				pVarResult->vt = VT_I4;
				pVarResult->lVal = m_param[SB_Options];
			}
			return S_OK;
		//ViewFlags	
		case 0x10000016:
			if (nArg >= 0) {
				m_param[SB_ViewFlags] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
			}
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = m_param[SB_ViewFlags];
			}
			return S_OK;
		//Id
		case 0x10000017:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = MAX_FV - m_nSB;
			}
			return S_OK;
		//FilterView
		case 0x10000018:
			if (nArg >= 0) {
				VARIANT v;
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				::SysReAllocString(&m_bsFilter, v.bstrVal);
				VariantClear(&v);
			}
			if (pVarResult) {
				pVarResult->vt = VT_BSTR;
				pVarResult->bstrVal = SysAllocString(m_bsFilter);
			}
			return S_OK;
		//FolderItem
		case 0x10000020:	
			if (pVarResult) {
				if (m_pFolderItem) {
					teSetObject(pVarResult, m_pFolderItem);
				}
				else {
					FolderItem *pid;
					if (GetFolderItemFromPidl(&pid, m_pidl)) {
						teSetObject(pVarResult, pid);
						pid->Release();
					}
				}
			}
			return S_OK;
		//Parent
		case 0x10000024:
			if (pVarResult) {
				IDispatch *pdisp;
				if SUCCEEDED(get_Parent(&pdisp)) {
					teSetObject(pVarResult, pdisp);
					pdisp->Release();
				}
			}
			return S_OK;
		//Close
		case 0x10000031:
			Close(false);
			return S_OK;
		//Title
		case 0x10000032:
			nIndex = GetTabIndex();
			if (nIndex >= 0) {
				LPWSTR pszText = new WCHAR[MAX_PATH];
				tcItem.pszText = pszText;
				tcItem.mask = TCIF_TEXT;
				tcItem.cchTextMax = MAX_PATH;
				if (nArg >= 0) {
					VARIANT vText;
					teVariantChangeType(&vText, &pDispParams->rgvarg[nArg], VT_BSTR);
					int nCount = SysStringLen(vText.bstrVal);
					if (nCount >= MAX_PATH) {
						nCount = MAX_PATH - 1;
					}
					int j = 0;
					for (int i = 0; i < nCount; i++) {
						pszText[j++] = vText.bstrVal[i];
						if (vText.bstrVal[i] == '&') {
							pszText[j++] = '&';
						}
					}
					pszText[j] = NULL;
					TabCtrl_SetItem(m_pTabs->m_hwnd, nIndex, &tcItem);
					VariantClear(&vText);
					ArrangeWindow();
				}
				if (pVarResult) {
					TabCtrl_GetItem(m_pTabs->m_hwnd, nIndex, &tcItem);
					int nCount = lstrlen(tcItem.pszText);
					int j = 0;
					WCHAR c = NULL;
					for (int i = 0; i < nCount; i++) {
						pszText[j] = pszText[i];
						if (c != '&' || pszText[i] != '&') {
							j++;
							c = pszText[i];
						}
						else {
							c = NULL;
						}
					}
					pszText[j] = NULL;
					pVarResult->vt = VT_BSTR;
					pVarResult->bstrVal = SysAllocString(pszText);
				}
				delete [] pszText;
			}
			return S_OK;
		//ShellFolderView
		case 0x10000050:
			if (pVarResult) {
				teSetObject(pVarResult, m_pDSFV);
			}
			return S_OK;			
		//hwndView
		case 0x10000103:
			if (pVarResult) {
#ifdef _VISTA7
				if (m_pExplorerBrowser) {
					HWND hwnd = NULL;
					EnumChildWindows(m_hwnd, EnumChildProc, (LPARAM)&hwnd);
					SetVariantLL(pVarResult, (LONGLONG)hwnd);
				}
				else 
#endif
				{
					SetVariantLL(pVarResult, (LONGLONG)m_hwnd);
				}
			}
			return S_OK;
		//SortColumn
		case 0x10000104:
			if (nArg >= 0) {
				VARIANT v;
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				SetSort(v.bstrVal);
				VariantClear(&v);
			}
			if (pVarResult) {
				pVarResult->vt = VT_BSTR;
				GetSort(&pVarResult->bstrVal);
			}
			return S_OK;
		//Focus
		case 0x10000106:
			if (m_pShellView && m_bVisible) {
				m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
			}
			return S_OK;
		//HitTest (Screen coordinates)
		case 0x10000107:
			if (nArg >= 1 && pVarResult) {
				LVHITTESTINFO info;
				GetPointFormVariant(&info.pt, &pDispParams->rgvarg[nArg]);
				UINT flags = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
				pVarResult->lVal = (int)DoHitTest(this, info.pt, flags);
				pVarResult->vt = VT_I4;
				if (pVarResult->lVal < 0) {
					HWND hwnd = m_hwnd;
#ifdef _VISTA7
					if (m_pExplorerBrowser) {
						EnumChildWindows(m_hwnd, EnumChildProc, (LPARAM)&hwnd);
					}
#endif
					hwnd = FindWindowEx(hwnd, 0, WC_LISTVIEW, NULL);
					if (hwnd) {
						ScreenToClient(hwnd, &info.pt);
						info.flags = flags;
						ListView_HitTest(hwnd, &info);
						if (info.flags & flags) {
							pVarResult->lVal = info.iItem;
						}
					}
				}
			}
			return S_OK;			
		//DropTarget
		case 0x10000108:
			if (pVarResult) {
				CteDropTarget *pCDT;
				if (m_pDropTarget) {
					pCDT = new CteDropTarget(static_cast<IDropTarget *>(this), NULL);
				}
				else {
					HWND hwnd = m_hwnd;
#ifdef _VISTA7
					if (m_pExplorerBrowser) {
						EnumChildWindows(m_hwnd, EnumChildProc, (LPARAM)&hwnd);
					}
#endif
					hwnd = FindWindowEx(hwnd, 0, WC_LISTVIEW, NULL);
					pCDT = new CteDropTarget(static_cast<IDropTarget *>(GetProp(hwnd, L"OleDropTargetInterface")), NULL);
				}
				if (pCDT) {
					teSetObject(pVarResult, pCDT);
					pCDT->Release();
				}
			}
			return S_OK;
		//Columns
		case 0x10000109:
			if (nArg >= 0) {
				VARIANT v;
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				int nCount = 0;
				LPTSTR *lplpszArgs = NULL;
				method *methodArgs = NULL;
				int *pi = NULL;
				if (v.bstrVal) {
					lplpszArgs = CommandLineToArgvW(v.bstrVal, &nCount);
					nCount /= 2;
					methodArgs = new method[nCount];
					for (int i = nCount - 1; i >= 0; i--) {
						methodArgs[i].name = lplpszArgs[i * 2];
						if (!StrToIntEx(lplpszArgs[i * 2 + 1], STIF_DEFAULT, (int *)&methodArgs[i].id)) {
							methodArgs[i].id = -1;
						}
					}
					pi = SortMethod(methodArgs, nCount * sizeof(method));
				}
				VariantClear(&v);
#ifdef _VISTA7
				IColumnManager *pColumnManager;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager))) {
					UINT uCount;
					if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_ALL, &uCount)) {
						PROPERTYKEY *propKey = new PROPERTYKEY[uCount];
						if SUCCEEDED(pColumnManager->GetColumns(CM_ENUM_ALL, propKey, uCount)) {
							method *methodProp = new method[uCount];
							int *piprop = new int[uCount];
							CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_NAME };
							for (UINT i = 0; i < uCount; i++) {
								if SUCCEEDED(pColumnManager->GetColumnInfo(propKey[i], &cmci)) {
									methodProp[i].name = ::SysAllocString(cmci.wszName);
								}
								else {
									methodProp[i].name = NULL;
								}
							}
							piprop = SortMethod(methodProp, uCount * sizeof(method));
							int k = 0;
							PROPERTYKEY *propKey2 = new PROPERTYKEY[uCount];
							int *pnWidth = new int[uCount];
							int nIndex;
							for (int i = 0; i < nCount; i++) {
								nIndex = teBSearch(methodProp, uCount, piprop, methodArgs[i].name);
								if (nIndex >= 0) {
									pnWidth[k] = methodArgs[i].id;
									propKey2[k++] = propKey[piprop[nIndex]];
								}
							}
							if (k == 0) {	//Default Columns
								while (k < (int)m_nDefultColumns && k < (int)uCount) {
									propKey2[k] = m_pDefultColumns[k];
									pnWidth[k++] = -1;
								}
							}
							if (k > 0) {
								if SUCCEEDED(pColumnManager->SetColumns(propKey2, k)) {
									CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_WIDTH };
									while (k--) {
										cmci.uWidth = pnWidth[k];
										pColumnManager->SetColumnInfo(propKey2[k], &cmci);
									}
								}
							}
							delete [] pnWidth;
							delete [] propKey2;
							delete [] piprop;
							for (int i = uCount - 1; i >= 0; i--) {
								::SysFreeString(methodProp[i].name);
							}
							delete [] methodProp;
						}
						delete [] propKey;
					}
					pColumnManager->Release();
				}
				else
#endif
				{
#ifdef _2000XP
					UINT uCount = m_nColumns;
					IShellFolder2 *pSF2;
					SHELLDETAILS sd;
					int nWidth;
					if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pSF2))) {
						if (uCount == 0) {
							while (pSF2->GetDetailsOf(NULL, uCount, &sd) == S_OK && uCount < 8192) {
								uCount++;
							}
							m_pColumns = new TEColumn[uCount];
						}
						for (int i = uCount - 1; i >= 0; i--) {
							pSF2->GetDefaultColumnState(i, &m_pColumns[i].csFlags);
							m_pColumns[i].csFlags &= ~SHCOLSTATE_ONBYDEFAULT;
							m_pColumns[i].nWidth = -3;
						}
						m_nColumns = uCount;

						BSTR bs = GetColumnsStr();
						int nCur;
						LPTSTR *lplpszColumns = CommandLineToArgvW(bs, &nCur);
						::SysFreeString(bs);
						nCur /= 2;
						method *methodColumns = new method[nCur];
						for (int i = nCur - 1; i >= 0; i--) {
							methodColumns[i].name = lplpszColumns[i * 2];
						}
						BOOL bDiff = nCount != nCur;
						if (!bDiff) {
							int *piCur = SortMethod(methodColumns, nCur * sizeof(method));
							for (int i = nCur - 1; i >= 0; i--) {
								if (lstrcmpi(methodArgs[pi[i]].name, methodColumns[piCur[i]].name)) {
									bDiff = true;
									break;
								}
							}
							delete [] piCur;
						}
						delete [] methodColumns;
						LocalFree(lplpszColumns);

						if (bDiff) {
							int nIndex;
							BOOL bDefault = true;
							for (UINT i = 0; (pSF2->GetDetailsOf(NULL, i, &sd) == S_OK); i++) {
								BSTR bs;
								if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs)) {
									nIndex = teBSearch(methodArgs, nCount, pi, bs);
									if (nIndex >= 0) {
										m_pColumns[i].csFlags |= SHCOLSTATE_ONBYDEFAULT;
										m_pColumns[i].nWidth = methodArgs[pi[nIndex]].id;
										bDefault = false;
									}
									::SysFreeString(bs);
								}
							}
							if (bDefault && m_nColumns) {
								m_nColumns = 0;
								delete [] m_pColumns;
							}
							IShellView *pPreviusView = m_pShellView;
							FOLDERSETTINGS fs;
							pPreviusView->GetCurrentInfo(&fs);
							m_param[SB_ViewMode] = fs.ViewMode;
							Show(FALSE);
							if SUCCEEDED(CreateViewWindowEx(pPreviusView)) {
								pPreviusView->DestroyViewWindow();
								pPreviusView->Release();
								pPreviusView = NULL;
							}
							Show(TRUE);
						}
						//Columns Order and Width;
						HWND hList = FindWindowEx(m_hwnd, 0, WC_LISTVIEW, NULL);
						if (hList) {
							HWND hHeader = ListView_GetHeader(hList);
							if (hHeader) {
								UINT nHeader = Header_GetItemCount(hHeader);
								if (nHeader == nCount) {
									SetRedraw(false);
									int *piOrderArray = new int[nHeader];
									try {
										HD_ITEM hdi;
										::ZeroMemory(&hdi, sizeof(HD_ITEM));
										hdi.mask = HDI_TEXT | HDI_WIDTH;
										hdi.cchTextMax = MAX_COLUMN_NAME_LEN;
										WCHAR szText[MAX_COLUMN_NAME_LEN];
										hdi.pszText = (LPWSTR)&szText;
										int nIndex;
										for (int i = nHeader - 1; i >= 0; i--) {
											Header_GetItem(hHeader, i, &hdi);
											nIndex = teBSearch(methodArgs, nCount, pi, szText); 
											if (nIndex >= 0) {
												nWidth = methodArgs[pi[nIndex]].id; 
												piOrderArray[pi[nIndex]] = i;
												if (nWidth != hdi.cxy) {
													if (nWidth == -2) {
														nWidth = LVSCW_AUTOSIZE;
													}
													else if (nWidth < 0) {
														if (!bDiff) {
															SHELLDETAILS sd;
															for (UINT k = 0; (pSF2->GetDetailsOf(NULL, k, &sd) == S_OK); k++) {
																BSTR bs;
																if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs)) {
																	int nCC = lstrcmpi(szText, bs);
																	::SysFreeString(bs);
																	if (nCC == 0) {
																		nWidth = sd.cxChar * g_nAvgWidth;																
																		break;
																	}
																}
															}
														}
													}
													ListView_SetColumnWidth(hList, i, nWidth);
												}
											}
										}
										ListView_SetColumnOrderArray(hList, nHeader, piOrderArray);
									} catch (...) {
									}
									delete [] piOrderArray;
									SetRedraw(true);
								}
							}
						}
						pSF2->Release();
					}
#endif
				}
				if (lplpszArgs) {
					delete [] pi;
					delete [] methodArgs;
					LocalFree(lplpszArgs);
				}
			}
			if (pVarResult) {
				pVarResult->vt = VT_BSTR;
				pVarResult->bstrVal = GetColumnsStr();
			}
			return S_OK;
		//Items
		case 0x10000202:
			if (pVarResult && m_pShellView) {
				IDataObject *pDataObj;
				if (FAILED(m_pShellView->GetItemObject(SVGIO_ALLVIEW | SVGIO_FLAG_VIEWORDER, IID_PPV_ARGS(&pDataObj)))) {
					pDataObj = NULL;
				}
				FolderItems *pFolderItems;
				pFolderItems = new CteFolderItems(pDataObj, NULL, false);
				teSetObject(pVarResult, pFolderItems);
				pFolderItems->Release();
				if (pDataObj) {
					pDataObj->Release();
				}
			}
			return S_OK;
		//SelectedItems
		case 0x10000203:
			if (pVarResult) {
				FolderItems *pFolderItems;
				if SUCCEEDED(SelectedItems(&pFolderItems)) {
					teSetObject(pVarResult, pFolderItems);
					pFolderItems->Release();
				}
			}
			return S_OK;
		//Refresh
		case 0x10000206:
			Refresh();
			return S_OK;
		//ViewMenu
		case 0x10000207:
			if (pVarResult) {
				IContextMenu *pCM;
				if SUCCEEDED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&pCM))) {
					IDataObject *pDataObj = NULL;
					m_pShellView->GetItemObject(SVGIO_ALLVIEW, IID_PPV_ARGS(&pDataObj));
					CteContextMenu *pCCM;
					pCCM = new CteContextMenu(pCM, pDataObj);
					teSetObject(pVarResult, pCCM);
					pCCM->Release();
					pCM->Release();
				}
			}
			return S_OK;
		//TranslateAccelerator
		case 0x10000208:
			HRESULT hr;
			hr = E_NOTIMPL;
			if (m_pShellView) {
				if (nArg >= 3) {
					MSG msg;
					msg.hwnd = (HWND)GetLLFromVariant(&pDispParams->rgvarg[nArg]);
					msg.message = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
					msg.wParam = (WPARAM)GetLLFromVariant(&pDispParams->rgvarg[nArg - 2]);
					msg.lParam = (LPARAM)GetLLFromVariant(&pDispParams->rgvarg[nArg - 3]);
					hr = m_pShellView->TranslateAcceleratorW(&msg);
				}
			}
			if (pVarResult) {
				pVarResult->lVal = hr;
				pVarResult->vt = VT_I4;
			}
			return S_OK;		
		case 0x10000209:// GetItemPosition
			hr = E_NOTIMPL;
			if (m_pShellView) {
				if (nArg >= 1) {
					char *pc = GetpcFormVariant(&pDispParams->rgvarg[nArg - 1]);
					if (pc) {
						IFolderView *pFV;
						if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
							LPITEMIDLIST pidl;
							if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
								hr = pFV->GetItemPosition(ILFindLastID(pidl), (LPPOINT)pc);
								::CoTaskMemFree(pidl);
							}
							pFV->Release();
						}
					}
				}
			}
			if (pVarResult) {
				pVarResult->lVal = hr;
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//SelectAndPositionItem
		case 0x1000020A:
			hr = E_NOTIMPL;
			if (m_pShellView) {
				if (nArg >= 2) {
					char *pc = GetpcFormVariant(&pDispParams->rgvarg[nArg - 2]);
					if (pc) {
						IShellView2 *pSV2;
						if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSV2))) {
							LPITEMIDLIST pidl;
							if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
								hr = pSV2->SelectAndPositionItem(ILFindLastID(pidl), GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), (LPPOINT)pc);
								::CoTaskMemFree(pidl);
							}
							pSV2->Release();
						}
					}
				}
			}
			if (pVarResult) {
				pVarResult->lVal = hr;
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//TreeView
		case 0x10000210:
			if (pVarResult) {
				teSetObject(pVarResult, m_pTV);
			}
			return S_OK;
		//SelectItem
		case 0x10000280:
			if (nArg >= 1) {
				if (m_pShellView) {
					LPITEMIDLIST pidl;
					GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg]);
					m_pShellView->SelectItem(ILFindLastID(pidl), GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]));
					if (pidl) {
						::CoTaskMemFree(pidl);
					}
				}
			}
			return S_OK;
		//FocusedItem
		case 0x10000281:
			if (pVarResult) {
				FolderItem *pFolderItem = NULL;
				if (get_FocusedItem(&pFolderItem) == S_OK) {
					teSetObject(pVarResult, pFolderItem);
					pFolderItem->Release();
				}
			}
			return S_OK;
		//GetFocusedItem
		case 0x10000282:
			if (pVarResult) {
				IFolderView *pFV;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
					if SUCCEEDED(pFV->GetFocusedItem(&pVarResult->intVal)) {
						pVarResult->vt = VT_I4;
#ifndef _W2000
						return S_OK;
#endif
					}
				}
#ifdef _W2000
				//Windows 2000
				HWND hList = FindWindowEx(m_hwnd, 0, WC_LISTVIEW, NULL);
				if (hList) {
					pVarResult->lVal = ListView_GetNextItem(hList, -1, LVNI_ALL | LVNI_FOCUSED);
					pVarResult->vt = VT_I4;
				}
#endif
			}
			return S_OK;
		//
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
		default:
			if (dispIdMember >= START_OnFunc && dispIdMember < START_OnFunc + Count_SBFunc) {
				if (nArg >= 0) {
					if (m_pOnFunc[dispIdMember - START_OnFunc]) {
						m_pOnFunc[dispIdMember - START_OnFunc]->Release();
						m_pOnFunc[dispIdMember - START_OnFunc] = NULL;
					}
					IUnknown *punk;
					if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
						punk->QueryInterface(IID_PPV_ARGS(&m_pOnFunc[dispIdMember - START_OnFunc]));
					}
				}
				if (pVarResult) {
					if (m_pOnFunc[dispIdMember - START_OnFunc]) {
						teSetObject(pVarResult, m_pOnFunc[dispIdMember - START_OnFunc]); 
					}
				}
				return S_OK;
			}
	}
	if (m_pDSFV) {
		return m_pDSFV->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	}
	return DISP_E_MEMBERNOTFOUND;
}

STDMETHODIMP CteShellBrowser::get_Application(IDispatch **ppid)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->get_Application(ppid);
		pSFVD->Release();
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::get_Parent(IDispatch **ppid)
{
	return m_pTabs->QueryInterface(IID_PPV_ARGS(ppid));
}

STDMETHODIMP CteShellBrowser::get_Folder(Folder **ppid)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->get_Folder(ppid);
		pSFVD->Release();
	}
	return SUCCEEDED(hr) ? hr : GetFolderObjFromPidl(m_pidl, ppid);
}

STDMETHODIMP CteShellBrowser::SelectedItems(FolderItems **ppid)
{	
	FolderItems *pItems = NULL;
/*	IShellFolderViewDual *pSFVD;			// Unstable (Windows2000)
	if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		pSFVD->SelectedItems(&pItems);
		pSFVD->Release();
	}*/
	IDataObject *pDataObj;
	if (!m_pShellView || FAILED(m_pShellView->GetItemObject(SVGIO_SELECTION, IID_PPV_ARGS(&pDataObj)))) {
		pDataObj = NULL;
	}
	*ppid = new CteFolderItems(pDataObj, pItems, false);
	if (pDataObj) {
		pDataObj->Release();
	}
	return S_OK;
}

STDMETHODIMP CteShellBrowser::get_FocusedItem(FolderItem **ppid)
{
	HRESULT hr = E_NOTIMPL;
	if (m_pDSFV) {
		IShellFolderViewDual *pSFVD;
		if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
			hr = pSFVD->get_FocusedItem(ppid);
			pSFVD->Release();
		}
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::SelectItem(VARIANT *pvfi, int dwFlags)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->SelectItem(pvfi, dwFlags);
		pSFVD->Release();
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::PopupItemMenu(FolderItem *pfi, VARIANT vx, VARIANT vy, BSTR *pbs)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->PopupItemMenu(pfi, vx, vy, pbs);
		pSFVD->Release();
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::get_Script(IDispatch **ppDisp)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->get_Script(ppDisp);
		pSFVD->Release();
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::get_ViewOptions(long *plViewOptions)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->get_ViewOptions(plViewOptions);
		pSFVD->Release();
	}
	return hr;
}

int CteShellBrowser::GetTabIndex()
{
	if (m_pTabs) {
		int i;
		TC_ITEM tcItem;
		
		for (i = TabCtrl_GetItemCount(m_pTabs->m_hwnd) - 1; i >= 0; i--) {
			tcItem.mask = TCIF_PARAM;
			TabCtrl_GetItem(m_pTabs->m_hwnd, i, &tcItem);
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
	TEBrowse *pBrowse = new TEBrowse[1];
	pBrowse->pidl = ::ILClone(pidlFolder);
	pBrowse->nSB = m_nSB;
	SetTimer(g_hwndMain, (UINT_PTR)pBrowse, 10, teTimerProcBrowse);
	return E_FAIL;
}

STDMETHODIMP CteShellBrowser::OnViewCreated(IShellView *psv)
{
	if (psv) {
		psv->QueryInterface(IID_PPV_ARGS(&m_pShellView));
		GetPidlFromSV(&m_pidl, psv);
		if (!m_DefProc[1]) {
			HWND hwnd = 0;
			EnumChildWindows(m_hwnd, EnumChildProc, (LPARAM)&hwnd);
			if (hwnd) {
				m_DefProc[1] = SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LONG_PTR)TELVProc2);
			}
		}
	}
	if (m_pDSFV) {
		m_pDSFV->Release();
	}
	if FAILED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&m_pDSFV))) {
		m_pDSFV = NULL;
	}
	TC_ITEM tcItem;
	//View Tab Name
	int i;
	i = GetTabIndex();
	if (i >= 0) {
		VARIANT vResult;
		VariantInit(&vResult);
		vResult.vt = VT_I4;
		vResult.lVal = S_FALSE;
		if (DoFunc(TE_OnViewCreated, this, E_NOTIMPL) != S_OK) {
			BSTR bstr;
			if SUCCEEDED(GetDisplayNameFromPidl(&bstr, m_pidl, SHGDN_INFOLDER)) {
				tcItem.pszText = bstr;
				tcItem.mask = TCIF_TEXT;
				TabCtrl_SetItem(m_pTabs->m_hwnd, i, &tcItem);
				::SysFreeString(bstr);
				ArrangeWindow();
			}
		}
	}
	if (!m_pExplorerBrowser) {
		OnNavigationComplete(NULL);
	}
	return S_OK;
}

STDMETHODIMP CteShellBrowser::OnNavigationComplete(PCIDLIST_ABSOLUTE pidlFolder)
{
	m_bNoRowSelect = !(m_param[SB_FolderFlags] & FWF_FULLROWSELECT);
#ifdef _VISTA7
	if (m_param[SB_IconSize] && m_param[SB_ViewMode] == FVM_ICON) {
		IFolderView2 *pFV2;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
			pFV2->SetViewModeAndIconSize(static_cast<FOLDERVIEWMODE>(m_param[SB_ViewMode]), m_param[SB_IconSize]);
			pFV2->Release();
		}
	}
#endif
	m_bIconSize = TRUE;
	if ((m_param[SB_TreeAlign] & TE_TV_View) && m_pTV->m_bSetRoot) {
		if (m_pTV->Create()) {
			m_pTV->SetRoot();
		}
	}
	SetRedraw(TRUE);
	return S_OK;
}

STDMETHODIMP CteShellBrowser::OnNavigationFailed(PCIDLIST_ABSOLUTE pidlFolder)
{
	SetRedraw(TRUE);
	return S_OK;
}

//IExplorerPaneVisibility
STDMETHODIMP CteShellBrowser::GetPaneState(REFEXPLORERPANE ep, EXPLORERPANESTATE *peps)
{
	if (g_pOnFunc[TE_OnGetPaneState]) {
		VARIANTARG *pv = GetNewVARIANT(3);
		CteMemory *pstEps;
		pstEps = new CteMemory(sizeof(int), (char *)peps, VT_NULL, 1, L"DWORD");
		if SUCCEEDED(pstEps->QueryInterface(IID_PPV_ARGS(&pv[0].pdispVal))) {
			pv[0].vt = VT_DISPATCH;
		}
		pstEps->Release();
		WCHAR str[40];
		StringFromGUID2(ep, str, 40);
		pv[1].bstrVal = ::SysAllocString(str);
		pv[1].vt = VT_BSTR;

		if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pv[2].pdispVal))) {
			pv[2].vt = VT_DISPATCH;
		}
		VARIANT vResult;
		VariantInit(&vResult);
		if SUCCEEDED(Invoke4(g_pOnFunc[TE_OnGetPaneState], &vResult, 3, pv)) {
			return GetIntFromVariantClear(&vResult);
		}
	}
	return E_NOTIMPL;
}

#ifdef _2000XP
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
	if (osInfo.dwMajorVersion == 5 && osInfo.dwMinorVersion >=1 && IsEqualIID(riid, IID_IShellView)) {
		//only XP
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
	return m_pSF2->GetDetailsEx(pidl, pscid, pv);
}

STDMETHODIMP CteShellBrowser::GetDetailsOf(PCUITEMID_CHILD pidl, UINT iColumn, SHELLDETAILS *psd)
{
	HRESULT hr = m_pSF2->GetDetailsOf(pidl, iColumn, psd);

	if (m_nColumns) {
		if (iColumn < m_nColumns) {
			if (m_pColumns[iColumn].nWidth >= 0) {
				psd->cxChar = m_pColumns[iColumn].nWidth / g_nAvgWidth;
			}
		}
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::MapColumnToSCID(UINT iColumn, SHCOLUMNID *pscid)
{
	return m_pSF2->MapColumnToSCID(iColumn, pscid);
}

//IShellFolderViewCB
STDMETHODIMP CteShellBrowser::MessageSFVCB(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	try {
		switch(uMsg) {
			case SFVM_BACKGROUNDENUM:
				return S_OK;
			case SFVM_GETNOTIFY:
				if (wParam) {
					*(LPITEMIDLIST *) wParam = m_pidl;
				}
				if (lParam) {
					*(LONG *) lParam = SHCNE_ALLEVENTS;
				}
				return S_OK;
/*			case SFVM_BACKGROUNDENUMDONE:
				return S_OK;*/
/*			case SFVM_THISIDLIST:
				if (lParam) {
					*(LPITEMIDLIST *) lParam = ::ILClone(m_pidl);
					return S_OK;
				}
				break;
			case SFVM_DIDDRAGDROP:
				return S_OK;
			case SFVM_SIZE:
				return S_OK;
			case SFVM_FSNOTIFY:
				return S_OK;
			default:
				return E_NOTIMPL;*/
		}
	}
	catch (...) {
	}
	return E_NOTIMPL;
}
#endif

//IDropTarget
STDMETHODIMP CteShellBrowser::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);
	DWORD dwEffect = *pdwEffect;
	HRESULT hr = DragSub(TE_OnDragEnter, this, m_pDragItems, grfKeyState, pt, pdwEffect);
	if (hr != S_OK && m_pDropTarget) {
		*pdwEffect = dwEffect;
		hr = m_pDropTarget->DragEnter(pDataObj, grfKeyState, pt, pdwEffect);
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_grfKeyState = grfKeyState;
	DWORD dwEffect = *pdwEffect;
	HRESULT hr = DragSub(TE_OnDragOver, this, m_pDragItems, grfKeyState, pt, pdwEffect);
	if (hr != S_OK && m_pDropTarget) {
		*pdwEffect = dwEffect;
		hr = m_pDropTarget->DragOver(grfKeyState, pt, pdwEffect);
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::DragLeave()
{
	if (m_DragLeave == E_NOT_SET) {
		m_DragLeave = DragLeaveSub(this, &m_pDragItems);
		if (m_pDropTarget) {
			HRESULT hr = m_pDropTarget->DragLeave();
			if (m_DragLeave != S_OK) {
				m_DragLeave = hr;
			}
		}
	}
	return m_DragLeave;
}

STDMETHODIMP CteShellBrowser::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	GetDragItems(&m_pDragItems, pDataObj);
	DWORD dwEffect = *pdwEffect;
	HRESULT hr = DragSub(TE_OnDrop, this, m_pDragItems, m_grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		if (hr != S_OK) {
			*pdwEffect = dwEffect;
			hr = m_pDropTarget->Drop(pDataObj, grfKeyState, pt, pdwEffect);
		}
	}
	DragLeave();
	return hr;
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
	*ppidl = ::ILClone(m_pidl);
	return S_OK;
}

//
void CteShellBrowser::Show(BOOL bShow)
{
	if (bShow) {
		if (!m_pTabs->m_bVisible) {
			return;
		}
		if (m_nUnload == 1 && !g_nLockUpdate) {
			m_nUnload = 2;
			BrowseObject(NULL, SBSP_RELATIVE | SBSP_SAMEBROWSER);
			m_nUnload = 0;
		}
	}
	m_bVisible = bShow;
	if (m_pShellView) {
		if (bShow) {
			ShowWindow(m_hwnd, SW_SHOW);
			SetRedraw(TRUE);
			if (m_param[SB_TreeAlign] & TE_TV_View) {
				ShowWindow(m_pTV->m_hwnd, SW_SHOW);
				BringWindowToTop(m_pTV->m_hwnd);
			}
			m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
			BringWindowToTop(m_hwnd);
			HWND hwndDV = m_hwnd;
#ifdef _VISTA7
			if (m_pExplorerBrowser) {
				if (!m_DefProc[1]) {
					EnumChildWindows(m_hwnd, EnumChildProc, (LPARAM)&hwndDV);
					if (hwndDV != m_hwnd) {
						m_DefProc[1] = SetWindowLongPtr(hwndDV, GWLP_WNDPROC, (LONG_PTR)TELVProc2);
					}
				}
			}
			else 
#endif
			{
				if (!m_DefProc[0]) {
					m_DefProc[0] = SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)TELVProc);
				}
			}
			HWND hList = FindWindowEx(hwndDV, 0, WC_LISTVIEW, NULL);
			if (hList) {
				ShowWindow(hList, SW_SHOW);
			}
			HookDragDrop(g_dragdrop & 1);
			ArrangeWindow();
			if (m_nUnload == 3) {
				Refresh();
			}
		}
		else {
			SetRedraw(FALSE);
			m_pShellView->UIActivate(SVUIA_DEACTIVATE);
			MoveWindow(m_hwnd, -1, -1, 0, 0, false);
			ShowWindow(m_hwnd, SW_HIDE);
			HWND hwndDV = m_hwnd;
#ifdef _VISTA7
			if (m_pExplorerBrowser) {
				if (m_DefProc[1]) {
					EnumChildWindows(m_hwnd, EnumChildProc, (LPARAM)&hwndDV);
					if (hwndDV != m_hwnd) {
						SetWindowLongPtr(hwndDV, GWLP_WNDPROC, (LONG_PTR)m_DefProc[1]);
					}
				}
			}
			else
#endif
			if (m_DefProc[0]) {
				SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)m_DefProc[0]);
			}
			if (m_pDropTarget) {
				HWND hwnd = FindWindowEx(hwndDV, 0, WC_LISTVIEW, NULL);
				SetProp(hwnd, L"OleDropTargetInterface", (HANDLE)m_pDropTarget);
			}
			m_DefProc[1] = NULL;
			m_DefProc[0] = NULL;
			m_pDropTarget = NULL;
			ShowWindow(m_pTV->m_hwnd, SW_HIDE);
		}
	}
}

VOID CteShellBrowser::Close(BOOL bForce)
{
	if (m_bEmpty) {
		return;
	}
	if (CanClose(this) || bForce) {
		int i = GetTabIndex();
		m_bEmpty = true;
		if (i >= 0)	{
			SendMessage(m_pTabs->m_hwnd, TCM_DELETEITEM, i, 0);
		}
		DestroyView(0);
		ShowWindow(m_pTV->m_hwnd, SW_HIDE);
		VariantClear(&m_vRoot);
		m_vRoot.vt = VT_I4;
		m_vRoot.lVal = 0;
	}
}

VOID CteShellBrowser::DestroyView(int nFlags)
{
#ifdef _VISTA7
	if ((nFlags & 2) == 0 && m_pExplorerBrowser) {
		try {
			if (m_dwEventCookie) {
			  m_pExplorerBrowser->Unadvise(m_dwEventCookie);
			}
			IUnknown_SetSite(m_pExplorerBrowser, NULL);
			Show(false);
			m_pExplorerBrowser->Destroy();
			m_pExplorerBrowser->Release();
		} catch (...) {}
		m_pExplorerBrowser = NULL;
	}
#endif
	if (m_pShellView) {
		if ((nFlags & 1) == 0) {
			Show(false);
			m_pShellView->DestroyViewWindow();
		}
		if (nFlags == 0) {
			m_pShellView->Release();
			m_pShellView = NULL;
		}
	}
}

VOID CteShellBrowser::GetSort(BSTR* pbs)
{
	*pbs = NULL;
	if (!m_pShellView) {
		return;
	}
#ifdef _VISTA7
	IFolderView2 *pFV2;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		SORTCOLUMN col;
		if SUCCEEDED(pFV2->GetSortColumns(&col, 1)) {
			IColumnManager *pColumnManager;
			if SUCCEEDED(pFV2->QueryInterface(IID_PPV_ARGS(&pColumnManager))) {
				UINT uCount;
				if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_VISIBLE, &uCount)) {
					PROPERTYKEY *propKey = new PROPERTYKEY[uCount];
					if SUCCEEDED(pColumnManager->GetColumns(CM_ENUM_VISIBLE, propKey, uCount)) {
						for (UINT i = 0; i < uCount; i++) {
							if (col.propkey.pid == propKey[i].pid && IsEqualGUID(col.propkey.fmtid, propKey[i].fmtid)) {
								CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_NAME };
								if SUCCEEDED(pColumnManager->GetColumnInfo(propKey[i], &cmci)) {
									*pbs = SysAllocString(cmci.wszName);
									if (col.direction < 0) {
										ToMinus(pbs);
									}
									break;
								}
							}
						}
					}
					delete [] propKey;
				}
				pColumnManager->Release();
			}
		}
		pFV2->Release();
		return;
	}
#endif
#ifdef _2000XP
	BOOL bEmpty = TRUE;
	LPARAM Sort = 0;
	IShellFolderView *pSFV;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
		if SUCCEEDED(pSFV->GetArrangeParam(&Sort)) {//Sort column
			Sort = LOWORD(Sort);
		}
		pSFV->Release();
	}
	HWND hList = FindWindowEx(m_hwnd, 0, WC_LISTVIEW, NULL);
	if (hList) {
		HWND hHeader = ListView_GetHeader(hList);
		if (hHeader) {
			WCHAR szColumn[MAX_PATH];
			int nCount = Header_GetItemCount(hHeader);
			for (int i = (int)Sort; i < nCount; i++) {
				HD_ITEM hdi;
				::ZeroMemory(&hdi, sizeof(hdi));
#ifdef _W2000
				hdi.mask = HDI_FORMAT | HDI_TEXT | HDI_BITMAP;
#else
				hdi.mask = HDI_FORMAT | HDI_TEXT;
#endif
				hdi.cchTextMax = _countof(szColumn);
				hdi.pszText = szColumn;
				Header_GetItem(hHeader, i, &hdi);
				if (hdi.fmt & (HDF_SORTUP | HDF_SORTDOWN | HDF_BITMAP)) {
					bEmpty = FALSE;
					*pbs = SysAllocString(szColumn);
					if (hdi.fmt & HDF_SORTDOWN) {
						ToMinus(pbs);
					}
#ifdef _W2000
					if (hdi.fmt & HDF_BITMAP) {	//Windows 2000
						if (!(hdi.fmt & (HDF_SORTUP | HDF_SORTDOWN))) {
							HDC dc = GetDC(hHeader);
							HDC memDC = CreateCompatibleDC(dc);
							ReleaseDC(hHeader, dc);
							HGDIOBJ hDefault = SelectObject(memDC, hdi.hbm);
							if (GetPixel(memDC, 3, 4) != GetSysColor(COLOR_BTNFACE)) {
								ToMinus(pbs);
							}
							SelectObject(memDC, hDefault);
							DeleteDC(memDC);
						}
						DeleteObject(hdi.hbm);
					}
#endif
					break;
				}
			}
		}
	}
	if (bEmpty) {
		IShellFolder2 *pSF2;
		if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pSF2))) {
			SHELLDETAILS sd;
			if (pSF2->GetDetailsOf(NULL, (UINT)Sort, &sd) == S_OK) {
				if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, pbs)) {
					bEmpty = FALSE;
				}
			}
			pSF2->Release();
		}
	}
	if (bEmpty) {
		::SysFreeString(*pbs);
		*pbs = NULL;
	}
#endif
}

VOID CteShellBrowser::SetSort(BSTR bs)
{
	BSTR bs1;
	GetSort(&bs1);
	if (lstrlen(bs) < 1 || lstrcmpi(bs, bs1) == 0) {
		::SysFreeString(bs1);
		return;
	}
	int dir = (bs[0] == '-') ? 1 : 0;
	int dir1;
	if (bs1) {
		dir1 = lstrcmpi(&bs[dir], &bs1[(bs1 && bs1[0] == '-')]) ? 1 : 0;
		::SysFreeString(bs1);
	}
	else {
		dir1 = 1;
	}
#ifdef _VISTA7
	IColumnManager *pColumnManager;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager))) {
		UINT uCount;
		if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_VISIBLE, &uCount)) {
			if (uCount) {
				PROPERTYKEY *propKey = new PROPERTYKEY[uCount];
				if SUCCEEDED(pColumnManager->GetColumns(CM_ENUM_VISIBLE, propKey, uCount)) {
					CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_NAME };
					for (UINT i = 0; i < uCount; i++) {
						if SUCCEEDED(pColumnManager->GetColumnInfo(propKey[i], &cmci)) {
							if (lstrcmpi(&bs[dir], cmci.wszName) == 0) {
								SORTCOLUMN col;
								col.direction = 1 - dir * 2;
								IFolderView2 *pFV2;
								if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
									col.propkey = propKey[i];
									pFV2->SetSortColumns(&col, 1);
									pFV2->Release();
									break;
								}
							}
						}
					}
				}
				delete [] propKey;
			}
		}
		pColumnManager->Release();
		return;
	}
#endif
#ifdef _2000XP
	IShellFolderView *pSFV;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
		IShellFolder2 *pSF2;
		if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pSF2))) {
			SHELLDETAILS sd;
			HRESULT hr = E_FAIL;
			for (int i = 0; (pSF2->GetDetailsOf(NULL, i, &sd) == S_OK && FAILED(hr)); i++) {
				if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs1)) {
					if (lstrcmpi(bs1, &bs[dir]) == 0) {
						hr = pSFV->Rearrange(i);
						if (dir && dir1) {
							hr = pSFV->Rearrange(i);
						}
					}
					::SysFreeString(bs1);
				}
			}
			pSF2->Release();
		}
		pSFV->Release();
	}
#endif
}

HRESULT CteShellBrowser::SetRedraw(BOOL bRedraw)
{
	HRESULT hr = E_FAIL;
	if (m_pShellView) {
#ifdef _VISTA7
		IFolderView2 *pFV2;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
			hr = pFV2->SetRedraw(bRedraw);
			pFV2->Release();
		}
		else
#endif
		{
#ifdef _2000XP
			IShellFolderView *pSFV;
			if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
				hr = pSFV->SetRedraw(bRedraw);
				pSFV->Release();
			}
#endif
		}
	}
	return hr;
}

HRESULT CteShellBrowser::CreateViewWindowEx(IShellView *pPreviusView)
{
	HRESULT hr = E_FAIL;
	m_pShellView = NULL;
	RECT rc;
	if (m_pSF2) {
#ifdef _2000XP
		if SUCCEEDED(CreateViewObject(NULL, IID_PPV_ARGS(&m_pShellView)) && m_pShellView) {
			if (osInfo.dwMajorVersion <= 5 && m_param[SB_ViewMode] == FVM_ICON && m_param[SB_IconSize] >= 96) {
				m_param[SB_ViewMode] = FVM_THUMBNAIL;
			}
#else
		if SUCCEEDED(m_pSF2->CreateViewObject(NULL, IID_PPV_ARGS(&m_pShellView)) && m_pShellView) {
#endif
			if (::InterlockedIncrement(&m_nCreate) == 1) {
				try {
					hr = m_pShellView->CreateViewWindow(pPreviusView, (LPCFOLDERSETTINGS)&m_param[SB_ViewMode], static_cast<IShellBrowser *>(this), &rc, &m_hwnd);
					GetDefaultColumns();
				} catch (...) {
					hr = E_FAIL;
				}
				InterlockedDecrement(&m_nCreate);
			}
			else {
				hr = E_FAIL;
			}
			if (hr == S_OK) {
				if (!(m_param[SB_FolderFlags] & FWF_NOCLIENTEDGE)) {
					HWND hwnd = FindWindowEx(m_hwnd, 0, WC_LISTVIEW, NULL);
					if (hwnd) {
						SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) | WS_BORDER);
					}
				}
			}
		}
	}
	return hr;
}

//CTE

CTE::CTE()
{
	m_cRef = 1;
	m_pDragItems = NULL;
	m_param[TE_Type] = CTRL_TE;
	VariantInit(&m_Data);

	for (int i = 1; i < _countof(m_param); i++) {
		m_param[i] = 0;
	}
	RegisterDragDrop(g_hwndMain, static_cast<IDropTarget *>(this));
}

CTE::~CTE()
{
	VariantClear(&m_Data);
	RevokeDragDrop(g_hwndMain);
}

STDMETHODIMP CTE::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else if (IsEqualIID(riid, IID_IDropTarget)) {
		*ppvObject = static_cast<IDropTarget *>(this);
	}
	else if (IsEqualIID(riid, IID_IDropSource)) {
		*ppvObject = static_cast<IDropSource *>(this);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
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
	return teGetDispId(methodTE, sizeof(methodTE), g_maps[MAP_TE], *rgszNames, rgDispId, true);
}

STDMETHODIMP CTE::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	IUnknown *punk = NULL;
	IDispatch *pdisp = NULL;

	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	if (dispIdMember >= TE_METHOD && wFlags == DISPATCH_PROPERTYGET) {
		if (pVarResult) {
			CteDispatch *pClone = new CteDispatch(this, 0);
			pClone->m_dispIdMember = dispIdMember;
			teSetObject(pVarResult, pClone);
			pClone->Release();
			return S_OK;
		}
		return S_OK;
	}
	switch(dispIdMember) {
		//Data
		case 1001:
			if (nArg >= 0) {
				VariantClear(&m_Data);
				VariantCopy(&m_Data, &pDispParams->rgvarg[nArg]);
			}
			if (pVarResult) {
				VariantCopy(pVarResult, &m_Data);
			}
			return S_OK;
		//hwnd
		case 1002:
			SetVariantLL(pVarResult, (LONGLONG)g_hwndMain);
			return S_OK;
		//About
		case 1004:
			if (pVarResult) {
				pVarResult->vt = VT_BSTR;
				pVarResult->bstrVal = SysAllocString(g_szTE);
			}
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
								int nId = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
								if (nId) {
									pSB = g_pSB[MAX_FV - nId];
								}
							}
							if (!pSB && g_pTabs) {
								pSB = g_pTabs->GetShellBrowser(g_pTabs->m_nIndex);
								if (!pSB) {
									pSB = GetNewShellBrowser(g_pTabs);
								}
							}
							if (pSB) {
								if (nCtrl == CTRL_TV) {
									teSetObject(pVarResult, pSB->m_pTV);
								}
								else {
									teSetObject(pVarResult, pSB);
								}
							}
							break;
						case CTRL_TC:
							CteTabs *pTC;
							pTC = g_pTabs;
							if (nArg >= 1) {
								int nId = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
								if (nId) {
									pTC = g_pTC[MAX_TC - nId];
								}
							}
							if (pTC) {
								teSetObject(pVarResult, pTC);
							}
							break;
						case CTRL_WB:
							teSetObject(pVarResult, g_pWebBrowser);
							break;
					}
				}
			}
			return S_OK;
		//Ctrls
		case TE_METHOD + 1006:
			if (pVarResult) {
				if (nArg >= 0) {
					int nCtrl;
					nCtrl = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					CteMemory *pMem = NULL;
					IDispatch **ppDisp;
					int nCount = 0;
					int i;

					switch (nCtrl) {
						case CTRL_FV:
						case CTRL_SB:
						case CTRL_EB:
						case CTRL_TV:
							i = MAX_FV;
							while (--i >= 0) {
								if (g_pSB[i]) {
									if (!g_pSB[i]->m_bEmpty) {
										nCount++;
									}
								}
								else {
									break;
								}
							}
							pMem = new CteMemory(nCount * sizeof(IDispatch *), NULL, VT_DISPATCH, nCount, NULL);
							ppDisp = (IDispatch **)pMem->m_pc;
							nCount = 0;
							i = MAX_FV;
							while (--i >= 0) {
								if (g_pSB[i]) {
									if (!g_pSB[i]->m_bEmpty) {
										if (nCtrl == CTRL_TV) {
											g_pSB[i]->m_pTV->QueryInterface(IID_PPV_ARGS(&ppDisp[nCount++]));
										}
										else {
											g_pSB[i]->QueryInterface(IID_PPV_ARGS(&ppDisp[nCount++]));
										}
									}
								}
								else {
									break;
								}
							}
							break;
						case CTRL_TC:
							i = MAX_TC;
							while (--i >= 0) {
								if (g_pTC[i]) {
									if (!g_pTC[i]->m_bEmpty) {
										nCount++;
									}
								}
								else {
									break;
								}
							}
							pMem = new CteMemory(nCount * sizeof(IDispatch *), NULL, VT_DISPATCH, nCount, NULL);
							ppDisp = (IDispatch **)pMem->m_pc;
							nCount = 0;
							i = MAX_TC;
							while (--i >= 0) {
								if (g_pTC[i]) {
									if (!g_pTC[i]->m_bEmpty) {
										g_pTC[i]->QueryInterface(IID_PPV_ARGS(&ppDisp[nCount++]));
									}
								}
								else {
									break;
								}
							}
							break;
						case CTRL_WB:
							pMem = new CteMemory(sizeof(IDispatch *), NULL, VT_DISPATCH, 1, NULL);
							ppDisp = (IDispatch **)pMem->m_pc;
							g_pWebBrowser->QueryInterface(IID_PPV_ARGS(&ppDisp[0]));
							break;
					}
					if (!pMem) {
						pMem = new CteMemory(0, NULL, VT_DISPATCH, 0, NULL);
					}
					teSetObject(pVarResult, pMem);
					pMem->Release();
				}
			}
			return S_OK;
		//Version
		case 1007:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = 20000000 + VER_Y * 10000 + VER_M * 100 + VER_D;
			}
			return S_OK;
		//ClearEvents
		case TE_METHOD + 1008:
			ClearEvents();
			g_bReload = false;
			return S_OK;
		//Reload
		case TE_METHOD + 1009:
			g_nLockUpdate = 0;
			g_bReload = true;
			ClearEvents();
			g_pWebBrowser->m_pWebBrowser->Refresh();
			SetTimer(g_hwndMain, TET_Reload, 2000, teTimerProc);
			return S_OK;
		case DISPID_WINDOWREGISTERED:
			if (g_pOnFunc[TE_OnWindowRegistered]) {
				VARIANT vResult;
				VariantInit(&vResult);
				VARIANTARG *pv = GetNewVARIANT(1);
				teSetObject(&pv[0], g_pTE);
				Invoke4(g_pOnFunc[TE_OnWindowRegistered], NULL, 1, pv);
			}
			return S_OK;
		//CreateObject
		case TE_METHOD + 1010:
		//GetObject
		case TE_METHOD + 1020:
			if (nArg >= 0) {
				if (pVarResult) {
					CLSID clsid;
					VARIANT vClass;
					teVariantChangeType(&vClass, &pDispParams->rgvarg[nArg], VT_BSTR);
					if SUCCEEDED(teCLSIDFromProgID(vClass.bstrVal, &clsid)) {
						if (dispIdMember == TE_METHOD + 1010) {
							IUnknown *punk;
							if SUCCEEDED(CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&punk))) {
								teSetObject(pVarResult, punk);
								punk->Release();
							}
						}
						else if (dispIdMember == TE_METHOD + 1020) {
							IUnknown *punk;
							if SUCCEEDED(GetActiveObject(clsid, NULL, &punk)) {
								teSetObject(pVarResult, punk);
								punk->Release();
							}
						}
					}
					VariantInit(&vClass);
				}
			}
			return S_OK;
		//WindowsAPI
		case 1030:
			if (pVarResult) {
				CteWindowsAPI *pAPI = new CteWindowsAPI();
				teSetObject(pVarResult, pAPI);
				pAPI->Release();
			}
			return S_OK;
		//CommonDialog
		case 1131:
			if (pVarResult) {
				CteCommonDialog *pCD = new CteCommonDialog();
				teSetObject(pVarResult, pCD);
				pCD->Release();
			}
			return S_OK;
		//GdiplusBitmap
		case 1132:
			if (pVarResult) {
				CteGdiplusBitmap *pGB = new CteGdiplusBitmap();
				teSetObject(pVarResult, pGB);
				pGB->Release();
			}
			return S_OK;
		//FolderItems
		case TE_METHOD + 1133:
			if (pVarResult) {
				CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL, true);
				teSetObject(pVarResult, pFolderItems);
				pFolderItems->Release();
			}
			return S_OK;
		//Object
		case TE_METHOD + 1134:
			if (pVarResult && g_pObject) {
				Invoke4(g_pObject, pVarResult, 0, NULL);
			}
			return S_OK;
		//Array
		case TE_METHOD + 1135:
			if (pVarResult && g_pArray) {
				Invoke4(g_pArray, pVarResult, 0, NULL);
			}
			return S_OK;
		//Collection
		case TE_METHOD + 1136:
			if (pVarResult && nArg >= 0 && GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
				CteDispatch *pObj = new CteDispatch(pdisp, 1);
				pdisp->Release();
				teSetObject(pVarResult, pObj);
				pObj->Release();
			}
			return S_OK;
#ifdef _USE_TESTOBJECT
		//TestObj
		case 1200:
			if (pVarResult) {
				CteTest *pO = new CteTest();
				teSetObject(pVarResult, pO);
				pO->Release();
			}
			return S_OK;
#endif
		//CtrlFromPoint
		case TE_METHOD + 1040:
			if (nArg >= 0) {
				if (pVarResult) {
					POINT pt;
					GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg]);
					IDispatch *pdisp;
					if SUCCEEDED(ControlFromhwnd(&pdisp, WindowFromPoint(pt))) {
						teSetObject(pVarResult, pdisp);
						pdisp->Release();
					}
				}
			}
			return S_OK;
		//CreateCtrl
		case TE_METHOD + 1050:
			if (nArg >= 4) {
				int i;
				i = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				switch (HIWORD(i)) {
					case 3:	//Tab
						CteTabs *pTC;
						pTC = GetNewTC();
						pTC->Init();
						for (i = 0; i <= nArg && i < _countof(pTC->m_param); i++) {
							pTC->m_param[i] = GetIntFromVariantPP(&pDispParams->rgvarg[nArg - i], i);
						}
						if (pTC->Create()) {
							if (pVarResult) {
								teSetObject(pVarResult, pTC);
							}
							g_pTabs = pTC;
							pTC->Show(TRUE);
						}
						else {
							pTC->Release();
						}
						break;
				}
			}
			return S_OK;

		//MainMenu
		case TE_METHOD + 1060:
			if (pVarResult) {
				if (nArg >= 0) {
					if (!g_hMenu) {
						g_hMenu = LoadMenu(g_hShell32, MAKEINTRESOURCE(216));
						if (!g_hMenu) {
							CteShellBrowser *pSB = GetNewShellBrowser(NULL);
							if (g_pDesktopFolder->CreateViewObject(g_hwndMain, IID_PPV_ARGS(&pSB->m_pShellView)) == S_OK) {
								FOLDERSETTINGS fs;
								fs.fFlags = FWF_NOICONS;
								fs.ViewMode = FVM_LIST;
								RECT rc;
								SetRectEmpty(&rc);
								if SUCCEEDED(pSB->m_pShellView->CreateViewWindow(NULL, &fs, pSB, &rc, &pSB->m_hwnd)) {
									pSB->m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
								}
							}
							pSB->Close(true);
						}
					}
					MENUITEMINFO mii;
					::ZeroMemory(&mii, sizeof(MENUITEMINFO));
					mii.cbSize = sizeof(MENUITEMINFO);
					mii.fMask  = MIIM_SUBMENU;
					GetMenuItemInfo(g_hMenu, GetIntFromVariant(&pDispParams->rgvarg[nArg]), FALSE, &mii);

					HMENU menu = CreatePopupMenu();
					UINT fState = MFS_DISABLED;
					if (g_pTabs) {
						CteShellBrowser *pSB = g_pTabs->GetShellBrowser(g_pTabs->m_nIndex);
						if (pSB) {
							fState = MFS_ENABLED;
						}
					}
					teCopyMenu(menu, mii.hSubMenu, fState);
					SetVariantLL(pVarResult, (LONGLONG)menu);
				}
			}
			return S_OK;
		//CtrlFromWindow
		case TE_METHOD + 1070:
			if (nArg >= 0) {
				if (pVarResult) {
					IDispatch *pdisp;
					if SUCCEEDED(ControlFromhwnd(&pdisp, (HWND)GetLLFromVariant(&pDispParams->rgvarg[nArg]))) {
						teSetObject(pVarResult, pdisp);
						pdisp->Release();
					}
				}
			}
			return S_OK;
		//LockUpdate
		case TE_METHOD + 1080:
			if (::InterlockedIncrement(&g_nLockUpdate) == 1) {
				SendMessage(g_hwndBrowser, WM_SETREDRAW, FALSE, 0);
			}
			return S_OK;
		//UnlockUpdate
		case TE_METHOD + 1090:
			if (::InterlockedDecrement(&g_nLockUpdate) <= 0) {
				g_nLockUpdate = 0;
				SendMessage(g_hwndBrowser, WM_SETREDRAW, TRUE, 0);
				RedrawWindow(g_hwndBrowser, NULL, 0, RDW_INVALIDATE | RDW_ALLCHILDREN);

				int i = MAX_TC;
				while (--i >= 0) {
					CteTabs *pTC = g_pTC[i];
					if (pTC) {
						if (pTC->m_bVisible) {
							CteShellBrowser *pSB = pTC->GetShellBrowser(pTC->m_nIndex);
							if (pSB && !pSB->m_bEmpty && pSB->m_nUnload & 1) {
								pSB->Show(true);
							}
						}
					}
					else {
						break;
					}
				}
				if (g_nSize >= MAXWORD) {
					g_nSize -= MAXWORD;
					ArrangeWindow();
				}
			}
			return S_OK;
		//HookDragDrop
		case TE_METHOD + 1100:
			if (nArg >= 0) {
				DWORD n = GetIntFromVariant(&pDispParams->rgvarg[nArg]) < CTRL_TV ? 1 : 2;
				if ((n | g_dragdrop) != g_dragdrop) {
					g_dragdrop |= n;
					int i = MAX_FV;
					while (--i >= 0) {
						if (g_pSB[i]) {
							g_pSB[i]->HookDragDrop(n);
						}
						else {
							break;
						}
					}
				}
			}
			return S_OK;
		//Value
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
		default:
			if (dispIdMember >= START_OnFunc && dispIdMember < START_OnFunc + Count_OnFunc) {
				if (nArg >= 0) {
					if (g_pOnFunc[dispIdMember - START_OnFunc]) {
						g_pOnFunc[dispIdMember - START_OnFunc]->Release();
						g_pOnFunc[dispIdMember - START_OnFunc] = NULL;
					}
					if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
						punk->QueryInterface(IID_PPV_ARGS(&g_pOnFunc[dispIdMember - START_OnFunc]));
					}
				}
				if (pVarResult) {
					if (g_pOnFunc[dispIdMember - START_OnFunc]) {
						teSetObject(pVarResult, g_pOnFunc[dispIdMember - START_OnFunc]);
					}
				}
				return S_OK;
			}
			//Type, offsetLeft etc.
			else if (dispIdMember >= TE_OFFSET && dispIdMember <= TE_OFFSET + TE_CmdShow) {
				if (nArg >= 0 && dispIdMember != TE_OFFSET) {
					m_param[dispIdMember - TE_OFFSET] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					if (dispIdMember >= TE_OFFSET + TE_Left && dispIdMember <= TE_Bottom + TE_OFFSET) {
						ArrangeWindow();
					}
				}
				if (pVarResult) {
					pVarResult->vt = VT_I4;
					pVarResult->lVal = m_param[dispIdMember - TE_OFFSET];
				}
			}
			return S_OK;
	}
//	return DISP_E_MEMBERNOTFOUND;
}

//IDropTarget
STDMETHODIMP CTE::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);
	return DragSub(TE_OnDragEnter, this, m_pDragItems, grfKeyState, pt, pdwEffect);
}

STDMETHODIMP CTE::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_grfKeyState = grfKeyState;
	return DragSub(TE_OnDragOver, this, m_pDragItems, grfKeyState, pt, pdwEffect);
}

STDMETHODIMP CTE::DragLeave()
{
	if (m_DragLeave == E_NOT_SET) {
		m_DragLeave = DragLeaveSub(this, &m_pDragItems);
	}
	return m_DragLeave;
}

STDMETHODIMP CTE::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	GetDragItems(&m_pDragItems, pDataObj);
	HRESULT hr = DragSub(TE_OnDrop, this, m_pDragItems, m_grfKeyState, pt, pdwEffect);
	DragLeave();
	return hr;
}

//IDropSource
STDMETHODIMP CTE::QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState)
{
	if (fEscapePressed) {
		return DRAGDROP_S_CANCEL;
	}
	if ((grfKeyState & (MK_LBUTTON | MK_RBUTTON)) == 0) {
		if (m_bDrop) {
			m_bDrop = false;
			return DRAGDROP_S_DROP;
		}
		return DRAGDROP_S_CANCEL;
	}
	if ((grfKeyState & (MK_LBUTTON | MK_RBUTTON)) == (MK_LBUTTON | MK_RBUTTON)) {
		return DRAGDROP_S_CANCEL;
	}
	return S_OK;
}

STDMETHODIMP CTE::GiveFeedback(DWORD dwEffect)
{
	return DRAGDROP_S_USEDEFAULTCURSORS;
}

// CteWebBrowser

CteWebBrowser::CteWebBrowser(HWND hwnd, WCHAR *szPath)
{
	m_cRef = 1;
	m_hwnd = hwnd;
	m_dwCookie = 0;
	m_pDragItems = NULL;
	VariantInit(&m_Data);

	HRESULT hr;
	MSG        msg;
	RECT       rc;

	IOleObject *pOleObject;
	hr = CoCreateInstance(CLSID_WebBrowser, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pWebBrowser));

	if SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleObject))) {
		pOleObject->SetClientSite(static_cast<IOleClientSite *>(this));

		SetRectEmpty(&rc);
		hr = pOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, &msg, static_cast<IOleClientSite *>(this), 0, g_hwndMain, &rc);

		IConnectionPointContainer *pCPC;
		hr = m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pCPC));
		if (SUCCEEDED(hr)) {
			IConnectionPoint *pCP;
			hr = pCPC->FindConnectionPoint(DIID_DWebBrowserEvents2, &pCP);
			if (SUCCEEDED(hr)) {
				IUnknown *punk;
				QueryInterface(IID_PPV_ARGS(&punk));
				hr = pCP->Advise(punk, &m_dwCookie);
				punk->Release();
				pCP->Release();
			}
			pCPC->Release();
		}
		pOleObject->Release();

		m_pWebBrowser->put_Offline(VARIANT_TRUE);
		m_bstrPath = SysAllocString(szPath);
		hr = m_pWebBrowser->Navigate(m_bstrPath, NULL, NULL, NULL, NULL);
		m_pWebBrowser->put_Visible(VARIANT_TRUE);
	}
}

CteWebBrowser::~CteWebBrowser()
{
	::SysFreeString(m_bstrPath);
	Close();
	VariantClear(&m_Data);
}

STDMETHODIMP CteWebBrowser::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IOleClientSite)) {
		*ppvObject = static_cast<IOleClientSite *>(this);
	}
	else if (IsEqualIID(riid, IID_IOleWindow) || IsEqualIID(riid, IID_IOleInPlaceSite)) {
		*ppvObject = static_cast<IOleInPlaceSite *>(this);
	}
	else if (IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else if (IsEqualIID(riid, IID_IDocHostUIHandler)) {
		*ppvObject = static_cast<IDocHostUIHandler *>(this);
	}
	else if (IsEqualIID(riid, IID_IDropTarget)) {
		*ppvObject = static_cast<IDropTarget *>(this);
	}
	else if (IsEqualIID(riid, IID_IWebBrowser2)) {
		return m_pWebBrowser->QueryInterface(riid, (LPVOID *)ppvObject);
	}
	else {
		return E_NOINTERFACE;
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
	return teGetDispId(methodWB, sizeof(methodWB), g_maps[MAP_WB], *rgszNames, rgDispId, true); 
}

STDMETHODIMP CteWebBrowser::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	switch(dispIdMember) {
		//Data
		case 0x10000001:
			if (nArg >= 0) {
				VariantClear(&m_Data);
				VariantCopy(&m_Data, &pDispParams->rgvarg[nArg]);
			}
			if (pVarResult) {
				VariantCopy(pVarResult, &m_Data);
			}
			return S_OK;
		//hwnd
		case 0x10000002:
			SetVariantLL(pVarResult, (LONGLONG)get_HWND());
			return S_OK;
		//Type
		case 0x10000003:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = CTRL_WB;
			}
			return S_OK;
		//TranslateAccelerator
		case 0x10000004:
			HRESULT hr;
			hr = E_NOTIMPL;
			if (nArg >= 3) {
				IOleInPlaceActiveObject *pActiveObject = NULL;
				if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pActiveObject))) {
					MSG msg;
					msg.hwnd = (HWND)GetLLFromVariant(&pDispParams->rgvarg[nArg]);
					msg.message = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
					msg.wParam = (WPARAM)GetLLFromVariant(&pDispParams->rgvarg[nArg - 2]);
					msg.lParam = (LPARAM)GetLLFromVariant(&pDispParams->rgvarg[nArg - 3]);
					hr = pActiveObject->TranslateAcceleratorW(&msg);
					pActiveObject->Release();
				}
			}
			if (pVarResult) {
				pVarResult->lVal = hr;
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//Application
		case 0x10000005:
			if (pVarResult) {
				IDispatch *pdisp;
				if SUCCEEDED(m_pWebBrowser->get_Application(&pdisp)) {
					teSetObject(pVarResult, pdisp);
					pdisp->Release();
				}
			}
			return S_OK;
		//Document
		case 0x10000006:
			if (pVarResult) {
				IDispatch *pdisp;
				if SUCCEEDED(m_pWebBrowser->get_Document(&pdisp)) {
					teSetObject(pVarResult, pdisp);
					pdisp->Release();
				}
			}
			return S_OK;
/*		//ShowBrowserBar
		case 0x10000007:
			if (nArg >= 2) {
				VARIANT v1, v2, v3;
				VariantInit(&v1);
				VariantInit(&v2);
				VariantInit(&v3);
				VariantCopy(&v1, &pDispParams->rgvarg[nArg]);
				VariantCopy(&v2, &pDispParams->rgvarg[nArg - 1]);
				VariantCopy(&v3, &pDispParams->rgvarg[nArg - 2]);
//				HRESULT hr = m_pWebBrowser->ShowBrowserBar(&pDispParams->rgvarg[nArg], &pDispParams->rgvarg[nArg - 1], &pDispParams->rgvarg[nArg - 2]);
				HRESULT hr = m_pWebBrowser->ShowBrowserBar(&v1, &v2, &v3);
				VariantClear(&v1);
				VariantClear(&v2);
				VariantClear(&v3);
			}
			return S_OK;*/
/*
//		::CoCreateInstance(CLSID_Explorer,NULL,CLSCTX_SERVER,IID_IWebBrowser2,(void**)&m_pWebBrowser);
		VARIANT v1, v2, v3;
		CLSID clsid;
		v1.bstrVal = SysAllocString(L"{EFA24E64-B078-11D0-89E4-00C04FC9E26E}");
		v1.vt = VT_BSTR;
//		CLSIDFromString(L"{EFA24E64-B078-11D0-89E4-00C04FC9E26E}", &clsid);
//		v1.vt = VT_CLSID;
//		v1.puuid = &clsid;
		v2.vt = VT_BOOL;
		v2.boolVal = VARIANT_TRUE;
		v3.vt = VT_I4;
		v3.lVal = 200;
		hr = m_pWebBrowser->Navigate(L"c:\\", NULL, NULL, NULL, NULL);
		m_pWebBrowser->ShowBrowserBar((VARIANT *)&v1, &v2, &v3);
*/
		case DISPID_STATUSTEXTCHANGE:
/*			if (nArg >= 0) {
				if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
					DoStatusText(this, (LPCWSTR)pDispParams->rgvarg[nArg].bstrVal, 0);
				}
			}*/
			return E_NOTIMPL;
/*		case DISPID_COMMANDSTATECHANGE:
			return E_NOTIMPL;
		case DISPID_DOWNLOADCOMPLETE:
			return S_OK;*/
		case DISPID_NAVIGATECOMPLETE:
			VARIANT vURL;
			teVariantChangeType(&vURL, &pDispParams->rgvarg[nArg - 1], VT_BSTR);		
			VariantClear(&vURL);
			return E_NOTIMPL;
		case DISPID_BEFORENAVIGATE2:
			if (nArg >= 6) {
				if (pDispParams->rgvarg[nArg - 6].vt == (VT_BYREF | VT_BOOL)) {
					VARIANT vURL;
					teVariantChangeType(&vURL, &pDispParams->rgvarg[nArg - 1], VT_BSTR);		
					for (int i = SysStringLen(vURL.bstrVal) - 1; i >= 0; i--) {
						if (vURL.bstrVal[i] == '/') {
							*pDispParams->rgvarg[nArg - 6].pboolVal = true;
							break;
						}
					}
					VariantClear(&vURL);
				}
			}
			break;
		case DISPID_DOCUMENTCOMPLETE:
			IOleWindow *pOleWindow;
			if SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
				pOleWindow->GetWindow(&g_hwndBrowser);
				pOleWindow->Release();
			}
			if (g_bInit) {
				g_bInit = false;
				SetTimer(g_hwndMain, TET_Create, 100, teTimerProc);
			}
			return S_OK;
		case DISPID_SECURITYDOMAIN:
			return S_OK;
		case DISPID_AMBIENT_DLCONTROL:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = DLCTL_DLIMAGES | DLCTL_VIDEOS | DLCTL_BGSOUNDS | DLCTL_FORCEOFFLINE | DLCTL_OFFLINE;
			}
			return S_OK;
		case DISPID_NEWWINDOW3:
			if (g_pOnFunc[TE_OnNewWindow]) {
				VARIANT vResult;
				VariantInit(&vResult);
				VARIANTARG *pv = GetNewVARIANT(4);
				teSetObject(&pv[3], g_pWebBrowser);
				for (int i = 2; i >= 0; i--) {
					VariantCopy(&pv[2 - i], &pDispParams->rgvarg[nArg - i - 2]);
				}
				Invoke4(g_pOnFunc[TE_OnNewWindow], &vResult, 4, pv);
				if (GetIntFromVariantClear(&vResult) == S_OK) {
					*pDispParams->rgvarg[nArg - 1].pboolVal = VARIANT_TRUE;
				}
			}
			return S_OK;
		case DISPID_NAVIGATEERROR:
			return S_OK;
		case DISPID_WINDOWCLOSING:
			return S_OK;
		case DISPID_FILEDOWNLOAD:
			if (nArg >= 1) {
				pDispParams->rgvarg[nArg - 1].vt = VT_BOOL;
				pDispParams->rgvarg[nArg - 1].boolVal = VARIANT_TRUE;
			}
			return S_OK;
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
	}
//	int i = dispIdMember;
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
	*phwnd = g_hwndMain;
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
	
	GetClientRect(g_hwndMain, lprcPosRect);
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
//IDocHostUIHandler
STDMETHODIMP CteWebBrowser::ShowContextMenu(DWORD dwID, POINT *ppt, IUnknown *pcmdtReserved, IDispatch *pdispReserved)
{
	if (g_pOnFunc[TE_OnShowContextMenu]) {
		MSG msg1;
		msg1.hwnd = m_hwnd;
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
	pInfo->dwFlags       = DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_SCROLL_NO/* | DOCHOSTUIFLAG_THEME*/;
	pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT;
	pInfo->pchHostCss    = NULL;
	pInfo->pchHostNS     = NULL;
	return S_OK;
}

STDMETHODIMP CteWebBrowser::ShowUI(DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc)
{
/*
IOleInPlaceUIWindow *pUIWindow;

	if SUCCEEDED(pCommandTarget->QueryInterface(IID_PPV_ARGS(&pUIWindow))) {
	}
*/
	return S_FALSE;//Resize Window
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
/*	if (!pchKey) {
		return E_INVALIDARG;
	}
	const LPWSTR szBuf = L"";
	int i = lstrlen(szBuf);
	*pchKey = reinterpret_cast<LPOLESTR>(::CoTaskMemAlloc((i + 1) * sizeof(OLECHAR)));
	lstrcpy(*pchKey, szBuf);*/
	return S_FALSE;
}

STDMETHODIMP CteWebBrowser::GetDropTarget(IDropTarget *pDropTarget, IDropTarget **ppDropTarget)
{
	return QueryInterface(IID_PPV_ARGS(ppDropTarget));
}

STDMETHODIMP CteWebBrowser::GetExternal(IDispatch **ppDispatch)
{
	return g_pTE->QueryInterface(IID_PPV_ARGS(ppDispatch));
}

STDMETHODIMP CteWebBrowser::TranslateUrl(DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut)
{
	return S_FALSE;
}

STDMETHODIMP CteWebBrowser::FilterDataObject(IDataObject *pDO, IDataObject **ppDORet)
{
	return E_NOTIMPL;
}

//IDropTarget
STDMETHODIMP CteWebBrowser::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);
	return DragSub(TE_OnDragEnter, this, m_pDragItems, grfKeyState, pt, pdwEffect);
}

STDMETHODIMP CteWebBrowser::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_grfKeyState = grfKeyState;
	return DragSub(TE_OnDragOver, this, m_pDragItems, grfKeyState, pt, pdwEffect);
}

STDMETHODIMP CteWebBrowser::DragLeave()
{
	if (m_DragLeave == E_NOT_SET) {
		m_DragLeave = DragLeaveSub(this, &m_pDragItems);
	}
	return m_DragLeave;
}

STDMETHODIMP CteWebBrowser::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	GetDragItems(&m_pDragItems, pDataObj);
	HRESULT hr = DragSub(TE_OnDrop, this, m_pDragItems, m_grfKeyState, pt, pdwEffect);
	DragLeave();
	return hr;
}

HWND CteWebBrowser::get_HWND()
{
	HWND hwnd = 0;
	IOleWindow *pOleWindow;
	if (m_pWebBrowser) {
		if SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
			pOleWindow->GetWindow(&hwnd);
			pOleWindow->Release();
		}
	}
	return hwnd;
}
/*
BOOL CteWebBrowser::IsBusy()
{
	VARIANT_BOOL vb;
	m_pWebBrowser->get_Busy(&vb);
	if (vb) {
		return vb;
	}
	READYSTATE rs;
	m_pWebBrowser->get_ReadyState(&rs);
	return rs < READYSTATE_COMPLETE;
}
*/
void CteWebBrowser::Close()
{
	IConnectionPointContainer *pCPC;
	if (m_pWebBrowser) {
		m_pWebBrowser->Quit();
		if SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pCPC))) {
			IConnectionPoint *pCP;
			if (SUCCEEDED(pCPC->FindConnectionPoint(DIID_DWebBrowserEvents2, &pCP))) {
				pCP->Unadvise(m_dwCookie);
				pCP->Release();
			}
			pCPC->Release();
		}
		IOleObject *pOleObject;
		if SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleObject))) {
			RECT rc;
			SetRectEmpty(&rc);
			pOleObject->DoVerb(OLEIVERB_HIDE, NULL, NULL, 0, g_hwndMain, &rc);
			pOleObject->Release();
		}
		for (int i = 3; i--;) {	//Error measures at the final
			m_pWebBrowser->AddRef();
		}
		PostMessage(get_HWND(), WM_CLOSE, 0, 0);
		m_pWebBrowser = NULL;
	}
}


// CteTabs

CteTabs::CteTabs()
{
	m_cRef = 1;
	m_nLockUpdate = 0;
	m_bVisible = FALSE;
	::ZeroMemory(&m_si, sizeof(m_si));
	m_si.cbSize = sizeof(m_si);
}

void CteTabs::Init()
{
	m_nIndex = -1;
	m_pDragItems = NULL;
	m_bEmpty = false;
	VariantInit(&m_Data);
	m_param[TE_Left] = 0;
	m_param[TE_Top] = 0;
	m_param[TE_Width] = 100;
	m_param[TE_Height] = 100;
	m_param[TE_Flags] = TCS_FOCUSNEVER | TCS_HOTTRACK | TCS_MULTILINE | TCS_RAGGEDRIGHT | TCS_SCROLLOPPOSITE | TCS_TOOLTIPS;
	m_param[TE_Align] = 0;
	m_param[TC_TabWidth] = 0;
	m_param[TC_TabHeight] = 0;
}

VOID CteTabs::GetItem(int i, VARIANT *pVarResult)
{
	if (pVarResult) {
		CteShellBrowser *pSB = GetShellBrowser(i);
		if (pSB) {
			teSetObject(pVarResult, pSB);
		}
	}
}

VOID CteTabs::Show(BOOL bVisible)
{
	if (bVisible ^ m_bVisible & 1) {
		m_bVisible = bVisible & 1;
		int nCmdShow = (bVisible ? SW_SHOW : SW_HIDE);
		if (bVisible) {
			ShowWindow(m_hwnd, nCmdShow);
			ShowWindow(m_hwndButton, nCmdShow);
			ShowWindow(m_hwndStatic, nCmdShow);
			g_pTabs = this;
			ArrangeWindow();
		}
		else {
			MoveWindow(m_hwndStatic, -1, -1, 0, 0, false);
			CteShellBrowser *pSB = GetShellBrowser(m_nIndex);
			if (pSB && pSB->m_bVisible) {
				pSB->Show(FALSE);
			}
		}
	}
	DoFunc(TE_OnVisibleChanged, this, E_NOTIMPL);
}

void CteTabs::Close(BOOL bForce)
{
	if (CanClose(this) || bForce) {
		int nCount = TabCtrl_GetItemCount(m_hwnd); 
		while (nCount--) {
			CteShellBrowser *pSB = GetShellBrowser(0);
			if (pSB) {
				pSB->Close(true);
			}
		}
		Show(FALSE);
		RevokeDragDrop(m_hwnd);
		SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)g_DefTCProc);
		DestroyWindow(m_hwnd);
		m_hwnd = NULL;

		RevokeDragDrop(m_hwndButton);
		SetWindowLongPtr(m_hwndButton, GWLP_WNDPROC, (LONG_PTR)g_DefBTProc);
		DestroyWindow(m_hwndButton);
		m_hwndButton = NULL;

		DestroyWindow(m_hwndStatic);
		m_hwndStatic = NULL;
		m_bEmpty = true;
		if (this == g_pTabs) {
			g_pTabs = NULL;
			int i = MAX_TC;
			while (--i >= 0) {
				if (g_pTC[i]) {
					if (!g_pTC[i]->m_bEmpty) {
						g_pTabs =  g_pTC[i];
						return;
					}
				}
				else {
					return;
				}
			}
		}
	}
}

DWORD CteTabs::GetStyle()
{
	DWORD dwStyle = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | TCS_FOCUSNEVER | m_param[TE_Flags];
	if (dwStyle & TCS_BUTTONS) {
		if (dwStyle & TCS_SCROLLOPPOSITE) {
			dwStyle &= ~TCS_SCROLLOPPOSITE;
		}
		if (dwStyle & TCS_BOTTOM && m_param[TE_Align] > 1) {
			dwStyle &= ~TCS_BOTTOM;
		}
	}
	return dwStyle;
}

VOID CteTabs::CreateTC()
{
	m_hwnd = CreateWindowEx(
		WS_EX_CONTROLPARENT, //Extended style
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
	ArrangeWindow();
	DWORD dwSize = MAKELPARAM(m_param[TC_TabWidth], m_param[TC_TabHeight]);
	SendMessage(m_hwnd, TCM_SETITEMSIZE, 0, dwSize);
	m_dwSize = dwSize;

	g_DefTCProc = (WNDPROC)SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)TETCProc);
	RegisterDragDrop(m_hwnd, static_cast<IDropTarget *>(this));
	BringWindowToTop(m_hwnd);
}

BOOL CteTabs::Create()
{
	m_hwndStatic = CreateWindowEx(
		0, L"STATIC", NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN /*| BS_OWNERDRAW*/,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 
		g_hwndMain, (HMENU)0, hInst, NULL); 
	BringWindowToTop(m_hwndStatic);

	m_hwndButton = CreateWindowEx(
		0, L"BUTTON", NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | BS_OWNERDRAW, // | WS_VSCROLL,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 
		m_hwndStatic, (HMENU)0, hInst, NULL);

	g_DefBTProc = (WNDPROC)SetWindowLongPtr(m_hwndButton, GWLP_WNDPROC, (LONG_PTR)TEBTProc);
	RegisterDragDrop(m_hwndButton, static_cast<IDropTarget *>(this));
	BringWindowToTop(m_hwndButton);

	CreateTC();
	DoFunc(TE_OnCreate, this, E_NOTIMPL);
	DoFunc(TE_OnViewCreated, this, E_NOTIMPL);
	ArrangeWindow();
	return true;
}

CteTabs::~CteTabs()
{
	Close(true);
	VariantClear(&m_Data);
}

STDMETHODIMP CteTabs::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else if (IsEqualIID(riid, IID_IDispatchEx)) {
		*ppvObject = static_cast<IDispatchEx *>(this);
	}
	else if (IsEqualIID(riid, IID_IDropTarget)) {
		*ppvObject = static_cast<IDropTarget *>(this);
	}
	else if (IsEqualIID(riid, g_ClsIdTC)) {
		*ppvObject = this;
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
}

STDMETHODIMP_(ULONG) CteTabs::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteTabs::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteTabs::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteTabs::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTabs::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	return teGetDispId(methodTC, sizeof(methodTC), g_maps[MAP_TC], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteTabs::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	IUnknown *punk = NULL;

	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	switch(dispIdMember) {
		//Data
		case 1:
			if (nArg >= 0) {
				VariantClear(&m_Data);
				VariantCopy(&m_Data, &pDispParams->rgvarg[nArg]);
			}
			if (pVarResult) {
				VariantCopy(pVarResult, &m_Data);
			}
			return S_OK;
		//hwnd
		case 2:
			SetVariantLL(pVarResult, (LONGLONG)m_hwnd);
			return S_OK;
		//Type
		case 3:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = CTRL_TC;
			}
			return S_OK;
		//HitTest (Screen coordinates)
		case 6:	
			if (nArg >= 1 && pVarResult) {
				TCHITTESTINFO info;
				GetPointFormVariant(&info.pt, &pDispParams->rgvarg[nArg]);
				UINT flags = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
				pVarResult->lVal = static_cast<int>(DoHitTest(this, info.pt, flags));
				if (pVarResult->lVal < 0) {
					ScreenToClient(m_hwnd, &info.pt);
					info.flags = flags;
					pVarResult->lVal = TabCtrl_HitTest(m_hwnd, &info);
					if (!(info.flags & flags)) {
						pVarResult->lVal = -1;
					}
				}
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//Move
		case 7:	
			if (nArg >= 1) {
				int nSrc, nDest;
				nDest = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
				CteShellBrowser *pSB;
				if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
					if SUCCEEDED(punk->QueryInterface(g_ClsIdSB, (LPVOID *)&pSB)) {
						pSB->m_pTabs->Move(pSB->GetTabIndex(), nDest, this);
						pSB->Release();
						return S_OK;
					}
				}
				nSrc = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				if (nArg >= 2) {
					if (FindUnknown(&pDispParams->rgvarg[nArg - 2], &punk)) {
						CteTabs *pTC;
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
		//Selected
		case 8:
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
		//Close
		case 9:
			Close(false);
			return S_OK;
		//SelectedIndex
		case 10:
			if (nArg >= 0) {
				TabCtrl_SetCurSel(m_hwnd, GetIntFromVariant(&pDispParams->rgvarg[nArg]));
			}
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = TabCtrl_GetCurSel(m_hwnd);
			}
			return S_OK;
		//Visible
		case 11:
			if (nArg >= 0) {
				Show(GetIntFromVariant(&pDispParams->rgvarg[nArg])); 
			}
			if (pVarResult) {
				pVarResult->vt = VT_BOOL;
				pVarResult->boolVal = m_bVisible;
			}
			return S_OK;
		//Id
		case 12:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = MAX_TC - m_nTC;
			}
			return S_OK;
		//Item
		case DISPID_TE_ITEM:
			if (nArg >= 0) {
				GetItem(GetIntFromVariant(&pDispParams->rgvarg[nArg]), pVarResult);
			}
			return S_OK;
		//Count
		case DISPID_TE_COUNT:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = TabCtrl_GetItemCount(m_hwnd);
			}
			return S_OK;
		//_NewEnum
		case DISPID_NEWENUM:
			if (pVarResult) {
				CteDispatch *pid = new CteDispatch(this, 0);
				teSetObject(pVarResult, pid);
				pid->Release();
				return S_OK;
			}
			return E_FAIL;
		//this
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
	}
	if (dispIdMember >= TE_OFFSET && dispIdMember <= TE_OFFSET + TC_TabHeight) {
		if (nArg >= 0 && dispIdMember != TE_OFFSET) {
			m_param[dispIdMember - TE_OFFSET] = GetIntFromVariantPP(&pDispParams->rgvarg[nArg], dispIdMember - TE_OFFSET);
			if (m_hwnd) {
				if (dispIdMember == TE_OFFSET + TE_Flags) {
					DWORD dwStyle = GetWindowLong(m_hwnd, GWL_STYLE);
					if ((dwStyle ^ m_param[TE_Flags]) & 0x7fff) {
						if ((dwStyle ^ m_param[TE_Flags]) & (TCS_SCROLLOPPOSITE | TCS_BUTTONS | TCS_TOOLTIPS)) {
							//Remake
							int nTabs = TabCtrl_GetItemCount(m_hwnd);
							int nSel = TabCtrl_GetCurFocus(m_hwnd);
							TC_ITEM *tabs = new TC_ITEM[nTabs];
							LPWSTR *ppszText = new LPWSTR[nTabs];
							for (int i = nTabs - 1; i >= 0; i--) {
								ppszText[i] = new WCHAR[MAX_PATH];
								tabs[i].cchTextMax = MAX_PATH;
								tabs[i].pszText = ppszText[i];
								tabs[i].mask = TCIF_IMAGE | TCIF_PARAM | TCIF_TEXT;
								TabCtrl_GetItem(m_hwnd, i, &tabs[i]);
							}
							HIMAGELIST hImage = TabCtrl_GetImageList(m_hwnd);
							HFONT hFont = (HFONT)SendMessage(m_hwnd, WM_GETFONT, 0, 0);
							RevokeDragDrop(m_hwnd);
							SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)g_DefTCProc);
							DestroyWindow(m_hwnd);

							CreateTC();
							SendMessage(m_hwnd, WM_SETFONT, (WPARAM)hFont, 0);
							TabCtrl_SetImageList(m_hwnd, hImage);
							for (int i = 0; i < nTabs; i++) {
								TabCtrl_InsertItem(m_hwnd, i, &tabs[i]);
								delete [] ppszText[i];
							}
							delete [] ppszText;
							delete [] tabs;
							TabCtrl_SetCurSel(m_hwnd, nSel);
							DoFunc(TE_OnViewCreated, this, E_NOTIMPL);
						}
					}
				}
				else if (dispIdMember == TE_OFFSET + TC_TabWidth || dispIdMember == TE_OFFSET + TC_TabHeight) {
					DWORD dwSize = MAKELPARAM(m_param[TC_TabWidth], m_param[TC_TabHeight]);
					if (dwSize != m_dwSize) {
						SendMessage(m_hwnd, TCM_SETITEMSIZE, 0, dwSize);
						m_dwSize = dwSize;
					}
				}
				SetWindowLong(m_hwnd, GWL_STYLE, GetStyle());
				ArrangeWindow();
			}
		}
		if (pVarResult) {
			int i = m_param[dispIdMember - TE_OFFSET];
			if (dispIdMember >= TE_OFFSET + TE_Left && dispIdMember <= TE_OFFSET + TE_Height) {
				if (i & 0x3fff) {															
					VARIANT v, vs;
					vs.vt = VT_R4;
					vs.fltVal = ((float)(m_param[dispIdMember - TE_OFFSET])) / 100;
					teVariantChangeType(&v, &vs, VT_BSTR);
					pVarResult->vt = VT_BSTR;
					pVarResult->bstrVal = SysAllocStringLen(NULL, SysStringLen(v.bstrVal) + 1);
					lstrcpy(pVarResult->bstrVal, v.bstrVal);
					lstrcat(pVarResult->bstrVal, L"%");
					VariantClear(&v);
				}
				else {
					pVarResult->vt = VT_I4;
					pVarResult->lVal = i / 0x4000;
				}
			}
			else {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = i;
			}
		}
		return S_OK;
	}
	if (dispIdMember >= DISPID_COLLECTION_MIN && dispIdMember <= DISPID_TE_MAX) {
		GetItem(dispIdMember - DISPID_COLLECTION_MIN, pVarResult);
		return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteTabs::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	return teGetDispId(methodTC, sizeof(methodTC), g_maps[MAP_TC], bstrName, pid, true);
}

STDMETHODIMP CteTabs::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller)
{
	return Invoke(id, IID_NULL, lcid, wFlags, pdp, pvarRes, pei, NULL);
}

STDMETHODIMP CteTabs::DeleteMemberByName(BSTR bstrName, DWORD grfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTabs::DeleteMemberByDispID(DISPID id)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTabs::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTabs::GetMemberName(DISPID id, BSTR *pbstrName)
{
	return teGetMemberName(id, pbstrName);
}

STDMETHODIMP CteTabs::GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid)
{
	*pid = (id < DISPID_COLLECTION_MIN) ? DISPID_COLLECTION_MIN : id + 1;
	return *pid < TabCtrl_GetItemCount(m_hwnd) + DISPID_COLLECTION_MIN ? S_OK : S_FALSE;
}

STDMETHODIMP CteTabs::GetNameSpaceParent(IUnknown **ppunk)
{
	return E_NOTIMPL;
}

VOID CteTabs::Move(int nSrc, int nDest, CteTabs *pDestTab)
{
	TC_ITEM tcItem;
	if (nDest < 0) {
		nDest = TabCtrl_GetItemCount(m_hwnd) - 1;
		if (nDest < 0) {
			nDest = 0;
		}
	}
	BOOL bOtherTab = m_hwnd != pDestTab->m_hwnd;
	if (nSrc != nDest || bOtherTab) {
		int nIndex = TabCtrl_GetCurSel(m_hwnd);
		WCHAR *pszText = NULL;
		::ZeroMemory(&tcItem, sizeof(tcItem));
		tcItem.mask = TCIF_TEXT | TCIF_IMAGE | TCIF_PARAM;
		tcItem.cchTextMax = MAX_PATH;
		pszText = new WCHAR[MAX_PATH];
		tcItem.pszText = pszText;
		TabCtrl_GetItem(m_hwnd, nSrc, &tcItem);
		SendMessage(m_hwnd, TCM_DELETEITEM, nSrc, 1);	//SB is not delete!
		CteShellBrowser *pSB = g_pSB[tcItem.lParam];
		if (pSB) {
			teSetParent(pSB->m_hwnd, pDestTab->m_hwndStatic);
			pSB->m_pTabs = pDestTab;
			TabCtrl_InsertItem(pDestTab->m_hwnd, nDest, &tcItem);
			if (nSrc == nIndex) {
				m_nIndex = -1;
				if (bOtherTab) {
					if (nSrc > 0) {
						nSrc--;
					}
					TabCtrl_SetCurSel(m_hwnd, nSrc);
				}
				else {
					TabCtrl_SetCurSel(m_hwnd, nDest);
					TabChanged(true);
				}
			}
		}
		delete [] pszText;
		if (bOtherTab) {
			pDestTab->m_nIndex = nDest;
			TabCtrl_SetCurSel(pDestTab->m_hwnd, nDest);
			ArrangeWindow();
		}
	}
}

VOID CteTabs::LockUpdate()
{
	if (InterlockedIncrement(&m_nLockUpdate) == 1) {
		SendMessage(m_hwndStatic, WM_SETREDRAW, FALSE, 0);
		SendMessage(g_hwndBrowser, WM_SETREDRAW, FALSE, 0);
	}
}

VOID CteTabs::UnLockUpdate()
{
	if (::InterlockedDecrement(&m_nLockUpdate) == 0) {
		m_bRedraw = true;
		KillTimer(g_hwndMain, TET_Redraw);
		SetTimer(g_hwndMain, TET_Redraw, 1000, teTimerProc);
	}
}

VOID CteTabs::RedrawUpdate()
{
	if (m_bRedraw && g_nLockUpdate == 0) {
		m_bRedraw = false;
		SendMessage(m_hwndStatic, WM_SETREDRAW, TRUE, 0);
		SendMessage(g_hwndBrowser, WM_SETREDRAW, TRUE, 0);
		CteShellBrowser *pSB = GetShellBrowser(m_nIndex);
		if (pSB) {
			pSB->SetRedraw(TRUE);
		}
		RedrawWindow(m_hwndStatic, NULL, 0, RDW_INVALIDATE | RDW_ALLCHILDREN);
		RedrawWindow(g_hwndBrowser, NULL, 0, RDW_INVALIDATE | RDW_ALLCHILDREN);
	}
}

//IDropTarget
STDMETHODIMP CteTabs::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);
	return DragSub(TE_OnDragEnter, this, m_pDragItems, grfKeyState, pt, pdwEffect);
}

STDMETHODIMP CteTabs::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_grfKeyState = grfKeyState;
	return DragSub(TE_OnDragOver, this, m_pDragItems, grfKeyState, pt, pdwEffect);
}

STDMETHODIMP CteTabs::DragLeave()
{
	if (m_DragLeave == E_NOT_SET) {
		m_DragLeave = DragLeaveSub(this, &m_pDragItems);
	}
	return m_DragLeave;
}

STDMETHODIMP CteTabs::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	GetDragItems(&m_pDragItems, pDataObj);
	HRESULT hr = DragSub(TE_OnDrop, this, m_pDragItems, m_grfKeyState, pt, pdwEffect);
	DragLeave();
	return hr;
}

VOID CteTabs::TabChanged(BOOL bSameTC)
{
	int i;

	CteShellBrowser *pHide = NULL;
	if (bSameTC) {
		i = TabCtrl_GetCurSel(m_hwnd);
		if (i != m_nIndex) {
			pHide = GetShellBrowser(m_nIndex);
		}
		m_nIndex = i;
	}
	CteShellBrowser *pSB = GetShellBrowser(m_nIndex);
	if (pSB) {
		pSB->Show(true);
		if (pHide) {
			pHide->Show(false);
		}
		DoFunc(TE_OnSelectionChanged, this, E_NOTIMPL);
	}
	else if (pHide) {
		pHide->Show(false);
	}
}

CteShellBrowser* CteTabs::GetShellBrowser(int nPage)
{
	if (nPage < 0) {
		return NULL;
	}
	TC_ITEM tcItem;
	tcItem.mask = TCIF_PARAM;
	TabCtrl_GetItem(m_hwnd, nPage, &tcItem);
	if (tcItem.lParam < _countof(g_pSB)) {
		return g_pSB[tcItem.lParam];
	}
	return NULL;
}


// CteFolderItems

CteFolderItems::CteFolderItems(IDataObject *pDataObj, FolderItems *pFolderItems, BOOL bEditable)
{
	m_cRef = 1;
	m_pDataObj = pDataObj;
	if (m_pDataObj) {
		m_pDataObj->AddRef();
	}
	m_pidllist = NULL;
	m_nCount = -1;
	m_oFolderItems = NULL;
	if (bEditable) {
		GetNewArray(&m_oFolderItems);
	}
	m_pFolderItems = pFolderItems;
	m_nIndex = 0;
	m_dwEffect = (DWORD)-1;
	m_dwTick = 0;
	m_bUseILF = true;
}

CteFolderItems::~CteFolderItems()
{
	Clear();
}

VOID CteFolderItems::Clear()
{
	if (m_pidllist) {
		for (int i = m_nCount; i >= 0; i--) {
			::CoTaskMemFree(m_pidllist[i]);
		}
		delete [] m_pidllist;
		m_pidllist = NULL;
	}
	if (m_oFolderItems) {
		m_oFolderItems->Release();
		m_oFolderItems = NULL;
	}
	if (m_pDataObj) {
		m_pDataObj->Release();
		m_pDataObj = NULL;
	}
	if (m_pFolderItems) {
		m_pFolderItems->Release();
		m_pFolderItems = NULL;
	}
}

VOID CteFolderItems::ItemEx(int nIndex, VARIANT *pVarResult, VARIANT *pVarNew)
{
	VARIANT v;
	v.lVal = nIndex;
	v.vt = VT_I4;
	if (pVarNew) {
		DISPID dispid;
		if (!m_oFolderItems) {
			long nCount;
			get_Count(&nCount);
			IDispatch *oFolderItems;
			GetNewArray(&oFolderItems);
			VARIANT v1, v2;
			v1.vt = VT_I4;
			FolderItem *pid;
			LPOLESTR sz = L"push";
			for (v1.lVal = 0; v1.lVal < nCount; v1.lVal++) {
				if (Item(v1, &pid) == S_OK) {
					teSetObject(&v2, pid);
					pid->Release();
					oFolderItems->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &dispid);
					Invoke5(oFolderItems, dispid, DISPATCH_METHOD, NULL, -1, &v2);
					VariantClear(&v2);
				}
			}
			Clear();
			m_oFolderItems = oFolderItems;
		}
		if (m_oFolderItems) {
			VARIANT v1;
			teVariantChangeType(&v1, &v, VT_BSTR);
			WORD wFlags = DISPATCH_PROPERTYPUT;
			m_oFolderItems->GetIDsOfNames(IID_NULL, &v1.bstrVal, 1, LOCALE_USER_DEFAULT, &dispid);
			if (dispid < 0) {
				::SysReAllocString(&v1.bstrVal, L"push");
				m_oFolderItems->GetIDsOfNames(IID_NULL, &v1.bstrVal, 1, LOCALE_USER_DEFAULT, &dispid);
				wFlags = DISPATCH_METHOD;
			}
			VariantClear(&v1);
			VARIANTARG *pv = GetNewVARIANT(1);
			pv[0].punkVal = NULL;
			IUnknown *punk;
			CteFolderItem *pid = NULL;
			if (FindUnknown(pVarNew, &punk)) {
				punk->QueryInterface(g_ClsIdFI, (LPVOID *)&pid);
			}
			if (pid == NULL) {
				pid = new CteFolderItem(pVarNew);
			}
			teSetObject(&pv[0], pid);
			pid->Release();
			Invoke5(m_oFolderItems, dispid, wFlags, NULL, 1, pv);
		}
	}
	if (pVarResult) {
		FolderItem *pid = NULL;
		if SUCCEEDED(Item(v, &pid)) {
			teSetObject(pVarResult, pid);
		}
	}
}

VOID CteFolderItems::AdjustIDListEx()
{
	if (!m_pidllist && m_oFolderItems) {
		get_Count(&m_nCount);
		m_pidllist = new LPITEMIDLIST[m_nCount + 1];
		m_pidllist[0] = NULL;
		VARIANT v;
		v.lVal = m_nCount;
		v.vt = VT_I4;
		FolderItem *pid = NULL;
		while (--v.lVal >= 0) {
			if (Item(v, &pid) == S_OK) {
				GetIDListFromObjectEx(pid, &m_pidllist[v.lVal + 1]);
				pid->Release();
			}
		}
		m_oFolderItems->Release();
		m_oFolderItems = NULL;
	}
	m_bUseILF = AdjustIDList(m_pidllist, m_nCount);
}

STDMETHODIMP CteFolderItems::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = reinterpret_cast<IDispatch *>(this);
	}
	else if (IsEqualIID(riid, IID_FolderItems)) {
		*ppvObject = static_cast<FolderItems *>(this);
	}
	else if (IsEqualIID(riid, IID_IDispatchEx)) {
		*ppvObject = static_cast<IDispatchEx *>(this);
	}
	else if (IsEqualIID(riid, IID_IDataObject)) {
		if (!m_pDataObj && m_pidllist) {
			m_bUseILF = AdjustIDList(m_pidllist, m_nCount);
			IShellFolder *pSF;
			if (GetShellFolder(&pSF, m_pidllist[0])) {
				pSF->GetUIObjectOf(g_hwndMain, m_nCount, (LPCITEMIDLIST *)&m_pidllist[1], IID_IDataObject, NULL, (LPVOID *)&m_pDataObj);
				pSF->Release();
			}
		}
		if (m_pDataObj) {
			if (m_pDataObj->QueryGetData(&UNICODEFormat) == S_OK) {
				return m_pDataObj->QueryInterface(riid, ppvObject);
			}
		}
		*ppvObject = static_cast<IDataObject *>(this);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
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
	return teGetDispId(methodFIs, sizeof(methodFIs), g_maps[MAP_FIs], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteFolderItems::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	if (dispIdMember >= DISPID_COLLECTION_MIN && dispIdMember <= DISPID_TE_MAX) {
		ItemEx(dispIdMember - DISPID_COLLECTION_MIN, pVarResult, nArg >= 0 ? &pDispParams->rgvarg[nArg] : NULL);
		return S_OK;
	}
	if (m_pFolderItems) {
		return m_pFolderItems->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	}
	switch(dispIdMember) {
		//Application
		case 2:
			if (pVarResult) {
				IDispatch *pdisp;
				if SUCCEEDED(get_Application(&pdisp)) {
					teSetObject(pVarResult, pdisp);
					pdisp->Release();
				}
			}
			return S_OK;
		//Parent
		case 3:
			return S_OK;
		//Item
		case DISPID_TE_ITEM:
			if (nArg >= 0) {
				ItemEx(GetIntFromVariant(&pDispParams->rgvarg[nArg]), pVarResult, nArg >= 1 ? &pDispParams->rgvarg[nArg - 1] : NULL);
			}
			return S_OK;
		//Count
		case DISPID_TE_COUNT:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				get_Count(&pVarResult->lVal);
			}
			return S_OK;
		//_NewEnum
		case DISPID_NEWENUM:
			if (pVarResult) {
				IUnknown *punk;
				if SUCCEEDED(_NewEnum(&punk)) {
					teSetObject(pVarResult, punk);
					punk->Release();
				}
				return S_OK;
			}
			return E_FAIL;
		//AddItem
		case 8:
			if (nArg >= 0) {
				ItemEx(-1, NULL, &pDispParams->rgvarg[nArg]);
			}
			return S_OK;
		//hDrop
		case 9:
			if (pVarResult) {
				int x = (nArg >= 0) ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : 0;
				int y = (nArg >= 1) ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : 0;
				int fNC = (nArg >= 2) ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]) : FALSE;
				int bSp = (nArg >= 3) ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 3]) : FALSE;
				SetVariantLL(pVarResult, (LONGLONG)GethDrop(x, y, fNC, bSp));
			}
			return S_OK;
		//Index
		case DISPID_TE_INDEX:
			if (nArg >= 0) {
				m_nIndex = GetIntFromVariant(&pDispParams->rgvarg[nArg]); 
			}
			if (pVarResult) {
				pVarResult->lVal = m_nIndex;
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//dwEffect
		case 0x10000001:
			if (nArg >= 0) {
				m_dwEffect = GetIntFromVariant(&pDispParams->rgvarg[nArg]); 
			}
			if (pVarResult) {
				pVarResult->lVal = m_dwEffect;
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//pdwEffect
		case 0x10000002:
			if (pVarResult) {
				CteMemory *pMem = new CteMemory(sizeof(int), (char *)&m_dwEffect, VT_NULL, 1, L"DWORD");
				teSetObject(pVarResult, pMem);
				pMem->Release();
			}
			return S_OK;
		//this
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

STDMETHODIMP CteFolderItems::get_Count(long *plCount)
{
	if (m_pFolderItems) {
		return m_pFolderItems->get_Count(plCount);
	}
	if (m_oFolderItems) {
		VARIANT v;
		VariantInit(&v);
		HRESULT hr = tePropertyGet(m_oFolderItems, L"length", &v);
		*plCount = GetIntFromVariantClear(&v);
		return hr;
	}
	if (m_nCount < 0) {
		if (m_pDataObj) {
			STGMEDIUM Medium;
			if (m_pDataObj->GetData(&IDLISTFormat, &Medium) == S_OK) {
				CIDA *pIDA = (CIDA *)GlobalLock(Medium.hGlobal);
				m_nCount = pIDA ? pIDA->cidl : 0;
				GlobalUnlock(Medium.hGlobal);
				ReleaseStgMedium(&Medium);
			}
			else if (m_pDataObj->GetData(&HDROPFormat, &Medium) == S_OK) {
				m_nCount = DragQueryFile((HDROP)Medium.hGlobal, (UINT)-1, NULL, 0);
				ReleaseStgMedium(&Medium);
			}
		}
		else {
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
	return g_pWebBrowser->m_pWebBrowser->get_Application(ppid);
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
			VARIANT v;
			VariantInit(&v);
			index.lVal = i;
			index.vt = VT_I4;
			teVariantChangeType(&v, &index, VT_BSTR);
			tePropertyGet(m_oFolderItems, v.bstrVal, &v);
			IUnknown *punk;
			if (FindUnknown(&v, &punk)) {
				HRESULT hr = punk->QueryInterface(IID_PPV_ARGS(ppid));
				VariantClear(&v);
				return hr;
			}
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
			GetFolderItemFromPidl(ppid, pidl);
			::CoTaskMemFree(pidl);
		}
		else {
			GetFolderItemFromPidl(ppid, m_pidllist[i + 1]);
		}
	}
	else if (i == -1) {
		if (m_pidllist[0] == NULL) {
			AdjustIDListEx();
		}
		if (m_pidllist[0]) {
			GetFolderItemFromPidl(ppid, m_pidllist[0]);
		}
	}
	return (*ppid != NULL) ? S_OK : E_FAIL;
}

STDMETHODIMP CteFolderItems::_NewEnum(IUnknown **ppunk)
{
	if (m_pFolderItems) {
		return m_pFolderItems->_NewEnum(ppunk);
	}
	*ppunk = reinterpret_cast<IUnknown *>(new CteDispatch(reinterpret_cast<IDispatch *>(this), 0));
	return S_OK;
}

//IDispatchEx
STDMETHODIMP CteFolderItems::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	return teGetDispId(methodFIs, sizeof(methodFIs), g_maps[MAP_FIs], bstrName, pid, true);
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

HDROP CteFolderItems::GethDrop(int x, int y, BOOL fNC, BOOL bSpecial)
{
	if (m_pDataObj) {
		STGMEDIUM Medium;
		if (m_pDataObj->GetData(&HDROPFormat, &Medium) == S_OK) {
			HDROP hDrop = (HDROP)Medium.hGlobal;
			if (fNC) {
				LPDROPFILES lpDropFiles = (LPDROPFILES)GlobalLock(hDrop);
				try {
					lpDropFiles->pt.x = x;
					lpDropFiles->pt.y = y;
					lpDropFiles->fNC = fNC;
				}
				catch (...) {
				}
				GlobalUnlock(hDrop);
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
	BSTR *pbslist = new BSTR[m_nCount];
	UINT uSize = sizeof(WCHAR);
	for (int i = m_nCount - 1; i >= 0; i--) {
		LPITEMIDLIST pidl = ILCombine(m_pidllist[0], m_pidllist[i + 1]);
		if (bSpecial) {
			GetDisplayNameFromPidl(&pbslist[i], pidl, SHGDN_FORPARSING);
			int j = SysStringByteLen(pbslist[i]);
			if (j) {
				uSize += j + sizeof(WCHAR);
			}
			else {
				::SysFreeString(pbslist[i]);
				pbslist[i] = NULL;
			}
		}
		else {
			BSTR bsPath;
			pbslist[i] = NULL;
			if SUCCEEDED(GetDisplayNameFromPidl(&bsPath, pidl, SHGDN_FORPARSING)) {
				int nLen = ::SysStringByteLen(bsPath);
				if (nLen) {
					pbslist[i] = bsPath;
					uSize += nLen + sizeof(WCHAR);
				}
				else {
					::SysFreeString(bsPath);
				}
			}
		}
		::CoTaskMemFree(pidl);
	}
	HDROP hDrop = (HDROP)GlobalAlloc(GHND, sizeof(DROPFILES) + uSize);
	LPDROPFILES lpDropFiles = (LPDROPFILES)GlobalLock(hDrop);
	try {
		lpDropFiles->pFiles = sizeof(DROPFILES);
		lpDropFiles->pt.x = x;
		lpDropFiles->pt.y = y;
		lpDropFiles->fNC = fNC;
		lpDropFiles->fWide = TRUE;

		LPWSTR lp = (LPWSTR)&lpDropFiles[1];
		for (int i = 0; i< m_nCount; i++) {
			if (pbslist[i]) {
				lstrcpy(lp, pbslist[i]);
				lp += (SysStringByteLen(pbslist[i]) / 2) + 1;
				::SysFreeString(pbslist[i]);
			}
		}
		*lp = 0;
		delete [] pbslist;
	}
	catch (...) {
	}
	GlobalUnlock(hDrop);
	return hDrop;
}

//IDataObject
STDMETHODIMP CteFolderItems::GetData(FORMATETC *pformatetcIn, STGMEDIUM *pmedium)
{
	if (m_dwEffect != (DWORD)-1 && pformatetcIn->cfFormat == DROPEFFECTFormat.cfFormat) {
		HGLOBAL hGlobal = GlobalAlloc(GHND | GMEM_SHARE, sizeof(DWORD));
		DWORD *pdwEffect = (DWORD *)GlobalLock(hGlobal);
		try {
			if (pdwEffect) {
				*pdwEffect = m_dwEffect;
			}
		} catch (...) {
		}
		GlobalUnlock(hGlobal);

		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = hGlobal;
		pmedium->pUnkForRelease = NULL;
		return S_OK;
	}
	if (!m_dwTick || GetTickCount() - m_dwTick > 16) {
		if (m_pDataObj) {
			HRESULT hr = m_pDataObj->GetData(pformatetcIn, pmedium);
			if (hr == S_OK) {
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
			for (int i = 0; i <= m_nCount; i++) {
				uSize += ILGetSize(m_pidllist[i]);
			}
			HGLOBAL hGlobal = GlobalAlloc(GHND, uSize);
			CIDA *pIDA = (CIDA *)GlobalLock(hGlobal);
			try {
				if (pIDA) {
					pIDA->cidl = m_nCount;
					for (int i = 0; i <= m_nCount; i++) {
						pIDA->aoffset[i] = uIndex;
						UINT u = ILGetSize(m_pidllist[i]);
						char *pc = (char *)pIDA + uIndex;
						::CopyMemory(pc, m_pidllist[i], u);
						uIndex += u;
					}
				}
			}
			catch (...) {
			}
			GlobalUnlock(hGlobal);
			pmedium->tymed = TYMED_HGLOBAL;
			pmedium->hGlobal = hGlobal;
			pmedium->pUnkForRelease = NULL;
			return S_OK;
		}
	}
	if (pformatetcIn->cfFormat == CF_HDROP) {
		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = (HGLOBAL)GethDrop(0, 0, FALSE, FALSE);
		pmedium->pUnkForRelease = NULL;
		return S_OK;
	}

	if (pformatetcIn->cfFormat == CF_UNICODETEXT) {
		BSTR bsText = NULL;
		try {
			if (g_pOnFunc[TE_OnClipboardText]) {
				VARIANT vResult, vStr;
				VariantInit(&vResult);
				VARIANTARG *pv = GetNewVARIANT(1);
				teSetObject(&pv[0], this);
				Invoke4(g_pOnFunc[TE_OnClipboardText], &vResult, 1, pv);
				teVariantChangeType(&vStr, &vResult, VT_BSTR);
				VariantClear(&vResult);
				bsText = vStr.bstrVal;
				vStr.bstrVal = NULL;
			}
		}
		catch (...) {
		}
		HGLOBAL hGlobal = GlobalAlloc(GHND, SysStringByteLen(bsText) + sizeof(WCHAR));
		LPWSTR lp = static_cast<LPWSTR>(GlobalLock(hGlobal));
		try {
			if (lp) {
				lstrcpy(lp, bsText);
			}
		} catch (...) {
		}
		GlobalUnlock(hGlobal);
		::SysFreeString(bsText);

		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = hGlobal;
		pmedium->pUnkForRelease = NULL;
		m_dwTick = GetTickCount();
		return S_OK;
	}
	return DATA_E_FORMATETC;
}

STDMETHODIMP CteFolderItems::GetDataHere(FORMATETC *pformatetc, STGMEDIUM *pmedium)
{
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
	if (pformatetc->cfFormat == CF_UNICODETEXT) {
		return S_OK;
	}
	if (pformatetc->cfFormat == CF_HDROP) {
		return S_OK;
	}
	if (pformatetc->cfFormat == IDLISTFormat.cfFormat && m_bUseILF) {
		return S_OK;
	}
	return DATA_E_FORMATETC;
}

STDMETHODIMP CteFolderItems::GetCanonicalFormatEtc(FORMATETC *pformatectIn, FORMATETC *pformatetcOut)
{
	if (m_pDataObj) {
		return m_pDataObj->GetCanonicalFormatEtc(pformatectIn, pformatetcOut);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::SetData(FORMATETC *pformatetc, STGMEDIUM *pmedium, BOOL fRelease)
{
	if (m_pDataObj) {
		return m_pDataObj->SetData(pformatetc, pmedium, fRelease);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenumFormatEtc)
{
	if (dwDirection == DATADIR_GET) {
		FORMATETC formats[MAX_FORMATS];
		FORMATETC teformats[] = {  IDLISTFormat, HDROPFormat, UNICODEFormat, DROPEFFECTFormat };
		UINT nFormat = 0;

		if (m_pDataObj) {
			IEnumFORMATETC *penumFormatEtc;
			if SUCCEEDED(m_pDataObj->EnumFormatEtc(DATADIR_GET, &penumFormatEtc)) {
				while (nFormat < MAX_FORMATS - _countof(teformats) && penumFormatEtc->Next(1, &formats[nFormat], NULL) == S_OK) {
					if (QueryGetData2(&formats[nFormat]) != S_OK) {
						nFormat++;
					}
				}
				penumFormatEtc->Release();
			}
		}
		AdjustIDListEx();
		int nMax = _countof(teformats);
		if (m_dwEffect == (DWORD)-1) {
			nMax--;
		}
		for (int i = m_bUseILF ? 0 : 1; i < nMax; i++) {
			formats[nFormat++] = teformats[i];
		}
		return CreateFormatEnumerator(nFormat, formats, ppenumFormatEtc);
	}
	if (m_pDataObj) {
		return m_pDataObj->EnumFormatEtc(dwDirection, ppenumFormatEtc);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItems::DAdvise(FORMATETC *pformatetc, DWORD advf, IAdviseSink *pAdvSink, DWORD *pdwConnection)
{
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
	if (m_pDataObj) {
		return m_pDataObj->EnumDAdvise(ppenumAdvise);
	}
	return E_NOTIMPL;
}

//CteMemory

CteMemory::CteMemory(int nSize, char *pc, int nMode, int nCount, LPWSTR lpStruct)
{
	m_cRef = 1;
	m_pc = pc;
	if (lpStruct) {
		m_bsStruct = ::SysAllocString(lpStruct);
	}
	else {
		m_bsStruct = NULL;
	}
	m_nMode = nMode;
	m_nCount = nCount;
	m_nAlloc = 0;
	m_nSize = 0;
	if (nSize > 0) {
		m_nSize = nSize;
		if (pc == NULL) {
			if (nMode == VT_VARIANT) {
				m_pc = (char *)GetNewVARIANT(nCount);
				if (m_pc) {
					m_nAlloc = 3;
				}
			}
			else {
				m_pc = new char[nSize];
				if (m_pc) {
					::ZeroMemory(m_pc, nSize);
					m_nAlloc = 1;
				}
			}
		}
		if (pc == (char *)1) {
			m_pc = (char *)CoTaskMemAlloc(nSize);
			if (m_pc) {
				::ZeroMemory(m_pc, nSize);
				m_nAlloc = 2;
			}
		}
	}
	m_nbs = 0;
	m_ppbs = NULL;
}


CteMemory::~CteMemory()
{
	Free();
}

void CteMemory::Free()
{
	switch (m_nMode) {
		case VT_BSTR:
			if (m_nCount > 0) {
				BSTR *lpBSTRData = (BSTR*)m_pc;
				while (--m_nCount >= 0) {
					if (lpBSTRData[m_nCount]) {
						::SysFreeString(lpBSTRData[m_nCount]);
					}
				}
			}
			break;
		case VT_DISPATCH:
		case VT_UNKNOWN:
			if (m_nCount > 0) {
				IUnknown **ppunk = (IUnknown **)m_pc;
				while (--m_nCount >= 0) {
					if (ppunk[m_nCount]) {
						ppunk[m_nCount]->Release();
					}
				}
			}
			break;
	}

	switch(m_nAlloc) {
		case 1:
			delete [] m_pc;
			break;
		case 2:
			::CoTaskMemFree((LPVOID)m_pc);
			break;
		case 3:
			VARIANT *pv = (VARIANT *)m_pc;
			while (--m_nCount >= 0) {
				VariantInit(&pv[m_nCount]);
			}
			delete [] pv;
			break;
	}
	if (m_bsStruct) {
		::SysFreeString(m_bsStruct);
	}
	if (m_ppbs) {
		while (--m_nbs >= 0) {
			::SysFreeString(m_ppbs[m_nbs]);
		}
		delete [] m_ppbs;
		m_ppbs = NULL;
	}
	m_nAlloc = 0;
	m_pc = NULL;
}

STDMETHODIMP CteMemory::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else if (IsEqualIID(riid, IID_IDispatchEx)) {
		*ppvObject = static_cast<IDispatchEx *>(this);
	}
	else if (IsEqualIID(riid, g_ClsIdStruct)) {
		*ppvObject = this;
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
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
	return GetDispID(*rgszNames, 0, rgDispId);
}

STDMETHODIMP CteMemory::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	switch(dispIdMember) {
		//P
		case 1:	
			SetVariantLL(pVarResult, (LONGLONG)m_pc);
			return S_OK;
		//Item
		case DISPID_TE_ITEM:
			if (nArg >= 0) {
				ItemEx(GetIntFromVariant(&pDispParams->rgvarg[nArg]), pVarResult, nArg >= 1 ? &pDispParams->rgvarg[nArg - 1] : NULL);
			}
			return S_OK;
		//Count
		case DISPID_TE_COUNT:
			if (pVarResult) {
				pVarResult->lVal = m_nCount;
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//_NewEnum
		case DISPID_NEWENUM:
			if (pVarResult) {
				CteDispatch *pid = new CteDispatch(this, 0);
				teSetObject(pVarResult, pid);
				pid->Release();
				return S_OK;
			}
			return E_FAIL;
		//Read
		case 4:
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
		case 5:
			if (nArg >= 2) {
				int nLen = -1;
				if (nArg >= 3) {
					nLen = GetIntFromVariant(&pDispParams->rgvarg[nArg] - 3);
				}
				Write(GetIntFromVariant(&pDispParams->rgvarg[nArg]), nLen, GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), &pDispParams->rgvarg[nArg - 2]);
			}
			return S_OK;
		//Size
		case 6:	
			if (pVarResult) {
				pVarResult->lVal = m_nSize;
				pVarResult->vt = VT_I4;
			}
		return S_OK;
		//Free
		case 7:
			Free();
			return S_OK;
		//Clone
		case 8:
			if (pVarResult) {
				if (m_nSize) {
					CteMemory *pmem = new CteMemory(m_nSize, NULL, m_nMode, m_nCount, m_bsStruct);
					::CopyMemory(pmem->m_pc, m_pc, m_nSize);
					teSetObject(pVarResult, pmem);
					pmem->Release();
				}
			}
			return S_OK;
		//Value
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
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
						if (pVarResult) {
							pVarResult->vt = VT_I4;
							pVarResult->lVal = dispIdMember & TE_VI;
						}
					}
					else {
						if (nArg >= 0) {
							if (vt == VT_BSTR && pDispParams->rgvarg[nArg].vt == VT_BSTR) {
								VARIANT v;
								VariantInit(&v);
								v.llVal = 0;
								v.bstrVal = AddBSTR(pDispParams->rgvarg[nArg].bstrVal);
								v.vt = VT_UI8;
								Write(dispIdMember & TE_VI, -1, vt, &v);
							}
							else {
								Write(dispIdMember & TE_VI, -1, vt, &pDispParams->rgvarg[nArg]);
							}
						}
						if (pVarResult) {
							pVarResult->vt = vt;
							Read(dispIdMember & TE_VI, -1, pVarResult);
						}
					}
				}
				return S_OK;
			}
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteMemory::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	if (lstrcmpi(m_bsStruct, L"POINT") == 0) {
		POINT *p = NULL;
		if (lstrcmpi(bstrName, L"x") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->x;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"y") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->y;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"RECT") == 0) {
		RECT *p = NULL;
		if (lstrcmpi(bstrName, L"left") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->left;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"top") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->top;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"right") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->right;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"bottom") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->bottom;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"BITMAP") == 0) {
		BITMAP *p = NULL;
		if (lstrcmpi(bstrName, L"bmType") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->bmType;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"bmWidth") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->bmWidth;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"bmHeight") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->bmHeight;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"bmWidthBytes") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->bmWidthBytes;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"bmPlanes") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->bmPlanes;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"bmBitsPixel") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->bmBitsPixel;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"bmBits") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->bmBits;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"CHOOSECOLOR") == 0) {
		CHOOSECOLOR *p = NULL;
		if (lstrcmpi(bstrName, L"lStructSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->lStructSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hwndOwner") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hwndOwner;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hInstance") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hInstance;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"rgbResult") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->rgbResult;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpCustColors") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpCustColors;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"Flags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->Flags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lCustData") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lCustData;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpfnHook") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpfnHook;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpTemplateName") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpTemplateName;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"CHOOSEFONT") == 0) {
		CHOOSEFONT *p = NULL;
		if (lstrcmpi(bstrName, L"lStructSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->lStructSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hwndOwner") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hwndOwner;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hDC") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hDC;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpLogFont") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpLogFont;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iPointSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iPointSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"Flags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->Flags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"rgbColors") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->rgbColors;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lCustData") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lCustData;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpfnHook") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpfnHook;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpTemplateName") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpTemplateName;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hInstance") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hInstance;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpszStyle") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpszStyle;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"nFontType") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->nFontType;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"___MISSING_ALIGNMENT__") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->___MISSING_ALIGNMENT__;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"nSizeMin") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->nSizeMin;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"nSizeMax") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->nSizeMax;
			return S_OK;
		}
	}

	if (lstrcmpi(m_bsStruct, L"COPYDATASTRUCT") == 0) {
		COPYDATASTRUCT *p = NULL;
 		if (lstrcmpi(bstrName, L"dwData") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->dwData;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cbData") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cbData;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpData") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpData;
			return S_OK;
		}
	}

	if (lstrcmpi(m_bsStruct, L"EXCEPINFO") == 0) {
		EXCEPINFO *p = NULL;
		if (lstrcmpi(bstrName, L"wCode") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->wCode;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wReserved") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->wReserved;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"bstrSource") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->bstrSource;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"bstrDescription") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->bstrDescription;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"bstrHelpFile") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->bstrHelpFile;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwHelpContext") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwHelpContext;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"pvReserved") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->pvReserved;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"pfnDeferredFillIn") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->pfnDeferredFillIn;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"scode") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->scode;
			return S_OK;
		}
	}

	if (lstrcmpi(m_bsStruct, L"FINDREPLACE") == 0) {
		FINDREPLACE *p = NULL;
		if (lstrcmpi(bstrName, L"lStructSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->lStructSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hwndOwner") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hwndOwner;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hInstance") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hInstance;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"Flags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->Flags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpstrFindWhat") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpstrFindWhat;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpstrReplaceWith") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpstrReplaceWith;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wFindWhatLen") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->wFindWhatLen;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wReplaceWithLen") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->wReplaceWithLen;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lCustData") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lCustData;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpfnHook") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpfnHook;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpTemplateName") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpTemplateName;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"DIBSECTION") == 0) {
		DIBSECTION *p = NULL;
		if (lstrcmpi(bstrName, L"dsBm") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->dsBm;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dsBmih") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->dsBmih;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dsBitfields0") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dsBitfields[0];
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dsBitfields1") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dsBitfields[1];
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dsBitfields2") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dsBitfields[2];
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dshSection") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->dshSection;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dsOffset") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dsOffset;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"EncoderParameter") == 0) {
		EncoderParameter *p = NULL;
		if (lstrcmpi(bstrName, L"Guid") == 0) {
			*pid = (VT_CLSID << TE_VT) + (DWORD)(UINT_PTR)&p->Guid;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"NumberOfValues") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->NumberOfValues;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"Type") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->Type;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"Value") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->Value;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"EncoderParameters") == 0) {
		EncoderParameters *p = NULL;
		if (lstrcmpi(bstrName, L"Count") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->Count;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"Parameter") == 0) {
			*pid = DISPID_TE_ITEM;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"FOLDERSETTINGS") == 0) {
		FOLDERSETTINGS *p = NULL;
		if (lstrcmpi(bstrName, L"ViewMode") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->ViewMode;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"fFlags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->fFlags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"Options") == 0) {
			*pid = (VT_I4 << TE_VT) + (SB_Options - 1) * 4;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"ViewFlags") == 0) {
			*pid = (VT_I4 << TE_VT) + (SB_ViewFlags - 1) * 4;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"ImageSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (SB_IconSize - 1) * 4;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"ICONINFO") == 0) {
		ICONINFO *p = NULL;
		if (lstrcmpi(bstrName, L"fIcon") == 0) {
			*pid = (VT_BOOL << TE_VT) + (DWORD)(UINT_PTR)&p->fIcon;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"xHotspot") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->xHotspot;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"yHotspot") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->yHotspot;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hbmMask") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hbmMask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hbmColor") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hbmColor;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"ICONMETRICS") == 0) {
		ICONMETRICS *p = NULL;
		if (lstrcmpi(bstrName, L"cbSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cbSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iHorzSpacing") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iHorzSpacing;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iVertSpacing") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iVertSpacing;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iTitleWrap") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iTitleWrap;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfFont") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lfFont;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"KEYBDINPUT") == 0) {
		KEYBDINPUT *p = (KEYBDINPUT *)(char*)sizeof(DWORD);
		if (lstrcmpi(bstrName, L"type") == 0) {
			*pid = (VT_I4 << TE_VT);
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wVk") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->wVk;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wScan") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->wScan;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwFlags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwFlags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"time") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->time;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwExtraInfo") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->dwExtraInfo;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"LOGFONT") == 0) {
		LOGFONT *p = NULL;
		if (lstrcmpi(bstrName, L"lfHeight") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->lfHeight;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfWidth") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->lfWidth;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfEscapement") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->lfEscapement;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfOrientation") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->lfOrientation;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfWeight") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->lfWeight;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfItalic") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->lfItalic;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfUnderline") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->lfUnderline;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfStrikeOut") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->lfStrikeOut;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfCharSet") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->lfCharSet;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfOutPrecision") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->lfOutPrecision;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfClipPrecision") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->lfClipPrecision;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfQuality") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->lfQuality;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfPitchAndFamily") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->lfPitchAndFamily;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfFaceName") == 0) {
			*pid = (VT_LPWSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lfFaceName;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"LVBKIMAGE") == 0) {
		LVBKIMAGE *p = NULL;
		if (lstrcmpi(bstrName, L"ulFlags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->ulFlags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hbm") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hbm;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"pszImage") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->pszImage;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cchImageMax") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cchImageMax;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"xOffsetPercent") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->xOffsetPercent;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"yOffsetPercent") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->yOffsetPercent;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"LVFINDINFO") == 0) {
		LVFINDINFO *p = NULL;
		if (lstrcmpi(bstrName, L"flags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->flags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"psz") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->psz;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lParam") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->lParam;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"pt") == 0) {
			*pid = (VT_CY << TE_VT) + (DWORD)(UINT_PTR)&p->pt;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"vkDirection") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->vkDirection;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"LVHITTESTINFO") == 0) {
		LVHITTESTINFO *p = NULL;
		if (lstrcmpi(bstrName, L"pt") == 0) {
			*pid = (VT_CY << TE_VT) + (DWORD)(UINT_PTR)&p->pt;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"flags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->flags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iItem") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iItem;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iSubItem") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iSubItem;
			return S_OK;
		} 
		if (lstrcmpi(bstrName, L"iGroup") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iGroup;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"LVITEM") == 0) {
		LVITEM *p = NULL;
		if (lstrcmpi(bstrName, L"mask") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->mask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iItem") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iItem;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iSubItem") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iSubItem;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"state") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->state;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"stateMask") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->stateMask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"pszText") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->pszText;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cchTextMax") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cchTextMax;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iImage") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iImage;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lParam") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->lParam;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iIndent") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iIndent;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iGroupId") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iGroupId;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cColumns") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cColumns;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"puColumns") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->puColumns;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"piColFmt") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->piColFmt;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iGroup") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iGroup;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"MENUITEMINFO") == 0) {
		MENUITEMINFO *p = NULL;
		if (lstrcmpi(bstrName, L"cbSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cbSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"fMask") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->fMask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"fType") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->fType;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"fState") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->fState;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wID") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->wID;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hSubMenu") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hSubMenu;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hbmpChecked") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hbmpChecked;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hbmpUnchecked") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hbmpUnchecked;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwItemData") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->dwItemData;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwTypeData") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->dwTypeData;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cch") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cch;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hbmpItem") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hbmpItem;
			return S_OK;
		}
	}

	if (lstrcmpi(m_bsStruct, L"MOUSEINPUT") == 0) {
		MOUSEINPUT *p = (MOUSEINPUT *)(char*)sizeof(DWORD);
		if (lstrcmpi(bstrName, L"type") == 0) {
			*pid = (VT_I4 << TE_VT);
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dx") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dx;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dy") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dy;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"mouseData") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->mouseData;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwFlags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwFlags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"time") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->time;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwExtraInfo") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->dwExtraInfo;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"MSG") == 0) {
		MSG *p = NULL;
		if (lstrcmpi(bstrName, L"hwnd") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hwnd;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"message") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->message;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wParam") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->wParam;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lParam") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lParam;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"time") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->time;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"pt") == 0) {
			*pid = (VT_CY << TE_VT) + (DWORD)(UINT_PTR)&p->pt;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"NONCLIENTMETRICS") == 0) {
		NONCLIENTMETRICS *p = NULL;
		if (lstrcmpi(bstrName, L"cbSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cbSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iBorderWidth") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iBorderWidth;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iScrollWidth") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iScrollWidth;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iScrollHeight") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iScrollHeight;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iCaptionWidth") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iCaptionWidth;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iCaptionHeight") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iCaptionHeight;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfCaptionFont") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lfCaptionFont;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iSmCaptionWidth") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iSmCaptionWidth;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iSmCaptionHeight") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iSmCaptionHeight;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfSmCaptionFont") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lfSmCaptionFont;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iMenuWidth") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iMenuWidth;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iMenuHeight") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iMenuHeight;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfMenuFont") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lfMenuFont;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfStatusFont") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lfStatusFont;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lfMessageFont") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lfMessageFont;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iPaddedBorderWidth") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iPaddedBorderWidth;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"NOTIFYICONDATA") == 0) {
		NOTIFYICONDATA *p = NULL;
		if (lstrcmpi(bstrName, L"cbSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cbSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hWnd") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hWnd;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"uID") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->uID;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"uFlags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->uFlags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"uCallbackMessage") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->uCallbackMessage;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hIcon") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hIcon;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"szTip") == 0) {
			*pid = (VT_LPWSTR << TE_VT) + (DWORD)(UINT_PTR)&p->szTip;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwState") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwState;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwState") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwState;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwStateMask") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwStateMask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"szInfo") == 0) {
			*pid = (VT_LPWSTR << TE_VT) + (DWORD)(UINT_PTR)&p->szInfo;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"uTimeout") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->uTimeout;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"uVersion") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->uVersion;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"szInfoTitle") == 0) {
			*pid = (VT_LPWSTR << TE_VT) + (DWORD)(UINT_PTR)&p->szInfoTitle;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwInfoFlags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwInfoFlags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"guidItem") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->guidItem;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hBalloonIcon") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hBalloonIcon;
			return S_OK;
		}
	}

	if (lstrcmpi(m_bsStruct, L"NMHDR") == 0) {
		NMHDR *p = NULL;
		if (lstrcmpi(bstrName, L"hwndFrom") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hwndFrom;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"idFrom") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->idFrom;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"code") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->code;
			return S_OK;
		}
	}

	if (lstrcmpi(m_bsStruct, L"OSVERSIONINFO") == 0 || lstrcmpi(m_bsStruct, L"OSVERSIONINFOEX") == 0) {
		OSVERSIONINFOEX *p = NULL;
		if (lstrcmpi(bstrName, L"dwOSVersionInfoSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwOSVersionInfoSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwMajorVersion") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwMajorVersion;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwMinorVersion") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwMinorVersion;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwBuildNumber") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwBuildNumber;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwPlatformId") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwPlatformId;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"szCSDVersion") == 0) {
			*pid = (VT_LPWSTR << TE_VT) + (DWORD)(UINT_PTR)&p->szCSDVersion;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wServicePackMajor") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->wServicePackMajor;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wServicePackMinor") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->wServicePackMinor;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wSuiteMask") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->wSuiteMask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wProductType") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->wProductType;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wReserved") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->wReserved;
			return S_OK;
		}
	}

	if (lstrcmpi(m_bsStruct, L"PAINTSTRUCT") == 0) {
		PAINTSTRUCT *p = NULL;
		if (lstrcmpi(bstrName, L"hdc") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hdc;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"fErase") == 0) {
			*pid = (VT_BOOL << TE_VT) + (DWORD)(UINT_PTR)&p->fErase;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"rcPaint") == 0) {
			*pid = (VT_CARRAY << TE_VT) + (DWORD)(UINT_PTR)&p->rcPaint;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"fRestore") == 0) {
			*pid = (VT_BOOL << TE_VT) + (DWORD)(UINT_PTR)&p->fRestore;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"fIncUpdate") == 0) {
			*pid = (VT_BOOL << TE_VT) + (DWORD)(UINT_PTR)&p->fIncUpdate;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"rgbReserved") == 0) {
			*pid = (VT_UI1 << TE_VT) + (DWORD)(UINT_PTR)&p->rgbReserved;
			return S_OK;
		}
	}

	if (lstrcmpi(m_bsStruct, L"SHELLEXECUTEINFO") == 0) {
		SHELLEXECUTEINFO *p = NULL;
		if (lstrcmpi(bstrName, L"cbSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cbSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"fMask") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->fMask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hwnd") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hwnd;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpVerb") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpVerb;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpFile") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpFile;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpParameters") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpParameters;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpDirectory") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpDirectory;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"nShow") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->nShow;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hInstApp") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hInstApp;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpIDList") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpIDList;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpClass") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpClass;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hkeyClass") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hkeyClass;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwHotKey") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwHotKey;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hIcon") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hIcon;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hMonitor") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hMonitor;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hProcess") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hProcess;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"SHFILEINFO") == 0) {
		SHFILEINFO *p = NULL;
		if (lstrcmpi(bstrName, L"hIcon") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hIcon;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iIcon") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iIcon;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwAttributes") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwAttributes;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"szDisplayName") == 0) {
			*pid = (VT_LPWSTR << TE_VT) + (DWORD)(UINT_PTR)&p->szDisplayName;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"szTypeName") == 0) {
			*pid = (VT_LPWSTR << TE_VT) + (DWORD)(UINT_PTR)&p->szTypeName;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"SHFILEOPSTRUCT") == 0) {
		SHFILEOPSTRUCT *p = NULL;
		if (lstrcmpi(bstrName, L"hwnd") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hwnd;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"wFunc") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->wFunc;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"pFrom") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->pFrom;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"pTo") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->pTo;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"fFlags") == 0) {
			*pid = (VT_UI2 << TE_VT) + (DWORD)(UINT_PTR)&p->fFlags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"fAnyOperationsAborted") == 0) {
			*pid = (VT_BOOL << TE_VT) + (DWORD)(UINT_PTR)&p->fAnyOperationsAborted;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hNameMappings") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hNameMappings;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lpszProgressTitle") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->lpszProgressTitle;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"SIZE") == 0) {
		SIZE *p = NULL;
		if (lstrcmpi(bstrName, L"cx") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cx;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cy") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cy;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"TCITEM") == 0) {
		TCITEM *p = NULL;
		if (lstrcmpi(bstrName, L"mask") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->mask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwState") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwState;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwStateMask") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwStateMask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"pszText") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->pszText;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cchTextMax") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cchTextMax;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iImage") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iImage;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lParam") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lParam;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"TVHITTESTINFO") == 0) {
		TVHITTESTINFO *p = NULL;
		if (lstrcmpi(bstrName, L"pt") == 0) {
			*pid = (VT_CY << TE_VT) + (DWORD)(UINT_PTR)&p->pt;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"flags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->flags;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hItem") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hItem;
			return S_OK;
		}
	}
	if (lstrcmpi(m_bsStruct, L"TVITEM") == 0) {
		TVITEM *p = NULL;
		if (lstrcmpi(bstrName, L"mask") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->mask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"hItem") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->hItem;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"state") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->state;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"stateMask") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->stateMask;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"pszText") == 0) {
			*pid = (VT_BSTR << TE_VT) + (DWORD)(UINT_PTR)&p->pszText;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cchTextMax") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cchTextMax;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iImage") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iImage;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"iSelectedImage") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->iSelectedImage;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cChildren") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cChildren;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"lParam") == 0) {
			*pid = (VT_PTR << TE_VT) + (DWORD)(UINT_PTR)&p->lParam;
			return S_OK;
		}
	}

	if (lstrcmpi(m_bsStruct, L"TCHITTESTINFO") == 0) {
		TCHITTESTINFO *p = NULL;
		if (lstrcmpi(bstrName, L"pt") == 0) {
			*pid = (VT_CY << TE_VT) + (DWORD)(UINT_PTR)&p->pt;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"flags") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->flags;
			return S_OK;
		}
	}

	if (lstrcmpi(m_bsStruct, L"WIN32_FIND_DATA") == 0) {
		WIN32_FIND_DATA *p = NULL;
		if (lstrcmpi(bstrName, L"dwFileAttributes") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwFileAttributes;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"ftCreationTime") == 0) {
			*pid = (VT_FILETIME << TE_VT) + (DWORD)(UINT_PTR)&p->ftCreationTime;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"ftLastAccessTime") == 0) {
			*pid = (VT_FILETIME << TE_VT) + (DWORD)(UINT_PTR)&p->ftLastAccessTime;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"ftLastWriteTime") == 0) {
			*pid = (VT_FILETIME << TE_VT) + (DWORD)(UINT_PTR)&p->ftLastWriteTime;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"nFileSizeHigh") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->nFileSizeHigh;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"nFileSizeLow") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->nFileSizeLow;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwReserved0") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwReserved0;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"dwReserved1") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->dwReserved1;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cFileName") == 0) {
			*pid = (VT_LPWSTR << TE_VT) + (DWORD)(UINT_PTR)&p->cFileName;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"cAlternateFileName") == 0) {
			*pid = (VT_LPWSTR << TE_VT) + (DWORD)(UINT_PTR)&p->cAlternateFileName;
			return S_OK;
		}
	}
	/*
	if (lstrcmpi(m_bsStruct, L"") == 0) {
		? *p = NULL;
		if (lstrcmpi(bstrName, L"cbSize") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->cbSize;
			return S_OK;
		}
		if (lstrcmpi(bstrName, L"") == 0) {
			*pid = (VT_I4 << TE_VT) + (DWORD)(UINT_PTR)&p->;
			return S_OK;
		}
	}
	*/
	if (teGetDispId(methodMem, sizeof(methodMem), g_maps[MAP_Mem], bstrName, pid, true) == S_OK) {
		return S_OK;
	}
	return teGetDispId(methodMem2, sizeof(methodMem2), g_maps[MAP_M2], m_bsStruct, pid, false);
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
			if (m_nCount) {
				if (m_nMode == VT_VARIANT) {
					VARIANT *p = (VARIANT *)m_pc;
					VariantCopy(&p[i], pVarNew);
				}
				else {
					int nSize = m_nSize / m_nCount;
					if (nSize >= 1 && nSize <= 8) {
						LARGE_INTEGER ll;
						ll.QuadPart = GetLLFromVariant(pVarNew);
						if (ll.QuadPart == 0 && pVarNew->vt == VT_BSTR && pVarNew->bstrVal &&
						(lstrcmpi(m_bsStruct, L"BSTR") == 0 || lstrcmpi(m_bsStruct, L"LPWSTR") == 0)) {
							ll.QuadPart = (LONGLONG)AddBSTR(pVarNew->bstrVal);
						}
						if (nSize == 1) {//BYTE
							m_pc[i] = ll.LowPart & 0xff;
						}
						else if (nSize == 2) {//WORD
							WORD *p = (WORD *)m_pc;
							p[i] = LOWORD(ll.LowPart);
						}
						else if (nSize == 4) {//DWORD
							int *p = (int *)m_pc;
							p[i] = ll.LowPart;
						}
						else if (nSize == 8) {//LONGLONG
							LONGLONG *p = (LONGLONG *)m_pc;
							p[i] = ll.QuadPart;
						}
					}
				}
			}
		}
		if (pVarResult) {
			switch (m_nMode) {
				case VT_BSTR:
					pVarResult->bstrVal = SysAllocString(((BSTR*)m_pc)[i]);
					pVarResult->vt = VT_BSTR;
					break;
				case VT_UNKNOWN:
				case VT_DISPATCH:
					IUnknown **ppUnk;
					ppUnk = (IUnknown **)m_pc;
					if (ppUnk[i]) {
						teSetObject(pVarResult, ppUnk[i]);
					}
					break;
				case VT_VARIANT:
					VARIANT *p;
					p = (VARIANT *)m_pc;
					VariantCopy(pVarResult, &p[i]);
					break;
				default:
					if (m_nCount) {
						int nSize = m_nSize / m_nCount;
						if (nSize) {
							if (nSize == 1) {//BYTE/char
								pVarResult->lVal = m_pc[i];
								pVarResult->vt = VT_I4;
							}
							else if (nSize == 2) {//WORD/WCHAR
								WORD *p = (WORD *)m_pc;
								pVarResult->lVal = p[i];
								pVarResult->vt = VT_I4;
							}
							else if (nSize == 4) {//DWORD/int
								int *p = (int *)m_pc;
								pVarResult->lVal = p[i];
								pVarResult->vt = VT_I4;
							}
							else if (nSize == 8 && lstrcmpi(m_bsStruct, L"POINT")) {//LONGLONG
								LONGLONG *p = (LONGLONG *)m_pc;
								SetVariantLL(pVarResult, p[i]);
							}
							else {
								int nOffset = GetOffsetOfStruct(m_bsStruct);
								nSize -= nOffset;
								CteMemory *pmem = new CteMemory(nSize, m_pc + nOffset + (nSize * i), m_nMode, 1, GetSubStruct(m_bsStruct));
								teSetObject(pVarResult, pmem);
								pmem->Release();
							}
						}
					}
					break;
			}
		}
	}
}


BSTR CteMemory::AddBSTR(BSTR bs)
{
	BSTR *p = new BSTR[m_nbs + 1];
	if (m_ppbs) {
		::CopyMemory(p, m_ppbs, m_nbs * sizeof(BSTR));
		delete [] m_ppbs;
	}
	m_ppbs = p;
	return m_ppbs[m_nbs++] = ::SysAllocStringByteLen(reinterpret_cast<char*>(bs), ::SysStringByteLen(bs));
}

VOID CteMemory::Read(int nIndex, int nLen, VARIANT *pVarResult)
{
	if (pVarResult->vt & VT_ARRAY) {
		VARTYPE vt = pVarResult->vt & 0xff;
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
	switch (pVarResult->vt) {
		case VT_BOOL:
			pVarResult->boolVal = *(BOOL *)&m_pc[nIndex];
			break;
		case VT_I1:
		case VT_UI1:
			pVarResult->bVal = m_pc[nIndex];
			break;
		case VT_I2:
		case VT_UI2:
			pVarResult->iVal = *(short *)&m_pc[nIndex];
			break;
		case VT_I4:
		case VT_UI4:
			pVarResult->lVal = *(int *)&m_pc[nIndex];
			break;
		case VT_I8:
		case VT_UI8:
			SetVariantLL(pVarResult, *(LONGLONG *)&m_pc[nIndex]);
			break;
		case VT_R4:
			pVarResult->fltVal = *(FLOAT *)&m_pc[nIndex];
			break;
		case VT_R8:
			pVarResult->dblVal = *(DOUBLE *)&m_pc[nIndex];
			break;
		case VT_PTR:
		case VT_BSTR:
			SetVariantLL(pVarResult, *(INT_PTR *)&m_pc[nIndex]);
			break;
/*			if (nLen < 0) {
				nLen = m_nSize - nIndex;
			}
			pVarResult->vt = VT_BSTR;
			pVarResult->bstrVal = SysAllocStringByteLen(NULL, nLen);
			::CopyMemory(pVarResult->bstrVal, &m_pc[nIndex], nLen);
			break;*/
		case VT_LPWSTR:
			if (nLen >= 0) {
				if (nLen * (int)sizeof(WCHAR) > m_nSize) {
					nLen = m_nSize / sizeof(WCHAR);
				}
				pVarResult->bstrVal = SysAllocStringLenEx((WCHAR *)&m_pc[nIndex], nLen, lstrlen((WCHAR *)&m_pc[nIndex]));
			}
			else {
				pVarResult->bstrVal = SysAllocString((WCHAR *)&m_pc[nIndex]);
			}
			pVarResult->vt = VT_BSTR;
			break;
		case VT_LPSTR:
			int nLenW;
			if (nLen > m_nSize) {
				nLen = m_nSize;
			}
			nLenW = MultiByteToWideChar(CP_ACP, 0, (LPCSTR)&m_pc[nIndex], nLen, NULL, NULL);
			LPWSTR lpszW;
			lpszW = new WCHAR[nLenW];
			MultiByteToWideChar(CP_ACP, 0, (LPCSTR)&m_pc[nIndex], nLen, lpszW, nLenW);
			pVarResult->bstrVal = SysAllocStringLen(lpszW, nLenW - 1);
			pVarResult->vt = VT_BSTR;
			delete []lpszW;
			break;
		case VT_FILETIME:
			teFileTimeToVariantTime((LPFILETIME)&m_pc[nIndex], &pVarResult->date);
			pVarResult->vt = VT_DATE;
			break;
		case VT_CY:
			CteMemory *pstPt;
			pstPt = new CteMemory(sizeof(POINT), NULL, VT_NULL, 1, L"POINT");
			::CopyMemory(pstPt->m_pc, &m_pc[nIndex], 2 * sizeof(int));	
			teSetObject(pVarResult, pstPt);
			pstPt->Release();
			break;
		case VT_CARRAY:
			CteMemory *pstR;
			pstR = new CteMemory(sizeof(RECT), NULL, VT_NULL, 1, L"RECT");
			::CopyMemory(pstR->m_pc, &m_pc[nIndex], 4 * sizeof(int));	
			teSetObject(pVarResult, pstR);
			pstR->Release();
			break;
		case VT_CLSID:
			LPOLESTR lpsz;
			StringFromCLSID(*(IID *)&m_pc[nIndex], &lpsz);
			pVarResult->bstrVal = SysAllocString(lpsz);
			::CoTaskMemFree(lpsz);
			pVarResult->vt = VT_BSTR;
			break;
		case VT_VARIANT:
			VariantCopy(pVarResult, (VARIANT *)&m_pc[nIndex]);
//			*pVarResult = *(VARIANT *)&m_pc[nIndex];
			break;
		default:
			pVarResult->vt = VT_EMPTY;
			break;
	}
}

//Write
VOID CteMemory::Write(int nIndex, int nLen, VARTYPE vt, VARIANT *pv)
{
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
	LONGLONG ll;
	ll = GetLLFromVariant(pv);
	switch (vt) {
		case VT_I1:
		case VT_UI1:
			::CopyMemory(&m_pc[nIndex], &ll, sizeof(char));
			break;
		case VT_I2:
		case VT_UI2:
			::CopyMemory(&m_pc[nIndex], &ll, sizeof(short));
			break;
		case VT_BOOL:
		case VT_I4:
		case VT_UI4:
			::CopyMemory(&m_pc[nIndex], &ll, sizeof(int));
			break;
		case VT_I8:
		case VT_UI8:
			::CopyMemory(&m_pc[nIndex], &ll, sizeof(LONGLONG));
			break;
		case VT_PTR:
		case VT_BSTR:
			::CopyMemory(&m_pc[nIndex], &ll, sizeof(INT_PTR));
			break;
/*
			teVariantChangeType(&v, pv, VT_BSTR);
			if (v.bstrVal) {
				if (nLen < 0) {
					nLen = SysStringByteLen(v.bstrVal) + 2;
				}
				if (nLen > m_nSize) {
					nLen = m_nSize;
				}
				::CopyMemory(&m_pc[nIndex], v.bstrVal, nLen);
			}
			else {
				*(WCHAR *)&m_pc[nIndex] = NULL;
			}
			VariantClear(&v);
			break;*/
		case VT_LPWSTR:
			VARIANT v;
			teVariantChangeType(&v, pv, VT_BSTR);
			if (v.bstrVal) {
				if (nLen < 0) {
					nLen = (SysStringLen(v.bstrVal) + 1);
				}
				nLen *= sizeof(WCHAR);
				if (nLen > m_nSize) {
					nLen = m_nSize;
				}
				::CopyMemory(&m_pc[nIndex], v.bstrVal, nLen);
			}
			else {
				*(WCHAR *)&m_pc[nIndex] = NULL;
			}
			VariantClear(&v);
			break;
		case VT_LPSTR:
			teVariantChangeType(&v, pv, VT_BSTR);
			int nLenA;
			nLenA = 0;
			if (v.bstrVal) {
				nLenA = WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)v.bstrVal, nLen, NULL, 0, NULL, NULL);
				if (nLenA > m_nSize) {
					nLenA = m_nSize;
				}
				WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)v.bstrVal, nLen, &m_pc[nIndex], nLenA, NULL, NULL);
			}
			m_pc[nIndex + nLenA] = NULL;
			VariantClear(&v);
			break;
		case VT_FILETIME:
			teVariantTimeToFileTime(pv->date, (LPFILETIME)&m_pc[nIndex]);
			break;
		case VT_CY:
			GetPointFormVariant((LPPOINT)&m_pc[nIndex], pv);
			break;
		case VT_CARRAY:
			CteMemory *pMem;
			if (GetMemFormVariant(pv, &pMem)) {
				::CopyMemory(&m_pc[nIndex], pMem->m_pc, 16);
				pMem->Release();
			}
			else {
				::ZeroMemory(&m_pc[nIndex], 16);
			}
			break;
		case VT_CLSID:
			if (pv->vt == VT_BSTR) {
				CLSIDFromString(pv->bstrVal, (LPCLSID)&m_pc[nIndex]);
			}
			break;
		case VT_VARIANT:
			VariantCopy((VARIANT *)&m_pc[nIndex], pv);
			break;
	}
}

int CteMemory::GetLong(int nIndex)
{
	int *pi;
	pi = (int *)m_pc;
	return pi[nIndex];
}

void CteMemory::SetLong(int nIndex, int nValue)
{
	int *pi;
	pi = (int *)m_pc;
	pi[nIndex] = nValue;
}

//CteContextMenu

CteContextMenu::CteContextMenu(IContextMenu *pContextMenu, IDataObject *pDataObj)
{
	m_cRef = 1;
	m_pContextMenu = NULL;
	m_pContextMenu2 = NULL;
	m_pContextMenu3 = NULL;
	m_pDataObj = pDataObj;

	if (pContextMenu) {
		if SUCCEEDED(pContextMenu->QueryInterface(IID_PPV_ARGS(&m_pContextMenu))) {
			pContextMenu->QueryInterface(IID_PPV_ARGS(&m_pContextMenu2));
			pContextMenu->QueryInterface(IID_PPV_ARGS(&m_pContextMenu3));
		}
	}
}

CteContextMenu::~CteContextMenu()
{
	if (m_pContextMenu3) {
		m_pContextMenu3->Release();
	}
	if (m_pContextMenu2) {
		m_pContextMenu2->Release();
	}
	if (m_pContextMenu) {
		m_pContextMenu->Release();
	}
	if (m_pDataObj) {
		m_pDataObj->Release();
	}
}

STDMETHODIMP CteContextMenu::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else if (IsEqualIID(riid, IID_IContextMenu)) {
		return m_pContextMenu->QueryInterface(IID_IContextMenu, ppvObject);
	}
	else if (IsEqualIID(riid, g_ClsIdContextMenu)) {
		*ppvObject = this;
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
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
	return teGetDispId(methodCM, sizeof(methodCM), g_maps[MAP_CM], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteContextMenu::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	HRESULT hr;

	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	switch(dispIdMember) {
		//QueryContextMenu
		case 1:	
			if (nArg >= 4) {
				hr = m_pContextMenu->QueryContextMenu((HMENU)GetLLFromVariant(&pDispParams->rgvarg[nArg]),
					(UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]),
					(UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]),
					(UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg - 3]),
					(UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg - 4]));
				if (pVarResult) {
					pVarResult->lVal = hr;
					pVarResult->vt = VT_I4;
				}
			}
			return S_OK;
		//InvokeCommand
		case 2:	
			if (nArg >= 7) {
				if (g_pOnFunc[TE_OnInvokeCommand]) {
					VARIANT vResult;
					VariantInit(&vResult);
					VARIANTARG *pv = GetNewVARIANT(nArg + 2);
					teSetObject(&pv[nArg + 1], this);
					for (int i = nArg; i >= 0; i--) {
						VariantCopy(&pv[i], &pDispParams->rgvarg[i]);
					}
					Invoke4(g_pOnFunc[TE_OnInvokeCommand], &vResult, nArg + 2, pv);
					if (GetIntFromVariantClear(&vResult) == S_OK) {
						return S_OK;
					}
				}
				CMINVOKECOMMANDINFOEX cmi;
				VARIANTARG *pv = GetNewVARIANT(3);
				WCHAR **ppwc = new WCHAR*[3];
				char **ppc = new char*[3];
				for (int i = 0; i <= 2; i++) {
					if (pDispParams->rgvarg[nArg - i - 2].vt == VT_I4) {
						ppwc[i] = (WCHAR *)(UINT_PTR)pDispParams->rgvarg[nArg - i - 2].lVal;
						ppc[i] = (char *)(UINT_PTR)pDispParams->rgvarg[nArg - i - 2].lVal;
					}
					else {
						teVariantChangeType(&pv[i], &pDispParams->rgvarg[nArg - i - 2], VT_BSTR);
						ppwc[i] = pv[i].bstrVal;
						if ((ULONG_PTR)ppwc[i] > 0xffff) {
							int nLenW = SysStringLen(ppwc[i]);
							int nLenA = WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)ppwc[i], nLenW, NULL, 0, NULL, NULL);
							ppc[i] = new char[nLenA + 1];
							WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)ppwc[i], nLenW, (LPSTR)ppc[i], nLenA, NULL, NULL);
							ppc[i][nLenA] = NULL;
						}
						else {
							ppc[i] = (char *)ppwc[i];
						}
					}
				}
				::ZeroMemory(&cmi, sizeof(cmi));
				cmi.cbSize = sizeof(cmi);
				cmi.fMask = (DWORD)GetIntFromVariant(&pDispParams->rgvarg[nArg]) | CMIC_MASK_UNICODE;
				cmi.hwnd  = (HWND)GetLLFromVariant(&pDispParams->rgvarg[nArg - 1]);
				cmi.lpVerbW = ppwc[0];
				cmi.lpVerb = ppc[0];
				cmi.lpParametersW = ppwc[1];
				cmi.lpParameters = ppc[1];
				cmi.lpDirectoryW = ppwc[2];
				cmi.lpDirectory = ppc[2];
				cmi.nShow = GetIntFromVariant(&pDispParams->rgvarg[nArg - 5]);
				cmi.dwHotKey = (DWORD)GetIntFromVariant(&pDispParams->rgvarg[nArg - 6]);
				cmi.hIcon =(HANDLE)GetLLFromVariant(&pDispParams->rgvarg[nArg - 7]);
				try {
					hr = m_pContextMenu->InvokeCommand((LPCMINVOKECOMMANDINFO)&cmi);
				} catch (...) {
					hr = E_UNEXPECTED;
				}
				if (pVarResult) {
					pVarResult->lVal = hr;
					pVarResult->vt = VT_I4;
				}
				for (int i = 2; i >= 0; i--) {
					VariantClear(&pv[i]);
					if ((ULONG_PTR)ppc[i] >= 0xffff) {
						delete [] ppc[i];
					}
				}
				delete [] pv;
			}
			return S_OK;
		//Items
		case 3:	
			if (pVarResult) {
				FolderItems *pFolderItems;
				pFolderItems = new CteFolderItems(m_pDataObj, NULL, false);
				teSetObject(pVarResult, pFolderItems);
				pFolderItems->Release();
			}
			return S_OK;
		//GetCommandString
		case 4:	
			if (nArg >= 1) {
				if (pVarResult) {
					WCHAR szName[MAX_PATH];
					szName[0] = NULL;
					if SUCCEEDED(m_pContextMenu->GetCommandString(
						(INT_PTR)GetLLFromVariant(&pDispParams->rgvarg[nArg]),
						GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]),
						NULL, (LPSTR)&szName, MAX_PATH))
					{
						pVarResult->bstrVal = SysAllocString(szName);
						pVarResult->vt = VT_BSTR;
					}
				}
			}
			return S_OK;

		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteDropTarget

CteDropTarget::CteDropTarget(IDropTarget *pDropTarget, LPITEMIDLIST pidl)
{
	m_cRef = 1;
	m_pidl = pidl;
	pDropTarget->QueryInterface(IID_PPV_ARGS(&m_pDropTarget));
}

CteDropTarget::~CteDropTarget()
{
	if (m_pidl) {
		::CoTaskMemFree(m_pidl);
	}
	if (m_pDropTarget) {
		m_pDropTarget->Release();
	}
}

STDMETHODIMP CteDropTarget::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
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
	return teGetDispId(methodDT, sizeof(methodDT), g_maps[MAP_DT], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteDropTarget::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	HRESULT hr;

	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	switch(dispIdMember) {
		//DragEnter
		case 1:	
		//DragOver
		case 2:		
		//Drop
		case 3:	
			hr = E_INVALIDARG;
			if (nArg >= 2 && m_pDropTarget) {
				IDataObject *pDataObj;
				if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
					DWORD grfKeyState = (DWORD)GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
					POINTL pt;
					GetPointFormVariant((POINT *)(char *)&pt, &pDispParams->rgvarg[nArg - 2]);
					DWORD *pdwEffect = NULL;
					if (nArg >= 3) {
						pdwEffect = (LPDWORD)GetpcFormVariant(&pDispParams->rgvarg[nArg - 3]);
					}
					DWORD dwEffect7 = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
					if (pdwEffect == NULL) {
						pdwEffect = &dwEffect7;
					}
					DWORD dwEffect = *pdwEffect;
					hr = m_pDropTarget->DragEnter(pDataObj, grfKeyState, pt, pdwEffect);
					if (hr == S_OK && dispIdMember >= 2) {
						CteFolderItems *pDragItems = new CteFolderItems(pDataObj, NULL, false);
						POINTL pt0 = {0, 0};
						hr = S_FALSE;
						if (m_pidl) {
							*pdwEffect = dwEffect;
							hr = DragSub(TE_OnDragOver, this, pDragItems, grfKeyState, pt0, pdwEffect);
						}
						if (hr != S_OK) {
							*pdwEffect = dwEffect;
							hr = m_pDropTarget->DragOver(grfKeyState, pt, pdwEffect);
						}
						if (hr == S_OK && dispIdMember >= 3) {
							hr = S_FALSE;
							if (m_pidl) {
								*pdwEffect = dwEffect;
								hr = DragSub(TE_OnDrop, this, pDragItems, grfKeyState, pt0, pdwEffect);
							}
							if (hr != S_OK) {
								*pdwEffect = dwEffect;
								hr = m_pDropTarget->Drop(pDataObj, grfKeyState, pt, pdwEffect);
							}
						}
						pDragItems->Release();
					}
					m_pDropTarget->DragLeave();
					if (dispIdMember >= 3) {
						DragLeaveSub(this, NULL);
					}
				}
			}
			if (pVarResult) {
				pVarResult->lVal = hr;
				pVarResult->vt = VT_I4;
			}
			return S_OK;

		//DragLeave
		case 4:	
			hr = m_pDropTarget->DragLeave();
			DragLeaveSub(this, NULL);
			if (pVarResult) {
				pVarResult->lVal = hr;
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//Type
		case 5:	
			if (pVarResult) {
				pVarResult->lVal = CTRL_DT;
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//FolderItem
		case 6:
			if (pVarResult) {
				FolderItem *pid;
				if (GetFolderItemFromPidl(&pid, m_pidl)) {
					teSetObject(pVarResult, pid);
					pid->Release();
				}
			}
			return S_OK;

		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteTreeView

CteTreeView::CteTreeView()
{
	m_cRef = 1;
	m_DefProc = NULL;
	m_DefProc2 = NULL;
	m_bMain = true;
	m_pDragItems = NULL;
	m_pDropTarget = NULL;
	VariantInit(&m_Data);
#ifdef _W2000
	VariantInit(&m_vSelected);
#endif
}

CteTreeView::~CteTreeView()
{
	Close();
	if (m_DefProc) {
		SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)m_DefProc);
	}
	if (m_DefProc2) {
		HWND hwnd = FindTreeWindow(m_hwnd);
		SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LONG_PTR)m_DefProc2);
		if (m_pDropTarget) {
			SetProp(hwnd, L"OleDropTargetInterface", (HANDLE)m_pDropTarget);
		}
	}
#ifdef _W2000
	VariantClear(&m_vSelected);
#endif
	VariantClear(&m_Data);
}

void CteTreeView::Init()
{
	m_hwnd = NULL;
	m_pNameSpaceTreeControl = NULL;
	m_pShellNameSpace = NULL;
	m_bSetRoot = true;
#ifdef _VISTA7
	m_dwCookie = 0;
#endif
//	m_dwEventCookie = 0;
}

void CteTreeView::Close()
{
#ifdef _VISTA7
	if (m_pNameSpaceTreeControl) {
		m_pNameSpaceTreeControl->RemoveAllRoots();
		m_pNameSpaceTreeControl->TreeUnadvise(m_dwCookie);
		m_dwCookie = 0;

		IOleWindow *pOleWindow;
		HWND hwnd;
		if SUCCEEDED(m_pNameSpaceTreeControl->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
			pOleWindow->GetWindow(&hwnd);
			pOleWindow->Release();
			PostMessage(hwnd, WM_CLOSE, 0, 0);	
		}
		m_pNameSpaceTreeControl = NULL;
	}
#endif
#ifdef _2000XP
	if (m_pShellNameSpace) {
		IOleObject *pOleObject;
		if SUCCEEDED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pOleObject))) {
			RECT rc;
			SetRectEmpty(&rc);
			pOleObject->DoVerb(OLEIVERB_HIDE, NULL, NULL, 0, g_hwndMain, &rc);
			pOleObject->Close(OLECLOSE_NOSAVE);
			pOleObject->SetClientSite(NULL);
			pOleObject->Release();
		}
		m_pShellNameSpace->Release();
		m_pShellNameSpace = NULL;
	}
#endif
}

BOOL CteTreeView::Create()
{
#ifdef _VISTA7
	if (
//		false &&
		SUCCEEDED(CoCreateInstance(CLSID_NamespaceTreeControl, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pNameSpaceTreeControl)))) {
		RECT rc;
		SetRectEmpty(&rc);
		if SUCCEEDED(m_pNameSpaceTreeControl->Initialize(GetParentWindow(), &rc, m_pFV->m_param[SB_TreeFlags])) {
			m_pNameSpaceTreeControl->TreeAdvise(static_cast<INameSpaceTreeControlEvents *>(this), &m_dwCookie);
			IOleWindow *pOleWindow;
			if SUCCEEDED(m_pNameSpaceTreeControl->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
				if (pOleWindow->GetWindow(&m_hwnd) == S_OK) {
					HWND hwnd = FindTreeWindow(m_hwnd);
					if (hwnd) {
						m_DefProc2 = (WNDPROC)SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LONG_PTR)TETVProc2);
						m_pFV->HookDragDrop(g_dragdrop & 2);
					}
					BringWindowToTop(m_hwnd);
					ArrangeWindow();
				}
				pOleWindow->Release();
			}
			DoFunc(TE_OnCreate, this, E_NOTIMPL);
			return TRUE;
		}
	}
#endif
#ifdef _2000XP
	if FAILED(CoCreateInstance(CLSID_ShellShellNameSpace, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pShellNameSpace))) {
		if FAILED(CoCreateInstance(CLSID_ShellNameSpace, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pShellNameSpace))) {
			return FALSE;
		}
	}
	IQuickActivate *pQuickActivate;
	if (FAILED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pQuickActivate)))) {
		return FALSE;
	}
	QACONTAINER qaContainer;
	::ZeroMemory(&qaContainer, sizeof(QACONTAINER));
	qaContainer.cbSize        = sizeof(QACONTAINER);
	qaContainer.pClientSite   = static_cast<IOleClientSite *>(this);
	qaContainer.pUnkEventSink = static_cast<IDispatch *>(this);
	QACONTROL qaControl;
	qaControl.cbSize = sizeof(QACONTROL);
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
/*	else if (m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pPersistStorage)) == S_OK) {
		pPersistStorage->InitNew(NULL);
		pPersistStorage->Release();
	}
	else if (m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pPersistMemory)) == S_OK) {
		pPersistMemory->InitNew();
		pPersistMemory->Release();
	}
	else if (m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pPersistPropertyBag)) == S_OK) {
		pPersistPropertyBag->InitNew();
		pPersistPropertyBag->Release();
	}*/
	else {
		return FALSE;
	}
/*	IRunnableObject *pRunnableObject;
	if (qaControl.dwMiscStatus & OLEMISC_ALWAYSRUN) {
		m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pRunnableObject));
		if (pRunnableObject) {
			pRunnableObject->Run(NULL);
			pRunnableObject->Release();
		}
	}*/
	RECT rc;
	SetRectEmpty(&rc);

	IOleObject *pOleObject;
	if FAILED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pOleObject))) {
		return FALSE;
	}
	pOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, static_cast<IOleClientSite *>(this), 0, g_hwndMain, &rc);
	pOleObject->Release();

	IOleWindow *pOleWindow;
	if SUCCEEDED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
		if (pOleWindow->GetWindow(&m_hwnd) == S_OK) {
			m_pShellNameSpace->put_Flags(m_pFV->m_param[SB_TreeFlags] & NSTCS_FAVORITESMODE);
			HWND hwnd = FindTreeWindow(m_hwnd);
			if (hwnd) {
				m_DefProc = (WNDPROC)SetWindowLongPtr(GetParent(hwnd), GWLP_WNDPROC, (LONG_PTR)TETVProc);
				m_DefProc2 = (WNDPROC)SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LONG_PTR)TETVProc2);
				m_pFV->HookDragDrop(g_dragdrop & 2);

				DWORD dwStyle = GetWindowLong(hwnd, GWL_STYLE);
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
					{ NSTCS_HORIZONTALSCROLL, TVS_NOHSCROLL },//Reverse
					//NSTCS_ROOTHASEXPANDO
					{ NSTCS_SHOWSELECTIONALWAYS, TVS_SHOWSELALWAYS },
					{ NSTCS_NOINFOTIP, TVS_INFOTIP },//Reverse
					{ NSTCS_EVENHEIGHT, TVS_NONEVENHEIGHT },//Reverse
					//NSTCS_NOREPLACEOPEN	= 0x800,
					{ NSTCS_DISABLEDRAGDROP, TVS_DISABLEDRAGDROP },
					//NSTCS_NOORDERSTREAM
					{ NSTCS_BORDER, WS_BORDER },
					{ NSTCS_NOEDITLABELS, TVS_EDITLABELS },//Reverse
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
					if (m_pFV->m_param[SB_TreeFlags] & style[i].nstc) {
						dwStyle |= style[i].nsc;
					}
					else {
						dwStyle &= ~style[i].nsc;
					}
				}
				dwStyle ^= (TVS_NOHSCROLL | TVS_INFOTIP | TVS_NONEVENHEIGHT | TVS_EDITLABELS);
				SetWindowLongPtr(hwnd, GWL_STYLE, dwStyle);
			}
			BringWindowToTop(m_hwnd);
			ArrangeWindow();
		}
		pOleWindow->Release();
	}
	DoFunc(TE_OnCreate, this, E_NOTIMPL);
	return TRUE;
#else
	return FALSE;
#endif
}

STDMETHODIMP CteTreeView::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch) || IsEqualIID(riid, DIID_DShellNameSpaceEvents)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
/*	else if (IsEqualIID(riid, IID_IContextMenu)) {
		*ppvObject = static_cast<IContextMenu *>(this);
	}*/
	else if (IsEqualIID(riid, IID_IDropTarget)) {
		*ppvObject = static_cast<IDropTarget *>(this);
	}
#ifdef _VISTA7
	else if (IsEqualIID(riid, IID_INameSpaceTreeControlEvents)) {
		*ppvObject = static_cast<INameSpaceTreeControlEvents *>(this);
	}
	else if (m_pNameSpaceTreeControl && IsEqualIID(riid, IID_INameSpaceTreeControl)) {
		return m_pNameSpaceTreeControl->QueryInterface(riid, ppvObject);
	}
#endif
#ifdef _2000XP
	else if (IsEqualIID(riid, IID_IOleClientSite)) {
		*ppvObject = static_cast<IOleClientSite *>(this);
	}
	else if (IsEqualIID(riid, IID_IOleWindow) || IsEqualIID(riid, IID_IOleInPlaceSite)) {
		*ppvObject = static_cast<IOleInPlaceSite *>(this);
	}
	else if (m_pShellNameSpace && IsEqualIID(riid, IID_IOleObject)) {
		return m_pShellNameSpace->QueryInterface(riid, ppvObject);
	}
	else if (m_pShellNameSpace && IsEqualIID(riid, IID_IShellNameSpace)) {
		return m_pShellNameSpace->QueryInterface(riid, ppvObject);
	}
/*	else if (IsEqualIID(riid, IID_IOleControlSite)) {
		*ppvObject = static_cast<IOleControlSite *>(this);
	}*/
#endif
	else if (IsEqualIID(riid, g_ClsIdTV)) {
		*ppvObject = this;
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
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
	return teGetDispId(methodTV, sizeof(methodTV), g_maps[MAP_TV], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteTreeView::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}

	if (dispIdMember >= 0x10000101 && dispIdMember < TE_OFFSET) {
		if (m_pFV->m_param[SB_TreeAlign] & TE_TV_View) {
			if (!m_pNameSpaceTreeControl && !m_pShellNameSpace) {
				Create();
			}
			if (m_bSetRoot && (m_pNameSpaceTreeControl || m_pShellNameSpace)) {
				SetRoot();
			}
		}
		else if (dispIdMember != 0x20000002) {
			return S_FALSE;
		}
	}

	switch(dispIdMember) {
		//Data
		case 0x10000001:
			if (nArg >= 0) {
				VariantClear(&m_Data);
				VariantCopy(&m_Data, &pDispParams->rgvarg[nArg]);
			}
			if (pVarResult) {
				VariantCopy(pVarResult, &m_Data);
			}
			return S_OK;
		//Type
		case 0x10000002:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = CTRL_TV;
			}
			return S_OK;
		//hwnd
		case 0x10000003:
			SetVariantLL(pVarResult, (LONGLONG)m_hwnd);
			return S_OK;
		//Close
		case 0x10000004:
			m_pFV->m_param[SB_TreeAlign] &= ~TE_TV_View;
			ArrangeWindow();
			return S_OK;

		//FolderView
		case 0x10000007:
			if (pVarResult) {
				teSetObject(pVarResult, m_pFV);
			}
			return S_OK;

		//Align
		case 0x10000008:
			if (nArg >= 0) {
				m_pFV->m_param[SB_TreeAlign] = GetIntFromVariant(&pDispParams->rgvarg[nArg]) | TE_TV_Use;
				if (!m_pNameSpaceTreeControl && !m_pShellNameSpace) {
					if (m_pFV->m_param[SB_TreeAlign] & TE_TV_View) {
						if (Create()) {
							if (m_bSetRoot) {
								SetRoot();
							}
						}
					}
				}
				ArrangeWindow();
			}
			if (pVarResult) {
				pVarResult->lVal = m_pFV->m_param[SB_TreeAlign];
				pVarResult->vt = VT_I4;
			}
			return S_OK;
		//Focus
		case 0x10000105:
			SetFocus(FindTreeWindow(m_hwnd));
			return S_OK;
		//HitTest (Screen coordinates)
		case 0x10000106:
			if (nArg >= 1 && pVarResult) {
				TVHITTESTINFO info;
				GetPointFormVariant(&info.pt, &pDispParams->rgvarg[nArg]);
				UINT flags = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
				HTREEITEM hItem = (HTREEITEM)DoHitTest(this, info.pt, flags);
				if ((LONGLONG)hItem < 1) {
					HWND hwnd = FindTreeWindow(m_hwnd);
					ScreenToClient(hwnd, &info.pt);
					info.flags = flags;
					hItem = TreeView_HitTest(hwnd, &info);
					if (!(info.flags & flags)) {
						hItem = 0;
					}
				}
				SetVariantLL(pVarResult, (LONGLONG)hItem);
			}
			return S_OK;

		//SelectedItem
		case 0x20000000:
			if (pVarResult) {
				IDispatch *pdisp;
				if (getSelected(&pdisp) == S_OK) {
					teSetObject(pVarResult, pdisp);
					pdisp->Release();
					return S_OK;
				}
			}
			break;
		//Root
		case 0x20000002:
			if (pVarResult) {
				VariantCopy(pVarResult, &m_pFV->m_vRoot);
				if (pVarResult->vt == VT_NULL) {
					pVarResult->vt = VT_I4;
					pVarResult->lVal = 0;
				}
				if (nArg < 0) {
					return S_OK;
				}
			}
		//SetRoot
		case 0x20000003:
			if (nArg >= 0) {
				m_pFV->m_param[SB_EnumFlags] = g_paramFV[SB_EnumFlags];
				m_pFV->m_param[SB_RootStyle] = g_paramFV[SB_RootStyle];
				if (nArg >= 1) {
					m_pFV->m_param[SB_EnumFlags] = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
					if (nArg >= 2) {
						m_pFV->m_param[SB_RootStyle] = GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]);
					}
				}

				VariantCopy(&m_pFV->m_vRoot, &pDispParams->rgvarg[nArg]);
				return SetRoot();
			}
			break;
		//Expand
#ifdef _VISTA7
		case 0x20000004:
			if (m_pNameSpaceTreeControl) {
				if (nArg >= 1) {
					LPITEMIDLIST pidl;
					GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg]);
					IShellItem *pShellItem;
					DWORD dwState;
					dwState = NSTCIS_SELECTED;
					if (GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) != 0) {
						dwState += NSTCIS_EXPANDED;
					}
					if (lpfnSHCreateItemFromIDList) {
						if SUCCEEDED(lpfnSHCreateItemFromIDList(pidl, IID_PPV_ARGS(&pShellItem))) {
							m_pNameSpaceTreeControl->SetItemState(pShellItem, dwState, dwState);
							pShellItem->Release();
						}
					}
					::CoTaskMemFree(pidl);
				}
				return S_OK;
			}
			break;
#endif
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
			SetChildren();
			return DoFunc(TE_OnSelectionChanged, this, E_NOTIMPL);
		case DISPID_DOUBLECLICK:
			return S_OK;
		case DISPID_READYSTATECHANGE:
			return S_OK;
/*		case DISPID_AMBIENT_BACKCOLOR:
//			return E_NOTIMPL;
		case DISPID_INITIALIZED:
			return E_NOTIMPL;*/
		case DISPID_AMBIENT_LOCALEID:
			pVarResult->vt = VT_I4;
			pVarResult->lVal = (SHORT)GetThreadLocale();
			return S_OK;
		case DISPID_AMBIENT_USERMODE:
			pVarResult->vt = VT_BOOL;
			pVarResult->boolVal = VARIANT_TRUE;	
			return S_OK;
		case DISPID_AMBIENT_DISPLAYASDEFAULT:
			pVarResult->vt = VT_BOOL;
			pVarResult->boolVal = VARIANT_FALSE;
			return S_OK;
#endif
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
		default:
			if (dispIdMember >= TE_OFFSET && dispIdMember <= TE_OFFSET + SB_RootStyle) {
				if (nArg >= 0) {
					DWORD dw = GetIntFromVariantPP(&pDispParams->rgvarg[nArg], dispIdMember - TE_OFFSET);
					if (dw != m_pFV->m_param[dispIdMember - TE_OFFSET]) {
						m_pFV->m_param[dispIdMember - TE_OFFSET] = dw;
						g_paramFV[dispIdMember - TE_OFFSET] = dw;
						if (dispIdMember == TE_OFFSET + SB_TreeFlags) {
							Close();
							m_bSetRoot = true;
						}
						ArrangeWindow();
					}
				}
				if (pVarResult) {
					pVarResult->lVal = m_pFV->m_param[dispIdMember - TE_OFFSET];
					pVarResult->vt = VT_I4;
				}
				return S_OK;
			}
			break;
	}
	if (dispIdMember >= 0x20000000 && dispIdMember < 0x20000000 + TVVERBS) {
		if (m_pNameSpaceTreeControl) {
			return S_OK;
		}
#ifdef _2000XP
		if (m_pShellNameSpace) {
			for (int i = _countof(methodTV); i--;) {
				if (methodTV[i].id == dispIdMember) {
					NSInvoke(methodTV[i].name, wFlags, pDispParams, pVarResult);
					break;
				}
			}
		}
#endif
		return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

#ifdef _VISTA7
//INameSpaceTreeControlEvents
STDMETHODIMP CteTreeView::OnItemClick(IShellItem *psi, NSTCEHITTEST nstceHitTest, NSTCSTYLE nsctsFlags)
{
	if (g_pOnFunc[TE_OnItemClick]) {
		VARIANTARG *pv;
		VARIANT vResult;
		pv = GetNewVARIANT(4);
		teSetObject(&pv[3], this);

		if (m_pNameSpaceTreeControl) {
			GetVariantFromShellItem(&pv[2], psi);
		}
		else {
			IDispatch *pdisp;
			if (getSelected(&pdisp) == S_OK) {
				teSetObject(&pv[2], pdisp);
				pdisp->Release();
			}
		}
		pv[1].vt = VT_I4;
		pv[1].lVal = nstceHitTest;
		pv[0].vt = VT_I4;
		pv[0].lVal = nsctsFlags;
		VariantInit(&vResult);
		Invoke4(g_pOnFunc[TE_OnItemClick], &vResult, 4, pv);
		return GetIntFromVariantClear(&vResult);
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
	if (g_pOnFunc[TE_OnShowContextMenu]) {
		m_pNameSpaceTreeControl->SetItemState(psi, NSTCIS_SELECTED, NSTCIS_SELECTED);
		if (MessageSubPt(TE_OnShowContextMenu, this, &msg1) == S_OK) {
			return QueryInterface(IID_IContextMenu, ppv);
		}
	}
/*	if (pcmIn) {
		HMENU hMenu = CreatePopupMenu();
		if SUCCEEDED(pcmIn->QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXPLORE)) {
			int nVerb = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, msg1.pt.x, msg1.pt.y, g_hwndMain, NULL);

		}
		DestroyMenu(hMenu);
		return QueryInterface(IID_IContextMenu, ppv);
	}*/
	*ppv = NULL;
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
#endif
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
	*phwnd = GetParentWindow();
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

	CopyRect(lprcClipRect, &m_rc);
	CopyRect(lprcPosRect, &m_rc);
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

//IOleControlSite
/*
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
*/
#endif

//IContextMenu
/*STDMETHODIMP CteTreeView::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
	return S_OK;
}

STDMETHODIMP CteTreeView::GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteTreeView::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
	return S_OK;
}
*/
#ifdef _2000XP
//Adjust Tree button
VOID CteTreeView::SetChildren()
{
	if (m_pShellNameSpace && !(m_pFV->m_param[SB_EnumFlags] & SHCONTF_NONFOLDERS)) {
		HWND hwnd = FindTreeWindow(m_hwnd);
		TVITEM item;
		::ZeroMemory(&item, sizeof(TVITEM));
		item.hItem = TreeView_GetSelection(hwnd);
		if (item.hItem) {
			TreeView_EnsureVisible(hwnd, item.hItem);
			IDispatch *pItem;
			if SUCCEEDED(getSelected(&pItem)) {
				LPITEMIDLIST pidl;
				if (GetIDListFromObject(pItem, &pidl)) {
					IShellFolder *pSF;
					LPCITEMIDLIST ItemID;
					if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &ItemID)) {
						SFGAOF dwAttr = SFGAO_HASSUBFOLDER;
						if SUCCEEDED(pSF->GetAttributesOf(1, &ItemID, &dwAttr)) {
							item.mask = TVIF_CHILDREN;
							item.cChildren = dwAttr & SFGAO_HASSUBFOLDER ? 1 : 0;
							TreeView_SetItem(hwnd, &item);
						}
						pSF->Release();
					}
					::CoTaskMemFree(pidl);
				}
				pItem->Release();
			}
		}
	}
}

HRESULT CteTreeView::NSInvoke(LPWSTR lpVerb, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HRESULT hr = E_NOTIMPL;
	if (m_pShellNameSpace) {
		IDispatch *pDispatch;
		if SUCCEEDED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pDispatch))) {
			DISPID dispid;
			if SUCCEEDED(pDispatch->GetIDsOfNames(IID_NULL, &lpVerb , 1, LOCALE_USER_DEFAULT, &dispid)) {
				hr = pDispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
					wFlags, pDispParams, pVarResult, NULL, NULL);
			}
			pDispatch->Release();
		}
	}
	return hr;
}
#endif

HRESULT CteTreeView::SetRoot()
{
	HRESULT hr = E_FAIL;
#ifdef _VISTA7
	if (m_pNameSpaceTreeControl) {
		IShellItem *pShellItem;
		if (lpfnSHCreateItemFromIDList) {
			LPITEMIDLIST pidl;
			if (!GetPidlFromVariant(&pidl, &m_pFV->m_vRoot)) {
				pidl = ::ILClone(g_pidls[CSIDL_DESKTOP]);
			}
			if SUCCEEDED(lpfnSHCreateItemFromIDList(pidl, IID_PPV_ARGS(&pShellItem))) {
				m_pNameSpaceTreeControl->RemoveAllRoots();
				hr = m_pNameSpaceTreeControl->AppendRoot(pShellItem, m_pFV->m_param[SB_EnumFlags], m_pFV->m_param[SB_RootStyle], NULL);
				pShellItem->Release();
			}
			::CoTaskMemFree(pidl);
		}
	}
#endif
#ifdef _2000XP
	if (hr != S_OK && m_pShellNameSpace) {
		HWND hwnd = FindTreeWindow(m_hwnd);
		DWORD dwStyle = GetWindowLong(hwnd, GWL_STYLE);
		if (m_pFV->m_param[SB_RootStyle] & NSTCRS_HIDDEN) {
			dwStyle &= ~TVS_LINESATROOT;
		}
		else {
			dwStyle |= TVS_LINESATROOT;
		}
		SetWindowLong(hwnd, GWL_STYLE, dwStyle);

		//running to call DISPID_FAVSELECTIONCHANGE SetRoot(Windows 2000)
		m_pShellNameSpace->put_EnumOptions(m_pFV->m_param[SB_EnumFlags]);

		LPITEMIDLIST pidl;
		if (GetPidlFromVariant(&pidl, &m_pFV->m_vRoot)) {
			BSTR bs;
			if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORPARSING)) {
				if (SysStringLen(bs) && bs[0] == ':') {
					hr = m_pShellNameSpace->SetRoot(bs);
				}
				::SysFreeString(bs);
			}
			::CoTaskMemFree(pidl);
		}
		if (hr != S_OK) {
			if (m_pFV->m_vRoot.vt == VT_BSTR) {
				hr = m_pShellNameSpace->SetRoot(m_pFV->m_vRoot.bstrVal);
				if (hr != S_OK) {
					VARIANT v;
					teVariantChangeType(&v, &m_pFV->m_vRoot, VT_I4); 
					hr = m_pShellNameSpace->put_Root(v);
					m_pShellNameSpace->SetRoot(NULL);
				}
			}
			else {
				hr = m_pShellNameSpace->put_Root(m_pFV->m_vRoot);
				m_pShellNameSpace->SetRoot(NULL);
			}
		}
	}
#endif
	if (hr == S_OK) {
		m_bSetRoot = false;
		DoFunc(TE_OnViewCreated, this, E_NOTIMPL);
	}
	return hr;
}

HWND CteTreeView::GetParentWindow()
{
	if (m_pFV->m_param[SB_TreeAlign] & TE_TV_Left && g_pTE->m_param[TE_Tab] && m_pFV->m_pTabs->m_param[TE_Align]) {
		return g_hwndMain;
	}
	else {
		return m_pFV->m_pTabs->m_hwndStatic;
	}
}

HRESULT CteTreeView::getSelected(IDispatch **ppid)
{
	*ppid = NULL;
	HRESULT hr = E_FAIL;
#ifdef _VISTA7
	if (m_pNameSpaceTreeControl) {
		IShellItemArray *psia;
		if SUCCEEDED(m_pNameSpaceTreeControl->GetSelectedItems(&psia)) {
			IShellItem *psi;
			if SUCCEEDED(psia->GetItemAt(0, &psi)) {
				FolderItem *pFI;
				if SUCCEEDED(GetFolderItemFromShellItem(&pFI, psi)) {
					hr = pFI->QueryInterface(IID_PPV_ARGS(ppid));
				}
				psi->Release();
			}
			psia->Release();
		}
		return hr;
	}
#endif
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

//IDropTarget
STDMETHODIMP CteTreeView::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);
	DWORD dwEffect = *pdwEffect;
	HRESULT hr = DragSub(TE_OnDragEnter, this, m_pDragItems, grfKeyState, pt, pdwEffect);
	if (hr != S_OK && m_pDropTarget) {
		*pdwEffect = dwEffect;
		hr = m_pDropTarget->DragEnter(pDataObj, grfKeyState, pt, pdwEffect);
	}
	return hr;
}

STDMETHODIMP CteTreeView::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_grfKeyState = grfKeyState;
	DWORD dwEffect = *pdwEffect;
	HRESULT hr = DragSub(TE_OnDragOver, this, m_pDragItems, grfKeyState, pt, pdwEffect);
	if (hr != S_OK && m_pDropTarget) {
		*pdwEffect = dwEffect;
		hr = m_pDropTarget->DragOver(grfKeyState, pt, pdwEffect);
	}
	return hr;
}

STDMETHODIMP CteTreeView::DragLeave()
{
	if (m_DragLeave == E_NOT_SET) {
		m_DragLeave = DragLeaveSub(this, &m_pDragItems);
		if (m_pDropTarget) {
			HRESULT hr = m_pDropTarget->DragLeave();
			if (m_DragLeave != S_OK) {
				m_DragLeave = hr;
			}
		}
	}
	return m_DragLeave;
}

STDMETHODIMP CteTreeView::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	GetDragItems(&m_pDragItems, pDataObj);
	DWORD dwEffect = *pdwEffect;
	HRESULT hr = DragSub(TE_OnDrop, this, m_pDragItems, m_grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		if (hr != S_OK) {
			*pdwEffect = dwEffect;
			hr = m_pDropTarget->Drop(pDataObj, grfKeyState, pt, pdwEffect);
		}
	}
	DragLeave();
	return hr;
}

// CteFolderItem

CteFolderItem::CteFolderItem(VARIANT *pv)
{
	m_cRef = 1;
	m_pidl = NULL;
	VariantInit(&m_v);
	if (pv) {
		VariantCopy(&m_v, pv);
	}
}

CteFolderItem::~CteFolderItem()
{
	if (m_pidl) {
		::CoTaskMemFree(m_pidl);
	}
	::VariantClear(&m_v);
}

LPITEMIDLIST CteFolderItem::GetPidl()
{
	if (m_v.vt != VT_EMPTY && m_pidl == NULL) {
		if (GetPidlFromVariant(&m_pidl, &m_v)) {
			::VariantClear(&m_v);
		}
	}
	return m_pidl ? m_pidl : g_pidls[CSIDL_DESKTOP];
}

HRESULT CteFolderItem::GetFolderItem(FolderItem **ppid)
{
	IPersistFolder *pPF = NULL;
	HRESULT hr = CoCreateInstance(CLSID_FolderItem, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pPF));
	if SUCCEEDED(hr) {
		hr = pPF->Initialize(GetPidl());
		if SUCCEEDED(hr) {
			hr = pPF->QueryInterface(IID_PPV_ARGS(ppid));
		}
		pPF->Release();
	}
	return hr;
}

HRESULT CteFolderItem::GetAttibute(VARIANT_BOOL *pb, DWORD dwFlag)
{
	HRESULT hr = E_FAIL;
	*pb = false;
	IShellFolder *pSF;
	LPCITEMIDLIST ItemID;
	if SUCCEEDED(SHBindToParent(GetPidl(), IID_PPV_ARGS(&pSF), &ItemID)) {
		DWORD attr = dwFlag;
		if SUCCEEDED(pSF->GetAttributesOf(1, (LPCITEMIDLIST *)&ItemID, &attr)) {
			*pb = (attr & dwFlag) != 0;
			hr = S_OK;
		}
		pSF->Release();
	}
	return hr;
}

STDMETHODIMP CteFolderItem::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else if (IsEqualIID(riid, IID_FolderItem)) {
		*ppvObject = static_cast<FolderItem *>(this);
	}
	else if (IsEqualIID(riid, IID_IPersist)) {
		*ppvObject = static_cast<IPersist *>(this);
	}
	else if (IsEqualIID(riid, IID_IPersistFolder)) {
		*ppvObject = static_cast<IPersistFolder *>(this);
	}
	else if (IsEqualIID(riid, IID_IPersistFolder2)) {
		*ppvObject = static_cast<IPersistFolder2 *>(this);
	}
	else if (IsEqualIID(riid, g_ClsIdFI)) {
		*ppvObject = this;
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
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
	if (m_pidl) {
		::CoTaskMemFree(m_pidl);
	}
	m_pidl = ::ILClone(pidl);
	VariantClear(&m_v);
	return S_OK;
}

//IPersistFolder2
STDMETHODIMP CteFolderItem::GetCurFolder(PIDLIST_ABSOLUTE *ppidl)
{
	*ppidl = ::ILClone(GetPidl());
	return S_OK;
}

STDMETHODIMP CteFolderItem::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	return teGetDispId(methodFI, sizeof(methodFI), g_maps[MAP_FI], *rgszNames, rgDispId, true); 
}

STDMETHODIMP CteFolderItem::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	HRESULT hr = E_FAIL;
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	switch(dispIdMember) {
		//Application
		case 1:
			if (pVarResult) {
				IDispatch *pdisp;
				hr = get_Application(&pdisp);
				if SUCCEEDED(hr) {
					teSetObject(pVarResult, pdisp);
					pdisp->Release();
				}
			}
			return hr;
		//Parent
		case 2:
			if (pVarResult) {
				IDispatch *pdisp;
				hr = get_Parent(&pdisp);
				if SUCCEEDED(hr) {
					teSetObject(pVarResult, pdisp);
					pdisp->Release();
				}
			}
			return hr;
		//Name
		case 3:
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
		//Path
		case 4:
			if (pVarResult) {
				hr = get_Path(&pVarResult->bstrVal);
				if SUCCEEDED(hr) {
					pVarResult->vt = VT_BSTR;
				}
			}
			return hr;
		//GetLink
		case 5:
			if (pVarResult) {
				IDispatch *pdisp;
				hr = get_GetLink(&pdisp);
				if SUCCEEDED(hr) {
					teSetObject(pVarResult, pdisp);
					pdisp->Release();
				}
			}
			return hr;
		//GetFolder
		case 6:
			if (pVarResult) {
				IDispatch *pdisp;
				hr = get_GetFolder(&pdisp);
				if SUCCEEDED(hr) {
					teSetObject(pVarResult, pdisp);
				}
			}
			return hr;
		//IsLink
		case 7:
			if (pVarResult) {
				hr = get_IsLink(&pVarResult->boolVal);
				if SUCCEEDED(hr) {
					pVarResult->vt = VT_BOOL;
				}
			}
			return hr;
		//IsFolder
		case 8:
			if (pVarResult) {
				hr = get_IsFolder(&pVarResult->boolVal);
				if SUCCEEDED(hr) {
					pVarResult->vt = VT_BOOL;
				}
			}
			return hr;
		//IsFileSystem
		case 9:
			if (pVarResult) {
				hr = get_IsFileSystem(&pVarResult->boolVal);
				if SUCCEEDED(hr) {
					pVarResult->vt = VT_BOOL;
				}
			}
			return hr;
		//IsBrowsable
		case 10:
			if (pVarResult) {
				hr = get_IsBrowsable(&pVarResult->boolVal);
				if SUCCEEDED(hr) {
					pVarResult->vt = VT_BOOL;
				}
			}
			return hr;
		//ModifyDate
		case 11:
			if (nArg >= 0) {
				VARIANT v;
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_DATE);
				if (v.vt == VT_DATE) {
					hr = put_ModifyDate(v.date);
				}
			}
			if (pVarResult) {
				hr = get_ModifyDate(&pVarResult->date);
				if SUCCEEDED(hr) {
					pVarResult->vt = VT_DATE;
				}
			}
			return hr;
		//Size
		case 12:
			if (pVarResult) {
				hr = get_Size(&pVarResult->lVal);
				if SUCCEEDED(hr) {
					pVarResult->vt = VT_I8;
				}
			}
			return hr;
		//Type
		case 13:
			if (pVarResult) {
				hr = get_Type(&pVarResult->bstrVal);
				if SUCCEEDED(hr) {
					pVarResult->vt = VT_BSTR;
				}
			}
			return hr;
		//Verbs
		case 14:
			if (pVarResult) {
				FolderItemVerbs *pid;
				hr = Verbs(&pid);
				if SUCCEEDED(hr) {
					teSetObject(pVarResult, pid);
					pid->Release();
				}
			}
			return hr;
		//InvokeVerb
		case 15:
			if (nArg >= 0) {
				hr = InvokeVerb(pDispParams->rgvarg[nArg]);
			}
			return hr;
		//Alt
		case 99:
			if (nArg >= 0) {
				if (m_pidl) {
					::CoTaskMemFree(m_pidl);
					m_pidl = NULL;
				}
				GetPidlFromVariant(&m_pidl, &pDispParams->rgvarg[nArg]);
			}
			return S_OK;
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

STDMETHODIMP CteFolderItem::get_Application(IDispatch **ppid)
{
	return g_pWebBrowser->m_pWebBrowser->get_Application(ppid);
}

STDMETHODIMP CteFolderItem::get_Parent(IDispatch **ppid)
{
	LPITEMIDLIST pidl = ::ILClone(GetPidl());
	ILRemoveLastID(pidl);
	HRESULT hr = GetFolderObjFromPidl(pidl, (Folder **)ppid);
	::CoTaskMemFree(pidl);
	return hr;
}

STDMETHODIMP CteFolderItem::get_Name(BSTR *pbs)
{
	if (m_v.vt == VT_BSTR && teIsFileSystem(m_v.bstrVal)) {
		*pbs = ::SysAllocString(PathFindFileName(m_v.bstrVal));
		return S_OK;
	}
	return GetDisplayNameFromPidl(pbs, GetPidl(), SHGDN_INFOLDER);
}

STDMETHODIMP CteFolderItem::put_Name(BSTR bs)
{
	HRESULT hr = E_FAIL;
	IShellFolder *pSF;
	LPCITEMIDLIST ItemID;
	if SUCCEEDED(SHBindToParent(GetPidl(), IID_PPV_ARGS(&pSF), &ItemID)) {
		if SUCCEEDED(pSF->SetNameOf(0, ItemID, bs, SHGDN_FOREDITING, NULL)) {
			hr = S_OK;
		}
		pSF->Release();
	}
	return hr;
}

STDMETHODIMP CteFolderItem::get_Path(BSTR *pbs)
{
	if (m_v.vt == VT_BSTR) {
		*pbs = ::SysAllocString(m_v.bstrVal);
		return S_OK;
	}
	return GetDisplayNameFromPidl(pbs, GetPidl(), SHGDN_FORPARSING);
}

STDMETHODIMP CteFolderItem::get_GetLink(IDispatch **ppid)
{
	FolderItem *pFI = NULL;
	HRESULT hr = GetFolderItem(&pFI);
	if SUCCEEDED(hr) {
		pFI->get_GetLink(ppid);
		pFI->Release();
	}
	return hr;
}

STDMETHODIMP CteFolderItem::get_GetFolder(IDispatch **ppid)
{
	return GetFolderObjFromPidl(GetPidl(), (Folder **)ppid);
}

STDMETHODIMP CteFolderItem::get_IsLink(VARIANT_BOOL *pb)
{
	return GetAttibute(pb, SFGAO_LINK);
}

STDMETHODIMP CteFolderItem::get_IsFolder(VARIANT_BOOL *pb)
{
	return GetAttibute(pb, SFGAO_FOLDER);
}

STDMETHODIMP CteFolderItem::get_IsFileSystem(VARIANT_BOOL *pb)
{
	return GetAttibute(pb, SFGAO_FILESYSTEM);
}

STDMETHODIMP CteFolderItem::get_IsBrowsable(VARIANT_BOOL *pb)
{
	return GetAttibute(pb, SFGAO_BROWSABLE);
}

STDMETHODIMP CteFolderItem::get_ModifyDate(DATE *pdt)
{
	WIN32_FIND_DATA data;
	HRESULT hr = GetDataFromPidl(GetPidl(), SHGDFIL_FINDDATA, &data, sizeof(WIN32_FIND_DATA));
	if SUCCEEDED(hr) {
		if (!teFileTimeToVariantTime(&data.ftLastWriteTime, pdt)) {
			return E_FAIL;
		}
	}
	return hr;
}

STDMETHODIMP CteFolderItem::put_ModifyDate(DATE dt)
{
	HRESULT hr = E_FAIL;

	FILETIME ft;
	if (teVariantTimeToFileTime(dt, &ft)) {
		BSTR bs;
		if SUCCEEDED(GetDisplayNameFromPidl(&bs, GetPidl(), SHGDN_FORPARSING)) {
			HANDLE hFile = ::CreateFile(bs, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0);
			if (hFile != INVALID_HANDLE_VALUE) {
				if (::SetFileTime(hFile, NULL, NULL, &ft)) {
					hr = S_OK;
				}
				::CloseHandle(hFile);
			}
			::SysFreeString(bs);
		}
	}
	return hr;
}

STDMETHODIMP CteFolderItem::get_Size(LONG *pul)
{
	WIN32_FIND_DATA data;
	HRESULT hr = GetDataFromPidl(GetPidl(), SHGDFIL_FINDDATA, &data, sizeof(WIN32_FIND_DATA));
	if SUCCEEDED(hr) {
		*pul = data.nFileSizeLow;
	}
	return hr;
}

STDMETHODIMP CteFolderItem::get_Type(BSTR *pbs)
{
	SHFILEINFO info;
	SHGetFileInfo((LPCWSTR)GetPidl(), 0, &info, sizeof(SHFILEINFO), SHGFI_PIDL | SHGFI_TYPENAME);
	*pbs = ::SysAllocString(info.szTypeName);
	return S_OK;
}

STDMETHODIMP CteFolderItem::Verbs(FolderItemVerbs **ppfic)
{
	FolderItem *pFI = NULL;
	HRESULT hr = GetFolderItem(&pFI);
	if SUCCEEDED(hr) {
		hr = pFI->Verbs(ppfic);
		pFI->Release();
	}
	return hr;
}

STDMETHODIMP CteFolderItem::InvokeVerb(VARIANT vVerb)
{
	FolderItem *pFI = NULL;
	HRESULT hr = GetFolderItem(&pFI);
	if SUCCEEDED(hr) {
		hr = pFI->InvokeVerb(vVerb);
		pFI->Release();
		return hr;
	}
#ifdef _W2000
	SHELLEXECUTEINFO    stExeInfo;	//Windows 2000
    stExeInfo.cbSize = sizeof(SHELLEXECUTEINFO);
	stExeInfo.hwnd = g_hwndMain;
    stExeInfo.fMask = SEE_MASK_INVOKEIDLIST;
	if (vVerb.vt == VT_BSTR) {
		stExeInfo.lpVerb = vVerb.bstrVal;
	}
	else if (vVerb.vt == VT_I4) {
		stExeInfo.lpVerb = (LPCWSTR)(UINT_PTR)vVerb.lVal;
	}
	else {
		stExeInfo.lpVerb = NULL;
	}
    stExeInfo.lpFile = NULL;
    stExeInfo.lpParameters = NULL;
    stExeInfo.lpDirectory = NULL;
    stExeInfo.nShow = SW_SHOWNORMAL;
    stExeInfo.hInstApp = NULL;
	stExeInfo.lpIDList = (LPVOID)GetPidl();

	if (::ShellExecuteEx(&stExeInfo)) {
		return S_OK;
	}
#endif
	return E_FAIL;
}

//CteCommonDialog

CteCommonDialog::CteCommonDialog()
{
	m_cRef = 1;

	::ZeroMemory(&m_ofn, sizeof(OPENFILENAME));
	m_ofn.lStructSize = sizeof(OPENFILENAME);
	m_ofn.hwndOwner = g_hwndMain;
//	m_ofn.hInstance = hInst;
	m_ofn.nMaxFile = MAX_PATH;
}

CteCommonDialog::~CteCommonDialog()
{
	
	if (m_ofn.lpstrFile) {
		delete [] m_ofn.lpstrFile;
	}
	if (m_ofn.lpstrInitialDir) {
		delete [] m_ofn.lpstrInitialDir;
	}
	if (m_ofn.lpstrFilter) {
		delete [] m_ofn.lpstrFilter;
	}
	if (m_ofn.lpstrDefExt) {
		delete [] m_ofn.lpstrDefExt;
	}
	if (m_ofn.lpstrTitle) {
		delete [] m_ofn.lpstrTitle;
	}
}

STDMETHODIMP CteCommonDialog::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
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
	return teGetDispId(methodCD, sizeof(methodCD), g_maps[MAP_CD], *rgszNames, rgDispId, true); 
}

STDMETHODIMP CteCommonDialog::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	LPCWSTR *lplpwstr = NULL;
	DWORD *pdw = NULL;
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	switch(dispIdMember) {
		//FileName
		case 10:	
			if (nArg >= 0) {
				VARIANT vFile;
				teVariantChangeType(&vFile, &pDispParams->rgvarg[nArg], VT_BSTR);
				if (!m_ofn.lpstrFile) {
					m_ofn.lpstrFile = new WCHAR[m_ofn.nMaxFile];
				}
				lstrcpyn(m_ofn.lpstrFile, vFile.bstrVal, m_ofn.nMaxFile);
				VariantClear(&vFile);
			}
			if (pVarResult) {
				pVarResult->bstrVal = SysAllocStringLen(m_ofn.lpstrFile, m_ofn.nMaxFile);
				pVarResult->vt = VT_BSTR;
			}
			return S_OK;
		//FilterView
		case 13:	
			if (nArg >= 0) {
				if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
					if (m_ofn.lpstrFilter) {
						delete [] m_ofn.lpstrFilter;
					}
					int i = SysStringLen(pDispParams->rgvarg[nArg].bstrVal);
					LPWSTR str = new WCHAR[i + 2];
					lstrcpy(str, pDispParams->rgvarg[nArg].bstrVal);
					str[i + 1] = NULL;
					while (i >= 0) {
						if (str[i] == '|') {
							str[i] = NULL;
						}
						i--;
					}
					m_ofn.lpstrFilter = str;
				}
			}
			if (pVarResult) {
				pVarResult->bstrVal = SysAllocString(m_ofn.lpstrInitialDir);
				pVarResult->vt = VT_BSTR;
			}
			return S_OK;
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
	}
	if (dispIdMember >= 40) {
		if (!m_ofn.lpstrFile) {
			m_ofn.lpstrFile = new WCHAR[m_ofn.nMaxFile];
			m_ofn.lpstrFile[0] = NULL;
		}
		BOOL bResult = FALSE;
		switch (dispIdMember) {
			case 40:
				bResult = GetOpenFileName(&m_ofn);
				break;
			case 41:
				bResult = GetSaveFileName(&m_ofn);
				break;
		}
		if (pVarResult) {
			pVarResult->boolVal = bResult ? VARIANT_TRUE : VARIANT_FALSE;
			pVarResult->vt = VT_BOOL;
		}
		return S_OK;
	}
	if (dispIdMember >= 30) {
		switch (dispIdMember) {
			case 30:
				pdw = &m_ofn.nMaxFile;
				break;
			case 31:
				pdw = &m_ofn.Flags;
				break;
			case 32:	
				pdw = &m_ofn.nFilterIndex;
				break;
			case 33:
				pdw = &m_ofn.FlagsEx;
				break;
		}
		if (nArg >= 0) {
			*pdw = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
		}
		if (pVarResult) {
			pVarResult->lVal = *pdw;
			pVarResult->vt = VT_I4;
		}
		return S_OK;
	}
	if (dispIdMember >= 20) {
		switch (dispIdMember) {
			case 20:
				lplpwstr = &m_ofn.lpstrInitialDir;
				break;
			case 21:
				lplpwstr = &m_ofn.lpstrDefExt;
				break;
			case 22:
				lplpwstr = &m_ofn.lpstrTitle;
				break;
		}
		if (nArg >= 0) {
			if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
				if (*lplpwstr) {
					delete [] *lplpwstr;
					*lplpwstr = NULL;
				}
				if (pDispParams->rgvarg[nArg].bstrVal) {
					*lplpwstr = new WCHAR[SysStringLen(pDispParams->rgvarg[nArg].bstrVal) + 1];
					lstrcpy(const_cast<LPWSTR>(*lplpwstr), pDispParams->rgvarg[nArg].bstrVal);
				}
			}
		}
		if (pVarResult) {
			pVarResult->bstrVal = SysAllocString(*lplpwstr);
			pVarResult->vt = VT_BSTR;
		}
		return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteGdiplusBitmap

CteGdiplusBitmap::CteGdiplusBitmap()
{
	m_cRef = 1;
	m_pImage = NULL;
}

CteGdiplusBitmap::~CteGdiplusBitmap()
{
	if (m_pImage) {
		delete m_pImage;
	}
}

STDMETHODIMP CteGdiplusBitmap::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
}

STDMETHODIMP_(ULONG) CteGdiplusBitmap::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteGdiplusBitmap::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}
	return m_cRef;
}

STDMETHODIMP CteGdiplusBitmap::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteGdiplusBitmap::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteGdiplusBitmap::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	return teGetDispId(methodGB, sizeof(methodGB), g_maps[MAP_GB], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteGdiplusBitmap::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	Gdiplus::Status Result = GenericError;
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	if (dispIdMember < 100) {
		if (m_pImage) {
			delete m_pImage;
			m_pImage = NULL;
		}
	}
	else if (!m_pImage) {
		return DISP_E_EXCEPTION;
	}
	switch(dispIdMember) {
		//FromHBITMAP
		case 1:	
			HPALETTE pal;
			pal = (nArg >= 1) ? (HPALETTE)GetLLFromVariant(&pDispParams->rgvarg[nArg - 1]) : 0;
			if (nArg >= 0) {
				m_pImage = Gdiplus::Bitmap::FromHBITMAP((HBITMAP)GetLLFromVariant(&pDispParams->rgvarg[nArg]), pal);
			}
			return S_OK;
		//FromHICON
		case 2:	
			if (nArg >= 0) {
				HICON hicon = (HICON)GetLLFromVariant(&pDispParams->rgvarg[nArg]);
				if (nArg >= 1) {
					ICONINFO iconinfo;
					GetIconInfo(hicon, &iconinfo);
					DIBSECTION dib;
					GetObject(iconinfo.hbmColor, sizeof(DIBSECTION), &dib);
					HIMAGELIST himl;
					himl = ImageList_Create(dib.dsBm.bmWidth, dib.dsBm.bmHeight, ILC_COLOR24 | ILC_MASK, 0, 0);
					ImageList_SetBkColor(himl, GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]));
					ImageList_AddIcon(himl, hicon);
					HICON icon2	= ImageList_GetIcon(himl, 0, ILD_NORMAL);
					ImageList_Destroy(himl);

					m_pImage = Gdiplus::Bitmap::FromHICON(icon2);
					DeleteObject(iconinfo.hbmColor);
					DeleteObject(iconinfo.hbmMask);
					DestroyIcon(icon2);
				}
				else {
					m_pImage = Gdiplus::Bitmap::FromHICON(hicon);
				}
			}
			return S_OK;
		//FromResource
		case 3:	
			if (nArg >= 1) {
				m_pImage = Gdiplus::Bitmap::FromResource(
					(HINSTANCE)GetLLFromVariant(&pDispParams->rgvarg[nArg]), 
					GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]));
			}
			return S_OK;
		//FromFile
		case 4:
			if (nArg >= 0) {
				BOOL b = FALSE;
				if (nArg) {
					b = (BOOL)GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
				}
				m_pImage = Gdiplus::Bitmap::FromFile(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), b);
//				m_pImage = Gdiplus::Image::FromFile(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), b);
			}
			return S_OK;
		//Free(99)
		//Save
		case 100:
			if (nArg >= 0) {
				VARIANT vText;
				teVariantChangeType(&vText, &pDispParams->rgvarg[nArg], VT_BSTR);
				CLSID encoderClsid;
				if (GetEncoderClsid(vText.bstrVal, &encoderClsid, NULL) >= 0) {
					EncoderParameters *pEncoderParameters = NULL;
					CteMemory *pMem = NULL;
					if (nArg && GetMemFormVariant(&pDispParams->rgvarg[nArg - 1], &pMem)) {
						pEncoderParameters = (EncoderParameters *)pMem->m_pc;
					}

					Result = m_pImage->Save(vText.bstrVal, &encoderClsid, pEncoderParameters);
					if (pMem) {
						pMem->Release();
					}
				}
			}				
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				pVarResult->lVal = Result;
			}
			return S_OK;
		//Base64
		case 101:
		//DataURI
		case 102:
			if (nArg >= 0 && lpfnCryptBinaryToStringW) {
				VARIANT vText;
				teVariantChangeType(&vText, &pDispParams->rgvarg[nArg], VT_BSTR);
				CLSID encoderClsid;
				WCHAR szMime[16];
				if (GetEncoderClsid(vText.bstrVal, &encoderClsid, szMime) >= 0) {
					IStream* oStream = NULL;
					if SUCCEEDED(CreateStreamOnHGlobal(NULL, TRUE, (LPSTREAM*)&oStream)) {
						if (m_pImage->Save(oStream, &encoderClsid) == 0) {
							ULARGE_INTEGER ulnSize;
							LARGE_INTEGER lnOffset;
							lnOffset.QuadPart = 0;
							oStream->Seek(lnOffset, STREAM_SEEK_END, &ulnSize);
							oStream->Seek(lnOffset, STREAM_SEEK_SET, NULL);
							UCHAR *pBuff = new UCHAR[ulnSize.LowPart];
							ULONG ulBytesRead;
							if SUCCEEDED(oStream->Read(pBuff, (ULONG)ulnSize.QuadPart, &ulBytesRead)) {
								WCHAR szHead[32] = L"data:";
								if (dispIdMember == 102) {
									lstrcat(szHead, szMime);
									lstrcat(szHead, L";base64,");
								}
								else {
									szHead[0] = NULL;
								}
								int nDest = lstrlen(szHead);
								DWORD dwSize;
								lpfnCryptBinaryToStringW(pBuff, ulBytesRead, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, NULL, &dwSize);
								if (dwSize > 0) {
									pVarResult->vt = VT_BSTR;
									pVarResult->bstrVal = SysAllocStringLen(NULL, nDest + dwSize - 1);
									lstrcpy(pVarResult->bstrVal, szHead);
									lpfnCryptBinaryToStringW(pBuff, ulBytesRead, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, &pVarResult->bstrVal[nDest], &dwSize);
								}
							}
							delete [] pBuff;
						}
						oStream->Release();
					}
				}
			}
			return S_OK;
		//GetWidth
		case 110:
			SetVariantLL(pVarResult, m_pImage->GetWidth());
			return S_OK;
		//GetHeight
		case 111:
			SetVariantLL(pVarResult, m_pImage->GetHeight());
			return S_OK;
		//GetThumbnailImage
		case 120:
			if (nArg >= 1) {
				CteGdiplusBitmap *pGB = new CteGdiplusBitmap();
				pGB->m_pImage = m_pImage->GetThumbnailImage(GetIntFromVariant(&pDispParams->rgvarg[nArg]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]));
				teSetObject(pVarResult, pGB);
				pGB->Release();
			}
			return S_OK;
		//RotateFlip
		case 130:
			if (nArg >= 0) {
				m_pImage->RotateFlip(static_cast<Gdiplus::RotateFlipType>(GetIntFromVariant(&pDispParams->rgvarg[nArg])));
			}
			return S_OK;
		//GetHBITMAP
		case 210:
			if (pVarResult) {
				int cl = (nArg >= 0) ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : 0;
				HBITMAP hBM;
				Gdiplus::Bitmap *pBitmap = dynamic_cast<Gdiplus::Bitmap*>(m_pImage);
				if (pBitmap && pBitmap->GetHBITMAP(((cl & 0xff) << 16) + (cl & 0xff00) + ((cl & 0xff0000) >> 16), &hBM) == 0) {
					SetVariantLL(pVarResult, (LONGLONG)hBM);
				}
			}
			return S_OK;
		//GetHICON
		case 211:
			if (pVarResult) {
				HICON hIcon;
				Gdiplus::Bitmap *pBitmap = dynamic_cast<Gdiplus::Bitmap*>(m_pImage);
				if (pBitmap->GetHICON(&hIcon) == 0) {
					SetVariantLL(pVarResult, (LONGLONG)hIcon);
				}
			}
			return S_OK;
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteWindowsAPI

CteWindowsAPI::CteWindowsAPI()
{
	m_cRef = 1;
}

CteWindowsAPI::~CteWindowsAPI()
{
}

STDMETHODIMP CteWindowsAPI::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
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
	return teGetDispId(methodAPI, sizeof(methodAPI), g_maps[MAP_API], *rgszNames, rgDispId, false);
}

STDMETHODIMP CteWindowsAPI::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	LONGLONG Result = 0;
	HANDLE *phResult = (HANDLE *)&Result;
	LONG *plResult = (LONG *)&Result;
	BOOL *pbResult = (BOOL *)&Result;
	WCHAR *pszResult = NULL;
	IUnknown *punk = NULL;

	if (wFlags == DISPATCH_PROPERTYGET) {
		if (dispIdMember != DISPID_VALUE) {
			if (pVarResult) {
				CteDispatch *pClone = new CteDispatch(this, 0);
				pClone->m_dispIdMember = dispIdMember;
				teSetObject(pVarResult, pClone);
				pClone->Release();
				return S_OK;
			}
		}
	}
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	if (dispIdMember >= 10000) {
		INT_PTR param[16];
		for (int i = 0; i < 16 && i <= nArg; i++) {
			param[i] = (INT_PTR)GetLLFromVariant(&pDispParams->rgvarg[nArg - i]);
		}
		if (dispIdMember >= 50000) {//string
			if (pVarResult) {
				LPWSTR pszName = NULL;
				switch(dispIdMember) {
					case 50005:
						if (nArg >= 0) {
							int i = GetWindowTextLength((HWND)param[0]);
							pszName = new WCHAR[i + 2];
							GetWindowText((HWND)param[0], pszName, i + 1);
							pszName[i + 1] = 0;
						}
						break;
					case 50015:
						if (nArg >= 0) {
							pszName = new WCHAR[MAX_CLASS_NAME + 1];
							int i = GetClassName((HWND)param[0], pszName, MAX_CLASS_NAME);
							pszName[i + 1] = 0;
						}
						break;
					case 50145:
						if (nArg >= 1) {
							pszName = new WCHAR[4096];
							int i = LoadString((HINSTANCE)param[0], (UINT)param[1], pszName, 4096);
							pszName[i + 1] = 0;
						}
						break;
					case 50025:
						if (nArg >= 0) {
							teGetModuleFileName((HMODULE)param[0], &pszName);
						}
						break;
					case 50035:
						if (pVarResult) {
							pVarResult->vt = VT_BSTR;
							pVarResult->bstrVal = SysAllocString(teGetCommandLine());
						}
						return S_OK;//Exception
					case 50045:
						pszName = new WCHAR[MAX_PATH];//use wsh.CurrentDirectory
						GetCurrentDirectory(MAX_PATH, pszName);
						break;
					case 50055:
						if (nArg >= 2) {
							pszName = teGetMenuString((HMENU)param[0], (UINT)param[1], param[2] != MF_BYCOMMAND);
						}
						break;
					case 50065: //GetDisplayNameOf
						if (nArg >= 1) {
							int i = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
							if (i & (SHGDN_FORPARSING | SHGDN_FORPARSINGEX) && (FindUnknown(&pDispParams->rgvarg[nArg], &punk))) {
								CteFolderItem *pFI = NULL;
								if SUCCEEDED(punk->QueryInterface(g_ClsIdFI, (LPVOID *)&pFI)) {
									if (pFI->m_v.vt == VT_BSTR) {
										VariantCopy(pVarResult, &pFI->m_v);
										pFI->Release();
										return S_OK;
									}
									pFI->Release();
								}
							}
							LPITEMIDLIST pidl;	
							if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
								GetVarPathFromPidl(pVarResult, pidl, i);
								::CoTaskMemFree(pidl);
							}
						}
						return S_OK;
					case 50075:
						if (nArg >= 0) {
							pszName = new WCHAR[MAX_PATH];
							int i = GetKeyNameText((LONG)param[0], pszName, MAX_PATH);
							pszName[i] = NULL;
						}
						break;
					//DragQueryFile
					case 50085:
						if (nArg >= 1) {
							if (param[1] & ~MAXINT) {
								if (pVarResult) {
									pVarResult->vt = VT_I4;
									pVarResult->lVal = DragQueryFile((HDROP)param[0], (UINT)param[1], NULL, 0);
								}
								return S_OK;
							}
							else {
								teDragQueryFile((HDROP)param[0], (UINT)param[1], &pszName);
							}
						}
						break;
					case 50095:
						if (nArg >= 0) {
							pVarResult->vt = VT_BSTR;
							pVarResult->bstrVal = SysAllocString(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]));
							return S_OK;
						}
						break;
					case 50105:
						if (nArg >= 1) {
							pVarResult->vt = VT_BSTR;
							pVarResult->bstrVal = SysAllocStringLen(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), (UINT)param[1]);
							return S_OK;
						}
						break;
					case 50115:
						if (nArg >= 1) {
							pVarResult->vt = VT_BSTR;
							pVarResult->bstrVal = SysAllocStringByteLen(reinterpret_cast<char*>(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg])), (UINT)param[1]);
							return S_OK;
						}
						break;
					case 50125://sprintf_s
						if (nArg >= 2) {
							try {
								int nSize = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
								LPWSTR pszFormat = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]);
								if (pszFormat && nSize > 0) {
									pszName = new WCHAR[nSize];
									pszName[0] = NULL;
									int nIndex = 1;
									int nLen = 0;
									while (++nIndex <= nArg && pszFormat[0] && nLen < nSize) {
										int nPos = 0;
										while (pszFormat[nPos] && pszFormat[nPos++] != '%') {
										}
										while (WCHAR wc = pszFormat[nPos]) {
											nPos++;
											if (wc == '%') {//Escape %
												lstrcpyn(&pszName[nLen], pszFormat, nPos);
												nLen += nPos - 1;
												pszFormat += nPos;
												nPos = 0;
												break;
											}
											if (StrChrIW(L"diuoxc", wc)) {//Integer
												wc = pszFormat[nPos];
												pszFormat[nPos] = NULL;
												swprintf_s(&pszName[nLen], nSize - nLen, pszFormat, GetLLFromVariant(&pDispParams->rgvarg[nArg - nIndex]));
												nLen += lstrlen(&pszName[nLen]);
												pszFormat += nPos;
												nPos = 0;
												pszFormat[0] = wc;
												break;
											}
											if (StrChrIW(L"efga", wc)) {//Real
												wc = pszFormat[nPos];
												pszFormat[nPos] = NULL;
												VARIANT v;
												teVariantChangeType(&v, &pDispParams->rgvarg[nArg - nIndex], VT_R8);
												swprintf_s(&pszName[nLen], nSize - nLen, pszFormat, v.dblVal);
												nLen += lstrlen(&pszName[nLen]);
												pszFormat += nPos;
												nPos = 0;
												pszFormat[0] = wc;
												break;
											}
											if (StrChrIW(L"s", wc)) {//String
												wc = pszFormat[nPos];
												pszFormat[nPos] = NULL;
												VARIANT v;
												teVariantChangeType(&v, &pDispParams->rgvarg[nArg - nIndex], VT_BSTR);
												swprintf_s(&pszName[nLen], nSize - nLen, pszFormat, v.bstrVal);
												VariantClear(&v);
												nLen += lstrlen(&pszName[nLen]);
												pszFormat += nPos;
												nPos = 0;
												pszFormat[0] = wc;
												break;
											}
											if (!StrChrIW(L"0123456789-+#.hl", wc)) {//not Specifier
												lstrcpyn(&pszName[nLen], pszFormat, nPos + 1);
												nLen += nPos;
												pszFormat += nPos;
												nPos = 0;
												break;
											}
										}
									}
									if (nLen < nSize) {
										lstrcpyn(&pszName[nLen], pszFormat, nSize - nLen);
									}
								}
							} catch (...) {
								return E_UNEXPECTED;
							}
						}
						break;
					case 50135://base64_encode
						if (nArg >= 0) {
							if (lpfnCryptBinaryToStringW) {
								UCHAR *pc;
								int nLen;
								GetpDataFromVariant(&pc, &nLen, &pDispParams->rgvarg[nArg]);
								if (pc) {
									DWORD dwSize;
									lpfnCryptBinaryToStringW(pc, nLen, CRYPT_STRING_BASE64, NULL, &dwSize);
									if (dwSize > 0) {
										pVarResult->vt = VT_BSTR;
										pVarResult->bstrVal = SysAllocStringLen(NULL, dwSize - 1);
										lpfnCryptBinaryToStringW(pc, nLen, CRYPT_STRING_BASE64, pVarResult->bstrVal, &dwSize);
									}
								}
							}
						}
						break;
					default:				
						return DISP_E_MEMBERNOTFOUND;
				}
				if (pszName) {
					pVarResult->vt = VT_BSTR;
					pVarResult->bstrVal = SysAllocString(pszName);
					delete [] pszName;
				}
			}
			return S_OK;
		}
		if (dispIdMember >= 40000) {//Handle
			switch(dispIdMember) {
				case 40003:
					if (nArg >= 2) {
						*phResult = ImageList_GetIcon((HIMAGELIST)param[0], (int)param[1], (UINT)param[2]);
					}
					break;
				case 40013:
					if (nArg >= 4) {
						*phResult = ImageList_Create((int)param[0], (int)param[1], (UINT)param[2], (int)param[3], (int)param[4]);
					}
					break;
				case 40023:
					if (nArg >= 1) {
						*phResult = (HANDLE)GetWindowLongPtr((HWND)param[0], (int)param[1]);
					}
					break;
				case 40033:
					if (nArg >= 1) {
						*phResult = (HANDLE)GetClassLongPtr((HWND)param[0], (int)param[1]);
					}
					break;
				case 40043:
					if (nArg >= 1) {
						*phResult = GetSubMenu((HMENU)param[0], (int)param[1]);
					}
					break;
				case 40053:
					if (nArg >= 1) {
						*phResult = SelectObject((HDC)param[0], (HGDIOBJ)param[1]);
					}
					break;
				case 40063:
					if (nArg >= 0) {
						*phResult = GetStockObject((int)param[0]);
					}
					break;
				case 40073:
					if (nArg >= 0) {
						*phResult = GetSysColorBrush((int)param[0]);
					}
					break;
				case 40083:
					if (nArg >= 0) {
						*phResult = SetFocus((HWND)param[0]);
					}
					break;
				case 40093:
					if (nArg >= 0) {
						*phResult = GetDC((HWND)param[0]);
					}
					break;
				case 40103:
					if (nArg >= 0) {
						*phResult = CreateCompatibleDC((HDC)param[0]);
					}
					break;
				case 40113:
					*phResult = CreatePopupMenu();
					break;
				case 40123:
					*phResult = CreateMenu();
					break;
				case 40133:
					if (nArg >= 2) {
						*phResult = CreateCompatibleBitmap((HDC)param[0], (int)param[1], (int)param[2]);
					}
					break;
				case 40143:
					if (nArg >= 2) {
						*phResult = (HANDLE)SetWindowLongPtr((HWND)param[0], (int)param[1], (LONG_PTR)param[2]);
					}
					break;
				case 40153:
					if (nArg >= 2) {
						*phResult = (HANDLE)SetClassLongPtr((HWND)param[0], (int)param[1], (LONG_PTR)param[2]);
					}
				case 40163:
					if (nArg >= 0) {
						*phResult = ImageList_Duplicate((HIMAGELIST)param[0]);
					}
					break;
				case 40173:
					if (nArg >= 1) {
						*phResult = (HANDLE)SendMessage((HWND)param[0], (UINT)param[1], (WPARAM)param[2], (LPARAM)param[3]);
					}
					break;
				case 40183:
					if (nArg >= 1) {
						*phResult = (HMENU)GetSystemMenu((HWND)param[0], (BOOL)param[1]);
					}
					break;
				case 40193:
					if (nArg >= 0) {
						*phResult = GetWindowDC((HWND)param[0]);
					}
					break;
				case 40203:
					if (nArg >= 2) {
						*phResult = CreatePen((int)param[0], (int)param[1], (COLORREF)param[2]);
					}
					break;
				case 40213:
					if (nArg >= 0) {
						*phResult = SetCapture((HWND)param[0]);
					}
					break;
				case 40223:
					if (nArg >= 0) {
						*phResult = SetCursor((HCURSOR)param[0]);
					}
					break;
				case 40233:
					if (nArg >= 4) {
						*phResult = (HANDLE)CallWindowProc((WNDPROC)param[0], (HWND)param[1], (UINT)param[2], (WPARAM)param[3], (LPARAM)param[4]);
					}
					break;					
				case 40243:
					if (nArg >= 1) {
						*phResult = (HANDLE)GetWindow((HWND)param[0], (UINT)param[1]);
					}
					break;					
				case 40253:
					if (nArg >= 0) {
						*phResult = (HANDLE)GetTopWindow((HWND)param[0]);
					}
					break;

				case 40263:
					if (nArg >= 2) {
						*phResult = (HANDLE)OpenProcess((DWORD)param[0], (BOOL)param[1], (DWORD)param[2]);
					}
					break;
				//other
				case 45003:
					if (nArg >= 1) {
						*phResult = LoadMenu((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]));
					}
					break;
				case 45013:
					if (nArg >= 1) {
						*phResult = LoadIcon((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]));
					}
					break;
				case 45023:
					if (nArg >= 2) {
						VARIANT vFile;
						teVariantChangeType(&vFile, &pDispParams->rgvarg[nArg], VT_BSTR);
						*phResult = LoadLibraryEx(vFile.bstrVal, (HANDLE)param[1], (DWORD)param[2]);
						VariantClear(&vFile);
					}
					break;
				case 45033:
					if (nArg >= 5) {
						*phResult = LoadImage((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
							(UINT)param[2],	(int)param[3], (int)param[4], (UINT)param[5]);
					}
					break;
				case 45043:
					if (nArg >= 6) {
						*phResult = ImageList_LoadImage((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
						(int)param[2], (int)param[3], (COLORREF)param[4], (UINT)param[5], (UINT)param[6]);
					}
					break;
				case 45053://SHGetFileInfo
					if (nArg >= 4) {
						VARIANT vPath;
						BSTR bs = NULL;
						if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
							GetPidlFromVariant((LPITEMIDLIST *)&bs, &pDispParams->rgvarg[nArg]);
						}
						else {
							teVariantChangeType(&vPath, &pDispParams->rgvarg[nArg], VT_BSTR);
							bs = vPath.bstrVal;
						}
						*phResult = (HANDLE)SHGetFileInfo(bs, (DWORD)param[1], 
							(SHFILEINFO *)GetpcFormVariant(&pDispParams->rgvarg[nArg - 2]), 
							(UINT)param[3], (UINT)param[4]);
						if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
							::CoTaskMemFree(bs);
						}
						else {
							VariantClear(&vPath);
						}
					}
					break;
				case 45083://CreateWindowEx
					if (nArg >= 11) {
						VARIANT vClass, vWindow;
						teVariantChangeType(&vClass, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
						teVariantChangeType(&vWindow, &pDispParams->rgvarg[nArg - 2], VT_BSTR);
						*phResult = CreateWindowEx((DWORD)param[0], vClass.bstrVal, vWindow.bstrVal,
							(DWORD)param[3], (int)param[4], (int)param[5], (int)param[6], (int)param[7],
							(HWND)param[8], (HMENU)param[9], (HINSTANCE)param[10], 
							(LPVOID)GetpcFormVariant(&pDispParams->rgvarg[nArg - 11])); 
					}
					break;
				case 45093://ShellExecute
					if (nArg >= 5) {
						VARIANT vFile, vParam, vDir;
						teVariantChangeType(&vFile, &pDispParams->rgvarg[nArg - 2], VT_BSTR);
						teVariantChangeType(&vParam, &pDispParams->rgvarg[nArg - 3], VT_BSTR);
						teVariantChangeType(&vDir, &pDispParams->rgvarg[nArg - 4], VT_BSTR);
						*phResult = ShellExecute((HWND)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
							vFile.bstrVal, vParam.bstrVal, vDir.bstrVal, (int)param[5]);
					}
					break;
				case 45103:
					if (nArg >= 1) {
						*phResult = BeginPaint((HWND)param[0], (LPPAINTSTRUCT)GetpcFormVariant(&pDispParams->rgvarg[nArg - 1]));
					}
					break;
				case 45113:
					if (nArg >= 1) {
						*phResult = LoadCursor((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]));
					}
					break;
				case 45123:
					if (nArg >= 0) {
						*phResult = LoadCursorFromFile(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]));
					}
					break;
				case 45133://SHParseDisplayName
					if (nArg >= 1) {
						TEInvoke *pInvoke = new TEInvoke[1];
						if (GetDispatch(&pDispParams->rgvarg[nArg], &pInvoke->pdisp)) {
							pInvoke->dispid = DISPID_VALUE;
							pInvoke->cArgs = nArg;
							pInvoke->pv = GetNewVARIANT(nArg);
							for (int i = nArg; i--;) {
								VariantCopy(&pInvoke->pv[i], &pDispParams->rgvarg[i]);
							}
							if (pInvoke->pv[pInvoke->cArgs - 1].vt == VT_BSTR) {
								*phResult = (HANDLE)_beginthread(&threadParseDisplayName, 0, pInvoke);
							}
							else {
								GetPidlFromVariant((LPITEMIDLIST *)&pInvoke->pResult, &pInvoke->pv[pInvoke->cArgs - 1]);
								ParseInvoke(pInvoke);
								*phResult = 0;								
							}
						}
						else {
							delete [] pInvoke;
						}
					}
					break;
				default:
					return DISP_E_MEMBERNOTFOUND;
			}
			SetReturnValue(pVarResult, dispIdMember, &Result);
			return S_OK;
		}
		if (dispIdMember >= 30000) {//int
			switch(dispIdMember) {
				case 30002:
					if (nArg >= 1) {
						*plResult = ImageList_SetBkColor((HIMAGELIST)param[0], (COLORREF)param[1]);
					}
					break;
				case 30012:
					if (nArg >= 1) {
						*plResult = ImageList_AddIcon((HIMAGELIST)param[0], (HICON)param[1]);
					}
					break;
				case 30022:
					if (nArg >= 2) {
						*plResult = ImageList_Add((HIMAGELIST)param[0], (HBITMAP)param[1], (HBITMAP)param[2]);
					}
					break;
				case 30032:
					if (nArg >= 2) {
						*plResult = ImageList_AddMasked((HIMAGELIST)param[0], (HBITMAP)param[1], (COLORREF)param[2]);
					}
					break;
				case 30042:
					if (nArg >= 0) {
						*plResult = GetKeyState((int)param[0]);
					}
					break;
				case 30052:
					if (nArg >= 0) {
						*plResult = GetSystemMetrics((int)param[0]);
					}
					break;
				case 30062:
					if (nArg >= 0) {
						*plResult = GetSysColor((int)param[0]);
					}
					break;
				case 30082:
					if (nArg >= 0) {
						*plResult = GetMenuItemCount((HMENU)param[0]);
					}
					break;
				case 30092:
					if (nArg >= 0) {
						*plResult = ImageList_GetImageCount((HIMAGELIST)param[0]);
					}
					break;
				case 30102:
					if (nArg >= 1) {
						*plResult = ReleaseDC((HWND)param[0], (HDC)param[1]);
					}
					break;
				case 30112:
					if (nArg >= 1) {
						*plResult = GetMenuItemID((HMENU)param[0], (int)param[1]);
					}
					break;
				case 30122:
					if (nArg >= 3) {
						*plResult = ImageList_Replace((HIMAGELIST)param[0], (int)param[1],
							(HBITMAP)param[2], (HBITMAP)param[3]);
					}
					break;
				case 30132:
					if (nArg >= 2) {
						*plResult = ImageList_ReplaceIcon((HIMAGELIST)param[0], (int)param[1],
							(HICON)param[2]);
					}
					break;
				case 30142:
					if (nArg >= 1) {
						*plResult = SetBkMode((HDC)param[0], (int)param[1]);
					}
					break;
				case 30152:
					if (nArg >= 1) {
						*plResult = SetBkColor((HDC)param[0], (COLORREF)param[1]);
					}
					break;
				case 30162:
					if (nArg >= 1) {
						*plResult = SetTextColor((HDC)param[0], (COLORREF)param[1]);
					}
					break;
				case 30172:
					if (nArg >= 1) {
						*plResult = MapVirtualKey((UINT)param[0], (UINT)param[1]);
					}
					break;
				case 30182:
					if (nArg >= 1) {
						*plResult = WaitForInputIdle((HANDLE)param[0], (DWORD)param[1]);
					}
					break;
				case 30192:
					if (nArg >= 1) {
						*plResult = WaitForSingleObject((HANDLE)param[0], (DWORD)param[1]);
					}
					break;
				case 30202:
					if (nArg >= 2) {
						*plResult = GetMenuDefaultItem((HMENU)param[0], (UINT)param[1], (UINT)param[2]);
					}
					break;
				case 30212://CRC32
					if (nArg >= 0) {
						BYTE *pc;
						int nLen;
						GetpDataFromVariant(&pc, &nLen, &pDispParams->rgvarg[nArg]);
						*plResult = CalcCrc32(pc, nLen, ((nArg >= 1) ? (UINT)param[1] : 0));
					}
					break;
				case 35002://TrackPopupMenuEx
					if (nArg >= 5) {
						if (nArg >= 6) {
							if (FindUnknown(&pDispParams->rgvarg[nArg - 6], &punk)) {
								punk->QueryInterface(g_ClsIdContextMenu, (LPVOID *)&g_pCCM);
							}
						}
						g_hMenuKeyHook = g_hMenuKeyHook ? g_hMenuKeyHook : SetWindowsHookEx(WH_KEYBOARD, (HOOKPROC)MenuKeyProc, hInst, GetCurrentThreadId());
						*plResult = TrackPopupMenuEx((HMENU)param[0], (UINT)param[1], (int)param[2], (int)param[3],
							(HWND)param[4], (LPTPMPARAMS)GetpcFormVariant(&pDispParams->rgvarg[nArg - 5]));
						UnhookWindowsHookEx(g_hMenuKeyHook);
						g_hMenuKeyHook = NULL;
						if (g_pCCM) {
							g_pCCM->Release();
							g_pCCM = NULL;
						}
					}
					break;
				case 35012://ExtractIconEx
					if (nArg >= 4) {
						*plResult = ExtractIconEx(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), (int)param[1], 
							(HICON *)GetpcFormVariant(&pDispParams->rgvarg[nArg - 2]), 
							(HICON *)GetpcFormVariant(&pDispParams->rgvarg[nArg - 3]), 
							(UINT)param[4]);
					}
					break;
				case 35022://GetObject
					if (nArg >= 2) {
						*plResult = GetObject((HWND)param[0], (int)param[1], GetpcFormVariant(&pDispParams->rgvarg[nArg - 2]));
					}
					break;
				case 35032:
					if (nArg >= 5) {
						*plResult = MultiByteToWideChar((UINT)param[0], (DWORD)param[1], 
							(LPCSTR)GetpcFormVariant(&pDispParams->rgvarg[nArg - 2]), (int)param[3],
							(LPWSTR)GetpcFormVariant(&pDispParams->rgvarg[nArg - 4]), (int)param[5]);
					}
					break;
				case 35042:
					if (nArg >= 7) {
						VARIANT v;
						teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 2], VT_BSTR);
						*plResult = WideCharToMultiByte((UINT)param[0], (DWORD)param[1], v.bstrVal, (int)param[3], 
							GetpcFormVariant(&pDispParams->rgvarg[nArg - 4]), (int)param[5], 
							(LPCSTR)GetpcFormVariant(&pDispParams->rgvarg[nArg - 6]),
							(LPBOOL)GetpcFormVariant(&pDispParams->rgvarg[nArg - 7]));
						VariantClear(&v);
					}
					break;
				case 35052:
					if (nArg >= 1) {
						*plResult = (ULONG)param[1]; 
						IDataObject *pDataObj;
						if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
							LPITEMIDLIST *ppidllist;
							long nCount;
							ppidllist = IDListFormDataObj(pDataObj, &nCount);
							AdjustIDList(ppidllist, nCount);
							if (nCount >= 1) {
								IShellFolder *pSF;
								if (GetShellFolder(&pSF, ppidllist[0])) {
									pSF->GetAttributesOf(nCount, (LPCITEMIDLIST *)&ppidllist[1], (SFGAOF *)(plResult));
									*plResult &= (ULONG)param[1]; 
									pSF->Release();
								}
							}
							while (--nCount >= 0) {
								::CoTaskMemFree(ppidllist[nCount]);
							}
							delete [] ppidllist;
							pDataObj->Release();
						}
					}
					break;
				case 35062://DoDragDrop
					if (nArg >= 2) {
						IDataObject *pDataObj;
						if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
							g_pTE->m_bDrop = true;
							try {
								*plResult = ::DoDragDrop(pDataObj, 
									static_cast<IDropSource *>(g_pTE), (DWORD)param[1],
									(LPDWORD)GetpcFormVariant(&pDispParams->rgvarg[nArg - 2]));
							} catch(...) {}
							pDataObj->Release();
							pDataObj = NULL;
						}
					}
					break;
				case 35072://CompareIDs
					if (nArg >= 2) {
						LPITEMIDLIST pidl1, pidl2;
						if (GetPidlFromVariant(&pidl1, &pDispParams->rgvarg[nArg - 1])) {
							if (GetPidlFromVariant(&pidl2, &pDispParams->rgvarg[nArg - 2])) {
								*plResult = g_pDesktopFolder->CompareIDs(param[0], pidl1, pidl2);
								::CoTaskMemFree(pidl2);
							}
							::CoTaskMemFree(pidl1);
						}
					}
					break;
				//GetScriptDispatch
				case 35080:
				//ExecScript
				case 35082:
					if (nArg >= 1) {
						VARIANT v[2];
						for (int i = 1; i >= 0; i--) {
							teVariantChangeType(&v[i], &pDispParams->rgvarg[nArg - i], VT_BSTR);
						}
						VARIANT *pv = NULL;
						if (nArg >= 2) {
							pv = &pDispParams->rgvarg[nArg - 2];
						}
						IUnknown *pOnError = NULL;
						if (nArg >= 3) {
							FindUnknown(&pDispParams->rgvarg[nArg - 3], &pOnError);
						}
						IDispatch *pdisp = NULL;
						IDispatch **ppdisp = pVarResult && (dispIdMember == 35080) ? &pdisp : NULL;
						*plResult = ParseScript(v[0].bstrVal, v[1].bstrVal, pv, pOnError, ppdisp);
						if (pdisp) {
							teSetObject(pVarResult, pdisp);
							pdisp->Release();
						}
					}
					break;
				//GetDispatch
				case 35090:
					if (nArg >= 1 && pVarResult) {
						IDispatch *pdisp;
						if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
							VARIANT v;
							teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
							DISPID dispid = DISPID_UNKNOWN;
							if (pdisp->GetIDsOfNames(IID_NULL, &v.bstrVal, 1, LOCALE_USER_DEFAULT, &dispid) == S_OK) {
								CteDispatch *oDisp = new CteDispatch(pdisp, 0);
								oDisp->m_dispIdMember = dispid;
								teSetObject(pVarResult, oDisp);
								oDisp->Release();
							}
							pdisp->Release();
						}
					}
					return S_OK;
				default:
					return DISP_E_MEMBERNOTFOUND;
			}
			SetReturnValue(pVarResult, dispIdMember, &Result);
			return S_OK;
		}
		if (dispIdMember == 29002) {//ShFileOperation
			if (nArg >= 4) {
				LPSHFILEOPSTRUCT pFO = new SHFILEOPSTRUCT[1];
				::ZeroMemory(pFO, sizeof(SHFILEOPSTRUCT));
				VARIANT v[2];
				for (int i = 1; i >= 0; i--) {
					teVariantChangeType(&v[i], &pDispParams->rgvarg[nArg - 1 - i], VT_BSTR);
				}
				pFO->wFunc = (UINT)param[0];
				int nLen = SysStringLen(v[0].bstrVal);
				BSTR bs = SysAllocStringLen(v[0].bstrVal, nLen + 2);
				bs[++nLen] = 0;
				pFO->pFrom = bs;
				nLen = SysStringLen(v[1].bstrVal);
				bs = SysAllocStringLen(v[1].bstrVal, nLen + 2);
				bs[++nLen] = 0;
				pFO->pTo = bs;
				pFO->fFlags = (FILEOP_FLAGS)param[3];
				if (param[4]) {
					SetVariantLL(pVarResult, (LONGLONG)_beginthread(&threadFileOperation, 0, pFO));
				}
				else {
					try {
						*plResult = ::SHFileOperation(pFO);
					} catch (...) {}
					::SysFreeString((BSTR)(pFO->pTo));
					::SysFreeString((BSTR)(pFO->pFrom));
				}
				SetReturnValue(pVarResult, dispIdMember, &Result);
				return S_OK;
			}
		}
		if (dispIdMember >= 29000) {//BOOL MEM()
			int i = (dispIdMember / 100) % 10;
			if (nArg >= i) {
				char *pc = GetpcFormVariant(&pDispParams->rgvarg[nArg - i]);
				if (pc) {
					switch (dispIdMember) {
						case 29001:
							if (nArg >= 4) {
								*pbResult = PeekMessage((LPMSG)pc, (HWND)param[1], (UINT)param[2], (UINT)param[3], (UINT)param[4]);
							}
							break;
						case 29002:
							*plResult = ::SHFileOperation((LPSHFILEOPSTRUCT)pc);
							break;
						case 29011:
							if (nArg >= 3) {
								*pbResult = GetMessage((LPMSG)pc, (HWND)param[1], (UINT)param[2], (UINT)param[3]);
							}
							break;
						case 29101:
							*pbResult = GetWindowRect((HWND)param[0], (LPRECT)pc);
							break;
						case 29102:
							*plResult = GetWindowThreadProcessId((HWND)param[0], (LPDWORD)pc);
							break;			
						case 29111:
							*pbResult = GetClientRect((HWND)param[0], (LPRECT)pc);
							break;
						case 29112:
							if (nArg >= 2) {
								*plResult = SendInput((UINT)param[0], (LPINPUT)pc, (int)param[2]);
							}
							break;
						case 29121:
							*pbResult = ScreenToClient((HWND)param[0], (LPPOINT)pc);
							break;
						case 29122:
							if (nArg >= 4) {
								*plResult = MsgWaitForMultipleObjectsEx((DWORD)param[0], (LPHANDLE)pc, (DWORD)param[2], (DWORD)param[3], (DWORD)param[4]);
							}
							break;
						case 29131:
							*pbResult = ClientToScreen((HWND)param[0], (LPPOINT)pc);
							break;
						case 29141:
							*pbResult = GetIconInfo((HICON)param[0], (PICONINFO)pc);
							break;
						case 29151:
							*pbResult = FindNextFile((HANDLE)param[0], (LPWIN32_FIND_DATAW)pc);
							break;
						case 29161:
							if (nArg >= 2) {
								*pbResult = FillRect((HDC)param[0], (LPRECT)pc, (HBRUSH)param[2]);
							}
							break;
						case 29171:
							*pbResult = Shell_NotifyIcon((DWORD)param[0], (PNOTIFYICONDATA)pc);
							break;
						case 29191://EndPaint
							if (nArg >= 1) {
								PAINTSTRUCT *ps = (LPPAINTSTRUCT)pc;
								*pbResult = EndPaint((HWND)param[0], ps);
							}
							break;
						case 29201:
							if (nArg >= 3) {
								*pbResult = SystemParametersInfo((UINT)param[0], (UINT)param[1], (PVOID)pc, (UINT)param[3]);
							}
							break;
						case 29211:
							*pbResult = ImageList_GetIconSize((HIMAGELIST)param[0], 
								(int *)GetpcFormVariant(&pDispParams->rgvarg[nArg - 1]), (int *)pc);
							break;
						case 29221:
							LPCWSTR lpString; 
							lpString = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]); 
							*pbResult = GetTextExtentPoint32((HDC)param[0], lpString, lstrlen(lpString), (LPSIZE)pc);
							break;
						case 29232:
							if (nArg >= 3) {
								LPITEMIDLIST pidl;
								GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg]);
								if (pidl) {
									*plResult = GetDataFromPidl(pidl, (int)param[1], (PVOID)pc, (int)param[3]);
								}
							}
							break;
						case 29301:
							*pbResult = InsertMenuItem((HMENU)param[0], (UINT)param[1], (BOOL)param[2], (LPCMENUITEMINFO)pc);
							break;
						case 29311:
							*pbResult = GetMenuItemInfo((HMENU)param[0], (UINT)param[1], (BOOL)param[2], (LPMENUITEMINFO)pc);
							break;
						case 29321:
							*pbResult = SetMenuItemInfo((HMENU)param[0], (UINT)param[1], (BOOL)param[2], (LPMENUITEMINFO)pc);
							break;
						default:
							return DISP_E_MEMBERNOTFOUND;
					}
				}
			}
			SetReturnValue(pVarResult, dispIdMember, &Result);
			return S_OK;
		}

		if (dispIdMember >= 20000) {//BOOL
			switch (dispIdMember) {
				case 20001:
					if (nArg >= 3) {
						*pbResult = DrawIcon((HDC)param[0], (int)param[1], (int)param[2], (HICON)param[3]);
					}
					break;
				case 20011:
					if (nArg >= 0) {
						*pbResult = DestroyMenu((HMENU)param[0]);
					}
					break;
				case 20021:
					if (nArg >= 0) {
						*pbResult = FindClose((HANDLE)param[0]);
					}
					break;
				case 20031:
					if (nArg >= 0) {
						*pbResult = FreeLibrary((HMODULE)param[0]);
					}
					break;
				case 20041:
					if (nArg >= 0) {
						*pbResult = ImageList_Destroy((HIMAGELIST)param[0]);
					}
					break;
				case 20051:
					if (nArg >= 0) {
						*pbResult = DeleteObject((HGDIOBJ)param[0]);
					}
					break;
				case 20061:
					if (nArg >= 0) {
						*pbResult = DestroyIcon((HICON)param[0]);
					}
					break;
				case 20071:
					if (nArg >= 0) {
						*pbResult = IsWindow((HWND)param[0]);
					}
					break;
				case 20081:
					if (nArg >= 0) {
						*pbResult = IsWindowVisible((HWND)param[0]);
					}
					break;
				case 20091:
					if (nArg >= 0) {
						*pbResult = IsZoomed((HWND)param[0]);
					}
					break;
				case 20101:
					if (nArg >= 0) {
						*pbResult = IsIconic((HWND)param[0]);
					}
					break;
				case 20111:
					if (nArg >= 0) {
						*pbResult = OpenIcon((HWND)param[0]);
					}
					break;
				case 20121:
					if (nArg >= 0) {
						*pbResult = SetForegroundWindow((HWND)param[0]);
					}
					break;
				case 20131:
					if (nArg >= 0) {
						*pbResult = BringWindowToTop((HWND)param[0]);
					}
					break;
				case 20141:
					if (nArg >= 0) {
						*pbResult = DeleteDC((HDC)param[0]);
					}
					break;
				case 20151:
					if (nArg >= 0) {
						*pbResult = CloseHandle((HANDLE)param[0]);
					}
					break;
				case 20161:
					if (nArg >= 0) {
						*pbResult = IsMenu((HMENU)param[0]);
					}
					break;
				case 20171:
					if (nArg >= 5) {
						*pbResult = MoveWindow((HWND)param[0], (int)param[1], (int)param[2], (int)param[3], (int)param[4], (BOOL)param[5]);
					}
					break;
				case 20181:
					if (nArg >= 4) {
						*pbResult = SetMenuItemBitmaps((HMENU)param[0], (UINT)param[1], (UINT)param[2], (HBITMAP)param[3], (HBITMAP)param[4]);
					}
					break;
				case 20191:
					if (nArg >= 1) {
						g_nCmdShow = (int)param[1];
						*pbResult = ShowWindow((HWND)param[0], g_nCmdShow);
					}
					break;
				case 20201:
					if (nArg >= 2) {
						*pbResult = DeleteMenu((HMENU)param[0], (UINT)param[1], (UINT)param[2]);
					}
					break;
				case 20211:
					if (nArg >= 2) {
						*pbResult = RemoveMenu((HMENU)param[0], (UINT)param[1], (UINT)param[2]);
					}
					break;
				case 20231:
					if (nArg >= 8) {
						*pbResult = DrawIconEx((HDC)param[0], (int)param[1], (int)param[2], (HICON)param[3],
							(int)param[4], (int)param[5], (UINT)param[6], (HBRUSH)param[7], (UINT)param[8]);
					}
					break;
				case 20241:
					if (nArg >= 2) {
						*pbResult = EnableMenuItem((HMENU)param[0], (UINT)param[1], (UINT)param[2]);
					}
					break;
				case 20251:
					if (nArg >= 1) {
						*pbResult = ImageList_Remove((HIMAGELIST)param[0], (int)param[1]);
					}
					break;
				case 20261:
					if (nArg >= 2) {
						*pbResult = ImageList_SetIconSize((HIMAGELIST)param[0], (int)param[1], (int)param[2]);
					}
					break;
				case 20271:
					if (nArg >= 5) {
						*pbResult = ImageList_Draw((HIMAGELIST)param[0], (int)param[1], (HDC)param[2],
							(int)param[3], (int)param[4], (UINT)param[5]);
					}
					break;
				case 20281:
					if (nArg >= 9) {
						*pbResult = ImageList_DrawEx((HIMAGELIST)param[0], (int)param[1], (HDC)param[2],
							(int)param[3], (int)param[4], (int)param[5], (int)param[6],
							(COLORREF)param[7], (COLORREF)param[8], (UINT)param[9]);
					}
					break;
				case 20291:
					if (nArg >= 1) {
						*pbResult = ImageList_SetImageCount((HIMAGELIST)param[0], (UINT)param[1]);
					}
					break;
				case 20301:
					if (nArg >= 2) {
						*pbResult = ImageList_SetOverlayImage((HIMAGELIST)param[0], (int)param[1], (int)param[2]);
					}
					break;
				case 20321:
					if (nArg >= 4) {
						*pbResult = ImageList_Copy((HIMAGELIST)param[0], (int)param[1], 
							(HIMAGELIST)param[2], (int)param[3], (UINT)param[4]);
					}
					break;
				case 20331:
					if (nArg >= 0) {
						*pbResult = DestroyWindow((HWND)param[0]);
					}
					break;
				case 20341:
					if (nArg >= 2) {
						*pbResult = LineTo((HDC)param[0], (int)param[1], (int)param[2]); 
					}
					break;
				case 20351:
					*pbResult = ReleaseCapture();
					break;
				case 20361:
					if (nArg >= 1) {
						*pbResult = SetCursorPos((int)param[0], (int)param[1]);
					}
				case 20371:
					if (nArg >= 0) {
						*pbResult = DestroyCursor((HCURSOR)param[0]);
					}
					break;
				case 25001://InsertMenu
					if (nArg >= 4) {
						VARIANT vNewItem;
						teVariantChangeType(&vNewItem, &pDispParams->rgvarg[nArg - 4], VT_BSTR);
						*pbResult = InsertMenu((HMENU)param[0], (UINT)param[1], (UINT)param[2],
							(UINT_PTR)param[3], vNewItem.bstrVal);
						VariantClear(&vNewItem);
					}
					break;
				case 25011://SetWindowText
					if (nArg >= 1) {
						if (!g_nLockUpdate || (HWND)param[0] != g_hwndMain) {
							VARIANT v;
							teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
							*pbResult = SetWindowText((HWND)param[0], v.bstrVal);
							VariantInit(&v);
						}
					}
					break;
				case 25021:
					if (nArg >= 3) {
						*pbResult = RedrawWindow((HWND)param[0], (LPRECT)GetpcFormVariant(&pDispParams->rgvarg[nArg - 1]), 
							(HRGN)param[2], (UINT)param[3]);
					}
					break;
				case 25031:
					if (nArg >= 3) {
						*pbResult = MoveToEx((HDC)param[0], (int)param[1], (int)param[2], (LPPOINT)GetpcFormVariant(&pDispParams->rgvarg[nArg - 3]));
					}
					break;
				case 25041:
					if (nArg >= 2) {
						*pbResult = InvalidateRect((HWND)param[0], (LPRECT)GetpcFormVariant(&pDispParams->rgvarg[nArg - 1]), (BOOL)param[2]);
					}
					break;
				case 25051:
					if (nArg >= 3) {
						*pbResult = SendNotifyMessage((HWND)param[0], (UINT)param[1], (WPARAM)param[2], (LPARAM)param[3]);
					}
					break;
				default:
					return DISP_E_MEMBERNOTFOUND;
			}
			SetReturnValue(pVarResult, dispIdMember, &Result);
			return S_OK;
		}
		//No return value
		switch(dispIdMember) {
			case 10000:
				if (nArg >= 1) {
					PostMessage((HWND)param[0], (UINT)param[1], (WPARAM)param[2], (LPARAM)param[3]);
				}
				break;
			case 10010:
				if (nArg >= 0) {
					SHFreeNameMappings((HANDLE)param[0]);
				}
				break;
			case 10020:
				if (nArg >= 0) {
					::CoTaskMemFree((LPVOID)param[0]);
				}
				break;
			case 10030:
				if (nArg >= 0) {
					Sleep((DWORD)param[0]);
				}
				break;
			case 10040:
				if (nArg >= 5) {
					if (lpfnSHRunDialog) {
						VARIANTARG va[3];
						for (int i = 0; i < 3; i++) {
							teVariantChangeType(&va[i], &pDispParams->rgvarg[nArg - i - 2], VT_BSTR);
						}
						lpfnSHRunDialog((HWND)param[0], (HICON)param[1], va[0].bstrVal, va[1].bstrVal, va[2].bstrVal, (DWORD)param[5]);
					}
				}
				break;
			case 10050:
				if (nArg >= 1) {
					DragAcceptFiles((HWND)param[0], (BOOL)param[1]);
				}
				break;
			case 10060:
				if (nArg >= 0) {
					DragFinish((HDROP)param[0]);
				}
				break;
			case 10070:
				if (nArg >= 1) {
					mouse_event((DWORD)param[0], (DWORD)param[1], (DWORD)param[2], (DWORD)param[3], (ULONG_PTR)param[4]); 
				}
				break;
			default:
				return DISP_E_MEMBERNOTFOUND;
		}
		return S_OK;
	}	
	if (dispIdMember >= 7000) {//(mem)
		if (nArg >= 0) {
			CteMemory *pMem;
			if (GetMemFormVariant(&pDispParams->rgvarg[nArg], &pMem)) {
				switch(dispIdMember) {
					//BOOL
					case 7001:
						*pbResult = GetCursorPos((LPPOINT)pMem->m_pc);
						break;
					case 7011:
						*pbResult = GetKeyboardState((PBYTE)pMem->m_pc);
						break;
					case 7021:
						*pbResult = SetKeyboardState((PBYTE)pMem->m_pc);
						break;
					case 7031:
						*pbResult = GetVersionEx((LPOSVERSIONINFO)pMem->m_pc);
						break;
					case 7041:
						*pbResult = ChooseFont((LPCHOOSEFONT)pMem->m_pc);
						break;
					case 7051:
						try {
							*pbResult = ::ChooseColor((LPCHOOSECOLOR)pMem->m_pc);
						} catch (...) {}
						break;
					case 7061:
						*pbResult = ::TranslateMessage((LPMSG)pMem->m_pc);
						break;
					case 7071:
						*pbResult = ::ShellExecuteEx((SHELLEXECUTEINFO *)pMem->m_pc);
						break;
					//int
					//Handle
					case 9003:
						*phResult = ::CreateFontIndirect((PLOGFONT)pMem->m_pc);
						break;
					case 9013:
						*phResult = ::CreateIconIndirect((PICONINFO)pMem->m_pc);
						break;
					case 9023:
						*phResult = ::FindText((LPFINDREPLACE)pMem->m_pc);
						break;
					case 9033:
						*phResult = ::ReplaceText((LPFINDREPLACE)pMem->m_pc);
						break;
					case 9043:
						*phResult = (HANDLE)::DispatchMessage((LPMSG)pMem->m_pc);
						break;
					default:
						pMem->Release();
						return DISP_E_MEMBERNOTFOUND;
				}
				pMem->Release();
			}
			else {
				return DISP_E_TYPEMISMATCH;
			}
		}
		SetReturnValue(pVarResult, dispIdMember, &Result);
		return S_OK;
	}
	if (dispIdMember >= 6000) {//(string)
		if (nArg >= 0) {
			VARIANT v;
			teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
			switch(dispIdMember) {
				case 6000://ILCreateFromPath
					LPITEMIDLIST pidl;
					GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg]);
					if (pidl) {
						FolderItem *pid;
						if (GetFolderItemFromPidl(&pid, pidl)) {
							teSetObject(pVarResult, pid);
							pid->Release();
						}
						::CoTaskMemFree(pidl);
					}
					break;
				case 6010://GetProcObject
					if (nArg >= 1) {
						CHAR szProcNameA[MAX_LOADSTRING];
						LPSTR lpProcNameA = (LPSTR)&szProcNameA;
						VARIANT v;
						if (pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
							WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)pDispParams->rgvarg[nArg - 1].bstrVal, -1, lpProcNameA, MAX_LOADSTRING, NULL, NULL);
						}
						else {
							teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 1], VT_I4);
							lpProcNameA = MAKEINTRESOURCEA(v.lVal);
						}
						LPFNGetProcObjectW lpfnGetProcObjectW = (LPFNGetProcObjectW)GetProcAddress((HMODULE)GetLLFromVariant(&pDispParams->rgvarg[nArg]), lpProcNameA);
						if (lpfnGetProcObjectW) {
							lpfnGetProcObjectW(pVarResult);
						}
					}
					break;
				case 6001:
					*pbResult = SetCurrentDirectory(v.bstrVal);//use wsh.CurrentDirectory
					break;
				case 6002:
					*plResult = RegisterWindowMessage(v.bstrVal);
					break;
				case 6011:
					if (lpfnSetDllDirectoryW) {
						*pbResult = lpfnSetDllDirectoryW(v.bstrVal);
					}
					break;
				case 6012:
					if (nArg >= 1) {
						VARIANT v2;
						teVariantChangeType(&v2, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
						*plResult = lstrcmpi(v.bstrVal, v2.bstrVal);
						VariantClear(&v2);
					}
					break;
				case 6021:
					*pbResult = PathIsNetworkPath(v.bstrVal);
					break;
				case 6022:
					*plResult = lstrlen(v.bstrVal);
					break;
				case 6032:
					if (nArg >= 2) {
						VARIANT v2;
						teVariantChangeType(&v2, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
						*plResult = StrCmpNI(v.bstrVal, v2.bstrVal, GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]));
						VariantClear(&v2);
					}
					break;
				case 6042:
					if (nArg >= 1) {
						VARIANT v2;
						teVariantChangeType(&v2, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
						*plResult = StrCmpLogicalW(v.bstrVal, v2.bstrVal);
						VariantClear(&v2);
					}
					break;
				case 6005:
					if (v.bstrVal) {
						int nLen = SysStringLen(v.bstrVal);
						pszResult = new WCHAR[nLen + 3];
						::CopyMemory(pszResult, v.bstrVal, (nLen + 1) * sizeof(WCHAR));
						pszResult[nLen] = NULL;
						PathQuoteSpaces(pszResult);
					}
					break;
				case 6015:
					if (v.bstrVal) {
						int nLen = SysStringLen(v.bstrVal);
						pszResult = new WCHAR[nLen + 1];
						::CopyMemory(pszResult, v.bstrVal, (nLen + 1) * sizeof(WCHAR));
						pszResult[nLen] = NULL;
						PathUnquoteSpaces(pszResult);
					}
					break;
				case 6025:
					if (v.bstrVal) {
						int nLen = GetShortPathName(v.bstrVal, NULL, 0);
						pszResult = new WCHAR[nLen + 1];
						GetShortPathName(v.bstrVal, pszResult, nLen);
						if (nLen == 0) {
							nLen = SysStringLen(v.bstrVal);
							pszResult = new WCHAR[nLen + 1];
							::CopyMemory(pszResult, v.bstrVal, (nLen + 1) * sizeof(WCHAR));
						}
						pszResult[nLen] = NULL;
					}
					break;
				//{ 6035, L"PathCreateFromUrl"},
				case 6035:
					if (v.bstrVal) {
						pszResult = new WCHAR[MAX_PATH];
						DWORD dwLen = MAX_PATH;
						if FAILED(PathCreateFromUrl(v.bstrVal, pszResult, &dwLen, NULL)) {
							delete [] pszResult;
							pszResult = NULL;
						}
					}
					break;
				default:
					VariantClear(&v);
					return DISP_E_MEMBERNOTFOUND;
			}
			VariantClear(&v);
		}
		if ((dispIdMember % 10) == 5) {
			if (pszResult) {
				pVarResult->vt = VT_BSTR;
				pVarResult->bstrVal = SysAllocString(pszResult);
				delete [] pszResult;
			}
		}
		else {
			SetReturnValue(pVarResult, dispIdMember, &Result);
		}
		return S_OK;
	}
	switch(dispIdMember) {
		//Memory
		case 1040:	
			if (nArg >= 0) {
				if (pVarResult) {
					CteMemory *pMem;
					char *pc = NULL;;
					LPWSTR sz = NULL;
					int nCount = 1;
					int i = 0;
					if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
						sz = pDispParams->rgvarg[nArg].bstrVal;
						i = GetSizeOfStruct(sz);
					}
					if (pDispParams->rgvarg[nArg].vt & VT_ARRAY) {
						pc = (char *)pDispParams->rgvarg[nArg].parray->pvData;
						nCount = pDispParams->rgvarg[nArg].parray->rgsabound[0].cElements;
						i = SizeOfvt(pDispParams->rgvarg[nArg].vt);
					}
					if (i == 0) {
						i = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
						if (i == 0) {
							return S_OK;
						}
					}
					BSTR bs = NULL;
					if (nArg >= 1) {
						if (i == 2 && pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
							bs = pDispParams->rgvarg[nArg - 1].bstrVal;
							nCount = SysStringByteLen(bs) / 2 + 1;
						}
						else {
							nCount = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
						}
						if (nCount < 1) {
							nCount = 1;
						}
						if (nArg >= 2) {
							pc = (char *)GetLLFromVariant(&pDispParams->rgvarg[nArg - 2]);
						}
					}
					pMem = new CteMemory(i * nCount + GetOffsetOfStruct(sz), pc, GetVTOfStruct(sz), nCount, sz);
					if (bs) {
						::CopyMemory(pMem->m_pc, bs, nCount * 2);
					}
					teSetObject(pVarResult, pMem);
					pMem->Release();
				}
			}
			return S_OK;
		//ContextMenu
		case 1061:
			if (nArg >= 0) {
				if (pVarResult) {
					IDataObject *pDataObj;
					if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
						LPITEMIDLIST *ppidllist;
						long nCount;
						ppidllist = IDListFormDataObj(pDataObj, &nCount);
						AdjustIDList(ppidllist, nCount);
						if (nCount >= 1) {
							IShellFolder *pSF;
							if (GetShellFolder(&pSF, ppidllist[0])) {
								IContextMenu *pCM;
								if SUCCEEDED(pSF->GetUIObjectOf(g_hwndMain, nCount, (LPCITEMIDLIST *)&ppidllist[1], IID_IContextMenu, NULL, (LPVOID*)&pCM)) {
									CteContextMenu *pCCM;
									pCCM = new CteContextMenu(pCM, pDataObj);
									pDataObj = NULL;
									teSetObject(pVarResult, pCCM);
									pCCM->Release();
									pCM->Release();
								}
							}
							while (--nCount >= 0) {
								::CoTaskMemFree(ppidllist[nCount]);
							}
							delete [] ppidllist;
						}
						if (pDataObj) {
							pDataObj->Release();
						}
					}
				}
			}
			return S_OK;

		//DropTarget
		case 1070:
			if (nArg >= 0) {
				if (pVarResult) {
					LPITEMIDLIST pidl;
					if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
						IShellFolder *pSF;
						LPCITEMIDLIST ItemID;
						if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &ItemID)) {
							IDropTarget *pDT;
							if SUCCEEDED(pSF->GetUIObjectOf(g_hwndMain, 1, &ItemID, IID_IDropTarget, NULL, (LPVOID*)&pDT)) {
								CteDropTarget *pCDT = new CteDropTarget(pDT, pidl);
								pidl = NULL;
								teSetObject(pVarResult, pCDT);
								pCDT->Release();
								pDT->Release();
							}
							pSF->Release();
						}
						::CoTaskMemFree(pidl);
					}
				}
			}
			return S_OK;
		//OleGetClipboard
		case 1080:
			if (pVarResult) {
				IDataObject *pDataObj;
				if (OleGetClipboard(&pDataObj) == S_OK) {
					FolderItems *pFolderItems = new CteFolderItems(pDataObj, NULL, false);
					if (teSetObject(pVarResult, pFolderItems)) {
						STGMEDIUM Medium;
						if (pDataObj->GetData(&DROPEFFECTFormat, &Medium) == S_OK) {
							DWORD *pdwEffect = (DWORD *)GlobalLock(Medium.hGlobal);
							try {
								VARIANTARG *pv = GetNewVARIANT(1);
								pv[0].lVal = *pdwEffect;
								pv[0].vt = VT_I4;
								IDispatch *pdisp;
								pVarResult->punkVal->QueryInterface(IID_PPV_ARGS(&pdisp));
								Invoke5(pdisp, 0x10000001, DISPATCH_PROPERTYPUT, NULL, 1, pv);
								pdisp->Release();
							} catch (...) {
							}
							GlobalUnlock(Medium.hGlobal);
							ReleaseStgMedium(&Medium);
						}
					}
					pFolderItems->Release();
					pDataObj->Release();
				}
			}
			return S_OK;
		//OleSetClipboard
		case 1082:
			if (nArg >= 0) {
				*plResult = E_FAIL;
				IDataObject *pDataObj;
				if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
					*plResult = OleSetClipboard(pDataObj);
					pDataObj->Release();
				}
			}
			break;
		//sizeof
		case 1122:
			if (nArg >= 0) {
				*plResult = GetSizeOf(&pDispParams->rgvarg[nArg]);
			}
			break;
		//LowPart
		case 1132:
			if (nArg >= 0) {
				LARGE_INTEGER li;
				li.QuadPart = GetLLFromVariant(&pDispParams->rgvarg[nArg]);
				*plResult = li.LowPart;
			}
			break;
		//HighPart
		case 1142:
			if (nArg >= 0) {
				LARGE_INTEGER li;
				li.QuadPart = GetLLFromVariant(&pDispParams->rgvarg[nArg]);
				*plResult = li.HighPart;
			}
			break;
		//QuadPart
		case 1154:
			if (nArg >= 1) {
				LARGE_INTEGER li;
				li.LowPart = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				li.HighPart = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
				Result = li.QuadPart;
			}
			break;
		//pvData
		case 1163:
			if (nArg >= 0 && pDispParams->rgvarg[nArg].vt & VT_ARRAY) {
				*phResult = (HANDLE)pDispParams->rgvarg[nArg].parray->pvData;
			}
			break;
		//ExecMethod
		case 1170:
			if (nArg >= 2) {
				IDispatch *pdisp;
				if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
					VARIANT v;
					teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
					DISPID dispid;
					if (pdisp->GetIDsOfNames(IID_NULL, &v.bstrVal, 1, LOCALE_USER_DEFAULT, &dispid) != S_OK) {
						dispid = DISPID_VALUE;
					}
					VariantClear(&v);
					VARIANTARG *pv = &pDispParams->rgvarg[nArg - 2];
					int nCount = 1;
					CteMemory *pMem;
					if (GetMemFormVariant(&pDispParams->rgvarg[nArg - 2], &pMem)) {
						pv = (VARIANTARG *)pMem->m_pc;
						nCount = pMem->m_nCount;
						pMem->Release();
					}
					Invoke5(pdisp, dispid, DISPATCH_METHOD, pVarResult, -nCount, pv);
					pdisp->Release();
				}
			}
			break;
		//FindWindow
		case 2063:
			if (nArg >= 0) {
				VARIANT vClass, vWindow;
				teVariantChangeType(&vClass, &pDispParams->rgvarg[nArg], VT_BSTR);
				teVariantChangeType(&vWindow, &pDispParams->rgvarg[nArg], VT_BSTR);

				*phResult = FindWindow(vClass.bstrVal, vWindow.bstrVal);
			}
			break;
		//FindWindowEx
		case 2073:
			if (nArg >= 3) {
				VARIANT vClass, vWindow;
				teVariantChangeType(&vClass, &pDispParams->rgvarg[nArg - 2], VT_BSTR);
				teVariantChangeType(&vWindow, &pDispParams->rgvarg[nArg - 3], VT_BSTR);
				*phResult = FindWindowEx((HWND)GetLLFromVariant(&pDispParams->rgvarg[nArg]),
					(HWND)GetLLFromVariant(&pDispParams->rgvarg[nArg - 1]),
					vClass.bstrVal, vWindow.bstrVal);
				VariantClear(&vClass);
				VariantClear(&vWindow);
			}
			break;
		//PathMatchSpec
		case 2111:
			if (nArg >= 1) {
				VARIANT vFile, vSpec;
				teVariantChangeType(&vFile, &pDispParams->rgvarg[nArg], VT_BSTR);
				teVariantChangeType(&vSpec, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
				*pbResult = PathMatchSpec(vFile.bstrVal, vSpec.bstrVal);
				VariantClear(&vFile);
				VariantClear(&vSpec);
			}
			break;
		//CommandLineToArgv
		case 2230:
			if (pVarResult) {
				if (nArg >= 0) {
					int i = 0;
					VARIANT vCmdLine;
					teVariantChangeType(&vCmdLine, &pDispParams->rgvarg[nArg], VT_BSTR);
					LPTSTR *lplpszArgs = NULL;
					if (vCmdLine.bstrVal) {
						lplpszArgs = CommandLineToArgvW(vCmdLine.bstrVal, &i);
					}
					VariantClear(&vCmdLine);
					if (i > 0) {
						CteMemory *pMem;
						pMem = new CteMemory(i * sizeof(BSTR*), NULL, VT_BSTR, i, NULL);
						if (pMem) {
							BSTR* lpBSTRData;
							lpBSTRData = (BSTR*)pMem->m_pc;
							while (--i >= 0) {
								int n = lstrlen(lplpszArgs[i]);
								if (lplpszArgs[i][n - 1] == '"') {
									lplpszArgs[i][n - 1] = '\\';
								}
								lpBSTRData[i] = SysAllocString(lplpszArgs[i]);
							}
							teSetObject(pVarResult, pMem);
							pMem->Release();
						}
					}
					if (lplpszArgs) {
						LocalFree(lplpszArgs);
					}
				}
			}
			return S_OK;
		//IsEqual
		case 2431:			
		//ILIsParent
		case 2441:
			if (nArg >= 1) {
				LPITEMIDLIST pidl, pidl2;
				if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
					if (GetPidlFromVariant(&pidl2, &pDispParams->rgvarg[nArg - 1])) {
						if (dispIdMember == 2431) {
							*pbResult = ::ILIsEqual(pidl, pidl2);
						}
						else if (dispIdMember == 2441) {
							if (nArg >= 2) {
								*pbResult = ::ILIsParent(pidl, pidl2, GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]));
							}
						}
						::CoTaskMemFree(pidl2);
					}
					::CoTaskMemFree(pidl);
				}
			}
			break;
		//ILClone
		case 2440:
			if (pVarResult && nArg >= 0) {
				LPITEMIDLIST pidl;
				if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
					FolderItem *pFI;
					if (GetFolderItemFromPidl(&pFI, pidl)) {
						teSetObject(pVarResult, pFI);
						pFI->Release();
					}
					::CoTaskMemFree(pidl);
				}
			}
			break;
		//ILRemoveLastID
		case 2450:
			if (pVarResult && nArg >= 0) {
				LPITEMIDLIST pidl;
				if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
					if (!ILIsEmpty(pidl)) {
						if (ILRemoveLastID(pidl)) {
							FolderItem *pFI;
							if (GetFolderItemFromPidl(&pFI, pidl)) {
								teSetObject(pVarResult, pFI);
								pFI->Release();
							}
						}
					}
					::CoTaskMemFree(pidl);
				}
			}
			break;
		//ILFindLastID
		case 2460:
			if (pVarResult && nArg >= 0) {
				LPITEMIDLIST pidl, pidlLast;
				if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
					pidlLast = ILFindLastID(pidl);
					if (pidlLast) {
						FolderItem *pFI;
						if (GetFolderItemFromPidl(&pFI, pidlLast)) {
							teSetObject(pVarResult, pFI);
							pFI->Release();
						}
					}
					::CoTaskMemFree(pidl);
				}
			}
			break;
		//ILIsEmpty
		case 2461:
			if (nArg >= 0) {
				LPITEMIDLIST pidl;
				if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
					*pbResult = ILIsEmpty(pidl);
					::CoTaskMemFree(pidl);
				}
			}
			break;
		//FindFirstFile(str, mem): handle
		case 2763:
			if (nArg >= 1) {
				VARIANT vFile;
				teVariantChangeType(&vFile, &pDispParams->rgvarg[nArg], VT_BSTR);
				*phResult = FindFirstFile(vFile.bstrVal, (LPWIN32_FIND_DATA)GetpcFormVariant(&pDispParams->rgvarg[nArg - 1]));
			}
			break;
		case 2773:
			if (nArg >= 0) {
				POINT pt;
				GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg]);
				*phResult = WindowFromPoint(pt);
			}
			break;
/*		case 2782://Deprecated   sha.GetSystemInformation("DoubleClickTime")
			*plResult = GetDoubleClickTime(); 
			break;*/
		case 2792:
			*plResult = g_nThreads;
			break;
		case 2801:	//DoEvents
			MSG msg;
			*pbResult = PeekMessage(&msg, NULL, 0, 0, PM_REMOVE);
			if (*pbResult) {
				if (MessageProc(&msg) != S_OK) {
					TranslateMessage(&msg);
					DispatchMessage(&msg);
				}
			}
			break;
		case 2804:	//swscanf
			if (nArg >= 1 && pVarResult) {
				LPWSTR pszData = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]);
				LPWSTR pszFormat = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]);
				if (pszData && pszFormat) {
					int nPos = 0;
					while (pszFormat[nPos] && pszFormat[nPos++] != '%') {
					}
					while (WCHAR wc = pszFormat[nPos]) {
						nPos++;
						if (wc == '%') {
							break;
						}
						if (StrChrIW(L"diuoxc", wc)) {//Integer
							swscanf_s(pszData, pszFormat, &Result);
							break;
						}
						if (StrChrIW(L"efga", wc)) {//Real
							pVarResult->vt = VT_R8;
							swscanf_s(pszData, pszFormat, &pVarResult->dblVal);
							return S_OK;
						}
						if (StrChrIW(L"s", wc)) {//String
							int nLen = lstrlen(pszData);
							LPWSTR sz = new WCHAR[nLen];
							swscanf_s(pszData, pszFormat, sz, nLen);
							pVarResult->bstrVal = ::SysAllocString(sz);
							pVarResult->vt = VT_BSTR;
							delete [] sz;
							return S_OK;
						}
					}
				}
			}
			break;
		case 2811:	//SetFileTime
			*pbResult = 0;
			if (nArg >= 3) {
				LPITEMIDLIST pidl;
				if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
					BSTR bs;
					if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORPARSING)) {
						HANDLE hFile = ::CreateFile(bs, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0);
						if (hFile != INVALID_HANDLE_VALUE) {
							FILETIME **ppft = new LPFILETIME[3];
							for (int i = 2; i >= 0; i--) {
								ppft[i] = NULL;
								VARIANT v;
								teVariantChangeType(&v, &pDispParams->rgvarg[nArg - i - 1], VT_DATE);
								if (v.vt == VT_DATE && v.date >= 0) {
									ppft[i] = new FILETIME[1];
									teVariantTimeToFileTime(v.date, ppft[i]);
								}
								VariantClear(&v);
							}
							*pbResult = ::SetFileTime(hFile, ppft[0], ppft[1], ppft[2]);
							::CloseHandle(hFile);
						}
						::SysFreeString(bs);
					}
				}
			}
			break;
		//Value
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
		default:
			return DISP_E_MEMBERNOTFOUND;
	}
	SetReturnValue(pVarResult, dispIdMember, &Result);
	return S_OK;
}

//CteDispatch

CteDispatch::CteDispatch(IDispatch *pDispatch, int nMode)
{
	m_cRef = 1;
	m_pDispatch = pDispatch;
	m_pDispatch->AddRef();
	m_nMode = nMode;
	m_dispIdMember = DISPID_UNKNOWN;
	m_pActiveScript = NULL;
	m_nIndex = 0;
}

CteDispatch::~CteDispatch()
{
	if (m_pDispatch) {
		m_pDispatch->Release();
	}
	if (m_pActiveScript) {
		m_pActiveScript->SetScriptState(SCRIPTSTATE_CLOSED);
		m_pActiveScript->Close();
		m_pActiveScript->Release();
	}
}

STDMETHODIMP CteDispatch::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else if (IsEqualIID(riid, IID_IEnumVARIANT)) {
		*ppvObject = static_cast<IEnumVARIANT *>(this);
	}
	else if (m_nMode) {
		return m_pDispatch->QueryInterface(riid, ppvObject);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
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
	if (m_nMode) {
		if (m_nMode == 2) {
			if (lstrcmp(*rgszNames, L"Item") == 0) {
				*rgDispId = DISPID_VALUE;
				return S_OK;
			}
		}
		return m_pDispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
	}
	return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP CteDispatch::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	if (m_nMode) {
		if (dispIdMember == DISPID_NEWENUM) {
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
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
				}
				else {
					delete [] pv;
				}
				if (pVarResult) {
					Invoke5(m_pDispatch, dispIdMember, DISPATCH_PROPERTYGET, pVarResult, 0, NULL);
				}
				return S_OK;
			}
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
		}
		return m_pDispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	}
	if (wFlags & DISPATCH_METHOD) {
		return m_pDispatch->Invoke(m_dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	}
	if (pVarResult) {
		teSetObject(pVarResult, this);
	}
	return S_OK;
}

STDMETHODIMP CteDispatch::Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched)
{
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
			pv[0].vt = VT_I4;
			pv[0].lVal = m_nIndex++;
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
	if (ppEnum) {
		CteDispatch *pdisp = new CteDispatch(m_pDispatch, m_nMode);
		pdisp->m_dispIdMember = m_dispIdMember;
		pdisp->m_nMode = m_nMode;
		*ppEnum = static_cast<IEnumVARIANT *>(pdisp);
		return S_OK;
	}
	return E_POINTER;
}

//CteActiveScriptSite
CteActiveScriptSite::CteActiveScriptSite(IUnknown *punk, IUnknown *pOnError)
{
	m_cRef = 1;
	m_pDispatchEx = NULL;
	m_pOnError = NULL;
	if (punk) {
		punk->QueryInterface(IID_PPV_ARGS(&m_pDispatchEx));
	}
	if (pOnError) {
		pOnError->QueryInterface(IID_PPV_ARGS(&m_pOnError));
	}
}

CteActiveScriptSite::~CteActiveScriptSite()
{
	m_pOnError && m_pOnError->Release();
	m_pDispatchEx && m_pDispatchEx->Release();
}

STDMETHODIMP CteActiveScriptSite::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IActiveScriptSite)) {
		*ppvObject = static_cast<IActiveScriptSite *>(this);
	}
	else if (IsEqualIID(riid, IID_IActiveScriptSiteWindow)) {
		*ppvObject = static_cast<IActiveScriptSiteWindow *>(this);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
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
			BSTR bs = ::SysAllocString(pstrName);
			DISPID dispid;
			if (m_pDispatchEx->GetDispID(bs, 0, &dispid) == S_OK) {
				DISPPARAMS noargs = { NULL, NULL, 0, 0 };
				VARIANT v;
				VariantInit(&v);
				if (m_pDispatchEx->InvokeEx(dispid, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noargs, &v, NULL, NULL) == S_OK) {
					if (FindUnknown(&v, ppiunkItem)) {
						hr = S_OK;
					}
					else {
						VariantClear(&v);
					}
				}
			}
			::SysFreeString(bs);
		}
		else if (g_pWebBrowser && lstrcmpi(pstrName, L"window") == 0) {
			IDispatch *pdisp;
			if (g_pWebBrowser->m_pWebBrowser->get_Document(&pdisp) == S_OK) {
				IHTMLDocument2 *pdoc;
				if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pdoc))) {
					IHTMLWindow2 *pwin;
					if SUCCEEDED(pdoc->get_parentWindow(&pwin)) {
						pwin->QueryInterface(IID_PPV_ARGS(ppiunkItem));
						pwin->Release();
					}
					pdoc->Release();
				}
				pdisp->Release();
			}
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
	if (!pscripterror) {
		return E_POINTER;
	}
	HRESULT hr = E_NOTIMPL;
	if (m_pOnError) {
		EXCEPINFO ei;
		hr = pscripterror->GetExceptionInfo(&ei);
		CteMemory *pei = new CteMemory(sizeof(EXCEPINFO), NULL, VT_NULL, 1, L"EXCEPINFO");
		::CopyMemory(pei->m_pc, &ei, sizeof(EXCEPINFO));
		VARIANTARG *pv = GetNewVARIANT(5);
		teSetObject(&pv[4], pei);
		pei->Release();
		pv[3].vt = VT_BSTR;
		if (pscripterror->GetSourceLineText(&pv[3].bstrVal) != S_OK) {
			pv[3].bstrVal = NULL;
		}
		if (pscripterror->GetSourcePosition(&pv[2].ulVal, &pv[1].ulVal, &pv[0].lVal) == S_OK) {
			pv[2].vt = VT_I4;
			pv[1].vt = VT_I4;
			pv[0].vt = VT_I4;
		}
		Invoke4(m_pOnError, NULL, 5, pv);
	}
	return hr;
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
	*phwnd = g_pWebBrowser->get_HWND();
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::EnableModeless(BOOL fEnable)
{
	return E_NOTIMPL;
}


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
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
}

STDMETHODIMP_(ULONG) CteTest::AddRef()
{
	g_bReload = false;
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteTest::Release()
{
	g_bReload = false;
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
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}
#endif