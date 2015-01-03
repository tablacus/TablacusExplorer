// TE.cpp
// Tablacus Explorer (C)2011- Gaku
// MIT Lisence
// Visual C++ 2008 Express Edition SP1
// Windows SDK v7.0
// http://www.eonet.ne.jp/~gakana/tablacus/

#include "stdafx.h"
#include <windows.h>
#include "TE.h"

// Global Variables:
HINSTANCE hInst;								// current instance
WCHAR	g_szTE[MAX_LOADSTRING];
WCHAR	g_szText[MAX_STATUS];
WCHAR	g_szStatus[MAX_STATUS];
HWND	g_hwndMain = NULL;
HWND	g_hwndBrowser = NULL;
CteTabs *g_pTabs = NULL;
CteTabs *g_pTC[MAX_TC];
HINSTANCE	g_hShell32 = NULL;
HINSTANCE	g_hCrypt32 = NULL;
HWND		g_hDialog = NULL;
IShellWindows *g_pSW = NULL;

LPFNSHCreateItemFromIDList lpfnSHCreateItemFromIDList = NULL;
LPFNSetDllDirectoryW lpfnSetDllDirectoryW = NULL;
LPFNSHParseDisplayName lpfnSHParseDisplayName = NULL;
LPFNSHGetImageList lpfnSHGetImageList = NULL;
LPFNSetWindowTheme lpfnSetWindowTheme = NULL;
LPFNSHRunDialog lpfnSHRunDialog = NULL;
LPFNRegenerateUserEnvironment lpfnRegenerateUserEnvironment = NULL;
LPFNCryptBinaryToStringW lpfnCryptBinaryToStringW = NULL;
LPFNPSPropertyKeyFromString lpfnPSPropertyKeyFromString = NULL;
LPFNPSGetPropertyKeyFromName lpfnPSGetPropertyKeyFromName = NULL;
LPFNPSPropertyKeyFromString lpfnPSPropertyKeyFromStringEx = NULL;
LPFNPSGetPropertyDescription lpfnPSGetPropertyDescription = NULL;
LPFNSHGetIDListFromObject lpfnSHGetIDListFromObject = NULL;
LPFNIsWow64Process lpfnIsWow64Process = NULL;
LPFNChangeWindowMessageFilter lpfnChangeWindowMessageFilter = NULL;
LPFNChangeWindowMessageFilterEx lpfnChangeWindowMessageFilterEx = NULL;
LPFNRtlGetVersion lpfnRtlGetVersion = NULL;

LPITEMIDLIST g_pidlCP = NULL;
FORMATETC HDROPFormat = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC IDLISTFormat = {0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC DROPEFFECTFormat = {0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC UNICODEFormat = {CF_UNICODETEXT, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
FORMATETC TEXTFormat = {CF_TEXT, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
GUID		g_ClsIdStruct;
GUID		g_ClsIdSB;
GUID		g_ClsIdTC;
GUID		g_ClsIdFI;
GUID		g_ClsIdFIs;
IUIAutomation *g_pAutomation;
PROPERTYID	g_PID_ItemIndex;

CTE			*g_pTE;
CteShellBrowser *g_pSB[MAX_FV];
CteWebBrowser *g_pWebBrowser;

LPITEMIDLIST g_pidls[MAX_CSIDL];
LPITEMIDLIST g_pidlResultsFolder;
LPITEMIDLIST g_pidlLibrary = NULL;

IDispatch	*g_pOnFunc[Count_OnFunc];
IDispatch	*g_pUnload  = NULL;
IDispatch	*g_pJS		= NULL;
IDropTargetHelper *g_pDropTargetHelper = NULL;

IContextMenu *g_pCM = NULL;
ULONG_PTR g_Token;
Gdiplus::GdiplusStartupInput g_StartupInput;
HHOOK	g_hHook;
HHOOK	g_hMouseHook;
HHOOK	g_hMessageHook;
HHOOK	g_hMenuKeyHook = NULL;
HMENU	g_hMenu = NULL;
BSTR	g_bsCmdLine = NULL;
HANDLE g_hMutex = NULL;

UINT	g_uCrcTable[256];
LONG	g_nSize = MAXWORD;
LONG	g_nLockUpdate = 0;
LONG	nUpdate = 0;
DWORD	g_paramFV[SB_Count];
DWORD	g_dragdrop = 0;
DWORD	g_dwMainThreadId;
DWORD	g_dwTickScroll;
DWORD	g_dwSWCookie = 0;
long	g_nProcKey	   = 0;
long	g_nProcMouse   = 0;
long	g_nProcDrag    = 0;
long	g_nProcFV      = 0;
long	g_nProcTV      = 0;
long	g_nProcFI      = 0;
long	g_nThreads	   = 0;
long	g_lSWCookie = 0;
int		g_nCmdShow = SW_SHOWNORMAL;
int		g_nReload = 0;
int		g_nException = 64;
int		*g_maps[MAP_LENGTH];
int		g_x = MAXINT;
int		g_nPidls = MAX_CSIDL;
int		g_nTCCount = 0;
int		g_nTCIndex = 0;
BOOL	g_bDialogOk = FALSE;
BOOL	g_bInit = true;
BOOL	g_bUnload = FALSE;
BOOL	g_bSetRedraw;
#ifdef _2000XP
int		g_nCharWidth = 7;
BOOL	g_bCharWidth = true;
BOOL	g_bUpperVista;
BOOL	g_bIsXP;
BOOL	g_bIs2000;
#endif

TEmethod tesNULL[] =
{
	{ 0, NULL }
};

TEmethod tesBITMAP[] =
{
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmType), L"bmType" },
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmWidth), L"bmWidth" },
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmHeight), L"bmHeight" },
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmWidthBytes), L"bmWidthBytes" },
	{ (VT_UI2 << TE_VT) + offsetof(BITMAP, bmPlanes), L"bmPlanes" },
	{ (VT_UI2 << TE_VT) + offsetof(BITMAP, bmBitsPixel), L"bmBitsPixel" },
	{ (VT_PTR << TE_VT) + offsetof(BITMAP, bmBits), L"bmBits" },
	{ 0, NULL }
};

TEmethod tesCHOOSECOLOR[] =
{
	{ (VT_I4 << TE_VT) + offsetof(CHOOSECOLOR, lStructSize), L"lStructSize" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, hwndOwner), L"hwndOwner" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, hInstance), L"hInstance" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSECOLOR, rgbResult), L"rgbResult" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, lpCustColors), L"lpCustColors" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSECOLOR, Flags), L"Flags" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, lCustData), L"lCustData" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, lpfnHook), L"lpfnHook" },
	{ (VT_BSTR << TE_VT) + offsetof(CHOOSECOLOR, lpTemplateName), L"lpTemplateName" },
	{ 0, NULL }
};

TEmethod tesCHOOSEFONT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, lStructSize), L"lStructSize" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, hwndOwner), L"hwndOwner" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, hDC), L"hDC" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, lpLogFont), L"lpLogFont" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, iPointSize), L"iPointSize" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, Flags), L"Flags" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, rgbColors), L"rgbColors" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, lCustData), L"lCustData" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, lpfnHook), L"lpfnHook" },
	{ (VT_BSTR << TE_VT) + offsetof(CHOOSEFONT, lpTemplateName), L"lpTemplateName" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, hInstance), L"hInstance" },
	{ (VT_BSTR << TE_VT) + offsetof(CHOOSEFONT, lpszStyle), L"lpszStyle" },
	{ (VT_UI2 << TE_VT) + offsetof(CHOOSEFONT, nFontType), L"nFontType" },
	{ (VT_UI2 << TE_VT) + offsetof(CHOOSEFONT, ___MISSING_ALIGNMENT__), L"___MISSING_ALIGNMENT__" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, nSizeMin), L"nSizeMin" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, nSizeMax), L"nSizeMax" },
	{ 0, NULL }
};

TEmethod tesCOPYDATASTRUCT[] =
{
	{ (VT_PTR << TE_VT) + offsetof(COPYDATASTRUCT, dwData), L"dwData" },
	{ (VT_I4 << TE_VT) + offsetof(COPYDATASTRUCT, cbData), L"cbData" },
	{ (VT_PTR << TE_VT) + offsetof(COPYDATASTRUCT, lpData), L"lpData" },
	{ 0, NULL }
};

TEmethod tesDIBSECTION[] =
{
	{ (VT_PTR << TE_VT) + offsetof(DIBSECTION, dsBm), L"dsBm" },
	{ (VT_PTR << TE_VT) + offsetof(DIBSECTION, dsBmih), L"dsBmih" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsBitfields[0]), L"dsBitfields0" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsBitfields[1]), L"dsBitfields1" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsBitfields[2]), L"dsBitfields2" },
	{ (VT_PTR << TE_VT) + offsetof(DIBSECTION, dshSection), L"dshSection" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsOffset), L"dsOffset" },
	{ 0, NULL }
};

TEmethod tesEncoderParameter[] =
{
	{ (VT_CLSID << TE_VT) + offsetof(EncoderParameter, Guid), L"Guid" },
	{ (VT_I4 << TE_VT) + offsetof(EncoderParameter, NumberOfValues), L"NumberOfValues" },
	{ (VT_I4 << TE_VT) + offsetof(EncoderParameter, Type), L"Type" },
	{ (VT_PTR << TE_VT) + offsetof(EncoderParameter, Value), L"Value" },
	{ 0, NULL }
};

TEmethod tesEncoderParameters[] =
{
	{ (VT_I4 << TE_VT) + offsetof(EncoderParameters, Count), L"Count" },
	{ DISPID_TE_ITEM, L"Parameter" },
	{ 0, NULL }
};

TEmethod tesEXCEPINFO[] =
{
	{ (VT_UI2 << TE_VT) + offsetof(EXCEPINFO, wCode), L"wCode" },
	{ (VT_UI2 << TE_VT) + offsetof(EXCEPINFO, wReserved), L"wReserved" },
	{ (VT_BSTR << TE_VT) + offsetof(EXCEPINFO, bstrSource), L"bstrSource" },
	{ (VT_BSTR << TE_VT) + offsetof(EXCEPINFO, bstrDescription), L"bstrDescription" },
	{ (VT_BSTR << TE_VT) + offsetof(EXCEPINFO, bstrHelpFile), L"bstrHelpFile" },
	{ (VT_I4 << TE_VT) + offsetof(EXCEPINFO, dwHelpContext), L"dwHelpContext" },
	{ (VT_PTR << TE_VT) + offsetof(EXCEPINFO, pvReserved), L"pvReserved" },
	{ (VT_PTR << TE_VT) + offsetof(EXCEPINFO, pfnDeferredFillIn), L"pfnDeferredFillIn" },
	{ (VT_I4 << TE_VT) + offsetof(EXCEPINFO, scode), L"scode" },
	{ 0, NULL }
};

TEmethod tesFINDREPLACE[] =
{
	{ (VT_I4 << TE_VT) + offsetof(FINDREPLACE, lStructSize), L"lStructSize" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, hwndOwner), L"hwndOwner" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, hInstance), L"hInstance" },
	{ (VT_I4 << TE_VT) + offsetof(FINDREPLACE, Flags), L"Flags" },
	{ (VT_BSTR << TE_VT) + offsetof(FINDREPLACE, lpstrFindWhat), L"lpstrFindWhat" },
	{ (VT_BSTR << TE_VT) + offsetof(FINDREPLACE, lpstrReplaceWith), L"lpstrReplaceWith" },
	{ (VT_UI2 << TE_VT) + offsetof(FINDREPLACE, wFindWhatLen), L"wFindWhatLen" },
	{ (VT_UI2 << TE_VT) + offsetof(FINDREPLACE, wReplaceWithLen), L"wReplaceWithLen" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, lCustData), L"lCustData" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, lpfnHook), L"lpfnHook" },
	{ (VT_BSTR << TE_VT) + offsetof(FINDREPLACE, lpTemplateName), L"lpTemplateName" },
	{ 0, NULL }
};

TEmethod tesFOLDERSETTINGS[] =
{
	{ (VT_I4 << TE_VT) + offsetof(FOLDERSETTINGS, ViewMode), L"ViewMode" },
	{ (VT_I4 << TE_VT) + offsetof(FOLDERSETTINGS, fFlags), L"fFlags" },
	{ (VT_I4 << TE_VT) + (SB_Options - 1) * 4, L"Options" },
	{ (VT_I4 << TE_VT) + (SB_ViewFlags - 1) * 4, L"ViewFlags" },
	{ (VT_I4 << TE_VT) + (SB_IconSize - 1) * 4, L"ImageSize" },
	{ 0, NULL }
};

TEmethod tesHDITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, mask), L"mask" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, cxy), L"cxy" },
	{ (VT_BSTR << TE_VT) + offsetof(HDITEM, pszText), L"pszText" },
	{ (VT_PTR << TE_VT) + offsetof(HDITEM, hbm), L"hbm" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, cchTextMax), L"cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, fmt), L"fmt" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, lParam), L"lParam" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, iImage), L"iImage" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, iOrder), L"iOrder" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, type), L"type" },
	{ (VT_PTR << TE_VT) + offsetof(HDITEM, pvFilter), L"pvFilter" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, state), L"state" },
	{ 0, NULL }
};

TEmethod tesICONINFO[] =
{
	{ (VT_BOOL << TE_VT) + offsetof(ICONINFO, fIcon), L"fIcon" },
	{ (VT_I4 << TE_VT) + offsetof(ICONINFO, xHotspot), L"xHotspot" },
	{ (VT_I4 << TE_VT) + offsetof(ICONINFO, yHotspot), L"yHotspot" },
	{ (VT_PTR << TE_VT) + offsetof(ICONINFO, hbmMask), L"hbmMask" },
	{ (VT_PTR << TE_VT) + offsetof(ICONINFO, hbmColor), L"hbmColor" },
	{ 0, NULL }
};

TEmethod tesICONMETRICS[] =
{
	{ (VT_I4 << TE_VT) + offsetof(ICONMETRICS, iHorzSpacing), L"iHorzSpacing" },
	{ (VT_I4 << TE_VT) + offsetof(ICONMETRICS, iVertSpacing), L"iVertSpacing" },
	{ (VT_I4 << TE_VT) + offsetof(ICONMETRICS, iTitleWrap), L"iTitleWrap" },
	{ (VT_PTR << TE_VT) + offsetof(ICONMETRICS, lfFont), L"lfFont" },
	{ 0, NULL }
};

TEmethod tesKEYBDINPUT[] =
{
	{ (VT_I4 << TE_VT), L"type" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, wVk), L"wVk" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, wScan), L"wScan" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, dwFlags), L"dwFlags" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, time), L"time" },
	{ (VT_PTR << TE_VT) + offsetof(KEYBDINPUT, dwExtraInfo), L"dwExtraInfo" },
	{ 0, NULL }
};

TEmethod tesLOGFONT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfHeight), L"lfHeight" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfWidth), L"lfWidth" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfEscapement), L"lfEscapement" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfOrientation), L"lfOrientation" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfWeight), L"lfWeight" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfItalic), L"lfItalic" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfUnderline), L"lfUnderline" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfStrikeOut), L"lfStrikeOut" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfCharSet), L"lfCharSet" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfOutPrecision), L"lfOutPrecision" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfClipPrecision), L"lfClipPrecision" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfQuality), L"lfQuality" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfPitchAndFamily), L"lfPitchAndFamily" },
	{ (VT_LPWSTR << TE_VT) + offsetof(LOGFONT, lfFaceName), L"lfFaceName" },
	{ 0, NULL }
};

TEmethod tesLVBKIMAGE[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, ulFlags), L"ulFlags" },
	{ (VT_PTR << TE_VT) + offsetof(LVBKIMAGE, hbm), L"hbm" },
	{ (VT_BSTR << TE_VT) + offsetof(LVBKIMAGE, pszImage), L"pszImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, cchImageMax), L"cchImageMax" },
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, xOffsetPercent), L"xOffsetPercent" },
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, yOffsetPercent), L"yOffsetPercent" },
	{ 0, NULL }
};

TEmethod tesLVFINDINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVFINDINFO, flags), L"flags" },
	{ (VT_BSTR << TE_VT) + offsetof(LVFINDINFO, psz), L"psz" },
	{ (VT_I4 << TE_VT) + offsetof(LVFINDINFO, lParam), L"lParam" },
	{ (VT_CY << TE_VT) + offsetof(LVFINDINFO, pt), L"pt" },
	{ (VT_I4 << TE_VT) + offsetof(LVFINDINFO, vkDirection), L"vkDirection" },
	{ 0, NULL }
};

TEmethod tesLVGROUP[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cbSize), L"cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, mask), L"mask" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszHeader), L"pszHeader" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchHeader), L"cchHeader" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszFooter), L"pszFooter" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchFooter), L"cchFooter" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iGroupId), L"iGroupId" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, stateMask), L"stateMask" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, state), L"state" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, uAlign), L"uAlign" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszSubtitle), L"pszSubtitle" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchSubtitle), L"cchSubtitle" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszTask), L"pszTask" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchTask), L"cchTask" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszDescriptionTop), L"pszDescriptionTop" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchDescriptionTop), L"cchDescriptionTop" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszDescriptionBottom), L"pszDescriptionBottom" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchDescriptionBottom), L"cchDescriptionBottom" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iTitleImage), L"iTitleImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iExtendedImage), L"iExtendedImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iFirstItem), L"iFirstItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cItems), L"cItems" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszSubsetTitle), L"pszSubsetTitle" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchSubsetTitle), L"cchSubsetTitle" },
	{ 0, NULL }
};

TEmethod tesLVHITTESTINFO[] =
{
	{ (VT_CY << TE_VT) + offsetof(LVHITTESTINFO, pt), L"pt" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, flags), L"flags" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, iItem), L"iItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, iSubItem ), L"iSubItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, iGroup), L"iGroup" },
	{ 0, NULL }
};

TEmethod tesLVITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, mask), L"mask" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iItem), L"iItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iSubItem), L"iSubItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, state), L"state" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, stateMask), L"stateMask" },
	{ (VT_BSTR << TE_VT) + offsetof(LVITEM, pszText), L"pszText" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, cchTextMax), L"cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iImage), L"iImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, lParam), L"lParam" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iIndent), L"iIndent" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iGroupId), L"iGroupId" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, cColumns), L"cColumns" },
	{ (VT_PTR << TE_VT) + offsetof(LVITEM, puColumns), L"puColumns" },
	{ (VT_PTR << TE_VT) + offsetof(LVITEM, piColFmt), L"piColFmt" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iGroup), L"iGroup" },
	{ 0, NULL }
};

TEmethod tesMENUITEMINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, cbSize), L"cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, fMask), L"fMask" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, fType), L"fType" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, fState), L"fState" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, wID), L"wID" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hSubMenu), L"hSubMenu" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hbmpChecked), L"hbmpChecked" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hbmpUnchecked), L"hbmpUnchecked" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, dwItemData), L"dwItemData" },
	{ (VT_BSTR << TE_VT) + offsetof(MENUITEMINFO, dwTypeData), L"dwTypeData" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, cch), L"cch" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hbmpItem), L"hbmpItem" },
	{ 0, NULL }
};

TEmethod tesMOUSEINPUT[] =
{
	{ (VT_I4 << TE_VT), L"type" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, dx), L"dx" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, dy), L"dy" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, mouseData), L"mouseData" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, dwFlags), L"dwFlags" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, time), L"time" },
	{ (VT_PTR << TE_VT) + offsetof(MOUSEINPUT, dwExtraInfo), L"dwExtraInfo" },
	{ 0, NULL }
};

TEmethod tesMSG[] =
{
	{ (VT_PTR << TE_VT) + offsetof(MSG, hwnd), L"hwnd" },
	{ (VT_I4 << TE_VT) + offsetof(MSG, message), L"message" },
	{ (VT_PTR << TE_VT) + offsetof(MSG, wParam), L"wParam" },
	{ (VT_PTR << TE_VT) + offsetof(MSG, lParam), L"lParam" },
	{ (VT_I4 << TE_VT) + offsetof(MSG, time), L"time" },
	{ (VT_CY << TE_VT) + offsetof(MSG, pt), L"pt" },
	{ 0, NULL }
};

TEmethod tesNONCLIENTMETRICS[] =
{
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, cbSize), L"cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iBorderWidth), L"iBorderWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iScrollWidth), L"iScrollWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iScrollHeight), L"iScrollHeight" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iCaptionWidth), L"iCaptionWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iCaptionHeight), L"iCaptionHeight" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfCaptionFont), L"lfCaptionFont" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iSmCaptionWidth), L"iSmCaptionWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iSmCaptionHeight), L"iSmCaptionHeight" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfSmCaptionFont), L"lfSmCaptionFont" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iMenuWidth), L"iMenuWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iMenuHeight), L"iMenuHeight" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfMenuFont), L"lfMenuFont" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfStatusFont), L"lfStatusFont" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfMessageFont), L"lfMessageFont" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iPaddedBorderWidth), L"iPaddedBorderWidth" },
	{ 0, NULL }
};

TEmethod tesNOTIFYICONDATA[] =
{
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, cbSize), L"cbSize" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, hWnd), L"hWnd" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uID), L"uID" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uFlags), L"uFlags" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uCallbackMessage), L"uCallbackMessage" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, hIcon), L"hIcon" },
	{ (VT_LPWSTR << TE_VT) + offsetof(NOTIFYICONDATA, szTip), L"szTip" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwState), L"dwState" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwState), L"dwState" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwStateMask), L"dwStateMask" },
	{ (VT_LPWSTR << TE_VT) + offsetof(NOTIFYICONDATA, szInfo), L"szInfo" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uTimeout), L"uTimeout" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uVersion), L"uVersion" },
	{ (VT_LPWSTR << TE_VT) + offsetof(NOTIFYICONDATA, szInfoTitle), L"szInfoTitle" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwInfoFlags), L"dwInfoFlags" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, guidItem), L"guidItem" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, hBalloonIcon), L"hBalloonIcon" },
	{ 0, NULL }
};

TEmethod tesNMHDR[] =
{
	{ (VT_PTR << TE_VT) + offsetof(NMHDR, hwndFrom), L"hwndFrom" },
	{ (VT_I4 << TE_VT) + offsetof(NMHDR, idFrom), L"idFrom" },
	{ (VT_I4 << TE_VT) + offsetof(NMHDR, code), L"code" },
	{ 0, NULL }
};

TEmethod tesOSVERSIONINFOEX[] =
{
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwOSVersionInfoSize), L"dwOSVersionInfoSize" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwMajorVersion), L"dwMajorVersion" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwMinorVersion), L"dwMinorVersion" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwBuildNumber), L"dwBuildNumber" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwPlatformId), L"dwPlatformId" },
	{ (VT_LPWSTR << TE_VT) + offsetof(OSVERSIONINFOEX, szCSDVersion), L"szCSDVersion" },
	{ (VT_UI2 << TE_VT) + offsetof(OSVERSIONINFOEX, wServicePackMajor), L"wServicePackMajor" },
	{ (VT_UI2 << TE_VT) + offsetof(OSVERSIONINFOEX, wServicePackMinor), L"wServicePackMinor" },
	{ (VT_UI2 << TE_VT) + offsetof(OSVERSIONINFOEX, wSuiteMask), L"wSuiteMask" },
	{ (VT_UI1 << TE_VT) + offsetof(OSVERSIONINFOEX, wProductType), L"wProductType" },
	{ (VT_UI1 << TE_VT) + offsetof(OSVERSIONINFOEX, wReserved), L"wReserved" },
	{ 0, NULL }
};

TEmethod tesPAINTSTRUCT[] =
{
	{ (VT_PTR << TE_VT) + offsetof(PAINTSTRUCT, hdc), L"hdc" },
	{ (VT_BOOL << TE_VT) + offsetof(PAINTSTRUCT, fErase), L"fErase" },
	{ (VT_CARRAY << TE_VT) + offsetof(PAINTSTRUCT, rcPaint), L"rcPaint" },
	{ (VT_BOOL << TE_VT) + offsetof(PAINTSTRUCT, fRestore), L"fRestore" },
	{ (VT_BOOL << TE_VT) + offsetof(PAINTSTRUCT, fIncUpdate), L"fIncUpdate" },
	{ (VT_UI1 << TE_VT) + offsetof(PAINTSTRUCT, rgbReserved), L"rgbReserved" },
	{ 0, NULL }
};

TEmethod tesPOINT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(POINT, x), L"x" },
	{ (VT_I4 << TE_VT) + offsetof(POINT, y), L"y" },
	{ 0, NULL }
};

TEmethod tesRECT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(RECT, left), L"left" },
	{ (VT_I4 << TE_VT) + offsetof(RECT, top), L"top" },
	{ (VT_I4 << TE_VT) + offsetof(RECT, right), L"right" },
	{ (VT_I4 << TE_VT) + offsetof(RECT, bottom), L"bottom" },
	{ 0, NULL }
};

TEmethod tesSHELLEXECUTEINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, cbSize), L"cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, fMask), L"fMask" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hwnd), L"hwnd" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpVerb), L"lpVerb" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpFile), L"lpFile" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpParameters), L"lpParameters" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpDirectory), L"lpDirectory" },
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, nShow), L"nShow" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hInstApp), L"hInstApp" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpIDList), L"lpIDList" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpClass), L"lpClass" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hkeyClass), L"hkeyClass" },
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, dwHotKey), L"dwHotKey" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hIcon), L"hIcon" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hMonitor), L"hMonitor" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hProcess), L"hProcess" },
	{ 0, NULL }
};

TEmethod tesSHFILEINFO[] =
{
	{ (VT_PTR << TE_VT) + offsetof(SHFILEINFO, hIcon), L"hIcon" },
	{ (VT_I4 << TE_VT) + offsetof(SHFILEINFO, iIcon), L"iIcon" },
	{ (VT_I4 << TE_VT) + offsetof(SHFILEINFO, dwAttributes), L"dwAttributes" },
	{ (VT_LPWSTR << TE_VT) + offsetof(SHFILEINFO, szDisplayName), L"szDisplayName" },
	{ (VT_LPWSTR << TE_VT) + offsetof(SHFILEINFO, szTypeName), L"szTypeName" },
	{ 0, NULL }
};

TEmethod tesSHFILEOPSTRUCT[] =
{
	{ (VT_PTR << TE_VT) + offsetof(SHFILEOPSTRUCT, hwnd), L"hwnd" },
	{ (VT_I4 << TE_VT) + offsetof(SHFILEOPSTRUCT, wFunc), L"wFunc" },
	{ (VT_BSTR << TE_VT) + offsetof(SHFILEOPSTRUCT, pFrom), L"pFrom" },
	{ (VT_BSTR << TE_VT) + offsetof(SHFILEOPSTRUCT, pTo), L"pTo" },
	{ (VT_UI2 << TE_VT) + offsetof(SHFILEOPSTRUCT, fFlags), L"fFlags" },
	{ (VT_BOOL << TE_VT) + offsetof(SHFILEOPSTRUCT, fAnyOperationsAborted), L"fAnyOperationsAborted" },
	{ (VT_PTR << TE_VT) + offsetof(SHFILEOPSTRUCT, hNameMappings), L"hNameMappings" },
	{ (VT_BSTR << TE_VT) + offsetof(SHFILEOPSTRUCT, lpszProgressTitle), L"lpszProgressTitle" },
	{ 0, NULL }
};

TEmethod tesSIZE[] =
{
	{ (VT_I4 << TE_VT) + offsetof(SIZE, cx), L"cx" },
	{ (VT_I4 << TE_VT) + offsetof(SIZE, cy), L"cy" },
	{ 0, NULL }
};

TEmethod tesTCHITTESTINFO[] =
{
	{ (VT_CY << TE_VT) + offsetof(TCHITTESTINFO, pt), L"pt" },
	{ (VT_I4 << TE_VT) + offsetof(TCHITTESTINFO, flags), L"flags" },
	{ 0, NULL }
};

TEmethod tesTCITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, mask), L"mask" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, dwState), L"dwState" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, dwStateMask), L"dwStateMask" },
	{ (VT_BSTR << TE_VT) + offsetof(TCITEM, pszText), L"pszText" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, cchTextMax), L"cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, iImage), L"iImage" },
	{ (VT_PTR << TE_VT) + offsetof(TCITEM, lParam), L"lParam" },
	{ 0, NULL }
};

TEmethod tesTVHITTESTINFO[] =
{
	{ (VT_CY << TE_VT) + offsetof(TVHITTESTINFO, pt), L"pt" },
	{ (VT_I4 << TE_VT) + offsetof(TVHITTESTINFO, flags), L"flags" },
	{ (VT_PTR << TE_VT) + offsetof(TVHITTESTINFO, hItem), L"hItem" },
	{ 0, NULL }
};

TEmethod tesTVITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, mask), L"mask" },
	{ (VT_PTR << TE_VT) + offsetof(TVITEM, hItem), L"hItem" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, state), L"state" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, stateMask), L"stateMask" },
	{ (VT_BSTR << TE_VT) + offsetof(TVITEM, pszText), L"pszText" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, cchTextMax), L"cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, iImage), L"iImage" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, iSelectedImage), L"iSelectedImage" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, cChildren), L"cChildren" },
	{ (VT_PTR << TE_VT) + offsetof(TVITEM, lParam), L"lParam" },
	{ 0, NULL }
};

TEmethod tesWIN32_FIND_DATA[] =
{
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, dwFileAttributes), L"dwFileAttributes" },
	{ (VT_FILETIME << TE_VT) + offsetof(WIN32_FIND_DATA, ftCreationTime), L"ftCreationTime" },
	{ (VT_FILETIME << TE_VT) + offsetof(WIN32_FIND_DATA, ftLastAccessTime), L"ftLastAccessTime" },
	{ (VT_FILETIME << TE_VT) + offsetof(WIN32_FIND_DATA, ftLastWriteTime), L"ftLastWriteTime" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, nFileSizeHigh), L"nFileSizeHigh" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, nFileSizeLow), L"nFileSizeLow" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, dwReserved0), L"dwReserved0" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, dwReserved1), L"dwReserved1" },
	{ (VT_LPWSTR << TE_VT) + offsetof(WIN32_FIND_DATA, cFileName), L"cFileName" },
	{ (VT_LPWSTR << TE_VT) + offsetof(WIN32_FIND_DATA, cAlternateFileName), L"cAlternateFileName" },
	{ 0, NULL }
};

TEmethod tesDRAWITEMSTRUCT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, CtlType), L"CtlType" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, CtlID), L"CtlID" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, itemID), L"itemID" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, itemAction), L"itemAction" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, itemState), L"itemState" },
	{ (VT_PTR << TE_VT) + offsetof(DRAWITEMSTRUCT, hwndItem), L"hwndItem" },
	{ (VT_PTR << TE_VT) + offsetof(DRAWITEMSTRUCT, hDC), L"hDC" },
	{ (VT_CARRAY << TE_VT) + offsetof(DRAWITEMSTRUCT, rcItem), L"rcItem" },
	{ (VT_PTR << TE_VT) + offsetof(DRAWITEMSTRUCT, itemData), L"itemData" },
	{ 0, NULL }
};

TEmethod tesMEASUREITEMSTRUCT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, CtlType), L"CtlType" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, CtlID), L"CtlID" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemID), L"itemID" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemWidth), L"itemWidth" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemHeight), L"itemHeight" },
	{ (VT_PTR << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemData), L"itemData" },
	{ 0, NULL }
};

TEmethod tesMENUINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, cbSize), L"cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, fMask), L"fMask" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, dwStyle), L"dwStyle" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, cyMax), L"cyMax" },
	{ (VT_PTR << TE_VT) + offsetof(MENUINFO, hbrBack), L"hbrBack" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, dwContextHelpID), L"dwContextHelpID" },
	{ (VT_PTR << TE_VT) + offsetof(MENUINFO, dwMenuData), L"dwMenuData" },
	{ 0, NULL }
};

TEmethod *pteStruct[] = {
	tesBITMAP,
	tesNULL,	//BSTR
	tesNULL,	//BYTE
	tesNULL,	//char
	tesCHOOSECOLOR,
	tesCHOOSEFONT,
	tesCOPYDATASTRUCT,
	tesNULL,	//DWORD
	tesDIBSECTION,
	tesEncoderParameter,
	tesEncoderParameters,
	tesEXCEPINFO,
	tesFINDREPLACE,
	tesFOLDERSETTINGS,
	tesNULL,	//GUID
	tesNULL,	//HANDLE
	tesHDITEM,
	tesICONINFO,
	tesICONMETRICS,
	tesNULL,	//int
	tesKEYBDINPUT,
	tesNULL,	//KEYSTATE
	tesLOGFONT,
	tesNULL,	//LPWSTR
	tesLVBKIMAGE,
	tesLVFINDINFO,
	tesLVGROUP,
	tesLVHITTESTINFO,
	tesLVITEM,
	tesMENUITEMINFO,
	tesMOUSEINPUT,
	tesMSG,
	tesNONCLIENTMETRICS,
	tesNOTIFYICONDATA,
	tesNMHDR,
	tesOSVERSIONINFOEX,	//OSVERSIONINFO
	tesOSVERSIONINFOEX,
	tesPAINTSTRUCT,
	tesPOINT,
	tesRECT,
	tesSHELLEXECUTEINFO,
	tesSHFILEINFO,
	tesSHFILEOPSTRUCT,
	tesSIZE,
	tesTCHITTESTINFO,
	tesTCITEM,
	tesTVHITTESTINFO,
	tesTVITEM,
	tesWIN32_FIND_DATA,
	tesNULL,	//VARIANT
	tesNULL,	//WCHAR
	tesNULL,	//WORD
	tesDRAWITEMSTRUCT,
	tesMEASUREITEMSTRUCT,
	tesMENUINFO,
};


TEmethod structSize[] = {
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
	{ sizeof(LVITEM), L"HDITEM" },
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
	{ sizeof(DRAWITEMSTRUCT), L"DRAWITEMSTRUCT" },
	{ sizeof(MEASUREITEMSTRUCT), L"MEASUREITEMSTRUCT" },
	{ sizeof(MENUINFO), L"MENUINFO" },
};

TEmethod methodMem2[] = {
	{ VT_I4  << TE_VT, L"int" },
	{ VT_UI4 << TE_VT, L"DWORD" },
	{ VT_UI1 << TE_VT, L"BYTE" },
	{ VT_UI2 << TE_VT, L"WORD" },
	{ VT_UI2 << TE_VT, L"WCHAR" },
	{ VT_PTR << TE_VT, L"HANDLE" },
	{ VT_PTR << TE_VT, L"LPWSTR" },
	{ 0, NULL }
};

TEmethod methodTE[] = {
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
	{ START_OnFunc + TE_OnTranslatePath, L"OnTranslatePath" },
	{ START_OnFunc + TE_OnNavigateComplete, L"OnNavigateComplete" },
	{ START_OnFunc + TE_OnILGetParent, L"OnILGetParent" },
	{ 0, NULL }
};

TEmethod methodAPI[] = {
	{ 1040, L"Memory" },
	{ 1061, L"ContextMenu" },
	{ 1070, L"DropTarget" },
	{ 1080, L"DataObject" },
	{ 1090, L"OleCmdExec" },
	{ 1122, L"sizeof" },
	{ 1132, L"LowPart" },
	{ 1142, L"HighPart" },
	{ 1154, L"QuadPart" },
	{ 1163, L"pvData" },
	{ 1170, L"ExecMethod" },
	{ 1182, L"Extract" },
	{ 1194, L"Add" },
	{ 1204, L"Sub" },
	{ 2063, L"FindWindow" },
	{ 2073, L"FindWindowEx" },
	{ 2080, L"OleGetClipboard" },
	{ 2082, L"OleSetClipboard" },
	{ 2111, L"PathMatchSpec" },
	{ 2230, L"CommandLineToArgv" },
	{ 2431, L"ILIsEqual" },
	{ 2440, L"ILClone" },
	{ 2441, L"ILIsParent" },
	{ 2450, L"ILRemoveLastID" },
	{ 2460, L"ILFindLastID" },
	{ 2461, L"ILIsEmpty" },
	{ 2470, L"ILGetParent" },
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
	{ 6045, L"PathSearchAndQualify"},
	{ 6055, L"PSFormatForDisplay"},
	{ 6065, L"PSGetDisplayName"},
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
	{ 10080, L"keybd_event" },
	{ 10090, L"SHChangeNotify" },
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
	{ 20391, L"EndMenu" },
	{ 20401, L"SHChangeNotifyDeregister" },
	{ 20411, L"SHChangeNotification_Unlock" },
	{ 20421, L"IsWow64Process" },
	{ 20431, L"BitBlt" },
	{ 20441, L"ImmSetOpenStatus" },
	{ 20451, L"ImmReleaseContext" },
	{ 20461, L"IsChild" },
	{ 20471, L"KillTimer" },
	{ 20481, L"AllowSetForegroundWindow" },
	{ 20491, L"SetWindowPos" },
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
	{ 29601, L"ImageList_GetIconSize" },
	{ 29611, L"GetMenuInfo" },
	{ 29621, L"SetMenuInfo" },
	//+mem(2)
	{ 29201, L"SystemParametersInfo" },
	{ 29221, L"GetTextExtentPoint32" },
	{ 29232, L"SHGetDataFromIDList" },
	//+mem(3)
	{ 29301, L"InsertMenuItem" },
	{ 29311, L"GetMenuItemInfo" },
	{ 29321, L"SetMenuItemInfo" },
	{ 29331, L"ChangeWindowMessageFilterEx" },
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
	{ 30222, L"SHEmptyRecycleBin" },
	{ 30232, L"GetMessagePos" },
	{ 30242, L"ImageList_GetOverlayImage" },
	{ 30252, L"ImageList_GetBkColor" },
	{ 30262, L"SetWindowTheme" },
	{ 30272, L"ImmGetVirtualKey" },
	{ 30282, L"GetAsyncKeyState" },
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
	{ 35100, L"SHChangeNotifyRegister" },
	{ 35102, L"MessageBox" },
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
	{ 40273, L"GetParent" },
	{ 40283, L"GetCapture" },
	{ 40293, L"GetModuleHandle" },
	{ 40303, L"SHGetImageList" },
	{ 40313, L"CopyImage" },
	{ 40323, L"GetCurrentProcess" },
	{ 40333, L"ImmGetContext" },
	{ 40343, L"GetFocus" },
	{ 40353, L"GetForegroundWindow" },
	{ 40363, L"SetTimer" },
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
	{ 45143, L"SHChangeNotification_Lock"},
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
	{ 50155, L"AssocQueryString"},
	{ 50165, L"StrFormatByteSize"},
//	{ 51005, L"Test" },
	{ 0, NULL }
};

TEmethod methodSB[] = {
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
	{ 0x10000021, L"TreeView" },
	{ 0x10000024, L"Parent" },
	{ 0x10000031, L"Close" },
	{ 0x10000032, L"Title" },
	{ 0x10000033, L"Suspend" },
	{ 0x10000040, L"Items" },
	{ 0x10000041, L"SelectedItems" },
	{ 0x10000050, L"ShellFolderView" },
	{ 0x10000102, L"hwndList" },
	{ 0x10000103, L"hwndView" },
	{ 0x10000104, L"SortColumn" },
	{ 0x10000106, L"Focus" },
	{ 0x10000107, L"HitTest" },
	{ 0x10000108, L"Droptarget" },
	{ 0x10000109, L"Columns"},
	{ 0x10000110, L"ItemCount" },
	{ 0x10000206, L"Refresh" },
	{ 0x10000207, L"ViewMenu" },
	{ 0x10000208, L"TranslateAccelerator" },
	{ 0x10000209, L"GetItemPosition" },
	{ 0x1000020A, L"SelectAndPositionItem" },
	{ 0x10000280, L"SelectItem" },
	{ 0x10000281, L"FocusedItem" },
	{ 0x10000282, L"GetFocusedItem" },
	{ 0x10000501, L"AddItem" },
	{ 0x10000502, L"RemoveItem" },
	{ START_OnFunc + SB_OnIncludeObject, L"OnIncludeObject" },
	{ 0, NULL }
};

TEmethod methodWB[] = {
	{ 0x10000001, L"Data" },
	{ 0x10000002, L"hwnd" },
	{ 0x10000003, L"Type" },
	{ 0x10000004, L"TranslateAccelerator" },
	{ 0x10000005, L"Application" },
	{ 0x10000006, L"Document" },
	{ 0, NULL }
};

TEmethod methodTC[] = {
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
	{ 0, NULL }
};

TEmethod methodFIs[] = {
	{ 2, L"Application" },
	{ 3, L"Parent" },
	{ 8, L"AddItem" },
	{ 9, L"hDrop" },
	{ 10, L"GetData" },
	{ 11, L"SetData" },
	{ DISPID_NEWENUM, L"_NewEnum" },
	{ DISPID_TE_ITEM, L"Item" },
	{ DISPID_TE_COUNT, L"Count" },
	{ DISPID_TE_COUNT, L"length" },
	{ DISPID_TE_INDEX, L"Index" },
	{ 0x10000001, L"lEvent" },
	{ 0x10000001, L"dwEffect" },
	{ 0x10000002, L"pdwEffect" },
	{ 0, NULL }
};

TEmethod methodDT[] = {
	{ 1, L"DragEnter" },
	{ 2, L"DragOver" },
	{ 3, L"Drop" },
	{ 4, L"DragLeave" },
	{ 5, L"Type" },
	{ 6, L"FolderItem" },
	{ 0, NULL }
};

TEmethod methodTV[] = {
	{ 0x10000001, L"Data" },
	{ 0x10000002, L"Type" },
	{ 0x10000003, L"hwnd" },
	{ 0x10000004, L"Close" },
	{ 0x10000005, L"hwndTree" },
	{ 0x10000007, L"FolderView" },
	{ 0x10000008, L"Align" },
	{ 0x10000009, L"Visible" },
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
	{ 0, NULL }
};

TEmethod methodFI[] = {
	{ 1, L"Name" },
	{ 2, L"Path" },
	{ 3, L"Alt" },
//	{ 4, L"FocusedItem" },
	{ 9, L"_BLOB" },
	{ 0, NULL }
};

TEmethod methodMem[] = {
	{ 1, L"P" },
	{ 4, L"Read" },
	{ 5, L"Write" },
	{ 6, L"Size" },
	{ 7, L"Free" },
	{ 8, L"Clone" },
	{ 9, L"_BLOB" },
	{ DISPID_NEWENUM, L"_NewEnum" },
	{ DISPID_TE_ITEM, L"Item" },
	{ DISPID_TE_COUNT, L"Count" },
	{ DISPID_TE_COUNT, L"length" },
	{ 0, NULL }
};

TEmethod methodCM[] = {
	{ 1, L"QueryContextMenu" },
	{ 2, L"InvokeCommand" },
	{ 3, L"Items" },
	{ 4, L"GetCommandString" },
	{ 5, L"FolderView" },
	{ 6, L"HandleMenuMsg" },
	{ 10, L"hmenu" },
	{ 11, L"indexMenu" },
	{ 12, L"idCmdFirst" },
	{ 13, L"idCmdLast" },
	{ 14, L"uFlags" },
	{ 0, NULL }
};

TEmethod methodCD[] = {
	{ 40, L"ShowOpen" },
	{ 41, L"ShowSave" },
//	{ 42, L"ShowFolder" },
	{ 10, L"FileName" },
	{ 13, L"Filter" },
	{ 20, L"InitDir" },
	{ 21, L"DefExt" },
	{ 22, L"Title" },
	{ 30, L"MaxFileSize" },
	{ 31, L"Flags" },
	{ 32, L"FilterIndex" },
	{ 31, L"FlagsEx" },
	{ 0, NULL }
};

TEmethod methodGB[] = {
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
	{ 0, NULL }
};

// Forward declarations of functions included in this code module:
ATOM MyRegisterClass(HINSTANCE hInstance, LPWSTR szClassName);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
VOID ArrangeWindow();

BOOL GetIDListFromObject(IUnknown *punk, LPITEMIDLIST *ppidl);
VOID CALLBACK teTimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);

//Unit

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
					SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) & ~WS_EX_TOOLWINDOW);
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
	if (lpfnChangeWindowMessageFilterEx) {
		return lpfnChangeWindowMessageFilterEx(hwnd, message, action, pChangeFilterStruct);
	}
	if (lpfnChangeWindowMessageFilter) {
		return lpfnChangeWindowMessageFilter(message, action);
	}
	return FALSE;
}

BOOL teGetBSTRFromFolderItem(BSTR *pbs, IUnknown *punk)
{
	CteFolderItem *pid;
	if (punk && SUCCEEDED(punk->QueryInterface(g_ClsIdFI, (LPVOID *)&pid))) {
		if (pid->m_v.vt == VT_BSTR) {
			*pbs = pid->m_v.bstrVal;
			pid->Release();
			return (*pbs != NULL);
		}
		pid->Release();
	}
	return FALSE;
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

VOID teRevoke()
{
	try {
		if (g_hMutex) {
			ReleaseMutex(g_hMutex);
			::CloseHandle(g_hMutex);
			g_hMutex = NULL;
		}
		if (g_pSW) {
			if (g_lSWCookie) {
				g_pSW->Revoke(g_lSWCookie);
				g_lSWCookie = 0;
			}
			if (g_dwSWCookie) {
				teUnadviseAndRelease(g_pSW, DIID_DShellWindowsEvents, &g_dwSWCookie);
				g_dwSWCookie = 0;
			}
		}
	} catch (...) {}

}

VOID teSetRedraw(BOOL bSetRedraw)
{
	SendMessage(g_hwndBrowser, WM_SETREDRAW, bSetRedraw, 0);
	if (bSetRedraw && !g_bSetRedraw) {
		RedrawWindow(g_hwndBrowser, NULL, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
	}
	g_bSetRedraw = bSetRedraw;
}

HRESULT teException(EXCEPINFO *pExcepInfo, UINT *puArgErr, LPWSTR lpszObj, TEmethod* pMethod, DISPID dispIdMember)
{
	if (pExcepInfo && puArgErr) {
		LPWSTR pszName = NULL;
		if (pMethod) {
			int i = 0;
			while (pMethod[i].id && pMethod[i].id != dispIdMember) {
				i++;
			}
			pszName = pMethod[i].name;
		}
		int nLen = lstrlen(lpszObj) + lstrlen(pszName) + 15;
		pExcepInfo->bstrDescription = SysAllocStringLen(NULL, nLen);
		swprintf_s(pExcepInfo->bstrDescription, nLen, L"Exception in %s.%s", lpszObj, pszName ? pszName : L"");
		pExcepInfo->bstrSource = SysAllocString(g_szTE);
		pExcepInfo->scode = DISP_E_EXCEPTION;

	}
	return DISP_E_EXCEPTION;
}

BOOL teStrSameIFree(BSTR bs, LPWSTR lpstr2)
{
	BOOL b = lstrcmpi(bs, lpstr2) == 0;
	::SysFreeString(bs);
	return b;
}

VOID teCoTaskMemFree(LPVOID pv)
{
	if (pv) {
		::CoTaskMemFree(pv);
	}
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
	if (ppidl && *ppidl) {
		::CoTaskMemFree(*ppidl);
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

#ifndef _WIN64
VOID Wow64ControlPanel(LPITEMIDLIST *ppidl, LPITEMIDLIST pidlReplace)
{
	BOOL bIsWow64 = FALSE;
	if (lpfnIsWow64Process) {
		lpfnIsWow64Process(GetCurrentProcess(), &bIsWow64);
		if (bIsWow64) {
			if (ILIsParent(g_pidlCP, *ppidl, FALSE) && !ILIsEqual(*ppidl, g_pidls[CSIDL_CONTROLS])) {
				teILCloneReplace(ppidl, pidlReplace);
			}
		}
	}
}
#endif

BOOL tePathMatchSpec(LPCWSTR pszFile, LPCWSTR pszSpec)//bugfix PathMatchSpec for Win 8/8.1
{
	if (pszSpec[0] == '*') {
		WCHAR wc;
		int i = 0;
		do {
			wc = pszSpec[++i];
		} while (wc && wc != '*' && wc != '?' && wc != ';');
		if (wc == '*') {
			if (i > 1 && !pszSpec[i + 1]) {
				LPWSTR psz = const_cast<LPWSTR>(pszFile);
				i -= 2;
				while (psz = StrChrIW(psz, pszSpec[1])) {
					if (i == 0 || StrCmpNIW(++psz, &pszSpec[2], i) == 0) {
						return TRUE;
					}
				}
				return i == 0 && pszSpec[1] == '.';
			}
		}
		if (!pszSpec[1]) {
			return TRUE;
		}
	}
	return PathMatchSpec(pszFile, pszSpec);
}

HRESULT STDAPICALLTYPE tePSPropertyKeyFromStringEx(__in LPCWSTR pszString,  __out PROPERTYKEY *pkey)
{
	HRESULT hr = lpfnPSPropertyKeyFromString(pszString, pkey);
	return SUCCEEDED(hr) ? hr : lpfnPSGetPropertyKeyFromName(pszString, pkey);
}

int CalcCrc32(BYTE *pc, int nLen, UINT c)
{
	c ^= MAXUINT;
	for (int i = 0; i < nLen; i++) {
		c = g_uCrcTable[(c ^ pc[i]) & 0xff] ^ (c >> 8);
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
		*pbsPath = SysAllocStringLen(NULL, nSize);
		i = DragQueryFile(hDrop, iFile, *pbsPath, nSize);
		if (i + 1 < nSize) {
			(*pbsPath)[i] = NULL;
			break;
		}
		SysFreeString(*pbsPath);
	}
	return i;
}

VOID tePathAppend(BSTR *pbsPath, LPWSTR pszPath, LPWSTR pszFile)
{
	*pbsPath = SysAllocStringLen(pszPath, lstrlen(pszPath) + lstrlen(pszFile) + 1);
	PathAppend(*pbsPath, pszFile);
}

BOOL teSetObject(VARIANT *pv, PVOID pObj)
{
	if (pv && pObj) {
		try {
			IUnknown *punk = static_cast<IUnknown *>(pObj);
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->pdispVal))) {
				pv->vt = VT_DISPATCH;
				return true;
			}
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->punkVal))) {
				pv->vt = VT_UNKNOWN;
				return true;
			}
		} catch (...) {}
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
					return true;
				}
				if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->punkVal))) {
					pv->vt = VT_UNKNOWN;
					punk->Release();
					return true;
				}
			}
			punk->Release();
		} catch (...) {}
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
	return false;
}

VOID ClearEvents()
{
	for (int j = Count_OnFunc; j-- > 0;) {
		if (g_pOnFunc[j]) {
			g_pOnFunc[j]->Release();
			g_pOnFunc[j] = NULL;
		}
	}

	int i = MAX_FV;
	while (--i >= 0) {
		if (g_pSB[i]) {
			for (int j = Count_SBFunc; j-- > 0;) {
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

BSTR SysAllocStringLenEx(const OLECHAR *strIn, UINT uSize, UINT uOrg)
{
	BSTR bs = SysAllocStringLen(NULL, uSize);
	lstrcpyn(bs, strIn, uSize < uOrg ? uSize : uOrg);
	return bs;
}

int ILGetCount(LPITEMIDLIST pidl)
{
	if (ILIsEmpty(pidl)) {
		return 0;
	}
	return ILGetCount(ILGetNext(pidl)) + 1;
}

HRESULT teGetDisplayNameBSTR(IShellFolder *pSF, PCUITEMID_CHILD pidl, SHGDNF uFlags, BSTR *pbs)
{
	STRRET strret;
	HRESULT hr = pSF->GetDisplayNameOf(pidl, uFlags, &strret);
	if SUCCEEDED(hr) {
		hr = StrRetToBSTR(&strret, pidl, pbs);
	}
	return hr;
}

LPITEMIDLIST teILCreateFromPath2(IShellFolder *pSF, LPWSTR pszPath, SHGDNF uFlags, BSTR bsParent)
{
	LPITEMIDLIST pidl = NULL;
	IEnumIDList *peidl = NULL;
	BSTR bstr = NULL;
	int ashgdn[] = {SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, SHGDN_INFOLDER | SHGDN_FORPARSING, SHGDN_INFOLDER};
	int nLen;
	if SUCCEEDED(pSF->EnumObjects(NULL, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN, &peidl)) {
		LPITEMIDLIST pidlPart;
		ULONG celtFetched;
		while (peidl->Next(1, &pidlPart, &celtFetched) == S_OK) {
			STRRET strret;
			for (int j = 0; j < 3; j++) {
				SHGDNF uFlag = uFlags & ashgdn[j];
				if (uFlag && SUCCEEDED(pSF->GetDisplayNameOf(pidlPart, uFlag, &strret))) {
					if SUCCEEDED(StrRetToBSTR(&strret, pidlPart, &bstr)) {
						BSTR bsFull = bstr;
						if (j && bsParent) {
							tePathAppend(&bsFull, bsParent, bstr);
							::SysFreeString(bstr);
						}
						if (lstrcmpi(const_cast<LPCWSTR>(pszPath), (LPCWSTR)bsFull) == 0) {
							::SysFreeString(bsFull);
							LPITEMIDLIST pidlParent;
							if (GetIDListFromObject(pSF, &pidlParent)) {
								pidl = ILCombine(pidlParent, pidlPart);
								teCoTaskMemFree(pidlParent);
								teCoTaskMemFree(pidlPart);
								peidl->Release();
								return pidl;
							}
							return NULL;
						}
						nLen = lstrlen(bsFull);
						if (nLen && StrCmpNI(const_cast<LPCWSTR>(pszPath), (LPCWSTR)bsFull, nLen) == 0) {
							IShellFolder *pSF1;
							if SUCCEEDED(pSF->BindToObject(pidlPart, NULL, IID_PPV_ARGS(&pSF1))) {
								pidl = teILCreateFromPath2(pSF1, pszPath, uFlag, bsFull);
								pSF1->Release();
								if (pidl) {
									return pidl;
								}
							}
						}
						::SysFreeString(bsFull);
					}
				}
			}
			teCoTaskMemFree(pidlPart);
		}
		peidl->Release();
	}
	return NULL;
}

BOOL GetCSIDLFromPath(int *i, LPWSTR pszPath)
{
	int n = lstrlen(pszPath);
	if (n <= 3 && pszPath[0] >= '0' && pszPath[0] <= '9') {
		swscanf_s(pszPath, L"%d", i);
		return (*i <= 9 && n == 1) || (*i >= 10 && *i <= 99 && n == 2) || (*i >= 100 && *i < MAX_CSIDL);
	}
	return FALSE;
}

HRESULT GetDisplayNameFromPidl(BSTR *pbs, LPITEMIDLIST pidl, SHGDNF uFlags)
{
	HRESULT hr = E_FAIL;
	IShellFolder *pSF;
	LPCITEMIDLIST ItemID;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &ItemID)) {
		STRRET strret;
		if SUCCEEDED(pSF->GetDisplayNameOf(ItemID, uFlags & (~SHGDN_FORPARSINGEX), &strret)) {
			if SUCCEEDED(StrRetToBSTR(&strret, ItemID, pbs)) {
				hr = S_OK;
			}
		}
		if (hr == S_OK) {
			if (teIsFileSystem(*pbs)) {
				if (uFlags & SHGDN_FORPARSINGEX) {
					LPITEMIDLIST pidl2 = SHSimpleIDListFromPath(*pbs);
					if (!ILIsEqual(pidl, pidl2)) {
						teILCloneReplace(&pidl2, pidl);
						BSTR bs2;
						BSTR bs3 = *pbs;
						*pbs = NULL;
						while (!ILIsEmpty(pidl2)) {
							BSTR bs;
							if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl2, SHGDN_INFOLDER | SHGDN_FORPARSING)) {
								if (*pbs) {
									tePathAppend(&bs2, bs, *pbs);
									::SysFreeString(bs);
									::SysFreeString(*pbs);
									*pbs = bs2;
								}
								else {
									*pbs = bs;
								}
								::ILRemoveLastID(pidl2);
							}
						}
						int n;
						if (GetCSIDLFromPath(&n, *pbs)) {
							bs2 = *pbs;
							*pbs = bs3;
							bs3 = bs2;
						}
						::SysFreeString(bs3);
					}
					teCoTaskMemFree(pidl2);
				}
			}
			else if (((uFlags & (SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) == (SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) &&
				ILGetCount(pidl) == 1) {
				LPITEMIDLIST pidl2 = teILCreateFromPath2(pSF, *pbs, SHGDN_FORPARSING | SHGDN_FORADDRESSBAR | SHGDN_INFOLDER, NULL);
				if (!ILIsEqual(pidl, pidl2)) {
					if SUCCEEDED(pSF->GetDisplayNameOf(ItemID, uFlags & (~(SHGDN_FORPARSINGEX | SHGDN_FORADDRESSBAR)), &strret)) {
						if SUCCEEDED(StrRetToBSTR(&strret, ItemID, pbs)) {
							hr = S_OK;
						}
					}
				}
				teCoTaskMemFree(pidl2);
			}
		}
		pSF->Release();
	}
	return hr;
}

BOOL teIsDesktopPath(BSTR bs)
{
	BOOL bResult = TRUE;
	BSTR bs2, bs3;
	if SUCCEEDED(GetDisplayNameFromPidl(&bs2, g_pidls[CSIDL_DESKTOP], SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
		if (!teStrSameIFree(bs2, bs)) {
			if SUCCEEDED(GetDisplayNameFromPidl(&bs3, g_pidls[CSIDL_DESKTOP], SHGDN_FORPARSING)) {
				bResult = teStrSameIFree(bs3, bs);
			}
		}
	}
	return bResult;
}

BOOL IsSearchPath(LPITEMIDLIST pidl)
{
	BOOL bSearch = FALSE;
	BSTR bs;
	if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
		bSearch = PathMatchSpec(bs, L"search-ms:*");
		::SysFreeString(bs);
	}
	return bSearch;
}

FOLDERVIEWOPTIONS teGetFolderViewOptions(LPITEMIDLIST pidl, UINT uViewMode)
{
	if (IsSearchPath(pidl)) {
		if (uViewMode == FVM_DETAILS) {
			return FVO_VISTALAYOUT;
		}
	}
	else if (uViewMode < FVM_THUMBSTRIP) {
		return FVO_VISTALAYOUT;
	}
	return FVO_DEFAULT;
}

void GetVarPathFromPidl(VARIANT *pVarResult, LPITEMIDLIST pidl, int uFlags)
{
	int i;
	BOOL bSpecial = FALSE;

	if (uFlags & SHGDN_FORPARSINGEX) {
		for (i = 0; i < g_nPidls; i++) {
			if (g_pidls[i] && ::ILIsEqual(pidl, g_pidls[i])) {
				LPITEMIDLIST pidl2 = NULL;
				BSTR bsPath;
				if SUCCEEDED(GetDisplayNameFromPidl(&bsPath, pidl, SHGDN_FORPARSING)) {
					pidl2 = SHSimpleIDListFromPath(bsPath);
					::SysFreeString(bsPath);
				}
				if (!::ILIsEqual(pidl, pidl2)) {
					pVarResult->lVal = i;
					pVarResult->vt = VT_I4;
				}
				teCoTaskMemFree(pidl2);
				break;
			}
		}
	}
	if (pVarResult->vt != VT_I4) {
		if SUCCEEDED(GetDisplayNameFromPidl(&pVarResult->bstrVal, pidl, uFlags)) {
			pVarResult->vt = VT_BSTR;
		}
	}
}

BOOL GetMemFromVariant(VARIANT *pv, CteMemory **ppMem)
{
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		return SUCCEEDED(punk->QueryInterface(g_ClsIdStruct, (LPVOID *)ppMem));
	}
	return FALSE;
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
	teClearVariantArgs(nArgs, pvArgs);
	return hr;
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

HRESULT teGetPropertyAt(IDispatch *pdisp, int i, VARIANT *pv)
{
	wchar_t szName[8];
	swprintf_s((wchar_t *)&szName, 8, L"%d", i);
	return teGetProperty(pdisp, szName, pv);
}

HRESULT tePutPropertyAt(PVOID pObj, int nIndex, VARIANT *pvValue)
{
	VARIANTARG *pv = GetNewVARIANT(3);
	teSetObject(&pv[2], pObj);
	pv[1].vt = VT_I4;
	pv[1].lVal = nIndex;
	VariantCopy(&pv[0], pvValue);
	return teExecMethod(g_pJS, L"_p", NULL, 3, pv);
}

char* GetpcFromVariant(VARIANT *pv)
{
	char *pc1 = NULL;
	CteMemory *pMem;
	if (GetMemFromVariant(pv, &pMem)) {
		char *pc = pMem->m_pc;
		pMem->Release();
		return pc;
	}
	IDispatch *pdisp;
	if (GetDispatch(pv, &pdisp)) {
		VARIANT v;
		VariantInit(&v);
		teGetProperty(pdisp, L"_BLOB", &v);
		pdisp->Release();
		if (v.vt == VT_BSTR) {
			return (char *)v.bstrVal;
		}
		VariantClear(&v);
	}
	return NULL;
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
		char *pc = GetpcFromVariant(pv);
		if (pc) {
			return (LONGLONG)pc;
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
		int nSize = lstrlen(strCmdLine) + MAX_PROP;
		BSTR bsCmdLine = SysAllocStringLen(NULL, nSize);
		int j = 0;
		int i = 0;
		while (i < nSize && strCmdLine[j]) {
			if (StrCmpNI(&strCmdLine[j], L"/idlist,:", 9) == 0) {
				LONGLONG hData;
				DWORD dwProcessId;
				WCHAR sz[MAX_PATH + 2];
				swscanf_s(&strCmdLine[j + 9], L"%32[1234567890]s", sz, _countof(sz));
				swscanf_s(sz, L"%lld", &hData);
				int k = lstrlen(sz);
				swscanf_s(&strCmdLine[j + k + 10], L"%32[1234567890]s", sz, _countof(sz));
				swscanf_s(sz, L"%d", &dwProcessId);
				j += k + lstrlen(sz) + 11;
#ifdef _2000XP
				LPITEMIDLIST pidl = (LPITEMIDLIST)SHLockShared((HANDLE)hData, dwProcessId);
				if (pidl) {
					VARIANT v, v2;
					GetVarPathFromPidl(&v, pidl, SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
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
						j++;
					}
				}
#endif
			}
			else {
				bsCmdLine[i++] = strCmdLine[j++];
			}
		}
		g_bsCmdLine = ::SysAllocStringLen(bsCmdLine, i);
		::SysFreeString(bsCmdLine);
	}
	return ::SysAllocString(g_bsCmdLine);
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
	if (*ppSF == NULL) {
		SHGetDesktopFolder(ppSF);
	}
#ifdef _2000XP
	if (!g_bUpperVista) {
		if (ILIsEqual(pidl, g_pidlResultsFolder)) {
			IPersistFolder *pPF;
			if SUCCEEDED((*ppSF)->QueryInterface(IID_PPV_ARGS(&pPF))) {
				pPF->Initialize(pidl);
				pPF->Release();
			}
		}
	}
#endif
	return true;
}

BOOL teILIsExists(LPITEMIDLIST pidl)
{
	BOOL bResult = FALSE;
	IShellFolder *pSF;
	if (GetShellFolder(&pSF, pidl)) {
		IEnumIDList *pEI;
		if SUCCEEDED(pSF->EnumObjects(NULL, SHCONTF_NONFOLDERS |SHCONTF_INCLUDEHIDDEN | SHCONTF_FOLDERS, &pEI)) {
			bResult = TRUE;
			pEI->Release();
		}
		pSF->Release();
	}
	return bResult;
}

LPITEMIDLIST teILCreateFromPath(LPWSTR pszPath)
{
	LPITEMIDLIST pidl = NULL;
	BSTR bstr = NULL;
	int n;

	if (pszPath) {
		BSTR bsPath2 = NULL;
		if (pszPath[0] == _T('"')) {
			bsPath2 = ::SysAllocStringLen(pszPath, lstrlen(pszPath) + 1);
			PathUnquoteSpaces(bsPath2);
			pszPath = bsPath2;
		}
		BSTR bsPath3 = NULL;
		if (PathMatchSpec(pszPath, L"*\\..\\*;*\\..;*\\.\\*;*\\.;*%*%*")) {
			UINT uLen = lstrlen(pszPath) + MAX_PATH;
			bsPath3 = ::SysAllocStringLen(NULL, uLen);
			PathSearchAndQualify(pszPath, bsPath3, uLen);
			pszPath = bsPath3;
		}
		else if (lstrlen(pszPath) == 1 && pszPath[0] >= 'A') {
			bsPath3 = ::SysAllocStringLen(L"?:\\", 4);
			bsPath3[0] = pszPath[0];
			pszPath = bsPath3;
		}
		if (GetCSIDLFromPath(&n, pszPath)) {
			pidl = ::ILClone(g_pidls[n]);
			pszPath = NULL;
		}
		n = PathGetDriveNumber(pszPath);
		if (n >= 0 && DriveType(n) == DRIVE_NO_ROOT_DIR && lstrlen(pszPath) > 3) {
			WCHAR szDrive[4];
			lstrcpyn(szDrive, pszPath, 4);
			lpfnSHParseDisplayName(szDrive, NULL, &pidl, 0, NULL);
			if (!pidl || g_dwMainThreadId == GetCurrentThreadId() || !teILIsExists(pidl)) {
				pszPath = NULL;
			}
			teILFreeClear(&pidl);
		}
		if (pszPath) {
			lpfnSHParseDisplayName(pszPath, NULL, &pidl, 0, NULL);
/*/// To parse too much.
			if (pidl == NULL && PathMatchSpec(pszPath, L"::{*")) {  
				int nSize = lstrlen(pszPath) + 6;
				BSTR bsPath4 = ::SysAllocStringLen(L"shell:", nSize);
				lstrcat(bsPath4, pszPath);
				lpfnSHParseDisplayName(bsPath4, NULL, &pidl, 0, NULL);
				::SysFreeString(bsPath4);
			}
//*/
			if (pidl == NULL && !teIsFileSystem(pszPath)) {
				for (int i = 0; i < g_nPidls; i++) {
					if (g_pidls[i]) {
						if SUCCEEDED(GetDisplayNameFromPidl(&bstr, g_pidls[i], SHGDN_FORADDRESSBAR)) {
							if (teStrSameIFree(bstr, pszPath)) {
								pidl = ILClone(g_pidls[i]);
								break;
							}
						}
					}
				}
				if (pidl == NULL) {
					IShellFolder *pSF;
					if (SHGetDesktopFolder(&pSF) == S_OK) {
						pidl = teILCreateFromPath2(pSF, pszPath, SHGDN_FORPARSING | SHGDN_FORADDRESSBAR | SHGDN_INFOLDER, NULL);
						pSF->Release();
					}
					if (pidl == NULL) {
						if (GetShellFolder(&pSF, g_pidls[CSIDL_DRIVES])) {
							pidl = teILCreateFromPath2(pSF, pszPath, SHGDN_FORPARSING | SHGDN_INFOLDER, NULL);
							pSF->Release();
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
	if (lpfnSHCreateItemFromIDList) {
		LPITEMIDLIST pidl = teILCreateFromPath(const_cast<LPWSTR>(pszPath));
		if (pidl) {
			Result = SUCCEEDED(lpfnSHCreateItemFromIDList(pidl, IID_PPV_ARGS(ppSI)));
			CoTaskMemFree(pidl);
		}
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
		Result = TRUE;
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
		if (lpfnSHCreateItemFromIDList && SUCCEEDED(CoCreateInstance(CLSID_FileOperation, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pFO)))) {
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
							}
							else if (pszTo && *pszTo) {
								if (pFOS->wFunc == FO_RENAME) {
									hr = pFO->RenameItem(psiFrom, pszTo, NULL);
								}
								else {
									if (psiTo || GetDestAndName(pszTo, &psiTo, &pszName)) {
										if (pFOS->wFunc == FO_COPY) {
											hr = pFO->CopyItem(psiFrom, psiTo, pszName, NULL);
										}
										else if (pFOS->wFunc == FO_MOVE) {
											if (!pszName || PathFindFileName(pszFrom) - pszFrom != pszName - pszTo || StrCmpNI(pszFrom, pszTo, (int)(pszName - pszTo))) {
												hr = pFO->MoveItem(psiFrom, psiTo, pszName, NULL);
											}
											else {
												hr = pFO->RenameItem(psiFrom, pszName, NULL);
											}
										}
									}
								}
								if (pFOS->fFlags & FOF_MULTIDESTFILES) {
									if (psiTo) {
										psiTo->Release();
										psiTo = NULL;
									}
									while (*(pszTo++));
								}
							}
							psiFrom->Release();
						}
						while (*(pszFrom++));
					}
					psiTo && psiTo->Release();
					if SUCCEEDED(hr) {
						hr = pFO->PerformOperations();
						pFO->GetAnyOperationsAborted(&pFOS->fAnyOperationsAborted);
					}
				}
				catch (...) {
				}
			}
			pFO->Release();
		}
	}
	return hr == E_NOTIMPL ? ::SHFileOperation(pFOS) : hr;
}

static void threadFileOperation(void *args)
{
	::InterlockedIncrement(&g_nThreads);
	::CoInitialize(NULL);
	LPSHFILEOPSTRUCT pFO = (LPSHFILEOPSTRUCT)args;
	try {
		teSHFileOperation(pFO);
	}
	catch (...) {
	}
	::SysFreeString(const_cast<BSTR>(pFO->pTo));
	::SysFreeString(const_cast<BSTR>(pFO->pFrom));
	delete [] pFO;
	::CoUninitialize();
	::InterlockedDecrement(&g_nThreads);
	::_endthread();
}

/*/// test thread and interface
static void threadDrop(void *args)
{
	::InterlockedIncrement(&g_nThreads);
	::OleInitialize(NULL);
	TEDrop *pDrop = (TEDrop *)args;
	try {
		IDropTarget *pDropTarget;
		if SUCCEEDED(CoGetInterfaceAndReleaseStream(pDrop->pDropTarget, IID_PPV_ARGS(&pDropTarget))) {
			IDataObject *pDataObj;
			if SUCCEEDED(CoGetInterfaceAndReleaseStream(pDrop->pDataObj, IID_PPV_ARGS(&pDataObj))) {
				DWORD dwEffect = pDrop->dwEffect;
				HRESULT hr = pDropTarget->DragEnter(pDataObj, pDrop->grfKeyState, pDrop->pt, &dwEffect);
				if (hr == S_OK) {
					dwEffect = pDrop->dwEffect;
					hr = pDropTarget->DragOver(pDrop->grfKeyState, pDrop->pt, &dwEffect);
					if (hr == S_OK) {
						dwEffect = pDrop->dwEffect;
						hr = pDropTarget->Drop(pDataObj, pDrop->grfKeyState, pDrop->pt, &dwEffect);
					}
				}
				pDropTarget->DragLeave();
			}
		}
	}
	catch (...) {
	}
	delete [] pDrop;
	::OleUninitialize();
	::InterlockedDecrement(&g_nThreads);
	::_endthread();
}

HRESULT teDrop(IDropTarget *pDropTarget, IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	if (g_dwMainThreadId == GetCurrentThreadId()) {
		TEDrop *pDrop = new TEDrop[1];
		CoMarshalInterThreadInterfaceInStream(IID_IDropTarget, pDropTarget, &pDrop->pDropTarget);
		CoMarshalInterThreadInterfaceInStream(IID_IDataObject, pDataObj, &pDrop->pDataObj);
		pDrop->grfKeyState = grfKeyState;
		pDrop->pt = pt;
		pDrop->dwEffect = *pdwEffect;
		_beginthread(&threadDrop, 0, pDrop);
		return S_OK;
	}
	return pDropTarget->Drop(pDataObj, grfKeyState, pt, pdwEffect);
}
//*/
HRESULT GetShellFolder2(IShellFolder2 **ppSF2, LPCITEMIDLIST pidl, BSTR bsFilter)
{
	HRESULT hr = E_FAIL;
	IShellFolder *pSF = NULL;
	if (*ppSF2) {
		(*ppSF2)->Release();
		*ppSF2 = NULL;
	}
	if (bsFilter && g_pidlLibrary && ILIsParent(g_pidlLibrary, pidl, FALSE)) {
		BSTR bs;
		GetDisplayNameFromPidl(&bs, const_cast<LPITEMIDLIST>(pidl), SHGDN_FORPARSING);
		if (teIsFileSystem(bs)) {
			LPITEMIDLIST pidl2 = teILCreateFromPath(bs);
			GetShellFolder(&pSF, pidl2);
			teCoTaskMemFree(pidl2);
		}
		::SysFreeString(bs);
	}
	if (!pSF) {
		GetShellFolder(&pSF, pidl);
	}
	if (pSF) {
		hr = pSF->QueryInterface(IID_PPV_ARGS(ppSF2));
		pSF->Release();
	}
	return hr;
}

int teGetMenuString(BSTR *pbs, HMENU hMenu, UINT uIDItem, BOOL fByPosition)
{
	MENUITEMINFO mii;
	::ZeroMemory(&mii, sizeof(MENUITEMINFO));
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask  = MIIM_STRING;
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
	BSTR bs = SysAllocStringByteLen(NULL, nLen);
	bs[0] = L'-';
	::CopyMemory(&bs[1], *pbs, nLen);
	::SysFreeString(*pbs);
	*pbs = bs;
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
	}//end_switch
	return 0;
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
		BSTR bsTypeData = SysAllocStringLen(NULL, mii.cch++);
		mii.dwTypeData = bsTypeData;
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
		SysFreeString(bsTypeData);
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

CteShellBrowser* SBfromhwnd(HWND hwnd)
{
	int i = MAX_FV;
	while (--i >= 0 && g_pSB[i]) {
		if (g_pSB[i]->m_hwnd == hwnd || IsChild(g_pSB[i]->m_hwnd, hwnd)) {
			return g_pSB[i];
		}
	}
	return NULL;
}

CteTabs* TCfromhwnd(HWND hwnd)
{
	int i = MAX_TC;
	while (--i >= 0 && g_pTC[i]) {
		if (g_pTC[i]->m_hwnd == hwnd || g_pTC[i]->m_hwndStatic == hwnd || g_pTC[i]->m_hwndButton == hwnd) {
			return g_pTC[i];
		}
	}
	return NULL;
}

CteTreeView* TVfromhwnd(HWND hwnd, BOOL bAll)
{
	int i = MAX_FV;
	while (--i >= 0 && g_pSB[i]) {
		HWND hwndTV = g_pSB[i]->m_pTV->m_hwnd;
		if (hwndTV == hwnd || IsChild(hwndTV, hwnd)) {
			if (bAll || g_pSB[i]->m_pTV->m_bMain) {
				return g_pSB[i]->m_pTV;
			}
			return NULL;
		}
	}
	return NULL;
}

void CheckChangeTabSB(HWND hwnd)
{
	if (g_pTabs) {
		CteShellBrowser *pSB = SBfromhwnd(hwnd);
		if (!pSB) {
			CteTreeView *pTV = TVfromhwnd(hwnd, false);
			if (pTV) {
				pSB = pTV->m_pFV;
			}
		}
		if (pSB && pSB->m_pTabs->m_bVisible) {
			if (g_pTabs->m_hwnd != pSB->m_pTabs->m_hwnd) {
				g_pTabs = pSB->m_pTabs;
				pSB->m_pTabs->TabChanged(false);
			}
		}
	}
}

void CheckChangeTabTC(HWND hwnd, BOOL bFocusSB)
{
	CteTabs *pTC = TCfromhwnd(hwnd);
	if (pTC && pTC->m_bVisible) {
		if (g_pTabs->m_hwnd != pTC->m_hwnd) {
			g_pTabs = pTC;
			pTC->TabChanged(false);
		}
		if (bFocusSB) {
			CteShellBrowser *pSB;
			pSB = pTC->GetShellBrowser(pTC->m_nIndex);
			if (pSB) {
				pSB->SetActive(TRUE);
			}
		}
	}
}

BOOL AdjustIDList(LPITEMIDLIST *ppidllist, int nCount)
{
	if (ppidllist == NULL || nCount <= 0) {
		return FALSE;
	}
	if (ppidllist[0]) {
#ifdef _2000XP
		if (g_bUpperVista || !ILIsEqual(ppidllist[0], g_pidlResultsFolder)) {
			return FALSE;
		}
		for (int i = nCount; i > 1; i--) {
			LPITEMIDLIST pidl = ppidllist[i];
			ppidllist[i] = ILCombine(ppidllist[i], pidl);
			teCoTaskMemFree(pidl);
		}
		teILFreeClear(&ppidllist[0]);
#else
		return FALSE;
#endif
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
		teCoTaskMemFree(pidl);
	}

	if (nCommon) {
		for (int i = nCount; i > 0; i--) {
			LPITEMIDLIST pidl = ppidllist[i];
			for (int j = nCommon; j-- > 0;) {
				pidl = ILGetNext(pidl);
			}
			teILCloneReplace(&ppidllist[i], pidl);
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
			try {
				CIDA *pIDA = (CIDA *)GlobalLock(Medium.hGlobal);
				if (pIDA) {
					*pnCount = pIDA->cidl;
					ppidllist = new LPITEMIDLIST[*pnCount + 1];
					for (int i = *pnCount; i >= 0; i--) {
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
			GlobalUnlock(Medium.hGlobal);
			ReleaseStgMedium(&Medium);
			if (ppidllist) {
				return ppidllist;
			}
		}
		if (pDataObj->GetData(&HDROPFormat, &Medium) == S_OK) {
			*pnCount = DragQueryFile((HDROP)Medium.hGlobal, (UINT)-1, NULL, 0);
			ppidllist = new LPITEMIDLIST[*pnCount + 1];
			ppidllist[0] = NULL;
			for (int i = *pnCount; i-- > 0;) {
				BSTR bsPath;
				teDragQueryFile((HDROP)Medium.hGlobal, i, &bsPath);
				ppidllist[i + 1] = teILCreateFromPath(bsPath);
				SysFreeString(bsPath);
			}
			ReleaseStgMedium(&Medium);
		}
	}
	return ppidllist;
}

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

LPITEMIDLIST teILCreateResultsXP(LPITEMIDLIST pidl)
{
	LPITEMIDLIST pidl2 = NULL;
	LPCITEMIDLIST pidlLast;
	IShellFolder *pSF;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlLast)) {
		SFGAOF sfAttr = SFGAO_FILESYSTEM | SFGAO_FILESYSANCESTOR | SFGAO_STORAGE | SFGAO_STREAM;
		if (SUCCEEDED(pSF->GetAttributesOf(1, &pidlLast, &sfAttr)) &&
			(sfAttr & (SFGAO_FILESYSTEM | SFGAO_FILESYSANCESTOR | SFGAO_STORAGE | SFGAO_STREAM))) {
			UINT uSize = ILGetSize(pidl) + 28;
			pidl2 = (LPITEMIDLIST)::CoTaskMemAlloc(uSize + sizeof(USHORT));
			::ZeroMemory(pidl2, uSize + sizeof(USHORT));

			UINT uSize2 = ILGetSize(pidlLast);
			::CopyMemory(pidl2, pidlLast, uSize2);
			*(PUSHORT)pidl2 = uSize - 2;
			UINT uSize3 = uSize - uSize2 - 28;
			PBYTE p = (PBYTE)pidl2;
			*(PUSHORT)&p[uSize2 - 2] = uSize3 + 28;
			*(PDWORD)&p[uSize2 + 2] = 0xbeef0005;
			::CopyMemory(&p[uSize2 + 22], pidl, uSize3);
			*(PUSHORT)&p[uSize - 4] = uSize2 - 2;
			CLSID clsid;
			IPersist *pPersist;
			if SUCCEEDED(pSF->QueryInterface(IID_PPV_ARGS(&pPersist))) {
				if SUCCEEDED(pPersist->GetClassID(&clsid)) {
					if (IsEqualCLSID(clsid, CLSID_ShellFSFolder)) {
						*(PUSHORT)&p[uSize2 + 24 + uSize3] = *(PUSHORT)&p[uSize2 - 4];
					}
				}
				pPersist->Release();
			}
			STRRET strret;
			if SUCCEEDED(pSF->GetDisplayNameOf(pidl2, SHGDN_NORMAL, &strret)) {
				if (strret.uType == STRRET_WSTR) {
					::CoTaskMemFree(strret.pOleStr);
				}
			}
			else {
				teILFreeClear(&pidl2);
			}
		}
		pSF->Release();
	}
	return pidl2;
}

HRESULT STDAPICALLTYPE teSHParseDisplayName2000(LPCWSTR pszName, IBindCtx *pbc, PIDLIST_ABSOLUTE *ppidl, SFGAOF sfgaoIn, SFGAOF *psfgaoOut)
{
	*ppidl = ILCreateFromPath(pszName);
	return *ppidl ? S_OK : E_FAIL;
}

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

HRESULT STDAPICALLTYPE tePSPropertyKeyFromString(__in LPCWSTR pszString,  __out PROPERTYKEY *pkey)
{
	HRESULT hr = E_NOTIMPL;
	ULONG chEaten = 0;
	IPropertyUI *pPUI;
	if SUCCEEDED(CoCreateInstance(CLSID_PropertiesUI, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&pPUI))) {
		hr = pPUI->ParsePropertyName(pszString, &pkey->fmtid, &pkey->pid, &chEaten);
		pPUI->Release();
	}
	return hr;
}

#endif

HMODULE teCreateInstance(LPVARIANTARG pvDllFile, LPVARIANTARG pvClass, REFIID riid, PVOID *ppvObj)
{
	CLSID clsid;
	VARIANT v;
	teVariantChangeType(&v, pvClass, VT_BSTR);
	CLSIDFromString(v.bstrVal, &clsid);
	VariantClear(&v);
	*ppvObj = NULL;
	HMODULE hDll = NULL;
	if FAILED(CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, riid, ppvObj)) {
		teVariantChangeType(&v, pvDllFile, VT_BSTR);
		hDll = LoadLibrary(v.bstrVal);
		VariantClear(&v);
		if (hDll) {
			LPFNDllGetClassObject lpfnDllGetClassObject = (LPFNDllGetClassObject)GetProcAddress(hDll, "DllGetClassObject");
			if (lpfnDllGetClassObject) {
				IClassFactory *pCF;
				if (lpfnDllGetClassObject(clsid, IID_PPV_ARGS(&pCF)) == S_OK) {
					pCF->CreateInstance(NULL, riid, ppvObj);
					pCF->Release();
				}
			}
		}
	}
	return hDll;
}

BOOL GetIDListFromObject(IUnknown *punk, LPITEMIDLIST *ppidl)
{
	*ppidl = NULL;
	if (!punk) {
		return FALSE;
	}
	if SUCCEEDED(lpfnSHGetIDListFromObject(punk, ppidl)) {
		return TRUE;
	}
	IServiceProvider *pSP;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pSP))) {
		IShellBrowser *pSB;
		if SUCCEEDED(pSP->QueryService(SID_SShellBrowser, IID_PPV_ARGS(&pSB))) {
			IShellView *pSV;
			if SUCCEEDED(pSB->QueryActiveShellView(&pSV)) {
				if SUCCEEDED(lpfnSHGetIDListFromObject(pSV, ppidl)) {
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
						delete [] pidllist;
						pDataObj->Release();
					}
#endif
				}
				pSV->Release();
			}
			pSB->Release();
		}
		pSP->Release();
		if (*ppidl != NULL) {
			return TRUE;
		}
	}
	return FALSE;
}

BOOL GetIDListFromObjectEx(IUnknown *punk, LPITEMIDLIST *ppidl)
{
	if (!GetIDListFromObject(punk, ppidl)) {
		*ppidl = ::ILClone(g_pidls[CSIDL_DESKTOP]);
	}
	return true;
}

BOOL teILIsEqualNet(BSTR bs1, IUnknown *punk, BOOL *pbResult)
{
	if (PathIsNetworkPath(bs1)) {
		LPITEMIDLIST pidl;
		if (GetIDListFromObject(punk, &pidl)) {
			BSTR bs2;
			if SUCCEEDED(GetDisplayNameFromPidl(&bs2, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
				*pbResult = teStrSameIFree(bs2, bs1);
			}
			teCoTaskMemFree(pidl);
		}
		return TRUE;
	}
	return FALSE;
}

BOOL teILIsEqual(IUnknown *punk1, IUnknown *punk2)
{
	BOOL bResult = FALSE;
	BSTR bs1 = NULL;
	BSTR bs2 = NULL;
	teGetBSTRFromFolderItem(&bs2, punk2);
	if (teGetBSTRFromFolderItem(&bs1, punk1)) {
		if (bs2) {
			return lstrcmpi(bs1, bs2) == 0;
		}
		if (teILIsEqualNet(bs1, punk2, &bResult)) {
			return bResult;
		}
	}
	if (bs2 && teILIsEqualNet(bs2, punk1, &bResult)) {
		return bResult;
	}
	if (punk1 && punk2) {
		LPITEMIDLIST pidl1, pidl2;
		if (GetIDListFromObject(punk1, &pidl1)) {
			if (GetIDListFromObject(punk2, &pidl2)) {
				bResult = ::ILIsEqual(pidl1, pidl2);
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
	IShellDispatch *psha;
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Shell.Application", &clsid);
	if SUCCEEDED(hr) {
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&psha));
		if SUCCEEDED(hr) {
			hr = psha->NameSpace(v, ppsdf);
		}
		psha->Release();
	}
	VariantClear(&v);
	return hr;
}

BOOL GetFolderItemFromPidl2(FolderItem **ppid, LPITEMIDLIST pidl)
{
	BOOL Result;
	Result = FALSE;
	*ppid = NULL;
	IPersistFolder *pPF = NULL;
	if SUCCEEDED(CoCreateInstance(CLSID_FolderItem, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pPF))) {
		if SUCCEEDED(pPF->Initialize(pidl)) {
			Result = SUCCEEDED(pPF->QueryInterface(IID_PPV_ARGS(ppid)));
		}
		pPF->Release();
		return Result;
	}
#ifdef _W2000
	Folder *pFolder = NULL;
	if (GetFolderObjFromPidl(pidl, &pFolder) == S_OK) {
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
	if (ILRemoveLastID(pidlParent) && GetFolderObjFromPidl(pidlParent, &pFolder) == S_OK) {
		BSTR bs;
		if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORPARSING | SHGDN_INFOLDER)) {
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

BOOL GetFolderItemFromPidl(FolderItem **ppid, LPITEMIDLIST pidl)
{
	if (pidl && GetFolderItemFromPidl2(ppid, pidl)) {
		return TRUE;
	}
	BOOL Result;
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
	v.vt = VT_I4;
	v.lVal = nIndex;
	if (pFolderItems->Item(v, ppFolderItem) == S_OK) {
		return true;
	}
	GetFolderItemFromPidl(ppFolderItem, g_pidls[CSIDL_DESKTOP]);
	return true;
}

/*//
VOID GetFolderItemsFromPCZZSTR(CteFolderItems **ppFolderItems, LPCWSTR pszPath)
{
	*ppFolderItems = new CteFolderItems(NULL, NULL, true);
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
					teCoTaskMemFree(pidl);
				}
			}
			if (pSF) {
				pSF->GetUIObjectOf(g_hwndMain, lCount, pidlList, IID_IDataObject, NULL, (LPVOID*)ppDataObj);
				pSF->Release();
			}
			while (--lCount >= 0) {
				teCoTaskMemFree((LPVOID)pidlList[lCount]);
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
		teCoTaskMemFree(pidl);
	}
	return *ppDataObj != NULL;
}

HRESULT GetFolderItemFromShellItem(FolderItem **ppid, IShellItem *psi)
{
	if (psi) {
		LPITEMIDLIST pidl;
		if (GetIDListFromObject(psi, &pidl)) {
			GetFolderItemFromPidl(ppid, pidl);
			teCoTaskMemFree(pidl);
			return S_OK;
		}
	}
	return E_FAIL;
}

VOID teSetIDList(VARIANT *pv, LPITEMIDLIST pidl)
{
	FolderItem *pFI;
	if (GetFolderItemFromPidl(&pFI, pidl)) {
		teSetObjectRelease(pv, pFI);
	}
}

BOOL GetVariantFromShellItem(VARIANT *pv, IShellItem *psi)
{
	FolderItem *pid;
	if SUCCEEDED(GetFolderItemFromShellItem(&pid, psi)) {
		return teSetObjectRelease(pv, pid);
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
		}//end_switch
	}
}

int* SortTEMethod(TEmethod *method, int nCount)
{
	int *pi = new int[nCount];

	for (int j = 0; j < nCount; j++) {
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

int teBSearch(TEmethod *method, int nSize, int* pMap, LPOLESTR bs)
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

HRESULT teGetDispId(TEmethod *method, int nCount, int* pMap, LPOLESTR bs, DISPID *rgDispId, BOOL bNum)
{
	if (pMap) {
		int nIndex = teBSearch(method, nCount, pMap, bs);
		if (nIndex >= 0) {
			*rgDispId = method[pMap[nIndex]].id;
			return S_OK;
		}
	}
	else {
		for (int i = 0; method[i].name; i++) {
			if (lstrcmpi(bs, method[i].name) == 0) {
				*rgDispId = method[i].id;
				return S_OK;
			}
		}
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
	teGetDispId(structSize, _countof(structSize), g_maps[MAP_SS], bs, &nSize, false);
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

int GetOffsetOfStruct(BSTR bs)
{
	return lstrcmpi(bs, L"EncoderParameters") ? 0 : sizeof(UINT);
}

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
			if (PathMatchSpec(pv->bstrVal, L"*%")) {
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
	}
	else {
		return GetIntFromVariant(pv);
	}
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
	}//end_switch
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
	char *pc = GetpcFromVariant(pv);
	if (pc) {
		IDispatch *pdisp;
		if (GetDispatch(pv, &pdisp)) {
			VARIANT v;
			if (teGetProperty(pdisp, L"Size", &v) == S_OK) {
				*ppc = (UCHAR *)pc;
				*pnLen = GetIntFromVariantClear(&v);
			}
			pdisp->Release();
			return;
		}
	}
	*ppc = NULL;
	*pnLen = 0;
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
				if (g_pWebBrowser) {
					pas->AddNamedItem(L"window", SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE | SCRIPTITEM_GLOBALMEMBERS);
				}
			}
			VARIANT v;
			VariantInit(&v);
			hr = pasp->ParseScriptText(lpScript, NULL, NULL, NULL, 0, 1, SCRIPTTEXT_ISPERSISTENT | SCRIPTTEXT_ISVISIBLE, &v, NULL);
			if (hr == S_OK) {
				pas->SetScriptState(SCRIPTSTATE_CONNECTED);
				if (ppdisp) {
					hr = E_FAIL;
					IDispatch *pdisp;
					if (pas->GetScriptDispatch(NULL, &pdisp) == S_OK) {
						CteDispatch *odisp = new CteDispatch(pdisp, 2);
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
	int npt;
	npt = GetIntFromVariant(pv);
	ppt->x = GET_X_LPARAM(npt);
	ppt->y = GET_Y_LPARAM(npt);
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
	return SUCCEEDED(DoFunc1(nFunc, pObj, &vResult)) ? GetIntFromVariant(&vResult) : hr;
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
	if (GetIDListFromObject(pid, &pidl)) {
		bResult = ::ILRemoveLastID(pidl);
#ifndef _WIN64
		Wow64ControlPanel(&pidl, g_pidls[CSIDL_DESKTOP]);
#endif
		if (!bResult) {
			BSTR bs;
			if SUCCEEDED(pid->get_Path(&bs)) {
				if (teIsFileSystem(bs) && !teIsDesktopPath(bs)) {
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
		GetFolderItemFromPidl(ppid, pidl);
		teCoTaskMemFree(pidl);
	}
	return bResult;
}

VOID GetNewArray(IDispatch **ppArray)
{
	VARIANT v;
	VariantInit(&v);
	if (teExecMethod(g_pJS, L"_a", &v, 0, NULL) == S_OK) {
		GetDispatch(&v, ppArray);
		VariantClear(&v);
	}
}

VOID teArrayPushVariantClear(IDispatch *pdisp, VARIANT *pv)
{
	teExecMethod(pdisp, L"push", NULL, 1, pv);
}

VOID teArrayPush(IDispatch *pdisp, PVOID pObj)
{
	VARIANT *pv = GetNewVARIANT(1);
	teSetObject(pv, pObj);
	teArrayPushVariantClear(pdisp, pv);
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
		pv[0].lVal = iPart;
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
		pstPt->SetPoint(pt.x, pt.y);
		teSetObject(&pv[1], pstPt);
		pstPt->Release();
		pv[0].lVal = flags;
		pv[0].vt = VT_I4;
		Invoke4(g_pOnFunc[TE_OnHitTest], &vResult, 3, pv);
		return GetLLFromVariant(&vResult);
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
				teSetObjectRelease(&pv[2], new CteMemory(sizeof(int), (char *)pgrfKeyState, VT_NULL, 1, L"DWORD"));
				pstPt = new CteMemory(2 * sizeof(int), NULL, VT_NULL, 1, L"POINT");
				pstPt->SetPoint(pt.x, pt.y);
				teSetObjectRelease(&pv[1], pstPt);
				teSetObjectRelease(&pv[0], new CteMemory(sizeof(int), (char *)pdwEffect, VT_NULL, 1, L"DWORD"));
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

HRESULT teDoCommand(PVOID pObj, HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	if (g_pOnFunc[TE_OnCommand]) {
		MSG msg1;
		msg1.hwnd = hwnd;
		msg1.message = msg;
		msg1.wParam = wParam;
		msg1.lParam = lParam;
		return MessageSub(TE_OnCommand, pObj, &msg1, S_FALSE);
	}
	return S_FALSE;
}

HRESULT MessageSubPt(int nFunc, PVOID pObj, MSG *pMsg)
{
	VARIANT vResult;
	if (g_pOnFunc[nFunc]) {
		VARIANTARG *pv = GetNewVARIANT(5);
		teSetObject(&pv[4], pObj);
		SetVariantLL(&pv[3], (LONGLONG)pMsg->hwnd);
		pv[2].lVal = (LONG)pMsg->message;
		pv[2].vt = VT_I4;
		pv[1].lVal = (LONG)pMsg->wParam;
		pv[1].vt = VT_I4;
		CteMemory *pstPt = new CteMemory(2 * sizeof(int), NULL, VT_NULL, 1, L"POINT");
		pstPt->SetPoint(pMsg->pt.x, pMsg->pt.y);
		teSetObjectRelease(&pv[0], pstPt);

		if SUCCEEDED(Invoke4(g_pOnFunc[nFunc], &vResult, 5, pv)) {
			return GetIntFromVariantClear(&vResult);
		}
	}
	return S_FALSE;
}

int GetShellBrowser2(CteShellBrowser *pSB)
{
	for (int i = MAX_FV; i--;) {
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
	for (int i = MAX_TC; i--;) {
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
	CteShellBrowser *pSB = SBfromhwnd(hwnd);
	if (pSB) {
		return pSB->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	CteTabs *pTC = TCfromhwnd(hwnd);
	if (pTC) {
		return pTC->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	CteTreeView *pTV = TVfromhwnd(hwnd, false);
	if (pTV) {
		return pTV->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	if (g_hwndMain == hwnd && g_pTE) {
		return g_pTE->QueryInterface(IID_PPV_ARGS(ppdisp));
	}
	if (g_pWebBrowser && IsChild(g_pWebBrowser->get_HWND(), hwnd)) {
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
	static UINT dwDoubleTime;

	HRESULT hrResult = S_FALSE;
	IDispatch *pdisp = NULL;
	if (nCode >= 0 && g_x == MAXINT) {
		if (nCode == HC_ACTION) {
			if (wParam != WM_MOUSEWHEEL) {
				if (g_pOnFunc[TE_OnMouseMessage]) {
					MOUSEHOOKSTRUCTEX *pMHS = (MOUSEHOOKSTRUCTEX *)lParam;
					if (ControlFromhwnd(&pdisp, pMHS->hwnd) == S_OK) {
						try {
							if (InterlockedIncrement(&g_nProcMouse) == 1) {
								CteTreeView *pTV = NULL;
								TVHITTESTINFO ht;
								pTV = TVfromhwnd(pMHS->hwnd, false);
								if (pTV) {
									if (wParam == WM_LBUTTONUP || wParam == WM_LBUTTONDBLCLK ||
										wParam == WM_RBUTTONUP || wParam == WM_RBUTTONDBLCLK ||
										wParam == WM_MBUTTONUP || wParam == WM_MBUTTONDBLCLK ||
										wParam == WM_XBUTTONUP || wParam == WM_XBUTTONDBLCLK) {
										ht.pt.x = pMHS->pt.x;
										ht.pt.y = pMHS->pt.y;
										ht.flags = TVHT_ONITEM;
										ScreenToClient(pMHS->hwnd, &ht.pt);
										TreeView_HitTest(pMHS->hwnd, &ht);
										if (ht.hItem && ht.flags & TVHT_ONITEM) {
											SetFocus(pMHS->hwnd);
											TreeView_SelectItem(pMHS->hwnd, ht.hItem);
										}
									}
								}
								MSG msg;
								msg.hwnd = pMHS->hwnd;
								msg.message = (LONG)wParam;
								if (msg.message == WM_LBUTTONDOWN) {
									WCHAR szClass[MAX_CLASS_NAME];
									GetClassName(msg.hwnd, szClass, MAX_CLASS_NAME);
									if (lstrcmp(szClass, L"DirectUIHWND") == 0) {
										DWORD dwTick = GetTickCount();
										if (dwTick - dwDoubleTime < GetDoubleClickTime()) {
											msg.message = WM_LBUTTONDBLCLK;
										}
										else {
											dwDoubleTime = dwTick;
										}
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
											}
											else if (wParam == WM_MBUTTONUP) {
												nStyle = NSTCECT_MBUTTON;
											}
											hrResult = pTV->OnItemClick(NULL, ht.flags, nStyle);
										}
									}
								}
#endif
							}
						}
						catch(...) {}
						::InterlockedDecrement(&g_nProcMouse);
						pdisp->Release();
					}
				}
			}
		}
	}
	return (hrResult != S_OK) ? CallNextHookEx(g_hMouseHook, nCode, wParam, lParam) : TRUE;
}

LRESULT CALLBACK TETVProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	BOOL Handled = FALSE;
	try {
		CteTreeView *pTV = (CteTreeView *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (msg == WM_NOTIFY) {
			LPNMTVCUSTOMDRAW lptvcd = (LPNMTVCUSTOMDRAW)lParam;
/*/// Color
			if (lptvcd->nmcd.hdr.code == NM_CUSTOMDRAW) {
				if (lptvcd->nmcd.dwDrawStage == CDDS_PREPAINT) {
					return CDRF_NOTIFYITEMDRAW;
				}
				if (lptvcd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
					if (!(lptvcd->nmcd.uItemState & CDIS_SELECTED)) {
						lptvcd->clrText = RGB(0, 0, 255);
					}
					return CDRF_DODEFAULT;
				}
			}
*///
#ifdef _2000XP
			if (pTV->m_pShellNameSpace && lptvcd->nmcd.hdr.code == NM_RCLICK) {
				try {
					if (InterlockedIncrement(&g_nProcTV) < 5) {
						MSG msg1;
						msg1.hwnd = hwnd;
						msg1.message = WM_CONTEXTMENU;
						msg1.wParam = (WPARAM)hwnd;
						GetCursorPos(&msg1.pt);
						if (MessageSubPt(TE_OnShowContextMenu, pTV, &msg1) == S_OK) {
							Handled = TRUE;
						}
					}
				}
				catch(...) {}
				::InterlockedDecrement(&g_nProcTV);
			}
#endif
		}
		if (Handled) {
			return 1;
		}
		return CallWindowProc(pTV->m_DefProc, hwnd, msg, wParam, lParam);
	}
	catch (...) {
		g_nException = 0;
	}
	return 0;
}

LRESULT CALLBACK TETVProc2(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	try {
		CteTreeView *pTV = (CteTreeView *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (pTV->m_bMain) {
			if (msg == WM_ENTERMENULOOP) {
				pTV->m_bMain = false;
			}
		}
		else {
			if (msg == WM_EXITMENULOOP) {
				pTV->m_bMain = true;
			}
			else if (msg >= TV_FIRST && msg < TV_FIRST + 0x100) {
				return 0;
			}
		}
		if (msg == WM_COMMAND) {
			LRESULT Result = S_FALSE;
			try {
				if (InterlockedIncrement(&g_nProcTV) == 1) {
					Result = teDoCommand(pTV, hwnd, msg, wParam, lParam);
				}
			}
			catch(...) {}
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
		return CallWindowProc(pTV->m_DefProc2, hwnd, msg, wParam, lParam);
	}
	catch (...) {
		g_nException = 0;
	}
	return 0;
}

VOID CALLBACK teNavigateCompleteProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);
	CteShellBrowser *pSB = (CteShellBrowser *)idEvent;
	CteFolderItem *pid;
	if (pSB->m_pFolderItem && SUCCEEDED(pSB->m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid))) {
		if (pid->m_pidlFocused) {
			pSB->SelectItemEx(&pid->m_pidlFocused, 
				pid->m_nSelected != 1 ? SVSI_FOCUSED | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS :
				SVSI_SELECT | SVSI_FOCUSED | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS);
		}
		else {
			IFolderView *pFV;
			if (pSB->m_pShellView && SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV)))) {
				pFV->SelectItem(0, SVSI_FOCUSED | SVSI_ENSUREVISIBLE);
				pFV->Release();
			}
		}
		pid->Release();
	}
}

LRESULT CALLBACK TELVProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hwndEdit = NULL;
	static FolderItem *pidEdit = NULL;

	try {
		CteShellBrowser *pSB = (CteShellBrowser *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		LRESULT lResult = S_FALSE;
/*/// Color
		if (msg == WM_NOTIFY) {
			LPNMLVCUSTOMDRAW lplvcd = (LPNMLVCUSTOMDRAW)lParam;
			if (lplvcd->nmcd.hdr.code == NM_CUSTOMDRAW) {
				if (lplvcd->nmcd.dwDrawStage == CDDS_PREPAINT) {
					return CDRF_NOTIFYITEMDRAW;
				}
				if (lplvcd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
					if (lplvcd->nmcd.uItemState) {
						lplvcd->clrText = RGB(0, 0, 255);
						lplvcd->clrTextBk = CLR_NONE;
					}
					return CDRF_DODEFAULT;
				}
			}
		}
//*/
		if (msg == WM_CONTEXTMENU || msg == SB_SETTEXT || msg == WM_COMMAND || msg == WM_COPYDATA) {
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
									ClientToScreen(hwndLV, &msg1.pt);
									int nSelected = ListView_GetNextItem(hwndLV, -1, LVNI_ALL | LVNI_SELECTED);
									if (nSelected >= 0) {
										RECT rc;
										ListView_GetItemRect(hwndLV, nSelected, &rc, LVIR_ICON);
										OffsetRect(&rc, msg1.pt.x, msg1.pt.y);
										msg1.pt.x = (rc.left + rc.right) / 2;
										msg1.pt.y = (rc.top + rc.bottom) / 2;
									}
								}
								else {
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
							if (wParam == 0x7103) {//Refresh
								if (ILIsEqual(pSB->m_pidl, g_pidlResultsFolder)) {
									pSB->BrowseObject(NULL, SBSP_RELATIVE | SBSP_SAMEBROWSER);
								}
							}
							break;
						case WM_COPYDATA:
							if (g_pOnFunc[TE_OnSystemMessage]) {
								MSG msg1;
								msg1.hwnd = hwnd;
								msg1.message = msg;
								msg1.wParam = wParam;
								msg1.lParam = lParam;
								lResult = MessageSub(TE_OnSystemMessage, pSB, &msg1, S_FALSE);
							}
							break;
					}//end_switch
				}
			}
			catch(...) {}
			::InterlockedDecrement(&g_nProcFV);
		}
		if (msg == WM_WINDOWPOSCHANGED) {
			pSB->SetFolderFlags();
		}
		return (lResult && hwnd == pSB->m_hwndDV) ? CallWindowProc((WNDPROC)pSB->m_DefProc, hwnd, msg, wParam, lParam) : 0;
	}
	catch (...) {
		g_nException = 0;
	}
	return 0;
}

LRESULT CALLBACK TELVProc2(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	try {
		CteShellBrowser *pSB = (CteShellBrowser *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
//		LRESULT lResult = S_FALSE;
		if (msg == WM_HSCROLL || msg == WM_VSCROLL) {
			g_dwTickScroll = GetTickCount();
		}
		return (/*lResult &&*/ hwnd == pSB->m_hwndLV) ? CallWindowProc((WNDPROC)pSB->m_DefProc2, hwnd, msg, wParam, lParam) : 0;
	}
	catch (...) {
		g_nException = 0;
	}
	return 0;
}

VOID teSetTreeWidth(CteShellBrowser	*pSB, HWND hwnd, LPARAM lParam)
{
	int nWidth = pSB->m_param[SB_TreeWidth] + GET_X_LPARAM(lParam) - g_x;
	RECT rc;
	GetWindowRect(hwnd, &rc);
	int nMax = rc.right - rc.left - 1;
	if (nWidth > nMax) {
		nWidth = nMax;
	}
	else if (nWidth < 3) {
		nWidth = 3;
	}
	else {
		g_x =GET_X_LPARAM(lParam);
	}
	if (pSB->m_param[SB_TreeWidth] != nWidth) {
		pSB->m_param[SB_TreeWidth] = nWidth;
		ArrangeWindow();
	}
}

VOID teSetTabWidth(CteTabs *pTC, LPARAM lParam)
{
	int nWidth = pTC->m_param[TC_TabWidth];
	RECT rc;
	GetWindowRect(pTC->m_hwndStatic, &rc);
	int nMax = rc.right - rc.left - 1;
	int nDiff = GET_X_LPARAM(lParam) - g_x;
	nWidth += (pTC->m_param[TE_Align] == 4) ? nDiff : -nDiff;
	if (nWidth > nMax) {
		nWidth = nMax;
	}
	else if (nWidth < 1) {
		nWidth = 1;
	}
	else {
		g_x = GET_X_LPARAM(lParam);
	}
	if (pTC->m_param[TC_TabWidth] != nWidth) {
		pTC->m_param[TC_TabWidth] = nWidth;
		ArrangeWindow();
	}
}

LRESULT CALLBACK TESTProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT Result = 1;
	try {
		CteTabs *pTC = (CteTabs *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		CteShellBrowser	*pSB = pTC->GetShellBrowser(pTC->m_nIndex);
		if (pSB && pSB->m_pTV && pSB->m_pidl && pSB->m_hwnd) {
			switch (msg) {
				case WM_MOUSEMOVE:
					if (g_x == MAXINT) {
						SetCursor(LoadCursor(NULL, IDC_SIZEWE));
					}
					else {
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
		return Result ? CallWindowProc(pTC->m_DefSTProc, hwnd, msg, wParam, lParam) : 0;
	}
	catch (...) {
		g_nException = 0;
	}
	return 0;
}

LRESULT TETCProc2(CteTabs *pTC, HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT Result = 1;
	switch (msg) {
		case WM_LBUTTONDOWN:
			if ((pTC->m_param[TE_Align] == 4 && WORD(lParam) >= pTC->m_param[TC_TabWidth] - pTC->m_nScrollWidth - 2 && WORD(lParam) < pTC->m_param[TC_TabWidth]) ||
				(pTC->m_param[TE_Align] == 5 && WORD(lParam) < 2)) {
				SetCapture(hwnd);
				g_x = GET_X_LPARAM(lParam);
				SetCursor(LoadCursor(NULL, IDC_SIZEWE));
				Result = 0;
			}
		case WM_SETFOCUS:
		case WM_RBUTTONDOWN:
			CheckChangeTabTC(hwnd, true);
			break;
		case WM_MOUSEMOVE:
			if (g_x == MAXINT) {
				if ((pTC->m_param[TE_Align] == 4 && WORD(lParam) >= pTC->m_param[TC_TabWidth] - pTC->m_nScrollWidth - 2 && WORD(lParam) < pTC->m_param[TC_TabWidth]) ||
					(pTC->m_param[TE_Align] == 5 && WORD(lParam) < 2)) {
					SetCursor(LoadCursor(NULL, IDC_SIZEWE));
				}
			}
			else {
				if (pTC->m_param[TE_Align] == 4) {
					teSetTabWidth(pTC, lParam);
				}
				Result = 0;
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
	return Result;
}

LRESULT CALLBACK TEBTProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT Result = 1;
	try {
		CteTabs *pTC = (CteTabs *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		switch (msg) {
			case WM_NOTIFY:
				switch (((LPNMHDR)lParam)->code) {
					case TCN_SELCHANGE:
						if (g_pTabs->m_hwnd != pTC->m_hwnd && pTC->m_bVisible) {
							g_pTabs = pTC;
						}
						pTC->TabChanged(true);
						Result = 0;
						break;
					case TCN_SELCHANGING:
						Result = DoFunc(TE_OnSelectionChanging, pTC, 0);
						break;
					case TTN_GETDISPINFO:
						if (g_pOnFunc[TE_OnToolTip]) {
							VARIANT vResult;
							VariantInit(&vResult);
							VARIANTARG *pv = GetNewVARIANT(2);
							VariantInit(&pv[1]);
							teSetObject(&pv[1], pTC);
							VariantInit(&pv[0]);
							SetVariantLL(&pv[0], ((LPNMHDR)lParam)->idFrom);
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
				Result = 0;
				break;
			case WM_VSCROLL:
				switch(LOWORD(wParam)){
					case SB_LINEUP:
						pTC->m_si.nPos -= 16;
						break;
					case SB_LINEDOWN:
						pTC->m_si.nPos += 16; break;
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
		if (Result) {
			Result = TETCProc2(pTC, hwnd, msg, wParam, lParam);
		}
		return Result ? CallWindowProc(pTC->m_DefBTProc, hwnd, msg, wParam, lParam) : 0;
	}
	catch (...) {
		g_nException = 0;
	}
	return 0;
}

LRESULT CALLBACK TETCProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT Result = 1;
	int nIndex, nCount;
	try {
		CteTabs *pTC = (CteTabs *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		switch (msg) {
			case TCM_DELETEITEM:
				if (lParam) {
					lParam = 0;
				}
				else {
					CteShellBrowser *pSB = NULL;
					pSB = pTC->GetShellBrowser((int)wParam);
					CallWindowProc(pTC->m_DefTCProc, hwnd, msg, wParam, lParam);
					Result = 0;
					if (pSB) {
						pSB->Close(true);
					}
					nIndex = TabCtrl_GetCurSel(hwnd);
					if (nIndex >= 0) {
						pTC->m_nIndex = nIndex;
					}
					else {
						nCount = TabCtrl_GetItemCount(hwnd);
						nIndex = pTC->m_nIndex;
						if ((int)wParam > nIndex) {
							nIndex--;
						}
						if (nIndex >= nCount) {
							nIndex = nCount - 1;
						}
						if (nIndex < 0) {
							nIndex = 0;
						}
						pTC->m_nIndex = -1;
						CallWindowProc(pTC->m_DefTCProc, hwnd, TCM_SETCURSEL, nIndex, 0);
					}
					pTC->TabChanged(true);
					ArrangeWindow();
				}
				break;
			case TCM_SETCURSEL:
				if (wParam != (UINT_PTR)TabCtrl_GetCurSel(hwnd)) {
					CallWindowProc(pTC->m_DefTCProc, hwnd, msg, wParam, lParam);
					Result = 0;
					pTC->TabChanged(true);
					if (g_pTE->m_param[TE_Tab] && pTC->m_param[TE_Align] == 1) {
						if (pTC->m_param[TE_Flags] & TCS_SCROLLOPPOSITE) {
							ArrangeWindow();
						}
					}
				}
				break;
			case TCM_SETITEM:
				if (pTC->m_param[TE_Flags] & TCS_FIXEDWIDTH) {
					CallWindowProc(pTC->m_DefTCProc, hwnd, msg, wParam, lParam);
					Result = 0;
					CallWindowProc(pTC->m_DefTCProc, pTC->m_hwnd, TCM_SETITEMSIZE, 0, MAKELPARAM(pTC->m_param[TC_TabWidth], pTC->m_param[TC_TabHeight] + 1));
					CallWindowProc(pTC->m_DefTCProc, pTC->m_hwnd, TCM_SETITEMSIZE, 0, pTC->m_dwSize);
				}
				break;
		}//end_switch
		if (Result) {
			Result = TETCProc2(pTC, hwnd, msg, wParam, lParam);
		}
		return Result ? CallWindowProc(pTC->m_DefTCProc, hwnd, msg, wParam, lParam) : 0;
	}
	catch (...) {
		g_nException = 0;
	}
	return 0;
}

HRESULT MessageProc(MSG *pMsg)
{
	HRESULT hrResult = S_FALSE;
	IDispatch *pdisp = NULL;
	IShellBrowser *pSB = NULL;
	IShellView *pSV = NULL;
	IWebBrowser2 *pWB = NULL;
	IOleInPlaceActiveObject *pActiveObject = NULL;

	if (pMsg->message == WM_MOUSEWHEEL) {
		try {
			if (InterlockedIncrement(&g_nProcMouse) < 10) {
				if (ControlFromhwnd(&pdisp, pMsg->hwnd) == S_OK) {
					if (MessageSubPt(TE_OnMouseMessage, pdisp, pMsg) == S_OK) {
						hrResult = S_OK;
					}
					pdisp->Release();
				}
			}
		}
		catch(...) {}
		::InterlockedDecrement(&g_nProcMouse);
	}

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
	g_maps[MAP_TE] = SortTEMethod(methodTE, _countof(methodTE));
	g_maps[MAP_SB] = SortTEMethod(methodSB, _countof(methodSB));
	g_maps[MAP_TC] = SortTEMethod(methodTC, _countof(methodTC));
	g_maps[MAP_TV] = SortTEMethod(methodTV, _countof(methodTV));
	g_maps[MAP_API] = SortTEMethod(methodAPI, _countof(methodAPI));
	g_maps[MAP_Mem] = SortTEMethod(methodMem, _countof(methodMem));
	g_maps[MAP_GB] = SortTEMethod(methodGB, _countof(methodGB));
	g_maps[MAP_FIs] = SortTEMethod(methodFIs, _countof(methodFIs));
	g_maps[MAP_CD] = SortTEMethod(methodCD, _countof(methodCD));
	g_maps[MAP_SS] = SortTEMethod(structSize, _countof(structSize));

#ifdef _WIN64
	g_pidlResultsFolder = ILCreateFromPath(L"shell:::{2965E715-EB66-4719-B53F-1672673BBEFA}");
#else
	DWORDLONG dwlConditionMask = 0;
    OSVERSIONINFOEX osvi;
    ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));
    osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
    osvi.dwMajorVersion = 6;
    VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, VER_GREATER_EQUAL);
	g_bUpperVista = VerifyVersionInfo(&osvi, VER_MAJORVERSION, dwlConditionMask);
    dwlConditionMask = 0;
	osvi.dwMajorVersion = 5;
    osvi.dwMinorVersion = 1;
	VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, VER_EQUAL);
    VER_SET_CONDITION(dwlConditionMask, VER_MINORVERSION, VER_GREATER_EQUAL);
	g_bIsXP = VerifyVersionInfo(&osvi, VER_MAJORVERSION | VER_MINORVERSION, dwlConditionMask);
	g_bIs2000 = !g_bUpperVista && !g_bIsXP;
	g_pidlResultsFolder = ILCreateFromPath(L"::{e17d4fc0-5564-11d1-83f2-00a0c90dc849}");
	if (!g_pidlResultsFolder) {
		g_pidlResultsFolder = ILCreateFromPath(L"shell:::{2965E715-EB66-4719-B53F-1672673BBEFA}");
	}
#endif
	g_pidlLibrary = teILCreateFromPath(L"shell:libraries");
	g_pidlCP = ILClone(g_pidls[CSIDL_CONTROLS]);
	ILRemoveLastID(g_pidlCP);
	if (ILIsEmpty(g_pidlCP) || ILIsEqual(g_pidlCP, g_pidls[CSIDL_DRIVES])) {
		teILCloneReplace(&g_pidlCP, g_pidls[CSIDL_CONTROLS]);
	}
}

VOID Finalize()
{
	try {
		teRevoke();
		for (int i = _countof(g_maps); i--;) {
			delete [] g_maps[i];
		}
		lpfnSHCreateItemFromIDList = NULL;
		while (g_nPidls--) {
			teILFreeClear(&g_pidls[g_nPidls]);
		}
		teCoTaskMemFree(g_pidlResultsFolder);
		teCoTaskMemFree(g_pidlCP);
		teCoTaskMemFree(g_pidlLibrary);
		if (g_hCrypt32) {
			FreeLibrary(g_hCrypt32);
		}
		teSysFreeString(&g_bsCmdLine);
	}
	catch (...) {
	}
	try {
		Gdiplus::GdiplusShutdown(g_Token);
	} catch (...) {}
	try {
		::OleUninitialize();
	}
	catch (...) {
	}
}

HRESULT teExtract(IStorage *pStorage, LPWSTR lpszFolderPath)
{
	STATSTG statstg;
	IEnumSTATSTG *pEnumSTATSTG;
#ifdef _2000XP
	IShellFolder2 *pSF2;
	LPITEMIDLIST pidl;
	ULONG chEaten;
	ULONG dwAttributes;
	SYSTEMTIME SysTime;
	SHCOLUMNID scid;
	scid.fmtid = FMTID_Storage;
	scid.pid = 14;
	VARIANT v;
#endif
	CreateDirectory(lpszFolderPath, NULL);

	HRESULT hr = pStorage->EnumElements(NULL, NULL, NULL, &pEnumSTATSTG);
	if FAILED(hr) {
		return hr;
	}
	while (pEnumSTATSTG->Next(1, &statstg, NULL) == S_OK) {
		BSTR bsPath;
		tePathAppend(&bsPath, lpszFolderPath, statstg.pwcsName);
		if (statstg.type == STGTY_STORAGE) {
			IStorage *pStorageNew;
			pStorage->OpenStorage(statstg.pwcsName, NULL, STGM_READ, 0, 0, &pStorageNew);
			teExtract(pStorageNew, bsPath);
			pStorageNew->Release();
		}
		else if (statstg.type == STGTY_STREAM) {
			IStream *pStream;
			ULONG uRead;
			DWORD dwWriteByte;

			hr = pStorage->OpenStream(statstg.pwcsName, NULL, STGM_READ, NULL, &pStream);
			if SUCCEEDED(hr) {
				LPBYTE lpData = (LPBYTE)CoTaskMemAlloc(SIZE_BUFF);
				if (lpData) {
					HANDLE hFile = CreateFile(bsPath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
					if (hFile) {
						while (SUCCEEDED(pStream->Read(lpData, SIZE_BUFF, &uRead)) && uRead) {
							WriteFile(hFile, lpData, uRead, &dwWriteByte, NULL);
						}
#ifdef _2000XP
						if (statstg.mtime.dwLowDateTime == 0) {
							if SUCCEEDED(pStorage->QueryInterface(IID_PPV_ARGS(&pSF2))) {
								if SUCCEEDED(pSF2->ParseDisplayName(NULL, NULL, statstg.pwcsName, &chEaten, &pidl, &dwAttributes)) {
									VariantInit(&v);
									if SUCCEEDED(pSF2->GetDetailsEx(pidl, &scid, &v) && v.vt == VT_DATE) {
										if (::VariantTimeToSystemTime(v.date, &SysTime)) {
											::SystemTimeToFileTime(&SysTime, &statstg.mtime);
										}
									}
									VariantClear(&v);
									teCoTaskMemFree(pidl);
								}
								pSF2->Release();
							}
						}
#endif
						SetFileTime(hFile, &statstg.ctime, &statstg.atime, &statstg.mtime);
						CloseHandle(hFile);
						hr = S_OK;
					}
					else {
						hr = E_ACCESSDENIED;
					}
					teCoTaskMemFree(lpData);
				}
				else {
					hr = E_OUTOFMEMORY;
				}
				pStream->Release();
			}
		}
		SysFreeString(bsPath);
		teCoTaskMemFree(statstg.pwcsName);
	}
	pEnumSTATSTG->Release();
	return hr;
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
ATOM MyRegisterClass(HINSTANCE hInstance, LPWSTR szClassName)
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
	wcex.lpszClassName	= szClassName;
	wcex.hIconSm		= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_16));

	return RegisterClassEx(&wcex);
}

BOOL GetEncoderClsid(const WCHAR* pszName, CLSID* pClsid, LPWSTR pszMimeType)
{
	UINT num = 0;			// number of image encoders
	UINT size = 0;			// size of the image encoder array in bytes

	ImageCodecInfo* pImageCodecInfo = NULL;

	GetImageEncodersSize(&num, &size);
	if (size) {
		pImageCodecInfo = (ImageCodecInfo*) new char[size];
		if (pImageCodecInfo) {
			GetImageEncoders(num, size, pImageCodecInfo);
			while (num--)
			{
				if (PathMatchSpec(pszName, pImageCodecInfo[num].FilenameExtension) || lstrcmpi(pszName, pImageCodecInfo[num].MimeType) == 0) {
					if (pClsid) {
						*pClsid = pImageCodecInfo[num].Clsid;
					}
					if (pszMimeType) {
						lstrcpyn(pszMimeType, pImageCodecInfo[num].MimeType, 15);
					}
					delete [] pImageCodecInfo;
					return TRUE;
				}
			}
			delete [] pImageCodecInfo;
		}
	}
	return FALSE;
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
	rcTree.right = rc->left + pSB->m_param[SB_TreeWidth] - 2;
	CteTreeView *pTV = pSB->m_pTV;
	if (pTV) {
		MoveWindow(pTV->m_hwnd, rcTree.left, rcTree.top,
			rcTree.right - rcTree.left, rcTree.bottom - rcTree.top,	true);
#ifdef _2000XP
		if (pTV->m_pShellNameSpace) {
			GetClientRect(pTV->m_hwnd, &rcTree);
			IOleInPlaceObject *pOleInPlaceObject;
			pTV->m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pOleInPlaceObject));
			pOleInPlaceObject->SetObjectRects(&rcTree, &rcTree);
			pOleInPlaceObject->Release();
		}
#endif
		rc->left += pSB->m_param[SB_TreeWidth];
	}
}

VOID CALLBACK teTimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	IDispatch *pdisp;
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
			if (g_nReload) {
				teRevoke();
				BSTR bsPath, bsQuoted;
				int i = teGetModuleFileName(NULL, &bsPath);
				bsQuoted = SysAllocStringLen(bsPath, i + 2);
				SysFreeString(bsPath);
				PathQuoteSpaces(bsQuoted);
				ShellExecute(hwnd, NULL, bsQuoted, NULL, NULL, SW_SHOWNORMAL);
				SysFreeString(bsQuoted);
				PostMessage(hwnd, WM_CLOSE, 0, 0);
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
						teCalcClientRect(pTC->m_param, &rc, &rcClient);
						if (g_pOnFunc[TE_OnArrange]) {
							VARIANTARG *pv;
							pv = GetNewVARIANT(2);
							teSetObject(&pv[1], pTC);
							teSetObjectRelease(&pv[0], new CteMemory(sizeof(RECT), (char *)&rc, VT_NULL, 1, L"RECT"));
							Invoke4(g_pOnFunc[TE_OnArrange], NULL, 2, pv);
						}
						if (!pTC->m_bEmpty && pTC->m_bVisible) {
							int nAlign = g_pTE->m_param[TE_Tab] ? pTC->m_param[TE_Align] : 0;
							if (rc.left) {
								rc.left++;
							}
							MoveWindow(pTC->m_hwndStatic, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, TRUE);
							int left = rc.left;
							int top = rc.top;

							OffsetRect(&rc, -left, -top);
							CopyRect(&rcTab, &rc);
							int nCount = TabCtrl_GetItemCount(pTC->m_hwnd);
							int h = rc.bottom;
							int nBottom = rcTab.bottom;
							if (nCount) {
								DWORD dwStyle = pTC->GetStyle();
								if (nAlign == 4 || nAlign == 5) {
									rc.right = pTC->m_param[TC_TabWidth];
								}
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
										rcTab.top = rcTab.bottom;
										rcTab.left = rcTab.right;
										nBottom = rcTab.bottom;
										break;
									case 2:						//top
										CopyRect(&rc, &rcTab);
										rcTab.bottom = h;
										rc.top += h;
										nBottom = h;
										break;
									case 3:						//bottom
										CopyRect(&rc, &rcTab);
										rcTab.top = rcTab.bottom - h;
										rc.bottom -= h;
										nBottom = rcTab.bottom;
										break;
									case 4:						//left
										SetRect(&rc, pTC->m_param[TC_TabWidth], 0, rcTab.right, rcTab.bottom);
										rcTab.right = rc.left;
										nBottom = h;
										break;
									case 5:						//Right
										SetRect(&rc, 0, 0, rcTab.right - pTC->m_param[TC_TabWidth], rcTab.bottom);
										rcTab.left = rc.right;
										nBottom = h;
										break;
								}
								while (--nCount >= 0) {
									if (nCount != pTC->m_nIndex) {
										pSB = pTC->GetShellBrowser(nCount);
										if (pSB && pSB->m_bVisible) {
											pSB->Show(FALSE);
										}
									}
								}
							}
							MoveWindow(pTC->m_hwndButton, rcTab.left, rcTab.top,
								rcTab.right - rcTab.left, rcTab.bottom - rcTab.top, TRUE);
							pTC->m_nScrollWidth = 0;
							if (nAlign == 4 || nAlign == 5) {
								int h2 = h - rcTab.bottom;
								if (h2 <= 0) {
									h2 = 0;
								}
								else {
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
								if (pTC->m_param[TE_Flags] & TCS_FIXEDWIDTH) {
									pTC->SetItemSize();
								}
								rcTab.right -= pTC->m_nScrollWidth;
							}
							else {
								ShowScrollBar(pTC->m_hwndButton, SB_VERT, FALSE);
								if (pTC->m_si.nTrackPos && pTC->m_param[TE_Flags] & TCS_FIXEDWIDTH) {
									pTC->SetItemSize();
								}
								pTC->m_si.nTrackPos = 0;
								pTC->m_si.nPos = 0;
							}
							if (teSetRect(pTC->m_hwnd, 0, -pTC->m_si.nPos, rcTab.right - rcTab.left, (nBottom - rcTab.top) + pTC->m_si.nPos)) {
								ArrangeWindow();
							}
							pSB = pTC->GetShellBrowser(pTC->m_nIndex);
							if (pSB) {
								if (pSB->m_param[SB_TreeAlign] & 2) {
									ArrangeTree(pSB, &rc);
								}
								if (pSB->m_bVisible) {
									teSetParent(pSB->m_pTV->m_hwnd, pSB->m_pTabs->m_hwndStatic);
									ShowWindow(pSB->m_pTV->m_hwnd, (pSB->m_param[SB_TreeAlign] & 2) ? SW_SHOW : SW_HIDE);
								}
								else {
									pSB->Show(TRUE);
								}
								MoveWindow(pSB->m_hwnd, rc.left, rc.top, rc.right - rc.left,
									rc.bottom - rc.top, TRUE);
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
							ShowWindow(pSB->m_hwndLV, SW_SHOW);
						}
					}
				}
				else {
					break;
				}
			}
			break;
		case TET_Unload:
			g_bUnload = FALSE;
			if (g_pUnload) {
				pdisp = g_pUnload;
				GetNewArray(&g_pUnload);
				pdisp->Release();
			}
			break;
		case TET_Status:
			pSB = g_pTabs->GetShellBrowser(g_pTabs->m_nIndex);
			if (pSB) {
				if (g_szStatus[0] == NULL) {
					IFolderView *pFV;
					if (pSB->m_pShellView && SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV)))) {
						int nCount;
						if SUCCEEDED(pFV->ItemCount(SVGIO_SELECTION, &nCount)) {
							UINT uID;
							if (nCount) {
								uID = nCount > 1 ? 38194 : 38195;
							}
							else if SUCCEEDED(pFV->ItemCount(SVGIO_ALLVIEW, &nCount)) {
								uID = nCount > 1 ? 38192 : 38193;
							}
							WCHAR sz[MAX_STATUS + 4];
							if (LoadString(g_hShell32, uID, &sz[4], MAX_STATUS) <= 2) {
								uID = uID < 38194 ? 6466 : 6477;
								if (LoadString(g_hShell32, uID, sz, MAX_STATUS) <= 6) {
									lstrcpy(&sz[6], nCount <= 1 ? L" item" : L" items");
									if (uID == 6477) {
										lstrcat(&sz[6], L" selected");
									}
								}
							}
							sz[4] = '%';
							sz[5] = 'd';
							swprintf_s(g_szStatus, MAX_STATUS, &sz[4], nCount);
						}
						pFV->Release();
					}
				}
				DoStatusText(pSB, g_szStatus, 0);
				g_szStatus[0] = NULL;
			}
			break;
		case TET_Refresh:
			if (GetKeyState(VK_LBUTTON) >= 0) {
				pSB = SBfromhwnd(hwnd);
				if (pSB) {
					pSB->BrowseObject(NULL, SBSP_RELATIVE | SBSP_SAMEBROWSER);
				}
			}
			else {
				SetTimer(hwnd, TET_Refresh, 500, teTimerProc);
			}
			break;
	}//end_switch
}

VOID CALLBACK teTimerProcKillFocus(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);

}

LRESULT CALLBACK HookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LPCWPSTRUCT pcwp;

	if (nCode >= 0) {
		if (nCode == HC_ACTION) {
			if (wParam == NULL) {
				pcwp = (LPCWPSTRUCT)lParam;
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
											if (pSB->m_hwndLV == pcwp->hwnd) {
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
					case WM_COMMAND:
						if (pcwp->message == WM_COMMAND && g_hDialog == pcwp->hwnd) {
							if (LOWORD(pcwp->wParam) == IDOK) {
								g_bDialogOk = TRUE;
							}
						}
					case WM_KILLFOCUS:
						if (pcwp->message == WM_KILLFOCUS && GetFocus() != g_hwndMain) {
							break;
						}
					case WM_NOTIFY:
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
								MessageSub(TE_OnSystemMessage, pdisp, &msg1, S_FALSE);
							}
						}
						break;
				}//end_switch
			}
		}
	}
	return CallNextHookEx(g_hHook, nCode, wParam, lParam);
}

UINT_PTR CALLBACK OFNHookProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LPOFNOTIFY pNotify;
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
						LPITEMIDLIST pidl = (LPITEMIDLIST)CoTaskMemAlloc(nLen);
						SendMessage(hDlg, CDM_GETFOLDERIDLIST, nLen, (LPARAM)pidl);
						BSTR bs;
						GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORPARSING | SHGDN_FORADDRESSBAR);
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
    return FALSE;
}

int APIENTRY _tWinMain(HINSTANCE hInstance,
					HINSTANCE hPrevInstance,
					LPTSTR    lpCmdLine,
					int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	g_nCmdShow = nCmdShow;
	BSTR bsPath, bsLib, bsIndex;
	HINSTANCE hDll;

	hInst = hInstance;

	::OleInitialize(NULL);
	InitCommonControls();

	for (UINT i = 0; i < 256; i++) {
		UINT c = i;
		for (int j = 8; j--;) {
			c = (c & 1) ? (0xedb88320 ^ (c >> 1)) : (c >> 1);
		}
		g_uCrcTable[i] = c;
	}
	teGetModuleFileName(NULL, &bsPath);//Executable Path
	for (int i = 0; bsPath[i]; i++) {
		if (bsPath[i] == '\\') {
			bsPath[i] = '/';
		}
		bsPath[i] = towupper(bsPath[i]);
	}
	int nCrc32 = CalcCrc32((BYTE *)bsPath, lstrlen(bsPath) * sizeof(WCHAR), 0);
	//Command Line
	BOOL bVisible = !PathMatchSpec(lpCmdLine, L"/run *");
	BOOL bNewProcess = PathMatchSpec(lpCmdLine, L"/open *");
	LPWSTR szClass = WINDOW_CLASS2;
	if (bVisible && !bNewProcess) {
		//Multiple Launch
		szClass = WINDOW_CLASS;
		g_hMutex = CreateMutex(NULL, FALSE, bsPath);
		if (WaitForSingleObject(g_hMutex, 0) == WAIT_TIMEOUT) {
			HWND hwndTE = NULL;
			while (hwndTE = FindWindowEx(NULL, hwndTE, szClass, NULL)) {
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
						SysFreeString(bsPath);
						Finalize();
						teSetForegroundWindow(hwndTE);
						return FALSE;
					}
					DWORD dwPid;
					GetWindowThreadProcessId(hwndTE, &dwPid);
					HANDLE hProcess = OpenProcess(PROCESS_TERMINATE, FALSE, dwPid);
					if (hProcess) {
						if (TerminateProcess(hProcess, 0)) {
							WaitForSingleObject(g_hMutex, 0);
						}
						CloseHandle(hProcess);

					}
				}
			}
		}
	}
	SysFreeString(bsPath);

	for (int i = MAX_CSIDL; i--;) {
		SHGetFolderLocation(NULL, i, NULL, NULL, &g_pidls[i]);
		if (!g_pidls[i] && g_nPidls == i + 1) {
			g_nPidls--;
		}
	}
	//Late Binding
	GetDisplayNameFromPidl(&bsPath, g_pidls[CSIDL_SYSTEM], SHGDN_FORPARSING);

	tePathAppend(&bsLib, bsPath, L"shell32.dll");
	g_hShell32 = GetModuleHandle(bsLib);
	SysFreeString(bsLib);
	if (g_hShell32) {
		lpfnSHCreateItemFromIDList = (LPFNSHCreateItemFromIDList)GetProcAddress(g_hShell32, "SHCreateItemFromIDList");
		lpfnSHParseDisplayName = (LPFNSHParseDisplayName)GetProcAddress(g_hShell32, "SHParseDisplayName");
		lpfnSHGetImageList = (LPFNSHGetImageList)GetProcAddress(g_hShell32, "SHGetImageList");
		lpfnSHGetIDListFromObject = (LPFNSHGetIDListFromObject)GetProcAddress(g_hShell32, "SHGetIDListFromObject");
		lpfnSHRunDialog = (LPFNSHRunDialog)GetProcAddress(g_hShell32, MAKEINTRESOURCEA(61));
		lpfnRegenerateUserEnvironment = (LPFNRegenerateUserEnvironment)GetProcAddress(g_hShell32, "RegenerateUserEnvironment");
	}
#ifdef _2000XP
	if (!lpfnSHGetIDListFromObject) {
		lpfnSHGetIDListFromObject = teGetIDListFromObjectXP;
	}
	if (!lpfnSHParseDisplayName) {
		lpfnSHParseDisplayName = teSHParseDisplayName2000;
	}
	if (!lpfnSHGetImageList) {
		lpfnSHGetImageList = teSHGetImageList2000;
	}
#endif

	tePathAppend(&bsLib, bsPath, L"kernel32.dll");
	hDll = GetModuleHandle(bsLib);
	SysFreeString(bsLib);
	if (hDll) {
		lpfnSetDllDirectoryW = (LPFNSetDllDirectoryW)GetProcAddress(hDll, "SetDllDirectoryW");
		if (lpfnSetDllDirectoryW) {
			lpfnSetDllDirectoryW(L"");
		}
		lpfnIsWow64Process = (LPFNIsWow64Process)GetProcAddress(hDll, "IsWow64Process");
	}

	tePathAppend(&bsLib, bsPath, L"propsys.dll");
	hDll = GetModuleHandle(bsLib);
	if (hDll) {
		lpfnPSPropertyKeyFromString = (LPFNPSPropertyKeyFromString)GetProcAddress(hDll, "PSPropertyKeyFromString");
		lpfnPSGetPropertyKeyFromName = (LPFNPSGetPropertyKeyFromName)GetProcAddress(hDll, "PSGetPropertyKeyFromName");
		lpfnPSGetPropertyDescription = (LPFNPSGetPropertyDescription)GetProcAddress(hDll, "PSGetPropertyDescription");
		lpfnPSPropertyKeyFromStringEx = tePSPropertyKeyFromStringEx;
	}
#ifdef _2000XP
	else {
		IPropertyUI *pPUI;
		if SUCCEEDED(CoCreateInstance(CLSID_PropertiesUI, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&pPUI))) {
			lpfnPSPropertyKeyFromStringEx = tePSPropertyKeyFromString;
			pPUI->Release();
		}
	}
#endif
	SysFreeString(bsLib);

	tePathAppend(&bsLib, bsPath, L"uxtheme.dll");
	hDll = GetModuleHandle(bsLib);
	if (hDll) {
		lpfnSetWindowTheme = (LPFNSetWindowTheme)GetProcAddress(hDll, "SetWindowTheme");
	}
	SysFreeString(bsLib);

	tePathAppend(&bsLib, bsPath, L"user32.dll");
	hDll = GetModuleHandle(bsLib);
	if (hDll) {
		lpfnChangeWindowMessageFilterEx = (LPFNChangeWindowMessageFilterEx)GetProcAddress(hDll, "ChangeWindowMessageFilterEx");
		lpfnChangeWindowMessageFilter = (LPFNChangeWindowMessageFilter)GetProcAddress(hDll, "ChangeWindowMessageFilter");
	}
	SysFreeString(bsLib);

	tePathAppend(&bsLib, bsPath, L"ntdll.dll");
	hDll = GetModuleHandle(bsLib);
	if (hDll) {
		lpfnRtlGetVersion = (LPFNRtlGetVersion)GetProcAddress(hDll, "RtlGetVersion");
	}
	SysFreeString(bsLib);

	tePathAppend(&bsLib, bsPath, L"crypt32.dll");
	g_hCrypt32 = LoadLibrary(bsLib);
	if (g_hCrypt32) {
		lpfnCryptBinaryToStringW = (LPFNCryptBinaryToStringW)GetProcAddress(g_hCrypt32, "CryptBinaryToStringW");
	}
	SysFreeString(bsLib);

	SysFreeString(bsPath);

	Initlize();
	MSG msg;

	//Initialize FolderView & TreeView settings
	g_paramFV[SB_Type] = 1;
	g_paramFV[SB_ViewMode] = FVM_DETAILS;
	g_paramFV[SB_FolderFlags] = FWF_SHOWSELALWAYS | FWF_NOWEBVIEW | FWF_AUTOARRANGE;
	g_paramFV[SB_IconSize] = 0;
	g_paramFV[SB_Options] = EBO_SHOWFRAMES | EBO_ALWAYSNAVIGATE;
	g_paramFV[SB_ViewFlags] = 0;

	g_paramFV[SB_TreeAlign] = 1;
	g_paramFV[SB_TreeWidth] = 200;
	g_paramFV[SB_TreeFlags] = NSTCS_HASEXPANDOS | NSTCS_SHOWSELECTIONALWAYS | NSTCS_BORDER | NSTCS_HASLINES;
	g_paramFV[SB_EnumFlags] = SHCONTF_FOLDERS;
	g_paramFV[SB_RootStyle] = NSTCRS_VISIBLE | NSTCRS_EXPANDED;

	// Initialize GDI+
	Gdiplus::GdiplusStartup(&g_Token, &g_StartupInput, NULL);
	MyRegisterClass(hInstance, szClass);
	// Title & Version
	lstrcpy(g_szTE, L"Tablacus Explorer " _T(STRING(VER_Y)) L"." _T(STRING(VER_M)) L"." _T(STRING(VER_D)) L" Gaku");
	if (bVisible) {
		g_hwndMain = CreateWindow(szClass, g_szTE, WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		  CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInstance, NULL);
		if (!bNewProcess) {
			SetWindowLongPtr(g_hwndMain, GWLP_USERDATA, nCrc32);
		}
	}
	else {
		g_hwndMain = CreateWindowEx(WS_EX_TOOLWINDOW, szClass, g_szTE, WS_VISIBLE | WS_POPUP,
		  CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, NULL, NULL, hInstance, NULL);
	}
	if (!g_hwndMain)
	{
		Finalize();
		return FALSE;
	}
	// ClipboardFormat
	IDLISTFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_SHELLIDLIST);
	DROPEFFECTFormat.cfFormat = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
	// Hook
	g_dwMainThreadId = GetCurrentThreadId();
	g_hHook = SetWindowsHookEx(WH_CALLWNDPROC, (HOOKPROC)HookProc, hInst, g_dwMainThreadId);
	g_hMouseHook = SetWindowsHookEx(WH_MOUSE, (HOOKPROC)MouseProc, hInst, g_dwMainThreadId);
	// Create own class
	CoCreateGuid(&g_ClsIdSB);
	CoCreateGuid(&g_ClsIdTC);
	CoCreateGuid(&g_ClsIdStruct);
	CoCreateGuid(&g_ClsIdFI);
	CoCreateGuid(&g_ClsIdFIs);
	// IUIAutomation
	if FAILED(CoCreateInstance(CLSID_CUIAutomation, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_pAutomation))) {
		g_pAutomation = NULL;
	}
	IUIAutomationRegistrar *pUIAR;
	if SUCCEEDED(CoCreateInstance(CLSID_CUIAutomationRegistrar, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pUIAR))) {
		UIAutomationPropertyInfo info = {ItemIndex_Property_GUID, L"ItemIndex", UIAutomationType_Int};
		pUIAR->RegisterProperty(&info, &g_PID_ItemIndex);
		pUIAR->Release();
	}
#ifdef _2000XP
	if (g_bUpperVista) {
#endif
		teChangeWindowMessageFilterEx(g_hwndMain, WM_COPYDATA, MSGFLT_ALLOW, NULL);
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_pDropTargetHelper));
#ifdef _2000XP
	}
#endif
	//Get JScript Object
	LPOLESTR lp = L"\
function _o() {\
	return {};\
}\
function _a() {\
	return [];\
}\
function _p(o, i, v) {\
	o[i] = v;\
}\
function _t(o) {\
	return {window: {external: o}};\
}\
";
	if (ParseScript(lp, L"JScript", NULL, NULL, &g_pJS) == S_OK) {
		GetNewArray(&g_pUnload);
	}
	// CTE
	g_pTE = new CTE();
	g_pTE->m_param[TE_CmdShow] = nCmdShow;

	IDispatch *pJS = NULL;
#ifdef _USE_HTMLDOC
	IHTMLDocument2	*pHtmlDoc = NULL;
#else
	VARIANT vWindow;
	VariantInit(&vWindow);
#endif
	if (bVisible) {
		teGetModuleFileName(NULL, &bsPath);//Executable Path
		PathRemoveFileSpec(bsPath);
		LPTSTR *lplpszArgs = NULL;
		LPWSTR lpFile = L"script\\index.html";
		if (bNewProcess) {
			int nCount = 0;
			lplpszArgs = CommandLineToArgvW(lpCmdLine, &nCount);
			if (nCount > 1) {
				lpFile = lplpszArgs[1];
			}
		}
		if (teIsFileSystem(lpFile) || PathMatchSpec(lpFile, L"file://*")) {
			bsIndex = ::SysAllocString(lpFile);			
		}
		else {
			tePathAppend(&bsIndex, bsPath, lpFile);
		}
		SysFreeString(bsPath);
		if (lplpszArgs) {
			LocalFree(lplpszArgs);
		}
		g_pWebBrowser = new CteWebBrowser(g_hwndMain, bsIndex);
		SysFreeString(bsIndex);
	}
	else {
		lp = L"\
function _s() {\
	try {\
		window.te = external;\
		api = te.WindowsAPI;\
		fso = te.CreateObject('Scripting.FileSystemObject');\
		sha = te.CreateObject('Shell.Application');\
		wsh = te.CreateObject('WScript.Shell');\
		arg = api.CommandLineToArgv(api.GetCommandLine());\
		location = {href: arg[2], hash: ''};\
		if (!api.PathMatchSpec(location.href, '?:\\*;\\\\*')) {\
			location.href = fso.BuildPath(fso.GetParentFolderName(arg[0]), location.href);\
		}\
		var wins = sha.Windows();\
		for (var i = wins.Count; i--;) {\
			var x = wins.Item(i);\
			if (x && api.strcmpi(x.FullName, arg[0]) == 0) {\
				var w = x.Document.parentWindow;\
				if (!window.MainWindow || window.MainWindow.Exchange && window.MainWindow.Exchange[arg[3]]) {\
					window.MainWindow = w;\
					var rc = api.Memory('RECT');\
					api.GetWindowRect(w.te.hwnd, rc);\
					api.MoveWindow(te.hwnd, (rc.Left + rc.Right) / 2, (rc.Top + rc.Bottom) / 2, 0, 0, false);\
				}\
			}\
		}\
		api.SetForegroundWindow(te.hwnd);\
		api.AllowSetForegroundWindow(-1);\
		return _es(location.href);\
	} catch (e) {\
		wsh.Popup((e.description || e.toString()), 0, 'Tablacus Explorer', 0x10);\
	}\
}\
function _es(fn) {\
	if (!api.PathMatchSpec(fn, '?:\\*;\\\\*')) {\
		fn = fso.BuildPath(fso.GetParentFolderName(location.href), fn);\
	}\
	try {\
		var ado = te.CreateObject('Adodb.Stream');\
		ado.CharSet = 'utf-8';\
		ado.Open();\
		ado.LoadFromFile(fn);\
		var fn = new Function(ado.ReadText());\
		ado.Close();\
		return fn();\
	} catch (e) {\
		wsh.Popup((e.description || e.toString()) + '\\n' + fn, 0, 'Tablacus Explorer', 0x10);\
	}\
}\
function importScripts() {\
	for (var i = 0; i < arguments.length; i++) {\
		_es(arguments[i]);\
	}\
}\
";
		VARIANT vResult;
		VariantInit(&vResult);
#ifdef _USE_HTMLDOC
		// CLSID_HTMLDocument
		if SUCCEEDED(CoCreateInstance(CLSID_HTMLDocument, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pHtmlDoc))) {
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
				swprintf_s((wchar_t *)pv->bstrVal, nSize, lpHtml, lp);
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
		ParseScript(lp, L"JScript", &vWindow, NULL, &pJS);
		teExecMethod(pJS, L"_s", &vResult, 0, NULL);
#endif
		bVisible = (vResult.vt == VT_BSTR) && (lstrcmpi(vResult.bstrVal, L"wait") == 0);
		VariantClear(&vResult);
	}
	if (!bVisible) {
		PostMessage(g_hwndMain, WM_CLOSE, 0, 0);
	}

	// Init ShellWindows
	if (SUCCEEDED(CoCreateInstance(CLSID_ShellWindows,
		NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&g_pSW)))) {
		if (g_pWebBrowser) {
			g_pSW->Register(g_pWebBrowser->m_pWebBrowser, (LONG)g_hwndMain, SWC_3RDPARTY, &g_lSWCookie);
		}
		teAdvise(g_pSW, DIID_DShellWindowsEvents, static_cast<IDispatch *>(g_pTE), &g_dwSWCookie);
	}

	// Main message loop:
	for (;;) {
		try {
			if (!GetMessage(&msg, NULL, 0, 0)) {
				break;
			}
			if (MessageProc(&msg) == S_OK) {
				continue;
			}
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		catch (...) {
			g_nException = 0;
		}
	}
	//At the end of processing
	try {
#ifdef _USE_HTMLDOC
		if (pHtmlDoc) {
			pHtmlDoc->Release();
		}
#else
		if (pJS) {
			pJS->Release();
		}
		VariantClear(&vWindow);
#endif
		if (g_pUnload) {
			g_pUnload->Release();
			g_pUnload = NULL;
		}
		if (g_pJS) {
			g_pJS->Release();
			g_pJS = NULL;
		}
		if (g_pAutomation) {
			g_pAutomation->Release();
			g_pAutomation = NULL;
		}
		if (g_pDropTargetHelper) {
			g_pDropTargetHelper->Release();
			g_pDropTargetHelper = NULL;
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
		KillTimer(g_hwndMain, TET_Size);
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
				teSetObjectRelease(&pInvoke->pv[pInvoke->cArgs - 1], pid);
			}
			teCoTaskMemFree(pidl);
		} catch (...) {}
	}
	Invoke5(pInvoke->pdisp, pInvoke->dispid, DISPATCH_METHOD, NULL, pInvoke->cArgs, pInvoke->pv);
	pInvoke->pdisp->Release();
	delete [] pInvoke;
}

VOID CALLBACK teTimerProcParse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);
	ParseInvoke((TEInvoke *)idEvent);
}

static void threadParseDisplayName(void *args)
{
	::OleInitialize(NULL);
	TEInvoke *pInvoke = (TEInvoke *)args;
	pInvoke->pResult = teILCreateFromPath(pInvoke->pv[pInvoke->cArgs - 1].bstrVal);
	SetTimer(g_hwndMain, (UINT_PTR)args, 100, teTimerProcParse);
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
			g_nLockUpdate += 1000;
			if (CanClose(g_pTE)) {
				teRevoke();
				DestroyWindow(hWnd);
			}
			g_nLockUpdate -= 1000;
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
			if (g_pWebBrowser) {
				g_pWebBrowser->Close();
				g_pWebBrowser = NULL;
			}
			PostQuitMessage(0);
			break;
		case WM_CONTEXTMENU:	//System Menu
			MSG msg1;
			msg1.hwnd = hWnd;
			msg1.message = message;
			msg1.wParam = wParam;
			msg1.pt.x = GET_X_LPARAM(lParam);
			msg1.pt.y = GET_Y_LPARAM(lParam);
			MessageSubPt(TE_OnShowContextMenu, g_pTE, &msg1);
			return DefWindowProc(hWnd, message, wParam, lParam);
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
			if (bDone && g_pCM) {
				LRESULT lResult = 0;
				IContextMenu3 *pCM3;
				IContextMenu2 *pCM2;
				if SUCCEEDED(g_pCM->QueryInterface(IID_PPV_ARGS(&pCM3))) {
					pCM3->HandleMenuMsg2(message, wParam, lParam, &lResult);
					pCM3->Release();
					return lResult;
				}
				else if SUCCEEDED(g_pCM->QueryInterface(IID_PPV_ARGS(&pCM2))) {
					pCM2->HandleMenuMsg(message, wParam, lParam);
					pCM2->Release();
				}
			}
			break;
		//System
		case WM_POWERBROADCAST:
			if (wParam == PBT_APMRESUMEAUTOMATIC) {
				for (int i = MAX_FV; i--;) {
					CteShellBrowser *pSB = g_pSB[i];
					if (pSB) {
						if (!pSB->m_bEmpty && pSB->m_nUnload == 0 && pSB->m_pidl) {
							VARIANT v;
							::VariantInit(&v);
							v.vt = VT_BSTR;
							if SUCCEEDED(GetDisplayNameFromPidl(&v.bstrVal, pSB->m_pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
								if (PathIsNetworkPath(v.bstrVal)) {
									pSB->Suspend(FALSE);
								}
								::VariantClear(&v);
							}
						}
					}
					else {
						break;
					}
				}
			}
		case WM_SETTINGCHANGE:
			if (message == WM_SETTINGCHANGE && lpfnRegenerateUserEnvironment) {
				try {
					if (lstrcmpi((LPCWSTR)lParam, L"Environment") == 0) {
						LPVOID lpEnvironment;
						lpfnRegenerateUserEnvironment(&lpEnvironment, TRUE);
						FreeEnvironmentStrings((LPTSTR)lpEnvironment);
					}
				}
				catch (...) {
				}
			}
		case WM_COPYDATA:
			if (g_nReload) {
				return E_FAIL;
			}
		case WM_TIMER:
		case WM_SETFOCUS:
		case WM_KILLFOCUS:
		case WM_DROPFILES:
		case WM_DEVICECHANGE:
		case WM_FONTCHANGE:
		case WM_ENABLE:
		case WM_NOTIFY:
		case WM_SETCURSOR:
		case WM_SYSCOLORCHANGE:
		case WM_SYSCOMMAND:
		case WM_THEMECHANGED:
		case WM_USERCHANGED:
		case WM_QUERYENDSESSION:
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
			if (teDoCommand(g_pTE, hWnd, message, wParam, lParam) == S_OK) {
				return 1;
			}
			return DefWindowProc(hWnd, message, wParam, lParam);
		default:
			if ((message >= WM_APP && message <= MAXWORD)) {
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
	}//end_switch
	return 0;
}

// CteShellBrowser

CteShellBrowser::CteShellBrowser(CteTabs *pTabs)
{
	m_cRef = 1;
	m_bEmpty = true;
	m_bInit = false;
	m_bsFilter = NULL;
	VariantInit(&m_vRoot);
	m_vRoot.vt = VT_I4;
	m_vRoot.lVal = 0;
	m_pTV = new CteTreeView();
	m_dwCookie = 0;
	VariantInit(&m_Data);
	Init(pTabs, true);
}

void CteShellBrowser::Init(CteTabs *pTabs, BOOL bNew)
{
	m_hwnd = NULL;
	m_hwndLV = NULL;
	m_hwndDV = NULL;
	m_pShellView = NULL;
	m_pExplorerBrowser = NULL;
	m_pServiceProvider = NULL;
	m_dwEventCookie = 0;
	m_pidl = NULL;
	m_pFolderItem = NULL;
	m_pFolderItem1 = NULL;
	m_nUnload = 0;
	m_DefProc = NULL;
	m_DefProc2 = NULL;
	m_pDropTarget = NULL;
	m_pDragItems = NULL;
	m_pDSFV = NULL;
	m_pSF2 = NULL;
#ifdef _2000XP
	m_pSFVCB = NULL;
#endif
	m_nColumns = 0;
	m_nCreate = 0;
	m_nDefultColumns = 0;
	m_bNavigateComplete = FALSE;
	VariantClear(&m_Data);

	for (int i = SB_Count; i--;) {
		m_param[i] = g_paramFV[i];
	}

	m_pTabs = pTabs;
	if (bNew) {
		m_ppLog = new FolderItem*[MAX_HISTORY];
	}
	for (int i = Count_SBFunc; i--;) {
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
	Clear();
}

void CteShellBrowser::Clear()
{
	for (int i = _countof(m_pOnFunc); i-- > 0;) {
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
		m_ppLog[m_nLogCount] = NULL;
	}
	teSysFreeString(&m_bsFilter);
	teUnadviseAndRelease(m_pDSFV, DIID_DShellFolderViewEvents, &m_dwCookie);
#ifdef _2000XP
	if (m_pSFVCB) {
		m_pSFVCB->Release();
		m_pSFVCB = NULL;
	}
#endif
	if (m_pSF2) {
		m_pSF2->Release();
		m_pSF2 = NULL;
	}
	if (m_nColumns) {
		delete [] m_pColumns;
		m_nColumns = NULL;
	}
	if (m_nDefultColumns) {
		delete [] m_pDefultColumns;
		m_nDefultColumns = NULL;
	}
	teILFreeClear(&m_pidl);
	if (m_pFolderItem) {
		m_pFolderItem->Release();
		m_pFolderItem= NULL;
	}
	if (m_pFolderItem1) {
		m_pFolderItem1->Release();
		m_pFolderItem1= NULL;
	}
	VariantClear(&m_Data);
}

STDMETHODIMP CteShellBrowser::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IShellBrowser)) {
		*ppvObject = static_cast<IShellBrowser *>(this);
	}
	else if (IsEqualIID(riid, IID_IOleWindow)) {
		*ppvObject = static_cast<IOleWindow *>(this);
	}
	else if (IsEqualIID(riid, IID_ICommDlgBrowser)) {
		*ppvObject = static_cast<ICommDlgBrowser *>(this);
	}
	else if (IsEqualIID(riid, IID_ICommDlgBrowser2)) {
		*ppvObject = static_cast<ICommDlgBrowser2 *>(this);
	}
/*	else if (IsEqualIID(riid, IID_ICommDlgBrowser3)) {
		*ppvObject = static_cast<ICommDlgBrowser3 *>(this);
	}*/
	else if (IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
	else if (IsEqualIID(riid, IID_IShellFolderViewDual)) {
		*ppvObject = static_cast<IShellFolderViewDual *>(this);
	}
	else if (IsEqualIID(riid, IID_IExplorerBrowserEvents)) {
		*ppvObject = static_cast<IExplorerBrowserEvents *>(this);
	}
	else if (IsEqualIID(riid, IID_IExplorerPaneVisibility)) {
		*ppvObject = static_cast<IExplorerPaneVisibility *>(this);
	}
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
			if (m_pShellView) {
				IFolderView2 *pFV2;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
					hr = pFV2->GetFolder(riid, ppvObject);
					pFV2->Release();
				}
			}
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
			if (m_pShellView) {
				IFolderView2 *pFV2;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
					hr = pFV2->GetFolder(riid, ppvObject);
					pFV2->Release();
				}
			}
			return hr;
		}
	}
	else if (IsEqualIID(riid, IID_IServiceProvider)) {
		*ppvObject = static_cast<IServiceProvider *>(new CteServiceProvider(reinterpret_cast<IUnknown *>(this), ILIsEmpty(m_pidl) ? NULL : m_pShellView));
		return S_OK;
	}
	else if (m_pShellView && IsEqualIID(riid, IID_IShellView)) {
		return m_pShellView->QueryInterface(riid, ppvObject);
	}
	else if (m_pFolderItem && (IsEqualIID(riid, IID_FolderItem) || IsEqualIID(riid, g_ClsIdFI))) {
		return m_pFolderItem->QueryInterface(riid, ppvObject);
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

//IOleWindow
STDMETHODIMP CteShellBrowser::GetWindow(HWND *phwnd)
{
	*phwnd = m_pTabs->m_hwndStatic;
	return S_OK;
}

STDMETHODIMP CteShellBrowser::ContextSensitiveHelp(BOOL fEnterMode)
{
	return E_NOTIMPL;
}

//IShellBrowser
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
			TEmethod pull_downs[] = {
				{ FCIDM_MENU_FILE, L"&File" },
				{ FCIDM_MENU_EDIT, L"&Edit" },
				{ FCIDM_MENU_VIEW, L"&View" },
				{ FCIDM_MENU_TOOLS, L"&Tools" },
				{ FCIDM_MENU_HELP, L"&Help" },
			};
			for (size_t i = 0; i < 5; i++) {
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
	for (int nCount = GetMenuItemCount(hmenuShared); nCount-- > 0;) {
		DeleteMenu(hmenuShared, nCount, MF_BYPOSITION);
	}
	return S_OK;
}

STDMETHODIMP CteShellBrowser::SetStatusTextSB(LPCWSTR lpszStatusText)
{
	KillTimer(g_hwndMain, TET_Status);
	if (lpszStatusText) {
		lstrcpyn(g_szStatus, lpszStatusText, _countof(g_szStatus) - 1);
	}
	else {
		g_szStatus[0] = NULL;
	}
	SetTimer(g_hwndMain, TET_Status, 500, teTimerProc);
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
	FolderItem *pid = NULL;
	if ((wFlags & SBSP_REDIRECT) && m_pidl) {
		if (m_nLogIndex >= 0 && m_nLogIndex < m_nLogCount && m_ppLog[m_nLogIndex]) {
			m_ppLog[m_nLogIndex]->QueryInterface(IID_PPV_ARGS(&pid));
		}
		else {
			VARIANT v;
			GetVarPathFromPidl(&v, m_pidl, SHGDN_FORPARSINGEX | SHGDN_FORPARSING);
			pid = new CteFolderItem(&v);
			VariantClear(&v);
		}
		teCoTaskMemFree(m_pidl);
		m_pidl = NULL;
		wFlags &= ~(SBSP_RELATIVE | SBSP_REDIRECT);
	}
	else {
		GetFolderItemFromPidl(&pid, const_cast<LPITEMIDLIST>(pidl));
	}
	return BrowseObject2(pid, wFlags);
}

HRESULT CteShellBrowser::BrowseObject2(FolderItem *pid, UINT wFlags)
{
	CteShellBrowser *pSB = NULL;
	DWORD param[SB_Count];
	for (int i = SB_Count; i--;) {
		param[i] = m_param[i];
	}
	param[SB_DoFunc] = 1;
	HRESULT hr = Navigate3(pid, wFlags, param, &pSB, NULL);
	pSB && pSB->Release();
	return hr;
}

HRESULT CteShellBrowser::Navigate3(FolderItem *pFolderItem, UINT wFlags, DWORD *param, CteShellBrowser **ppSB, FolderItems *pFolderItems)
{
	HRESULT hr = E_FAIL;
	m_pTabs->LockUpdate();
	try {
		if (!(wFlags & SBSP_NEWBROWSER || m_bEmpty)) {
			hr = Navigate2(pFolderItem, wFlags, param, pFolderItems, m_pFolderItem, this);
			if (hr == E_ACCESSDENIED) {
				wFlags = (wFlags | SBSP_NEWBROWSER) & (~SBSP_ACTIVATE_NOFOCUS);
			}
		}
		if (wFlags & SBSP_NEWBROWSER || (m_bEmpty && m_bVisible)) {
			CteShellBrowser *pSB = this;
			BOOL bNew = !m_bEmpty && (wFlags & SBSP_NEWBROWSER);
			if (bNew) {
				pSB = GetNewShellBrowser(m_pTabs);
				*ppSB = pSB;
				VariantCopy(&pSB->m_vRoot, &m_vRoot);
			}
			for (int i = SB_Count; i--;) {
				pSB->m_param[i] = param[i];
			}
			pSB->m_bEmpty = false;
			hr = pSB->Navigate2(pFolderItem, wFlags, param, pFolderItems, m_pFolderItem, this);
			if SUCCEEDED(hr) {
				if (bNew && (wFlags & SBSP_ACTIVATE_NOFOCUS) == 0) {
					TabCtrl_SetCurSel(m_pTabs->m_hwnd, m_pTabs->m_nIndex + 1);
					m_pTabs->m_nIndex = TabCtrl_GetCurSel(m_pTabs->m_hwnd);
				}
				else {
					m_pTabs->TabChanged(TRUE);
				}
			}
			else {
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

VOID CteShellBrowser::CheckNavigate(LPITEMIDLIST *ppidl, CteShellBrowser *pHistSB, int nLogIndex)
{
	try {
		if (!teILIsExists(*ppidl)) {
			teILFreeClear(ppidl);
		}

		if (pHistSB) {			//Delete path from the history if it fails.
			if (!*ppidl) {
				pHistSB->m_ppLog[nLogIndex]->Release();
				for (int i = nLogIndex; i < pHistSB->m_nLogCount - 1; i++) {
					pHistSB->m_ppLog[i] = pHistSB->m_ppLog[i + 1];
				}
				pHistSB->m_nLogCount--;
				return;
			}
			if (this == pHistSB) {
				m_nLogIndex = nLogIndex;
			}
		}
	}
	catch (...) {
	}
}

BOOL CteShellBrowser::Navigate1(FolderItem *pFolderItem, UINT wFlags, FolderItems *pFolderItems, FolderItem *pPrevious, LPITEMIDLIST *ppidl)
{
	CteFolderItem *pid = NULL;
	if (pFolderItem) {
		if SUCCEEDED(pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid)) {
			if (!pid->m_pidl) {
				if (pid->m_v.vt == VT_BSTR) {
					if (PathIsNetworkPath(pid->m_v.bstrVal)) {
						Navigate1Ex(pid->m_v.bstrVal, pFolderItems, wFlags, pPrevious);
						pid->Release();
						return TRUE;
					}
					if (ppidl) {
						pid->GetCurFolder(ppidl);
						if (!pid->m_pidl && m_pShellView) {
							teILFreeClear(ppidl);
							pid->Release();
							return TRUE;
						}
						pid->Release();
						return FALSE;
					}
				}
			}
			pid->Release();
		}
		if (ppidl) {
			GetIDListFromObject(pFolderItem, ppidl);
		}
	}
	return FALSE;
}

VOID CteShellBrowser::Navigate1Ex(LPOLESTR pstr, FolderItems *pFolderItems, UINT wFlags, FolderItem *pPrevious)
{
	TEInvoke *pInvoke = new TEInvoke[1];
	QueryInterface(IID_PPV_ARGS(&pInvoke->pdisp));
	pInvoke->dispid = 0x20000007;//Navigate2Ex
	pInvoke->cArgs = 4;
	pInvoke->pv = GetNewVARIANT(4);
	pInvoke->pv[3].bstrVal = ::SysAllocString(pstr);
	pInvoke->pv[3].vt = VT_BSTR;
	pInvoke->pv[2].lVal = wFlags;
	pInvoke->pv[2].vt = VT_I4;
	teSetObject(&pInvoke->pv[1], pFolderItems);
	teSetObject(&pInvoke->pv[0], pPrevious);
	_beginthread(&threadParseDisplayName, 0, pInvoke);
}

HRESULT CteShellBrowser::Navigate2(FolderItem *pFolderItem, UINT wFlags, DWORD *param, FolderItems *pFolderItems, FolderItem *pPrevious, CteShellBrowser *pHistSB)
{
	RECT rc;
	LPITEMIDLIST pidl = NULL;
	HRESULT hr;

	m_pFolderItem1 && m_pFolderItem1->Release();
	m_pFolderItem1 = NULL;

	if (m_hwnd) {
		KillTimer(m_hwnd, (UINT_PTR)this);
	}
	SetRedraw(FALSE);
	if (m_bInit) {
		return E_FAIL;
	}
	for (int i = SB_Count; i-- > 0;) {
		m_param[i] = param[i];
		g_paramFV[i] = param[i];
	}
	m_nPrevLogIndex = m_nLogIndex;
	SaveFocusedItemToHistory();
	hr = GetAbsPidl(&pidl, &m_pFolderItem1, pFolderItem, wFlags, pFolderItems, pPrevious, pHistSB);
	if (hr != S_OK) {
		if (hr == S_FALSE) {
			if (!m_pFolderItem) {
				if SUCCEEDED(pFolderItem->QueryInterface(IID_PPV_ARGS(&m_pFolderItem))) {
					CteFolderItem *pid1 = NULL;
					if (SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) && !pid1->m_pidl) {
						pid1->m_pidl = ::ILClone(g_pidlResultsFolder);
						pid1->Release();
					}
				}
			}
		}
		teCoTaskMemFree(pidl);
		return m_pidl ? S_OK : hr;
	}
	hr = S_FALSE;
#ifndef _WIN64
	Wow64ControlPanel(&pidl, g_pidls[CSIDL_CONTROLS]);
#endif
	if (ILIsEmpty(pidl) && !g_nLockUpdate) {
		BSTR bs;
		if (m_pFolderItem1) {
			if SUCCEEDED(m_pFolderItem1->get_Path(&bs)) {
				if (!teIsDesktopPath(bs)) {
					teILCloneReplace(&pidl, g_pidlResultsFolder);
				}
				SysFreeString(bs);
			}
			else {
				Navigate1Ex(L"0", NULL, SBSP_SAMEBROWSER, NULL);
				teILFreeClear(&pidl);
			}
		}
	}
	if (!pidl) {
		return S_OK;
	}
	IShellFolder *pSF;
	LPCITEMIDLIST ItemID;
	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &ItemID)) {
		SFGAOF sfAttr = SFGAO_LINK;
		if SUCCEEDED(pSF->GetAttributesOf(1, const_cast<LPCITEMIDLIST *>(&ItemID), &sfAttr)) {
			if (sfAttr & SFGAO_LINK) {
				IShellLink *pSL;
				if SUCCEEDED(pSF->GetUIObjectOf(NULL, 1, const_cast<LPCITEMIDLIST *>(&ItemID), IID_IShellLink, NULL, (LPVOID *)&pSL)) {
					if (pSL->Resolve(NULL, SLR_NO_UI) == S_OK) {
						teILFreeClear(&pidl);
						if (pSL->GetIDList(&pidl) == S_OK) {
							if (m_pFolderItem1) {
								m_pFolderItem1->Release();
								GetFolderItemFromPidl(&m_pFolderItem1, pidl);
							}
						}
					}
					pSL->Release();
				}
			}
		}
		pSF->Release();
	}
	BOOL bExplorerBrowser = FALSE;
	DWORD dwFrame = 0;
	FOLDERVIEWOPTIONS fvo;
#ifdef _2000XP
	if (g_bUpperVista) {
#endif
		bExplorerBrowser = (m_param[SB_Type] == 2) || ILIsParent(g_pidlCP, pidl, FALSE);
		if (bExplorerBrowser) {
			dwFrame = EBO_SHOWFRAMES;
		}
		fvo = teGetFolderViewOptions(pidl, m_param[SB_ViewMode]);
		bExplorerBrowser |= fvo == FVO_DEFAULT;
#ifdef _2000XP
	}
#endif
	//ExplorerBrowser
	if (bExplorerBrowser && m_pExplorerBrowser && !((dwFrame ^ m_param[SB_Options]) & EBO_SHOWFRAMES)) {
		if (!ILIsEqual(pidl, g_pidlResultsFolder) || !ILIsEqual(m_pidl, g_pidlResultsFolder)) {
			IFolderViewOptions *pOptions;
			FOLDERVIEWOPTIONS fvo0 = fvo;
			if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
				pOptions->GetFolderViewOptions(&fvo0);
				pOptions->Release();
			}
			if (fvo == (fvo0 & FVO_VISTALAYOUT)) {
				if SUCCEEDED(GetShellFolder2(&m_pSF2, pidl, m_bsFilter)) {
					m_pExplorerBrowser->SetOptions(static_cast<EXPLORER_BROWSER_OPTIONS>((m_param[SB_Options] & ~EBO_SHOWFRAMES) | dwFrame | EBO_NOTRAVELLOG));
					if SUCCEEDED(m_pExplorerBrowser->BrowseToObject(m_pSF2, SBSP_ABSOLUTE)) {
						IOleWindow *pOleWindow;
						if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
							pOleWindow->GetWindow(&m_hwnd);
							pOleWindow->Release();
						}
						teCoTaskMemFree(pidl);
						return S_OK;
					}
				}
			}
		}
	}
	teCoTaskMemFree(m_pidl);
	m_pidl = pidl;
	m_pFolderItem = m_pFolderItem1;
	m_pFolderItem1 = NULL;
	if (g_nLockUpdate || !m_pTabs->m_bVisible) {
		m_nUnload = MAX_TC - m_pTabs->m_nTC;
		m_nUnload = 1;
		//History / Management
		SetHistory(pFolderItems, wFlags);
		return S_OK;
	}
	//Previous display
	IShellView *pPreviousView = m_pShellView;
	EXPLORER_BROWSER_OPTIONS dwflag;
	dwflag = static_cast<EXPLORER_BROWSER_OPTIONS>(m_param[SB_Options]);

	m_bIconSize = FALSE;
	if (param[SB_DoFunc]) {
		hr = OnBeforeNavigate(pPrevious, wFlags);
		if (hr == E_ACCESSDENIED && (teILIsEqual(m_pFolderItem, pPrevious) || !m_pShellView)) {
			hr = S_OK;
		}
		if (hr != S_OK && m_nUnload != 2) {
			teILFreeClear(&m_pidl);
			GetIDListFromObject(pPrevious, &m_pidl);
			if (pPrevious) {
				pPrevious->AddRef();
				m_pFolderItem && m_pFolderItem->Release();
				m_pFolderItem = pPrevious;
			}
			m_nLogIndex = m_nPrevLogIndex;
			return hr;
		}
		if (!teILIsEqual(m_pFolderItem, pPrevious)) {
			teSysFreeString(&m_bsFilter);
			if (m_pOnFunc[SB_OnIncludeObject]) {
				m_pOnFunc[SB_OnIncludeObject]->Release();
				m_pOnFunc[SB_OnIncludeObject] = NULL;
			}
		}
	}

	SetRectEmpty(&rc);

	//History / Management
	SetHistory(pFolderItems, wFlags);
	DestroyView(m_param[SB_Type]);
	Show(false);
	//ShellBrowser
	if (!bExplorerBrowser) {
		IShellFolder *pShellFolder = NULL;
		if (m_nColumns) {
			delete [] m_pColumns;
			m_nColumns = 0;
		}
		if SUCCEEDED(GetShellFolder2(&m_pSF2, m_pidl, m_bsFilter)) {
			try {
				hr = CreateViewWindowEx(pPreviousView);
			}
			catch (...) {
				hr = E_FAIL;
			}
		}
		if (hr != S_OK) {
			if (hr == S_FALSE
#ifdef _2000XP
				&& g_bUpperVista
#endif
			) {
				bExplorerBrowser = true;//Use ExplorerBrowser
			}
			else {
				BSTR bs = NULL;
				m_pFolderItem->get_Path(&bs);
				if (hr == E_CANCELLED_BY_USER) {
					if (pPrevious) {
						BSTR bs2 = NULL;
						pPrevious->get_Path(&bs2);
						if (lstrcmpi(bs2, bs)) {
							teSysFreeString(&bs);
							bs = bs2;
							hr = OLE_E_BLANK;
						}
						else {
							::SysFreeString(bs2);
						}
					}
				}
				if (hr == OLE_E_BLANK && bs) {
					Navigate1Ex(bs, NULL, SBSP_SAMEBROWSER, NULL);
				}
				else {
					Error(&bs);
				}
				teSysFreeString(&bs);
				hr = S_FALSE;
			}
		}
	}
	//ExplorerBrowser
	if (bExplorerBrowser) {
		hr = NavigateEB(dwFrame);
	}
	if (hr != S_OK) {
		if (m_pShellView) {
			m_pShellView->Release();
		}
		m_pShellView = pPreviousView;

		teILFreeClear(&m_pidl);
		GetIDListFromObject(pPrevious, &m_pidl);
		if (pPrevious) {
			pPrevious->AddRef();
			m_pFolderItem && m_pFolderItem->Release();
			m_pFolderItem = pPrevious;
		}
		m_nLogIndex = m_nPrevLogIndex;
		return hr;
	}

	if (pPreviousView) {
//		pPreviousView->SaveViewState();
		pPreviousView->DestroyViewWindow();
		pPreviousView->Release();
		pPreviousView = NULL;
	}
	else {
		IFolderView2 *pFV2;
		if (m_pShellView) {
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
			else {
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

	m_nOpenedType = m_param[SB_Type];
	if (m_pTabs && GetTabIndex() == m_pTabs->m_nIndex) {
		Show(true);
	}
	if (!m_pExplorerBrowser) {
		OnViewCreated(NULL);
		ArrangeWindow();
	}
	return S_OK;
}

HRESULT CteShellBrowser::NavigateEB(DWORD dwFrame)
{
	HRESULT hr = E_FAIL;
	RECT rc;
	if (m_pExplorerBrowser) {
		DestroyView(CTRL_SB);
	}
	if (SUCCEEDED(CoCreateInstance(CLSID_ExplorerBrowser,
	NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pExplorerBrowser)))) {
		FOLDERSETTINGS fs = { m_param[SB_ViewMode], (m_param[SB_FolderFlags] | FWF_USESEARCHFOLDER) & ~FWF_NOENUMREFRESH };
		if (SUCCEEDED(m_pExplorerBrowser->Initialize(m_pTabs->m_hwndStatic, &rc, &fs))) {
			if SUCCEEDED(GetShellFolder2(&m_pSF2, m_pidl, m_bsFilter)) {
				m_pExplorerBrowser->Advise(static_cast<IExplorerBrowserEvents *>(this), &m_dwEventCookie);
				IFolderViewOptions *pOptions;
				if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
					pOptions->SetFolderViewOptions(FVO_VISTALAYOUT, teGetFolderViewOptions(m_pidl, m_param[SB_ViewMode]));
					pOptions->Release();
				}
				m_pExplorerBrowser->SetOptions(static_cast<EXPLORER_BROWSER_OPTIONS>((m_param[SB_Options] & ~EBO_SHOWFRAMES) | dwFrame | EBO_NOTRAVELLOG));
				m_pServiceProvider = new CteServiceProvider(reinterpret_cast<IUnknown *>(this), NULL);
				IUnknown_SetSite(m_pExplorerBrowser, m_pServiceProvider);
				hr = m_pExplorerBrowser->BrowseToObject(m_pSF2, SBSP_ABSOLUTE);
			}
			if (hr == S_OK) {
				if (!m_pShellView) {
					m_pExplorerBrowser->GetCurrentView(IID_PPV_ARGS(&m_pShellView));
				}
				if (m_pShellView) {
					GetDefaultColumns();
				}
				if (m_pExplorerBrowser) {
					IOleWindow *pOleWindow;
					if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
						pOleWindow->GetWindow(&m_hwnd);
						pOleWindow->Release();
					}
				}
			}
		}
	}
	return hr;
}

HRESULT CteShellBrowser::OnBeforeNavigate(FolderItem *pPrevious, UINT wFlags)
{
	HRESULT hr = S_OK;
	if (m_pShellView) {
		GetViewModeAndIconSize(TRUE);
		if (m_pExplorerBrowser) {
			m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
		}
	}
	if (g_pOnFunc[TE_OnBeforeNavigate]) {
		VARIANT vResult;
		VariantInit(&vResult);

		VARIANTARG *pv;
		CteMemory *pMem;
		pv = GetNewVARIANT(4);
		teSetObject(&pv[3], this);

		pMem = new CteMemory(4 * sizeof(int), (char *)&m_param[SB_ViewMode], VT_NULL, 1, L"FOLDERSETTINGS");
		pMem->QueryInterface(IID_PPV_ARGS(&pv[2].punkVal));
		teSetObjectRelease(&pv[2], pMem);

		pv[1].vt = VT_I4;
		pv[1].lVal = wFlags;
		if (!teILIsEqual(m_pFolderItem, pPrevious)) {
			teSetObject(&pv[0], pPrevious);
		}
		m_bInit = true;
		try {
			Invoke4(g_pOnFunc[TE_OnBeforeNavigate], &vResult, 4, pv);
		}
		catch (...) {}
		m_bInit = false;
		hr = GetIntFromVariantClear(&vResult);
	}
	return hr;
}

VOID CteShellBrowser::SaveFocusedItemToHistory()
{
	if (m_nLogIndex >= 0 && m_nLogIndex < m_nLogCount && m_pShellView && m_ppLog[m_nLogIndex]) {
		IFolderView *pFV;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
			int i;
			if SUCCEEDED(pFV->GetFocusedItem(&i)) {
				CteFolderItem *pid1 = NULL;
				if FAILED(m_ppLog[m_nLogIndex]->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) {
					pid1 = new CteFolderItem(NULL);
					if SUCCEEDED(pid1->Initialize(m_pidl)) {
						m_ppLog[m_nLogIndex]->Release();
						pid1->QueryInterface(IID_PPV_ARGS(&m_ppLog[m_nLogIndex]));
					}
				}
				pFV->Item(i, &pid1->m_pidlFocused);
				pFV->ItemCount(SVGIO_SELECTION, &pid1->m_nSelected);
				pid1->Release();
			}
			pFV->Release();
		}
	}
}

HRESULT CteShellBrowser::GetAbsPidl(LPITEMIDLIST *ppidlOut, FolderItem **ppid, FolderItem *pid, UINT wFlags, FolderItems *pFolderItems, FolderItem *pPrevious, CteShellBrowser *pHistSB)
{
	if (wFlags & SBSP_PARENT) {
		if (pPrevious) {
			if (m_bsFilter) {
				teSysFreeString(&m_bsFilter);
				Refresh(TRUE);
				return E_FAIL;
			}
			if (teILGetParent(pPrevious, ppid)) {
				GetIDListFromObject(*ppid, ppidlOut);
				VARIANT v;
				if SUCCEEDED(ppid[0]->QueryInterface(IID_PPV_ARGS(&v.pdispVal))) {
					v.vt = VT_DISPATCH;
					CteFolderItem *pid1 = new CteFolderItem(&v);
					GetIDListFromObject(pPrevious, &pid1->m_pidlFocused);
					pid1->m_nSelected = 1;
					ppid[0]->Release();
					ppid[0] = pid1;
					VariantClear(&v);
				}
				return S_OK;
			}
		}
		*ppidlOut = ::ILClone(g_pidls[CSIDL_DESKTOP]);
		GetFolderItemFromPidl(ppid, *ppidlOut);
		return S_OK;
	}
	if (wFlags & SBSP_NAVIGATEBACK) {
		int nLogIndex = pHistSB->m_nLogIndex;
		if (nLogIndex > 0 && GetIDListFromObjectEx(pHistSB->m_ppLog[--nLogIndex], ppidlOut)) {
			pHistSB->m_ppLog[nLogIndex]->QueryInterface(IID_PPV_ARGS(ppid));
			CheckNavigate(ppidlOut, pHistSB, nLogIndex);
			return S_OK;
		}
		return E_FAIL;
	}
	if (wFlags & SBSP_NAVIGATEFORWARD) {
		int nLogIndex = pHistSB->m_nLogIndex;
		if (nLogIndex < pHistSB->m_nLogCount - 1 && GetIDListFromObjectEx(pHistSB->m_ppLog[++nLogIndex], ppidlOut)) {
			pHistSB->m_ppLog[nLogIndex]->QueryInterface(IID_PPV_ARGS(ppid));
			CheckNavigate(ppidlOut, pHistSB, nLogIndex);
			return S_OK;
		}
		return E_FAIL;
	}
	LPITEMIDLIST pidl = NULL;
	if (Navigate1(pid, wFlags, pFolderItems, pPrevious, &pidl)) {
		return S_FALSE;
	}
	if (wFlags & SBSP_RELATIVE) {
		LPITEMIDLIST pidlPrevius = NULL;
		GetIDListFromObject(pPrevious, &pidlPrevius);
		if (pidlPrevius && !ILIsEqual(pidlPrevius, g_pidlResultsFolder) && !ILIsEmpty(pidl)) {
			*ppidlOut = ILCombine(pidlPrevius, pidl);
			GetFolderItemFromPidl(ppid, *ppidlOut);
		}
		else {
			*ppidlOut = ILClone(pidlPrevius);
			if (pPrevious) {
				pPrevious->QueryInterface(IID_PPV_ARGS(ppid));
			}
			else {
				GetFolderItemFromPidl(ppid, *ppidlOut);
			}
		}
		teCoTaskMemFree(pidlPrevius);
		teCoTaskMemFree(pidl);
		return S_OK;
	}
	if (pid) {
		pid->QueryInterface(IID_PPV_ARGS(ppid));
		if (pidl) {
			*ppidlOut = pidl;
			return S_OK;
		}
	}
	if (*ppid) {
		return GetIDListFromObject(*ppid, ppidlOut) ? S_OK : E_FAIL;
	}
	return E_FAIL;
}

VOID CteShellBrowser::HookDragDrop(int nMode)
{
	if (m_hwnd && nMode & 1) {
		if (!m_pDropTarget) {
			teRegisterDragDrop(m_hwndLV, this, &m_pDropTarget);
		}
	}
	if (nMode & 2 && m_pTV && m_pTV->m_hwndTV && !m_pTV->m_pDropTarget) {
		teRegisterDragDrop(m_pTV->m_hwndTV, static_cast<IDropTarget *>(m_pTV), &m_pTV->m_pDropTarget);
	}
}

VOID CteShellBrowser::Error(BSTR *pbs)
{
	if (!PathRemoveFileSpec(*pbs)) {
		SysReAllocString(pbs, lstrlen(*pbs) == 3 ? L"17" : L"0");
	}
	Navigate1Ex(*pbs, NULL, SBSP_SAMEBROWSER, NULL);
}

VOID CteShellBrowser::Refresh(BOOL bCheck)
{
	if (bCheck) {
		BSTR bs;
		if SUCCEEDED(GetDisplayNameFromPidl(&bs, m_pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
			if (PathIsNetworkPath(bs)) {
				TEInvoke *pInvoke = new TEInvoke[1];
				QueryInterface(IID_PPV_ARGS(&pInvoke->pdisp));
				pInvoke->dispid = 0x20000206;//Refresh2Ex
				pInvoke->cArgs = 1;
				pInvoke->pv = GetNewVARIANT(1);
				VariantInit(&pInvoke->pv[0]);
				pInvoke->pv[0].bstrVal = bs;
				pInvoke->pv[0].vt = VT_BSTR;
				_beginthread(&threadParseDisplayName, 0, pInvoke);
				return;
			}
			::SysFreeString(bs);
		}
	}
	if (m_pShellView) {
		if (m_bVisible) {
			if (m_nUnload == 4) {
				m_nUnload = 0;
			}
			if ((ILIsEqual(m_pidl, g_pidlResultsFolder) || (g_pidlLibrary && ILIsParent(g_pidlLibrary, m_pidl, FALSE)))) {
				BrowseObject(NULL, SBSP_RELATIVE | SBSP_SAMEBROWSER);
				return;
			}
			m_pShellView->Refresh();
		}
		else if (m_nUnload == 0) {
			m_nUnload = 4;
		}
	}
}

BOOL CteShellBrowser::SetActive(BOOL bForce)
{
	if (!m_bVisible) {
		return FALSE;
	}
	HWND hwnd = GetFocus();
	if (!bForce) {
		if (hwnd) {
			if (IsChild(g_hwndBrowser, hwnd) || (m_pTV && IsChild(m_pTV->m_hwnd, hwnd))) {
				return FALSE;
			}
			WCHAR szClass[MAX_CLASS_NAME];
			GetClassName(hwnd, szClass, MAX_CLASS_NAME);
			if (PathMatchSpec(szClass, WC_TREEVIEW)) {
				return FALSE;
			}
		}
	}
	if (m_pShellView) {
		if (IsChild(m_hwnd, hwnd)) {
			SetFocus(m_hwnd);
		}
		m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
	}
	return TRUE;
}

VOID CteShellBrowser::SetTitle(BSTR szName, int nIndex)
{
	TC_ITEM tcItem;
	BSTR bsText = SysAllocStringLen(NULL, MAX_PATH);
	BSTR bsOldText = SysAllocStringLen(NULL, MAX_PATH);
	bsOldText[0] = NULL;
	tcItem.pszText = bsOldText;
	tcItem.mask = TCIF_TEXT;
	tcItem.cchTextMax = MAX_PATH;
	int nCount = ::SysStringLen(szName);
	if (nCount >= MAX_PATH) {
		nCount = MAX_PATH - 1;
	}
	int j = 0;
	for (int i = 0; i < nCount; i++) {
		bsText[j++] = szName[i];
		if (szName[i] == '&') {
			bsText[j++] = '&';
		}
	}
	bsText[j] = NULL;
	TabCtrl_GetItem(m_pTabs->m_hwnd, nIndex, &tcItem);
	if (lstrcmpi(bsText, bsOldText)) {
		tcItem.pszText = bsText;
		TabCtrl_SetItem(m_pTabs->m_hwnd, nIndex, &tcItem);
		ArrangeWindow();
	}
	SysFreeString(bsOldText);
	SysFreeString(bsText);
}

VOID CteShellBrowser::GetShellFolderView()
{
	teUnadviseAndRelease(m_pDSFV, DIID_DShellFolderViewEvents, &m_dwCookie);
	if SUCCEEDED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&m_pDSFV))) {
		teAdvise(m_pDSFV, DIID_DShellFolderViewEvents, static_cast<IDispatch *>(this), &m_dwCookie);
	}
	else {
		m_pDSFV = NULL;
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
	IFolderView *pFV;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
		pFV->GetFocusedItem(piItem);
		pFV->Release();
	}
}

VOID CteShellBrowser::SetFolderFlags()
{
	if (m_pShellView) {
		IFolderView2 *pFV2;
		if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2)))) {
			pFV2->SetCurrentFolderFlags(~(FWF_NOENUMREFRESH | FWF_USESEARCHFOLDER), m_param[SB_FolderFlags]);
			pFV2->Release();
		}
#ifdef _2000XP
		else {
			PostMessage(m_hwndLV, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES, HIWORD(m_param[SB_FolderFlags]));
		}
#endif
		GetViewModeAndIconSize(FALSE);
	}
}

VOID CteShellBrowser::GetViewModeAndIconSize(BOOL bGetIconSize)
{
	FOLDERSETTINGS fs;
	UINT uViewMode = 0;
	int iImageSize = m_param[SB_IconSize];
	if SUCCEEDED(m_pShellView->GetCurrentInfo(&fs)) {
		uViewMode = fs.ViewMode;
		m_param[SB_FolderFlags] = fs.fFlags;
	}
	if (bGetIconSize || uViewMode != m_param[SB_ViewMode]) {
		IFolderView2 *pFV2;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
			pFV2->GetViewModeAndIconSize((FOLDERVIEWMODE *)&uViewMode, &iImageSize);
			pFV2->Release();
		}
#ifdef _2000XP
		else if (!g_bUpperVista) {
			if (uViewMode == FVM_ICON && iImageSize >= 96) {
				uViewMode = FVM_THUMBNAIL;
				IFolderView *pFV;
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
					pFV->SetCurrentViewMode(uViewMode);
					pFV->Release();
				}
			}
			if (uViewMode == FVM_ICON) {
				iImageSize = GetSystemMetrics(SM_CXICON);
			}
			else if (uViewMode == FVM_TILE) {
				iImageSize = 48;
			}
			else if (uViewMode == FVM_THUMBNAIL) {
				HKEY hKey;
				if (RegOpenKeyEx(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
					DWORD dwSize = sizeof(iImageSize);
					RegQueryValueEx(hKey, L"ThumbnailSize", NULL, NULL, (LPBYTE)&iImageSize, &dwSize);
					RegCloseKey(hKey);
				}
			}
			else {
				iImageSize = GetSystemMetrics(SM_CXSMICON);
			}
		}
#endif
		m_param[SB_IconSize] = iImageSize;
		if (uViewMode != m_param[SB_ViewMode]) {
			m_param[SB_ViewMode] = uViewMode;
			if (m_pExplorerBrowser) {
				IFolderViewOptions *pOptions;
				if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
					FOLDERVIEWOPTIONS fvo = teGetFolderViewOptions(m_pidl, fs.ViewMode);
					FOLDERVIEWOPTIONS fvo0 = fvo;
					pOptions->GetFolderViewOptions(&fvo0);
					teDoCommand(this, m_hwnd, WM_NULL, 0, 0);
					if (fvo != (fvo0 & FVO_VISTALAYOUT)) {
						SetTimer(m_hwnd, TET_Refresh, 16, teTimerProc);
					}
					pOptions->Release();
				}
			}
			else if (teGetFolderViewOptions(m_pidl, m_param[SB_ViewMode]) == FVO_DEFAULT) {
				teDoCommand(this, m_hwnd, WM_NULL, 0, 0);
				SetTimer(m_hwnd, TET_Refresh, 16, teTimerProc);
			}
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
		if (!(m_nLogCount > 0 && teILIsEqual(m_pFolderItem, m_ppLog[m_nLogIndex]))) {
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
	}
}

STDMETHODIMP CteShellBrowser::GetViewStateStream(DWORD grfMode, IStream **ppStrm)
{
//	return SHCreateStreamOnFileEx(L"", STGM_READWRITE | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, TRUE, NULL ppStrm);
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
	if (m_pTabs) {
		CheckChangeTabTC(m_pTabs->m_hwnd, false);
	}
/*/// For check
	BSTR bsPath;
	if SUCCEEDED(GetDisplayNameFromPidl(&bsPath, m_pidl, SHGDN_FORPARSING)) {
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
	if (m_bsFilter) {
		HRESULT hr = S_OK;
		if (m_pSF2) {
			STRRET strret;
			if SUCCEEDED(m_pSF2->GetDisplayNameOf(pidl, SHGDN_INFOLDER | SHGDN_FORPARSING, &strret)) {
				BSTR bs;
				if SUCCEEDED(StrRetToBSTR(&strret, pidl, &bs)) {
					if (!tePathMatchSpec(bs, m_bsFilter)) {
						hr = S_FALSE;
						if SUCCEEDED(m_pSF2->GetDisplayNameOf(pidl, SHGDN_NORMAL, &strret)) {
							BSTR bs2;
							if SUCCEEDED(StrRetToBSTR(&strret, pidl, &bs2)) {
								if (tePathMatchSpec(bs2, m_bsFilter)) {
									hr = S_OK;
								}
								::SysFreeString(bs2);
							}
						}
					}
					::SysFreeString(bs);
				}
			}
		}
		return hr;
	}
	if (m_pOnFunc[SB_OnIncludeObject]) {
		VARIANTARG *pv;
		VARIANT vResult;
		VariantInit(&vResult);
		pv = GetNewVARIANT(2);
		teSetObject(&pv[1], this);
		LPITEMIDLIST pidlParent, pidlFull;
		if (GetIDListFromObject(m_pSF2, &pidlParent)) {
			pidlFull = ILCombine(pidlParent, pidl);
			FolderItem *pFolderItem;
			if (GetFolderItemFromPidl(&pFolderItem, pidlFull)) {
				teSetObjectRelease(&pv[0], pFolderItem);
			}
			teCoTaskMemFree(pidlFull);
			teCoTaskMemFree(pidlParent);
		}
		Invoke4(m_pOnFunc[SB_OnIncludeObject], &vResult, 2, pv);
		return GetIntFromVariantClear(&vResult);
	}
	return S_OK;
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
	*pdwFlags = (m_param[SB_ViewFlags] & (~CDB2GVF_NOINCLUDEITEM)) | ((m_bsFilter || m_pOnFunc[SB_OnIncludeObject] || ILIsEqual(m_pidl, g_pidlResultsFolder)) ? 0 : CDB2GVF_NOINCLUDEITEM);
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
	if (m_pOnFunc[SB_OnIncludeObject]) {
		VARIANTARG *pv;
		VARIANT vResult;
		VariantInit(&vResult);
		pv = GetNewVARIANT(2);
		teSetObject(&pv[1], this);
		LPITEMIDLIST FullID;
		FullID = ILCombine(pidlFolder, pidlItem);
		FolderItem *pFolderItem;
		if (GetFolderItemFromPidl(&pFolderItem, FullID)) {
			teSetObject(&pv[0], pFolderItem);
			pFolderItem->Release();
		}
		teCoTaskMemFree(FullID);
		Invoke4(m_pOnFunc[SB_OnIncludeObject], &vResult, 2, pv);
		return GetIntFromVariantClear(&vResult);
	}
	if (m_bsFilter) {
		HRESULT hr = S_OK;
		STRRET strret;
		if SUCCEEDED(psf->GetDisplayNameOf(pidlItem, SHGDN_NORMAL, &strret)) {
			BSTR bsFile;
			if SUCCEEDED(StrRetToBSTR(&strret, pidlItem, &bsFile)) {
				if (!PathMatchSpec(bsFile, m_bsFilter)) {
					hr = S_FALSE;
				}
				::SysFreeString(bsFile);
			}
		}
		return hr;
	}
	return S_OK;
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
	HRESULT hr = teGetDispId(methodSB, _countof(methodSB), g_maps[MAP_SB], *rgszNames, rgDispId, true);
	if (hr != S_OK && m_pDSFV) {
		return m_pDSFV->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
	}
	return hr;
}

VOID AddColumnData(LPWSTR pszColumns, LPWSTR pszName, int nWidth)
{
	WCHAR szName[MAX_COLUMN_NAME_LEN + 3];
	lstrcpyn(szName, pszName, MAX_COLUMN_NAME_LEN);
	PathQuoteSpaces(szName);
	swprintf_s(&pszColumns[lstrlen(pszColumns)], MAX_COLUMN_NAME_LEN + 20, L" %s %d", szName, nWidth);
}

BSTR CteShellBrowser::GetColumnsStr()
{
	WCHAR szColumns[8192];
	szColumns[0] = 0;
	szColumns[1] = 0;

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
	else {
#ifdef _2000XP
		HWND hHeader = ListView_GetHeader(m_hwndLV);
		if (hHeader) {
			UINT nHeader = Header_GetItemCount(hHeader);
			if (nHeader) {
				int *piOrderArray = new int[nHeader];
				ListView_GetColumnOrderArray(m_hwndLV, nHeader, piOrderArray);
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
								AddColumnData(szColumns, bs, sd.cxChar * g_nCharWidth);
								::SysFreeString(bs);
							}
						}
					}
				}
				pSF2->Release();
			}
		}
#endif
	}
	return SysAllocString(&szColumns[1]);
}

VOID CteShellBrowser::GetDefaultColumns()
{
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
}

STDMETHODIMP CteShellBrowser::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		int nIndex;
		TC_ITEM tcItem;

		if ((m_bEmpty || m_nUnload == 1 || m_nUnload == 9) && dispIdMember >= 0x10000100 && dispIdMember) {
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
					for (int i = SB_Count; i--;) {
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
							try {
								teSetObject(&v, m_ppLog[m_nLogCount - i]);
								pFolderItems->ItemEx(-1, NULL, &v);
								VariantClear(&v);
							}
							catch (...) {
							}
						}
						pFolderItems->m_nIndex = m_nLogCount - 1 - m_nLogIndex;
						teSetObjectRelease(pVarResult, pFolderItems);
					}
				}
				return S_OK;
			//Navigate2Ex
			case 0x20000007:
				if (nArg >= 3) {
					FolderItem *pFolderItem = NULL;
					IUnknown *punk;
					if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
						punk->QueryInterface(IID_PPV_ARGS(&pFolderItem));
					}
					if (!pFolderItem && !m_pShellView) {
						GetFolderItemFromPidl(&pFolderItem, g_pidls[CSIDL_DESKTOP]);
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
							GetIDListFromObject(punk, &pidl);
						}
						if (SUCCEEDED(Navigate2(pFolderItem, GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), param, pFolderItems, m_pFolderItem, this)) && m_bVisible) {
							Show(TRUE);
						}
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
					teSetObject(pVarResult, pSB ? pSB : this);
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
					SetFolderFlags();
				}
				if (pVarResult) {
					pVarResult->lVal = m_param[SB_FolderFlags];
					pVarResult->vt = VT_I4;
				}
				return S_OK;
			//CurrentViewMode
			case 0x10000010:
				IFolderView2 *pFV2;
				if (nArg >= 1 && m_pShellView && SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2)))) {
					m_param[SB_ViewMode] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					m_param[SB_IconSize] = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
					pFV2->SetViewModeAndIconSize(static_cast<FOLDERVIEWMODE>(m_param[SB_ViewMode]), m_param[SB_IconSize]);
					pFV2->Release();
				}
				else if (nArg >= 0) {
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
						SetFolderFlags();
					}
					pVarResult->lVal = m_param[SB_ViewMode];
					pVarResult->vt = VT_I4;
				}
				return S_OK;
			//IconSize
			case 0x10000011:
				if (m_pShellView && m_bIconSize) {
					GetViewModeAndIconSize(TRUE);
					if (nArg >= 0) {
						int iImageSize = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
						IFolderView2 *pFV2;
						if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
							if (iImageSize != m_param[SB_IconSize]) {
								pFV2->SetViewModeAndIconSize(static_cast<FOLDERVIEWMODE>(m_param[SB_ViewMode]), m_param[SB_IconSize]);
							}
							pFV2->Release();
						}
#ifdef _2000XP
						else {
							m_param[SB_IconSize] = iImageSize;
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
					if (m_pExplorerBrowser) {
						m_pExplorerBrowser->SetOptions(static_cast<EXPLORER_BROWSER_OPTIONS>(m_param[SB_Options] | EBO_NOTRAVELLOG));
					}
				}
				if (pVarResult) {
					if (m_pExplorerBrowser) {
						m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
					}
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
					if (wFlags & DISPATCH_METHOD) {
						if (m_pDSFV) {
							IShellFolderViewDual3 *pSFVD3;
							if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD3))) {
								pSFVD3->FilterView(v.bstrVal);
								pSFVD3->Release();
							}
#ifdef _2000XP
							else {
								::SysReAllocString(&m_bsFilter, v.bstrVal);
								Refresh(TRUE);
							}
#endif
						}
					}
					else if (lstrcmpi(m_bsFilter, v.bstrVal)) {
						::SysReAllocString(&m_bsFilter, v.bstrVal);
					}
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
				}
				return S_OK;
			//TreeView
			case 0x10000021:
				teSetObject(pVarResult, m_pTV);
				return S_OK;
			//Parent
			case 0x10000024:
				IDispatch *pdisp;
				if SUCCEEDED(get_Parent(&pdisp)) {
					teSetObjectRelease(pVarResult, pdisp);
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
					if (nArg >= 0) {
						VARIANT vText;
						teVariantChangeType(&vText, &pDispParams->rgvarg[nArg], VT_BSTR);
						SetTitle(vText.bstrVal, nIndex);
						VariantClear(&vText);
					}
					if (pVarResult) {
						BSTR bsText = SysAllocStringLen(NULL, MAX_PATH);
						bsText[0] = NULL;
						tcItem.pszText = bsText;
						tcItem.mask = TCIF_TEXT;
						tcItem.cchTextMax = MAX_PATH;
						TabCtrl_GetItem(m_pTabs->m_hwnd, nIndex, &tcItem);
						int nCount = lstrlen(tcItem.pszText);
						int j = 0;
						WCHAR c = NULL;
						for (int i = 0; i < nCount; i++) {
							bsText[j] = bsText[i];
							if (c != '&' || bsText[i] != '&') {
								j++;
								c = bsText[i];
							}
							else {
								c = NULL;
							}
						}
						bsText[j] = NULL;
						pVarResult->vt = VT_BSTR;
						pVarResult->bstrVal = SysAllocString(bsText);
						SysFreeString(bsText);
					}
				}
				return S_OK;
			//Suspend
			case 0x10000033:
				Suspend(nArg >= 0 && GetIntFromVariant(&pDispParams->rgvarg[nArg]));
				return S_OK;
			//Items
			case 0x10000040:
				FolderItems *pFolderItems;
				if SUCCEEDED(Items(nArg >= 0 ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : SVGIO_ALLVIEW | SVGIO_FLAG_VIEWORDER, &pFolderItems)) {
					teSetObjectRelease(pVarResult, pFolderItems);
				}
				return S_OK;
			//SelectedItems
			case 0x10000041:
				if SUCCEEDED(Items(SVGIO_SELECTION, &pFolderItems)) {
					teSetObjectRelease(pVarResult, pFolderItems);
				}
				return S_OK;
			//ShellFolderView
			case 0x10000050:
				teSetObject(pVarResult, m_pDSFV);
				return S_OK;
			//hwndList
			case 0x10000102:
				SetVariantLL(pVarResult, (LONGLONG)m_hwndLV);
				return S_OK;
			//hwndView
			case 0x10000103:
				if (pVarResult && m_pShellView) {
					HWND hwndDV = m_hwnd;
					m_pShellView->GetWindow(&hwndDV);
					SetVariantLL(pVarResult, (LONGLONG)hwndDV);
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
				SetActive(TRUE);
				return S_OK;
			//HitTest (Screen coordinates)
			case 0x10000107:
				if (nArg >= 0 && pVarResult) {
					LVHITTESTINFO info;
					GetPointFormVariant(&info.pt, &pDispParams->rgvarg[nArg]);
					UINT flags = nArg >= 1 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : LVHT_ONITEM;
					pVarResult->lVal = (int)DoHitTest(this, info.pt, flags);
					pVarResult->vt = VT_I4;
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
						}
						else if (g_pAutomation) {
							POINT pt = info.pt;
							IUIAutomationElement *pElement, *pElement2;
							if SUCCEEDED(g_pAutomation->ElementFromPoint(pt, &pElement)) {
								if SUCCEEDED(pElement->GetCurrentPropertyValue(g_PID_ItemIndex, pVarResult)) {
									pVarResult->lVal--;
								}
								if (pVarResult->lVal < 0) {
									IUIAutomationTreeWalker* pWalker = NULL;
									if SUCCEEDED(g_pAutomation->get_ControlViewWalker(&pWalker)) {
										if SUCCEEDED(pWalker->GetParentElement(pElement, &pElement2)) {
											if SUCCEEDED(pElement2->GetCurrentPropertyValue(g_PID_ItemIndex, pVarResult)) {
												pVarResult->lVal--;
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
				}
				return S_OK;
			//DropTarget
			case 0x10000108:
				CteDropTarget *pCDT;
				if (m_pDropTarget) {
					pCDT = new CteDropTarget(static_cast<IDropTarget *>(this), NULL);
				}
				else {
					pCDT = new CteDropTarget(static_cast<IDropTarget *>(GetProp(m_hwndLV, L"OleDropTargetInterface")), NULL);
				}
				if (pCDT) {
					teSetObjectRelease(pVarResult, pCDT);
				}
				return S_OK;
			//Columns
			case 0x10000109:
				if (nArg >= 0) {
					VARIANT v;
					teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
					int nCount = 0;
					LPTSTR *lplpszArgs = NULL;
					TEmethod *methodArgs = NULL;
					int *pi = NULL;
					if (v.bstrVal) {
						lplpszArgs = CommandLineToArgvW(v.bstrVal, &nCount);
						int nAllWidth = 0;
						nCount /= 2;
						methodArgs = new TEmethod[nCount];
						for (int i = nCount; i-- > 0;) {
							methodArgs[i].name = lplpszArgs[i * 2];
							int n;
							if (StrToIntEx(lplpszArgs[i * 2 + 1], STIF_DEFAULT, &n)) {
								nAllWidth += n;
							}
							else {
								n = -1;
								nAllWidth += 100;
							}
							methodArgs[i].id = n;
						}
						if (!nCount || nAllWidth > nCount + 4) {
							pi = SortTEMethod(methodArgs, nCount);
						}
						else {
							if (methodArgs) {
								delete [] methodArgs;
							}
							LocalFree(lplpszArgs);
							VariantClear(&v);
							return S_OK;
						}
					}
					VariantClear(&v);
					IColumnManager *pColumnManager;
					if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager))) {
						UINT uCount;
						if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_ALL, &uCount)) {
							PROPERTYKEY *propKey = new PROPERTYKEY[uCount];
							if SUCCEEDED(pColumnManager->GetColumns(CM_ENUM_ALL, propKey, uCount)) {
								VARIANTARG *pv = GetNewVARIANT(3);
								teExecMethod(g_pJS, L"_o", &pv[2], 0, NULL);
								IDispatch *pdispColumns;
								if (GetDispatch(&pv[2], &pdispColumns)) {
									CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_NAME };
									for (int i = uCount; i-- > 0;) {
										if SUCCEEDED(pColumnManager->GetColumnInfo(propKey[i], &cmci)) {
											if (cmci.wszName[0]) {
												pv[1].vt = VT_BSTR;
												pv[1].bstrVal = ::SysAllocString(cmci.wszName);
												pv[0].vt = VT_I4;
												pv[0].lVal = i;
												teExecMethod(g_pJS, L"_p", NULL, -3, pv);
												VariantClear(&pv[1]);
											}
										}
									}
									int k = 0;
									PROPERTYKEY *propKey2 = new PROPERTYKEY[uCount];
									int *pnWidth = new int[uCount];
									for (int i = 0; i < nCount; i++) {
										if (teGetProperty(pdispColumns, methodArgs[i].name, &pv[0]) == S_OK) {
											pnWidth[k] = methodArgs[i].id;
											propKey2[k++] = propKey[GetIntFromVariantClear(&pv[0])];
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
									pdispColumns->Release();
								}
								teClearVariantArgs(3, pv);
							}
							delete [] propKey;
						}
						pColumnManager->Release();
					}
					else {
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
							for (int i = uCount; i-- > 0;) {
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
							TEmethod *methodColumns = new TEmethod[nCur];
							for (int i = nCur; i-- > 0;) {
								methodColumns[i].name = lplpszColumns[i * 2];
							}
							BOOL bDiff = nCount != nCur;
							if (!bDiff) {
								int *piCur = SortTEMethod(methodColumns, nCur);
								for (int i = nCur; i-- > 0;) {
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
								IShellView *pPreviousView = m_pShellView;
								FOLDERSETTINGS fs;
								pPreviousView->GetCurrentInfo(&fs);
								m_param[SB_ViewMode] = fs.ViewMode;
								if (!ILIsEqual(m_pidl, g_pidlResultsFolder)) {
									Show(FALSE);
									if SUCCEEDED(CreateViewWindowEx(pPreviousView)) {
										pPreviousView->DestroyViewWindow();
										pPreviousView->Release();
										pPreviousView = NULL;
										ResetPropEx();
										GetShellFolderView();
									}
									Show(TRUE);
									SetPropEx();
								}
							}
							//Columns Order and Width;
							HWND hwndDV;
							if (m_pShellView->GetWindow(&hwndDV) == S_OK) {
								HWND hwndLV = FindWindowEx(hwndDV, 0, WC_LISTVIEW, NULL);
								if (hwndLV) {
									HWND hHeader = ListView_GetHeader(hwndLV);
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
												for (int i = nHeader; i-- > 0;) {
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
																			if (teStrSameIFree(bs, szText)) {
																				nWidth = sd.cxChar * g_nCharWidth;
																				break;
																			}
																		}
																	}
																}
															}
															ListView_SetColumnWidth(hwndLV, i, nWidth);
														}
													}
												}
												ListView_SetColumnOrderArray(hwndLV, nHeader, piOrderArray);
											} catch (...) {
											}
											delete [] piOrderArray;
											SetRedraw(true);
										}
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
			//ItemCount
			case 0x10000110:
				if (pVarResult && m_pShellView) {
					IFolderView *pFV;
					if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
						if SUCCEEDED(pFV->ItemCount(nArg >= 0 ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : SVGIO_ALLVIEW, &pVarResult->intVal)) {
							pVarResult->vt = VT_I4;
						}
						pFV->Release();
					}
				}
				return S_OK;
			//Refresh
			case 0x10000206:
				Refresh(TRUE);
				return S_OK;
			//Refresh2Ex
			case 0x20000206:
				if (nArg >= 0 && pDispParams->rgvarg[nArg].vt != VT_EMPTY) {
					Refresh(FALSE);
				}
				return S_OK;
			//ViewMenu
			case 0x10000207:
				IContextMenu *pCM;
				CteContextMenu *pCCM;
				CteServiceProvider *pSP;
				pCM = NULL;
				pCCM = NULL;
				pSP = NULL;
				try {
					if SUCCEEDED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&pCM))) {
						pSP = new CteServiceProvider(reinterpret_cast<IUnknown *>(this), m_pShellView);
						IUnknown_SetSite(pCM, pSP);
						CteContextMenu *pCCM = new CteContextMenu(pCM, NULL);
						pCCM->m_pShellBrowser = this;
						teSetObject(pVarResult, pCCM);
					}
				}
				catch(...) {
				}
				if (pCCM) {
					pCCM->Release();
				}
				if (pSP) {
					pSP->Release();
				}
				if (pCM) {
					pCM->Release();
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
						char *pc = GetpcFromVariant(&pDispParams->rgvarg[nArg - 1]);
						if (pc) {
							IFolderView *pFV;
							if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
								LPITEMIDLIST pidl;
								if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
									hr = pFV->GetItemPosition(ILFindLastID(pidl), (LPPOINT)pc);
									teCoTaskMemFree(pidl);
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
						char *pc = GetpcFromVariant(&pDispParams->rgvarg[nArg - 2]);
						if (pc) {
							IShellView2 *pSV2;
							if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSV2))) {
								LPITEMIDLIST pidl;
								if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
									hr = pSV2->SelectAndPositionItem(ILFindLastID(pidl), GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), (LPPOINT)pc);
									teCoTaskMemFree(pidl);
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
			//SelectItem
			case 0x10000280:
				if (nArg >= 1) {
					SelectItem(&pDispParams->rgvarg[nArg], GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]));
				}
				return S_OK;
			//FocusedItem
			case 0x10000281:
				FolderItem *pFolderItem;
				if (get_FocusedItem(&pFolderItem) == S_OK) {
					teSetObjectRelease(pVarResult, pFolderItem);
				}
				return S_OK;
			//GetFocusedItem
			case 0x10000282:
				if (pVarResult) {
					GetFocusedIndex(&pVarResult->intVal);
					pVarResult->vt = VT_I4;
				}
				return S_OK;
			//AddItem
			case 0x10000501:
				if (nArg >= 0) {
					IFolderView *pFV = NULL;
					LPITEMIDLIST pidl;
					LPITEMIDLIST pidlChild = NULL;
					if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
						if (!ILIsEmpty(pidl)) {
							IResultsFolder *pRF;
							if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) && SUCCEEDED(pFV->GetFolder(IID_PPV_ARGS(&pRF)))) {
								pRF->AddIDList(pidl, &pidlChild);
								pRF->Release();
							}
#ifdef _2000XP
							else {
								BSTR bs;
								if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORADDRESSBAR)) {
									if (bs[0]) {
										IShellFolderView *pSFV;
										if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
											pidlChild = teILCreateResultsXP(pidl);
											UINT ui;
											pSFV->AddObject(pidlChild, &ui);
											pSFV->Release();
										}
									}
									SysFreeString(bs);
								}
							}
#endif
							if (pFV) {
								pFV->Release();
							}
							if (pidlChild) {
								LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidlChild);
								FolderItem *pFolderItem;
								if (GetFolderItemFromPidl(&pFolderItem, pidlFull)) {
									teSetObjectRelease(pVarResult, pFolderItem);
								}
								teCoTaskMemFree(pidlFull);
								CoTaskMemFree(pidlChild);
							}
						}
					}
					teCoTaskMemFree(pidl);
				}
				return S_OK;
			//RemoveItem
			case 0x10000502:
				hr = E_NOTIMPL;
				if (nArg >= 0) {
					LPITEMIDLIST pidl;
					if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
						IFolderView *pFV = NULL;
						IResultsFolder *pRF;
						if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) && SUCCEEDED(pFV->GetFolder(IID_PPV_ARGS(&pRF)))) {
							hr = pRF->RemoveIDList(pidl);
							pRF->Release();
						}
#ifdef _2000XP
						else {
							IShellFolderView *pSFV;
							if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
								LPITEMIDLIST pidlChild = teILCreateResultsXP(pidl);
								UINT ui;
								hr = pSFV->RemoveObject(pidlChild, &ui);
								teCoTaskMemFree(pidlChild);
								pSFV->Release();
							}
						}
#endif
						if (pFV) {
							pFV->Release();
						}
						teCoTaskMemFree(pidl);
					}
				}
				if (pVarResult) {
					pVarResult->lVal = hr;
					pVarResult->vt = VT_I4;
				}
				return S_OK;
			//
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
			//DIID_DShellFolderViewEvents
			case DISPID_SELECTIONCHANGED:
				SetStatusTextSB(NULL);
				return DoFunc(TE_OnSelectionChanged, this, S_OK);
			case DISPID_FILELISTENUMDONE:
				if (m_bNavigateComplete) {
					m_bNavigateComplete = FALSE;
					OnNavigationComplete(NULL);
				}
				return S_OK;
			case DISPID_INITIALENUMERATIONDONE:
			case DISPID_VIEWMODECHANGED:
				SetFolderFlags();
				return S_OK;
			default:
				if (dispIdMember >= START_OnFunc && dispIdMember < START_OnFunc + Count_SBFunc) {
					if (wFlags == DISPATCH_METHOD) {
						if (m_pOnFunc[dispIdMember - START_OnFunc]) {
							Invoke5(m_pOnFunc[dispIdMember - START_OnFunc], DISPID_VALUE, wFlags, pVarResult, - nArg - 1, pDispParams->rgvarg);
						}
						return S_OK;
					}
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
					if (m_pOnFunc[dispIdMember - START_OnFunc]) {
						teSetObject(pVarResult, m_pOnFunc[dispIdMember - START_OnFunc]);
					}
					return S_OK;
				}
		}
		if (m_pDSFV) {
			return m_pDSFV->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
	} catch (...) {
		return teException(pExcepInfo, puArgErr, L"FolderView", methodSB, dispIdMember);
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
/*/// Unstable (Windows2000)
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		pSFVD->SelectedItems(ppid);
		pSFVD->Release();
		return S_OK;
	}
//*/
	return Items(SVGIO_SELECTION, ppid);
}

HRESULT CteShellBrowser::Items(UINT uItem, FolderItems **ppid)
{
	FolderItems *pItems = NULL;
	IDataObject *pDataObj = NULL;
	if (m_pShellView) {
#ifdef _2000XP
		if (!g_bUpperVista && ListView_GetSelectedCount(m_hwndLV) > 1) {
			BOOL bResultsFolder = ILIsEqual(m_pidl, g_pidlResultsFolder);
			if (bResultsFolder || (uItem & (SVGIO_SELECTION | SVGIO_FLAG_VIEWORDER)) == (SVGIO_SELECTION | SVGIO_FLAG_VIEWORDER)) {
				CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL, true);
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
					}
					else {
						UINT uCount = 0;
						pSFV->GetObjectCount(&uCount);
						for (UINT u = 0; u < uCount; u++) {
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
	*ppid = new CteFolderItems(pDataObj, pItems, false);
	if (pDataObj) {
		pDataObj->Release();
	}
	return S_OK;
}


#ifdef _2000XP
VOID CteShellBrowser::AddPathXP(CteFolderItems *pFolderItems, IShellFolderView *pSFV, int nIndex, BOOL bResultsFolder)
{
	try {
		LPCITEMIDLIST pidl;
		if SUCCEEDED(pSFV->GetObjectW(const_cast<LPITEMIDLIST *>(&pidl), nIndex)) {
			VARIANT v;
			VariantInit(&v);
			if (bResultsFolder) {
				STRRET strret;
				m_pSF2->GetDisplayNameOf(pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &strret);
				if SUCCEEDED(StrRetToBSTR(&strret, pidl, &v.bstrVal)) {
					v.vt = VT_BSTR;
				}
			}
			else {
				FolderItem *pid;
				LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidl);
				if (GetFolderItemFromPidl(&pid, pidlFull)) {
					teSetObject(&v, pid);
					pid->Release();
				}
				::CoTaskMemFree(pidlFull);
			}
			if (v.vt != VT_EMPTY) {
				pFolderItems->ItemEx(-1, NULL, &v);
				VariantClear(&v);
			}
		}
	}
	catch (...) {}
}
#endif

STDMETHODIMP CteShellBrowser::get_FocusedItem(FolderItem **ppid)
{
	HRESULT hr = E_NOTIMPL;
	*ppid = NULL;
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
	LPITEMIDLIST pidl;
	GetPidlFromVariant(&pidl, pvfi);
	return SelectItemEx(&pidl, dwFlags);
}

HRESULT CteShellBrowser::SelectItemEx(LPITEMIDLIST *ppidl, int dwFlags)
{
	HRESULT hr = E_FAIL;
	if (m_pShellView) {
		if (m_hwndLV && (dwFlags & SVSI_FOCUSED)) {// bug fix
	#ifdef _2000XP
			if (g_bUpperVista) {
	#endif
				int nFocused = ListView_GetNextItem(m_hwndLV, -1, LVNI_ALL | LVNI_FOCUSED);
				if (nFocused >= 0) {
					ListView_SetItemState(m_hwndLV, nFocused, 0, LVIS_FOCUSED);
				}
	#ifdef _2000XP
			}
	#endif
		}
		hr = m_pShellView->SelectItem(ILFindLastID(*ppidl), dwFlags);
	}
	teILFreeClear(ppidl);
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
		
		for (i = TabCtrl_GetItemCount(m_pTabs->m_hwnd); i-- > 0;) {
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

	IFolderViewOptions *pOptions;
	if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
		FOLDERVIEWOPTIONS fvo = teGetFolderViewOptions(const_cast<LPITEMIDLIST>(pidlFolder), m_param[SB_ViewMode]);
		pOptions->SetFolderViewOptions(FVO_VISTALAYOUT, fvo);
		pOptions->Release();
	}
	LPITEMIDLIST pidlPrevius = m_pidl;
	m_pidl = ::ILClone(pidlFolder);
	FolderItem *pPrevious = m_pFolderItem;
	if (m_pFolderItem1) {
		m_pFolderItem = m_pFolderItem1;
		m_pFolderItem1 = NULL;
	}
	else {
		GetFolderItemFromPidl(&m_pFolderItem, m_pidl);
	}

	HRESULT hr = OnBeforeNavigate(pPrevious, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
	if FAILED(hr) {
		teCoTaskMemFree(m_pidl);
		m_pidl = pidlPrevius;
		FolderItem *pid = m_pFolderItem;
		m_pFolderItem = NULL;
		if (pPrevious) {
			pPrevious->QueryInterface(IID_PPV_ARGS(&m_pFolderItem));
		}
		if (hr == E_ACCESSDENIED) {
			BrowseObject2(pid, SBSP_NEWBROWSER | SBSP_ABSOLUTE);
		}
		pid ->Release();
		return hr;
	}
	ResetPropEx();
	//History / Management
	SetHistory(NULL, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
	teCoTaskMemFree(pidlPrevius);
	return hr;
}

STDMETHODIMP CteShellBrowser::OnViewCreated(IShellView *psv)
{
	if (psv) {
		psv->QueryInterface(IID_PPV_ARGS(&m_pShellView));
	}
	GetShellFolderView();
	if (m_pExplorerBrowser) {
		IOleWindow *pOleWindow;
		if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
			pOleWindow->GetWindow(&m_hwnd);
			pOleWindow->Release();
		}
	}
	m_bNavigateComplete = !m_pExplorerBrowser;
	SetPropEx();
#ifdef _2000XP
	if (!g_bUpperVista) {
		if (ILIsEqual(m_pidl, g_pidlResultsFolder)) {
			UINT i = 0;
			SHCOLUMNID scid;
			while (MapColumnToSCID(i, &scid) == S_OK) {
				if (scid.pid == 2 && IsEqualFMTID(scid.fmtid, FMTID_Storage)) {
					m_nFolderName = i;
					break;
				}
				i++;
			}
		}
	}
	if (g_bCharWidth) {
		if (m_hwndLV) {
			HWND hHeader = ListView_GetHeader(m_hwndLV);
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
							if (teStrSameIFree(bs, hdi.pszText)) {
								if (sd.cxChar) {
									g_nCharWidth = hdi.cxy / sd.cxChar;
								}
								break;
							}
						}
					}
					g_bCharWidth = false;
					pSF2->Release();
				}
			}
		}
	}
#endif
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
			if (m_pFolderItem && SUCCEEDED(m_pFolderItem->get_Name(&bstr))) {
				SetTitle(bstr, i);
				::SysFreeString(bstr);
			}
		}
	}
	SetStatusTextSB(NULL);
	return S_OK;
}

STDMETHODIMP CteShellBrowser::OnNavigationComplete(PCIDLIST_ABSOLUTE pidlFolder)
{
	IFolderView2 *pFV2;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		pFV2->SetViewModeAndIconSize(static_cast<FOLDERVIEWMODE>(m_param[SB_ViewMode]), m_param[SB_IconSize]);
		pFV2->Release();
	}
	m_bIconSize = TRUE;
	if ((m_param[SB_TreeAlign] & 2) && m_pTV->m_bSetRoot) {
		if (m_pTV->Create()) {
			m_pTV->SetRoot();
		}
	}
	SetRedraw(TRUE);
	SetActive(FALSE);
	DoFunc(TE_OnNavigateComplete, this, E_NOTIMPL);
	SetTimer(m_hwnd, (UINT_PTR)this, 100, teNavigateCompleteProc);
	SetFolderFlags();
	return S_OK;
}

STDMETHODIMP CteShellBrowser::OnNavigationFailed(PCIDLIST_ABSOLUTE pidlFolder)
{
	SetRedraw(TRUE);
	BSTR bs = NULL;
	m_pFolderItem && m_pFolderItem->get_Path(&bs);
	Error(&bs);
	teSysFreeString(&bs);
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
		pv[1].vt = VT_BSTR;
		pv[1].bstrVal = ::SysAllocStringLen(NULL, 38);
		StringFromGUID2(ep, pv[1].bstrVal, 39);

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
	if (g_bIsXP && IsEqualIID(riid, IID_IShellView)) {
		//only XP
		if (m_pSFVCB) {
			m_pSFVCB->Release();
			m_pSFVCB = NULL;
		}
		IShellView *pSV;
		SFV_CREATE sfvc;
		sfvc.cbSize = sizeof(SFV_CREATE);
		sfvc.pshf = static_cast<IShellFolder *>(this);
		sfvc.psfvcb = static_cast<IShellFolderViewCB *>(this);
		sfvc.psvOuter = NULL;
		if SUCCEEDED(m_pSF2->CreateViewObject(hwndOwner, IID_PPV_ARGS(&pSV))) {
			IShellFolderView *pSFV;
			if (pSV->QueryInterface(IID_PPV_ARGS(&pSFV)) == S_OK) {
				pSFV->SetCallback(NULL, &m_pSFVCB);
				pSFV->Release();
			}
			if (m_pSFVCB) {
				pSV->Release();
			}
			else {
				*ppv = pSV;
				return S_OK;
			}
		}
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
	if (pidl && iColumn == m_nFolderName) {
		if SUCCEEDED(m_pSF2->GetDisplayNameOf(pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &psd->str)) {
			PathRemoveFileSpec(psd->str.pOleStr);
			return S_OK;
		}
	}
	HRESULT hr = m_pSF2->GetDetailsOf(pidl, iColumn, psd);

	if (m_nColumns) {
		if (iColumn < m_nColumns) {
			if (m_pColumns[iColumn].nWidth >= 0) {
				psd->cxChar = m_pColumns[iColumn].nWidth / g_nCharWidth;
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
	if (uMsg != 54 && uMsg != 78 && m_pSFVCB) {
		return m_pSFVCB->MessageSFVCB(uMsg, wParam, lParam);
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
	HRESULT hr = DragSub(TE_OnDragEnter, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		*pdwEffect = dwEffect;
		if (m_pDropTarget->DragEnter(pDataObj, grfKeyState, pt, pdwEffect) == S_OK) {
			hr = S_OK;
		}
	}
	if (hr == S_OK && *pdwEffect) {
		m_pDragItems->Regenerate();
	}
	if (g_pDropTargetHelper) {
		g_pDropTargetHelper->DragEnter(m_hwnd, pDataObj, (LPPOINT)&pt, *pdwEffect);
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	DWORD dwEffect2 = *pdwEffect;
	HRESULT hr2 = E_NOTIMPL;
	HRESULT hr = DragSub(TE_OnDragOver, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		hr2 = m_pDropTarget->DragOver(grfKeyState, pt, &dwEffect2);
	}
	if (hr == S_OK) {
		if (*pdwEffect != dwEffect2) {
			if (g_pDropTargetHelper) {
				g_pDropTargetHelper->DragOver((LPPOINT)&pt, *pdwEffect);
			}
		}
	}
	else {
		*pdwEffect = dwEffect2;
		hr = hr2;
	}
	m_grfKeyState = grfKeyState;
	if (hr == S_OK && m_hwndLV && GetTickCount() - g_dwTickScroll > DRAG_INTERVAL
#ifdef _2000XP
	&& g_bUpperVista
#endif
	) {
		RECT rc;
		ScreenToClient(m_hwndLV, (LPPOINT)&pt);
		GetClientRect(m_hwndLV, &rc);
		if (pt.x < DRAG_SCROLL) {
			SendMessage(m_hwndLV, WM_HSCROLL, SB_LINELEFT, 0);
		}
		else if (pt.x > rc.right - DRAG_SCROLL) {
			SendMessage(m_hwndLV, WM_HSCROLL, SB_LINERIGHT, 0);
		}
		if (pt.y < DRAG_SCROLL) {
			SendMessage(m_hwndLV, WM_VSCROLL, SB_LINEUP, 0);
		}
		else if (pt.y > rc.bottom - DRAG_SCROLL) {
			SendMessage(m_hwndLV, WM_VSCROLL, SB_LINEDOWN, 0);
		}
	}
	return hr;
}

STDMETHODIMP CteShellBrowser::DragLeave()
{
	if (m_DragLeave == E_NOT_SET) {
		if (g_pDropTargetHelper) {
			g_pDropTargetHelper->DragLeave();
		}
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
	HRESULT hr = DragSub(TE_OnDrop, this, m_pDragItems, &m_grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		if (hr != S_OK) {
			*pdwEffect = dwEffect;
			hr = m_pDropTarget->Drop(pDataObj, m_grfKeyState, pt, pdwEffect);
		}
	}
	if (g_pDropTargetHelper) {
		g_pDropTargetHelper->Drop(pDataObj, (LPPOINT)&pt, *pdwEffect);
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

VOID CteShellBrowser::Suspend(BOOL bTree)
{
	SaveFocusedItemToHistory();
	teDoCommand(this, m_hwnd, WM_NULL, 0, 0);
	BOOL bVisible = m_bVisible;
	DestroyView(0);
	if (bTree && m_pTV) {
		m_pTV->Close();
		m_pTV->m_bSetRoot = true;
	}
	if (m_pidl && !IsSearchPath(m_pidl)) {
		if (m_pFolderItem) {
			m_pFolderItem->Release();
		}
		VARIANT v;
		VariantInit(&v);
		GetVarPathFromPidl(&v, m_pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
		m_pFolderItem = new CteFolderItem(&v);
		teILFreeClear(&m_pidl);
		VariantClear(&v);
	}
	m_nUnload = 9;
	Show(bVisible);
}

VOID CteShellBrowser::SetPropEx()
{
	if (m_pShellView->GetWindow(&m_hwndDV) == S_OK) {
		if (!m_DefProc) {
			SetWindowLongPtr(m_hwndDV, GWLP_USERDATA, (LONG_PTR)this);
			m_DefProc = SetWindowLongPtr(m_hwndDV, GWLP_WNDPROC, (LONG_PTR)TELVProc);
		}
		m_hwndLV = FindWindowEx(m_hwndDV, 0, WC_LISTVIEW, NULL);
		if (m_hwndLV) {
			if (!m_DefProc2) {
				SetWindowLongPtr(m_hwndLV, GWLP_USERDATA, (LONG_PTR)this);
				m_DefProc2 = SetWindowLongPtr(m_hwndLV, GWLP_WNDPROC, (LONG_PTR)TELVProc2);
			}
			HookDragDrop(g_dragdrop & 1);
			if (m_pExplorerBrowser) {
				if (m_param[SB_FolderFlags] & FWF_NOCLIENTEDGE) {
					SetWindowLong(m_hwnd, GWL_STYLE, GetWindowLong(m_hwnd, GWL_STYLE) & ~WS_BORDER);
				}
			}
			else if (!(m_param[SB_FolderFlags] & FWF_NOCLIENTEDGE)) {
				SetWindowLong(m_hwndLV, GWL_EXSTYLE, GetWindowLong(m_hwndLV, GWL_EXSTYLE) | WS_EX_CLIENTEDGE);
			}
		}
	}
}

VOID CteShellBrowser::ResetPropEx()
{
	if (m_DefProc) {
		SetWindowLongPtr(m_hwndDV, GWLP_WNDPROC, (LONG_PTR)m_DefProc);
		m_DefProc = NULL;
	}
	if (m_DefProc2) {
		SetWindowLongPtr(m_hwndLV, GWLP_WNDPROC, (LONG_PTR)m_DefProc2);
		m_DefProc2 = NULL;
	}
	if (m_pDropTarget) {
		RevokeDragDrop(m_hwndLV);
		RegisterDragDrop(m_hwndLV, m_pDropTarget);
		m_pDropTarget->Release();
		m_pDropTarget = NULL;
	}
}

void CteShellBrowser::Show(BOOL bShow)
{
	if (bShow) {
		if (!m_pTabs->m_bVisible) {
			return;
		}
		if (!g_nLockUpdate) {
			if (m_nUnload == 9) {
				m_nUnload = 2;
				BrowseObject(NULL, SBSP_RELATIVE | SBSP_WRITENOHISTORY | SBSP_REDIRECT);
				m_nUnload = 0;
			}
			else if (m_nUnload == 1 || !m_pShellView) {
				m_nUnload = 2;
				if SUCCEEDED(BrowseObject(NULL, SBSP_RELATIVE | SBSP_SAMEBROWSER)) {
					m_nUnload = 0;
				}
			}
		}
	}
	m_bVisible = bShow;
	if (m_pShellView) {
		if (bShow) {
			ShowWindow(m_hwnd, SW_SHOW);
			SetRedraw(TRUE);
			if (m_param[SB_TreeAlign] & 2) {
				ShowWindow(m_pTV->m_hwnd, SW_SHOW);
				BringWindowToTop(m_pTV->m_hwnd);
			}
			if (!SetActive(FALSE)) {
				m_pShellView->UIActivate(SVUIA_ACTIVATE_NOFOCUS);
			}
			BringWindowToTop(m_hwnd);
			ArrangeWindow();
			if (m_nUnload == 4) {
				Refresh(FALSE);
			}
		}
		else {
			SetRedraw(FALSE);
			m_pShellView->UIActivate(SVUIA_DEACTIVATE);
			MoveWindow(m_hwnd, -1, -1, 0, 0, false);
			ShowWindow(m_hwnd, SW_HIDE);
			MoveWindow(m_pTV->m_hwnd, -1, -1, 0, 0, false);
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
		ShowWindow(m_pTV->m_hwnd, SW_HIDE);
		Clear();
	}
}

VOID CteShellBrowser::DestroyView(int nFlags)
{
	KillTimer(m_hwnd, (UINT_PTR)this);
	ResetPropEx();
	if ((nFlags & 2) == 0 && m_pExplorerBrowser) {
		try {
			if (m_dwEventCookie) {
			  m_pExplorerBrowser->Unadvise(m_dwEventCookie);
			}
			IUnknown_SetSite(m_pExplorerBrowser, NULL);
			m_pServiceProvider->Release();
			m_pServiceProvider = NULL;
			Show(false);
			m_pExplorerBrowser->Destroy();
			m_pExplorerBrowser->Release();
		}
		catch (...) {
			g_nException = 0;
		}
		m_pExplorerBrowser = NULL;
	}
	if (m_pShellView) {
		if ((nFlags & 1) == 0) {
			Show(false);
			IUnknown_SetSite(m_pShellView, NULL);
			m_pShellView->DestroyViewWindow();
		}
		if (nFlags == 0 && m_pShellView) {
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
	IFolderView2 *pFV2;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		SORTCOLUMN col;
		if SUCCEEDED(pFV2->GetSortColumns(&col, 1)) {
			if (lpfnPSGetPropertyDescription) {
				IPropertyDescription *pdesc;
				if SUCCEEDED(lpfnPSGetPropertyDescription(col.propkey, IID_PPV_ARGS(&pdesc))) {
					LPWSTR pszName;
					if SUCCEEDED(pdesc->GetDisplayName(&pszName)) {
						*pbs = SysAllocString(pszName);
						teCoTaskMemFree(pszName);
					}
					pdesc->Release();
					if (col.direction < 0) {
						ToMinus(pbs);
					}
				}
			}
		}
		pFV2->Release();
		return;
	}
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
	HWND hHeader = ListView_GetHeader(m_hwndLV);
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
		teSysFreeString(pbs);
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
#ifdef _2000XP
	IShellFolderView *pSFV;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
		IShellFolder2 *pSF2;
		if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pSF2))) {
			SHELLDETAILS sd;
			HRESULT hr = E_FAIL;
			for (int i = 0; (pSF2->GetDetailsOf(NULL, i, &sd) == S_OK && FAILED(hr)); i++) {
				if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs1)) {
					if (teStrSameIFree(bs1, &bs[dir])) {
						hr = pSFV->Rearrange(i);
						if (dir && dir1) {
							hr = pSFV->Rearrange(i);
						}
					}
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
		IFolderView2 *pFV2;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
			hr = pFV2->SetRedraw(bRedraw);
			pFV2->Release();
		}
		else {
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

HRESULT CteShellBrowser::CreateViewWindowEx(IShellView *pPreviousView)
{
	HRESULT hr = E_FAIL;
	m_pShellView = NULL;
#ifdef _2000XP
	m_nFolderName = -1;
#endif
	RECT rc;
	if (m_pSF2) {
#ifdef _2000XP
		if SUCCEEDED(CreateViewObject(m_pTabs->m_hwndStatic, IID_PPV_ARGS(&m_pShellView)) && m_pShellView) {
			if (!g_bUpperVista && m_param[SB_ViewMode] == FVM_ICON && m_param[SB_IconSize] >= 96) {
				m_param[SB_ViewMode] = FVM_THUMBNAIL;
			}
#else
		if SUCCEEDED(m_pSF2->CreateViewObject(m_pTabs->m_hwndStatic, IID_PPV_ARGS(&m_pShellView)) && m_pShellView) {
#endif
			if (m_param[SB_ViewMode] == FVM_ICON && m_param[SB_IconSize] < 48
#ifdef _2000XP
				&& g_bUpperVista
#endif
			) {
				m_param[SB_IconSize] = 48;
			}
			try {
				if (::InterlockedIncrement(&m_nCreate) <= 1) {
					FOLDERSETTINGS fs = { m_param[SB_ViewMode], (m_param[SB_FolderFlags] | FWF_USESEARCHFOLDER) & ~FWF_NOENUMREFRESH };
					m_hwnd = NULL;
					hr = m_pShellView->CreateViewWindow(pPreviousView, &fs, static_cast<IShellBrowser *>(this), &rc, &m_hwnd);
				}
				else {
					hr = OLE_E_BLANK;
				}
			} catch (...) {
				hr = E_FAIL;
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

//CTE

CTE::CTE()
{
	m_cRef = 1;
	m_pDragItems = NULL;
	VariantInit(&m_Data);
	::ZeroMemory(m_param, sizeof(m_param));
	m_param[TE_Type] = CTRL_TE;
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
	return teGetDispId(methodTE, _countof(methodTE), g_maps[MAP_TE], *rgszNames, rgDispId, true);
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
		if (dispIdMember >= TE_METHOD && wFlags == DISPATCH_PROPERTYGET) {
			CteDispatch *pClone = new CteDispatch(this, 0);
			pClone->m_dispIdMember = dispIdMember;
			teSetObjectRelease(pVarResult, pClone);
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
						}//end_switch
					}
				}
				return S_OK;
			//Ctrls
			case TE_METHOD + 1006:
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
							if (g_pWebBrowser) {
								pMem = new CteMemory(sizeof(IDispatch *), NULL, VT_DISPATCH, 1, NULL);
								ppDisp = (IDispatch **)pMem->m_pc;
								g_pWebBrowser->QueryInterface(IID_PPV_ARGS(&ppDisp[0]));
							}
							break;
					}//end_switch
					if (!pMem) {
						pMem = new CteMemory(0, NULL, VT_DISPATCH, 0, NULL);
					}
					teSetObjectRelease(pVarResult, pMem);
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
				g_nReload = 0;
				return S_OK;
			//Reload
			case TE_METHOD + 1009:
				g_nReload = 1;
				if (g_nException-- <= 0 || (nArg >= 0 && GetIntFromVariant(&pDispParams->rgvarg[nArg]))) {
					SetTimer(g_hwndMain, TET_Reload, 100, teTimerProc);
					return S_OK;
				}
				if (g_pWebBrowser) {
					g_nLockUpdate = 0;
					ClearEvents();
					g_pWebBrowser->m_pWebBrowser->Refresh();
				}
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
					if SUCCEEDED(teCLSIDFromProgID(vClass.bstrVal, &clsid)) {
						HRESULT hr = E_FAIL;
						IUnknown *punk;
						if (dispIdMember == TE_METHOD + 1020) {
							hr = GetActiveObject(clsid, NULL, &punk);
						}
						if FAILED(hr) {
							hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&punk));
						}
						if SUCCEEDED(hr) {
							teSetObjectRelease(pVarResult, punk);
						}
					}
					VariantInit(&vClass);
				}
				return S_OK;
			//WindowsAPI
			case 1030:
				teSetObjectRelease(pVarResult, new CteWindowsAPI());
				return S_OK;
			//CommonDialog
			case 1131:
				teSetObjectRelease(pVarResult, new CteCommonDialog());
				return S_OK;
			//GdiplusBitmap
			case 1132:
				teSetObjectRelease(pVarResult, new CteGdiplusBitmap());
				return S_OK;
			//FolderItems
			case TE_METHOD + 1133:
				teSetObjectRelease(pVarResult, new CteFolderItems(NULL, NULL, true));
				return S_OK;
			//Object
			case TE_METHOD + 1134:
				if (pVarResult) {
					teExecMethod(g_pJS, L"_o", pVarResult, 0, NULL);
				}
				return S_OK;
			//Array
			case TE_METHOD + 1135:
				if (pVarResult) {
					GetNewArray(&pdisp);
					teSetObjectRelease(pVarResult, pdisp);
				}
				return S_OK;
			//Collection
			case TE_METHOD + 1136:
				if (pVarResult && nArg >= 0 && GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
					teSetObjectRelease(pVarResult, new CteDispatch(pdisp, 1));
					pdisp->Release();
				}
				return S_OK;
#ifdef _USE_TESTOBJECT
			//TestObj
			case 1200:
				punk = new CteTest();
				teSetObject(pVarResult, punk);
				punk->Release();
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
				if (nArg >= 4) {
					int i;
					i = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					switch (HIWORD(i)) {
						case 3:	//Tab
							CteTabs *pTC;
							pTC = GetNewTC();
							pTC->Init();
							i = _countof(pTC->m_param) - 1;
							for (i = nArg < i ? nArg : i; i >= 0; i--) {
								pTC->m_param[i] = GetIntFromVariantPP(&pDispParams->rgvarg[nArg - i], i);
							}
							if (pTC->Create()) {
								teSetObject(pVarResult, pTC);
								g_pTabs = pTC;
								pTC->Show(TRUE);
							}
							else {
								pTC->Release();
							}
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
								pSB->Close(true);
							}
						}
						MENUITEMINFO mii;
						::ZeroMemory(&mii, sizeof(MENUITEMINFO));
						mii.cbSize = sizeof(MENUITEMINFO);
						mii.fMask  = MIIM_SUBMENU;
						GetMenuItemInfo(g_hMenu, GetIntFromVariant(&pDispParams->rgvarg[nArg]), FALSE, &mii);

						HMENU hMenu = CreatePopupMenu();
						UINT fState = MFS_DISABLED;
						if (g_pTabs) {
							CteShellBrowser *pSB = g_pTabs->GetShellBrowser(g_pTabs->m_nIndex);
							if (pSB) {
								fState = MFS_ENABLED;
							}
						}
						teCopyMenu(hMenu, mii.hSubMenu, fState);
						SetVariantLL(pVarResult, (LONGLONG)hMenu);
					}
				}
				return S_OK;
			//CtrlFromWindow
			case TE_METHOD + 1070:
				if (nArg >= 0) {
					IDispatch *pdisp;
					if SUCCEEDED(ControlFromhwnd(&pdisp, (HWND)GetLLFromVariant(&pDispParams->rgvarg[nArg]))) {
						teSetObjectRelease(pVarResult, pdisp);
					}
				}
				return S_OK;
			//LockUpdate
			case TE_METHOD + 1080:
				if (::InterlockedIncrement(&g_nLockUpdate) == 1) {
					teSetRedraw(FALSE);
				}
				return S_OK;
			//UnlockUpdate
			case TE_METHOD + 1090:
				if (::InterlockedDecrement(&g_nLockUpdate) <= 0) {
					g_nLockUpdate = 0;
					teSetRedraw(TRUE);
					int i = MAX_TC;
					while (--i >= 0) {
						CteTabs *pTC = g_pTC[i];
						if (pTC) {
							if (pTC->m_bVisible) {
								CteShellBrowser *pSB = pTC->GetShellBrowser(pTC->m_nIndex);
								if (pSB && !pSB->m_bEmpty && pSB->m_nUnload & 5) {
									pSB->Show(TRUE);
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
				teSetObject(pVarResult, this);
				return S_OK;
			default:
				if (dispIdMember >= START_OnFunc && dispIdMember < START_OnFunc + Count_OnFunc) {
					if (wFlags == DISPATCH_METHOD) {
						if (g_pOnFunc[dispIdMember - START_OnFunc]) {
							Invoke5(g_pOnFunc[dispIdMember - START_OnFunc], DISPID_VALUE, wFlags, pVarResult, - nArg - 1, pDispParams->rgvarg);
						}
						return S_OK;
					}
					if (nArg >= 0) {
						if (g_pOnFunc[dispIdMember - START_OnFunc]) {
							g_pOnFunc[dispIdMember - START_OnFunc]->Release();
							g_pOnFunc[dispIdMember - START_OnFunc] = NULL;
						}
						if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
							punk->QueryInterface(IID_PPV_ARGS(&g_pOnFunc[dispIdMember - START_OnFunc]));
						}
					}
					if (g_pOnFunc[dispIdMember - START_OnFunc]) {
						teSetObject(pVarResult, g_pOnFunc[dispIdMember - START_OnFunc]);
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
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"external", methodTE, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDropTarget
STDMETHODIMP CTE::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);
	HRESULT hr = DragSub(TE_OnDragEnter, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (hr == S_OK && *pdwEffect) {
		m_pDragItems->Regenerate();
	}
	return hr;
}

STDMETHODIMP CTE::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	HRESULT hr = DragSub(TE_OnDragOver, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	m_grfKeyState = grfKeyState;
	return hr;
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
	HRESULT hr = DragSub(TE_OnDrop, this, m_pDragItems, &m_grfKeyState, pt, pdwEffect);
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
	*ppvObject = NULL;
	
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IInternetSecurityManager)) {
		*ppvObject = static_cast<IInternetSecurityManager *>(this);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
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
	*pPolicy = PathMatchSpec(pwszUrl, L"http*") ? URLPOLICY_DISALLOW : URLPOLICY_ALLOW;
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
	*ppvObject = NULL;
	
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_INewWindowManager)) {
		*ppvObject = static_cast<INewWindowManager *>(this);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
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
	return PathMatchSpec(pszUrl, L"http*") ? S_FALSE : S_OK;
}

// CteWebBrowser

CteWebBrowser::CteWebBrowser(HWND hwnd, WCHAR *szPath)
{
	m_cRef = 1;
	m_hwnd = hwnd;
	m_dwCookie = 0;
	m_pDragItems = NULL;
	m_pDropTarget = NULL;
	VariantInit(&m_Data);

	MSG        msg;
	RECT       rc;

	IOleObject *pOleObject;
	CLSID clsid;
	if (SUCCEEDED(CLSIDFromProgID(L"Shell.Explorer", &clsid)) &&
		SUCCEEDED(CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pWebBrowser))) &&
		SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleObject)))) {
		pOleObject->SetClientSite(static_cast<IOleClientSite *>(this));
		SetRectEmpty(&rc);
		pOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, &msg, static_cast<IOleClientSite *>(this), 0, g_hwndMain, &rc);
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
	/*	else if (IsEqualIID(riid, IID_IDocHostShowUI)) {
		*ppvObject = static_cast<IDocHostShowUI *>(this);
	}*/
	else if (IsEqualIID(riid, IID_IServiceProvider)) {
		*ppvObject = static_cast<IServiceProvider *>(new CteServiceProvider(reinterpret_cast<IUnknown *>(this), m_pWebBrowser));
		return S_OK;
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
	return teGetDispId(methodWB, 0, NULL, *rgszNames, rgDispId, true);
}

STDMETHODIMP CteWebBrowser::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
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
				IDispatch *pdisp;
				if SUCCEEDED(m_pWebBrowser->get_Application(&pdisp)) {
					teSetObjectRelease(pVarResult, pdisp);
				}
				return S_OK;
			//Document
			case 0x10000006:
				if SUCCEEDED(m_pWebBrowser->get_Document(&pdisp)) {
					teSetObjectRelease(pVarResult, pdisp);
				}
				return S_OK;
			//DIID_DWebBrowserEvents2
			case DISPID_DOWNLOADCOMPLETE:
				if (g_nReload == 1) {
					g_nReload = 2;
					SetTimer(g_hwndMain, TET_Reload, 2000, teTimerProc);
				}
				return S_OK;
			case DISPID_BEFORENAVIGATE2:
				if (nArg >= 6) {
					if (pDispParams->rgvarg[nArg - 6].vt == (VT_BYREF | VT_BOOL)) {
						VARIANT vURL;
						teVariantChangeType(&vURL, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
						VARIANT v = pDispParams->rgvarg[nArg - 1];
						VARIANT_BOOL bSB = VARIANT_FALSE; 
						FolderItem *pid = new CteFolderItem(&vURL);
						pid->get_IsFolder(&bSB);
						if (bSB || StrChrW(vURL.bstrVal, '/')) {
							*pDispParams->rgvarg[nArg - 6].pboolVal = VARIANT_TRUE;
							if (bSB) {
								CteShellBrowser *pSB = NULL;
								if (g_pTabs) {
									pSB = g_pTabs->GetShellBrowser(g_pTabs->m_nIndex);
									if (pSB) {
										pSB->BrowseObject2(pid, SBSP_SAMEBROWSER);
									}
								}
							}
						}
						pid->Release();
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
					for (int i = 3; i--;) {
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
/*///
			case DISPID_TITLECHANGE:
				if (nArg >= 0 && pDispParams->rgvarg[nArg].vt == VT_BSTR) {
					SetWindowText(g_hwndMain, pDispParams->rgvarg[nArg].bstrVal);
				}
				return S_OK;
//*/
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"WebBrowser", methodWB, dispIdMember);
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
/*// For check
	WCHAR* szKey = L"Software\\tablacus\\explorer";
    if (pchKey) {
		*pchKey = (LPOLESTR)CoTaskMemAlloc((lstrlen(szKey) + 1) * sizeof(WCHAR));
		if (*pchKey) {
			lstrcpy(*pchKey, szKey);
			return S_OK;
		}
    }
	return E_INVALIDARG;
//*/
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::GetDropTarget(IDropTarget *pDropTarget, IDropTarget **ppDropTarget)
{
	if (m_pDropTarget) {
		m_pDropTarget->Release();
		m_pDropTarget = NULL;
	}
	if (pDropTarget) {
		pDropTarget->QueryInterface(IID_PPV_ARGS(&m_pDropTarget));
	}
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

/*/// IDocHostShowUI
STDMETHODIMP CteWebBrowser::ShowMessage(HWND hwnd, LPOLESTR lpstrText, LPOLESTR lpstrCaption, DWORD dwType, LPOLESTR lpstrHelpFile, DWORD dwHelpContext, LRESULT *plResult)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteWebBrowser::ShowHelp(HWND hwnd, LPOLESTR pszHelpFile, UINT uCommand, DWORD dwData, POINT ptMouse, IDispatch *pDispatchObjectHit)
{
	return E_NOTIMPL;
}
//*/

//IDropTarget
STDMETHODIMP CteWebBrowser::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);

	DWORD dwEffect = *pdwEffect;
	HRESULT	hr = DragSub(TE_OnDragEnter, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		*pdwEffect = dwEffect;
		if (m_pDropTarget->DragEnter(pDataObj, grfKeyState, pt, pdwEffect) == S_OK) {
			hr = S_OK;
		}
	}
	if (hr == S_OK && *pdwEffect) {
		m_pDragItems->Regenerate();
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
		}
		else {
			m_dwEffect = DROPEFFECT_NONE;
		}
	}
	*pdwEffect |= m_dwEffect;
	m_grfKeyState = grfKeyState;
	return hr;
}

STDMETHODIMP CteWebBrowser::DragLeave()
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

STDMETHODIMP CteWebBrowser::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	HRESULT hr = E_NOTIMPL;
	DWORD dwEffect = *pdwEffect;
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
	IOleWindow *pOleWindow;
	if (m_pWebBrowser) {
		if SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
			pOleWindow->GetWindow(&hwnd);
			pOleWindow->Release();
		}
	}
	return hwnd;
}

/*//
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
//*/

void CteWebBrowser::Close()
{
	if (m_pWebBrowser) {
		m_pWebBrowser->Quit();
		IOleObject *pOleObject;
		if SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleObject))) {
			RECT rc;
			SetRectEmpty(&rc);
			pOleObject->DoVerb(OLEIVERB_HIDE, NULL, NULL, 0, g_hwndMain, &rc);
			pOleObject->Release();
		}
		teUnadviseAndRelease(m_pWebBrowser, DIID_DWebBrowserEvents2, &m_dwCookie);
		HWND hwnd = get_HWND();
		m_pWebBrowser = NULL;
		PostMessage(hwnd, WM_CLOSE, 0, 0);
	}
	if (m_pDropTarget) {
		m_pDropTarget->Release();
		m_pDropTarget = NULL;
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
	m_nScrollWidth = 0;
	m_DefTCProc = NULL;
	m_DefBTProc = NULL;
	m_DefSTProc = NULL;
}

VOID CteTabs::GetItem(int i, VARIANT *pVarResult)
{
	CteShellBrowser *pSB = GetShellBrowser(i);
	if (pSB) {
		teSetObject(pVarResult, pSB);
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
			for (int i = TabCtrl_GetItemCount(m_hwnd); i--;) {
				CteShellBrowser *pSB = GetShellBrowser(i);
				if (pSB && pSB->m_bVisible) {
					pSB->Show(FALSE);
				}
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
		SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)m_DefTCProc);
		DestroyWindow(m_hwnd);
		m_hwnd = NULL;

		RevokeDragDrop(m_hwndButton);
		SetWindowLongPtr(m_hwndButton, GWLP_WNDPROC, (LONG_PTR)m_DefBTProc);
		DestroyWindow(m_hwndButton);
		m_hwndButton = NULL;

		SetWindowLongPtr(m_hwndStatic, GWLP_WNDPROC, (LONG_PTR)m_DefSTProc);
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

VOID CteTabs::SetItemSize()
{
	DWORD dwSize = MAKELPARAM(m_param[TC_TabWidth] - m_nScrollWidth, m_param[TC_TabHeight]);
	if (dwSize != m_dwSize) {
		SendMessage(m_hwnd, TCM_SETITEMSIZE, 0, dwSize);
		m_dwSize = dwSize;
	}
}

DWORD CteTabs::GetStyle()
{
	DWORD dwStyle = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | TCS_FOCUSNEVER | m_param[TE_Flags];
	if ((dwStyle & (TCS_SCROLLOPPOSITE | TCS_BUTTONS)) == TCS_SCROLLOPPOSITE) {
		if (m_param[TE_Align] == 4 || m_param[TE_Align] == 5) {
			dwStyle |= TCS_BUTTONS;
		}
	}
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
		WS_EX_CONTROLPARENT | WS_EX_COMPOSITED, //Extended style
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
	SetItemSize();
	SetWindowLongPtr(m_hwnd, GWLP_USERDATA, (LONG_PTR)this);
	m_DefTCProc = (WNDPROC)SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)TETCProc);
	RegisterDragDrop(m_hwnd, static_cast<IDropTarget *>(this));
	BringWindowToTop(m_hwnd);
}

BOOL CteTabs::Create()
{
	m_hwndStatic = CreateWindowEx(
		0, WC_STATIC, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | SS_NOTIFY,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
		g_hwndMain, (HMENU)0, hInst, NULL);
	SetWindowLongPtr(m_hwndStatic, GWLP_USERDATA, (LONG_PTR)this);
	m_DefSTProc = (WNDPROC)SetWindowLongPtr(m_hwndStatic, GWLP_WNDPROC, (LONG_PTR)TESTProc);
	BringWindowToTop(m_hwndStatic);

	m_hwndButton = CreateWindowEx(
		0, WC_BUTTON, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | BS_OWNERDRAW,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
		m_hwndStatic, (HMENU)0, hInst, NULL);

	SetWindowLongPtr(m_hwndButton, GWLP_USERDATA, (LONG_PTR)this);
	m_DefBTProc = (WNDPROC)SetWindowLongPtr(m_hwndButton, GWLP_WNDPROC, (LONG_PTR)TEBTProc);
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
	return teGetDispId(methodTC, _countof(methodTC), g_maps[MAP_TC], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteTabs::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
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
				if (nArg >= 0 && pVarResult) {
					TCHITTESTINFO info;
					GetPointFormVariant(&info.pt, &pDispParams->rgvarg[nArg]);
					UINT flags = nArg >= 1 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : TCHT_ONITEM;
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
					pVarResult->boolVal = m_bVisible ? VARIANT_TRUE : VARIANT_FALSE;
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
				teSetObjectRelease(pVarResult, new CteDispatch(this, 0));
				return S_OK;
			//this
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
		if (dispIdMember >= TE_OFFSET && dispIdMember <= TE_OFFSET + TC_TabHeight) {
			if (nArg >= 0 && dispIdMember != TE_OFFSET) {
				m_param[dispIdMember - TE_OFFSET] = GetIntFromVariantPP(&pDispParams->rgvarg[nArg], dispIdMember - TE_OFFSET);
				if (m_hwnd) {
					if (dispIdMember == TE_OFFSET + TE_Flags) {
						DWORD dwStyle = GetWindowLong(m_hwnd, GWL_STYLE);
						DWORD dwPrev = GetStyle();
						if ((dwStyle ^ dwPrev) & 0x7fff) {
							if ((dwStyle ^ dwPrev) & (TCS_SCROLLOPPOSITE | TCS_BUTTONS | TCS_TOOLTIPS)) {
								//Remake
								int nTabs = TabCtrl_GetItemCount(m_hwnd);
								int nSel = TabCtrl_GetCurFocus(m_hwnd);
								TC_ITEM *tabs = new TC_ITEM[nTabs];
								LPWSTR *ppszText = new LPWSTR[nTabs];
								for (int i = nTabs; i-- > 0;) {
									ppszText[i] = new WCHAR[MAX_PATH];
									tabs[i].cchTextMax = MAX_PATH;
									tabs[i].pszText = ppszText[i];
									tabs[i].mask = TCIF_IMAGE | TCIF_PARAM | TCIF_TEXT;
									TabCtrl_GetItem(m_hwnd, i, &tabs[i]);
								}
								HIMAGELIST hImage = TabCtrl_GetImageList(m_hwnd);
								HFONT hFont = (HFONT)SendMessage(m_hwnd, WM_GETFONT, 0, 0);
								RevokeDragDrop(m_hwnd);
								SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)m_DefTCProc);
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
						SetItemSize();
					}
					SetWindowLong(m_hwnd, GWL_STYLE, GetStyle());
					ArrangeWindow();
				}
			}
			if (pVarResult) {
				int i = m_param[dispIdMember - TE_OFFSET];
				if (dispIdMember >= TE_OFFSET + TE_Left && dispIdMember <= TE_OFFSET + TE_Height) {
					if (i & 0x3fff) {
						wchar_t szSize[20];
						swprintf_s((wchar_t *)&szSize, 20, L"%g%%", ((float)(m_param[dispIdMember - TE_OFFSET])) / 100);
						pVarResult->vt = VT_BSTR;
						pVarResult->bstrVal = SysAllocString(szSize);
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
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"TabControl", methodTC, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteTabs::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	return teGetDispId(methodTC, _countof(methodTC), g_maps[MAP_TC], bstrName, pid, true);
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
		SendMessage(m_hwnd, TCM_DELETEITEM, nSrc, 1/* Not delete SB flag */);
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
		SendMessage(m_hwnd, WM_SETREDRAW, FALSE, 0);
		teSetRedraw(FALSE);
	}
}

VOID CteTabs::UnLockUpdate()
{
	if (::InterlockedDecrement(&m_nLockUpdate) == 0) {
		m_bRedraw = true;
		KillTimer(g_hwndMain, TET_Redraw);
		SetTimer(g_hwndMain, TET_Redraw, 500, teTimerProc);
	}
}

VOID CteTabs::RedrawUpdate()
{
	if (m_bRedraw && g_nLockUpdate == 0) {
		m_bRedraw = false;
		SendMessage(m_hwndStatic, WM_SETREDRAW, TRUE, 0);
		SendMessage(m_hwnd, WM_SETREDRAW, TRUE, 0);
		CteShellBrowser *pSB = GetShellBrowser(m_nIndex);
		if (pSB) {
			pSB->SetRedraw(TRUE);
		}
		RedrawWindow(m_hwndStatic, NULL, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
		teSetRedraw(TRUE);
	}
}

//IDropTarget
STDMETHODIMP CteTabs::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);
	HRESULT hr = DragSub(TE_OnDragEnter, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (hr == S_OK && *pdwEffect) {
		m_pDragItems->Regenerate();
	}
	if (g_pDropTargetHelper) {
		g_pDropTargetHelper->DragEnter(m_hwnd, pDataObj, (LPPOINT)&pt, *pdwEffect);
	}
	return hr;
}

STDMETHODIMP CteTabs::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	HRESULT hr = DragSub(TE_OnDragOver, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (g_pDropTargetHelper) {
		g_pDropTargetHelper->DragOver((LPPOINT)&pt, *pdwEffect);
	}
	m_grfKeyState = grfKeyState;
	return hr;
}

STDMETHODIMP CteTabs::DragLeave()
{
	if (m_DragLeave == E_NOT_SET) {
		if (g_pDropTargetHelper) {
			g_pDropTargetHelper->DragLeave();
		}
		m_DragLeave = DragLeaveSub(this, &m_pDragItems);
	}
	return m_DragLeave;
}

STDMETHODIMP CteTabs::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	GetDragItems(&m_pDragItems, pDataObj);
	if (g_pDropTargetHelper) {
		g_pDropTargetHelper->Drop(pDataObj, (LPPOINT)&pt, *pdwEffect);
	}
	HRESULT hr = DragSub(TE_OnDrop, this, m_pDragItems, &m_grfKeyState, pt, pdwEffect);
	DragLeave();
	return hr;
}

VOID CteTabs::TabChanged(BOOL bSameTC)
{
	if (bSameTC) {
		m_nIndex = TabCtrl_GetCurSel(m_hwnd);
	}
	CteShellBrowser *pSB = GetShellBrowser(m_nIndex);
	if (pSB) {
		pSB->SetStatusTextSB(NULL);
		DoFunc(TE_OnSelectionChanged, this, E_NOTIMPL);
		ArrangeWindow();
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
	m_bsText = NULL;
	if (bEditable) {
		GetNewArray(&m_oFolderItems);
	}
	m_pFolderItems = pFolderItems;
	m_nIndex = 0;
	m_dwEffect = (DWORD)-1;
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
			teCoTaskMemFree(m_pidllist[i]);
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
	teSysFreeString(&m_bsText);
}

VOID CteFolderItems::Regenerate()
{
	if (!m_oFolderItems) {
		long nCount;
		if SUCCEEDED(get_Count(&nCount)) {
			IDispatch *oFolderItems;
			GetNewArray(&oFolderItems);
			VARIANT v;
			v.vt = VT_I4;
			FolderItem *pid;
			for (v.lVal = 0; v.lVal < nCount; v.lVal++) {
				if (Item(v, &pid) == S_OK) {
					teArrayPush(oFolderItems, pid);
				}
			}
			Clear();
			m_oFolderItems = oFolderItems;
		}
	}
}

VOID CteFolderItems::ItemEx(int nIndex, VARIANT *pVarResult, VARIANT *pVarNew)
{
	VARIANT v;
	v.lVal = nIndex;
	v.vt = VT_I4;
	if (pVarNew) {
		teSysFreeString(&m_bsText);
		DISPID dispid;
		Regenerate();
		if (m_oFolderItems) {
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
			teSetObjectRelease(&pv[0], pid);
			VARIANT v1;
			teVariantChangeType(&v1, &v, VT_BSTR);
			WORD wFlags = DISPATCH_PROPERTYPUT;
			m_oFolderItems->GetIDsOfNames(IID_NULL, &v1.bstrVal, 1, LOCALE_USER_DEFAULT, &dispid);
			if (dispid < 0) {
				LPOLESTR lp = L"push";
				m_oFolderItems->GetIDsOfNames(IID_NULL, &lp, 1, LOCALE_USER_DEFAULT, &dispid);
				wFlags = DISPATCH_METHOD;
			}
			VariantClear(&v1);
			Invoke5(m_oFolderItems, dispid, wFlags, NULL, 1, pv);
		}
	}
	if (pVarResult) {
		FolderItem *pid = NULL;
		if SUCCEEDED(Item(v, &pid)) {
			teSetObjectRelease(pVarResult, pid);
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

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch) || IsEqualIID(riid, IID_FolderItems)) {
		*ppvObject = static_cast<FolderItems *>(this);
	}
	else if (IsEqualIID(riid, IID_IDispatchEx)) {
		*ppvObject = static_cast<IDispatchEx *>(this);
	}
	else if (IsEqualIID(riid, IID_IDataObject)) {
		if (!m_pDataObj && m_pidllist && m_nCount) {
			m_bUseILF = AdjustIDList(m_pidllist, m_nCount);
			IShellFolder *pSF;
			if (GetShellFolder(&pSF, m_pidllist[0])) {
				pSF->GetUIObjectOf(g_hwndMain, m_nCount, const_cast<LPCITEMIDLIST *>(&m_pidllist[1]), IID_IDataObject, NULL, (LPVOID *)&m_pDataObj);
				pSF->Release();
			}
		}
		if (m_pDataObj) {
			if (m_pDataObj->QueryGetData(&UNICODEFormat) == S_OK || m_pDataObj->QueryGetData(&TEXTFormat) == S_OK) {
				return m_pDataObj->QueryInterface(riid, ppvObject);
			}
		}
		*ppvObject = static_cast<IDataObject *>(this);
	}
	else if (IsEqualIID(riid, g_ClsIdFIs)) {
		*ppvObject = this;
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
	return teGetDispId(methodFIs, _countof(methodFIs), g_maps[MAP_FIs], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteFolderItems::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
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
				IDispatch *pdisp;
				if SUCCEEDED(get_Application(&pdisp)) {
					teSetObject(pVarResult, pdisp);
					pdisp->Release();
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
				IUnknown *punk;
				if SUCCEEDED(_NewEnum(&punk)) {
					teSetObject(pVarResult, punk);
					punk->Release();
				}
				return S_OK;
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
			//GetData
			case 10:
				if (pVarResult) {
					IDataObject *pDataObj;
					if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pDataObj))) {
						STGMEDIUM Medium;
						if (pDataObj->GetData(&UNICODEFormat, &Medium) == S_OK) {
							pVarResult->bstrVal = ::SysAllocString(static_cast<LPWSTR>(GlobalLock(Medium.hGlobal)));
							pVarResult->vt = VT_BSTR;
							GlobalUnlock(Medium.hGlobal);
							ReleaseStgMedium(&Medium);
						}
						else if (pDataObj->GetData(&TEXTFormat, &Medium) == S_OK) {
							LPCSTR lp = static_cast<LPCSTR>(GlobalLock(Medium.hGlobal));
							int nLenW = MultiByteToWideChar(CP_ACP, 0, lp, -1, NULL, NULL);
							pVarResult->bstrVal = SysAllocStringLen(NULL, nLenW - 1);
							MultiByteToWideChar(CP_ACP, 0, lp, -1, pVarResult->bstrVal, nLenW);
							pVarResult->vt = VT_BSTR;
							GlobalUnlock(Medium.hGlobal);
							ReleaseStgMedium(&Medium);
						}
						pDataObj->Release();
					}
				}
				return S_OK;
			//SetData
			case 11:
				if (nArg >= 0) {
					teSysFreeString(&m_bsText);
					VARIANT v;
					teVariantChangeType(&v, &pDispParams->rgvarg[0], VT_BSTR);
					m_bsText = v.bstrVal;
					v.bstrVal = NULL;
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
				punk = new CteMemory(sizeof(int), (char *)&m_dwEffect, VT_NULL, 1, L"DWORD");
				teSetObject(pVarResult, punk);
				punk->Release();
				return S_OK;
			//this
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"FolderItems", methodFIs, dispIdMember);
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
		HRESULT hr = teGetProperty(m_oFolderItems, L"length", &v);
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
			VARIANT v;
			VariantInit(&v);
			teGetPropertyAt(m_oFolderItems, i, &v);
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
			teCoTaskMemFree(pidl);
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
	return teGetDispId(methodFIs, _countof(methodFIs), g_maps[MAP_FIs], bstrName, pid, true);
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
	for (int i = m_nCount; i-- > 0;) {
		LPITEMIDLIST pidl = ILCombine(m_pidllist[0], m_pidllist[i + 1]);
		if (bSpecial) {
			GetDisplayNameFromPidl(&pbslist[i], pidl, SHGDN_FORPARSING);
			int j = SysStringByteLen(pbslist[i]);
			if (j) {
				uSize += j + sizeof(WCHAR);
			}
			else {
				teSysFreeString(&pbslist[i]);
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
		teCoTaskMemFree(pidl);
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
		}
		catch (...) {
		}
		GlobalUnlock(hGlobal);

		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = hGlobal;
		pmedium->pUnkForRelease = NULL;
		return S_OK;
	}
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
	if (pformatetcIn->cfFormat == CF_HDROP) {
		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = (HGLOBAL)GethDrop(0, 0, FALSE, FALSE);
		pmedium->pUnkForRelease = NULL;
		return S_OK;
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
			}
			catch (...) {
			}
		}
		HGLOBAL hMem;
		int nLen = ::SysStringLen(m_bsText) + 1;
		int nLenA = WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)m_bsText, nLen, NULL, 0, NULL, NULL);
		if (pformatetcIn->cfFormat == CF_UNICODETEXT) {
/*/// For paste problem at file open dialog.
			if (m_nCount) {
				LPSTR lpA = new char[nLenA + 1];
				WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)m_bsText, nLen, lpA, nLenA, NULL, NULL);
				int nLenW = MultiByteToWideChar(CP_ACP, 0, lpA, nLenA, NULL, NULL);
				LPWSTR lpW = new WCHAR[nLenW];
				MultiByteToWideChar(CP_ACP, 0, lpA, nLenA, lpW, nLenW);
				if (lstrcmpi(m_bsText, lpW) == 0) {
					delete [] lpA;
					delete [] lpW;
					return DATA_E_FORMATETC;
				}
				delete [] lpA;
				delete [] lpW;
			}
//*/
			nLen = ::SysStringByteLen(m_bsText) + sizeof(WCHAR);
			hMem = GlobalAlloc(GHND, nLen);
			try {
				LPWSTR lp = static_cast<LPWSTR>(GlobalLock(hMem));
				if (lp) {
					::CopyMemory(lp, m_bsText, nLen);
				}
			}
			catch (...) {
			}
		}
		else {
			hMem = GlobalAlloc(GHND, nLenA);
			try {
				LPSTR lpA = static_cast<LPSTR>(GlobalLock(hMem));
				if (lpA) {
					WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)m_bsText, nLen, lpA, nLenA, NULL, NULL);
				}
			}
			catch (...) {
			}
		}

		GlobalUnlock(hMem);
		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = hMem;
		pmedium->pUnkForRelease = NULL;
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
	if (pformatetc->cfFormat == CF_TEXT) {
		return S_OK;
	}
	if (m_nCount) {
		if (pformatetc->cfFormat == CF_HDROP) {
			return S_OK;
		}
		if (pformatetc->cfFormat == IDLISTFormat.cfFormat && m_bUseILF) {
			return S_OK;
		}
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
		FORMATETC teformats[] = {  IDLISTFormat, HDROPFormat, UNICODEFormat, TEXTFormat, DROPEFFECTFormat };
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
		for (int i = m_nCount ? (m_bUseILF ? 0 : 1) : 2 ; i < nMax; i++) {
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

//CteServiceProvider
STDMETHODIMP CteServiceProvider::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IServiceProvider)) {
		*ppvObject = static_cast<IServiceProvider *>(this);
		AddRef();
		return S_OK;
	}
	return m_pUnk->QueryInterface(riid, ppvObject);
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
	if (IsEqualIID(riid, IID_IShellBrowser) || IsEqualIID(riid, IID_ICommDlgBrowser) || IsEqualIID(riid, IID_ICommDlgBrowser2)) {//|| IsEqualIID(riid, IID_ICommDlgBrowser3)) {
		return m_pUnk->QueryInterface(riid, ppv);
	}
	if (IsEqualIID(riid, IID_IServiceProvider)) {
		return QueryInterface(riid, ppv);
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
	m_bAlloc = FALSE;
	m_nSize = 0;
	if (nSize > 0) {
		m_nSize = nSize;
		if (pc == NULL) {
			m_pc = reinterpret_cast<char *>(::SysAllocStringByteLen(NULL, nSize));
			if (m_pc) {
				::ZeroMemory(m_pc, nSize);
				m_bAlloc = TRUE;
			}
		}
	}
	m_nbs = 0;
	m_ppbs = NULL;
}


CteMemory::~CteMemory()
{
	Free(TRUE);
}

void CteMemory::Free(BOOL bpbs)
{
	switch (m_nMode) {
		case VT_BSTR:
			if (m_nCount > 0) {
				BSTR *lpBSTRData = (BSTR *)m_pc;
				for (int i = m_nCount; i--;) {
					teSysFreeString(&lpBSTRData[i]);
				}
			}
			break;
		case VT_DISPATCH:
		case VT_UNKNOWN:
			if (m_nCount > 0) {
				IUnknown **ppunk = (IUnknown **)m_pc;
				for (int i = m_nCount; i--;) {
					if (ppunk[i]) {
						ppunk[i]->Release();
					}
				}
			}
			break;
	}//end_switch

	if (m_bAlloc) {
		::SysFreeString(reinterpret_cast<BSTR>(m_pc));
		m_bAlloc = FALSE;
	}
	teSysFreeString(&m_bsStruct);
	if (bpbs && m_ppbs) {
		while (--m_nbs >= 0) {
			teSysFreeString(&m_ppbs[m_nbs]);
		}
		delete [] m_ppbs;
		m_ppbs = NULL;
	}
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
	try {
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
				teSetObjectRelease(pVarResult, new CteDispatch(this, 0));
				return S_OK;
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
				Free(TRUE);
				return S_OK;
			//Clone
			case 8:
				if (m_nSize) {
					CteMemory *pMem = new CteMemory(m_nSize, NULL, m_nMode, m_nCount, m_bsStruct);
					::CopyMemory(pMem->m_pc, m_pc, m_nSize);
					teSetObjectRelease(pVarResult, pMem);
				}
				return S_OK;
			//_BLOB
			case 9:
				if (nArg >= 0 && pDispParams->rgvarg[nArg].vt == VT_BSTR) {
					Free(FALSE);
					BSTR bs = ::SysAllocString(pDispParams->rgvarg[nArg].bstrVal);
					m_pc = reinterpret_cast<char *>(bs);
					if (m_pc) {
						m_nSize = ::SysStringByteLen(bs);
						m_bAlloc = TRUE;
					}
				}
				if (pVarResult && m_bAlloc) {
					pVarResult->vt = VT_BSTR;
					pVarResult->bstrVal = (BSTR)m_pc;
				}
				return S_OK;
			//Value
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
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
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"Memory", methodMem, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteMemory::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	int nIndex = teBSearch(structSize, _countof(structSize), g_maps[MAP_SS], m_bsStruct);
	if (nIndex >= 0) {
		if (teGetDispId(pteStruct[g_maps[MAP_SS][nIndex]], 0, NULL, bstrName, pid, false) == S_OK) {
			return S_OK;
		}
	}
	if (teGetDispId(methodMem, _countof(methodMem), g_maps[MAP_Mem], bstrName, pid, true) == S_OK) {
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
			if (m_nCount) {
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
								teSetObjectRelease(pVarResult, new CteMemory(nSize, m_pc + nOffset + (nSize * i), m_nMode, 1, GetSubStruct(m_bsStruct)));
							}
						}
					}
					break;
			}//end_switch
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
	CteMemory *pMem;
	try {
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
				pVarResult->bstrVal = SysAllocStringLen(NULL, nLenW);
				MultiByteToWideChar(CP_ACP, 0, (LPCSTR)&m_pc[nIndex], nLen, pVarResult->bstrVal, nLenW);
				pVarResult->vt = VT_BSTR;
				break;
			case VT_FILETIME:
				teFileTimeToVariantTime((LPFILETIME)&m_pc[nIndex], &pVarResult->date);
				pVarResult->vt = VT_DATE;
				break;
			case VT_CY:
				POINT *ppt;
				ppt = reinterpret_cast<POINT *>(&m_pc[nIndex]);
				pMem = new CteMemory(sizeof(POINT), NULL, VT_NULL, 1, L"POINT");
				pMem->SetPoint(ppt->x, ppt->y);
				teSetObjectRelease(pVarResult, pMem);
				break;
			case VT_CARRAY:
				pMem = new CteMemory(sizeof(RECT), NULL, VT_NULL, 1, L"RECT");
				::CopyMemory(pMem->m_pc, &m_pc[nIndex], 4 * sizeof(int));
				teSetObjectRelease(pVarResult, pMem);
				break;
			case VT_CLSID:
				LPOLESTR lpsz;
				StringFromCLSID(*(IID *)&m_pc[nIndex], &lpsz);
				pVarResult->bstrVal = SysAllocString(lpsz);
				teCoTaskMemFree(lpsz);
				pVarResult->vt = VT_BSTR;
				break;
			case VT_VARIANT:
				VariantCopy(pVarResult, (VARIANT *)&m_pc[nIndex]);
				break;
/*			case VT_RECORD:
				teSetIDList(pVarResult, *(LPITEMIDLIST *)&m_pc[nIndex]);
				break;*/
			default:
				pVarResult->vt = VT_EMPTY;
				break;
		}//end_switch
	}
	catch (...) {
		g_nException = 0;
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
					nLenA = WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)v.bstrVal, nLen, NULL, 0, NULL, NULL) + 1;
					if (nLenA > m_nSize) {
						nLenA = m_nSize;
					}
					WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)v.bstrVal, nLen, &m_pc[nIndex], nLenA - 1, NULL, NULL);
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
				char *pc;
				pc = GetpcFromVariant(pv);
				if (pc) {
					::CopyMemory(&m_pc[nIndex], pc, 16);
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
/*			case VT_RECORD:
				LPITEMIDLIST *ppidl;
				ppidl = (LPITEMIDLIST *)&m_pc[nIndex];
				teCoTaskMemFree(*ppidl);
				GetPidlFromVariant(ppidl, pv);
				break;*/
		}//end_switch
	}
	catch (...) {
		g_nException = 0;
	}
}

void CteMemory::SetPoint(int x, int y)
{
	POINT *ppt = reinterpret_cast<POINT *>(m_pc);
	ppt->x = x;
	ppt->y = y;
}

//CteContextMenu

CteContextMenu::CteContextMenu(IUnknown *punk, IDataObject *pDataObj)
{
	m_cRef = 1;
	m_pContextMenu = NULL;
	m_pShellBrowser = NULL;
	m_pDataObj = pDataObj;
	m_pDll = NULL;

	if (punk) {
		punk->QueryInterface(IID_PPV_ARGS(&m_pContextMenu));
	}
}

CteContextMenu::~CteContextMenu()
{
	if (m_pContextMenu) {
		m_pContextMenu->Release();
	}
	if (m_pDataObj) {
		m_pDataObj->Release();
	}
	if (m_pDll) {
		m_pDll->Release();
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
	return teGetDispId(methodCM, 0, NULL, *rgszNames, rgDispId, true);
}

STDMETHODIMP CteContextMenu::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		HRESULT hr;

		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		switch(dispIdMember) {
			//QueryContextMenu
			case 1:
				if (nArg >= 4) {
					for (int i = 5; i--;) {
						m_param[i] = (UINT_PTR)GetLLFromVariant(&pDispParams->rgvarg[nArg - i]);
					}
					hr = m_pContextMenu->QueryContextMenu((HMENU)m_param[0], (UINT)m_param[1], (UINT)m_param[2], (UINT)m_param[3], (UINT)m_param[4]);
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
					BOOL bExec = TRUE;
					::ZeroMemory(&cmi, sizeof(cmi));
					for (int i = 0; i <= 2; i++) {
						if (pDispParams->rgvarg[nArg - i - 2].vt == VT_I4) {
							UINT_PTR nVerb = (UINT_PTR)pDispParams->rgvarg[nArg - i - 2].lVal;
							ppwc[i] = (WCHAR *)nVerb;
							ppc[i] = (char *)nVerb;
							if (nVerb > MAXWORD) {
								bExec = FALSE;
							}
						}
						else {
							teVariantChangeType(&pv[i], &pDispParams->rgvarg[nArg - i - 2], VT_BSTR);
							ppwc[i] = pv[i].bstrVal;
							if ((ULONG_PTR)ppwc[i] > MAXWORD) {
								int nLenW = ::SysStringLen(ppwc[i]) + 1;
								int nLenA = WideCharToMultiByte(CP_ACP, 0, ppwc[i], nLenW, NULL, 0, NULL, NULL);
								ppc[i] = new char[nLenA];
								WideCharToMultiByte(CP_ACP, 0, ppwc[i], nLenW, ppc[i], nLenA, NULL, NULL);
								BSTR bs = SysAllocStringLen(NULL, nLenW);
								MultiByteToWideChar(CP_ACP, 0, ppc[i], nLenA, bs, nLenW);
								if (lstrcmp(bs, ppwc[i])) {
									cmi.fMask = CMIC_MASK_UNICODE;
								}
								SysFreeString(bs);
							}
							else {
								ppc[i] = (char *)ppwc[i];
							}
						}
					}
					if (bExec) {
						cmi.cbSize = sizeof(cmi);
						cmi.fMask |= (DWORD)GetIntFromVariant(&pDispParams->rgvarg[nArg]);
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
						}
						catch (...) {
							hr = E_UNEXPECTED;
						}
					}
					if (pVarResult) {
						pVarResult->lVal = hr;
						pVarResult->vt = VT_I4;
					}
					for (int i = 3; i--;) {
						VariantClear(&pv[i]);
						if ((ULONG_PTR)ppc[i] >= MAXWORD) {
							delete [] ppc[i];
						}
					}
					delete [] pv;
				}
				return S_OK;
			//Items
			case 3:
				teSetObjectRelease(pVarResult, new CteFolderItems(m_pDataObj, NULL, false));
				return S_OK;
			//GetCommandString
			case 4:
				if (nArg >= 1) {
					if (pVarResult) {
						WCHAR szName[MAX_PATH];
						szName[0] = NULL;
						UINT idCmd = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
						if (idCmd <= MAXWORD) {
							m_pContextMenu->GetCommandString(idCmd,
								GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]),
								NULL, (LPSTR)&szName, MAX_PATH);
						}
						pVarResult->bstrVal = SysAllocString(szName);
						pVarResult->vt = VT_BSTR;
					}
				}
				return S_OK;
			//FolderView
			case 5:
				IServiceProvider *pSP;
				if SUCCEEDED(IUnknown_GetSite(m_pContextMenu, IID_PPV_ARGS(&pSP))) {
					IShellBrowser *pSB;
					if SUCCEEDED(pSP->QueryService(SID_SShellBrowser, IID_PPV_ARGS(&pSB))) {
						IDispatch *pdisp;
						if SUCCEEDED(pSB->QueryInterface(IID_PPV_ARGS(&pdisp))) {
							teSetObjectRelease(pVarResult, pdisp);
						}
						pSB->Release();
					}
					pSP->Release();
				}
				return S_OK;
			//HandleMenuMsg
			case 6:
				LRESULT lResult;
				IContextMenu3 *pCM3;
				IContextMenu2 *pCM2;
				lResult = 0;
				if (nArg >= 2) {
					if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pCM3))) {
						pCM3->HandleMenuMsg2(
							(UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg]),
							(WPARAM)GetLLFromVariant(&pDispParams->rgvarg[nArg - 1]),
							(LPARAM)GetLLFromVariant(&pDispParams->rgvarg[nArg - 2]), &lResult
						);
						pCM3->Release();
					}
					else if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pCM2))) {
						pCM2->HandleMenuMsg(
							(UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg]),
							(WPARAM)GetLLFromVariant(&pDispParams->rgvarg[nArg - 1]),
							(LPARAM)GetLLFromVariant(&pDispParams->rgvarg[nArg - 2])
						);
						pCM2->Release();
					}
				}
				SetVariantLL(pVarResult, lResult);
				return S_OK;

			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
			default:
				if (dispIdMember >= 10 && dispIdMember <= 14) {
					SetVariantLL(pVarResult, m_param[dispIdMember - 10]);
					return S_OK;
				}
				break;
		}
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"ContextMenu", methodCM, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
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
	if (m_pFolderItem) {
		m_pFolderItem->Release();
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
	else if (m_pFolderItem && (IsEqualIID(riid, IID_FolderItem) || IsEqualIID(riid, g_ClsIdFI))) {
		return m_pFolderItem->QueryInterface(riid, ppvObject);
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
	return teGetDispId(methodDT, 0, NULL, *rgszNames, rgDispId, true);
}

STDMETHODIMP CteDropTarget::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
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
				if (nArg >= 2) {
					BOOL bSingle = (nArg >= 4) && GetIntFromVariant(&pDispParams->rgvarg[nArg - 4]);
					IDataObject *pDataObj;
					if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
						try {
							DWORD grfKeyState = (DWORD)GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
							POINTL pt;
							GetPointFormVariant(reinterpret_cast<POINT *>(&pt), &pDispParams->rgvarg[nArg - 2]);
							DWORD dwEffect1 = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
							VARIANT v;
							VariantInit(&v);
							IDispatch *pEffect = NULL;
							if (nArg >= 3 && GetDispatch(&pDispParams->rgvarg[nArg - 3], &pEffect)) {
								teGetPropertyAt(pEffect, 0, &v);
								dwEffect1 = GetIntFromVariantClear(&v);
							}
							DWORD dwEffect = dwEffect1;
							CteFolderItems *pDragItems = new CteFolderItems(pDataObj, NULL, false);
							try {
								POINTL pt0 = {0, 0};
								if (!bSingle || dispIdMember == 1) {//DragEnter
									if (m_pDropTarget) {
										try {
											hr = m_pDropTarget->DragEnter(pDataObj, grfKeyState, pt, &dwEffect1);
										} catch (...) {
											hr = E_UNEXPECTED;
										}
									}
									else {
										dwEffect1 = 0;
									}
									if (m_pFolderItem) {
										DWORD dwEffect2 = dwEffect;
										if (DragSub(TE_OnDragEnter, this, pDragItems, &grfKeyState, pt0, &dwEffect2) == S_OK) {
											hr = S_OK;
											dwEffect1 = dwEffect2;
										}
									}
								}
								else {
									hr = S_OK;
								}
								if (hr == S_OK && dispIdMember >= 2) {
									if (!bSingle || dispIdMember == 2) {//DragOver
										hr = S_FALSE;
										if (m_pFolderItem) {
											dwEffect1 = dwEffect;
											hr = DragSub(TE_OnDragOver, this, pDragItems, &grfKeyState, pt0, &dwEffect1);
										}
										if (hr != S_OK && m_pDropTarget) {
											dwEffect1 = dwEffect;
											try {
												hr = m_pDropTarget->DragOver(grfKeyState, pt, &dwEffect1);
											} catch (...) {
												hr = E_UNEXPECTED;
											}
										}
									}
									if (hr == S_OK && dispIdMember >= 3) {
										if (!bSingle || dispIdMember == 3) {//Drop
											hr = S_FALSE;
											if (m_pFolderItem) {
												dwEffect1 = dwEffect;
												hr = DragSub(TE_OnDrop, this, pDragItems, &grfKeyState, pt0, &dwEffect1);
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
							}
							catch (...) {
							}
							pDragItems->Release();
							if (!bSingle) {
								if (m_pDropTarget) {
									m_pDropTarget->DragLeave();
								}
								if (dispIdMember >= 3) {
									DragLeaveSub(this, NULL);
								}
							}
							if (pEffect) {
								v.vt = VT_I4;
								v.lVal = dwEffect1;
								tePutPropertyAt(pEffect, 0, &v);
								pEffect->Release();
							}
						} catch (...) {
						}
						pDataObj->Release();
					}
				}
				if (pVarResult) {
					pVarResult->lVal = hr;
					pVarResult->vt = VT_I4;
				}
				return S_OK;

			//DragLeave
			case 4:
				if (m_pDropTarget) {
					hr = m_pDropTarget->DragLeave();
				}
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
				teSetObject(pVarResult, m_pFolderItem);
				return S_OK;

			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"DropTarget", methodDT, dispIdMember);
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
		m_DefProc = NULL;
	}
	if (m_DefProc2) {
		SetWindowLongPtr(m_hwndTV, GWLP_WNDPROC, (LONG_PTR)m_DefProc2);
		m_DefProc2 = NULL;
	}
	if (m_pDropTarget) {
		RevokeDragDrop(m_hwndTV);
		RegisterDragDrop(m_hwndTV, m_pDropTarget);
		m_pDropTarget->Release();
		m_pDropTarget = NULL;
	}
#ifdef _W2000
	VariantClear(&m_vSelected);
#endif
	VariantClear(&m_Data);
}

void CteTreeView::Init()
{
	m_hwnd = NULL;
	m_hwndTV = NULL;
	m_pNameSpaceTreeControl = NULL;
	m_bSetRoot = true;
#ifdef _2000XP
	m_pShellNameSpace = NULL;
#endif
	m_dwCookie = 0;
}

void CteTreeView::Close()
{
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
			DestroyWindow(m_hwnd);
		}
		m_pShellNameSpace->Release();
		m_pShellNameSpace = NULL;
	}
#endif
}

BOOL CteTreeView::Create()
{
	if (
//		false &&
		SUCCEEDED(CoCreateInstance(CLSID_NamespaceTreeControl, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pNameSpaceTreeControl)))) {
		RECT rc;
		SetRectEmpty(&rc);
		if SUCCEEDED(m_pNameSpaceTreeControl->Initialize(m_pFV->m_pTabs->m_hwndStatic, &rc, m_pFV->m_param[SB_TreeFlags])) {
			m_pNameSpaceTreeControl->TreeAdvise(static_cast<INameSpaceTreeControlEvents *>(this), &m_dwCookie);
			IOleWindow *pOleWindow;
			if SUCCEEDED(m_pNameSpaceTreeControl->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
				if (pOleWindow->GetWindow(&m_hwnd) == S_OK) {
					m_hwndTV = FindTreeWindow(m_hwnd);
					if (m_hwndTV) {
						SetWindowLongPtr(m_hwnd, GWLP_USERDATA, (LONG_PTR)this);
						m_DefProc = (WNDPROC)SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)TETVProc);
						SetWindowLongPtr(m_hwndTV, GWLP_USERDATA, (LONG_PTR)this);
						m_DefProc2 = (WNDPROC)SetWindowLongPtr(m_hwndTV, GWLP_WNDPROC, (LONG_PTR)TETVProc2);
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
#ifdef _2000XP
	m_hwnd = CreateWindowEx(
		WS_EX_CLIENTEDGE, WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | ES_READONLY,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
		m_pFV->m_pTabs->m_hwndStatic, (HMENU)0, hInst, NULL);
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
/*/// Not use in IShellNameSpace
	else if (m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pPersistStorage)) == S_OK) {
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
	}
//*/
	else {
		return FALSE;
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
		return FALSE;
	}
	pOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, static_cast<IOleClientSite *>(this), 0, m_pFV->m_pTabs->m_hwndStatic, &rc);
	pOleObject->Release();

	IOleWindow *pOleWindow;
	if SUCCEEDED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
		HWND hwnd;
		if (pOleWindow->GetWindow(&hwnd) == S_OK) {
			m_pShellNameSpace->put_Flags(m_pFV->m_param[SB_TreeFlags] & NSTCS_FAVORITESMODE);
			m_hwndTV = FindTreeWindow(hwnd);
			if (m_hwndTV) {
				HWND hwndParent = GetParent(m_hwndTV);
				SetWindowLongPtr(hwndParent, GWLP_USERDATA, (LONG_PTR)this);
				m_DefProc = (WNDPROC)SetWindowLongPtr(hwndParent, GWLP_WNDPROC, (LONG_PTR)TETVProc);
				SetWindowLongPtr(m_hwndTV, GWLP_USERDATA, (LONG_PTR)this);
				m_DefProc2 = (WNDPROC)SetWindowLongPtr(m_hwndTV, GWLP_WNDPROC, (LONG_PTR)TETVProc2);
				m_pFV->HookDragDrop(g_dragdrop & 2);

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
				SetWindowLongPtr(m_hwndTV, GWL_STYLE, dwStyle & ~WS_BORDER);
				DWORD dwExStyle = GetWindowLong(m_hwnd, GWL_EXSTYLE);
				if (dwStyle & WS_BORDER) {
					dwExStyle |= WS_EX_CLIENTEDGE;
				}
				else {
					dwExStyle &= ~WS_EX_CLIENTEDGE;
				}
				SetWindowLong(m_hwnd, GWL_EXSTYLE, dwExStyle);
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

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)
#ifdef _2000XP
		|| IsEqualIID(riid, DIID_DShellNameSpaceEvents)
#endif
		) {
		*ppvObject = static_cast<IDispatch *>(this);
	}
/*	else if (IsEqualIID(riid, IID_IContextMenu)) {
		*ppvObject = static_cast<IContextMenu *>(this);
	}*/
	else if (IsEqualIID(riid, IID_IDropTarget)) {
		*ppvObject = static_cast<IDropTarget *>(this);
	}
	else if (IsEqualIID(riid, IID_INameSpaceTreeControlEvents)) {
		*ppvObject = static_cast<INameSpaceTreeControlEvents *>(this);
	}
	else if (m_pNameSpaceTreeControl && IsEqualIID(riid, IID_INameSpaceTreeControl)) {
		return m_pNameSpaceTreeControl->QueryInterface(riid, ppvObject);
	}
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
	return teGetDispId(methodTV, _countof(methodTV), g_maps[MAP_TV], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteTreeView::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}

		if (dispIdMember >= 0x10000101 && dispIdMember < TE_OFFSET) {
			if (m_pFV->m_param[SB_TreeAlign] & 2) {
				if (!m_pNameSpaceTreeControl
#ifdef _2000XP
					&& !m_pShellNameSpace
#endif
				) {
					Create();
				}
				if (m_bSetRoot && (m_pNameSpaceTreeControl
#ifdef _2000XP
					|| m_pShellNameSpace
#endif
					)) {
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
				m_pFV->m_param[SB_TreeAlign] = 0;
				ArrangeWindow();
				return S_OK;

			//hwnd
			case 0x10000005:
				SetVariantLL(pVarResult, (LONGLONG)m_hwndTV);
				return S_OK;
			//FolderView
			case 0x10000007:
				teSetObject(pVarResult, m_pFV);
				return S_OK;
			//Align
			case 0x10000008:
				if (nArg >= 0) {
					m_pFV->m_param[SB_TreeAlign] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					if (!m_pNameSpaceTreeControl
#ifdef _2000XP
						&& !m_pShellNameSpace
#endif
						) {
						if (m_pFV->m_param[SB_TreeAlign] & 2) {
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
			//Visible
			case 0x10000009:
				if (nArg >= 0) {
					m_pFV->m_param[SB_TreeAlign] = GetIntFromVariant(&pDispParams->rgvarg[nArg]) ? 3 : 1;
					if (!m_pNameSpaceTreeControl
#ifdef _2000XP
						&& !m_pShellNameSpace
#endif
						) {
						if (m_pFV->m_param[SB_TreeAlign] & 2) {
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
					pVarResult->boolVal = m_pFV->m_param[SB_TreeAlign] & 2 ? VARIANT_TRUE : VARIANT_FALSE;
					pVarResult->vt = VT_BOOL;
				}
				return S_OK;
			//Focus
			case 0x10000105:
				SetFocus(m_hwndTV);
				return S_OK;
			//HitTest (Screen coordinates)
			case 0x10000106:
				if (nArg >= 0 && pVarResult) {
					TVHITTESTINFO info;
					GetPointFormVariant(&info.pt, &pDispParams->rgvarg[nArg]);
					UINT flags = nArg >= 1 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : TVHT_ONITEM;
					HTREEITEM hItem = (HTREEITEM)DoHitTest(this, info.pt, flags);
					if ((LONGLONG)hItem < 1) {
						ScreenToClient(m_hwndTV, &info.pt);
						info.flags = flags;
						hItem = TreeView_HitTest(m_hwndTV, &info);
						if (!(info.flags & flags)) {
							hItem = 0;
						}
					}
					SetVariantLL(pVarResult, (LONGLONG)hItem);
				}
				return S_OK;

			//SelectedItem
			case 0x20000000:
				IDispatch *pdisp;
				if (getSelected(&pdisp) == S_OK) {
					teSetObjectRelease(pVarResult, pdisp);
					return S_OK;
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
			case 0x20000004:
				if (m_pNameSpaceTreeControl) {
					if (nArg >= 1) {
						LPITEMIDLIST pidl;
						GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg]);
						IShellItem *pShellItem;
						DWORD dwState;
						dwState = NSTCIS_SELECTED;
						if (GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) != 0) {
							dwState |= NSTCIS_EXPANDED;
						}
						if (lpfnSHCreateItemFromIDList) {
							if SUCCEEDED(lpfnSHCreateItemFromIDList(pidl, IID_PPV_ARGS(&pShellItem))) {
								m_pNameSpaceTreeControl->SetItemState(pShellItem, dwState, dwState);
								pShellItem->Release();
							}
						}
						teCoTaskMemFree(pidl);
					}
					return S_OK;
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
				teSetObject(pVarResult, this);
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
#ifdef _2000XP
			if (m_pNameSpaceTreeControl) {
				return S_OK;
			}
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
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"TreeView", methodTV, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

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
				teSetObjectRelease(&pv[2], pdisp);
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
/*/// For check
	if (pcmIn) {
		HMENU hMenu = CreatePopupMenu();
		if SUCCEEDED(pcmIn->QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXPLORE)) {
			int nVerb = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, msg1.pt.x, msg1.pt.y, g_hwndMain, NULL);

		}
		DestroyMenu(hMenu);
		return QueryInterface(IID_IContextMenu, ppv);
	}
*///
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

/*/// IContextMenu
STDMETHODIMP CteTreeView::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
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
//*/
#ifdef _2000XP
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
			teCoTaskMemFree(pidl);
		}
	}
#ifdef _2000XP
	if (hr != S_OK && m_pShellNameSpace) {
		DWORD dwStyle = GetWindowLong(m_hwndTV, GWL_STYLE);
		if (m_pFV->m_param[SB_RootStyle] & NSTCRS_HIDDEN) {
			dwStyle &= ~TVS_LINESATROOT;
		}
		else {
			dwStyle |= TVS_LINESATROOT;
		}
		SetWindowLong(m_hwndTV, GWL_STYLE, dwStyle);

		//running to call DISPID_FAVSELECTIONCHANGE SetRoot(Windows 2000)
		m_pShellNameSpace->put_EnumOptions(m_pFV->m_param[SB_EnumFlags]);

		LPITEMIDLIST pidl;
		if (GetPidlFromVariant(&pidl, &m_pFV->m_vRoot)) {
			BSTR bs;
			if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORPARSING)) {
				if (::SysStringLen(bs) && bs[0] == ':') {
					hr = m_pShellNameSpace->SetRoot(bs);
				}
				::SysFreeString(bs);
			}
			teCoTaskMemFree(pidl);
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
				if SUCCEEDED(GetFolderItemFromShellItem(&pFI, psi)) {
					hr = pFI->QueryInterface(IID_PPV_ARGS(ppid));
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

//IDropTarget
STDMETHODIMP CteTreeView::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	m_DragLeave = E_NOT_SET;
	GetDragItems(&m_pDragItems, pDataObj);
	DWORD dwEffect = *pdwEffect;
	HRESULT hr = DragSub(TE_OnDragEnter, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		*pdwEffect = dwEffect;
		if (m_pDropTarget->DragEnter(pDataObj, grfKeyState, pt, pdwEffect) == S_OK) {
			hr = S_OK;
		}
	}
	if (hr == S_OK && *pdwEffect) {
		m_pDragItems->Regenerate();
	}
	return hr;
}

STDMETHODIMP CteTreeView::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	DWORD dwEffect2 = *pdwEffect;
	HRESULT hr2 = E_NOTIMPL;
	HRESULT hr = DragSub(TE_OnDragOver, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		hr2 = m_pDropTarget->DragOver(grfKeyState, pt, &dwEffect2);
	}
	if (hr != S_OK) {
		*pdwEffect = dwEffect2;
		hr = hr2;
	}
	m_grfKeyState = grfKeyState;
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
	HRESULT hr = DragSub(TE_OnDrop, this, m_pDragItems, &m_grfKeyState, pt, pdwEffect);
	if (m_pDropTarget) {
		if (hr != S_OK) {
			*pdwEffect = dwEffect;
			hr = m_pDropTarget->Drop(pDataObj, m_grfKeyState, pt, pdwEffect);
		}
	}
	DragLeave();
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
		GetIDListFromObject(pdisp, ppidl);
		pdisp->Release();
	}
	return S_OK;
}

// CteFolderItem

CteFolderItem::CteFolderItem(VARIANT *pv)
{
	m_cRef = 1;
	m_pidl = NULL;
	m_pFolderItem = NULL;
	m_bStrict = FALSE;
	m_pidlFocused = NULL;
	VariantInit(&m_v);
	if (pv) {
		VariantCopy(&m_v, pv);
	}
}

CteFolderItem::~CteFolderItem()
{
	teCoTaskMemFree(m_pidl);
	m_pFolderItem && m_pFolderItem->Release();
	::VariantClear(&m_v);
	teCoTaskMemFree(m_pidlFocused);
}

LPITEMIDLIST CteFolderItem::GetPidl()
{
	if (m_v.vt != VT_EMPTY && m_pidl == NULL) {
		VARIANT vPath = m_v;
		VARIANT vResult;
		VariantInit(&vResult);
		if (m_v.vt == VT_BSTR) {
			if (g_pOnFunc[TE_OnTranslatePath]) {
				try {
					if (InterlockedIncrement(&g_nProcFI) < 5) {
						VARIANTARG *pv = GetNewVARIANT(2);
						teSetObject(&pv[1], this);
						VariantCopy(&pv[0], &m_v);
						Invoke4(g_pOnFunc[TE_OnTranslatePath], &vResult, 2, pv);
						if (vResult.vt == VT_BSTR) {
							vPath = vResult;
						}
					}
				}
				catch(...) {}
				::InterlockedDecrement(&g_nProcFI);
			}
		}
		if (GetPidlFromVariant(&m_pidl, &vPath)) {
			if (m_v.vt != VT_BSTR) {
				::VariantClear(&m_v);
			}
			else if (lstrcmp(m_v.bstrVal, vPath.bstrVal) == 0) {
				::VariantClear(&m_v);
			}
		}
		VariantClear(&vResult);
	}
	return m_pidl ? m_pidl : g_pidls[CSIDL_DESKTOP];
}

BOOL CteFolderItem::GetFolderItem()
{
	if (m_pFolderItem) {
		return TRUE;
	}
	return GetFolderItemFromPidl2(&m_pFolderItem, GetPidl());
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
	else if (IsEqualIID(riid, IID_FolderItem2) && GetFolderItem()) {
		return m_pFolderItem->QueryInterface(riid, ppvObject);
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
	teILCloneReplace(&m_pidl, pidl);
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
	HRESULT hr = teGetDispId(methodFI, 0, NULL, *rgszNames, rgDispId, true);
	if SUCCEEDED(hr) {
		return S_OK;
	}
	if (GetFolderItem()) {
		hr = m_pFolderItem->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
		if (SUCCEEDED(hr) && *rgDispId > 0) {
			*rgDispId += 10;
		}
	}
	return hr;
}

STDMETHODIMP CteFolderItem::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		HRESULT hr = E_FAIL;
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		switch(dispIdMember) {
			//Name
			case 1:
				if (nArg >= 0) {
					put_Name(pDispParams->rgvarg[nArg].bstrVal);
				}
				if (pVarResult) {
					if SUCCEEDED(get_Name(&pVarResult->bstrVal)) {
						pVarResult->vt = VT_BSTR;
					}
				}
				return S_OK;
			//Path
			case 2:
				if (pVarResult) {
					if SUCCEEDED(get_Path(&pVarResult->bstrVal)) {
						pVarResult->vt = VT_BSTR;
					}
				}
				return S_OK;
			//Alt
			case 3:
				if (nArg >= 0) {
					teILFreeClear(&m_pidl);
					GetPidlFromVariant(&m_pidl, &pDispParams->rgvarg[nArg]);
				}
				return S_OK;
			//_BLOB
			case 9:
				return DISP_E_MEMBERNOTFOUND;
/*/// Reserved future
			//FocusedItem
			case 0x4001004:
				if (nArg >= 0) {
					teILFreeClear(&m_pidlFocused);
					GetPidlFromVariant(&m_pidlFocused, &pDispParams->rgvarg[nArg]);
				}
				if (pVarResult) {
					teSetIDList(pVarResult, m_pidlFocused);
				}
				return S_OK;
//*/
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
		if (GetFolderItem()) {
			if (dispIdMember > 10) {
				dispIdMember -= 10;
			}
			return m_pFolderItem->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"FolderItem", methodFI, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

STDMETHODIMP CteFolderItem::get_Application(IDispatch **ppid)
{
	if (GetFolderItem()) {
		m_pFolderItem->get_Application(ppid);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::get_Parent(IDispatch **ppid)
{
	if (GetFolderItem()) {
		m_pFolderItem->get_Parent(ppid);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::get_Name(BSTR *pbs)
{
	if (m_v.vt == VT_BSTR) {
		*pbs = ::SysAllocString(PathFindFileName(m_v.bstrVal));
		return S_OK;
	}
	return GetDisplayNameFromPidl(pbs, GetPidl(), SHGDN_INFOLDER);
}

STDMETHODIMP CteFolderItem::put_Name(BSTR bs)
{
	if (GetFolderItem()) {
		m_pFolderItem->put_Name(bs);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::get_Path(BSTR *pbs)
{
	if (m_v.vt == VT_BSTR) {
		*pbs = ::SysAllocString(m_v.bstrVal);
		return S_OK;
	}
	return GetDisplayNameFromPidl(pbs, GetPidl(), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
}

STDMETHODIMP CteFolderItem::get_GetLink(IDispatch **ppid)
{
	if (GetFolderItem()) {
		m_pFolderItem->get_GetLink(ppid);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::get_GetFolder(IDispatch **ppid)
{
	return GetFolderObjFromPidl(GetPidl(), (Folder **)ppid);
}

STDMETHODIMP CteFolderItem::get_IsLink(VARIANT_BOOL *pb)
{
	if (GetFolderItem()) {
		m_pFolderItem->get_IsLink(pb);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::get_IsFolder(VARIANT_BOOL *pb)
{
	if (GetFolderItem()) {
		m_pFolderItem->get_IsFolder(pb);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::get_IsFileSystem(VARIANT_BOOL *pb)
{
	if (GetFolderItem()) {
		m_pFolderItem->get_IsFileSystem(pb);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::get_IsBrowsable(VARIANT_BOOL *pb)
{
	if (GetFolderItem()) {
		m_pFolderItem->get_IsBrowsable(pb);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::get_ModifyDate(DATE *pdt)
{
	if (GetFolderItem()) {
		m_pFolderItem->get_ModifyDate(pdt);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::put_ModifyDate(DATE dt)
{
	if (GetFolderItem()) {
		m_pFolderItem->put_ModifyDate(dt);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::get_Size(LONG *pul)
{
	if (GetFolderItem()) {
		m_pFolderItem->get_Size(pul);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::get_Type(BSTR *pbs)
{
	if (GetFolderItem()) {
		m_pFolderItem->get_Type(pbs);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::Verbs(FolderItemVerbs **ppfic)
{
	if (GetFolderItem()) {
		m_pFolderItem->Verbs(ppfic);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteFolderItem::InvokeVerb(VARIANT vVerb)
{
	if (GetFolderItem()) {
		m_pFolderItem->InvokeVerb(vVerb);
	}
	return E_NOTIMPL;
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
	return teGetDispId(methodCD, _countof(methodCD), g_maps[MAP_CD], *rgszNames, rgDispId, true);
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
		switch(dispIdMember) {
			//FileName
			case 10:
				if (nArg >= 0) {
					VARIANT vFile;
					teVariantChangeType(&vFile, &pDispParams->rgvarg[nArg], VT_BSTR);
					if (!m_ofn.lpstrFile) {
						m_ofn.lpstrFile = SysAllocStringLen(vFile.bstrVal, m_ofn.nMaxFile);
					}
					else {
						lstrcpyn(m_ofn.lpstrFile, vFile.bstrVal, m_ofn.nMaxFile);
					}
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
						teSysFreeString(const_cast<BSTR *>(&m_ofn.lpstrFilter));
						int i = ::SysStringLen(pDispParams->rgvarg[nArg].bstrVal);
						BSTR bs = SysAllocStringLen(pDispParams->rgvarg[nArg].bstrVal, i + 1);
						bs[i + 1] = NULL;
						while (i >= 0) {
							if (bs[i] == '|') {
								bs[i] = NULL;
							}
							i--;
						}
						m_ofn.lpstrFilter = bs;
					}
				}
				if (pVarResult) {
					pVarResult->bstrVal = SysAllocString(m_ofn.lpstrInitialDir);
					pVarResult->vt = VT_BSTR;
				}
				return S_OK;
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
		if (dispIdMember >= 40) {
			if (!m_ofn.lpstrFile) {
				m_ofn.lpstrFile = SysAllocStringLen(NULL, m_ofn.nMaxFile);
				m_ofn.lpstrFile[0] = NULL;
			}
			BOOL bResult = FALSE;
			switch (dispIdMember) {
				case 40:
					if (m_ofn.Flags & OFN_ENABLEHOOK) {
						m_ofn.lpfnHook = OFNHookProc;
					}
					g_bDialogOk = FALSE;
					bResult = GetOpenFileName(&m_ofn) || g_bDialogOk;
					break;
				case 41:
					bResult = GetSaveFileName(&m_ofn);
					break;
/*/// IFileOpenDialog
				case 42:
					{
						IFileOpenDialog *pFileOpenDialog;
						if (lpfnSHCreateItemFromIDList && SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pFileOpenDialog)))) {
							IShellItem *psi;
							LPITEMIDLIST pidl = teILCreateFromPath(const_cast<LPWSTR>(m_ofn.lpstrInitialDir));
							if (pidl) {
								if SUCCEEDED(lpfnSHCreateItemFromIDList(pidl, IID_PPV_ARGS(&psi))) {
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
										GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORPARSING);
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
							GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORPARSING);
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
			}//end_switch
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
					pbs = &m_ofn.lpstrInitialDir;
					break;
				case 21:
					pbs = &m_ofn.lpstrDefExt;
					break;
				case 22:
					pbs = &m_ofn.lpstrTitle;
					break;
			}//end_switch
			if (nArg >= 0) {
				if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
					SysReAllocString(const_cast<BSTR *>(pbs), pDispParams->rgvarg[nArg].bstrVal);
				}
			}
			if (pVarResult) {
				pVarResult->bstrVal = SysAllocString(*pbs);
				pVarResult->vt = VT_BSTR;
			}
			return S_OK;
		}
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"CommonDialog", methodCD, dispIdMember);
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
	return teGetDispId(methodGB, _countof(methodGB), g_maps[MAP_GB], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteGdiplusBitmap::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
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
				}
				return S_OK;
			//Free(99)
				return S_OK;
			//Save
			case 100:
				if (nArg >= 0) {
					VARIANT vText;
					teVariantChangeType(&vText, &pDispParams->rgvarg[nArg], VT_BSTR);
					CLSID encoderClsid;
					if (GetEncoderClsid(vText.bstrVal, &encoderClsid, NULL)) {
						EncoderParameters *pEncoderParameters = NULL;
						if (nArg) {
							char *pc = GetpcFromVariant(&pDispParams->rgvarg[nArg - 1]);
							if (pc) {
								pEncoderParameters = (EncoderParameters *)pc;
							}
						}
						Result = m_pImage->Save(vText.bstrVal, &encoderClsid, pEncoderParameters);
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
					if (GetEncoderClsid(vText.bstrVal, &encoderClsid, szMime)) {
						IStream* pStream = NULL;
						if SUCCEEDED(CreateStreamOnHGlobal(NULL, TRUE, &pStream)) {
							if (m_pImage->Save(pStream, &encoderClsid) == 0) {
								ULARGE_INTEGER ulnSize;
								LARGE_INTEGER lnOffset;
								lnOffset.QuadPart = 0;
								pStream->Seek(lnOffset, STREAM_SEEK_END, &ulnSize);
								pStream->Seek(lnOffset, STREAM_SEEK_SET, NULL);
								UCHAR *pBuff = new UCHAR[ulnSize.LowPart];
								ULONG ulBytesRead;
								if SUCCEEDED(pStream->Read(pBuff, (ULONG)ulnSize.QuadPart, &ulBytesRead)) {
									wchar_t szHead[32];
									szHead[0] = NULL;
									if (dispIdMember == 102) {
										swprintf_s(szHead, 32, L"data:%s;base64,", szMime);
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
							pStream->Release();
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
					CLSID encoderClsid;
					if (GetEncoderClsid(L"image/png", &encoderClsid, NULL)) {
						IStream* pStream = NULL;
						if SUCCEEDED(CreateStreamOnHGlobal(NULL, TRUE, &pStream)) {
							int x = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
							int y = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
							if (x > 96 || y > 96) {
								GUID guid;
								m_pImage->GetRawFormat(&guid);
								if (guid == ImageFormatJPEG || guid == ImageFormatTIFF || guid == ImageFormatUndefined) {
									HICON hIcon;
									m_pImage->GetHICON(&hIcon);
									delete m_pImage;
									m_pImage = Gdiplus::Bitmap::FromHICON(hIcon);
									DeleteObject(hIcon);
								}
							}
							Gdiplus::Image *pImage = m_pImage->GetThumbnailImage(x, y);
							if (pImage->Save(pStream, &encoderClsid) == 0) {
								CteGdiplusBitmap *pGB = new CteGdiplusBitmap();
								pGB->m_pImage = Gdiplus::Bitmap::FromStream(pStream, FALSE);
								teSetObjectRelease(pVarResult, pGB);
							}
							delete pImage;
							pStream->Release();
						}
					}
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
#ifdef _2000XP
					HICON hIcon;
					if (cl && !g_bUpperVista && m_pImage->GetHICON(&hIcon) == 0) {
						HDC hdc = GetDC(g_hwndMain);
						ICONINFO iconinfo;
						GetIconInfo(hIcon, &iconinfo);
						DIBSECTION dib;
						GetObject(iconinfo.hbmColor, sizeof(DIBSECTION), &dib);

						hBM = CreateCompatibleBitmap(hdc , dib.dsBm.bmWidth, dib.dsBm.bmHeight);
						HDC hCompatDC = CreateCompatibleDC(hdc);
						SelectObject(hCompatDC, hBM);
						RECT rc = {0, 0, dib.dsBm.bmWidth, dib.dsBm.bmHeight};
						SetBkColor(hCompatDC, cl);
						FillRect(hCompatDC, &rc, NULL);
						DrawIconEx(hCompatDC, 0, 0, hIcon, dib.dsBm.bmWidth, dib.dsBm.bmHeight, 0, 0, DI_NORMAL);
						DeleteDC(hCompatDC);
						ReleaseDC(g_hwndMain, hdc);
						DeleteObject(iconinfo.hbmColor);
						DeleteObject(iconinfo.hbmMask);
						DestroyIcon(hIcon);
						SetVariantLL(pVarResult, (LONGLONG)hBM);
						return S_OK;
					}
#endif
					if (m_pImage->GetHBITMAP(((cl & 0xff) << 16) + (cl & 0xff00) + ((cl & 0xff0000) >> 16), &hBM) == 0) {
						SetVariantLL(pVarResult, (LONGLONG)hBM);
					}
				}
				return S_OK;
			//GetHICON
			case 211:
				HICON hIcon;
				if (m_pImage->GetHICON(&hIcon) == 0) {
					SetVariantLL(pVarResult, (LONGLONG)hIcon);
				}
				return S_OK;
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"GdiplusBitmap", methodGB, dispIdMember);
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
	return teGetDispId(methodAPI, _countof(methodAPI), g_maps[MAP_API], *rgszNames, rgDispId, false);
}

STDMETHODIMP CteWindowsAPI::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		LONGLONG Result = 0;
		HANDLE *phResult = (HANDLE *)&Result;
		LONG *plResult = (LONG *)&Result;
		BOOL *pbResult = (BOOL *)&Result;
		BSTR bsResult = NULL;
		int nLen = 0;
		IUnknown *punk = NULL;

		if (wFlags == DISPATCH_PROPERTYGET) {
			if (dispIdMember != DISPID_VALUE) {
				CteDispatch *pClone = new CteDispatch(this, 0);
				pClone->m_dispIdMember = dispIdMember;
				teSetObjectRelease(pVarResult, pClone);
				return S_OK;
			}
		}
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (dispIdMember >= 10000) {
			INT_PTR param[16];
			for (int i = nArg < 15 ? nArg : 15; i >= 0; i--) {
				param[i] = (INT_PTR)GetLLFromVariant(&pDispParams->rgvarg[nArg - i]);
			}
			if (dispIdMember >= 50000) {//string
				if (pVarResult) {
					switch(dispIdMember) {
						case 50005:
							if (nArg >= 0) {
								nLen = GetWindowTextLength((HWND)param[0]);
								if (nLen) {
									bsResult = SysAllocStringLen(NULL, nLen);
									nLen = GetWindowText((HWND)param[0], bsResult, nLen + 1);
								}
							}
							break;
						case 50015:
							if (nArg >= 0) {
								bsResult = SysAllocStringLen(NULL, MAX_CLASS_NAME);
								nLen = GetClassName((HWND)param[0], bsResult, MAX_CLASS_NAME);
							}
							break;
						case 50145:
							if (nArg >= 1) {
								bsResult = SysAllocStringLen(NULL, SIZE_BUFF);
								nLen = LoadString((HINSTANCE)param[0], (UINT)param[1], bsResult, SIZE_BUFF);
							}
							break;
						case 50025:
							if (nArg >= 0) {
								nLen = teGetModuleFileName((HMODULE)param[0], &bsResult);
							}
							break;
						case 50035:
							bsResult = teGetCommandLine();
							nLen = -2;
							break;
						case 50045:
							//use wsh.CurrentDirectory
							bsResult = SysAllocStringLen(NULL, MAX_PATH);
							nLen = GetCurrentDirectory(MAX_PATH, bsResult);
							break;
						case 50055:
							if (nArg >= 2) {
								nLen = teGetMenuString(&bsResult, (HMENU)param[0], (UINT)param[1], param[2] != MF_BYCOMMAND);
							}
							break;
						case 50065: //GetDisplayNameOf
							try {
								if (nArg >= 1) {
									VARIANT *pv = &pDispParams->rgvarg[nArg];
									int i = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
									if (pv->vt == VT_BSTR) {
										if (!(i & SHGDN_INFOLDER)) {
											VariantCopy(pVarResult, pv);
											return S_OK;
										}
									}
									else if (FindUnknown(pv, &punk)) {
										BSTR bs;
										if (teGetBSTRFromFolderItem(&bs, punk)) {
											pVarResult->vt = VT_BSTR;
											if (i & SHGDN_INFOLDER) {
												pVarResult->bstrVal = ::SysAllocString(PathFindFileName(bs));
												return S_OK;
											}
											pVarResult->bstrVal = ::SysAllocString(bs);
											return S_OK;
										}
									}
									LPITEMIDLIST pidl;
									if (GetPidlFromVariant(&pidl, pv)) {
										GetVarPathFromPidl(pVarResult, pidl, i);
										teCoTaskMemFree(pidl);
									}
								}
							}
							catch(...) {
							}
							return S_OK;
						case 50075:
							if (nArg >= 0) {
								bsResult = SysAllocStringLen(NULL, MAX_PATH);
								nLen = GetKeyNameText((LONG)param[0], bsResult, MAX_PATH);
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
									nLen = teDragQueryFile((HDROP)param[0], (UINT)param[1], &bsResult);
								}
							}
							break;
						case 50095:
							if (nArg >= 0) {
								bsResult = SysAllocString(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]));
								nLen = -2;
							}
							break;
						case 50105:
							if (nArg >= 1) {
								bsResult = SysAllocStringLen(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), (UINT)param[1]);
								nLen = -2;
							}
							break;
						case 50115:
							if (nArg >= 1) {
								bsResult = SysAllocStringByteLen(reinterpret_cast<char*>(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg])), (UINT)param[1]);
								nLen = -2;
							}
							break;
						case 50125://sprintf_s
							if (nArg >= 2) {
								int nSize = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
								LPWSTR pszFormat = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]);
								if (pszFormat && nSize > 0) {
									bsResult = SysAllocStringLen(NULL, nSize);
									int nIndex = 1;
									int nPos = 0;
									while (pszFormat[nPos] && nLen < nSize) {
										while (pszFormat[nPos] && pszFormat[nPos++] != '%') {
										}
										while (WCHAR wc = pszFormat[nPos]) {
											nPos++;
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
												if (StrChrIW(L"s", wc)) {//String
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
												if (!StrChrIW(L"0123456789-+#.hl", wc)) {//not Specifier
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
								}
							}
							break;
						case 50135://base64_encode
							if (nArg >= 0) {
								if (lpfnCryptBinaryToStringW) {
									UCHAR *pc;
									GetpDataFromVariant(&pc, &nLen, &pDispParams->rgvarg[nArg]);
									if (pc) {
										DWORD dwSize;
										lpfnCryptBinaryToStringW(pc, nLen, CRYPT_STRING_BASE64, NULL, &dwSize);
										if (dwSize > 0) {
											bsResult = SysAllocStringLen(NULL, dwSize - 1);
											lpfnCryptBinaryToStringW(pc, nLen, CRYPT_STRING_BASE64, bsResult, &dwSize);
											nLen = -1;
										}
									}
								}
							}
							break;
						//AssocQueryString
						case 50155:
							if (nArg >= 3) {
								DWORD cch = 0;
								LPWSTR lp2 = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 2]);
								LPWSTR lp3 = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 3]);
								AssocQueryString((ASSOCF)param[0], (ASSOCSTR)param[1], lp2 , lp3, NULL, &cch);
								if (cch > 0) {
									nLen = cch - 1;
									bsResult = SysAllocStringLen(NULL, nLen);
									AssocQueryString((ASSOCF)param[0], (ASSOCSTR)param[1], lp2 , lp3, bsResult, &cch);
								}
							}
							break;
						//StrFormatByteSize
						case 50165:
							if (nArg >= 0) {
								nLen = (nArg >= 1) ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : MAX_COLUMN_NAME_LEN;
								if (nLen > 0) {
									bsResult = SysAllocStringLen(NULL, nLen);
									StrFormatByteSize(GetLLFromVariant(&pDispParams->rgvarg[nArg]), bsResult, nLen);
								}
								nLen = -1;
							}
							break;
						default:
							return DISP_E_MEMBERNOTFOUND;
					}
					pVarResult->vt = VT_BSTR;
					if (bsResult) {
						if (nLen < 0) {
							nLen = lstrlen(bsResult);
						}
						if (SysStringLen(bsResult) == nLen) {
							pVarResult->bstrVal = bsResult;
							return S_OK;
						}
					}
					pVarResult->bstrVal = SysAllocStringLen(bsResult, nLen);
					teSysFreeString(&bsResult);
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
					case 40123:
						if (dispIdMember == 40113) {
							*phResult = CreatePopupMenu();
						}
						else {
							*phResult = CreateMenu();
						}
						MENUINFO menuInfo;
						menuInfo.cbSize = sizeof(MENUINFO);
						menuInfo.fMask = MIM_STYLE;
						GetMenuInfo((HMENU)*phResult, &menuInfo);
						menuInfo.dwStyle = (menuInfo.dwStyle & ~MNS_NOCHECK) | MNS_CHECKORBMP;
						SetMenuInfo((HMENU)*phResult, &menuInfo);
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
						IUnknown *punk;
						if (nArg >= 0 && FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
							IUnknown_GetWindow(punk, (HWND *)phResult);
						}
						else if (nArg >= 1) {
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
					case 40273:
						if (nArg >= 0) {
							*phResult = (HANDLE)GetParent((HWND)param[0]);
						}
						break;
					case 40283:
						*phResult = (HANDLE)GetCapture();
						break;
					case 40293:
						VARIANT vFile;
						teVariantChangeType(&vFile, &pDispParams->rgvarg[nArg], VT_BSTR);
						*phResult = (HANDLE)GetModuleHandle(vFile.bstrVal);
						break;
					case 40303:
						if (nArg >= 0) {
							if FAILED(lpfnSHGetImageList((int)param[0], IID_IImageList, (LPVOID *)phResult)) {
								*phResult = 0;
							}
						}
						break;
					case 40313:
						if (nArg >= 4) {
							*phResult = CopyImage((HANDLE)param[0], (UINT)param[1], (int)param[2], (int)param[3], (UINT)param[4]);
						}
						break;
					case 40323:
						*phResult = GetCurrentProcess();
						break;
					case 40333:
						if (nArg >= 0) {
							*phResult = ImmGetContext((HWND)param[0]);
						}
						break;
					//other
					case 45003:
						if (nArg >= 1) {
							*phResult = LoadMenu((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]));
						}
						break;
					case 40343:
						*phResult = GetFocus();
						break;
					case 40353:
						*phResult = GetForegroundWindow();
						break;
					case 40363:
						if (nArg >= 2) {
							*phResult = (HANDLE)SetTimer((HWND)param[0], param[1], (UINT)param[2], NULL);
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
							if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
								GetPidlFromVariant((LPITEMIDLIST *)&vPath.bstrVal, &pDispParams->rgvarg[nArg]);
								vPath.vt = VT_NULL;
							}
							else {
								teVariantChangeType(&vPath, &pDispParams->rgvarg[nArg], VT_BSTR);
							}
							*phResult = (HANDLE)SHGetFileInfo(vPath.bstrVal, (DWORD)param[1],
								(SHFILEINFO *)GetpcFromVariant(&pDispParams->rgvarg[nArg - 2]),
								(UINT)param[3], (UINT)param[4]);
							if (vPath.vt == VT_NULL) {
								teCoTaskMemFree(vPath.bstrVal);
							}
							VariantClear(&vPath);
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
								(LPVOID)GetpcFromVariant(&pDispParams->rgvarg[nArg - 11]));
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
							*phResult = BeginPaint((HWND)param[0], (LPPAINTSTRUCT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1]));
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
									break;
								}
								FolderItem *pid;
								GetFolderItemFromVariant(&pid, &pInvoke->pv[pInvoke->cArgs - 1]);
								VariantClear(&pInvoke->pv[pInvoke->cArgs - 1]);
								teSetObjectRelease(&pInvoke->pv[pInvoke->cArgs - 1], pid);
								Invoke5(pInvoke->pdisp, DISPID_VALUE, DISPATCH_METHOD, NULL, pInvoke->cArgs, pInvoke->pv);
								pInvoke->pdisp->Release();
								*phResult = 0;
							}
							delete [] pInvoke;
						}
						break;
					case 45143: //SHChangeNotification_Lock
						if (nArg >= 2) {
							IUnknown *punk;
							if (FindUnknown(&pDispParams->rgvarg[nArg - 2], &punk)) {
								CteFolderItems *pFIs;
								if SUCCEEDED(punk->QueryInterface(g_ClsIdFIs, (LPVOID *)&pFIs)) {
									pFIs->Clear();
									PIDLIST_ABSOLUTE *ppidl;
									*phResult = SHChangeNotification_Lock((HANDLE)param[0], (DWORD)param[1], &ppidl, (LONG *)&pFIs->m_dwEffect);
									VARIANT v;
									teSetIDList(&v, ppidl[0]);
									pFIs->ItemEx(-1, NULL, &v);
									VariantClear(&v);
									teSetIDList(&v, ppidl[1]);
									pFIs->ItemEx(-1, NULL, &v);
									VariantClear(&v);
									pFIs->Release();
								}
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
					case 30222:
						if (nArg >= 2) {
							*plResult = SHEmptyRecycleBin((HWND)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]), (DWORD)param[2]);
						}
						break;
					case 30232:
						*plResult = GetMessagePos();
						break;
					case 30242:
						if (nArg >= 1) {
							try {
								*plResult = -1;
								((IImageList *)param[0])->GetOverlayImage((int)param[1], (int *)plResult);
							}
							catch (...) {
							}
						}
						break;
					case 30252:
						if (nArg >= 0) {
							*plResult = ImageList_GetBkColor((HIMAGELIST)param[0]);
						}
						break;
					case 30262:
						if (nArg >= 1 && lpfnSetWindowTheme) {
							*plResult = lpfnSetWindowTheme((HWND)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]), GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 2]));
						}
						break;
					case 30272:
						if (nArg >= 0) {
							*plResult = ImmGetVirtualKey((HWND)param[0]);
						}
						break;
					case 30282:
						if (nArg >= 0) {
							*plResult = GetAsyncKeyState((int)param[0]);
						}
						break;
					case 35002://TrackPopupMenuEx
						if (nArg >= 5) {
							if (nArg >= 6) {
								if (FindUnknown(&pDispParams->rgvarg[nArg - 6], &punk)) {
									punk->QueryInterface(IID_PPV_ARGS(&g_pCM));
								}
							}
							if (!g_hMenuKeyHook) {
								g_hMenuKeyHook = SetWindowsHookEx(WH_KEYBOARD, (HOOKPROC)MenuKeyProc, hInst, g_dwMainThreadId);
							}
							HWND hwnd = GetForegroundWindow();
							DWORD dwExStyle = 0;
							DWORD dwPid1, dwPid2;
							if (hwnd != g_hwndMain) {
								GetWindowThreadProcessId(g_hwndMain, &dwPid1);
								GetWindowThreadProcessId(hwnd, &dwPid2);
								if (dwPid1 != dwPid2) {
									dwExStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
									SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
									teSetForegroundWindow(g_hwndMain);
								}
							}
							*plResult = TrackPopupMenuEx((HMENU)param[0], (UINT)param[1], (int)param[2], (int)param[3],
								(HWND)param[4], (LPTPMPARAMS)GetpcFromVariant(&pDispParams->rgvarg[nArg - 5]));
							if (hwnd != g_hwndMain && dwPid1 != dwPid2) {
								if (!(dwExStyle & WS_EX_TOPMOST)) {
									SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
								}
								teSetForegroundWindow(hwnd);
							}
							UnhookWindowsHookEx(g_hMenuKeyHook);
							g_hMenuKeyHook = NULL;
							if (g_pCM) {
								g_pCM->Release();
								g_pCM = NULL;
							}
						}
						break;
					case 35012://ExtractIconEx
						if (nArg >= 4) {
							*plResult = ExtractIconEx(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), (int)param[1],
								(HICON *)GetpcFromVariant(&pDispParams->rgvarg[nArg - 2]),
								(HICON *)GetpcFromVariant(&pDispParams->rgvarg[nArg - 3]),
								(UINT)param[4]);
						}
						break;
					case 35022://GetObject
						if (nArg >= 2) {
							*plResult = GetObject((HWND)param[0], (int)param[1], GetpcFromVariant(&pDispParams->rgvarg[nArg - 2]));
						}
						break;
					case 35032:
						if (nArg >= 5) {
							*plResult = MultiByteToWideChar((UINT)param[0], (DWORD)param[1],
								(LPCSTR)GetpcFromVariant(&pDispParams->rgvarg[nArg - 2]), (int)param[3],
								(LPWSTR)GetpcFromVariant(&pDispParams->rgvarg[nArg - 4]), (int)param[5]);
						}
						break;
					case 35042:
						if (nArg >= 7) {
							VARIANT v;
							teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 2], VT_BSTR);
							*plResult = WideCharToMultiByte((UINT)param[0], (DWORD)param[1], v.bstrVal, (int)param[3], 
								GetpcFromVariant(&pDispParams->rgvarg[nArg - 4]), (int)param[5],
								(LPCSTR)GetpcFromVariant(&pDispParams->rgvarg[nArg - 6]),
								(LPBOOL)GetpcFromVariant(&pDispParams->rgvarg[nArg - 7]));
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
								if (ppidllist) {
									AdjustIDList(ppidllist, nCount);
									if (nCount >= 1) {
										IShellFolder *pSF;
										if (GetShellFolder(&pSF, ppidllist[0])) {
											pSF->GetAttributesOf(nCount, (LPCITEMIDLIST *)&ppidllist[1], (SFGAOF *)(plResult));
											*plResult &= (ULONG)param[1];
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
						}
						break;
					case 35062://DoDragDrop
						if (nArg >= 2) {
							IDataObject *pDataObj;
							if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
								g_pTE->m_bDrop = true;
								IDispatch *pEffect = NULL;
								DWORD dwEffect = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
								VARIANT v;
								VariantInit(&v);
								if (nArg >= 2 && GetDispatch(&pDispParams->rgvarg[nArg - 2], &pEffect)) {
									teGetPropertyAt(pEffect, 0, &v);
									dwEffect = GetIntFromVariantClear(&v);
								}
								try {
									*plResult = ::DoDragDrop(pDataObj, static_cast<IDropSource *>(g_pTE), (DWORD)param[1], &dwEffect);
								}
								catch(...) {
								}
								if (pEffect) {
									v.vt = VT_I4;
									v.lVal = dwEffect;
									tePutPropertyAt(pEffect, 0, &v);
									pEffect->Release();
								}
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
									IShellFolder *pDesktopFolder;
									SHGetDesktopFolder(&pDesktopFolder);
									*plResult = pDesktopFolder->CompareIDs(param[0], pidl1, pidl2);
									pDesktopFolder->Release();
									teCoTaskMemFree(pidl2);
								}
								teCoTaskMemFree(pidl1);
							}
						}
						break;
					//GetScriptDispatch
					case 35080:
					//ExecScript
					case 35082:
						if (nArg >= 1) {
							VARIANT v[2];
							for (int i = 2; i--;) {
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
								teSetObjectRelease(pVarResult, pdisp);
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
									teSetObjectRelease(pVarResult, oDisp);
								}
								pdisp->Release();
							}
						}
						return S_OK;
					//SHChangeNotifyRegister
					case 35100:
						if (nArg >= 5) {
							SHChangeNotifyEntry entry;
							GetPidlFromVariant(const_cast<LPITEMIDLIST *>(&entry.pidl), &pDispParams->rgvarg[nArg - 4]);
							entry.fRecursive = (BOOL)param[5];
							if (entry.pidl) {
								*plResult = SHChangeNotifyRegister((HWND)param[0], (int)param[1], (LONG)param[2], (UINT)param[3], 1, &entry);
								teCoTaskMemFree(const_cast<LPITEMIDLIST>(entry.pidl));
							}
						}
						break;
					case 35102:
						if (nArg >= 3) {
							*plResult = MessageBox((HWND)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]), GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 2]), (UINT)param[3]);
						}
						break;
					default:
						return DISP_E_MEMBERNOTFOUND;
				}
				SetReturnValue(pVarResult, dispIdMember, &Result);
				return S_OK;
			}
			if (dispIdMember == 29002) {//SHFileOperation
				if (nArg >= 4) {
					LPSHFILEOPSTRUCT pFO = new SHFILEOPSTRUCT[1];
					::ZeroMemory(pFO, sizeof(SHFILEOPSTRUCT));
					VARIANT v[2];
					for (int i = 2; i--;) {
						teVariantChangeType(&v[i], &pDispParams->rgvarg[nArg - 1 - i], VT_BSTR);
					}
					pFO->wFunc = (UINT)param[0];
					int nLen = ::SysStringLen(v[0].bstrVal);
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
							*plResult = teSHFileOperation(pFO);
						}
						catch (...) {
						}
						::SysFreeString(const_cast<BSTR>(pFO->pTo));
						::SysFreeString(const_cast<BSTR>(pFO->pFrom));
						delete [] pFO;
					}
					SetReturnValue(pVarResult, dispIdMember, &Result);
					return S_OK;
				}
			}
			if (dispIdMember >= 29000) {//BOOL MEM()
				int i = (dispIdMember / 100) % 5;
				if (nArg >= i) {
					char *pc = GetpcFromVariant(&pDispParams->rgvarg[nArg - i]);
					if (pc) {
						switch (dispIdMember) {
							case 29001:
								if (nArg >= 4) {
									*pbResult = PeekMessage((LPMSG)pc, (HWND)param[1], (UINT)param[2], (UINT)param[3], (UINT)param[4]);
								}
								break;
							case 29002:
								*plResult = teSHFileOperation((LPSHFILEOPSTRUCT)pc);
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
							case 29331:
								*pbResult = teChangeWindowMessageFilterEx((HWND)param[0], (UINT)param[1], (DWORD)param[2], (PCHANGEFILTERSTRUCT)pc);
								break;
							case 29601:
								if (nArg >= 2) {
									*pbResult = ImageList_GetIconSize((HIMAGELIST)param[0],
										(int *)pc, (int *)GetpcFromVariant(&pDispParams->rgvarg[nArg - 2]));
								}
								else {
									*pbResult = ImageList_GetIconSize((HIMAGELIST)param[0],
										(int *)&(((LPSIZE)pc)->cx), (int *)&(((LPSIZE)pc)->cy));
								}
								break;
							case 29611:
								*pbResult = GetMenuInfo((HMENU)param[0], (LPMENUINFO)pc);
								break;
							case 29621:
								*pbResult = SetMenuInfo((HMENU)param[0], (LPCMENUINFO)pc);
								break;
							default:
								return DISP_E_MEMBERNOTFOUND;
						}//end_switch
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
							*pbResult = teSetForegroundWindow((HWND)param[0]);
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
					case 20381:
						if (nArg >= 1) {
							*pbResult = SHFreeShared((HANDLE)param[0], (DWORD)param[1]);
						}
						break;
					case 20391:
						*pbResult = EndMenu();
						break;
					case 20401:
						if (nArg >= 0) {
							*pbResult = SHChangeNotifyDeregister((ULONG)param[0]);
						}
						break;
					case 20411:
						if (nArg >= 0) {
							*pbResult = SHChangeNotification_Unlock((HANDLE)param[0]);
						}
						break;
					case 20421:
						if (nArg >= 0 && lpfnIsWow64Process) {
							lpfnIsWow64Process((HANDLE)param[0], pbResult);
						}
						break;
					case 20431:
						if (nArg >= 8) {
							*pbResult = BitBlt((HDC)param[0], (int)param[1], (int)param[2], (int)param[3], (int)param[4],
								(HDC)param[5], (int)param[6], (int)param[7], (DWORD)param[8]);
						}
						break;
					case 20441:
						if (nArg >= 1) {
							*pbResult = ImmSetOpenStatus((HIMC)param[0], (BOOL)param[1]);
						}
						break;
					case 20451:
						if (nArg >= 1) {
							*pbResult = ImmReleaseContext((HWND)param[0], (HIMC)param[1]);
						}
						break;
					case 20461:
						if (nArg >= 1) {
							*pbResult = IsChild((HWND)param[0], (HWND)param[1]);
						}
						break;
					case 20471:
						if (nArg >= 1) {
							*pbResult = KillTimer((HWND)param[0], param[1]);
						}
						break;
					case 20481:
						if (nArg >= 0) {
							*pbResult = AllowSetForegroundWindow((DWORD)param[0]);
						}
						break;
					case 20491:
						if (nArg >= 6) {
							*pbResult = SetWindowPos((HWND)param[0], (HWND)param[1], (int)param[2], (int)param[3], (int)param[4], (int)param[5], (UINT)param[6]);
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
							if (!g_bSetRedraw && g_hwndMain == (HWND)param[0]) {
								*pbResult = FALSE;
							}
							else {
								*pbResult = RedrawWindow((HWND)param[0], (LPRECT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1]),
									(HRGN)param[2], (UINT)param[3]);
							}
						}
						break;
					case 25031:
						if (nArg >= 3) {
							*pbResult = MoveToEx((HDC)param[0], (int)param[1], (int)param[2], (LPPOINT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 3]));
						}
						break;
					case 25041:
						if (nArg >= 2) {
							*pbResult = InvalidateRect((HWND)param[0], (LPRECT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1]), (BOOL)param[2]);
						}
						break;
					case 25051:
						if (nArg >= 3) {
							*pbResult = SendNotifyMessage((HWND)param[0], (UINT)param[1], (WPARAM)param[2], (LPARAM)param[3]);
						}
						break;
					default:
						return DISP_E_MEMBERNOTFOUND;
				}//end_switch
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
						teCoTaskMemFree((LPVOID)param[0]);
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
				case 10080:
					if (nArg >= 3) {
						keybd_event((BYTE)param[0], (BYTE)param[1], (DWORD)param[2], (ULONG_PTR)param[3]);
					}
					break;
				case 10090://SHChangeNotify
					if (nArg >= 3) {
						VARIANT *pv = GetNewVARIANT(2);
						for (int i = 2; i--;) {
							pv[i].bstrVal = NULL;
							BYTE type = (BYTE)param[1];
							if (type == SHCNF_IDLIST) {
								GetPidlFromVariant((LPITEMIDLIST *)&pv[i].bstrVal, &pDispParams->rgvarg[nArg - i - 2]);
							}
							else if (type == SHCNF_PATH) {
								teVariantChangeType(&pv[i], &pDispParams->rgvarg[nArg - i - 2], VT_BSTR);
							}
						}
						SHChangeNotify((LONG)param[0], (UINT)param[1], pv[0].bstrVal, pv[1].bstrVal);
						for (int i = 2; i--;) {
							if (pv[i].vt != VT_BSTR) {
								teCoTaskMemFree(pv[i].bstrVal);
							}
							else {
								VariantClear(&pv[i]);
							}
						}
						delete pv;
					}
					break;
				default:
					return DISP_E_MEMBERNOTFOUND;
			}
			return S_OK;
		}
		if (dispIdMember >= 7000) {//(mem)
			if (nArg >= 0) {
				char *pc = GetpcFromVariant(&pDispParams->rgvarg[nArg]);
				if (pc) {
					switch(dispIdMember) {
						//BOOL
						case 7001:
							*pbResult = GetCursorPos((LPPOINT)pc);
							break;
						case 7011:
							*pbResult = GetKeyboardState((PBYTE)pc);
							break;
						case 7021:
							*pbResult = SetKeyboardState((PBYTE)pc);
							break;
						case 7031:
							*pbResult = lpfnRtlGetVersion ? !lpfnRtlGetVersion((PRTL_OSVERSIONINFOEXW)pc) : GetVersionEx((LPOSVERSIONINFO)pc);
							break;
						case 7041:
							*pbResult = ChooseFont((LPCHOOSEFONT)pc);
							break;
						case 7051:
							try {
								*pbResult = ::ChooseColor((LPCHOOSECOLOR)pc);
							} catch (...) {}
							break;
						case 7061:
							*pbResult = ::TranslateMessage((LPMSG)pc);
							break;
						case 7071:
							*pbResult = ::ShellExecuteEx((SHELLEXECUTEINFO *)pc);
							break;
						//int
						//Handle
						case 9003:
							*phResult = ::CreateFontIndirect((PLOGFONT)pc);
							break;
						case 9013:
							*phResult = ::CreateIconIndirect((PICONINFO)pc);
							break;
						case 9023:
							*phResult = ::FindText((LPFINDREPLACE)pc);
							break;
						case 9033:
							*phResult = ::ReplaceText((LPFINDREPLACE)pc);
							break;
						case 9043:
							*phResult = (HANDLE)::DispatchMessage((LPMSG)pc);
							break;
						default:
							return DISP_E_MEMBERNOTFOUND;
					}
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
						FolderItem *pid;
						GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
						teSetObjectRelease(pVarResult, pid);
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
							bsResult = ::SysAllocStringLen(v.bstrVal, ::SysStringLen(v.bstrVal) + 3);
							PathQuoteSpaces(bsResult);
							nLen = -1;
						}
						break;
					case 6015:
						if (v.bstrVal) {
							bsResult = ::SysAllocString(v.bstrVal);
							PathUnquoteSpaces(bsResult);
							nLen = -1;
						}
						break;
					case 6025:
						if (v.bstrVal) {
							nLen = GetShortPathName(v.bstrVal, NULL, 0);
							bsResult = ::SysAllocStringLen(NULL, nLen + MAX_PATH);
							nLen = GetShortPathName(v.bstrVal, bsResult, nLen);
						}
						break;
					case 6035:
						if (v.bstrVal) {
							DWORD dwLen = ::SysStringLen(v.bstrVal) + MAX_PATH;
							bsResult = ::SysAllocStringLen(NULL, dwLen);
							PathCreateFromUrl(v.bstrVal, bsResult, &dwLen, NULL);
							nLen = dwLen;
						}
						break;
					case 6045:
						if (v.bstrVal) {
							UINT uLen = ::SysStringLen(v.bstrVal) + MAX_PATH;
							bsResult = ::SysAllocStringLen(v.bstrVal, uLen);
							PathSearchAndQualify(v.bstrVal, bsResult, uLen);
							nLen = -1;
						}
						break;
			//PSFormatForDisplay
					case 6055:
						if (nArg >= 2 && lpfnPSPropertyKeyFromStringEx) {
							if (v.bstrVal) {
								PROPERTYKEY propKey;
								if SUCCEEDED(lpfnPSPropertyKeyFromStringEx(v.bstrVal, &propKey)) {
									if (lpfnPSGetPropertyDescription) {
										IPropertyDescription *pdesc;
										if SUCCEEDED(lpfnPSGetPropertyDescription(propKey, IID_PPV_ARGS(&pdesc))) {
											LPWSTR psz;
											if SUCCEEDED(pdesc->FormatForDisplay(*((PROPVARIANT *)&pDispParams->rgvarg[nArg - 1]), (PROPDESC_FORMAT_FLAGS)GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]), &psz)) {
												bsResult = ::SysAllocString(psz);
												nLen = -2;
												::CoTaskMemFree(psz);
											}
											pdesc->Release();
										}
									}
#ifdef _2000XP
									else {
										bsResult = ::SysAllocStringLen(NULL, MAX_PROP);
										IPropertyUI *pPUI;
										if SUCCEEDED(CoCreateInstance(CLSID_PropertiesUI, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&pPUI))) {
											pPUI->FormatForDisplay(propKey.fmtid, propKey.pid, (PROPVARIANT *)&pDispParams->rgvarg[nArg - 1], GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]), bsResult, MAX_PROP);
											pPUI->Release();
											nLen = -1;
										}
									}
#endif
								}
							}
						}
						break;
					//PSGetPropertyDescription
					case 6065:
						if (v.bstrVal && lpfnPSPropertyKeyFromStringEx) {
							PROPERTYKEY propKey;
							if SUCCEEDED(lpfnPSPropertyKeyFromStringEx(v.bstrVal, &propKey)) {
								if (lpfnPSGetPropertyDescription) {
									IPropertyDescription *pdesc;
									if SUCCEEDED(lpfnPSGetPropertyDescription(propKey, IID_PPV_ARGS(&pdesc))) {
										LPWSTR psz;
										if SUCCEEDED(pdesc->GetDisplayName(&psz)) {
											bsResult = ::SysAllocString(psz);
											nLen = -2;
											CoTaskMemFree(psz);
										}
										pdesc->Release();
									}
								}
#ifdef _2000XP
								else {
									bsResult = ::SysAllocStringLen(NULL, MAX_PROP);
									IPropertyUI *pPUI;
									if SUCCEEDED(CoCreateInstance(CLSID_PropertiesUI, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&pPUI))) {
										pPUI->GetDisplayName(propKey.fmtid, propKey.pid, PUIFNF_DEFAULT, bsResult, MAX_PROP);
										pPUI->Release();
										nLen = -1;
									}
								}
#endif
							}
						}
						break;
					default:
						VariantClear(&v);
						return DISP_E_MEMBERNOTFOUND;
				}
				VariantClear(&v);
			}
			int nMode = dispIdMember % 10;
			if (nMode >= 5) {
				pVarResult->vt = VT_BSTR;
				if (bsResult) {
					if (nLen < 0) {
						nLen = lstrlen(bsResult);
					}
					if (::SysStringLen(bsResult) == nLen) {
						pVarResult->bstrVal = bsResult;
						return S_OK;
					}
				}
				pVarResult->bstrVal = SysAllocStringLen(bsResult, nLen);
				teSysFreeString(&bsResult);
				return S_OK;
			}
			SetReturnValue(pVarResult, dispIdMember, &Result);
			return S_OK;
		}
		switch(dispIdMember) {
			//Memory
			case 1040:
				if (nArg >= 0) {
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
					pMem = new CteMemory(i * nCount + GetOffsetOfStruct(sz), pc, VT_NULL, nCount, sz);
					if (bs) {
						::CopyMemory(pMem->m_pc, bs, nCount * 2);
					}
					teSetObjectRelease(pVarResult, pMem);
				}
				return S_OK;
			//ContextMenu
			case 1061:
				if (nArg >= 0) {
					IDataObject *pDataObj = NULL;
					CteContextMenu *pCCM;
					if (nArg >= 6) {
						IShellExtInit *pSEI;
						HMODULE hDll = teCreateInstance(&pDispParams->rgvarg[nArg], &pDispParams->rgvarg[nArg - 1], IID_PPV_ARGS(&pSEI));
						if (pSEI) {
							LPITEMIDLIST pidl;
							if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg - 2])) {
								if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg - 3])) {
									VARIANT v;
									teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 5], VT_BSTR);
									HKEY hKey;
									if (RegOpenKeyEx((HKEY)GetLLFromVariant(&pDispParams->rgvarg[nArg - 4]), v.bstrVal, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
										VariantClear(&v);
										if SUCCEEDED(pSEI->Initialize(pidl, pDataObj, hKey)) {
											IUnknown *punk;
											if (FindUnknown(&pDispParams->rgvarg[nArg - 6], &punk)) {
												IUnknown_SetSite(pSEI, punk);
											}
											pCCM = new CteContextMenu(pSEI, pDataObj);
											pDataObj = NULL;
											pCCM->m_pDll = new CteDll(hDll);
											hDll = NULL;
											teSetObjectRelease(pVarResult, pCCM);
										}
										RegCloseKey(hKey);
									}
									VariantClear(&v);
								}
								teCoTaskMemFree(pidl);
							}
							pSEI->Release();
						}
						if (hDll) {
							FreeLibrary(hDll);
						}
					}
					else if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
						LPITEMIDLIST *ppidllist;
						long nCount;
						ppidllist = IDListFormDataObj(pDataObj, &nCount);
#ifdef _2000XP
						if (nCount >= 2 && !g_bUpperVista && ppidllist[0] && ILIsEmpty(ppidllist[0])) {
							for (int i = nCount; i-- > 0;) {
								LPITEMIDLIST pidl = ppidllist[i + 1];
								LPITEMIDLIST pidlFull = ILCombine(ppidllist[0], pidl);
								BSTR bs;
								if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidlFull, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) {
									ppidllist[i + 1] = teILCreateFromPath(bs);
									SysFreeString(bs);
								}
								teCoTaskMemFree(pidlFull);
								teCoTaskMemFree(pidl);
							}
							teILFreeClear(&ppidllist[0]);
						}
#endif
						AdjustIDList(ppidllist, nCount);
						if (nCount >= 1) {
							IShellFolder *pSF;
							if (GetShellFolder(&pSF, ppidllist[0])) {
								IContextMenu *pCM;
//								if SUCCEEDED(CDefFolderMenu_Create2(ppidllist[0], g_hwndMain, nCount, const_cast<LPCITEMIDLIST *>(&ppidllist[1]), pSF, NULL, 0, NULL, &pCM)) {
								if SUCCEEDED(pSF->GetUIObjectOf(g_hwndMain, nCount, const_cast<LPCITEMIDLIST *>(&ppidllist[1]), IID_IContextMenu, NULL, (LPVOID*)&pCM)) {
									if (nArg >= 1) {
										IUnknown *punk;
										if (FindUnknown(&pDispParams->rgvarg[nArg - 1], &punk)) {
											IUnknown_SetSite(pCM, punk);
										}
									}
									pCCM = new CteContextMenu(pCM, pDataObj);
									pDataObj = NULL;
									teSetObjectRelease(pVarResult, pCCM);
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
				return S_OK;

			//DropTarget
			case 1070:
				if (nArg >= 0) {
					FolderItem *pid;
					GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
					LPITEMIDLIST pidl;
					if (GetIDListFromObject(pid, &pidl)) {
						IShellFolder *pSF;
						LPCITEMIDLIST ItemID;
						if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &ItemID)) {
							IDropTarget *pDT = NULL;
							pSF->GetUIObjectOf(g_hwndMain, 1, &ItemID, IID_IDropTarget, NULL, (LPVOID*)&pDT);
							teSetObjectRelease(pVarResult, new CteDropTarget(pDT, pid));
							pDT && pDT->Release();
							pSF->Release();
						}
						teCoTaskMemFree(pidl);
					}
					pid->Release();
				}
				return S_OK;
			//DataObject
			case 1080:
				teSetObjectRelease(pVarResult, new CteFolderItems(NULL, NULL, true));
				return S_OK;
			case 1090:
				if (nArg >= 3) {
					IUnknown *punk;
					if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
						IOleCommandTarget *pCT;
						if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCT))) {
							GUID *pguid = NULL;
							GUID guidCmd;
							if (pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
								pguid = &guidCmd;
								CLSIDFromString(pDispParams->rgvarg[nArg - 1].bstrVal, pguid);
							}
							pCT->Exec(pguid, (OLECMDID)GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]),
								(OLECMDEXECOPT)GetIntFromVariant(&pDispParams->rgvarg[nArg - 3]),
								&pDispParams->rgvarg[nArg - 4],
								pVarResult
							);
							pCT->Release();
						}
					}
				}
				return S_OK;
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
				else if (nArg >= 0) {
					Result = GetLLFromVariant(&pDispParams->rgvarg[nArg]);
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
						IDispatch *pdisp1 = NULL;
						if (GetDispatch(&pDispParams->rgvarg[nArg - 2], &pdisp1)) {
							VARIANT v;
							VariantInit(&v);
							teGetProperty(pdisp1, L"length", &v);
							int nCount = GetIntFromVariantClear(&v);
							VARIANTARG *pv = GetNewVARIANT(nCount);
							for (int i = nCount; i-- > 0;) {
								teGetPropertyAt(pdisp1, nCount - i - 1, &pv[i]);//Reverse argument
							}
							teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
							teExecMethod(pdisp, v.bstrVal, pVarResult, -nCount, pv);
							VariantClear(&v);
							for (int i = nCount; i-- > 0;) {
								tePutPropertyAt(pdisp1, nCount - i - 1, &pv[i]);//for Reference
							}
							teClearVariantArgs(nCount, pv);
							pdisp1->Release();
						}
						pdisp->Release();
					}
				}
				break;
			//Extract
			case 1182:
				*plResult = E_FAIL;
				if (nArg >= 3) {
					IStorage *pStorage;
					HMODULE hDll = teCreateInstance(&pDispParams->rgvarg[nArg], &pDispParams->rgvarg[nArg - 1], IID_PPV_ARGS(&pStorage));
					if (pStorage) {
						VARIANT vZip, vDest;
						teVariantChangeType(&vZip, &pDispParams->rgvarg[nArg - 2], VT_BSTR);
						teVariantChangeType(&vDest, &pDispParams->rgvarg[nArg - 3], VT_BSTR);

						IPersistFile *pPF;
						IPersistFolder *pPF2;
						if SUCCEEDED(pStorage->QueryInterface(IID_PPV_ARGS(&pPF))) {
							if SUCCEEDED(pPF->Load(vZip.bstrVal, STGM_READ)) {
								*plResult = teExtract(pStorage, vDest.bstrVal);
							}
							pPF->Release();
						}
						else if SUCCEEDED(pStorage->QueryInterface(IID_PPV_ARGS(&pPF2))) {
							LPITEMIDLIST pidl = teILCreateFromPath(vZip.bstrVal);
							if (pidl) {
								if SUCCEEDED(pPF2->Initialize(pidl)) {
									*plResult = teExtract(pStorage, vDest.bstrVal);
								}
								::CoTaskMemFree(pidl);
							}
							pPF2->Release();
						}
						VariantClear(&vDest);
						VariantClear(&vZip);
						pStorage->Release();
					}
					hDll && FreeLibrary(hDll);
				}
				break;
			//Add
			case 1194:
				if (nArg >= 1) {
					Result = GetLLFromVariant(&pDispParams->rgvarg[nArg]) + GetLLFromVariant(&pDispParams->rgvarg[nArg - 1]);
				}
				break;
			//Sub
			case 1204:
				if (nArg >= 1) {
					Result = GetLLFromVariant(&pDispParams->rgvarg[nArg]) - GetLLFromVariant(&pDispParams->rgvarg[nArg - 1]);
				}
				break;
			//FindWindow
			case 2063:
				if (nArg >= 1) {
					VARIANT vClass, vWindow;
					teVariantChangeType(&vClass, &pDispParams->rgvarg[nArg], VT_BSTR);
					teVariantChangeType(&vWindow, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
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
			//OleGetClipboard
			case 2080:
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
			case 2082:
				if (nArg >= 0) {
					*plResult = E_FAIL;
					IDataObject *pDataObj;
					if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
						*plResult = OleSetClipboard(pDataObj);
						pDataObj->Release();
					}
				}
				break;
			//PathMatchSpec
			case 2111:
				if (nArg >= 1) {
					VARIANT vFile, vSpec;
					teVariantChangeType(&vFile, &pDispParams->rgvarg[nArg], VT_BSTR);
					teVariantChangeType(&vSpec, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
					*pbResult = tePathMatchSpec(vFile.bstrVal, vSpec.bstrVal);
					VariantClear(&vFile);
					VariantClear(&vSpec);
				}
				break;
			//CommandLineToArgv
			case 2230:
				if (nArg >= 0) {
					int nLen = 0;
					VARIANT vCmdLine;
					teVariantChangeType(&vCmdLine, &pDispParams->rgvarg[nArg], VT_BSTR);
					LPTSTR *lplpszArgs = NULL;
					IDispatch *pdisp;
					GetNewArray(&pdisp);
					if (vCmdLine.bstrVal && ::SysStringLen(vCmdLine.bstrVal)) {
						lplpszArgs = CommandLineToArgvW(vCmdLine.bstrVal, &nLen);
						for (int i = 0; i < nLen; i++) {
							VARIANT *pv = GetNewVARIANT(1);
							pv->bstrVal = ::SysAllocString(lplpszArgs[i]);
							pv->vt = VT_BSTR;
							int n = ::SysStringLen(pv->bstrVal);
							if (pv->bstrVal[n - 1] == '"') {
								pv->bstrVal[n - 1] = '\\';
							}
							teArrayPushVariantClear(pdisp, pv);
						}
						LocalFree(lplpszArgs);
					}
					VariantClear(&vCmdLine);
					teSetObjectRelease(pVarResult, pdisp);
				}
				return S_OK;
			//ILIsEqual
			case 2431:
				if (nArg >= 1) {
					FolderItem *pid1, *pid2;
					GetFolderItemFromVariant(&pid1, &pDispParams->rgvarg[nArg]);
					GetFolderItemFromVariant(&pid2, &pDispParams->rgvarg[nArg - 1]);
					*pbResult = teILIsEqual(pid1, pid2);
					pid2->Release();
					pid1->Release();
				}
				break;
			//ILIsParent
			case 2441:
				if (nArg >= 2) {
					LPITEMIDLIST pidl, pidl2;
					if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
						if (GetPidlFromVariant(&pidl2, &pDispParams->rgvarg[nArg - 1])) {
							*pbResult = ::ILIsParent(pidl, pidl2, GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]));
							teCoTaskMemFree(pidl2);
						}
						teCoTaskMemFree(pidl);
					}
				}
				break;
			//ILClone
			case 2440:
				if (pVarResult && nArg >= 0) {
					FolderItem *pid;
					GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
					teSetObjectRelease(pVarResult, pid);
				}
				break;
			//ILRemoveLastID
			case 2450:
				if (pVarResult && nArg >= 0) {
					LPITEMIDLIST pidl;
					if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
						if (!ILIsEmpty(pidl)) {
							if (::ILRemoveLastID(pidl)) {
								teSetIDList(pVarResult, pidl);
							}
						}
						teCoTaskMemFree(pidl);
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
							teSetIDList(pVarResult, pidl);
						}
						teCoTaskMemFree(pidl);
					}
				}
				break;
			//ILIsEmpty
			case 2461:
				*pbResult = TRUE;
				if (nArg >= 0) {
					LPITEMIDLIST pidl;
					if (GetPidlFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
						*pbResult = ILIsEmpty(pidl);
						teCoTaskMemFree(pidl);
					}
				}
				break;
			//ILGetParent
			case 2470:
				if (pVarResult && nArg >= 0) {
					FolderItem *pFolderItem, *pFolderItem1;
					GetFolderItemFromVariant(&pFolderItem, &pDispParams->rgvarg[nArg]);
					if (teILGetParent(pFolderItem, &pFolderItem1)) {
						teSetObjectRelease(pVarResult, pFolderItem1);
					}
					pFolderItem->Release();
				}
				break;
			//FindFirstFile(str, mem): handle
			case 2763:
				if (nArg >= 1) {
					VARIANT vFile;
					teVariantChangeType(&vFile, &pDispParams->rgvarg[nArg], VT_BSTR);
					*phResult = FindFirstFile(vFile.bstrVal, (LPWIN32_FIND_DATA)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1]));
				}
				break;
			case 2773:
				if (nArg >= 0) {
					POINT pt;
					GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg]);
					*phResult = WindowFromPoint(pt);
				}
				break;
/*/// Deprecated / Use sha.GetSystemInformation("DoubleClickTime")
			case 2782:
				*plResult = GetDoubleClickTime();
				break;
*///
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
								for (int i = 3; i--;) {
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
				teSetObject(pVarResult, this);
				return S_OK;
			default:
				return DISP_E_MEMBERNOTFOUND;
		}
		SetReturnValue(pVarResult, dispIdMember, &Result);
		return S_OK;
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"WindowsAPI", methodAPI, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
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
	try {
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (m_nMode) {
			if (dispIdMember == DISPID_NEWENUM) {
				teSetObject(pVarResult, this);
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
				teSetObject(pVarResult, this);
				return S_OK;
			}
			return m_pDispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
		if (wFlags & DISPATCH_METHOD) {
			return m_pDispatch->Invoke(m_dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
		teSetObject(pVarResult, this);
		return S_OK;
	}
	catch (...) {
		return teException(pExcepInfo, puArgErr, L"Dispatch", NULL, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
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
/*		else if (g_pWebBrowser && lstrcmpi(pstrName, L"te") == 0) {
			g_pTE->QueryInterface(IID_PPV_ARGS(ppiunkItem));
			return S_OK;
		}*/
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
		teSetObjectRelease(&pv[4], pei);
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
	*phwnd = g_hwndMain;
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::EnableModeless(BOOL fEnable)
{
	return E_NOTIMPL;
}

CteDll::CteDll(HMODULE hDll)
{
	m_hDll = hDll;
	m_cRef = 1;
}

CteDll::~CteDll()
{
	FreeLibrary(m_hDll);
}

STDMETHODIMP CteDll::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown)) {
		*ppvObject = static_cast<IUnknown *>(this);
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CteDll::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteDll::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		LPFNDllCanUnloadNow lpfnDllCanUnloadNow = (LPFNDllCanUnloadNow)GetProcAddress(m_hDll, "DllCanUnloadNow");
		if (g_pUnload && lpfnDllCanUnloadNow && lpfnDllCanUnloadNow()) {
			teArrayPush(g_pUnload, this);
			if (!g_bUnload) {
				g_bUnload = TRUE;
				SetTimer(g_hwndMain, TET_Unload, 1000 * 60 * 10, teTimerProc);
			}
			return m_cRef;
		}
		delete this;
		return 0;
	}
	return m_cRef;
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
	}
	else if (IsEqualIID(riid, IID_IServiceProvider)) {
		*ppvObject = static_cast<IServiceProvider *>(new CteServiceProvider(NULL, NULL));
		return S_OK;
	}
	else {
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
	}
	else {
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