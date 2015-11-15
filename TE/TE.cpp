// TE.cpp
// Tablacus Explorer (C)2011 Gaku
// MIT Lisence
// Visual C++ 2008 Express Edition SP1
// Windows SDK v7.0
// http://www.eonet.ne.jp/~gakana/tablacus/

#include "stdafx.h"
#include "TE.h"

// Global Variables:
HINSTANCE hInst;								// current instance
WCHAR	g_szTE[MAX_LOADSTRING];
WCHAR	g_szText[MAX_STATUS];
WCHAR	g_szStatus[MAX_STATUS];
WCHAR	g_szTitle[MAX_STATUS];
HWND	g_hwndMain = NULL;
CteTabCtrl *g_pTC = NULL;
CteTabCtrl *g_ppTC[MAX_TC];
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
LPFNPSStringFromPropertyKey lpfnPSStringFromPropertyKey = NULL;
LPFNSHGetIDListFromObject lpfnSHGetIDListFromObject = NULL;
LPFNIsWow64Process lpfnIsWow64Process = NULL;
LPFNChangeWindowMessageFilter lpfnChangeWindowMessageFilter = NULL;
LPFNChangeWindowMessageFilterEx lpfnChangeWindowMessageFilterEx = NULL;
LPFNAddClipboardFormatListener lpfnAddClipboardFormatListener = NULL;
LPFNRemoveClipboardFormatListener lpfnRemoveClipboardFormatListener = NULL;

LPFNRtlGetVersion lpfnRtlGetVersion = NULL;
LPFNSHTestTokenMembership lpfnSHTestTokenMembership = NULL;
#ifdef _USE_APIHOOK
LPFNRegQueryValueExW lpfnRegQueryValueExW = NULL;
#endif

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
IUIAutomation *g_pAutomation;
PROPERTYID	g_PID_ItemIndex;

CTE			*g_pTE;
CteShellBrowser *g_ppSB[MAX_FV];
CteWebBrowser *g_pWebBrowser;
IDispatch	*g_pAPI;

LPITEMIDLIST g_pidls[MAX_CSIDL];
LPITEMIDLIST g_pidlResultsFolder;
LPITEMIDLIST g_pidlLibrary = NULL;
BSTR		 g_bsPidls[MAX_CSIDL];

IDispatch	*g_pOnFunc[Count_OnFunc];
IDispatch	*g_pUnload  = NULL;
IDispatch	*g_pFreeLibrary = NULL;
IDispatch	*g_pJS		= NULL;
IDispatch	*g_pSubWindows = NULL;
IDropTargetHelper *g_pDropTargetHelper = NULL;
IUnknown	*g_pRelease = NULL;
IUnknown	*g_pDraggingCtrl = NULL;
IDataObject	*g_pDraggingItems = NULL;

IDispatchEx *g_pCM = NULL;
ULONG_PTR g_Token;
Gdiplus::GdiplusStartupInput g_StartupInput;
HHOOK	g_hHook;
HHOOK	g_hMouseHook;
HHOOK	g_hMessageHook;
HHOOK	g_hMenuKeyHook = NULL;
HMENU	g_hMenu = NULL;
BSTR	g_bsCmdLine = NULL;
BSTR	g_bsDocumentWrite = NULL;
BSTR	g_bsClipRoot = NULL;
HANDLE g_hMutex = NULL;

UINT	g_uCrcTable[256];
LONG	g_nSize = MAXWORD;
LONG	g_nLockUpdate = 0;
LONG	nUpdate = 0;
DWORD	g_paramFV[SB_Count];
DWORD	g_dwMainThreadId;
DWORD	g_dwTickScroll;
DWORD	g_dwSWCookie = 0;
DWORD	g_dwHideFileExt = 0;
long	g_nProcKey	   = 0;
long	g_nProcMouse   = 0;
long	g_nProcDrag    = 0;
long	g_nProcFV      = 0;
long	g_nProcTV      = 0;
long	g_nProcFI      = 0;
long	g_nThreads	   = 0;
long	g_lSWCookie = 0;
int		g_nReload = 0;
int		g_nException = 64;
int		*g_maps[MAP_LENGTH];
int		g_param[Count_TE_params];
int		g_x = MAXINT;
int		g_nPidls = MAX_CSIDL;
int		g_nTCCount = 0;
int		g_nTCIndex = 0;
BOOL	g_bLabelsMode;
BOOL	g_bMessageLoop = TRUE;
BOOL	g_bDialogOk = FALSE;
BOOL	g_bInit = FALSE;
BOOL	g_bUnload = FALSE;
BOOL	g_bSetRedraw;
BOOL	g_bShowParseError = TRUE;
BOOL	g_bDropFinished = TRUE;
#ifdef _2000XP
int		g_nCharWidth = 7;
BOOL	g_bCharWidth = true;
BOOL	g_bUpperVista;
BOOL	g_bIsXP;
HWND	g_hwndNextClip = NULL;
LPWSTR  g_szTotalFileSizeXP = GetUserDefaultLCID() == 1041 ? L"総ファイル サイズ" : L"Total file size";
LPWSTR  g_szLabelXP = GetUserDefaultLCID() == 1041 ? L"ラベル" : L"Label";
LPWSTR	g_szTotalFileSizeCodeXP = L"System.TotalFileSize";
LPWSTR	g_szLabelCodeXP = L"System.Contact.Label";
#endif
#ifdef _W2000
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

TEmethod tesMONITORINFOEX[] =
{
	{ (VT_I4 << TE_VT), L"cbSize" },
	{ (VT_CARRAY << TE_VT) + offsetof(MONITORINFOEX, rcMonitor), L"rcMonitor" },
	{ (VT_CARRAY << TE_VT) + offsetof(MONITORINFOEX, rcWork), L"rcWork" },
	{ (VT_I4 << TE_VT) + offsetof(MONITORINFOEX, dwFlags), L"dwFlags" },
	{ (VT_LPWSTR << TE_VT) + offsetof(MONITORINFOEX, szDevice), L"szDevice" },
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

TEmethod tesNMCUSTOMDRAW[] =
{
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, hdr), L"hdr" },
	{ (VT_I4 << TE_VT) + offsetof(NMCUSTOMDRAW, dwDrawStage), L"dwDrawStage" },
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, hdc), L"hdc" },
	{ (VT_CARRAY << TE_VT) + offsetof(NMCUSTOMDRAW, rc), L"rc" },
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, dwItemSpec), L"dwItemSpec" },
	{ (VT_I4 << TE_VT) + offsetof(NMCUSTOMDRAW, uItemState), L"uItemState" },
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, lItemlParam), L"lItemlParam" },
	{ 0, NULL }
};

TEmethod tesNMLVCUSTOMDRAW[] =
{
	{ (VT_PTR << TE_VT) + offsetof(NMLVCUSTOMDRAW, nmcd), L"nmcd" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, clrText), L"clrText" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, clrTextBk), L"clrTextBk" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iSubItem), L"iSubItem" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, dwItemType), L"dwItemType" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, clrFace), L"clrFace" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iIconEffect), L"iIconEffect" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iIconPhase), L"iIconPhase" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iPartId), L"iPartId" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iStateId), L"iStateId" },
	{ (VT_CARRAY << TE_VT) + offsetof(NMLVCUSTOMDRAW, rcText), L"rcText" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, uAlign), L"uAlign" },
	{ 0, NULL }
};

TEmethod tesNMTVCUSTOMDRAW[] =
{
	{ (VT_PTR << TE_VT) + offsetof(NMTVCUSTOMDRAW, nmcd), L"nmcd" },
	{ (VT_I4 << TE_VT) + offsetof(NMTVCUSTOMDRAW, clrText), L"clrText" },
	{ (VT_I4 << TE_VT) + offsetof(NMTVCUSTOMDRAW, clrTextBk), L"clrTextBk" },
	{ (VT_I4 << TE_VT) + offsetof(NMTVCUSTOMDRAW, iLevel), L"iLevel" },
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

TEStruct pTEStructs[] = {
	{ sizeof(BITMAP), L"BITMAP", tesBITMAP },
	{ sizeof(BSTR), L"BSTR", tesNULL },
	{ sizeof(BYTE), L"BYTE", tesNULL },
	{ sizeof(char), L"char", tesNULL },
	{ sizeof(CHOOSECOLOR), L"CHOOSECOLOR", tesCHOOSECOLOR },
	{ sizeof(CHOOSEFONT), L"CHOOSEFONT", tesCHOOSEFONT },
	{ sizeof(COPYDATASTRUCT), L"COPYDATASTRUCT", tesCOPYDATASTRUCT },
	{ sizeof(DIBSECTION), L"DIBSECTION", tesDIBSECTION },
	{ sizeof(DRAWITEMSTRUCT), L"DRAWITEMSTRUCT", tesDRAWITEMSTRUCT },
	{ sizeof(DWORD), L"DWORD", tesNULL },
	{ sizeof(EncoderParameter), L"EncoderParameter", tesEncoderParameter },
	{ sizeof(EncoderParameter), L"EncoderParameters", tesEncoderParameters },
	{ sizeof(EXCEPINFO), L"EXCEPINFO", tesEXCEPINFO },
	{ sizeof(FINDREPLACE), L"FINDREPLACE", tesFINDREPLACE },
	{ sizeof(FOLDERSETTINGS), L"FOLDERSETTINGS", tesFOLDERSETTINGS },
	{ sizeof(GUID), L"GUID", tesNULL },
	{ sizeof(HANDLE), L"HANDLE", tesNULL },
	{ sizeof(HDITEM), L"HDITEM", tesHDITEM },
	{ sizeof(ICONINFO), L"ICONINFO", tesICONINFO },
	{ sizeof(ICONMETRICS), L"ICONMETRICS", tesICONMETRICS },
	{ sizeof(int), L"int", tesNULL },
	{ sizeof(KEYBDINPUT) + sizeof(DWORD), L"KEYBDINPUT", tesKEYBDINPUT },
	{ 256, L"KEYSTATE", tesNULL },
	{ sizeof(LOGFONT), L"LOGFONT", tesLOGFONT },
	{ sizeof(LPWSTR), L"LPWSTR", tesNULL },
	{ sizeof(LVBKIMAGE), L"LVBKIMAGE", tesLVBKIMAGE },
	{ sizeof(LVFINDINFO), L"LVFINDINFO", tesLVFINDINFO },
	{ sizeof(LVGROUP), L"LVGROUP", tesLVGROUP },
	{ sizeof(LVHITTESTINFO), L"LVHITTESTINFO", tesLVHITTESTINFO },
	{ sizeof(LVITEM), L"LVITEM", tesLVITEM },
	{ sizeof(MEASUREITEMSTRUCT), L"MEASUREITEMSTRUCT", tesMEASUREITEMSTRUCT },
	{ sizeof(MENUINFO), L"MENUINFO", tesMENUINFO },
	{ sizeof(MENUITEMINFO), L"MENUITEMINFO", tesMENUITEMINFO },
	{ sizeof(MONITORINFOEX), L"MONITORINFOEX", tesMONITORINFOEX },
	{ sizeof(MOUSEINPUT) + sizeof(DWORD), L"MOUSEINPUT", tesMOUSEINPUT },
	{ sizeof(MSG), L"MSG", tesMSG },
	{ sizeof(NMCUSTOMDRAW), L"NMCUSTOMDRAW", tesNMCUSTOMDRAW },
	{ sizeof(NMLVCUSTOMDRAW), L"NMLVCUSTOMDRAW", tesNMLVCUSTOMDRAW },
	{ sizeof(NMTVCUSTOMDRAW), L"NMTVCUSTOMDRAW", tesNMTVCUSTOMDRAW },
	{ sizeof(NMHDR), L"NMHDR", tesNMHDR },
	{ sizeof(NONCLIENTMETRICS), L"NONCLIENTMETRICS", tesNONCLIENTMETRICS },
	{ sizeof(NOTIFYICONDATA), L"NOTIFYICONDATA", tesNOTIFYICONDATA },
	{ sizeof(OSVERSIONINFO), L"OSVERSIONINFO", tesOSVERSIONINFOEX },
	{ sizeof(OSVERSIONINFOEX), L"OSVERSIONINFOEX", tesOSVERSIONINFOEX },
	{ sizeof(PAINTSTRUCT), L"PAINTSTRUCT", tesPAINTSTRUCT },
	{ sizeof(POINT), L"POINT", tesPOINT },
	{ sizeof(RECT), L"RECT", tesRECT },
	{ sizeof(SHELLEXECUTEINFO), L"SHELLEXECUTEINFO", tesSHELLEXECUTEINFO },
	{ sizeof(SHFILEINFO), L"SHFILEINFO", tesSHFILEINFO },
	{ sizeof(SHFILEOPSTRUCT), L"SHFILEOPSTRUCT", tesSHFILEOPSTRUCT },
	{ sizeof(SIZE), L"SIZE", tesSIZE },
	{ sizeof(TCHITTESTINFO), L"TCHITTESTINFO", tesTCHITTESTINFO },
	{ sizeof(TCITEM), L"TCITEM", tesTCITEM },
	{ sizeof(TVHITTESTINFO), L"TVHITTESTINFO", tesTVHITTESTINFO },
	{ sizeof(TVITEM), L"TVITEM", tesTVITEM },
	{ sizeof(VARIANT), L"VARIANT", tesNULL },
	{ sizeof(WCHAR), L"WCHAR", tesNULL },
	{ sizeof(WIN32_FIND_DATA), L"WIN32_FIND_DATA", tesWIN32_FIND_DATA },
	{ sizeof(WORD), L"WORD", tesNULL },
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
	{ TE_METHOD + 1100, L"HookDragDrop" },//Deprecated
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
	{ TE_OFFSET + TE_Layout, L"Layout" },
	{ TE_OFFSET + TE_NetworkTimeout, L"NetworkTimeout" },
	{ START_OnFunc + TE_Labels, L"Labels" },
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
	{ START_OnFunc + TE_OnViewModeChanged, L"OnViewModeChanged" },
	{ START_OnFunc + TE_OnColumnsChanged, L"OnColumnsChanged" },
	{ START_OnFunc + TE_OnItemPrePaint, L"OnItemPrePaint" },
	{ START_OnFunc + TE_OnColumnClick, L"OnColumnClick" },
	{ START_OnFunc + TE_OnBeginDrag, L"OnBeginDrag" },
	{ START_OnFunc + TE_OnBeforeGetData, L"OnBeforeGetData" },
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
	{ 0x10000013, L"SizeFormat" },
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
	{ 0x10000058, L"Droptarget" },
	{ 0x10000059, L"Columns"},
	{ 0x10000102, L"hwndList" },
	{ 0x10000103, L"hwndView" },
	{ 0x10000104, L"SortColumn" },
	{ 0x10000105, L"GroupBy" },
	{ 0x10000106, L"Focus" },
	{ 0x10000107, L"HitTest" },
	{ 0x10000110, L"ItemCount" },
	{ 0x10000111, L"Item" },
	{ 0x10000206, L"Refresh" },
	{ 0x10000207, L"ViewMenu" },
	{ 0x10000208, L"TranslateAccelerator" },
	{ 0x10000209, L"GetItemPosition" },
	{ 0x1000020A, L"SelectAndPositionItem" },
	{ 0x10000280, L"SelectItem" },
	{ 0x10000281, L"FocusedItem" },
	{ 0x10000282, L"GetFocusedItem" },
	{ 0x10000283, L"GetItemRect" },
	{ 0x10000300, L"Notify" },
	{ 0x10000501, L"AddItem" },
	{ 0x10000502, L"RemoveItem" },
	{ 0x10000503, L"AddItems" },
	{ 0x10000504, L"RemoveAll" },
	{ 0x10000505, L"SessionId" },
	{ START_OnFunc + SB_TotalFileSize, L"TotalFileSize" },
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
	{ 0x10000007, L"Window" },
	{ 0x10000008, L"Focus" },
//	{ 0x10000009, L"Close" },
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
	{ 13, L"LockUpdate" },
	{ 14, L"UnlockUpdate" },
	{ DISPID_NEWENUM, L"_NewEnum" },
	{ DISPID_TE_ITEM, L"Item" },
	{ DISPID_TE_COUNT, L"Count" },
	{ DISPID_TE_COUNT, L"length" },
	{ TE_OFFSET + TE_Left, L"Left" },
	{ TE_OFFSET + TE_Top, L"Top" },
	{ TE_OFFSET + TE_Width, L"Width" },
	{ TE_OFFSET + TE_Height, L"Height" },
	{ TE_OFFSET + TE_Flags, L"Style" },
	{ TE_OFFSET + TC_Align, L"Align" },
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
//	{ 0x10000001, L"lEvent" },
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
	{ 0x10000106, L"Focus" },
	{ 0x10000107, L"HitTest" },
	{ 0x10000283, L"GetItemRect" },
	{ 0x10000300, L"Notify" },
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
	{ 5, L"Unavailable" },
	{ 9, L"_BLOB" },	//To be necessary
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
	{ 112, L"GetPixel" },
	{ 113, L"SetPixel" },
	{ 120, L"GetThumbnailImage" },
	{ 130, L"RotateFlip" },

	{ 210, L"GetHBITMAP" },
	{ 211, L"GetHICON" },
	{ 0, NULL }
};

// Forward declarations of functions included in this code module:
ATOM MyRegisterClass(HINSTANCE hInstance, LPWSTR szClassName, WNDPROC lpfnWndProc);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
VOID ArrangeWindow();

BOOL teGetIDListFromObject(IUnknown *punk, LPITEMIDLIST *ppidl);
VOID CALLBACK teTimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
VOID CALLBACK teTimerProcParse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
static void threadParseDisplayName(void *args);
LPITEMIDLIST teILCreateFromPath(LPWSTR pszPath);

//Unit

LONGLONG teGetU(LONGLONG ll)
{
	if ((ll & (~MAXINT)) == (~MAXINT)) {
		return ll & MAXUINT;
	}
	return ll & MAXINT64;
}

VOID teReleaseClear(PVOID ppObj)
{
	try {
		IUnknown **ppunk = static_cast<IUnknown **>(ppObj);
		if (*ppunk) {
			(*ppunk)->Release();
			*ppunk = NULL;
		}
	} catch (...) {
		g_nException = 0;
	}
}

BOOL teShowWindow(HWND hWnd, int nCmdShow)
{
	ANIMATIONINFO ai;
	ai.cbSize = sizeof(ANIMATIONINFO);
	int iMinAnimate = 0;
	if (nCmdShow != SW_SHOWMINIMIZED || IsIconic(hWnd)) {
		if (SystemParametersInfo(SPI_GETANIMATION, sizeof(ai), &ai, 0)) {
			iMinAnimate = ai.iMinAnimate;
		}
		if (iMinAnimate) {
			ai.iMinAnimate = FALSE;
			SystemParametersInfo(SPI_SETANIMATION, sizeof(ai), &ai, 0);
		}
	}
	BOOL bResult = ShowWindow(hWnd, nCmdShow);
	if (iMinAnimate) {
		ai.iMinAnimate = iMinAnimate;
		SystemParametersInfo(SPI_SETANIMATION, sizeof(ai), &ai, 0);
	}
	return bResult;
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
	LPWSTR pszPrefix = L"\0KMGTPE\0";
	int i = (dwFormat >> 4) & 7;
	int j = i;
	LONGLONG llPow = 1;
	if (i < 7) {
		while (i-- > 0) {
			llPow *= 1024;
		}
	}
	WCHAR pszNum[32];
	swprintf_s(pszNum, 32, L"%f", double(qdw) / llPow);
	teCommaSize(pszNum, pszBuf, cchBuf, dwFormat & 0xf);
	swprintf_s(&pszBuf[lstrlen(pszBuf)], 4, L" %cB", pszPrefix[j]);
}

VOID teGetWinSettings()
{
	HKEY hKey;
	if (RegOpenKeyEx(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
		DWORD dwSize = sizeof(g_dwHideFileExt);
		RegQueryValueEx(hKey, L"HideFileExt", NULL, NULL, (LPBYTE)&g_dwHideFileExt, &dwSize);
		RegCloseKey(hKey);
	}
}

VOID teSysFreeString(BSTR *pbs)
{
	if (*pbs) {
		::SysFreeString(*pbs);
		*pbs = NULL;
	}
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

VOID teSetBSTR(VARIANT *pv, BSTR bs, int nLen)
{
	if (pv) {
		pv->vt = VT_BSTR;
		if (bs) {
			if (nLen < 0) {
				nLen = lstrlen(bs);
			}
			if (::SysStringLen(bs) == nLen) {
				pv->bstrVal = bs;
				return;
			}
		}
		pv->bstrVal = SysAllocStringLen(bs, nLen);
		teSysFreeString(&bs);
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

BSTR tePSGetNameFromPropertyKeyEx(PROPERTYKEY propKey, int nFormat, IShellView *pSV)
{
	if (nFormat == 2) {
		WCHAR szProp[64];
		if (lpfnPSStringFromPropertyKey) {
			lpfnPSStringFromPropertyKey(propKey, szProp, 64);
		}
#ifdef _2000XP
		else {
			StringFromGUID2(propKey.fmtid, szProp, 39);
			wchar_t pszId[8];
			swprintf_s(pszId, 8, L" %u", propKey.pid);
			lstrcat(szProp, pszId);
		}
#endif
		return ::SysAllocString(szProp);
	}
	if (lpfnPSGetPropertyDescription) {
		BSTR bs = NULL;
		IPropertyDescription *pdesc;
		if SUCCEEDED(lpfnPSGetPropertyDescription(propKey, IID_PPV_ARGS(&pdesc))) {
			LPWSTR psz = NULL;
			CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_NAME };
			cmci.wszName[0] = NULL;
			if (nFormat) {
				pdesc->GetCanonicalName(&psz);
			} else {
				IColumnManager *pColumnManager;
				if (pSV && SUCCEEDED(pSV->QueryInterface(IID_PPV_ARGS(&pColumnManager)))) {
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
				CoTaskMemFree(psz);
			}
			pdesc->Release();
		} else {
			bs = tePSGetNameFromPropertyKeyEx(propKey, 2, NULL);
		}
		return bs;
	}
#ifdef _2000XP
	if (IsEqualPropertyKey(propKey, PKEY_TotalFileSize)) {
		return ::SysAllocString(nFormat ? g_szTotalFileSizeCodeXP : g_szTotalFileSizeXP);
	}
	if (IsEqualPropertyKey(propKey, PKEY_Contact_Label)) {
		return ::SysAllocString(nFormat ? g_szLabelCodeXP : g_szLabelXP);
	}
	WCHAR szProp[MAX_PROP];
	IPropertyUI *pPUI;
	if SUCCEEDED(CoCreateInstance(CLSID_PropertiesUI, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&pPUI))) {
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

BOOL teGetStrFromFolderItem(BSTR *pbs, IUnknown *punk)
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
			}
		}
	} catch (...) {}
	g_pSW = NULL;
}

VOID teSetRedraw(BOOL bSetRedraw)
{
	if (g_pWebBrowser) {
		SendMessage(g_pWebBrowser->m_hwndBrowser, WM_SETREDRAW, bSetRedraw, 0);
		if (bSetRedraw && !g_bSetRedraw) {
			RedrawWindow(g_pWebBrowser->m_hwndBrowser, NULL, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
		}
	}
	g_bSetRedraw = bSetRedraw;
}

HRESULT teExceptionEx(EXCEPINFO *pExcepInfo, LPWSTR pszObj, LPCWSTR pszName)
{
	if (pExcepInfo) {
		int nLen = lstrlen(pszObj) + lstrlen(pszName) + 15;
		pExcepInfo->bstrDescription = SysAllocStringLen(NULL, nLen);
		swprintf_s(pExcepInfo->bstrDescription, nLen, L"Exception in %s.%s", pszObj, pszName ? pszName : L"");
		pExcepInfo->bstrSource = SysAllocString(g_szTE);
		pExcepInfo->scode = DISP_E_EXCEPTION;

	}
	return DISP_E_EXCEPTION;
}

HRESULT teException(EXCEPINFO *pExcepInfo, LPWSTR pszObj, TEmethod* pMethod, DISPID dispIdMember)
{
	LPWSTR pszName = NULL;
	if (pMethod) {
		int i = 0;
		while (pMethod[i].id && pMethod[i].id != dispIdMember) {
			i++;
		}
		pszName = pMethod[i].name;
	}
	return teExceptionEx(pExcepInfo, pszObj, pszName);
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
		try {
			::CoTaskMemFree(pv);
		} catch (...) {}
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
	if (ppidl) {
		teCoTaskMemFree(*ppidl);
		*ppidl = NULL;
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
			return (tolower(pszFile[0]) == tolower(pszSpec[0])) && tePathMatchSpec2(pszFile + 1, pszSpec + 1);
	}
}
#endif

BOOL teStartsText(LPWSTR pszSub, LPCWSTR pszFile)
{
	BOOL bResult = pszFile ? TRUE : FALSE;
	WCHAR wc;
	while (bResult && (wc = *pszSub++)) {
		bResult = tolower(wc) == tolower(*pszFile++);
	}
	return bResult;
}

BOOL tePathMatchSpec1(LPCWSTR pszFile, LPWSTR pszSpec)
{
	WCHAR wc = *pszSpec;
	BOOL bResult = *pszFile != NULL || !wc || wc == '*';
	for (; bResult && *pszFile; pszFile++) {
		wc = *pszSpec++;
		if (wc == '*') {
			wc = tolower(*pszSpec++);
			do {
				if (!wc || wc == ';') {
					return TRUE; 
				}
				for (; *pszFile && tolower(*pszFile) != wc; pszFile++);
				if (!*pszFile) {
					return FALSE; 
				}
				bResult = tePathMatchSpec1(++pszFile, pszSpec);
			} while (!bResult);
			return bResult;
		}
		if (!wc || wc == ';') {
			break; 
		}
		if (wc != '?') {
			bResult = tolower(*pszFile) == tolower(wc);
		}
	}
	while ((wc = *pszSpec) == '*') {
		pszSpec++;
	}
	return bResult && (*pszFile == (wc == ';' ? NULL : wc));
}

BOOL tePathMatchSpec(LPCWSTR pszFile, LPCWSTR pszSpec)
{
	if (!pszSpec || !pszSpec[0]) {
		return TRUE;
	}
	if (!pszFile) {
		return FALSE;
	}
	LPWSTR pszSpec1 = const_cast<LPWSTR>(pszSpec);
	do {
#ifdef _USE_TESTPATHMATCHSPEC
		BOOL b1 = !!tePathMatchSpec1(pszFile, pszSpec1);
		BOOL b2 = !!tePathMatchSpec2(pszFile, pszSpec1);
		if (b1 != b2) {
			b2 = !!tePathMatchSpec1(pszFile, pszSpec1);
		}
#endif
		if (tePathMatchSpec1(pszFile, pszSpec1)) {
			return TRUE;
		}
		pszSpec1 = StrChr(pszSpec1, ';') + 1;
	} while (pszSpec1 > LPWSTR(sizeof(WCHAR)) && *pszSpec1);
	return FALSE;
}

BOOL tePathIsNetworkPath(LPCWSTR pszPath)//PathIsNetworkPath is slow in DRIVE_NO_ROOT_DIR.
{
	WCHAR pszDrive[4];
	lstrcpyn(pszDrive, pszPath, 4);
	if (pszDrive[0] >= 'A' && pszDrive[1] == ':' && pszDrive[2] == '\\') {
		UINT uDriveType = GetDriveType(pszDrive);
		return uDriveType == DRIVE_REMOTE || uDriveType == DRIVE_NO_ROOT_DIR;
	}
	return (pszDrive[0] == '\\' && pszDrive[1] == '\\');
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

VOID tePathAppend(BSTR *pbsPath, LPCWSTR pszPath, LPWSTR pszFile)
{
	*pbsPath = SysAllocStringLen(pszPath, lstrlen(pszPath) + lstrlen(pszFile) + 1);
	PathAppend(*pbsPath, pszFile);
}

ULONGLONG teGetFolderSize(LPCWSTR szPath, int nLevel, LPITEMIDLIST *ppidl, LPITEMIDLIST pidl)
{
	if (*ppidl != pidl) {
		return 0;
	}
	ULONGLONG Result = 0;
	BSTR bs;
	WIN32_FIND_DATA wfd;
	tePathAppend(&bs, szPath, L"*");
	HANDLE hFind = FindFirstFile(bs, &wfd);
	if (hFind != INVALID_HANDLE_VALUE) {
		do {
			if (tePathMatchSpec(wfd.cFileName, L".;..")) {
				continue;
			}
			if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
				if (nLevel) {
					::SysFreeString(bs);
					tePathAppend(&bs, szPath, wfd.cFileName);
					Result += teGetFolderSize(bs, nLevel - 1, ppidl, pidl);
				}
				continue;
			}
			ULARGE_INTEGER uli;
			uli.HighPart = wfd.nFileSizeHigh;
			uli.LowPart = wfd.nFileSizeLow;
			Result += uli.QuadPart;
		} while(*ppidl == pidl && FindNextFile(hFind, &wfd));
	}
	FindClose(hFind);
	::SysFreeString(bs);
	return Result;
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
	teReleaseClear(ppDragItems);
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
		if (ILIsEqual(pidl, g_pidlResultsFolder)) {
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

LPITEMIDLIST teILCreateFromPath3(IShellFolder *pSF, LPWSTR pszPath, HWND hwnd)
{
	LPITEMIDLIST pidlResult = NULL;
	IEnumIDList *peidl = NULL;
	BSTR bstr = NULL;
	LPWSTR lpDelimiter = StrChr(pszPath, '\\');
	int nDelimiter = 0;
	if (lpDelimiter) {
		nDelimiter = (int)(lpDelimiter - pszPath);
	}
	int ashgdn[] = {SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_INFOLDER, SHGDN_INFOLDER | SHGDN_FORPARSING};
	if SUCCEEDED(pSF->EnumObjects(hwnd, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN | SHCONTF_INCLUDESUPERHIDDEN, &peidl)) {
		LPITEMIDLIST pidlPart;
		while (!pidlResult && peidl->Next(1, &pidlPart, NULL) == S_OK) {
			for (int j = 0; j < 2; j++) {
				if SUCCEEDED(teGetDisplayNameBSTR(pSF, pidlPart, ashgdn[j], &bstr)) {
					if (nDelimiter) {
						if (nDelimiter == ::SysStringLen(bstr) && StrCmpNI(const_cast<LPCWSTR>(pszPath), (LPCWSTR)bstr, nDelimiter) == 0) {
							IShellFolder *pSF1;
							if SUCCEEDED(pSF->BindToObject(pidlPart, NULL, IID_PPV_ARGS(&pSF1))) {
								pidlResult = teILCreateFromPath3(pSF1, &lpDelimiter[1], NULL);
								pSF1->Release();
							}
							::SysFreeString(bstr);
							break;
						}
						::SysFreeString(bstr);
					}
					if (teStrSameIFree(bstr, pszPath)) {
						LPITEMIDLIST pidlParent;
						if (teGetIDListFromObject(pSF, &pidlParent)) {
							pidlResult = ILCombine(pidlParent, pidlPart);
							teCoTaskMemFree(pidlParent);
						}
						break;
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
		pidlResult = teILCreateFromPath3(pSF, pszPath, hwnd);
		pSF->Release();
	}
	return pidlResult;
}

VOID teReleaseILCreate(TEILCreate *pILC)
{
	if (::InterlockedDecrement(&pILC->cRef) == 0) {
		teCoTaskMemFree(pILC->pidlResult);
		delete [] pILC;
	}
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
	LPCITEMIDLIST pidlPart;

	if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
		hr = teGetDisplayNameBSTR(pSF, pidlPart, uFlags & (~SHGDN_FORPARSINGEX), pbs);
		if (hr == S_OK) {
			if (teIsFileSystem(*pbs)) {
				if (uFlags & SHGDN_FORPARSINGEX) {
					BOOL bSearch = FALSE;
					BSTR bs2;
					if SUCCEEDED(teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &bs2)) {
						bSearch = teStartsText(L"search-ms:", bs2);
						::SysFreeString(bs2);
					}
					if (!bSearch) {
						LPITEMIDLIST pidl2 = SHSimpleIDListFromPath(*pbs);
						if (!ILIsEqual(pidl, pidl2)) {
							teILCloneReplace(&pidl2, pidl);
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
									} else {
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
			} else if (((uFlags & (SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)) == (SHGDN_FORADDRESSBAR | SHGDN_FORPARSING))) {
				if (ILGetCount(pidl) == 1 || tePathMatchSpec1(*pbs, L"search-ms:*\\*")) {
					LPITEMIDLIST pidl2 = teILCreateFromPath2(g_pidls[CSIDL_DESKTOP], *pbs, NULL);
					if (!ILIsEqual(pidl, pidl2)) {
						teSysFreeString(pbs);
						hr = teGetDisplayNameBSTR(pSF, pidlPart, uFlags & (~(SHGDN_FORPARSINGEX | SHGDN_FORADDRESSBAR)), pbs);
					}
					teCoTaskMemFree(pidl2);
				}
			}
		}
		pSF->Release();
	}
	return hr;
}


BOOL teILIsFileSystem(LPITEMIDLIST pidl)
{
	BOOL bResult = FALSE;
	BSTR bs;
	if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORPARSING)) {
		bResult = teIsFileSystem(bs);
		::SysFreeString(bs);
	}
	return bResult;
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

void GetVarPathFromIDList(VARIANT *pVarResult, LPITEMIDLIST pidl, int uFlags)
{
	int i;
	BOOL bSpecial = FALSE;

	if (uFlags & SHGDN_FORPARSINGEX) {
		for (i = 0; i < g_nPidls; i++) {
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
	if (pVarResult->vt != VT_I4) {
		if SUCCEEDED(GetDisplayNameFromPidl(&pVarResult->bstrVal, pidl, uFlags)) {
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
	} catch (...) {
		hr = E_FAIL;
	}
	teClearVariantArgs(nArgs, pvArgs);
	return hr;
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
	return hr;
}

HRESULT tePutProperty(IUnknown *punk, LPOLESTR sz, VARIANT *pv)
{
	return tePutProperty0(punk, sz, pv, fdexNameEnsure);
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
	HRESULT hr = DISP_E_MEMBERNOTFOUND;
	IDispatchEx *pdex;
	if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pdex))) {
		BSTR bs = ::SysAllocString(sz);
		hr = pdex->GetDispID(bs, fdexNameCaseInsensitive, &dispid);
		::SysFreeString(bs);
		pdex->Release();
	} else {
		hr = pdisp->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &dispid);
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

HRESULT teGetPropertyAtLLX(IDispatch *pdisp, LONGLONG i, VARIANT *pv)
{
	wchar_t pszName[20];
	swprintf_s(pszName, 20, L"%llx", i);
	return teGetProperty(pdisp, pszName, pv);
}

HRESULT tePutPropertyAtLLX(PVOID pObj, LONGLONG i, VARIANT *pv)
{
	wchar_t pszName[20];
	swprintf_s(pszName, 20, L"%llx", i);
	return tePutProperty((IUnknown *)pObj, pszName, pv);
}

HRESULT teDelPropertyAtLLX(PVOID pObj, LONGLONG i)
{
	wchar_t pszName[20];
	swprintf_s(pszName, 20, L"%llx", i);
	return teDelProperty((IUnknown *)pObj, pszName);
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
		VARIANT vo;
		VariantInit(&vo);
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8)) {
			return vo.llVal;
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

LONGLONG GetLLPFromVariant(VARIANT *pv, VARIANT *pvMem)
{
	if (pv) {
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return GetLLPFromVariant(pv->pvarVal, pvMem);
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
		char *pc = GetpcFromVariant(pv, pvMem);
		if (pc) {
			return (LONGLONG)pc;
		}
		VARIANT vo;
		VariantInit(&vo);
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8)) {
			return vo.llVal;
		}
#ifdef _W2000
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_R8)) {
			return (int)(LONGLONG)vo.dblVal;
		}
#endif
	}
	return 0;
}

VOID teWriteBack(VARIANT *pvArg, VARIANT *pvDat)
{
	IUnknown *punk;
	if (pvDat->vt == VT_BSTR && pvDat->bstrVal) {
		if (FindUnknown(pvArg, &punk)) {
			tePutProperty(punk, L"_BLOB", pvDat);
		}
	}
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
		return;
	}
	if FAILED(VariantChangeType(pvargDest, pvarSrc, 0, vt)) {
		pvargDest->llVal = 0;
	}
}

BOOL teGetVariantTime(VARIANT *pv, DATE *pdt)
{
	VARIANT v;
	VariantInit(&v);
	VariantCopy(&v, pv);
	IDispatch *pdisp;
	if (GetDispatch(&v, &pdisp)) {
		VariantClear(&v);
		teExecMethod(pdisp, L"getVarDate", &v, 0, NULL);
		pdisp->Release();
	}
	if (v.vt != VT_DATE) {
		VariantClear(&v);
		teVariantChangeType(&v, pv, VT_DATE);
	}
	if (v.vt == VT_DATE) {
		*pdt = v.date;
		return TRUE;
	}
	VariantClear(&v);
	return FALSE;
}

BOOL teGetSystemTime(VARIANT *pv, SYSTEMTIME *pSysTime)
{
	DATE dt;
	if (teGetVariantTime(pv, &dt)) {
		return ::VariantTimeToSystemTime(dt, pSysTime);
	}
	return FALSE;
}

HRESULT teInvokeAPI(TEDispatchApi *pApi, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo)
{
	int nArg = pDispParams ? pDispParams->cArgs : 0;
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	if (nArg-- >= pApi->nArgs) {
		VARIANT vParam[12] = { 0 };
		LONGLONG param[12] = { 0 };
		for (int i = nArg < 11 ? nArg : 11; i >= 0; i--) {
			VariantInit(&vParam[i]);
			if (i == pApi->nStr1 || i == pApi->nStr2 || i == pApi->nStr3) {
				teVariantChangeType(&vParam[i], &pDispParams->rgvarg[nArg - i], VT_BSTR);
				param[i] = (LONGLONG)vParam[i].bstrVal;
			} else {
				param[i] = GetLLPFromVariant(&pDispParams->rgvarg[nArg - i], &vParam[i]);
			}
		}
		pApi->fn(nArg, param, pDispParams, pVarResult);
		for (int i = nArg < 11 ? nArg : 11; i >= 0; i--) {
			if (i != pApi->nStr1 && i != pApi->nStr2 && i != pApi->nStr3) {
				teWriteBack(&pDispParams->rgvarg[nArg - i], &vParam[i]);
			}
			VariantClear(&vParam[i]);
		}
	}
	return S_OK;
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
						j++;
					}
				}
#endif
			} else {
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
	HWND hwnd1 = FindWindowEx(hwnd, 0, WC_TREEVIEW, NULL);
	return hwnd1 ? hwnd1 : FindWindowEx(FindWindowEx(hwnd, 0, L"NamespaceTreeControl", NULL), 0, WC_TREEVIEW, NULL);
}

BOOL teSetRect(HWND hwnd, int left, int top, int right, int bottom)
{
	RECT rc;
	GetClientRect(hwnd, &rc);
	MoveWindow(hwnd, left, top, right - left, bottom - top, FALSE);
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
			hr = pSF1->EnumObjects(NULL, SHCONTF_NONFOLDERS | SHCONTF_FOLDERS, &peidl);
			if (peidl) {
				LPITEMIDLIST pidlPart;
				hr = peidl->Next(1, &pidlPart, NULL);
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
	if (!(iUseFS & 2)) {
		WCHAR pszDrive[0x80];
		lstrcpyn(pszDrive, pszPath, 4);
		if (pszDrive[0] >= 'A' && pszDrive[1] == ':' && pszDrive[2] == '\\') {
			if (!GetVolumeInformation(pszDrive, NULL, 0, NULL, NULL, NULL, pszDrive, sizeof(pszDrive))) {
				return E_NOT_READY;
			}
		}
	}
	LPITEMIDLIST pidl = NULL;
	if (iUseFS) {
		lpfnSHParseDisplayName(pszPath, NULL, &pidl, 0, NULL);
	} else {
		pidl = teILCreateFromPath(pszPath);
	}
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
	try {
		TEExists *pExists = (TEExists *)args;
		pExists->hr = tePathIsDirectory2(pExists->pszPath, pExists->iUseFS);
		SetEvent(pExists->hEvent);
		teReleaseExists(pExists);
	} catch (...) {
		g_nException = 0;
	}
	::_endthread();
}

HRESULT tePathIsDirectory(LPWSTR pszPath, int dwms, int iUseFS)
{
	if (g_dwMainThreadId == GetCurrentThreadId()) {
		if (dwms <= 0) {
			dwms = g_param[TE_NetworkTimeout];
		}
		if (dwms) {
			TEExists *pExists = new TEExists[1];
			pExists->cRef = 2;
			pExists->hr = E_ABORT;
			pExists->pszPath = pszPath;
			pExists->hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
			pExists->iUseFS = iUseFS;
			_beginthread(&threadExists, 0, pExists);
			WaitForSingleObject(pExists->hEvent, dwms);
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

	if (pszPath) {
		BSTR bsPath2 = NULL;
		if (pszPath[0] == _T('"')) {
			bsPath2 = ::SysAllocStringLen(pszPath, lstrlen(pszPath) + 1);
			PathUnquoteSpaces(bsPath2);
			pszPath = bsPath2;
		}
		BSTR bsPath3 = NULL;
		if (tePathMatchSpec(pszPath, L"search-ms:*&crumb=location:*")) {
			LPWSTR lp1, lp2;
			lp1 = StrChr(pszPath, ':');
			while (lp2 = StrChr(lp1 + 1, ':')) {
				lp1 = lp2;
			}
			lp1 -= 4;
			BSTR bs = ::SysAllocString(lp1);
			bs[0] = 'f';
			bs[1] = 'i';
			bs[2] = 'l';
			bs[3] = 'e';
			DWORD dwLen = ::SysStringLen(bs);
			bsPath3 = ::SysAllocStringLen(NULL, dwLen);
			if SUCCEEDED(PathCreateFromUrl(bs, bsPath3, &dwLen, 0)) {
				pszPath = bsPath3;
			}
		} else if (tePathMatchSpec(pszPath, L"*\\..\\*;*\\..;*\\.\\*;*\\.;*%*%*")) {
			DWORD dwLen = lstrlen(pszPath) + MAX_PATH;
			bsPath3 = ::SysAllocStringLen(NULL, dwLen);
			if (PathSearchAndQualify(pszPath, bsPath3, dwLen)) {
				pszPath = bsPath3;
			}
		} else if (lstrlen(pszPath) == 1 && pszPath[0] >= 'A') {
			bsPath3 = ::SysAllocStringLen(L"?:\\", 4);
			bsPath3[0] = pszPath[0];
			pszPath = bsPath3;
		}
		if (GetCSIDLFromPath(&n, pszPath)) {
			pidl = ::ILClone(g_pidls[n]);
			pszPath = NULL;
		}
		if (pszPath) {
			if (tePathMatchSpec1(pszPath, L"\\\\*\\*")) {
				LPWSTR lpDelimiter = StrChr(&pszPath[2], '\\');
				BSTR bsServer = ::SysAllocStringLen(pszPath, int(lpDelimiter - pszPath));
				LPITEMIDLIST pidlServer;
				if SUCCEEDED(lpfnSHParseDisplayName(bsServer, NULL, &pidlServer, 0, NULL)) {
					pidl = teILCreateFromPath2(pidlServer, &lpDelimiter[1], g_hwndMain);
					teCoTaskMemFree(pidlServer);
				}
				::SysFreeString(bsServer);
			}
			if (!pidl) {
				lpfnSHParseDisplayName(pszPath, NULL, &pidl, 0, NULL);
				if (pidl) {
					if (tePathIsNetworkPath(pszPath) && PathIsRoot(pszPath) && FAILED(tePathIsDirectory(pszPath, 0, 3))) {
						teILFreeClear(&pidl);
					}
				} else if (PathGetDriveNumber(pszPath) >= 0 && !PathIsRoot(pszPath)) {
					WCHAR pszDrive[4];
					lstrcpyn(pszDrive, pszPath, 4);
					int n = GetDriveType(pszDrive);
					if (n == DRIVE_NO_ROOT_DIR && SUCCEEDED(tePathIsDirectory(pszDrive, 0, 3))) {
						lpfnSHParseDisplayName(pszPath, NULL, &pidl, 0, NULL);
					}
				}
			}
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
						if (!lstrcmpi(bstr, g_bsPidls[i])) {
							pidl = ILClone(g_pidls[i]);
							break;
						}
					}
				}
				if (!pidl) {
					pidl = teILCreateFromPath2(g_pidls[CSIDL_DESKTOP], pszPath, NULL);
					if (!pidl) {
						pidl = teILCreateFromPath2(g_pidls[CSIDL_DRIVES], pszPath, NULL);
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
	try {
		TEILCreate *pILC = (TEILCreate *)args;
		pILC->pidlResult = teILCreateFromPath1(pILC->pszPath);
		SetEvent(pILC->hEvent);
		teReleaseILCreate(pILC);
	} catch (...) {
		g_nException = 0;
	}
	::_endthread();
}

LPITEMIDLIST teILCreateFromPath0(LPWSTR pszPath, BOOL bForceLimit)
{
	DWORD dwms = g_pTE ? g_param[TE_NetworkTimeout] : 2000;
	if (bForceLimit && (!dwms || dwms > 500)) {
		dwms = 500;
	}
	if (dwms && (bForceLimit || g_dwMainThreadId == GetCurrentThreadId())) {
		TEILCreate *pILC = new TEILCreate[1];
		pILC->cRef = 2;
		pILC->pidlResult = NULL;
		pILC->pszPath = pszPath;
		pILC->hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
		_beginthread(&threadILCreate, 0, pILC);
		LPITEMIDLIST pidl = NULL;
		if (WaitForSingleObject(pILC->hEvent, dwms) != WAIT_TIMEOUT) {
			pidl = pILC->pidlResult;
			if (pidl) {
				pILC->pidlResult = NULL;
			} else {
				pidl = teILCreateFromPath1(pszPath);
			}
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
							} else if (pszTo && *pszTo) {
								if (pFOS->wFunc == FO_RENAME) {
									hr = pFO->RenameItem(psiFrom, pszTo, NULL);
								} else {
									if (psiTo || GetDestAndName(pszTo, &psiTo, &pszName)) {
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
								}
								if (pFOS->fFlags & FOF_MULTIDESTFILES) {
									teReleaseClear(&psiTo);
									while (*(pszTo++));
								}
							}
							psiFrom->Release();
						}
						while (*(pszFrom++));
					}
					teReleaseClear(&psiTo);
					if SUCCEEDED(hr) {
						hr = pFO->PerformOperations();
						pFO->GetAnyOperationsAborted(&pFOS->fAnyOperationsAborted);
					}
				} catch (...) {}
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
	}
	::InterlockedDecrement(&g_nThreads);
	::SysFreeString(const_cast<BSTR>(pFO->pTo));
	::SysFreeString(const_cast<BSTR>(pFO->pFrom));
	delete [] pFO;
	::CoUninitialize();
	::_endthread();
}

static void threadFolderSize(void *args)
{
	::CoInitialize(NULL);
	TEFS *pFS = (TEFS *)args;
	IDispatch *pTotalFileSize;
	CoGetInterfaceAndReleaseStream(pFS->pStrmTotalFileSize, IID_PPV_ARGS(&pTotalFileSize));
	VARIANT v;
	VariantInit(&v);
	try {
		teSetLL(&v, teGetFolderSize(pFS->bsPath, 128, pFS->ppidl, pFS->pidl));
		if (*pFS->ppidl == pFS->pidl) {
			tePutProperty(pTotalFileSize, pFS->bsName, &v);
			if (pFS->hwnd) {
				InvalidateRect(pFS->hwnd, NULL, FALSE);
			}
		}
	} catch (...) {
		g_nException = 0;
	}
	VariantClear(&v);
	pTotalFileSize->Release();
	::SysFreeString(pFS->bsPath);
	::SysFreeString(pFS->bsName);
	delete [] pFS;
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

HRESULT teSHGetDataFromIDList(IShellFolder *pSF, LPCITEMIDLIST pidlPart, int nFormat, void *pv, int cb)
{
	HRESULT hr = SHGetDataFromIDList(pSF, pidlPart, nFormat, pv, cb);
	if (hr != S_OK && nFormat == SHGDFIL_FINDDATA) {
		// for Library
		BSTR bs;
		if SUCCEEDED(teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_FORPARSING, &bs)) {
			HANDLE hFind = FindFirstFile(bs, (LPWIN32_FIND_DATA)pv);
			::SysFreeString(bs);
			if (hFind != INVALID_HANDLE_VALUE) {
				FindClose(hFind);
				hr = S_OK;
			}
		}
	}
	return hr;
}

CteShellBrowser* SBfromhwnd(HWND hwnd)
{
	CteShellBrowser *pSB;
	for (int i = MAX_FV; i-- && (pSB = g_ppSB[i]);) {
		if (pSB->m_hwnd == hwnd || IsChild(pSB->m_hwnd, hwnd)) {
			return pSB;
		}
	}
	return NULL;
}

CteTabCtrl* TCfromhwnd(HWND hwnd)
{
	CteTabCtrl *pTC;
	for (int i = MAX_TC; i-- && (pTC = g_ppTC[i]);) {
		if (pTC->m_hwnd == hwnd || pTC->m_hwndStatic == hwnd || pTC->m_hwndButton == hwnd) {
			return pTC;
		}
	}
	return NULL;
}

CteTreeView* TVfromhwnd(HWND hwnd, BOOL bAll)
{
	CteShellBrowser *pSB;
	for (int i = MAX_FV; i-- && (pSB = g_ppSB[i]);) {
		HWND hwndTV = pSB->m_pTV->m_hwnd;
		if (hwndTV == hwnd || IsChild(hwndTV, hwnd)) {
			if (bAll || pSB->m_pTV->m_bMain) {
				return pSB->m_pTV;
			}
			return NULL;
		}
	}
	return NULL;
}

void CheckChangeTabSB(HWND hwnd)
{
	if (g_pTC) {
		CteShellBrowser *pSB = SBfromhwnd(hwnd);
		if (!pSB) {
			CteTreeView *pTV = TVfromhwnd(hwnd, false);
			if (pTV) {
				pSB = pTV->m_pFV;
			}
		}
		if (pSB && pSB->m_pTC->m_bVisible) {
			if (g_pTC->m_hwnd != pSB->m_pTC->m_hwnd) {
				g_pTC = pSB->m_pTC;
				pSB->m_pTC->TabChanged(false);
			}
		}
	}
}

void CheckChangeTabTC(HWND hwnd, BOOL bFocusSB)
{
	CteTabCtrl *pTC = TCfromhwnd(hwnd);
	if (pTC && pTC->m_bVisible) {
		if (g_pTC->m_hwnd != pTC->m_hwnd) {
			g_pTC = pTC;
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

#ifdef _USE_APIHOOK
LSTATUS APIENTRY teRegQueryValueExW(HKEY hKey, LPCWSTR lpValueName, LPDWORD lpReserved, LPDWORD lpType, LPBYTE lpData, LPDWORD lpcbData)
{
	return lpfnRegQueryValueExW(hKey, lpValueName, lpReserved, lpType, lpData, lpcbData);
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
	if SUCCEEDED(CoCreateInstance(CLSID_PropertiesUI, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&pPUI))) {
		hr = pPUI->ParsePropertyName(pszString, &pkey->fmtid, &pkey->pid, &chEaten);
		pPUI->Release();
	}
	return hr;
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
			lp1++;
		}
	}
}

#endif
#ifdef _W2000
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

#endif

HMODULE teCreateInstance(LPVARIANTARG pvDllFile, LPVARIANTARG pvClass, REFIID riid, PVOID *ppvObj)
{
	CLSID clsid;
	VARIANT v;
	teVariantChangeType(&v, pvClass, VT_BSTR);
	teCLSIDFromString(v.bstrVal, &clsid);
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

BOOL teGetIDListFromObject(IUnknown *punk, LPITEMIDLIST *ppidl)
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
				if FAILED(lpfnSHGetIDListFromObject(pSV, ppidl)) {
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
	}
	return *ppidl != NULL;
}

BOOL teGetIDListFromObjectEx(IUnknown *punk, LPITEMIDLIST *ppidl)
{
	if (!teGetIDListFromObject(punk, ppidl)) {
		*ppidl = ::ILClone(g_pidls[CSIDL_DESKTOP]);
	}
	return true;
}


BOOL teILIsEqualNet(BSTR bs1, IUnknown *punk, BOOL *pbResult)
{
	if (tePathIsNetworkPath(bs1)) {
		LPITEMIDLIST pidl;
		if (teGetIDListFromObject(punk, &pidl)) {
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
	teGetStrFromFolderItem(&bs2, punk2);
	if (teGetStrFromFolderItem(&bs1, punk1)) {
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
		if (teGetIDListFromObject(punk1, &pidl1)) {
			if (teGetIDListFromObject(punk2, &pidl2)) {
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
		if (pv->vt == VT_R8) {
			return (int)(LONGLONG)pv->dblVal;
		}
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
	return 0;
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
				*ppidl = (LPITEMIDLIST)CoTaskMemAlloc(nSize);
				::CopyMemory(*ppidl, pvData, nSize);
				::SafeArrayUnaccessData(psa);
			}
			break;
		case VT_NULL:
			break;
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

BOOL GetVarArrayFromIDList(VARIANT *vaPidl, LPITEMIDLIST pidl)
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

HRESULT GetFolderObjFromIDList(LPITEMIDLIST pidl, Folder** ppsdf)
{
	VARIANT v;
	GetVarArrayFromIDList(&v, pidl);
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
	teSetLong(&v, nIndex);
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


HRESULT GetFolderItemFromShellItem(FolderItem **ppid, IShellItem *psi)
{
	if (psi) {
		LPITEMIDLIST pidl;
		if (teGetIDListFromObject(psi, &pidl)) {
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

int GetIntFromVariantClear(VARIANT *pv)
{
	int i = GetIntFromVariant(pv);
	VariantClear(pv);
	return i;
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
			if (lCount) {
				CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL, true);
				VARIANT v, v2;
				VariantInit(&v);
				VariantInit(&v2);
				v.vt = VT_I4;
				for (v.lVal = 0; v.lVal < lCount; v.lVal++) {
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
			VARIANT v;
			VariantInit(&v);
			teGetProperty(pdex, L"length", &v);
			long lCount = GetIntFromVariantClear(&v);
			if (lCount) {
				CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL, true);
				for (int i = 0; i < lCount; i++) {
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
	}
	LPITEMIDLIST pidl;
	pidl = NULL;
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

#ifdef _USE_BSEARCHAPI
int* teSortDispatchApi(TEDispatchApi *method, int nCount)
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
			} else {
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

int teBSearchApi(TEDispatchApi *method, int nSize, int* pMap, LPOLESTR bs)
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
#endif

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
			} else {
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

	for (int j = 0; j < nCount; j++) {
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
		for (int i = j; i > nMin; i--) {
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
	if (pMap) {
		int nIndex = teBSearch(method, nCount, pMap, bs);
		if (nIndex >= 0) {
			*rgDispId = method[nIndex].id;
			return S_OK;
		}
	} else {
		for (int i = 0; method[i].name; i++) {
			if (lstrcmpi(bs, method[i].name) == 0) {
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
			if (tePathMatchSpec1(pv->bstrVal, L"*%")) {
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
	char *pc = GetpcFromVariant(pv, NULL);
	if (pc) {
		IDispatch *pdisp;
		if (GetDispatch(pv, &pdisp)) {
			VARIANT v;
			VariantInit(&v);
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

static void threadAddItems(void *args)
{
	::OleInitialize(NULL);
	LPOLESTR lpShift = L"shift";
	DISPID dispidShift;
	LPITEMIDLIST pidl;

	TEAddItems *pAddItems = (TEAddItems *)args;
	IDispatch *pSB;
	CoGetInterfaceAndReleaseStream(pAddItems->pStrmSB, IID_PPV_ARGS(&pSB));
	IDispatch *pArray;
	CoGetInterfaceAndReleaseStream(pAddItems->pStrmArray, IID_PPV_ARGS(&pArray));
	try {
		HRESULT hr = pArray->GetIDsOfNames(IID_NULL, &lpShift, 1, LOCALE_USER_DEFAULT, &dispidShift);
		if (hr == S_OK) {
			while (Invoke5(pArray, dispidShift, DISPATCH_METHOD, &pAddItems->pv[1], 0, NULL) == S_OK && pAddItems->pv[1].vt != VT_EMPTY) {
				if (teGetIDListFromVariant(&pidl, &pAddItems->pv[1], TRUE)) {
					VariantClear(&pAddItems->pv[1]);
					teSetIDList(&pAddItems->pv[1], pidl);
					teCoTaskMemFree(pidl);
					if (Invoke5(pSB, 0x10000501, DISPATCH_METHOD, NULL, -2, pAddItems->pv) != S_OK) {
						break;
					}
				}
				VariantClear(&pAddItems->pv[1]);
			}
		}
	} catch (...) {}
	teClearVariantArgs(2, pAddItems->pv);
	delete [] pAddItems;
	pArray->Release();
	pSB->Release();
	::OleUninitialize();
	::_endthread();
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
	if (tePathMatchSpec1(lpLang, L"J*Script")) {
		CoCreateInstance(CLSID_JScriptChakra, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pas));
	}
	if (pas == NULL && teCLSIDFromProgID(lpLang, &clsid) == S_OK) {
		CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pas));
	}
	if (pas) {
		IActiveScriptProperty *paspr;
		if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&paspr))) {
			VARIANT v;
			teSetLong(&v, 0);
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
					while (hr == S_OK) {
						BSTR bs;
						if (pdex->GetMemberName(dispid, &bs) == S_OK) {
							pas->AddNamedItem(bs, SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE | SCRIPTITEM_GLOBALMEMBERS);
							::SysFreeString(bs);
						}
						hr = pdex->GetNextDispID(fdexEnumAll, dispid, &dispid);
					}
					pdex->Release();
				}
			} else if (pv && pv->boolVal) {
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
	int npt = GetIntFromVariant(pv);
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
	if (teGetIDListFromObject(pid, &pidl)) {
		bResult = ::ILRemoveLastID(pidl);
#ifndef _WIN64
		Wow64ControlPanel(&pidl, g_pidls[CSIDL_DESKTOP]);
#endif
		if (!bResult || ILIsEmpty(pidl)) {
			BSTR bs;
			if SUCCEEDED(pid->get_Path(&bs)) {
				if (teIsFileSystem(bs) && bs[1] == ':' && !teIsDesktopPath(bs)) {
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

VOID GetNewObject(IDispatch **ppObj)
{
	VARIANT v;
	VariantInit(&v);
	teReleaseClear(ppObj);
	if (teExecMethod(g_pJS, L"_o", &v, 0, NULL) == S_OK) {
		GetDispatch(&v, ppObj);
		VariantClear(&v);
	}
}

VOID ClearEvents()
{
	GetNewObject(&g_pSubWindows);
	for (int j = Count_OnFunc; j-- > 1;) {
		teReleaseClear(&g_pOnFunc[j]);
	}

	CteShellBrowser *pSB;
	for (int i = MAX_FV; i-- && (pSB = g_ppSB[i]);) {
		for (int j = Count_SBFunc; j-- > 1;) {
			teReleaseClear(&pSB->m_pDispatch[j]);
		}
	}
	g_param[TE_Tab] = TRUE;
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
	if (g_pOnFunc[TE_OnStatusText]) {
		VARIANTARG *pv;
		VARIANT vResult;
		VariantInit(&vResult);
		pv = GetNewVARIANT(3);
		teSetObject(&pv[2], pObj);
		teSetSZ(&pv[1], lpszStatusText);
		teSetLong(&pv[0], iPart);
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
		CteMemory *pstPt = new CteMemory(2 * sizeof(int), NULL, 1, L"POINT");
		pstPt->SetPoint(pt.x, pt.y);
		teSetObject(&pv[1], pstPt);
		pstPt->Release();
		teSetLong(&pv[0], flags);
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
				teSetObjectRelease(&pv[2], new CteMemory(sizeof(int), pgrfKeyState, 1, L"DWORD"));
				pstPt = new CteMemory(2 * sizeof(int), NULL, 1, L"POINT");
				pstPt->SetPoint(pt.x, pt.y);
				teSetObjectRelease(&pv[1], pstPt);
				teSetObjectRelease(&pv[0], new CteMemory(sizeof(int), pdwEffect, 1, L"DWORD"));
				if SUCCEEDED(Invoke4(g_pOnFunc[nFunc], &vResult, 5, pv)) {
					hr = GetIntFromVariant(&vResult);
				}
			}
		} catch(...) {}
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
		teSetPtr(&pv[3], pMsg->hwnd);
		teSetLong(&pv[2], pMsg->message);
		teSetPtr(&pv[1], pMsg->wParam);
		teSetPtr(&pv[0], pMsg->lParam);
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
		teSetPtr(&pv[3], pMsg->hwnd);
		teSetLong(&pv[2], pMsg->message);
		teSetLong(&pv[1], (LONG)pMsg->wParam);
		CteMemory *pstPt = new CteMemory(2 * sizeof(int), NULL, 1, L"POINT");
		pstPt->SetPoint(pMsg->pt.x, pMsg->pt.y);
		teSetObjectRelease(&pv[0], pstPt);

		if SUCCEEDED(Invoke4(g_pOnFunc[nFunc], &vResult, 5, pv)) {
			return GetIntFromVariantClear(&vResult);
		}
	}
	return S_FALSE;
}

int GetShellBrowser2(CteShellBrowser *pSB1)
{
	CteShellBrowser *pSB;
	for (int i = MAX_FV; i-- && (pSB = g_ppSB[i]);) {
		if (pSB == pSB1 && !pSB->m_bEmpty) {
			return i;
		}
	}
	return 0;
}

CteShellBrowser* GetNewShellBrowser(CteTabCtrl *pTC)
{
	CteShellBrowser *pSB = NULL;
	int i = MAX_FV;
	while (i--) {
		pSB = g_ppSB[i];
		if (pSB) {
			if (pSB->m_bEmpty) {
				pSB->Init(pTC, FALSE);
				break;
			}
		} else {
			pSB = new CteShellBrowser(pTC);
			pSB->m_nSB = i;
			g_ppSB[i] = pSB;
			break;
		}
	}
	if (pTC && i >= 0) {
		TC_ITEM tcItem;
		tcItem.mask = TCIF_PARAM;
		tcItem.lParam = i;
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
	int i = MAX_TC;
	while (i--) {
		if (g_ppTC[i]) {
			if (g_ppTC[i]->m_bEmpty) {
				return g_ppTC[i];
			}
		} else {
			g_ppTC[i] = new CteTabCtrl();
			g_ppTC[i]->m_nTC = i;
			return g_ppTC[i];
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
	CteTabCtrl *pTC = TCfromhwnd(hwnd);
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
			} catch(...) {
				g_nException = 0;
			}
			::InterlockedDecrement(&g_nProcKey);
		}
	}
	return lResult ? CallNextHookEx(g_hMenuKeyHook, nCode, wParam, lParam) : lResult;
}

LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	static UINT dwDoubleTime;
	static WPARAM wParam2;

	HRESULT hrResult = S_FALSE;
	try {
		IDispatch *pdisp = NULL;
		if (nCode >= 0 && g_x == MAXINT) {
			if (nCode == HC_ACTION) {
				if (wParam != WM_MOUSEWHEEL) {
					if (g_pOnFunc[TE_OnMouseMessage]) {
						MOUSEHOOKSTRUCTEX *pMHS = (MOUSEHOOKSTRUCTEX *)lParam;
						if (ControlFromhwnd(&pdisp, pMHS->hwnd) == S_OK) {
							try {
								if (InterlockedIncrement(&g_nProcMouse) == 1 || wParam != wParam2) {
									wParam2 = wParam;
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
											} else {
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
												} else if (wParam == WM_MBUTTONUP) {
													nStyle = NSTCECT_MBUTTON;
												}
												hrResult = pTV->OnItemClick(NULL, ht.flags, nStyle);
											}
										}
									}
#endif
								}
							} catch(...) {}
							::InterlockedDecrement(&g_nProcMouse);
							pdisp->Release();
						}
					}
				}
			}
		}
	} catch (...) {
		g_nException = 0;
	}
	return (hrResult != S_OK) ? CallNextHookEx(g_hMouseHook, nCode, wParam, lParam) : TRUE;
}

#ifdef _2000XP
LRESULT CALLBACK TETVProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	BOOL bHandled = FALSE;
	try {
		CteTreeView *pTV = (CteTreeView *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (hwnd != GetParent(pTV->m_hwndTV)) {
			return 0;
		}
		if (msg == WM_NOTIFY) {
			LPNMTVCUSTOMDRAW lptvcd = (LPNMTVCUSTOMDRAW)lParam;
/// Color
			if (lptvcd->nmcd.hdr.code == NM_CUSTOMDRAW) {
				if (lptvcd->nmcd.dwDrawStage == CDDS_PREPAINT) {
					return CDRF_NOTIFYITEMDRAW;
				}
				if (lptvcd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
					LONG_PTR lRes = CDRF_DODEFAULT;
					if (g_pOnFunc[TE_OnItemPrePaint] && !(lptvcd->nmcd.uItemState & CDIS_SELECTED)) {
						VARIANTARG *pv = GetNewVARIANT(5);
						teSetObject(&pv[4], pTV);
/*//
						if (pTV->m_pShellNameSpace && (lptvcd->nmcd.uItemState & CDIS_SELECTED)) {
							IDispatch *pid;
							if (pTV->m_pShellNameSpace->get_SelectedItem(&pid) == S_OK) {
								teSetObjectRelease(&pv[3], pid);
							}
						}
*///
						teSetObjectRelease(&pv[2], new CteMemory(sizeof(NMCUSTOMDRAW), &lptvcd->nmcd, 1, L"NMCUSTOMDRAW"));
						teSetObjectRelease(&pv[1], new CteMemory(sizeof(NMTVCUSTOMDRAW), lptvcd, 1, L"NMTVCUSTOMDRAW"));
						teSetObjectRelease(&pv[0], new CteMemory(sizeof(HANDLE), &lRes, 1, L"HANDLE"));
						Invoke4(g_pOnFunc[TE_OnItemPrePaint], NULL, 5, pv);
					}
					return lRes;
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
				} catch(...) {}
				::InterlockedDecrement(&g_nProcTV);
			}
		}
		return bHandled ? 1 : CallWindowProc(pTV->m_DefProc, hwnd, msg, wParam, lParam);
	} catch (...) {
		g_nException = 0;
	}
	return 0;
}
#endif

LRESULT CALLBACK TETVProc2(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	try {
		CteTreeView *pTV = (CteTreeView *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (hwnd != pTV->m_hwndTV) {
			return 0;
		}
		if (pTV->m_bMain) {
			if (msg == WM_ENTERMENULOOP) {
				pTV->m_bMain = false;
			}
		} else {
			if (msg == WM_EXITMENULOOP) {
				pTV->m_bMain = true;
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
			} catch(...) {}
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
	} catch (...) {
		g_nException = 0;
	}
	return 0;
}

LRESULT CALLBACK TELVProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hwndEdit = NULL;
	static FolderItem *pidEdit = NULL;

	CteShellBrowser *pSB = (CteShellBrowser *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
	try {
		if (hwnd != pSB->m_hwndDV) {
			return 0;
		}
		LRESULT lResult = S_FALSE;
		if (msg == WM_NOTIFY) {
			if (((LPNMHDR)lParam)->code == LVN_GETDISPINFO) {
				NMLVDISPINFO *lpDispInfo = (NMLVDISPINFO *)lParam;
				if (lpDispInfo->item.mask & LVIF_TEXT) {
					if (pSB->m_param[SB_ViewMode] == FVM_DETAILS) {
						if (lpDispInfo->item.iSubItem == pSB->m_nFolderSizeIndex) {
							lResult = CallWindowProc((WNDPROC)pSB->m_DefProc, hwnd, msg, wParam, lParam);
							if (lpDispInfo->item.pszText[0] == 0) {
								IFolderView *pFV;
								LPITEMIDLIST pidl;
								if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
									if SUCCEEDED(pFV->Item(lpDispInfo->item.iItem, &pidl)) {
										pSB->SetFolderSize(pidl, lpDispInfo->item.pszText, lpDispInfo->item.cchTextMax);
										teCoTaskMemFree(pidl);
									}
									pFV->Release();
								}
							}
							return lResult;
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
						if (lpDispInfo->item.iSubItem == pSB->m_nSizeIndex && pSB->m_param[SB_SizeFormat]) {
							lResult = CallWindowProc((WNDPROC)pSB->m_DefProc, hwnd, msg, wParam, lParam);
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
							return lResult;
						}
					}
					if (g_dwHideFileExt && lpDispInfo->item.iSubItem == 0) {
						lResult = CallWindowProc((WNDPROC)pSB->m_DefProc, hwnd, msg, wParam, lParam);
						if (!StrChr(lpDispInfo->item.pszText, '.')) {
							IFolderView *pFV;
							LPITEMIDLIST pidl;
							if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
								if SUCCEEDED(pFV->Item(lpDispInfo->item.iItem, &pidl)) {
									BSTR bsFile;
									if SUCCEEDED(teGetDisplayNameBSTR(pSB->m_pSF2, pidl, SHGDN_FORPARSING, &bsFile)) {
										LPWSTR pszExt = PathFindExtension(bsFile);
										if (pszExt && lstrlen(lpDispInfo->item.pszText) + lstrlen(pszExt) < lpDispInfo->item.cchTextMax) {
											lstrcat(lpDispInfo->item.pszText, pszExt);
										}
										::SysFreeString(bsFile);
									}
									teCoTaskMemFree(pidl);
									pFV->Release();
								}
							}
						}
						return lResult;
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
							if (!GetIntFromVariantClear(&vResult)) {
								return 1;
							}
						}
					}
				}
			} else if (((LPNMHDR)lParam)->code == LVN_INCREMENTALSEARCH) {
				NMLVFINDITEM *lpFindItem = (NMLVFINDITEM *)lParam;
				if (lpFindItem->lvfi.flags & (LVFI_STRING | LVFI_PARTIAL)) {
					DoStatusText(pSB, lpFindItem->lvfi.psz, 0);
				}
			} else if (((LPNMHDR)lParam)->code == HDN_ITEMCHANGED) {
				pSB->SetLVSettings();
			}
/// Color
			if (g_pOnFunc[TE_OnItemPrePaint] && pSB->m_pShellView) {
				LPNMLVCUSTOMDRAW lplvcd = (LPNMLVCUSTOMDRAW)lParam;
				if (lplvcd->nmcd.hdr.code == NM_CUSTOMDRAW) {
					if (lplvcd->nmcd.dwDrawStage == CDDS_PREPAINT) {
						return CDRF_NOTIFYITEMDRAW;
					}
					if (lplvcd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
						LRESULT lRes = CDRF_DODEFAULT;
						IFolderView *pFV;
						if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
							LPITEMIDLIST pidl;
							if SUCCEEDED(pFV->Item((int)lplvcd->nmcd.dwItemSpec , &pidl)) {
								LPITEMIDLIST pidlFull = ILCombine(pSB->m_pidl, pidl);
								VARIANTARG *pv = GetNewVARIANT(5);
								teSetObject(&pv[4], pSB);
								teSetIDList(&pv[3], pidlFull);
								CoTaskMemFree(pidlFull);
								CoTaskMemFree(pidl);
								teSetObjectRelease(&pv[2], new CteMemory(sizeof(NMCUSTOMDRAW), &lplvcd->nmcd, 1, L"NMCUSTOMDRAW"));
								teSetObjectRelease(&pv[1], new CteMemory(sizeof(NMLVCUSTOMDRAW), lplvcd, 1, L"NMLVCUSTOMDRAW"));
								teSetObjectRelease(&pv[0], new CteMemory(sizeof(HANDLE), &lRes, 1, L"HANDLE"));
								Invoke4(g_pOnFunc[TE_OnItemPrePaint], NULL, 5, pv);
							}
							pFV->Release();
						}
						return lRes;
					}
				}
			}
//*/
		}
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
							if (wParam == 0x7103) {//Refresh
								if (ILIsEqual(pSB->m_pidl, g_pidlResultsFolder)) {
//									pSB->BrowseObject(NULL, SBSP_RELATIVE | SBSP_SAMEBROWSER);
									pSB->Refresh(FALSE);
								} else {
									pSB->m_bRefreshing = TRUE;
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
			} catch(...) {}
			::InterlockedDecrement(&g_nProcFV);
		}
		if (msg == WM_WINDOWPOSCHANGED) {
			pSB->SetFolderFlags();
		}
		if (hwnd != pSB->m_hwndDV) {
			return 0;
		}
		return lResult ? CallWindowProc((WNDPROC)pSB->m_DefProc, hwnd, msg, wParam, lParam) : 0;
	} catch (...) {
		g_nException = 0;
	}
	return 0;
}

/*/// Shift to HookProc
LRESULT CALLBACK TELVProc2(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	try {
		CteShellBrowser *pSB = (CteShellBrowser *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (hwnd != pSB->m_hwndLV) {
			return 0;
		}
		if (msg == WM_HSCROLL || msg == WM_VSCROLL) {
			g_dwTickScroll = GetTickCount();
		}
		return CallWindowProc((WNDPROC)pSB->m_DefProc2, hwnd, msg, wParam, lParam);
	} catch (...) {
		g_nException = 0;
	}
	return 0;
}
//*///
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

LRESULT CALLBACK TESTProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = 1;
	try {
		CteTabCtrl *pTC = (CteTabCtrl *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (hwnd != pTC->m_hwndStatic) {
			return 0;
		}
		CteShellBrowser	*pSB = pTC->GetShellBrowser(pTC->m_nIndex);
		RECT rc;
		if (pSB && pSB->m_pidl && GetWindowRect(pSB->m_hwnd, &rc) && rc.top != rc.bottom && IsWindowVisible(pSB->m_pTV->m_hwndTV)) {
			switch (msg) {
				case WM_MOUSEMOVE:
					if (g_x == MAXINT) {
						SetCursor(LoadCursor(NULL, IDC_SIZEWE));
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
		return lResult ? CallWindowProc(pTC->m_DefSTProc, hwnd, msg, wParam, lParam) : 0;
	} catch (...) {
		g_nException = 0;
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
			CheckChangeTabTC(hwnd, true);
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

LRESULT CALLBACK TEBTProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = 1;
	try {
		CteTabCtrl *pTC = (CteTabCtrl *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (hwnd != pTC->m_hwndButton) {
			return 0;
		}
		switch (msg) {
			case WM_NOTIFY:
				switch (((LPNMHDR)lParam)->code) {
					case TCN_SELCHANGE:
						if (g_pTC->m_hwnd != pTC->m_hwnd && pTC->m_bVisible) {
							g_pTC = pTC;
						}
						pTC->TabChanged(true);
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
							VariantInit(&pv[1]);
							teSetObject(&pv[1], pTC);
							VariantInit(&pv[0]);
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
		if (lResult) {
			lResult = TETCProc2(pTC, hwnd, msg, wParam, lParam);
		}
		return lResult ? CallWindowProc(pTC->m_DefBTProc, hwnd, msg, wParam, lParam) : 0;
	} catch (...) {
		g_nException = 0;
	}
	return 0;
}

LRESULT CALLBACK TETCProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT Result = 1;
	int nIndex, nCount;
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
					CallWindowProc(pTC->m_DefTCProc, hwnd, msg, wParam, lParam);
					Result = 0;
					if (pSB) {
						pSB->Close(true);
					}
					nIndex = TabCtrl_GetCurSel(hwnd);
					if (nIndex >= 0) {
						pTC->m_nIndex = nIndex;
					} else {
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
					if (g_param[TE_Tab] && pTC->m_param[TC_Align] == 1) {
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
	} catch (...) {
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
		} catch(...) {}
		::InterlockedDecrement(&g_nProcMouse);
	}

	if (pMsg->message >= WM_KEYFIRST && pMsg->message <= WM_KEYLAST) {
		try {
			if (InterlockedIncrement(&g_nProcKey) == 1) {
				if (ControlFromhwnd(&pdisp, pMsg->hwnd) == S_OK) {
					if (MessageSub(TE_OnKeyMessage, pdisp, pMsg, S_FALSE) == S_OK) {
						hrResult = S_OK;
					} else if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pSB))) {
						if SUCCEEDED(pSB->QueryActiveShellView(&pSV)) {
							if (pSV->TranslateAcceleratorW(pMsg) == S_OK) {
								hrResult = S_OK;
							}
							pSV->Release();
						}
						pSB->Release();
					} else if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pWB))) {
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
				VARIANT v;
				if SUCCEEDED(teGetPropertyAtLLX(g_pSubWindows, (LONGLONG)GetParent((HWND)GetWindowLongPtr(GetParent(pMsg->hwnd), GWLP_HWNDPARENT)), &v)) {
					if SUCCEEDED(v.punkVal->QueryInterface(IID_PPV_ARGS(&pWB))) {
						if SUCCEEDED(pWB->QueryInterface(IID_PPV_ARGS(&pActiveObject))) {
							if (pActiveObject->TranslateAcceleratorW(pMsg) == S_OK) {
								hrResult = S_OK;
							}
							pActiveObject->Release();
						}
						pWB->Release();
					}
					VariantClear(&v);
				}
			}
		} catch(...) {}
		::InterlockedDecrement(&g_nProcKey);
	}
	return hrResult;
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
			teSysFreeString(&g_bsPidls[g_nPidls]);
		}
		teCoTaskMemFree(g_pidlResultsFolder);
		teCoTaskMemFree(g_pidlCP);
		teCoTaskMemFree(g_pidlLibrary);
		if (g_hCrypt32) {
			FreeLibrary(g_hCrypt32);
		}
		teSysFreeString(&g_bsCmdLine);
	} catch (...) {}
	try {
		Gdiplus::GdiplusShutdown(g_Token);
	} catch (...) {}
	try {
		::OleUninitialize();
	} catch (...) {}
}

HRESULT teExtract(IStorage *pStorage, LPWSTR lpszFolderPath)
{
	STATSTG statstg;
	IEnumSTATSTG *pEnumSTATSTG;
	BYTE pszData[SIZE_BUFF];
#ifdef _2000XP
	IShellFolder2 *pSF2;
	LPITEMIDLIST pidl;
	ULONG chEaten;
	ULONG dwAttributes;
	SYSTEMTIME SysTime;
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
		} else if (statstg.type == STGTY_STREAM) {
			IStream *pStream;
			ULONG uRead;
			DWORD dwWriteByte;

			hr = pStorage->OpenStream(statstg.pwcsName, NULL, STGM_READ, NULL, &pStream);
			if SUCCEEDED(hr) {
				HANDLE hFile = CreateFile(bsPath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
				if (hFile != INVALID_HANDLE_VALUE) {
					while (SUCCEEDED(pStream->Read(pszData, SIZE_BUFF, &uRead)) && uRead) {
						WriteFile(hFile, pszData, uRead, &dwWriteByte, NULL);
					}
#ifdef _2000XP
					if (statstg.mtime.dwLowDateTime == 0) {
						if SUCCEEDED(pStorage->QueryInterface(IID_PPV_ARGS(&pSF2))) {
							if SUCCEEDED(pSF2->ParseDisplayName(NULL, NULL, statstg.pwcsName, &chEaten, &pidl, &dwAttributes)) {
								VariantInit(&v);
								if (SUCCEEDED(pSF2->GetDetailsEx(pidl, &PKEY_DateModified, &v)) && v.vt == VT_DATE) {
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
				} else {
					hr = E_ACCESSDENIED;
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
				LPITEMIDLIST pidl;
				if SUCCEEDED(lpfnSHParseDisplayName(bsPath, NULL, &pidl, 0, NULL)) {
					BSTR bs;
					if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORADDRESSBAR)) {
						if (StrChr(bsPath, '\\') && !StrChr(bs, '\\')) {
							::SysFreeString(bs);
							GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
						}
						lp[0] = '\\';
						*pbsPath = ::SysAllocStringLen(bs, SysStringLen(bs) + lstrlen(lp) + 1);
						lstrcat(*pbsPath, lp);
						::SysFreeString(bs);
					}
					CoTaskMemFree(pidl);
				}
				lp[0] = '\\';
			}
		}
		::SysFreeString(bsPath);
	}
	return *pbsPath != NULL;
}

VOID teFreeLibraries()
{
	VARIANT v;
	VariantInit(&v);
	for (;;) {
		teExecMethod(g_pFreeLibrary, L"shift", &v, 0, NULL);
		if (v.vt == VT_EMPTY) {
			return;
		}
		HMODULE hDll = (HMODULE)GetLLFromVariant(&v);
		VariantClear(&v);
		FreeLibrary(hDll);
	}
}

VOID teGetDisplayNameOf(VARIANT *pv, int uFlags, VARIANT *pVarResult)
{
	try {
		IUnknown *punk;
		if (pv->vt == VT_BSTR) {
			if (!(uFlags & SHGDN_INFOLDER)) {
				VariantCopy(pVarResult, pv);
				return;
			}
		} else if (FindUnknown(pv, &punk)) {
			BSTR bs;
			if (teGetStrFromFolderItem(&bs, punk)) {
				if (uFlags & SHGDN_INFOLDER) {
					teSetSZ(pVarResult, PathFindFileName(bs));
					return;
				}
				if (uFlags & SHGDN_FORADDRESSBAR) {
					BSTR bsLocal;
					if (teLocalizePath(bs, &bsLocal)) {
						teSetSZ(pVarResult, bsLocal);
						::SysFreeString(bsLocal);
						return;
					}
				}
				teSetSZ(pVarResult, bs);
				return;
			}
		}
		LPITEMIDLIST pidl;
		if (teGetIDListFromVariant(&pidl, pv)) {
			GetVarPathFromIDList(pVarResult, pidl, uFlags);
			teCoTaskMemFree(pidl);
		}
	} catch(...) {
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
			if (pVarResult->vt == VT_BSTR && teStartsText(L"search-ms:", pVarResult->bstrVal)) {
				LPITEMIDLIST pidl;
				if (teGetIDListFromObject(pFolderItem, &pidl)) {
					VariantClear(pVarResult);
					GetVarArrayFromIDList(pVarResult, pidl);
					CoTaskMemFree(pidl);
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

//Dispatch API

VOID teApiMemory(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
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
		i = (int)param[0];
		if (i == 0) {
			return;
		}
	}
	BSTR bs = NULL;
	if (nArg >= 1) {
		if (i == 2 && pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
			bs = pDispParams->rgvarg[nArg - 1].bstrVal;
			nCount = SysStringByteLen(bs) / 2 + 1;
		} else {
			nCount = (int)param[1];
		}
		if (nCount < 1) {
			nCount = 1;
		}
		if (nArg >= 2) {
			pc = (char *)param[2];
		}
	}
	pMem = new CteMemory(i * nCount + GetOffsetOfStruct(sz), pc, nCount, sz);
	if (bs) {
		::CopyMemory(pMem->m_pc, bs, nCount * 2);
	}
	teSetObjectRelease(pVarResult, pMem);
}

VOID teApiContextMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDataObject *pDataObj = NULL;
	CteContextMenu *pdispCM;
	if (nArg >= 6) {
		IShellExtInit *pSEI;
		HMODULE hDll = teCreateInstance(&pDispParams->rgvarg[nArg], &pDispParams->rgvarg[nArg - 1], IID_PPV_ARGS(&pSEI));
		if (pSEI) {
			LPITEMIDLIST pidl;
			if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg - 2])) {
				if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg - 3])) {
					HKEY hKey;
					if (RegOpenKeyEx((HKEY)param[4], (LPCWSTR)param[5], 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
						if SUCCEEDED(pSEI->Initialize(pidl, pDataObj, hKey)) {
							IUnknown *punk = NULL;
							FindUnknown(&pDispParams->rgvarg[nArg - 6], &punk);
							pdispCM = new CteContextMenu(pSEI, pDataObj, punk);
							pDataObj = NULL;
							pdispCM->m_pDll = new CteDll(hDll);
							hDll = NULL;
							teSetObjectRelease(pVarResult, pdispCM);
						}
						RegCloseKey(hKey);
					}
				}
				teCoTaskMemFree(pidl);
			}
			pSEI->Release();
		}
		if (hDll) {
			FreeLibrary(hDll);
		}
	} else if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
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

VOID teApiDropTarget(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
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

VOID teApiDataObject(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetObjectRelease(pVarResult, new CteFolderItems(NULL, NULL, true));
}

VOID teApiOleCmdExec(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		IOleCommandTarget *pCT;
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCT))) {
			GUID *pguid = NULL;
			teCLSIDFromString((LPCOLESTR)param[1], pguid);
			pCT->Exec(pguid, (OLECMDID)param[2], (OLECMDEXECOPT)param[3], &pDispParams->rgvarg[nArg - 4], pVarResult);
			pCT->Release();
		}
	}
}

VOID teApisizeof(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetSizeOf(&pDispParams->rgvarg[nArg]));
}

VOID teApiLowPart(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LARGE_INTEGER li;
	li.QuadPart = param[0];
	teSetLong(pVarResult, li.LowPart);
}

VOID teApiULowPart(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LARGE_INTEGER li;
	li.QuadPart = param[0];
	teSetULL(pVarResult, li.LowPart);
}

VOID teApiHighPart(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LARGE_INTEGER li;
	li.QuadPart = param[0];
	teSetLong(pVarResult, li.HighPart);
}

VOID teApiQuadPart(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 1) {
		LARGE_INTEGER li;
		li.LowPart = (DWORD)param[0];
		li.HighPart = (LONG)param[1];
		teSetLL(pVarResult, li.QuadPart);
		return;
	}
	teSetLL(pVarResult, param[0]);
}

VOID teApiUQuadPart(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 1) {
		LARGE_INTEGER li;
		li.LowPart = (DWORD)param[0];
		li.HighPart = (LONG)param[1];
		teSetULL(pVarResult, li.QuadPart);
		return;
	}
	teSetULL(pVarResult, teGetU(param[0]));
}

VOID teApipvData(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pDispParams->rgvarg[nArg].vt & VT_ARRAY) {
		teSetPtr(pVarResult, pDispParams->rgvarg[nArg].parray->pvData);
	}
}

VOID teApiExecMethod(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
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
			teExecMethod(pdisp, (LPOLESTR)param[1], pVarResult, -nCount, pv);
			for (int i = nCount; i-- > 0;) {
				tePutPropertyAt(pdisp1, nCount - i - 1, &pv[i]);//for Reference
			}
			teClearVariantArgs(nCount, pv);
			pdisp1->Release();
		}
		pdisp->Release();
	}
}

VOID teApiExtract(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HRESULT hr = E_FAIL;
	IStorage *pStorage;
	HMODULE hDll = teCreateInstance(&pDispParams->rgvarg[nArg], &pDispParams->rgvarg[nArg - 1], IID_PPV_ARGS(&pStorage));
	if (pStorage) {
		IPersistFile *pPF;
		IPersistFolder *pPF2;
		if SUCCEEDED(pStorage->QueryInterface(IID_PPV_ARGS(&pPF))) {
			if SUCCEEDED(pPF->Load((LPCWSTR)param[2], STGM_READ)) {
				hr = teExtract(pStorage, (LPWSTR)param[3]);
			}
			pPF->Release();
		} else if SUCCEEDED(pStorage->QueryInterface(IID_PPV_ARGS(&pPF2))) {
			LPITEMIDLIST pidl = teILCreateFromPath((LPWSTR)param[2]);
			if (pidl) {
				if SUCCEEDED(pPF2->Initialize(pidl)) {
					hr = teExtract(pStorage, (LPWSTR)param[3]);
				}
				teCoTaskMemFree(pidl);
			}
			pPF2->Release();
		}
		pStorage->Release();
	}
	hDll && FreeLibrary(hDll);
	teSetLong(pVarResult, hr);
}

VOID teApiQuadAdd(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLL(pVarResult, param[0] + param[1]);
}

VOID teApiUQuadAdd(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetULL(pVarResult, teGetU(param[0]) + teGetU(param[1]));
}

VOID teApiUQuadSub(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLL(pVarResult, teGetU(param[0]) - teGetU(param[1]));
}

VOID teApiUQuadCmp(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LONGLONG ll = teGetU(param[0]) - teGetU(param[1]);
	int i = 0;
	if (ll > 0) {
		i = 1;
	} else if (ll < 0) {
		i = -1;
	}
	teSetLong(pVarResult, i);
}

VOID teApiFindWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, FindWindow((LPCWSTR)param[0], (LPCWSTR)param[1]));
}

VOID teApiFindWindowEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, FindWindowEx((HWND)param[0], (HWND)param[1], (LPCWSTR)param[2], (LPCWSTR)param[3]));
}

VOID teApiOleGetClipboard(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		IDataObject *pDataObj;
		if (OleGetClipboard(&pDataObj) == S_OK) {
			FolderItems *pFolderItems = new CteFolderItems(pDataObj, NULL, false);
			if (teSetObject(pVarResult, pFolderItems)) {
				STGMEDIUM Medium;
				if (pDataObj->GetData(&DROPEFFECTFormat, &Medium) == S_OK) {
					VARIANT v;
					VariantInit(&v);
					DWORD *pdwEffect = (DWORD *)GlobalLock(Medium.hGlobal);
					try {
						teSetLong(&v, *pdwEffect);
					} catch (...) {}
					GlobalUnlock(Medium.hGlobal);
					ReleaseStgMedium(&Medium);
					IUnknown *punk;
					if (FindUnknown(pVarResult, &punk)) {
						tePutProperty(punk, L"dwEffect", &v);
					}
				}
			}
			pFolderItems->Release();
			pDataObj->Release();
		}
	}
}

VOID teApiOleSetClipboard(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LONG lResult = E_FAIL;
	IDataObject *pDataObj;
	if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg])) {
		lResult = OleSetClipboard(pDataObj);
		teSysFreeString(&g_bsClipRoot);
		teGetRootFromDataObj(&g_bsClipRoot, pDataObj);
		pDataObj->Release();
	}
	teSetLong(pVarResult, lResult);
}

VOID teApiPathMatchSpec(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, tePathMatchSpec((LPCWSTR)param[0], (LPCWSTR)param[1]));
}

VOID teApiCommandLineToArgv(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	int nLen = 0;
	LPTSTR *lplpszArgs = NULL;
	IDispatch *pArray;
	GetNewArray(&pArray);
	if (param[0] && ((LPCWSTR)param[0])[0]) {
		lplpszArgs = CommandLineToArgvW((LPCWSTR)param[0], &nLen);
		for (int i = nLen; i-- > 0;) {
			VARIANT v;
			teSetSZ(&v, lplpszArgs[i]);
			int n = ::SysStringLen(v.bstrVal);
			if (v.bstrVal[n - 1] == '"') {
				v.bstrVal[n - 1] = '\\';
			}
			tePutPropertyAt(pArray, i, &v);
			VariantClear(&v);
		}
		LocalFree(lplpszArgs);
	}
	teSetObjectRelease(pVarResult, new CteDispatchEx(pArray));
	pArray->Release();
}

VOID teApiILIsEqual(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	FolderItem *pid1, *pid2;
	GetFolderItemFromVariant(&pid1, &pDispParams->rgvarg[nArg]);
	GetFolderItemFromVariant(&pid2, &pDispParams->rgvarg[nArg - 1]);
	teSetBool(pVarResult, teILIsEqual(pid1, pid2));
	pid2->Release();
	pid1->Release();
}

VOID teApiILClone(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	FolderItem *pid;
	GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
	teSetObjectRelease(pVarResult, pid);
}

VOID teApiILIsParent(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl, pidl2;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		if (teGetIDListFromVariant(&pidl2, &pDispParams->rgvarg[nArg - 1])) {
			teSetBool(pVarResult, ::ILIsParent(pidl, pidl2, (BOOL)param[2]));
			teCoTaskMemFree(pidl2);
		}
		teCoTaskMemFree(pidl);
	}
}

VOID teApiILRemoveLastID(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		if (!ILIsEmpty(pidl)) {
			if (::ILRemoveLastID(pidl)) {
				teSetIDList(pVarResult, pidl);
			}
		}
		teCoTaskMemFree(pidl);
	}
}

VOID teApiILFindLastID(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
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

VOID teApiILIsEmpty(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BOOL b = TRUE;
	LPITEMIDLIST pidl;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		b = ILIsEmpty(pidl);
		teCoTaskMemFree(pidl);
	}
	teSetBool(pVarResult, b);
}

VOID teApiILGetParent(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
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

VOID teApiFindFirstFile(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, FindFirstFile((LPCWSTR)param[0], (LPWIN32_FIND_DATA)param[1]));
}

VOID teApiWindowFromPoint(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg]);
	teSetPtr(pVarResult, WindowFromPoint(pt));
}

VOID teApiGetThreadCount(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, g_nThreads);
}

VOID teApiDoEvents(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
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

VOID teApisscanf(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		LPWSTR pszData = (LPWSTR)param[0];
		LPWSTR pszFormat = (LPWSTR)param[1];
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
					teSetBSTR(pVarResult, bs, -1);
				}
			}
		}
	}
}

VOID teApiSetFileTime(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BOOL bResult = FALSE;
	LPITEMIDLIST pidl;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		BSTR bs;
		if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORPARSING)) {
			HANDLE hFile = ::CreateFile(bs, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0);
			if (hFile != INVALID_HANDLE_VALUE) {
				FILETIME **ppft = new LPFILETIME[3];
				for (int i = 3; i--;) {
					ppft[i] = NULL;
					DATE dt;
					if (teGetVariantTime(&pDispParams->rgvarg[nArg - i - 1], &dt) && dt >= 0) {
						ppft[i] = new FILETIME[1];
						teVariantTimeToFileTime(dt, ppft[i]);
					}
				}
				bResult = ::SetFileTime(hFile, ppft[0], ppft[1], ppft[2]);
				::CloseHandle(hFile);
			}
			::SysFreeString(bs);
		}
	}
	teSetBool(pVarResult, bResult);
}

VOID teApiILCreateFromPath(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	FolderItem *pid;
	GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
	teSetObjectRelease(pVarResult, pid);
}

VOID teApiGetProcObject(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HMODULE hDll = LoadLibrary((LPCWSTR)param[0]);
	if (hDll) {
		CHAR szProcNameA[MAX_LOADSTRING];
		LPSTR lpProcNameA = szProcNameA;
		if (pDispParams->rgvarg[nArg - 1].vt == VT_BSTR) {
			WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)pDispParams->rgvarg[nArg - 1].bstrVal, -1, lpProcNameA, MAX_LOADSTRING, NULL, NULL);
		} else {
			lpProcNameA = MAKEINTRESOURCEA(param[1]);
		}
		LPFNGetProcObjectW lpfnGetProcObjectW = (LPFNGetProcObjectW)GetProcAddress(hDll, lpProcNameA);
		if (lpfnGetProcObjectW) {
			if (nArg >= 2) {
				VariantCopy(pVarResult, &pDispParams->rgvarg[nArg - 2]);
			}
			lpfnGetProcObjectW(pVarResult);
			IDispatch *pdisp;
			if (GetDispatch(pVarResult, &pdisp)) {
				CteDispatch *odisp = new CteDispatch(pdisp, 4);
				odisp->m_pDll = new CteDll(hDll);
				hDll = 0;
				VariantClear(pVarResult);
				teSetObjectRelease(pVarResult, odisp);
			}
		}
		if (hDll) {
			FreeLibrary(hDll);
		}
	}
}

VOID teApiSetCurrentDirectory(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetCurrentDirectory((LPCWSTR)param[0]));//use wsh.CurrentDirectory
}

VOID teApiSetDllDirectory(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (lpfnSetDllDirectoryW) {
		teSetBool(pVarResult, lpfnSetDllDirectoryW((LPCWSTR)param[0]));
	}
}

VOID teApiPathIsNetworkPath(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, tePathIsNetworkPath((LPCWSTR)param[0]));
}

VOID teApiRegisterWindowMessage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, RegisterWindowMessage((LPCWSTR)param[0]));
}

VOID teApiStrCmpI(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, lstrcmpi((LPCWSTR)param[0], (LPCWSTR)param[1]));
}

VOID teApiStrLen(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, lstrlen((LPCWSTR)param[0]));
}

VOID teApiStrCmpNI(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, StrCmpNI((LPCWSTR)param[0], (LPCWSTR)param[1], (int)param[2]));
}

VOID teApiStrCmpLogical(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, StrCmpLogicalW((LPCWSTR)param[0], (LPCWSTR)param[1]));
}

VOID teApiPathQuoteSpaces(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0]) {
		BSTR bsResult = ::SysAllocStringLen((BSTR)param[0], ::SysStringLen((BSTR)param[0]) + 3);
		PathQuoteSpaces(bsResult);
		teSetBSTR(pVarResult, bsResult, -1);
	}
}

VOID teApiPathUnquoteSpaces(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0]) {
		BSTR bsResult = ::SysAllocString((OLECHAR *)param[0]);
		PathUnquoteSpaces(bsResult);
		teSetBSTR(pVarResult, bsResult, -1);
	}
}

VOID teApiGetShortPathName(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0]) {
		int nLen = GetShortPathName((LPCWSTR)param[0], NULL, 0);
		BSTR bsResult = ::SysAllocStringLen(NULL, nLen + MAX_PATH);
		nLen = GetShortPathName((LPCWSTR)param[0], bsResult, nLen);
		teSetBSTR(pVarResult, bsResult, -1);
	}
}

VOID teApiPathCreateFromUrl(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0]) {
		DWORD dwLen = ::SysStringLen((BSTR)param[0]) + MAX_PATH;
		BSTR bsResult = ::SysAllocStringLen(NULL, dwLen);
		PathCreateFromUrl((LPCWSTR)param[0], bsResult, &dwLen, NULL);
		teSetBSTR(pVarResult, bsResult, dwLen);
	}
}

VOID teApiPathSearchAndQualify(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0]) {
		UINT uLen = ::SysStringLen((BSTR)param[0]) + MAX_PATH;
		BSTR bsResult = ::SysAllocStringLen((OLECHAR *)param[0], uLen);
		PathSearchAndQualify((LPCWSTR)param[0], bsResult, uLen);
		teSetBSTR(pVarResult, bsResult, -1);
	}
}

VOID teApiPSFormatForDisplay(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0]) {
		PROPERTYKEY propKey;
		if SUCCEEDED(lpfnPSPropertyKeyFromStringEx((LPCWSTR)param[0], &propKey)) {
			if (lpfnPSGetPropertyDescription) {
				IPropertyDescription *pdesc;
				if SUCCEEDED(lpfnPSGetPropertyDescription(propKey, IID_PPV_ARGS(&pdesc))) {
					LPWSTR psz;
					if SUCCEEDED(pdesc->FormatForDisplay(*((PROPVARIANT *)&pDispParams->rgvarg[nArg - 1]), (PROPDESC_FORMAT_FLAGS)param[2], &psz)) {
						teSetSZ(pVarResult, psz);
						teCoTaskMemFree(psz);
					}
					pdesc->Release();
				}
			}
#ifdef _2000XP
			else {
				BSTR bsResult = ::SysAllocStringLen(NULL, MAX_PROP);
				IPropertyUI *pPUI;
				if SUCCEEDED(CoCreateInstance(CLSID_PropertiesUI, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&pPUI))) {
					pPUI->FormatForDisplay(propKey.fmtid, propKey.pid, (PROPVARIANT *)&pDispParams->rgvarg[nArg - 1], (PROPERTYUI_FORMAT_FLAGS)param[2], bsResult, MAX_PROP);
					pPUI->Release();
				}
				teSetBSTR(pVarResult, bsResult, -1);
			}
#endif
		}
	}
}

VOID teApiPSGetDisplayName(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0] && pVarResult) {
		PROPERTYKEY propKey;
		if SUCCEEDED(lpfnPSPropertyKeyFromStringEx((LPCWSTR)param[0], &propKey)) {
			int nFormat = (int)param[1];
			IShellView *pSV = NULL;
			if (nFormat == 0 && g_pTC) {
				CteShellBrowser *pSB = g_pTC->GetShellBrowser(g_pTC->m_nIndex);
				if (pSB) {
					pSV = pSB->m_pShellView;
				}
			}
			pVarResult->bstrVal = tePSGetNameFromPropertyKeyEx(propKey, nFormat, pSV);
			pVarResult->vt = VT_BSTR;
		}
	}
}

VOID teApiGetCursorPos(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetCursorPos((LPPOINT)param[0]));
}

VOID teApiGetKeyboardState(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetKeyboardState((PBYTE)param[0]));
}

VOID teApiSetKeyboardState(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetKeyboardState((PBYTE)param[0]));
}

VOID teApiGetVersionEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, lpfnRtlGetVersion ? !lpfnRtlGetVersion((PRTL_OSVERSIONINFOEXW)param[0]) : GetVersionEx((LPOSVERSIONINFO)param[0]));
}

VOID teApiChooseFont(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ChooseFont((LPCHOOSEFONT)param[0]));
}

VOID teApiChooseColor(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ::ChooseColor((LPCHOOSECOLOR)param[0]));
}

VOID teApiTranslateMessage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ::TranslateMessage((LPMSG)param[0]));
}

VOID teApiShellExecuteEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ::ShellExecuteEx((SHELLEXECUTEINFO *)param[0]));
}

VOID teApiCreateFontIndirect(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::CreateFontIndirect((PLOGFONT)param[0]));
}

VOID teApiCreateIconIndirect(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::CreateIconIndirect((PICONINFO)param[0]));
}

VOID teApiFindText(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::FindText((LPFINDREPLACE)param[0]));
}

VOID teApiFindReplace(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::ReplaceText((LPFINDREPLACE)param[0]));
}

VOID teApiDispatchMessage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::DispatchMessage((LPMSG)param[0]));
}

VOID teApiPostMessage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	PostMessage((HWND)param[0], (UINT)param[1], (WPARAM)param[2], (LPARAM)param[3]);
}

VOID teApiSHFreeNameMappings(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SHFreeNameMappings((HANDLE)param[0]);
}

VOID teApiCoTaskMemFree(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teCoTaskMemFree((LPVOID)param[0]);
}

VOID teApiSleep(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	Sleep((DWORD)param[0]);
}

VOID teApiShRunDialog(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (lpfnSHRunDialog) {
		lpfnSHRunDialog((HWND)param[0], (HICON)param[1], (LPWSTR)param[2], (LPWSTR)param[3], (LPWSTR)param[4], (DWORD)param[5]);
	}
}

VOID teApiDragAcceptFiles(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DragAcceptFiles((HWND)param[0], (BOOL)param[1]);
}

VOID teApiDragFinish(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DragFinish((HDROP)param[0]);
}

VOID teApimouse_event(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	mouse_event((DWORD)param[0], (DWORD)param[1], (DWORD)param[2], (DWORD)param[3], (ULONG_PTR)param[4]);
}

VOID teApikeybd_event(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	keybd_event((BYTE)param[0], (BYTE)param[1], (DWORD)param[2], (ULONG_PTR)param[3]);
}

VOID teApiSHChangeNotify(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	VARIANT pv[2];
	BYTE type = (BYTE)param[1];
	for (int i = 2; i--;) {
		VariantInit(&pv[i]);
		pv[i].bstrVal = NULL;
		if (type == SHCNF_IDLIST) {
			teGetIDListFromVariant((LPITEMIDLIST *)&pv[i].bstrVal, &pDispParams->rgvarg[nArg - i - 2]);
		} else if (type == SHCNF_PATH) {
			teVariantChangeType(&pv[i], &pDispParams->rgvarg[nArg - i - 2], VT_BSTR);
		}
	}
	SHChangeNotify((LONG)param[0], (UINT)param[1], pv[0].bstrVal, pv[1].bstrVal);
	for (int i = 2; i--;) {
		if (pv[i].vt != VT_BSTR) {
			teCoTaskMemFree(pv[i].bstrVal);
		} else {
			VariantClear(&pv[i]);
		}
	}
}

VOID teApiDrawIcon(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DrawIcon((HDC)param[0], (int)param[1], (int)param[2], (HICON)param[3]));
}

VOID teApiDestroyMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyMenu((HMENU)param[0]));
}

VOID teApiFindClose(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, FindClose((HANDLE)param[0]));
}

VOID teApiFreeLibrary(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 0) {
		VARIANT *pv = GetNewVARIANT(1);
		teSetLL(pv, param[0]);
		teExecMethod(g_pFreeLibrary, L"push", NULL, 1, pv);
		SetTimer(g_hwndMain, TET_FreeLibrary, (UINT)param[1], teTimerProc);
		return;
	}
	teSetBool(pVarResult, FreeLibrary((HMODULE)param[0]));
}

VOID teApiImageList_Destroy(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_Destroy((HIMAGELIST)param[0]));
}

VOID teApiDeleteObject(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteObject((HGDIOBJ)param[0]));
}

VOID teApiDestroyIcon(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyIcon((HICON)param[0]));
}

VOID teApiIsWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsWindow((HWND)param[0]));
}

VOID teApiIsWindowVisible(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsWindowVisible((HWND)param[0]));
}

VOID teApiIsZoomed(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsZoomed((HWND)param[0]));
}

VOID teApiIsIconic(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsIconic((HWND)param[0]));
}

VOID teApiOpenIcon(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, OpenIcon((HWND)param[0]));
}

VOID teApiSetForegroundWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teSetForegroundWindow((HWND)param[0]));
}

VOID teApiBringWindowToTop(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, BringWindowToTop((HWND)param[0]));
}

VOID teApiDeleteDC(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteDC((HDC)param[0]));
}

VOID teApiCloseHandle(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, CloseHandle((HANDLE)param[0]));
}

VOID teApiIsMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsMenu((HMENU)param[0]));
}

VOID teApiMoveWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, MoveWindow((HWND)param[0], (int)param[1], (int)param[2], (int)param[3], (int)param[4], (BOOL)param[5]));
}

VOID teApiSetMenuItemBitmaps(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuItemBitmaps((HMENU)param[0], (UINT)param[1], (UINT)param[2], (HBITMAP)param[3], (HBITMAP)param[4]));
}

VOID teApiShowWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teShowWindow((HWND)param[0], (int)param[1]));
}

VOID teApiDeleteMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteMenu((HMENU)param[0], (UINT)param[1], (UINT)param[2]));
}

VOID teApiRemoveMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, RemoveMenu((HMENU)param[0], (UINT)param[1], (UINT)param[2]));
}

VOID teApiDrawIconEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DrawIconEx((HDC)param[0], (int)param[1], (int)param[2], (HICON)param[3],
	(int)param[4], (int)param[5], (UINT)param[6], (HBRUSH)param[7], (UINT)param[8]));
}

VOID teApiEnableMenuItem(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, EnableMenuItem((HMENU)param[0], (UINT)param[1], (UINT)param[2]));
}

VOID teApiImageList_Remove(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_Remove((HIMAGELIST)param[0], (int)param[1]));
}

VOID teApiImageList_SetIconSize(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_SetIconSize((HIMAGELIST)param[0], (int)param[1], (int)param[2]));
}

VOID teApiImageList_Draw(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_Draw((HIMAGELIST)param[0], (int)param[1], (HDC)param[2],
	(int)param[3], (int)param[4], (UINT)param[5]));
}

VOID teApiImageList_DrawEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_DrawEx((HIMAGELIST)param[0], (int)param[1], (HDC)param[2],
	(int)param[3], (int)param[4], (int)param[5], (int)param[6],
	(COLORREF)param[7], (COLORREF)param[8], (UINT)param[9]));
}

VOID teApiImageList_SetImageCount(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_SetImageCount((HIMAGELIST)param[0], (UINT)param[1]));
}

VOID teApiImageList_SetOverlayImage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_SetOverlayImage((HIMAGELIST)param[0], (int)param[1], (int)param[2]));
}

VOID teApiImageList_Copy(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_Copy((HIMAGELIST)param[0], (int)param[1],
	(HIMAGELIST)param[2], (int)param[3], (UINT)param[4]));
}

VOID teApiDestroyWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyWindow((HWND)param[0]));
}

VOID teApiLineTo(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, LineTo((HDC)param[0], (int)param[1], (int)param[2]));
}

VOID teApiReleaseCapture(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ReleaseCapture());
}

VOID teApiSetCursorPos(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetCursorPos((int)param[0], (int)param[1]));
}

VOID teApiDestroyCursor(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyCursor((HCURSOR)param[0]));
}

VOID teApiSHFreeShared(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SHFreeShared((HANDLE)param[0], (DWORD)param[1]));
}

VOID teApiEndMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, EndMenu());
}

VOID teApiSHChangeNotifyDeregister(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SHChangeNotifyDeregister((ULONG)param[0]));
}

VOID teApiSHChangeNotification_Unlock(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SHChangeNotification_Unlock((HANDLE)param[0]));
}

VOID teApiIsWow64Process(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (lpfnIsWow64Process) {
		BOOL bResult;
		lpfnIsWow64Process((HANDLE)param[0], &bResult);
		teSetBool(pVarResult, bResult);
	}
}

VOID teApiBitBlt(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, BitBlt((HDC)param[0], (int)param[1], (int)param[2], (int)param[3], (int)param[4],
	(HDC)param[5], (int)param[6], (int)param[7], (DWORD)param[8]));
}

VOID teApiImmSetOpenStatus(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImmSetOpenStatus((HIMC)param[0], (BOOL)param[1]));
}

VOID teApiImmReleaseContext(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImmReleaseContext((HWND)param[0], (HIMC)param[1]));
}

VOID teApiIsChild(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsChild((HWND)param[0], (HWND)param[1]));
}

VOID teApiKillTimer(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, KillTimer((HWND)param[0], (UINT_PTR)param[1]));
}

VOID teApiAllowSetForegroundWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, AllowSetForegroundWindow((DWORD)param[0]));
}

VOID teApiSetWindowPos(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetWindowPos((HWND)param[0], (HWND)param[1], (int)param[2], (int)param[3], (int)param[4], (int)param[5], (UINT)param[6]));
}

VOID teApiInsertMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, InsertMenu((HMENU)param[0], (UINT)param[1], (UINT)param[2],
	(UINT_PTR)param[3], (LPCWSTR)param[4]));
}

VOID teApiSetWindowText(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if ((HWND)param[0] == g_hwndMain && !param[2]) {
		if (param[1]) {
			lstrcpyn(g_szTitle, (LPCWSTR)param[1], _countof(g_szTitle) - 1);
		} else {
			g_szTitle[0] = NULL;
		}
		SetTimer(g_hwndMain, TET_Title, 100, teTimerProc);
		return;
	}
	teSetBool(pVarResult, SetWindowText((HWND)param[0], (LPCWSTR)param[1]));
}

VOID teApiRedrawWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (!g_bSetRedraw && g_hwndMain == (HWND)param[0]) {
		teSetBool(pVarResult, FALSE);
	} else {
		teSetBool(pVarResult, RedrawWindow((HWND)param[0], (LPRECT)param[1], (HRGN)param[2], (UINT)param[3]));
	}
}

VOID teApiMoveToEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, MoveToEx((HDC)param[0], (int)param[1], (int)param[2], (LPPOINT)param[3]));
}

VOID teApiInvalidateRect(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, InvalidateRect((HWND)param[0], (LPRECT)param[1], (BOOL)param[2]));
}

VOID teApiSendNotifyMessage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SendNotifyMessage((HWND)param[0], (UINT)param[1], (WPARAM)param[2], (LPARAM)param[3]));
}

VOID teApiPeekMessage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, PeekMessage((LPMSG)param[0], (HWND)param[1], (UINT)param[2], (UINT)param[3], (UINT)param[4]));
}

VOID teApiSHFileOperation(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 4) {
		LPSHFILEOPSTRUCT pFO = new SHFILEOPSTRUCT[1];
		::ZeroMemory(pFO, sizeof(SHFILEOPSTRUCT));
		pFO->hwnd = g_hwndMain;
		pFO->wFunc = (UINT)param[0];
		BSTR bs = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]);
		int nLen = ::SysStringLen(bs) + 1;
		bs = ::SysAllocStringLen(bs, nLen);
		bs[nLen] = 0;
		pFO->pFrom = bs;
		bs = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 2]);
		nLen = ::SysStringLen(bs) + 1;
		bs = ::SysAllocStringLen(bs, nLen);
		bs[nLen] = 0;
		pFO->pTo = bs;
		pFO->fFlags = (FILEOP_FLAGS)param[3];
		if (param[4]) {
			teSetPtr(pVarResult, _beginthread(&threadFileOperation, 0, pFO));
			return;
		}
		try {
			teSetLong(pVarResult, teSHFileOperation(pFO));
		} catch (...) {}
		::SysFreeString(const_cast<BSTR>(pFO->pTo));
		::SysFreeString(const_cast<BSTR>(pFO->pFrom));
		delete [] pFO;
		return;
	}
	teSetLong(pVarResult, teSHFileOperation((LPSHFILEOPSTRUCT)param[0]));
}

VOID teApiGetMessage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMessage((LPMSG)param[0], (HWND)param[1], (UINT)param[2], (UINT)param[3]));
}

VOID teApiGetWindowRect(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetWindowRect((HWND)param[0], (LPRECT)param[1]));
}

VOID teApiGetWindowThreadProcessId(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetWindowThreadProcessId((HWND)param[0], (LPDWORD)param[1]));
}

VOID teApiGetClientRect(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetClientRect((HWND)param[0], (LPRECT)param[1]));
}

VOID teApiSendInput(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SendInput((UINT)param[0], (LPINPUT)param[1], (int)param[2]));
}

VOID teApiScreenToClient(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ScreenToClient((HWND)param[0], (LPPOINT)param[1]));
}

VOID teApiMsgWaitForMultipleObjectsEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, MsgWaitForMultipleObjectsEx((DWORD)param[0], (LPHANDLE)param[1], (DWORD)param[2], (DWORD)param[3], (DWORD)param[4]));
}

VOID teApiClientToScreen(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ClientToScreen((HWND)param[0], (LPPOINT)param[1]));
}

VOID teApiGetIconInfo(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetIconInfo((HICON)param[0], (PICONINFO)param[1]));
}

VOID teApiFindNextFile(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, FindNextFile((HANDLE)param[0], (LPWIN32_FIND_DATAW)param[1]));
}

VOID teApiFillRect(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, FillRect((HDC)param[0], (LPRECT)param[1], (HBRUSH)param[2]));
}

VOID teApiShell_NotifyIcon(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, Shell_NotifyIcon((DWORD)param[0], (PNOTIFYICONDATA)param[1]));
}

VOID teApiEndPaint(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	PAINTSTRUCT *ps = (LPPAINTSTRUCT)param[1];
	teSetBool(pVarResult, EndPaint((HWND)param[0], ps));
}

VOID teApiImageList_GetIconSize(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 2) {
		teSetBool(pVarResult, ImageList_GetIconSize((HIMAGELIST)param[0], (int *)param[1], (int *)param[2]));
	} else {
		teSetBool(pVarResult, ImageList_GetIconSize((HIMAGELIST)param[0], (int *)&(((LPSIZE)param[1])->cx), (int *)&(((LPSIZE)param[1])->cy)));
	}
}

VOID teApiGetMenuInfo(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMenuInfo((HMENU)param[0], (LPMENUINFO)param[1]));
}

VOID teApiSetMenuInfo(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuInfo((HMENU)param[0], (LPCMENUINFO)param[1]));
}

VOID teApiSystemParametersInfo(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SystemParametersInfo((UINT)param[0], (UINT)param[1], (PVOID)param[2], (UINT)param[3]));
}

VOID teApiGetTextExtentPoint32(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetTextExtentPoint32((HDC)param[0], (LPCWSTR)param[1], lstrlen((LPCWSTR)param[1]), (LPSIZE)param[2]));
}

VOID teApiSHGetDataFromIDList(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl;
	teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg]);
	if (pidl) {
		HRESULT hr = E_FAIL;
		IShellFolder *pSF;
		LPCITEMIDLIST pidlPart;
		if SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&pSF), &pidlPart)) {
			hr = teSHGetDataFromIDList(pSF, pidlPart, (int)param[1], (PVOID)param[2], (int)param[3]);
			pSF->Release();
		}
		teSetLong(pVarResult, hr);
		teCoTaskMemFree(pidl);
	}
}

VOID teApiInsertMenuItem(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, InsertMenuItem((HMENU)param[0], (UINT)param[1], (BOOL)param[2], (LPCMENUITEMINFO)param[3]));
}

VOID teApiGetMenuItemInfo(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMenuItemInfo((HMENU)param[0], (UINT)param[1], (BOOL)param[2], (LPMENUITEMINFO)param[3]));
}

VOID teApiSetMenuItemInfo(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuItemInfo((HMENU)param[0], (UINT)param[1], (BOOL)param[2], (LPMENUITEMINFO)param[3]));
}

VOID teApiChangeWindowMessageFilterEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teChangeWindowMessageFilterEx((HWND)param[0], (UINT)param[1], (DWORD)param[2], (PCHANGEFILTERSTRUCT)param[3]));
}

VOID teApiImageList_SetBkColor(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_SetBkColor((HIMAGELIST)param[0], (COLORREF)param[1]));
}

VOID teApiImageList_AddIcon(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_AddIcon((HIMAGELIST)param[0], (HICON)param[1]));
}

VOID teApiImageList_Add(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_Add((HIMAGELIST)param[0], (HBITMAP)param[1], (HBITMAP)param[2]));
}

VOID teApiImageList_AddMasked(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_AddMasked((HIMAGELIST)param[0], (HBITMAP)param[1], (COLORREF)param[2]));
}

VOID teApiGetKeyState(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetKeyState((int)param[0]));
}

VOID teApiGetSystemMetrics(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetSystemMetrics((int)param[0]));
}

VOID teApiGetSysColor(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetSysColor((int)param[0]));
}

VOID teApiGetMenuItemCount(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMenuItemCount((HMENU)param[0]));
}

VOID teApiImageList_GetImageCount(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_GetImageCount((HIMAGELIST)param[0]));
}

VOID teApiReleaseDC(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ReleaseDC((HWND)param[0], (HDC)param[1]));
}

VOID teApiGetMenuItemID(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMenuItemID((HMENU)param[0], (int)param[1]));
}

VOID teApiImageList_Replace(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_Replace((HIMAGELIST)param[0], (int)param[1],
		(HBITMAP)param[2], (HBITMAP)param[3]));
}

VOID teApiImageList_ReplaceIcon(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_ReplaceIcon((HIMAGELIST)param[0], (int)param[1], (HICON)param[2]));
}

VOID teApiSetBkMode(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetBkMode((HDC)param[0], (int)param[1]));
}

VOID teApiSetBkColor(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetBkColor((HDC)param[0], (COLORREF)param[1]));
}

VOID teApiSetTextColor(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetTextColor((HDC)param[0], (COLORREF)param[1]));
}

VOID teApiMapVirtualKey(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, MapVirtualKey((UINT)param[0], (UINT)param[1]));
}

VOID teApiWaitForInputIdle(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, WaitForInputIdle((HANDLE)param[0], (DWORD)param[1]));
}

VOID teApiWaitForSingleObject(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, WaitForSingleObject((HANDLE)param[0], (DWORD)param[1]));
}

VOID teApiGetMenuDefaultItem(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMenuDefaultItem((HMENU)param[0], (UINT)param[1], (UINT)param[2]));
}

VOID teApiSetMenuDefaultItem(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuDefaultItem((HMENU)param[0], (UINT)param[1], (UINT)param[2]));
}

VOID teApiCRC32(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BYTE *pc;
	int nLen;
	GetpDataFromVariant(&pc, &nLen, &pDispParams->rgvarg[nArg]);
	teSetLong(pVarResult, CalcCrc32(pc, nLen, (UINT)param[1]));
}

VOID teApiSHEmptyRecycleBin(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SHEmptyRecycleBin((HWND)param[0], (LPCWSTR)param[1], (DWORD)param[2]));
}

VOID teApiGetMessagePos(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMessagePos());
}

VOID teApiImageList_GetOverlayImage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	int iResult = -1;
	try {
		((IImageList *)param[0])->GetOverlayImage((int)param[1], &iResult);
	} catch (...) {}
	teSetLong(pVarResult, iResult);
}

VOID teApiImageList_GetBkColor(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_GetBkColor((HIMAGELIST)param[0]));
}

VOID teApiSetWindowTheme(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (lpfnSetWindowTheme) {
		teSetLong(pVarResult, lpfnSetWindowTheme((HWND)param[0], (LPCWSTR)param[1], (LPCWSTR)param[2]));
	}
}

VOID teApiImmGetVirtualKey(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImmGetVirtualKey((HWND)param[0]));
}

VOID teApiGetAsyncKeyState(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetAsyncKeyState((int)param[0]));
}

VOID teApiTrackPopupMenuEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	if (nArg >= 6) {
		if (FindUnknown(&pDispParams->rgvarg[nArg - 6], &punk)) {
			IContextMenu *pCM;
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pCM))) {
				IDispatch *pdisp;
				GetNewArray(&pdisp);
				teArrayPush(pdisp, pCM);
				pCM->Release();
				pdisp->QueryInterface(IID_PPV_ARGS(&g_pCM));
				pdisp->Release();
			}
			else {
				punk->QueryInterface(IID_PPV_ARGS(&g_pCM));
			}
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
	teSetLong(pVarResult, TrackPopupMenuEx((HMENU)param[0], (UINT)param[1], (int)param[2], (int)param[3],
		(HWND)param[4], (LPTPMPARAMS)param[5]));
	if (hwnd != g_hwndMain && dwPid1 != dwPid2) {
		if (!(dwExStyle & WS_EX_TOPMOST)) {
			SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
		}
		teSetForegroundWindow(hwnd);
	}
	UnhookWindowsHookEx(g_hMenuKeyHook);
	g_hMenuKeyHook = NULL;
	teReleaseClear(&g_pCM);
}

VOID teApiExtractIconEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ExtractIconEx((LPCWSTR)param[0], (int)param[1], (HICON *)param[2], (HICON *)param[3], (UINT)param[4]));
}

VOID teApiGetObject(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetObject((HWND)param[0], (int)param[1], (LPVOID)param[2]));
}

VOID teApiMultiByteToWideChar(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, MultiByteToWideChar((UINT)param[0], (DWORD)param[1],
		(LPCSTR)param[2], (int)param[3], (LPWSTR)param[4], (int)param[5]));
}

VOID teApiWideCharToMultiByte(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, WideCharToMultiByte((UINT)param[0], (DWORD)param[1], (LPCWSTR)param[2], (int)param[3], 
		(LPSTR)param[4], (int)param[5], (LPCSTR)param[6], (LPBOOL)param[7]));
}

VOID teApiGetAttributesOf(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	ULONG lResult = (ULONG)param[1];
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
					pSF->GetAttributesOf(nCount, (LPCITEMIDLIST *)&ppidllist[1], (SFGAOF *)(&lResult));
					lResult &= (ULONG)param[1];
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

VOID teApiSHDoDragDrop(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDataObject *pDataObj;
	if (GetDataObjFromVariant(&pDataObj, &pDispParams->rgvarg[nArg - 1])) {
		g_pDraggingItems = pDataObj;
		FindUnknown(&pDispParams->rgvarg[nArg - 2], &g_pDraggingCtrl);
		DWORD dwEffect = (DWORD)param[3];
		VARIANT v;
		VariantInit(&v);
		try {
			g_bDropFinished = FALSE;
			teSetLong(pVarResult, SHDoDragDrop((HWND)param[0], pDataObj, static_cast<IDropSource *>(g_pTE), dwEffect, &dwEffect));
		} catch(...) {}
		g_pDraggingCtrl = NULL;
		g_pDraggingItems = NULL;
		IDispatch *pEffect;
		if (nArg >= 4 && GetDispatch(&pDispParams->rgvarg[nArg - 4], &pEffect)) {
			teSetLong(&v, dwEffect);
			tePutPropertyAt(pEffect, 0, &v);
			pEffect->Release();
		}
		pDataObj->Release();
	}
}

VOID teApiCompareIDs(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl1, pidl2;
	if (teGetIDListFromVariant(&pidl1, &pDispParams->rgvarg[nArg - 1])) {
		if (teGetIDListFromVariant(&pidl2, &pDispParams->rgvarg[nArg - 2])) {
			IShellFolder *pDesktopFolder;
			SHGetDesktopFolder(&pDesktopFolder);
			teSetLong(pVarResult, pDesktopFolder->CompareIDs((LPARAM)param[0], pidl1, pidl2));
			pDesktopFolder->Release();
			teCoTaskMemFree(pidl2);
		}
		teCoTaskMemFree(pidl1);
	}
}

VOID teApiExecScript(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	VARIANT *pv = NULL;
	if (nArg >= 2) {
		pv = &pDispParams->rgvarg[nArg - 2];
	}
	IUnknown *pOnError = NULL;
	if (nArg >= 3) {
		FindUnknown(&pDispParams->rgvarg[nArg - 3], &pOnError);
	}
	teSetLong(pVarResult, ParseScript((LPOLESTR)param[0], (LPOLESTR)param[1], pv, pOnError, NULL));
}

VOID teApiGetScriptDispatch(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		VARIANT *pv = NULL;
		if (nArg >= 2) {
			pv = &pDispParams->rgvarg[nArg - 2];
		}
		IUnknown *pOnError = NULL;
		if (nArg >= 3) {
			FindUnknown(&pDispParams->rgvarg[nArg - 3], &pOnError);
		}
		IDispatch *pdisp = NULL;
		ParseScript((LPOLESTR)param[0], (LPOLESTR)param[1], pv, pOnError, &pdisp);
		teSetObjectRelease(pVarResult, pdisp);
	}
}

VOID teApiGetDispatch(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		IDispatch *pdisp;
		if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
			DISPID dispid = DISPID_UNKNOWN;
			LPOLESTR lp = (LPOLESTR)param[1];
			if (pdisp->GetIDsOfNames(IID_NULL, &lp, 1, LOCALE_USER_DEFAULT, &dispid) == S_OK) {
				CteDispatch *oDisp = new CteDispatch(pdisp, 0);
				oDisp->m_dispIdMember = dispid;
				teSetObjectRelease(pVarResult, oDisp);
			}
			pdisp->Release();
		}
	}
}

VOID teApiSHChangeNotifyRegister(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SHChangeNotifyEntry entry;
	teGetIDListFromVariant(const_cast<LPITEMIDLIST *>(&entry.pidl), &pDispParams->rgvarg[nArg - 4]);
	entry.fRecursive = (BOOL)param[5];
	if (entry.pidl) {
		teSetLong(pVarResult, SHChangeNotifyRegister((HWND)param[0], (int)param[1], (LONG)param[2], (UINT)param[3], 1, &entry));
		teChangeWindowMessageFilterEx((HWND)param[0], (UINT)param[3], MSGFLT_ALLOW, NULL);
		teCoTaskMemFree(const_cast<LPITEMIDLIST>(entry.pidl));
	}
}

VOID teApiMessageBox(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, MessageBox((HWND)param[0], (LPCWSTR)param[1], (LPCWSTR)param[2], (UINT)param[3]));
}

VOID teApiImageList_GetIcon(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_GetIcon((HIMAGELIST)param[0], (int)param[1], (UINT)param[2]));
}

VOID teApiImageList_Create(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_Create((int)param[0], (int)param[1], (UINT)param[2], (int)param[3], (int)param[4]));
}

VOID teApiGetWindowLongPtr(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetWindowLongPtr((HWND)param[0], (int)param[1]));
}

VOID teApiGetClassLongPtr(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetClassLongPtr((HWND)param[0], (int)param[1]));
}

VOID teApiGetSubMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetSubMenu((HMENU)param[0], (int)param[1]));
}

VOID teApiSelectObject(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SelectObject((HDC)param[0], (HGDIOBJ)param[1]));
}

VOID teApiGetStockObject(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetStockObject((int)param[0]));
}

VOID teApiGetSysColorBrush(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetSysColorBrush((int)param[0]));
}

VOID teApiSetFocus(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetFocus((HWND)param[0]));
	if ((HWND)param[0] == g_hwndMain) {
		CteShellBrowser *pSB = g_pTC->GetShellBrowser(g_pTC->m_nIndex);
		if (pSB) {
			pSB->SetActive(FALSE);
		}
	}
}

VOID teApiGetDC(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetDC((HWND)param[0]));
}

VOID teApiCreateCompatibleDC(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateCompatibleDC((HDC)param[0]));
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

VOID teApiCreatePopupMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teCreateMenu(CreatePopupMenu(), pVarResult);
}

VOID teApiCreateMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teCreateMenu(CreateMenu(), pVarResult);
}

VOID teApiCreateCompatibleBitmap(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateCompatibleBitmap((HDC)param[0], (int)param[1], (int)param[2]));
}

VOID teApiSetWindowLongPtr(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetWindowLongPtr((HWND)param[0], (int)param[1], (LONG_PTR)param[2]));
}

VOID teApiSetClassLongPtr(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetClassLongPtr((HWND)param[0], (int)param[1], (LONG_PTR)param[2]));
}

VOID teApiImageList_Duplicate(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_Duplicate((HIMAGELIST)param[0]));
}

VOID teApiSendMessage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SendMessage((HWND)param[0], (UINT)param[1], (WPARAM)param[2], (LPARAM)param[3]));
}

VOID teApiGetSystemMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetSystemMenu((HWND)param[0], (BOOL)param[1]));
}

VOID teApiGetWindowDC(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetWindowDC((HWND)param[0]));
}

VOID teApiCreatePen(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreatePen((int)param[0], (int)param[1], (COLORREF)param[2]));
}

VOID teApiSetCapture(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetCapture((HWND)param[0]));
}

VOID teApiSetCursor(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetCursor((HCURSOR)param[0]));
}

VOID teApiCallWindowProc(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CallWindowProc((WNDPROC)param[0], (HWND)param[1], (UINT)param[2], (WPARAM)param[3], (LPARAM)param[4]));
}

VOID teApiGetWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 1) {
		teSetPtr(pVarResult, GetWindow((HWND)param[0], (UINT)param[1]));
		return;
	}
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		HWND hwnd;
		IUnknown_GetWindow(punk, &hwnd);
		teSetPtr(pVarResult, hwnd);
	}
}

VOID teApiGetTopWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetTopWindow((HWND)param[0]));
}

VOID teApiOpenProcess(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, OpenProcess((DWORD)param[0], (BOOL)param[1], (DWORD)param[2]));
}

VOID teApiGetParent(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetParent((HWND)param[0]));
}

VOID teApiGetCapture(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetCapture());
}

VOID teApiGetModuleHandle(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetModuleHandle((LPCWSTR)param[0]));
}

VOID teApiSHGetImageList(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HANDLE h = 0;
	if FAILED(lpfnSHGetImageList((int)param[0], IID_IImageList, (LPVOID *)&h)) {
		h = 0;
	}
	teSetPtr(pVarResult, h);
}

VOID teApiCopyImage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CopyImage((HANDLE)param[0], (UINT)param[1], (int)param[2], (int)param[3], (UINT)param[4]));
}

VOID teApiGetCurrentProcess(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetCurrentProcess());
}

VOID teApiImmGetContext(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImmGetContext((HWND)param[0]));
}

VOID teApiGetFocus(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetFocus());
}

VOID teApiGetForegroundWindow(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetForegroundWindow());
}

VOID teApiSetTimer(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetTimer((HWND)param[0], (UINT_PTR)param[1], (UINT)param[2], NULL));
}

VOID teApiLoadMenu(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadMenu((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teApiLoadIcon(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadIcon((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teApiLoadLibraryEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadLibraryEx(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), (HANDLE)param[1], (DWORD)param[2]));
}

VOID teApiLoadImage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadImage((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
		(UINT)param[2],	(int)param[3], (int)param[4], (UINT)param[5]));
}

VOID teApiImageList_LoadImage(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_LoadImage((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
	(int)param[2], (int)param[3], (COLORREF)param[4], (UINT)param[5], (UINT)param[6]));
}

VOID teApiSHGetFileInfo(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl;
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg]);
	} else {
		pidl = (LPITEMIDLIST)param[0];//string
	}
	teSetPtr(pVarResult, SHGetFileInfo((LPCWSTR)pidl, (DWORD)param[1], (SHFILEINFO *)param[2], (UINT)param[3], (UINT)param[4]));
	if (pidl != (LPITEMIDLIST)param[0]) {
		teCoTaskMemFree(pidl);
	}
}

VOID teApiCreateWindowEx(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateWindowEx((DWORD)param[0], (LPCWSTR)param[1], (LPCWSTR)param[2],
		(DWORD)param[3], (int)param[4], (int)param[5], (int)param[6], (int)param[7],
		(HWND)param[8], (HMENU)param[9], (HINSTANCE)param[10], (LPVOID)param[11]));
}

VOID teApiShellExecute(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ShellExecute((HWND)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
		GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 2]), (LPCWSTR)param[3], (LPCWSTR)param[4], (int)param[5]));
}

VOID teApiBeginPaint(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, BeginPaint((HWND)param[0], (LPPAINTSTRUCT)param[1]));
}

VOID teApiLoadCursor(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadCursor((HINSTANCE)param[0], GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teApiLoadCursorFromFile(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadCursorFromFile(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teAsyncInvoke(WORD wMode, int nArg, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	TEInvoke *pInvoke = new TEInvoke[1];
	if (GetDispatch(&pDispParams->rgvarg[nArg], &pInvoke->pdisp)) {
		int dwms = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
		if (dwms == 0) {
			dwms = g_param[TE_NetworkTimeout];
		}
		pInvoke->dispid = DISPID_VALUE;
		pInvoke->cRef = 2;
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
			if (dwms > 0) {
				SetTimer(g_hwndMain, (UINT_PTR)pInvoke, dwms, teTimerProcParse);
			}
			teSetPtr(pVarResult, _beginthread(&threadParseDisplayName, 0, pInvoke));
			return;
		}
		FolderItem *pid;
		GetFolderItemFromVariant(&pid, &pInvoke->pv[pInvoke->cArgs - 1]);
		VariantClear(&pInvoke->pv[pInvoke->cArgs - 1]);
		teSetObjectRelease(&pInvoke->pv[pInvoke->cArgs - 1], pid);
		Invoke5(pInvoke->pdisp, DISPID_VALUE, DISPATCH_METHOD, NULL, pInvoke->cArgs, pInvoke->pv);
		pInvoke->pdisp->Release();
		teSetPtr(pVarResult, 0);
	}
	delete [] pInvoke;
}

VOID teApiSHParseDisplayName(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teAsyncInvoke(0, nArg, pDispParams, pVarResult);
}

VOID teApiSHChangeNotification_Lock(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDispatch *pdisp;
	if (GetDispatch(&pDispParams->rgvarg[nArg - 2], &pdisp)) {
		PIDLIST_ABSOLUTE *ppidl;
		LONG lEvent;
		teSetPtr(pVarResult, SHChangeNotification_Lock((HANDLE)param[0], (DWORD)param[1], &ppidl, &lEvent));
		VARIANT v;
		teSetIDList(&v, ppidl[0]);
		tePutPropertyAt(pdisp, 0, &v);
		VariantClear(&v);
		teSetIDList(&v, ppidl[1]);
		tePutPropertyAt(pdisp, 1, &v);
		VariantClear(&v);
		teSetLong(&v, lEvent);
		tePutProperty(pdisp, L"lEvent", &v);
		VariantClear(&v);
		pdisp->Release();
	}
}

VOID teApiGetWindowText(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if ((HWND)param[0] == g_hwndMain) {
		teSetSZ(pVarResult, g_szTitle);
		return;
	}
	BSTR bs = NULL;
	int nLen = GetWindowTextLength((HWND)param[0]);
	if (nLen) {
		bs = SysAllocStringLen(NULL, nLen);
		nLen = GetWindowText((HWND)param[0], bs, nLen + 1);
	}
	teSetBSTR(pVarResult, bs, nLen);
}

VOID teApiGetClassName(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = SysAllocStringLen(NULL, MAX_CLASS_NAME);
	int nLen = GetClassName((HWND)param[0], bsResult, MAX_CLASS_NAME);
	teSetBSTR(pVarResult, bsResult, nLen);
}

VOID teApiGetModuleFileName(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = NULL;
	int nLen = teGetModuleFileName((HMODULE)param[0], &bsResult);
	teSetBSTR(pVarResult, bsResult, nLen);
}

VOID teApiGetCommandLine(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBSTR(pVarResult, teGetCommandLine(), -2);
}

VOID teApiGetCurrentDirectory(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = SysAllocStringLen(NULL, MAX_PATH);
	int nLen = GetCurrentDirectory(MAX_PATH, bsResult);
	teSetBSTR(pVarResult, bsResult, nLen);
}

VOID teApiGetMenuString(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = NULL;
	int nLen = teGetMenuString(&bsResult, (HMENU)param[0], (UINT)param[1], param[2] != MF_BYCOMMAND);
	teSetBSTR(pVarResult, bsResult, nLen);
}

VOID teApiGetDisplayNameOf(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teGetDisplayNameOf(&pDispParams->rgvarg[nArg], (int)param[1], pVarResult);
}

VOID teApiGetKeyNameText(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = SysAllocStringLen(NULL, MAX_PATH);
	int nLen = GetKeyNameText((LONG)param[0], bsResult, MAX_PATH);
	teSetBSTR(pVarResult, bsResult, nLen);
}

VOID teApiDragQueryFile(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[1] & ~MAXINT) {
		teSetLong(pVarResult, DragQueryFile((HDROP)param[0], (UINT)param[1], NULL, 0));
		return;
	}
	BSTR bsResult;
	int nLen = teDragQueryFile((HDROP)param[0], (UINT)param[1], &bsResult);
	teSetBSTR(pVarResult, bsResult, nLen);
}

VOID teApiSysAllocString(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetSZ(pVarResult, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]));
}

VOID teApiSysAllocStringLen(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBSTR(pVarResult, SysAllocStringLen(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), (UINT)param[1]), -2);
}

VOID teApiSysAllocStringByteLen(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBSTR(pVarResult, SysAllocStringByteLen((char *)GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), (UINT)param[1]), -2);
}

VOID teApisprintf(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	int nSize = (int)param[0];
	LPWSTR pszFormat = (LPWSTR)param[1];
	if (pszFormat && nSize > 0) {
		BSTR bsResult = SysAllocStringLen(NULL, nSize);
		int nLen = 0;
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
					if (tolower(wc) == 's') {//String
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
					if (!StrChrIW(L"0123456789-+#.hljzt", wc)) {//not Specifier
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
		teSetBSTR(pVarResult, bsResult, nLen);
	}
}

VOID teApibase64_encode(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (lpfnCryptBinaryToStringW) {
		UCHAR *pc;
		int nLen = 0;
		GetpDataFromVariant(&pc, &nLen, &pDispParams->rgvarg[nArg]);
		if (pc) {
			DWORD dwSize;
			lpfnCryptBinaryToStringW(pc, nLen, CRYPT_STRING_BASE64, NULL, &dwSize);
			BSTR bsResult = NULL;
			if (dwSize > 0) {
				bsResult = SysAllocStringLen(NULL, dwSize - 1);
				lpfnCryptBinaryToStringW(pc, nLen, CRYPT_STRING_BASE64, bsResult, &dwSize);
			}
			teSetBSTR(pVarResult, bsResult, -1);
		}
	}
}

VOID teApiLoadString(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = SysAllocStringLen(NULL, SIZE_BUFF);
	int nLen = LoadString((HINSTANCE)param[0], (UINT)param[1], bsResult, SIZE_BUFF);
	teSetBSTR(pVarResult, bsResult, -1);
}

VOID teApiAssocQueryString(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DWORD cch = 0;
	AssocQueryString((ASSOCF)param[0], (ASSOCSTR)param[1], (LPCWSTR)param[2], (LPCWSTR)param[3], NULL, &cch);
	BSTR bsResult = NULL;
	int nLen = -1;
	if (cch > 0) {
		nLen = cch - 1;
		bsResult = SysAllocStringLen(NULL, nLen);
		AssocQueryString((ASSOCF)param[0], (ASSOCSTR)param[1], (LPCWSTR)param[2], (LPCWSTR)param[3], bsResult, &cch);
	}
	teSetBSTR(pVarResult, bsResult, nLen);
}

VOID teApiStrFormatByteSize(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	int nLen = (int)param[1];
	if (nLen <= 0) {
		nLen = MAX_COLUMN_NAME_LEN;
	}
	BSTR bsResult = ::SysAllocStringLen(NULL, nLen);
	StrFormatByteSize(param[0], bsResult, nLen);
	teSetBSTR(pVarResult, bsResult, -1);
}

VOID teApiGetDateFormat(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SYSTEMTIME SysTime;
	if (teGetSystemTime(&pDispParams->rgvarg[nArg - 2], &SysTime)) {
		int cch = GetDateFormat((LCID)param[0], (DWORD)param[1], &SysTime, (LPCWSTR)param[3], NULL, 0);
		if (cch > 0) {
			BSTR bsResult = ::SysAllocStringLen(NULL, cch - 1);
			GetDateFormat((LCID)param[0], (DWORD)param[1], &SysTime, (LPCWSTR)param[3], bsResult, cch);
			teSetBSTR(pVarResult, bsResult, cch - 1);
		}
	}
}

VOID teApiGetTimeFormat(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SYSTEMTIME SysTime;
	if (teGetSystemTime(&pDispParams->rgvarg[nArg - 2], &SysTime)) {
		int cch = GetTimeFormat((LCID)param[0], (DWORD)param[1], &SysTime, (LPCWSTR)param[3], NULL, 0);
		if (cch > 0) {
			BSTR bsResult = ::SysAllocStringLen(NULL, cch - 1);
			GetTimeFormat((LCID)param[0], (DWORD)param[1], &SysTime, (LPCWSTR)param[3], bsResult, cch);
			teSetBSTR(pVarResult, bsResult, cch - 1);
		}
	}
}

VOID teApiGetLocaleInfo(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	int cch = GetLocaleInfo((LCID)param[0], (LCTYPE)param[1], NULL, 0);
	if (cch > 0) {
		BSTR bsResult = ::SysAllocStringLen(NULL, cch - 1);
		GetLocaleInfo((LCID)param[0], (LCTYPE)param[1], bsResult, cch);
		teSetBSTR(pVarResult, bsResult, -1);
	}
}

VOID teApiMonitorFromPoint(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg]);
	teSetPtr(pVarResult, MonitorFromPoint(pt, (DWORD)param[1]));
	
}

VOID teApiMonitorFromRect(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, MonitorFromRect((LPCRECT)param[0], (DWORD)param[1]));
}

VOID teApiGetMonitorInfo(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMonitorInfo((HMONITOR)param[0], (LPMONITORINFO)param[1]));
}

VOID teApiGlobalAddAtom(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GlobalAddAtom((LPCWSTR)param[0]));
}

VOID teApiGlobalGetAtomName(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	TCHAR szBuf[256];
	UINT uSize = GlobalGetAtomName((ATOM)param[0], szBuf, sizeof(szBuf) / sizeof(TCHAR));
	teSetSZ(pVarResult, szBuf);
}

VOID teApiGlobalFindAtom(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GlobalFindAtom((LPCWSTR)param[0]));
}

VOID teApiGlobalDeleteAtom(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GlobalDeleteAtom((ATOM)param[0]));
}

VOID teApiRegisterHotKey(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, RegisterHotKey((HWND)param[0], (int)param[1], (UINT)param[2], (UINT)param[3]));
}

VOID teApiUnregisterHotKey(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, UnregisterHotKey((HWND)param[0], (int)param[1]));
}

VOID teApiILGetCount(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		teSetLong(pVarResult, ILGetCount(pidl));
		teCoTaskMemFree(pidl);
	}
}

VOID teApiSHTestTokenMembership(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (lpfnSHTestTokenMembership) {
		teSetBool(pVarResult, lpfnSHTestTokenMembership((HANDLE)param[0], (ULONG)param[1]));
	}
}

VOID teApiObjGetI(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDispatch *pdisp;
	if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
		teGetPropertyI(pdisp, (LPOLESTR)param[1], pVarResult);
		pdisp->Release();
	}
}

VOID teApiObjPutI(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		teSetLong(pVarResult, tePutProperty0(punk, (LPOLESTR)param[1], &pDispParams->rgvarg[nArg - 2], fdexNameEnsure | fdexNameCaseInsensitive));
	}
}

VOID teApiObjDispId(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DISPID dispid = DISPID_UNKNOWN;
	IDispatch *pdisp;
	if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
		LPOLESTR lpsz = (LPOLESTR)param[1];
		HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &lpsz, 1, LOCALE_USER_DEFAULT, &dispid);
		pdisp->Release();
	}
	teSetLong(pVarResult, dispid);
}

VOID teApiDrawText(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, DrawText((HDC)param[0], (LPCWSTR)param[1], (int)param[2], (LPRECT)param[3], (UINT)param[4]));
}

VOID teApiRectangle(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, Rectangle((HDC)param[0], (int)param[1], (int)param[2], (int)param[3], (int)param[4]));
}

VOID teApiPathIsDirectory(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 2 && pDispParams->rgvarg[nArg].vt == VT_DISPATCH) {
		teAsyncInvoke(1, nArg, pDispParams, pVarResult);
		return;
	}
	HRESULT hr = tePathIsDirectory((LPWSTR)param[0], (DWORD)param[1], (int)param[2]);
	if (hr != E_ABORT) {
		if (hr <= 0) {
			teSetBool(pVarResult, SUCCEEDED(hr));
		} else {
			teSetLong(pVarResult, hr);
		}
	}
}

VOID teApiPathIsSameRoot(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, PathIsSameRoot((LPCWSTR)param[0], (LPCWSTR)param[1]));
}

VOID teApiSHSimpleIDListFromPath(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl;
	if (nArg >= 1 && (param[1] & FILE_ATTRIBUTE_DIRECTORY)) {
		BSTR bs;
		tePathAppend(&bs, (LPWSTR)param[0], L"a");
		pidl = SHSimpleIDListFromPath(bs);
		ILRemoveLastID(pidl);
		::SysFreeString(bs);
	} else {
		pidl = SHSimpleIDListFromPath((LPWSTR)param[0]);
	}
	if (pidl) {
		if (nArg >= 0 && ILGetSize(pidl) > 16) {
			LPWORD pwIdl = (LPWORD)ILFindLastID(pidl);
			pwIdl[6] = (WORD)param[1];
			if (nArg >= 2) {
				DATE dt;
				if (teGetVariantTime(&pDispParams->rgvarg[nArg - 2], &dt) && dt >= 0) {
					FILETIME ft;
					if (teVariantTimeToFileTime(dt, &ft)) {
						::FileTimeToDosDateTime(&ft, &pwIdl[4], &pwIdl[5]);
					}
				}
				if (nArg >= 3) {
					*((UINT *)&pwIdl[2]) = (UINT)param[3];
				}
			}
		}
		teSetIDList(pVarResult, pidl);
		teCoTaskMemFree(pidl);
	}
}

VOID teApiOutputDebugString(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	::OutputDebugString((LPCWSTR)param[0]);
}

VOID teApiDllGetClassObject(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IDispatch *pdisp;
	HMODULE hDll = teCreateInstance(&pDispParams->rgvarg[nArg], &pDispParams->rgvarg[nArg - 1], IID_PPV_ARGS(&pdisp));
	if (hDll) {
		CteDispatch *odisp = new CteDispatch(pdisp, 4);
		odisp->m_pDll = new CteDll(hDll);
		hDll = 0;
		VariantClear(pVarResult);
		teSetObjectRelease(pVarResult, odisp);
	}
}

/*
VOID teApi(int nArg, LONGLONG *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
}

*/

//
TEDispatchApi dispAPI[] = {
	{ 1, -1, -1, -1, L"Memory", teApiMemory },
	{ 1,  5, -1, -1, L"ContextMenu", teApiContextMenu },
	{ 1, -1, -1, -1, L"DropTarget", teApiDropTarget },
	{ 0, -1, -1, -1, L"DataObject", teApiDataObject },
	{ 5,  1, -1, -1, L"OleCmdExec", teApiOleCmdExec },
	{ 1, -1, -1, -1, L"sizeof", teApisizeof },
	{ 1, -1, -1, -1, L"LowPart", teApiLowPart },
	{ 1, -1, -1, -1, L"ULowPart", teApiULowPart },
	{ 1, -1, -1, -1, L"HighPart", teApiHighPart },
	{ 1, -1, -1, -1, L"QuadPart", teApiQuadPart },
	{ 1, -1, -1, -1, L"UQuadPart", teApiUQuadPart },
	{ 1, -1, -1, -1, L"pvData", teApipvData },
	{ 3,  1, -1, -1, L"ExecMethod", teApiExecMethod },
	{ 4,  2,  3, -1, L"Extract", teApiExtract },
	{ 2, -1, -1, -1, L"Add", teApiQuadAdd },
	{ 2, -1, -1, -1, L"UQuadAdd", teApiUQuadAdd },
	{ 2, -1, -1, -1, L"UQuadSub", teApiUQuadSub },
	{ 2, -1, -1, -1, L"UQuadCmp", teApiUQuadCmp },
	{ 2,  0,  1, -1, L"FindWindow", teApiFindWindow },
	{ 4,  2,  3, -1, L"FindWindowEx", teApiFindWindowEx },
	{ 0, -1, -1, -1, L"OleGetClipboard", teApiOleGetClipboard },
	{ 1, -1, -1, -1, L"OleSetClipboard", teApiOleSetClipboard },
	{ 2,  0,  1, -1, L"PathMatchSpec", teApiPathMatchSpec },
	{ 1,  0, -1, -1, L"CommandLineToArgv", teApiCommandLineToArgv },
	{ 2, -1, -1, -1, L"ILIsEqual", teApiILIsEqual },
	{ 1, -1, -1, -1, L"ILClone", teApiILClone },
	{ 3, -1, -1, -1, L"ILIsParent", teApiILIsParent },
	{ 1, -1, -1, -1, L"ILRemoveLastID", teApiILRemoveLastID },
	{ 1, -1, -1, -1, L"ILFindLastID", teApiILFindLastID },
	{ 1, -1, -1, -1, L"ILIsEmpty", teApiILIsEmpty },
	{ 1, -1, -1, -1, L"ILGetParent", teApiILGetParent },
	{ 2,  0, -1, -1, L"FindFirstFile", teApiFindFirstFile },
	{ 1, -1, -1, -1, L"WindowFromPoint", teApiWindowFromPoint },
	{ 0, -1, -1, -1, L"GetThreadCount", teApiGetThreadCount },
	{ 0, -1, -1, -1, L"DoEvents", teApiDoEvents },
	{ 2,  0,  1, -1, L"sscanf", teApisscanf },
	{ 4, -1, -1, -1, L"SetFileTime", teApiSetFileTime },
	{ 1, -1, -1, -1, L"ILCreateFromPath", teApiILCreateFromPath },
	{ 2,  0, -1, -1, L"GetProcObject", teApiGetProcObject },
	{ 1,  0, -1, -1, L"SetDllDirectory", teApiSetDllDirectory },
	{ 1,  0, -1, -1, L"PathIsNetworkPath", teApiPathIsNetworkPath },
	{ 1,  0, -1, -1, L"RegisterWindowMessage", teApiRegisterWindowMessage },
	{ 2,  0,  1, -1, L"StrCmpI", teApiStrCmpI },
	{ 1,  0, -1, -1, L"StrLen", teApiStrLen },
	{ 3,  0,  1, -1, L"StrCmpNI", teApiStrCmpNI },
	{ 2,  0,  1, -1, L"StrCmpLogical", teApiStrCmpLogical },
	{ 1,  0, -1, -1, L"PathQuoteSpaces", teApiPathQuoteSpaces },
	{ 1,  0, -1, -1, L"PathUnquoteSpaces", teApiPathUnquoteSpaces },
	{ 1,  0, -1, -1, L"GetShortPathName", teApiGetShortPathName },
	{ 1,  0, -1, -1, L"PathCreateFromUrl", teApiPathCreateFromUrl },
	{ 1,  0, -1, -1, L"PathSearchAndQualify", teApiPathSearchAndQualify },
	{ 3,  0, -1, -1, L"PSFormatForDisplay", teApiPSFormatForDisplay },
	{ 1,  0, -1, -1, L"PSGetDisplayName", teApiPSGetDisplayName },
	{ 1, -1, -1, -1, L"GetCursorPos", teApiGetCursorPos },
	{ 1, -1, -1, -1, L"GetKeyboardState", teApiGetKeyboardState },
	{ 1, -1, -1, -1, L"SetKeyboardState", teApiSetKeyboardState },
	{ 1, -1, -1, -1, L"GetVersionEx", teApiGetVersionEx },
	{ 1, -1, -1, -1, L"ChooseFont", teApiChooseFont },
	{ 1, -1, -1, -1, L"ChooseColor", teApiChooseColor },
	{ 1, -1, -1, -1, L"TranslateMessage", teApiTranslateMessage },
	{ 1, -1, -1, -1, L"ShellExecuteEx", teApiShellExecuteEx },
	{ 1, -1, -1, -1, L"CreateFontIndirect", teApiCreateFontIndirect },
	{ 1, -1, -1, -1, L"CreateIconIndirect", teApiCreateIconIndirect },
	{ 1, -1, -1, -1, L"FindText", teApiFindText },
	{ 1, -1, -1, -1, L"FindReplace", teApiFindReplace },
	{ 1, -1, -1, -1, L"DispatchMessage", teApiDispatchMessage },
	{ 1, -1, -1, -1, L"PostMessage", teApiPostMessage },
	{ 1, -1, -1, -1, L"SHFreeNameMappings", teApiSHFreeNameMappings },
	{ 1, -1, -1, -1, L"CoTaskMemFree", teApiCoTaskMemFree },
	{ 1, -1, -1, -1, L"Sleep", teApiSleep },
	{ 6,  2,  3,  4, L"ShRunDialog", teApiShRunDialog },
	{ 2, -1, -1, -1, L"DragAcceptFiles", teApiDragAcceptFiles },
	{ 1, -1, -1, -1, L"DragFinish", teApiDragFinish },
	{ 2, -1, -1, -1, L"mouse_event", teApimouse_event },
	{ 4, -1, -1, -1, L"keybd_event", teApikeybd_event },
	{ 4, -1, -1, -1, L"SHChangeNotify", teApiSHChangeNotify },
	{ 4, -1, -1, -1, L"DrawIcon", teApiDrawIcon },
	{ 1, -1, -1, -1, L"DestroyMenu", teApiDestroyMenu },
	{ 1, -1, -1, -1, L"FindClose", teApiFindClose },
	{ 1, -1, -1, -1, L"FreeLibrary", teApiFreeLibrary },
	{ 1, -1, -1, -1, L"ImageList_Destroy", teApiImageList_Destroy },
	{ 1, -1, -1, -1, L"DeleteObject", teApiDeleteObject },
	{ 1, -1, -1, -1, L"DestroyIcon", teApiDestroyIcon },
	{ 1, -1, -1, -1, L"IsWindow", teApiIsWindow },
	{ 1, -1, -1, -1, L"IsWindowVisible", teApiIsWindowVisible },
	{ 1, -1, -1, -1, L"IsZoomed", teApiIsZoomed },
	{ 1, -1, -1, -1, L"IsIconic", teApiIsIconic },
	{ 1, -1, -1, -1, L"OpenIcon", teApiOpenIcon },
	{ 1, -1, -1, -1, L"SetForegroundWindow", teApiSetForegroundWindow },
	{ 1, -1, -1, -1, L"BringWindowToTop", teApiBringWindowToTop },
	{ 1, -1, -1, -1, L"DeleteDC", teApiDeleteDC },
	{ 1, -1, -1, -1, L"CloseHandle", teApiCloseHandle },
	{ 1, -1, -1, -1, L"IsMenu", teApiIsMenu },
	{ 6, -1, -1, -1, L"MoveWindow", teApiMoveWindow },
	{ 5, -1, -1, -1, L"SetMenuItemBitmaps", teApiSetMenuItemBitmaps },
	{ 2, -1, -1, -1, L"ShowWindow", teApiShowWindow },
	{ 3, -1, -1, -1, L"DeleteMenu", teApiDeleteMenu },
	{ 3, -1, -1, -1, L"RemoveMenu", teApiRemoveMenu },
	{ 9, -1, -1, -1, L"DrawIconEx", teApiDrawIconEx },
	{ 3, -1, -1, -1, L"EnableMenuItem", teApiEnableMenuItem },
	{ 2, -1, -1, -1, L"ImageList_Remove", teApiImageList_Remove },
	{ 3, -1, -1, -1, L"ImageList_SetIconSize", teApiImageList_SetIconSize },
	{ 6, -1, -1, -1, L"ImageList_Draw", teApiImageList_Draw },
	{10, -1, -1, -1, L"ImageList_DrawEx", teApiImageList_DrawEx },
	{ 2, -1, -1, -1, L"ImageList_SetImageCount", teApiImageList_SetImageCount },
	{ 3, -1, -1, -1, L"ImageList_SetOverlayImage", teApiImageList_SetOverlayImage },
	{ 5, -1, -1, -1, L"ImageList_Copy", teApiImageList_Copy },
	{ 1, -1, -1, -1, L"DestroyWindow", teApiDestroyWindow },
	{ 3, -1, -1, -1, L"LineTo", teApiLineTo },
	{ 0, -1, -1, -1, L"ReleaseCapture", teApiReleaseCapture },
	{ 2, -1, -1, -1, L"SetCursorPos", teApiSetCursorPos },
	{ 2, -1, -1, -1, L"DestroyCursor", teApiDestroyCursor },
	{ 2, -1, -1, -1, L"SHFreeShared", teApiSHFreeShared },
	{ 0, -1, -1, -1, L"EndMenu", teApiEndMenu },
	{ 1, -1, -1, -1, L"SHChangeNotifyDeregister", teApiSHChangeNotifyDeregister },
	{ 1, -1, -1, -1, L"SHChangeNotification_Unlock", teApiSHChangeNotification_Unlock },
	{ 1, -1, -1, -1, L"IsWow64Process", teApiIsWow64Process },
	{ 9, -1, -1, -1, L"BitBlt", teApiBitBlt },
	{ 2, -1, -1, -1, L"ImmSetOpenStatus", teApiImmSetOpenStatus },
	{ 2, -1, -1, -1, L"ImmReleaseContext", teApiImmReleaseContext },
	{ 2, -1, -1, -1, L"IsChild", teApiIsChild },
	{ 2, -1, -1, -1, L"KillTimer", teApiKillTimer },
	{ 1, -1, -1, -1, L"AllowSetForegroundWindow", teApiAllowSetForegroundWindow },
	{ 7, -1, -1, -1, L"SetWindowPos", teApiSetWindowPos },
	{ 5,  4, -1, -1, L"InsertMenu", teApiInsertMenu },
	{ 2, -1,  1, -1, L"SetWindowText", teApiSetWindowText },
	{ 4, -1, -1, -1, L"RedrawWindow", teApiRedrawWindow },
	{ 4, -1, -1, -1, L"MoveToEx", teApiMoveToEx },
	{ 3, -1, -1, -1, L"InvalidateRect", teApiInvalidateRect },
	{ 4, -1, -1, -1, L"SendNotifyMessage", teApiSendNotifyMessage },
	{ 5, -1, -1, -1, L"PeekMessage", teApiPeekMessage },
	{ 1, -1, -1, -1, L"SHFileOperation", teApiSHFileOperation },
	{ 4, -1, -1, -1, L"GetMessage", teApiGetMessage },
	{ 2, -1, -1, -1, L"GetWindowRect", teApiGetWindowRect },
	{ 2, -1, -1, -1, L"GetWindowThreadProcessId", teApiGetWindowThreadProcessId },
	{ 2, -1, -1, -1, L"GetClientRect", teApiGetClientRect },
	{ 3, -1, -1, -1, L"SendInput", teApiSendInput },
	{ 2, -1, -1, -1, L"ScreenToClient", teApiScreenToClient },
	{ 5, -1, -1, -1, L"MsgWaitForMultipleObjectsEx", teApiMsgWaitForMultipleObjectsEx },
	{ 2, -1, -1, -1, L"ClientToScreen", teApiClientToScreen },
	{ 2, -1, -1, -1, L"GetIconInfo", teApiGetIconInfo },
	{ 2, -1, -1, -1, L"FindNextFile", teApiFindNextFile },
	{ 3, -1, -1, -1, L"FillRect", teApiFillRect },
	{ 2, -1, -1, -1, L"Shell_NotifyIcon", teApiShell_NotifyIcon },
	{ 2, -1, -1, -1, L"EndPaint", teApiEndPaint },
	{ 2, -1, -1, -1, L"ImageList_GetIconSize", teApiImageList_GetIconSize },
	{ 2, -1, -1, -1, L"GetMenuInfo", teApiGetMenuInfo },
	{ 2, -1, -1, -1, L"SetMenuInfo", teApiSetMenuInfo },
	{ 4, -1, -1, -1, L"SystemParametersInfo", teApiSystemParametersInfo },
	{ 3,  1, -1, -1, L"GetTextExtentPoint32", teApiGetTextExtentPoint32 },
	{ 4, -1, -1, -1, L"SHGetDataFromIDList", teApiSHGetDataFromIDList },
	{ 4, -1, -1, -1, L"InsertMenuItem", teApiInsertMenuItem },
	{ 4, -1, -1, -1, L"GetMenuItemInfo", teApiGetMenuItemInfo },
	{ 4, -1, -1, -1, L"SetMenuItemInfo", teApiSetMenuItemInfo },
	{ 4, -1, -1, -1, L"ChangeWindowMessageFilterEx", teApiChangeWindowMessageFilterEx },
	{ 2, -1, -1, -1, L"ImageList_SetBkColor", teApiImageList_SetBkColor },
	{ 2, -1, -1, -1, L"ImageList_AddIcon", teApiImageList_AddIcon },
	{ 3, -1, -1, -1, L"ImageList_Add", teApiImageList_Add },
	{ 3, -1, -1, -1, L"ImageList_AddMasked", teApiImageList_AddMasked },
	{ 1, -1, -1, -1, L"GetKeyState", teApiGetKeyState },
	{ 1, -1, -1, -1, L"GetSystemMetrics", teApiGetSystemMetrics },
	{ 1, -1, -1, -1, L"GetSysColor", teApiGetSysColor },
	{ 1, -1, -1, -1, L"GetMenuItemCount", teApiGetMenuItemCount },
	{ 1, -1, -1, -1, L"ImageList_GetImageCount", teApiImageList_GetImageCount },
	{ 2, -1, -1, -1, L"ReleaseDC", teApiReleaseDC },
	{ 2, -1, -1, -1, L"GetMenuItemID", teApiGetMenuItemID },
	{ 4, -1, -1, -1, L"ImageList_Replace", teApiImageList_Replace },
	{ 3, -1, -1, -1, L"ImageList_ReplaceIcon", teApiImageList_ReplaceIcon },
	{ 2, -1, -1, -1, L"SetBkMode", teApiSetBkMode },
	{ 2, -1, -1, -1, L"SetBkColor", teApiSetBkColor },
	{ 2, -1, -1, -1, L"SetTextColor", teApiSetTextColor },
	{ 2, -1, -1, -1, L"MapVirtualKey", teApiMapVirtualKey },
	{ 2, -1, -1, -1, L"WaitForInputIdle", teApiWaitForInputIdle },
	{ 2, -1, -1, -1, L"WaitForSingleObject", teApiWaitForSingleObject },
	{ 3, -1, -1, -1, L"GetMenuDefaultItem", teApiGetMenuDefaultItem },
	{ 3, -1, -1, -1, L"SetMenuDefaultItem", teApiSetMenuDefaultItem },
	{ 1, -1, -1, -1, L"CRC32", teApiCRC32 },
	{ 3,  1, -1, -1, L"SHEmptyRecycleBin", teApiSHEmptyRecycleBin },
	{ 0, -1, -1, -1, L"GetMessagePos", teApiGetMessagePos },
	{ 2, -1, -1, -1, L"ImageList_GetOverlayImage", teApiImageList_GetOverlayImage },
	{ 1, -1, -1, -1, L"ImageList_GetBkColor", teApiImageList_GetBkColor },
	{ 3,  1,  2, -1, L"SetWindowTheme", teApiSetWindowTheme },
	{ 1, -1, -1, -1, L"ImmGetVirtualKey", teApiImmGetVirtualKey },
	{ 1, -1, -1, -1, L"GetAsyncKeyState", teApiGetAsyncKeyState },
	{ 6, -1, -1, -1, L"TrackPopupMenuEx", teApiTrackPopupMenuEx },
	{ 5,  0, -1, -1, L"ExtractIconEx", teApiExtractIconEx },
	{ 3, -1, -1, -1, L"GetObject", teApiGetObject },
	{ 6, -1, -1, -1, L"MultiByteToWideChar", teApiMultiByteToWideChar },
	{ 8,  2, -1, -1, L"WideCharToMultiByte", teApiWideCharToMultiByte },
	{ 2, -1, -1, -1, L"GetAttributesOf", teApiGetAttributesOf },
//	{ 2, -1, -1, -1, L"DoDragDrop", teApiDoDragDrop },
	{ 4, -1, -1, -1, L"SHDoDragDrop", teApiSHDoDragDrop },
	{ 3, -1, -1, -1, L"CompareIDs", teApiCompareIDs },
	{ 2,  0,  1, -1, L"ExecScript", teApiExecScript },
	{ 2,  0,  1, -1, L"GetScriptDispatch", teApiGetScriptDispatch },
	{ 2,  1, -1, -1, L"GetDispatch", teApiGetDispatch },
	{ 6, -1, -1, -1, L"SHChangeNotifyRegister", teApiSHChangeNotifyRegister },
	{ 4,  1,  2, -1, L"MessageBox", teApiMessageBox },
	//Handle
	{ 3, -1, -1, -1, L"ImageList_GetIcon", teApiImageList_GetIcon },
	{ 5, -1, -1, -1, L"ImageList_Create", teApiImageList_Create },
	{ 2, -1, -1, -1, L"GetWindowLongPtr", teApiGetWindowLongPtr },
	{ 2, -1, -1, -1, L"GetClassLongPtr", teApiGetClassLongPtr },
	{ 2, -1, -1, -1, L"GetSubMenu", teApiGetSubMenu },
	{ 2, -1, -1, -1, L"SelectObject", teApiSelectObject },
	{ 1, -1, -1, -1, L"GetStockObject", teApiGetStockObject },
	{ 1, -1, -1, -1, L"GetSysColorBrush", teApiGetSysColorBrush },
	{ 1, -1, -1, -1, L"SetFocus", teApiSetFocus },
	{ 1, -1, -1, -1, L"GetDC", teApiGetDC },
	{ 1, -1, -1, -1, L"CreateCompatibleDC", teApiCreateCompatibleDC },
	{ 0, -1, -1, -1, L"CreatePopupMenu", teApiCreatePopupMenu },
	{ 0, -1, -1, -1, L"CreateMenu", teApiCreateMenu },
	{ 3, -1, -1, -1, L"CreateCompatibleBitmap", teApiCreateCompatibleBitmap },
	{ 3, -1, -1, -1, L"SetWindowLongPtr", teApiSetWindowLongPtr },
	{ 3, -1, -1, -1, L"SetClassLongPtr", teApiSetClassLongPtr },
	{ 1, -1, -1, -1, L"ImageList_Duplicate", teApiImageList_Duplicate },
	{ 2, -1, -1, -1, L"SendMessage", teApiSendMessage },
	{ 2, -1, -1, -1, L"GetSystemMenu", teApiGetSystemMenu },
	{ 1, -1, -1, -1, L"GetWindowDC", teApiGetWindowDC },
	{ 3, -1, -1, -1, L"CreatePen", teApiCreatePen },
	{ 1, -1, -1, -1, L"SetCapture", teApiSetCapture },
	{ 1, -1, -1, -1, L"SetCursor", teApiSetCursor },
	{ 5, -1, -1, -1, L"CallWindowProc", teApiCallWindowProc },
	{ 1, -1, -1, -1, L"GetWindow", teApiGetWindow },
	{ 1, -1, -1, -1, L"GetTopWindow", teApiGetTopWindow },
	{ 3, -1, -1, -1, L"OpenProcess", teApiOpenProcess },
	{ 1, -1, -1, -1, L"GetParent", teApiGetParent },
	{ 0, -1, -1, -1, L"GetCapture", teApiGetCapture },
	{ 1, -1,  0, -1, L"GetModuleHandle", teApiGetModuleHandle },
	{ 1, -1, -1, -1, L"SHGetImageList", teApiSHGetImageList },
	{ 5, -1, -1, -1, L"CopyImage", teApiCopyImage },
	{ 0, -1, -1, -1, L"GetCurrentProcess", teApiGetCurrentProcess },
	{ 1, -1, -1, -1, L"ImmGetContext", teApiImmGetContext },
	{ 0, -1, -1, -1, L"GetFocus", teApiGetFocus },
	{ 0, -1, -1, -1, L"GetForegroundWindow", teApiGetForegroundWindow },
	{ 3, -1, -1, -1, L"SetTimer", teApiSetTimer },
	{ 2, -1, -1, -1, L"LoadMenu", teApiLoadMenu },
	{ 2, -1, -1, -1, L"LoadIcon", teApiLoadIcon },
	{ 3, -1, -1, -1, L"LoadLibraryEx", teApiLoadLibraryEx },
	{ 6, -1, -1, -1, L"LoadImage", teApiLoadImage },
	{ 7, -1, -1, -1, L"ImageList_LoadImage", teApiImageList_LoadImage },
	{ 5,  0, -1, -1, L"SHGetFileInfo", teApiSHGetFileInfo },
	{12, -1,  1,  2, L"CreateWindowEx", teApiCreateWindowEx },
	{ 6, -1,  3,  4, L"ShellExecute", teApiShellExecute },
	{ 2, -1, -1, -1, L"BeginPaint", teApiBeginPaint },
	{ 2, -1, -1, -1, L"LoadCursor", teApiLoadCursor },
	{ 2, -1, -1, -1, L"LoadCursorFromFile", teApiLoadCursorFromFile },
	{ 3, -1, -1, -1, L"SHParseDisplayName", teApiSHParseDisplayName },
	{ 3, -1, -1, -1, L"SHChangeNotification_Lock", teApiSHChangeNotification_Lock },
	{ 1, -1, -1, -1, L"GetWindowText", teApiGetWindowText },
	{ 1, -1, -1, -1, L"GetClassName", teApiGetClassName },
	{ 1, -1, -1, -1, L"GetModuleFileName", teApiGetModuleFileName },
	{ 0, -1, -1, -1, L"GetCommandLine", teApiGetCommandLine },
	{ 0, -1, -1, -1, L"GetCurrentDirectory", teApiGetCurrentDirectory },//use wsh.CurrentDirectory
	{ 3, -1, -1, -1, L"GetMenuString", teApiGetMenuString },
	{ 2, -1, -1, -1, L"GetDisplayNameOf", teApiGetDisplayNameOf },
	{ 1, -1, -1, -1, L"GetKeyNameText", teApiGetKeyNameText },
	{ 2, -1, -1, -1, L"DragQueryFile", teApiDragQueryFile },
	{ 1, -1, -1, -1, L"SysAllocString", teApiSysAllocString },
	{ 2, -1, -1, -1, L"SysAllocStringLen", teApiSysAllocStringLen },
	{ 2, -1, -1, -1, L"SysAllocStringByteLen", teApiSysAllocStringByteLen },
	{ 3, -1,  1, -1, L"sprintf", teApisprintf },
	{ 1, -1, -1, -1, L"base64_encode", teApibase64_encode },
	{ 2, -1, -1, -1, L"LoadString", teApiLoadString },
	{ 4,  2,  3, -1, L"AssocQueryString", teApiAssocQueryString },
	{ 1, -1, -1, -1, L"StrFormatByteSize", teApiStrFormatByteSize },
	{ 4,  3, -1, -1, L"GetDateFormat", teApiGetDateFormat },
	{ 4,  3, -1, -1, L"GetTimeFormat", teApiGetTimeFormat },
	{ 2, -1, -1, -1, L"GetLocaleInfo", teApiGetLocaleInfo },	
	{ 2, -1, -1, -1, L"MonitorFromPoint", teApiMonitorFromPoint },
	{ 2, -1, -1, -1, L"MonitorFromRect", teApiMonitorFromRect },
	{ 2, -1, -1, -1, L"GetMonitorInfo", teApiGetMonitorInfo },
	{ 1,  0, -1, -1, L"GlobalAddAtom", teApiGlobalAddAtom },
	{ 1, -1, -1, -1, L"GlobalGetAtomName", teApiGlobalGetAtomName },
	{ 1,  0, -1, -1, L"GlobalFindAtom", teApiGlobalFindAtom },
	{ 1, -1, -1, -1, L"GlobalDeleteAtom", teApiGlobalDeleteAtom },
	{ 4, -1, -1, -1, L"RegisterHotKey", teApiRegisterHotKey },
	{ 2, -1, -1, -1, L"UnregisterHotKey", teApiUnregisterHotKey },
	{ 1, -1, -1, -1, L"ILGetCount", teApiILGetCount },
	{ 2, -1, -1, -1, L"SHTestTokenMembership", teApiSHTestTokenMembership },
	{ 2,  1, -1, -1, L"ObjGetI", teApiObjGetI },
	{ 3,  1, -1, -1, L"ObjPutI", teApiObjPutI },
	{ 5,  1, -1, -1, L"DrawText", teApiDrawText },
	{ 5, -1, -1, -1, L"Rectangle", teApiRectangle },
	{ 1,  0, -1, -1, L"PathIsDirectory", teApiPathIsDirectory },
	{ 2,  0,  1, -1, L"PathIsSameRoot", teApiPathIsSameRoot },
	{ 2,  1, -1, -1, L"ObjDispId", teApiObjDispId },
	{ 1,  0, -1, -1, L"SHSimpleIDListFromPath", teApiSHSimpleIDListFromPath },
	{ 1,  0, -1, -1, L"OutputDebugString", teApiOutputDebugString },
	{ 2, -1, -1, -1, L"DllGetClassObject", teApiDllGetClassObject },
//	{ 0, -1, -1, -1, L"", teApi },
//	{ 0, -1, -1, -1, L"Test", teApiTest },
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
	for (int i = 0; i < _countof(pTEStructs); i++) {
		int j = map[i];
		if (i != j) {
			LPWSTR lpi = pTEStructs[i].name;
			LPWSTR lpj = pTEStructs[j].name;
			Sleep(0);
		}
	}
*///
#ifdef _USE_BSEARCHAPI
	g_maps[MAP_API] = teSortDispatchApi(dispAPI, _countof(dispAPI));
#endif
#ifndef _2000XP
	lpfnSHParseDisplayName(L"shell:::{2965E715-EB66-4719-B53F-1672673BBEFA}", NULL, &g_pidlResultsFolder, 0, NULL);
#else
	DWORDLONG dwlConditionMask = 0;
	OSVERSIONINFOEX osvi = { sizeof(OSVERSIONINFOEX), 6 };
    VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, VER_GREATER_EQUAL);
	g_bUpperVista = VerifyVersionInfo(&osvi, VER_MAJORVERSION, dwlConditionMask);
    dwlConditionMask = 0;
	osvi.dwMajorVersion = 5;
    osvi.dwMinorVersion = 1;
	VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, VER_EQUAL);
    VER_SET_CONDITION(dwlConditionMask, VER_MINORVERSION, VER_GREATER_EQUAL);
	g_bIsXP = VerifyVersionInfo(&osvi, VER_MAJORVERSION | VER_MINORVERSION, dwlConditionMask);
#endif
#ifdef _W2000
	g_bIs2000 = !g_bUpperVista && !g_bIsXP;
#endif
#ifdef _2000XP
	lpfnSHParseDisplayName(L"::{e17d4fc0-5564-11d1-83f2-00a0c90dc849}", NULL, &g_pidlResultsFolder, 0, NULL);
	if (!g_pidlResultsFolder) {
		lpfnSHParseDisplayName(L"shell:::{2965E715-EB66-4719-B53F-1672673BBEFA}", NULL, &g_pidlResultsFolder, 0, NULL);
	}
	g_bCharWidth = !g_bUpperVista;
#endif
	lpfnSHParseDisplayName(L"shell:libraries", NULL, &g_pidlLibrary, 0, NULL);
	g_pidlCP = ILClone(g_pidls[CSIDL_CONTROLS]);
	ILRemoveLastID(g_pidlCP);
	if (ILIsEmpty(g_pidlCP) || ILIsEqual(g_pidlCP, g_pidls[CSIDL_DRIVES])) {
		teILCloneReplace(&g_pidlCP, g_pidls[CSIDL_CONTROLS]);
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
				if (tePathMatchSpec(pszName, pImageCodecInfo[num].FilenameExtension) || lstrcmpi(pszName, pImageCodecInfo[num].MimeType) == 0) {
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
	} else {
		rc->left = param[TE_Left] / 0x4000 + (param[TE_Left] >= 0 ? rcClient->left : rcClient->right);
	}
	if (param[TE_Top] & 0x3fff) {
		rc->top = (param[TE_Top] * (rcClient->bottom - rcClient->top)) / 10000 + rcClient->top;
	} else {
		rc->top = param[TE_Top] / 0x4000 + (param[TE_Top] >= 0 ? rcClient->top : rcClient->bottom);
	}
	if (param[TE_Width] & 0x3fff) {
		rc->right = param[TE_Width] * (rcClient->right - rcClient->left) / 10000 + rc->left;
	} else {
		rc->right = param[TE_Width] / 0x4000 + (param[TE_Width] >= 0 ? rc->left : rcClient->right);
	}

	if (rc->right > rcClient->right) {
		rc->right = rcClient->right;
	}
	if (param[TE_Height] & 0x3fff) {
		rc->bottom = param[TE_Height] * (rcClient->bottom - rcClient->top) / 10000 + rc->top;
	} else {
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
	RECT rc, rcTab, rcClient;
	IDispatch *pdisp;
	KillTimer(hwnd, idEvent);
	try {
		switch (idEvent) {
			case TET_Create:
				if (g_pOnFunc[TE_OnCreate]) {
					DoFunc(TE_OnCreate, g_pTE, E_NOTIMPL);
					DoFunc(TE_OnCreate, g_pWebBrowser, E_NOTIMPL);
					break;
				}
				g_bsDocumentWrite = ::SysAllocStringLen(NULL, MAX_STATUS);
				lstrcpy(g_bsDocumentWrite, lstrcmpi(g_pWebBrowser->m_bstrPath, L"TE32.exe") ? PathFileExists(g_pWebBrowser->m_bstrPath) ? L"<h1>500 Internal Script Error</h1>" : L"<h1>404 File Not Found</h1>" : L"<h1>303 Exec Other</h1>");
				lstrcat(g_bsDocumentWrite, g_pWebBrowser->m_bstrPath);
				g_nSize = 0;
				g_nLockUpdate = 0;
				::SysReAllocString(&g_pWebBrowser->m_bstrPath, L"about:blank");
				g_pWebBrowser->m_pWebBrowser->Navigate(g_pWebBrowser->m_bstrPath, NULL, NULL, NULL, NULL);
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
			case TET_Size:
				GetClientRect(hwnd, &rcClient);
				rcClient.left += g_param[TE_Left];
				rcClient.top += g_param[TE_Top];
				rcClient.right -= g_param[TE_Width];
				rcClient.bottom -= g_param[TE_Height];
				CteShellBrowser *pSB;
				BOOL bRedraw, bDirect;
				bRedraw = FALSE;
				bDirect = TRUE;
				CteTabCtrl *pTC;
				g_nSize--;
				for (int i = MAX_TC; i-- && (pTC = g_ppTC[i]);) {
					if (pTC->m_bVisible) {
						pTC->LockUpdate();
						try {
							teCalcClientRect(pTC->m_param, &rc, &rcClient);
							if (g_pOnFunc[TE_OnArrange]) {
								VARIANTARG *pv;
								pv = GetNewVARIANT(2);
								teSetObject(&pv[1], pTC);
								teSetObjectRelease(&pv[0], new CteMemory(sizeof(RECT), &rc, 1, L"RECT"));
								Invoke4(g_pOnFunc[TE_OnArrange], NULL, 2, pv);
							}
							if (!pTC->m_bEmpty && pTC->m_bVisible) {
								int nAlign = g_param[TE_Tab] ? pTC->m_param[TC_Align] : 0;
								if (rc.left) {
									rc.left++;
								}
								if (teSetRect(pTC->m_hwndStatic, rc.left, rc.top, rc.right, rc.bottom)) {
									bRedraw = TRUE;
								}
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
									} else {
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
												pSB->Show(FALSE, 1);
											}
										}
									}
								}
								MoveWindow(pTC->m_hwndButton, rcTab.left, rcTab.top,
									rcTab.right - rcTab.left, rcTab.bottom - rcTab.top, FALSE);
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
									if (pTC->m_param[TE_Flags] & TCS_FIXEDWIDTH) {
										pTC->SetItemSize();
									}
									rcTab.right -= pTC->m_nScrollWidth;
								} else {
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
										teSetParent(pSB->m_pTV->m_hwnd, pSB->m_pTC->m_hwndStatic);
										ShowWindow(pSB->m_pTV->m_hwnd, (pSB->m_param[SB_TreeAlign] & 2) ? SW_SHOW : SW_HIDE);
									} else {
										if (!pSB->m_hwnd) {
											bDirect = FALSE;
										}
										pSB->Show(TRUE, 0);
									}
									MoveWindow(pSB->m_hwnd, rc.left, rc.top, rc.right - rc.left,
										rc.bottom - rc.top, FALSE);
								}
							}
						} catch (...) {}
						pTC->UnLockUpdate(bDirect);
					}
				}
				if (g_pWebBrowser) {
					IOleInPlaceObject *pOleInPlaceObject;
					g_pWebBrowser->m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleInPlaceObject));
					GetClientRect(hwnd, &rc);
					pOleInPlaceObject->SetObjectRects(&rc, &rc);
					pOleInPlaceObject->Release();
					if (bRedraw) {
						RedrawWindow(g_pWebBrowser->m_hwndBrowser, NULL, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
					}
				}
				break;
			case TET_Redraw:
				for (int i = MAX_TC; i-- && (pTC = g_ppTC[i]);) {
					pTC->RedrawUpdate();
				}
				break;
			case TET_Unload:
				g_bUnload = FALSE;
				if (g_pUnload) {
					pdisp = g_pUnload;
					GetNewArray(&g_pUnload);
					pdisp->Release();
					teExecMethod(g_pJS, L"CollectGarbage", NULL, 0, NULL);
				}
				break;
			case TET_FreeLibrary:
				teFreeLibraries();
				break;
			case TET_Status:
				pSB = g_pTC->GetShellBrowser(g_pTC->m_nIndex);
				if (pSB) {
					if (g_szStatus[0] == NULL) {
						IFolderView *pFV;
						if (pSB->m_pShellView && SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV)))) {
							int nCount;
							if SUCCEEDED(pFV->ItemCount(SVGIO_SELECTION, &nCount)) {
								UINT uID;
								if (nCount) {
									uID = nCount > 1 ? 38194 : 38195;
								} else if SUCCEEDED(pFV->ItemCount(SVGIO_ALLVIEW, &nCount)) {
									uID = nCount > 1 ? 38192 : 38193;
								}
								BSTR bsCount = ::SysAllocStringLen(NULL, 16);
								WCHAR pszNum[16];
								swprintf_s(pszNum, 12, L"%d", nCount);
								teCommaSize(pszNum, bsCount, 16, 0);

								WCHAR psz[MAX_STATUS];
								if (LoadString(g_hShell32, uID, psz, MAX_STATUS) > 2 && tePathMatchSpec1(psz, L"*%s*")) {
									swprintf_s(g_szStatus, MAX_STATUS, psz, bsCount);
								} else {
									uID = uID < 38194 ? 6466 : 6477;
									if (LoadString(g_hShell32, uID, psz, MAX_STATUS) > 6) {
										FormatMessage(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ARGUMENT_ARRAY,
											psz, 0, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), g_szStatus, MAX_STATUS, (va_list *)&bsCount);
									}
								}
								::SysFreeString(bsCount);
							}
							pFV->Release();
						}
					}
					DoStatusText(pSB, g_szStatus, 0);
					g_szStatus[0] = NULL;
				}
				break;
			case TET_Title:
				SetWindowText(hwnd, g_szTitle);
				break;
		}//end_switch
	} catch (...) {
		g_nException = 0;
	}
}

VOID CALLBACK teTimerProcSW(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	IDispatch *pdisp;
	try {
		KillTimer(hwnd, idEvent);
		CteWebBrowser *pWB = (CteWebBrowser *)idEvent;
		if (GetFocus() == pWB->m_hwndParent) {
			if (pWB->m_pWebBrowser->get_Document(&pdisp) == S_OK) {
				IHTMLDocument2 *pdoc;
				if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pdoc))) {
					IHTMLWindow2 *pwin;
					if SUCCEEDED(pdoc->get_parentWindow(&pwin)) {
						teExecMethod(pwin, L"focus", NULL, NULL, NULL);
						pwin->Release();
					}
					pdoc->Release();
				}
				pdisp->Release();
			}
		}
	} catch (...) {
		g_nException = 0;
	}
}

VOID CALLBACK teTimerProcSort(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);
	try {
		SORTCOLUMN *pCol = (SORTCOLUMN *)idEvent;
		CteShellBrowser *pSB = SBfromhwnd(hwnd);
		if (pSB) {
			IFolderView2 *pFV2;
			if SUCCEEDED(pSB->m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
				pFV2->SetSortColumns(pCol, 1);
				pFV2->Release();
			}
		}
		delete [] pCol;
	} catch (...) {
		g_nException = 0;
	}
}

VOID CALLBACK teTimerProcFocus(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	try {
		KillTimer(hwnd, idEvent);
		CteShellBrowser *pSB = SBfromhwnd(hwnd);
		if (pSB) {
			pSB->FocusItem(TRUE);
		}
	} catch (...) {
		g_nException = 0;
	}
}

LRESULT CALLBACK HookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LPCWPSTRUCT pcwp;
	CteShellBrowser *pSB;
	try {
		if (nCode >= 0) {
			if (nCode == HC_ACTION) {
				if (wParam == NULL) {
					pcwp = (LPCWPSTRUCT)lParam;
					switch (pcwp->message)
					{
						case WM_KILLFOCUS:
							if (pcwp->message == WM_KILLFOCUS && GetFocus() != g_hwndMain) {
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
									MessageSub(TE_OnSystemMessage, pdisp, &msg1, S_FALSE);
								}
							}
							break;
						case WM_SETFOCUS:
							CheckChangeTabSB(pcwp->hwnd);
							break;
						case WM_COMMAND:
							if (pcwp->message == WM_COMMAND && g_hDialog == pcwp->hwnd) {
								if (LOWORD(pcwp->wParam) == IDOK) {
									g_bDialogOk = TRUE;
								}
							}
							break;
						case WM_HSCROLL:
						case WM_VSCROLL:
							g_dwTickScroll = GetTickCount();
							break;
						case LVM_SETTEXTCOLOR:
							pSB = SBfromhwnd(pcwp->hwnd);
							if (pSB) {
								if (!pSB->m_bRefreshing) {
									pSB->m_clrText = (COLORREF)pcwp->lParam;
								}
							}
							break;
						case LVM_SETBKCOLOR:
							pSB = SBfromhwnd(pcwp->hwnd);
							if (pSB) {
								if (!pSB->m_bRefreshing) {
									pSB->m_clrBk = (COLORREF)pcwp->lParam;
								}
							}
							break;
						case LVM_SETTEXTBKCOLOR:
							pSB = SBfromhwnd(pcwp->hwnd);
							if (pSB) {
								if (!pSB->m_bRefreshing) {
									pSB->m_clrTextBk = (COLORREF)pcwp->lParam;
								}
							}
							break;
					}//end_switch
				}
			}
		}
	} catch (...) {
		g_nException = 0;
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
	} catch (...) {
		g_nException = 0;
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
	g_dwMainThreadId = GetCurrentThreadId();
	teGetModuleFileName(NULL, &bsPath);//Executable Path
	for (int i = 0; bsPath[i]; i++) {
		if (bsPath[i] == '\\') {
			bsPath[i] = '/';
		}
		bsPath[i] = towupper(bsPath[i]);
	}
	int nCrc32 = CalcCrc32((BYTE *)bsPath, lstrlen(bsPath) * sizeof(WCHAR), 0);
	//Command Line
	BOOL bVisible = !teStartsText(L"/run ", lpCmdLine);
	BOOL bNewProcess = teStartsText(L"/open ", lpCmdLine);
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
		g_bsPidls[i] = NULL;
		if (g_pidls[i]) {
			IShellFolder *pSF;
			LPCITEMIDLIST pidlPart;
			if SUCCEEDED(SHBindToParent(g_pidls[i], IID_PPV_ARGS(&pSF), &pidlPart)) {
				teGetDisplayNameBSTR(pSF, pidlPart, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &g_bsPidls[i]);
				pSF->Release();
			}
		} else if (g_nPidls == i + 1) {
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
		lpfnSHTestTokenMembership = (LPFNSHTestTokenMembership)GetProcAddress(g_hShell32, "SHTestTokenMembership");
	}
#ifdef _2000XP
	if (!lpfnSHGetIDListFromObject) {
		lpfnSHGetIDListFromObject = teGetIDListFromObjectXP;
	}
#endif
#ifdef _W2000
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
		lpfnPSStringFromPropertyKey = (LPFNPSStringFromPropertyKey)GetProcAddress(hDll, "PSStringFromPropertyKey");
		lpfnPSPropertyKeyFromStringEx = tePSPropertyKeyFromStringEx;
	}
#ifdef _2000XP
	else {
		lpfnPSPropertyKeyFromStringEx = tePSPropertyKeyFromStringXP;
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
		lpfnRemoveClipboardFormatListener = (LPFNRemoveClipboardFormatListener)GetProcAddress(hDll, "RemoveClipboardFormatListener");
		if (lpfnRemoveClipboardFormatListener) {
			lpfnAddClipboardFormatListener = (LPFNAddClipboardFormatListener)GetProcAddress(hDll, "AddClipboardFormatListener");
		}
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
// API Hook test
#ifdef _USE_APIHOOK
	tePathAppend(&bsLib, bsPath, L"advapi32.dll");
	hDll = GetModuleHandle(bsLib);
	if (hDll) {
		lpfnRegQueryValueExW = (LPFNRegQueryValueExW)GetProcAddress(hDll, "RegQueryValueExW");
	}
	SysFreeString(bsLib);

	tePathAppend(&bsLib, bsPath, L"comctl32.dll");
	hDll = GetModuleHandle(NULL);
	if (hDll) {
		DWORD_PTR dwBase = (DWORD_PTR)hDll;
		DWORD dwIdataSize, dwDummy;
		PIMAGE_IMPORT_DESCRIPTOR pImgDesc = (PIMAGE_IMPORT_DESCRIPTOR)ImageDirectoryEntryToData((HMODULE)dwBase,
			true, IMAGE_DIRECTORY_ENTRY_IMPORT, &dwIdataSize);
		while(pImgDesc->Name) {
			char *lpModule = (char*)(dwBase + pImgDesc->Name);
			if (!_strcmpi(lpModule, "advapi32.dll")) {
				PIMAGE_THUNK_DATA pIAT = (PIMAGE_THUNK_DATA)(dwBase + pImgDesc->FirstThunk);
				while (pIAT->u1.Function) {
					if ((LPFNRegQueryValueExW)pIAT->u1.Function == lpfnRegQueryValueExW) {
						VirtualProtect(&pIAT->u1.Function, sizeof(pIAT->u1.Function), PAGE_EXECUTE_READWRITE, &dwDummy);
						LPVOID NewProc = &teRegQueryValueExW;
						WriteProcessMemory(GetCurrentProcess(), &pIAT->u1.Function, &NewProc, sizeof(pIAT->u1.Function), &dwDummy);
						break;
					}
					pIAT++;
				}
				break;
			}
			pImgDesc++;
		}
	}
	SysFreeString(bsLib);
#endif
	SysFreeString(bsPath);

	Initlize();
	MSG msg;

	//Initialize FolderView & TreeView settings
	g_paramFV[SB_Type] = 1;
	g_paramFV[SB_ViewMode] = FVM_DETAILS;
	g_paramFV[SB_FolderFlags] = FWF_SHOWSELALWAYS | FWF_AUTOARRANGE;
	g_paramFV[SB_IconSize] = 0;
	g_paramFV[SB_Options] = EBO_ALWAYSNAVIGATE;
	g_paramFV[SB_ViewFlags] = 0;
	g_paramFV[SB_SizeFormat] = 0;

	g_paramFV[SB_TreeAlign] = 1;
	g_paramFV[SB_TreeWidth] = 200;
	g_paramFV[SB_TreeFlags] = NSTCS_HASEXPANDOS | NSTCS_SHOWSELECTIONALWAYS | NSTCS_BORDER | NSTCS_HASLINES | NSTCS_NOINFOTIP;
	g_paramFV[SB_EnumFlags] = SHCONTF_FOLDERS;
	g_paramFV[SB_RootStyle] = NSTCRS_VISIBLE | NSTCRS_EXPANDED;

	// Initialize GDI+
	Gdiplus::GdiplusStartup(&g_Token, &g_StartupInput, NULL);
	MyRegisterClass(hInstance, szClass, WndProc);
	// Title & Version
	lstrcpy(g_szTE, _T(PRODUCTNAME) L" " _T(STRING(VER_Y)) L"." _T(STRING(VER_M)) L"." _T(STRING(VER_D)) L" Gaku");
	lstrcpy(g_szTitle, g_szTE);
	if (bVisible) {
		g_bInit = !bNewProcess;
		g_hwndMain = CreateWindow(szClass, g_szTE, WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		  CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInstance, NULL);
		if (!bNewProcess) {
			SetWindowLongPtr(g_hwndMain, GWLP_USERDATA, nCrc32);
			ShowWindow(g_hwndMain, SW_SHOWMINIMIZED);
		}
	} else {
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
	g_hHook = SetWindowsHookEx(WH_CALLWNDPROC, (HOOKPROC)HookProc, hInst, g_dwMainThreadId);
	g_hMouseHook = SetWindowsHookEx(WH_MOUSE, (HOOKPROC)MouseProc, hInst, g_dwMainThreadId);
	// Create own class
	CoCreateGuid(&g_ClsIdSB);
	CoCreateGuid(&g_ClsIdTC);
	CoCreateGuid(&g_ClsIdStruct);
	CoCreateGuid(&g_ClsIdFI);
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
	if (lpfnAddClipboardFormatListener) {
		lpfnAddClipboardFormatListener(g_hwndMain);
	}
#ifdef _2000XP
	else {
		g_hwndNextClip = SetClipboardViewer(g_hwndMain);
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
function _t(o) {\
	return {window: {external: o}};\
}\
";
	if (ParseScript(lp, L"JScript", NULL, NULL, &g_pJS) != S_OK) {
		PostMessage(g_hwndMain, WM_CLOSE, 0, 0);
		Finalize();
		return FALSE;
	}
	GetNewArray(&g_pUnload);
	GetNewArray(&g_pFreeLibrary);
	//WindowAPI
#ifdef _USE_BSEARCHAPI
	g_pAPI = new CteWindowsAPI(NULL);
#else
	GetNewObject(&g_pAPI);
	for (int i = _countof(dispAPI); i--;) {
		VARIANT v;
		v.vt = VT_DISPATCH;
		v.pdispVal = new CteAPI(&dispAPI[i]);
		tePutProperty(g_pAPI, dispAPI[i].name, &v);
		VariantClear(&v);
	}
#endif
	// CTE
	g_pTE = new CTE(nCmdShow);
	RegisterDragDrop(g_hwndMain, static_cast<IDropTarget *>(g_pTE));

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
		SysFreeString(bsPath);
		if (lplpszArgs) {
			LocalFree(lplpszArgs);
		}
#ifdef _WIN64
		if (!lpfnSHCreateItemFromIDList) {
			SysReAllocString(&bsIndex, L"TE32.exe");
		}
#endif
		CoInternetSetFeatureEnabled(FEATURE_DISABLE_NAVIGATION_SOUNDS, SET_FEATURE_ON_PROCESS, TRUE);
		g_pWebBrowser = new CteWebBrowser(g_hwndMain, bsIndex, NULL);
		SysFreeString(bsIndex);
		SetTimer(g_hwndMain, TET_Create, 10000, teTimerProc);
	} else {
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
			if (x && api.StrCmpI(x.FullName, arg[0]) == 0) {\
				var w = x.Document.parentWindow;\
				if (!window.MainWindow || window.MainWindow.Exchange && window.MainWindow.Exchange[arg[3]]) {\
					window.MainWindow = w;\
					var rc = api.Memory('RECT');\
					api.GetWindowRect(w.te.hwnd, rc);\
					api.MoveWindow(te.hwnd, (rc.Left + rc.Right) / 2, (rc.Top + rc.Bottom) / 2, 0, 0, false);\
				}\
			}\
		}\
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
	teGetWinSettings();

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
		}
	}
	//At the end of processing
	try {
		RevokeDragDrop(g_hwndMain);
#ifdef _USE_HTMLDOC
		teReleaseClear(&pHtmlDoc);
#else
		teReleaseClear(&pJS);
		VariantClear(&vWindow);
#endif
		teFreeLibraries();
		teReleaseClear(&g_pFreeLibrary);
		teReleaseClear(&g_pSubWindows);
		teReleaseClear(&g_pUnload);
		teReleaseClear(&g_pAPI);
		teReleaseClear(&g_pJS);
		teReleaseClear(&g_pAutomation);
		teReleaseClear(&g_pDropTargetHelper);
		UnhookWindowsHookEx(g_hMouseHook);
		UnhookWindowsHookEx(g_hHook);
		if (g_hMenu) {
			DestroyMenu(g_hMenu);
		}
		if (lpfnAddClipboardFormatListener) {
			lpfnRemoveClipboardFormatListener(g_hwndMain);
		}
#ifdef	_2000XP
		else {
			ChangeClipboardChain(g_hwndMain, g_hwndNextClip);
		}
#endif
	} catch (...) {}
	Finalize();
	teReleaseClear(&g_pTE);
	return (int) msg.wParam;
}

VOID ArrangeWindow()
{
	if (g_nSize <= 0) {
		g_nSize++;
		SetTimer(g_hwndMain, TET_Size, 1, teTimerProc);
	}
}

VOID CALLBACK teTimerProcRelease(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	try {
		KillTimer(hwnd, idEvent);
		((IUnknown *)idEvent)->Release();
	} catch (...) {
		g_nException = 0;
	}
}

VOID teDelayRelease(PVOID ppObj)
{
	try {
		IUnknown **ppunk = static_cast<IUnknown **>(ppObj);
		if (*ppunk) {
			if (*ppunk != g_pRelease) {
				g_pRelease = *ppunk;
				SetTimer(g_hwndMain, (UINT_PTR)*ppunk, 100, teTimerProcRelease);
			} else {
				(*ppunk)->Release();
			}
			*ppunk = NULL;
		}
	} catch (...) {
		g_nException = 0;
	}
}

BOOL teReleaseParse(TEInvoke *pInvoke)
{
	BOOL bResult = ::InterlockedDecrement(&pInvoke->cRef) == 0;
	if (bResult) {
		teDelayRelease(&pInvoke->pdisp);
	}
	return bResult;
}

VOID CALLBACK teTimerProcParse(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	KillTimer(hwnd, idEvent);
	TEInvoke *pInvoke = (TEInvoke *)idEvent;
	try {
		if (pInvoke->pidl || pInvoke->wErrorHandling == 3) {
			VariantClear(&pInvoke->pv[pInvoke->cArgs - 1]);
			if (pInvoke->wMode) {
				teSetLong(&pInvoke->pv[pInvoke->cArgs - 1], pInvoke->hr);
			} else {
				FolderItem *pid;
				if (pInvoke->pidl && GetFolderItemFromPidl(&pid, pInvoke->pidl)) {
					teSetObjectRelease(&pInvoke->pv[pInvoke->cArgs - 1], pid);
				}
			}
			teILFreeClear(&pInvoke->pidl);
			Invoke5(pInvoke->pdisp, pInvoke->dispid, DISPATCH_METHOD, NULL, pInvoke->cArgs, pInvoke->pv);
		} else if (pInvoke->wErrorHandling == 1) {
			CteFolderItem *pid = new CteFolderItem(&pInvoke->pv[pInvoke->cArgs - 1]);
			teILCloneReplace(&pid->m_pidl, g_pidlResultsFolder);
			pid->m_dwUnavailable = GetTickCount();
			VariantInit(&pInvoke->pv[pInvoke->cArgs - 1]);
			teSetObjectRelease(&pInvoke->pv[pInvoke->cArgs - 1], pid);
			Invoke5(pInvoke->pdisp, pInvoke->dispid, DISPATCH_METHOD, NULL, pInvoke->cArgs, pInvoke->pv);
		} else if (pInvoke->wErrorHandling == 2) {
			if (g_bShowParseError) {
				LPWSTR lpBuf;
				WCHAR pszFormat[MAX_STATUS];
				if (LoadString(g_hShell32, 6456, pszFormat, MAX_STATUS)) {
					if (FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ARGUMENT_ARRAY,
						pszFormat, 0, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&lpBuf, MAX_STATUS, (va_list *)&pInvoke->pv[pInvoke->cArgs - 1].bstrVal)) {
						int r = MessageBox(g_hwndMain, lpBuf, _T(PRODUCTNAME), MB_ABORTRETRYIGNORE);
						LocalFree(lpBuf);
						if (r == IDRETRY) {
							::InterlockedIncrement(&pInvoke->cRef);
							_beginthread(&threadParseDisplayName, 0, pInvoke);
							return;
						}
						if (r == IDIGNORE) {
							g_bShowParseError = FALSE;
						}
					}
				}
			}
			teClearVariantArgs(pInvoke->cArgs, pInvoke->pv);
		}
	} catch (...) {
		g_nException = 0;
	}
	teReleaseParse(pInvoke);
}

static void threadParseDisplayName(void *args)
{
	::OleInitialize(NULL);
	TEInvoke *pInvoke = (TEInvoke *)args;
	try {
		if (!pInvoke->pidl) {
			pInvoke->pidl = teILCreateFromPath1(pInvoke->pv[pInvoke->cArgs - 1].bstrVal);
		}
		if (pInvoke->wMode) {
			pInvoke->hr = E_PATH_NOT_FOUND;
			if (pInvoke->pidl) {
				pInvoke->hr = teILFolderExists(pInvoke->pidl);
			}
		}
	} catch (...) {
		g_nException = 0;
	}
	if (!teReleaseParse(pInvoke)) {
		SetTimer(g_hwndMain, (UINT_PTR)args, 10, teTimerProcParse);
	}
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
			try {
				if (CanClose(g_pTE)) {
					teRevoke();
					try {
						DestroyWindow(hWnd);
					} catch (...) {
						g_bMessageLoop = FALSE;
					}
				}
			} catch (...) {
				DestroyWindow(hWnd);
			}
			g_nLockUpdate -= 1000;
			return 0;
		case WM_SIZE:
			ArrangeWindow();
			break;
		case WM_DESTROY:
			try {
				if (g_pOnFunc[TE_OnSystemMessage]) {
					msg1.hwnd = hWnd;
					msg1.message = message;
					msg1.wParam = wParam;
					msg1.lParam = lParam;
					MessageSub(TE_OnSystemMessage, g_pTE, &msg1, S_FALSE);
				}
				//Close Tab
				CteTabCtrl *pTC;
				for (i = MAX_TC; i-- && (pTC = g_ppTC[i]);) {
					pTC->Close(TRUE);
				}
				//Close Browser
				if (g_pWebBrowser) {
					g_pWebBrowser->Close();
					teReleaseClear(&g_pWebBrowser);
				}
			} catch (...) {}
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
			if (message == WM_SETTINGCHANGE) {
				teGetWinSettings();
				if (lpfnRegenerateUserEnvironment) {
					try {
						if (lstrcmpi((LPCWSTR)lParam, L"Environment") == 0) {
							LPVOID lpEnvironment;
							lpfnRegenerateUserEnvironment(&lpEnvironment, TRUE);
							//Not permitted to free lpEnvironment!
							//FreeEnvironmentStrings((LPTSTR)lpEnvironment);
						}
					} catch (...) {}
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
		case WM_SYSCOMMAND:
		case WM_THEMECHANGED:
		case WM_USERCHANGED:
		case WM_QUERYENDSESSION:
		case WM_HOTKEY:
		case WM_CLIPBOARDUPDATE:
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
#ifdef	_2000XP
		case WM_DRAWCLIPBOARD:
			if (g_hwndNextClip) {
				SendMessage(g_hwndNextClip, message, wParam, lParam);
			}
			PostMessage(g_hwndMain, WM_CLIPBOARDUPDATE, 0, 0);
	        return 0;
		case WM_CHANGECBCHAIN:
			if ((HWND)wParam == g_hwndNextClip) {
				g_hwndNextClip = (HWND)lParam;
			} else if (g_hwndNextClip) {
				SendMessage(g_hwndNextClip, message, wParam, lParam);
			}
			return 0;
#endif
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

// Sub window

BOOL teGetSubWindow(HWND hwnd, CteWebBrowser **ppWB)
{
	VARIANT v;
	VariantInit(&v);
	BOOL bResult = SUCCEEDED(teGetPropertyAtLLX(g_pSubWindows, (LONGLONG)hwnd, &v));
	*ppWB = (CteWebBrowser *)v.pdispVal;
	VariantClear(&v);
	return bResult;
}
LRESULT CALLBACK WndProc2(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	try {
		CteWebBrowser *pWB;
		switch (message)
		{
			case WM_CLOSE:
				try {
					if (teGetSubWindow(hWnd, &pWB)) {
						if (pWB->m_bstrPath) {
							::SysReAllocString(&pWB->m_bstrPath, L"about:blank");
							pWB->m_pWebBrowser->Navigate(pWB->m_bstrPath, NULL, NULL, NULL, NULL);
							teSysFreeString(&pWB->m_bstrPath);
							return 0;
						}
						pWB->Close();
					}
					if (GetForegroundWindow() == pWB->m_hwndParent) {
						teSetForegroundWindow((HWND)GetWindowLongPtr(pWB->m_hwndParent, GWLP_HWNDPARENT));
					}
				} catch(...) {
				}
				DestroyWindow(hWnd);
				return 0;
			case WM_SIZE:
				if (teGetSubWindow(hWnd, &pWB)) {
					RECT rc;
					GetClientRect(hWnd, &rc);
					IOleInPlaceObject *pOleInPlaceObject;
					pWB->m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleInPlaceObject));
					pOleInPlaceObject->SetObjectRects(&rc, &rc);
					pOleInPlaceObject->Release();
				}
				break;
			case WM_ACTIVATE:
				if (LOWORD(wParam)) {
					if (teGetSubWindow(hWnd, &pWB)) {
						SetTimer(pWB->m_hwndParent, (UINT_PTR)pWB, 100, teTimerProcSW);
					}
				}
				break;
			default:
				return DefWindowProc(hWnd, message, wParam, lParam);
		}
	} catch (...) {
		g_nException = 0;
	}
	return 0;
}

// CteShellBrowser

CteShellBrowser::CteShellBrowser(CteTabCtrl *pTC)
{
	m_cRef = 1;
	m_bEmpty = TRUE;
	m_bInit = FALSE;
	m_bsFilter = NULL;
	teSetLong(&m_vRoot, 0);
	m_pTV = new CteTreeView();
	m_dwCookie = 0;
	VariantInit(&m_vData);
	Init(pTC, TRUE);
}

void CteShellBrowser::Init(CteTabCtrl *pTC, BOOL bNew)
{
	m_hwnd = NULL;
	m_hwndLV = NULL;
	m_hwndDV = NULL;
	m_hwndDT = NULL;
	m_pShellView = NULL;
	m_pExplorerBrowser = NULL;
	m_pServiceProvider = NULL;
	m_dwEventCookie = 0;
	m_pidl = NULL;
	m_pFolderItem = NULL;
	m_pFolderItem1 = NULL;
	m_nUnload = 0;
	m_DefProc = NULL;
	m_pDropTarget = NULL;
	m_pDragItems = NULL;
	m_pDSFV = NULL;
	m_pSF2 = NULL;
	m_clrText = GetSysColor(COLOR_WINDOWTEXT);
	m_clrBk = GetSysColor(COLOR_WINDOW);
	m_clrTextBk = m_clrBk;
	m_bRefreshing = FALSE;
#ifdef _2000XP
	m_pSFVCB = NULL;
	m_nColumns = 0;
#endif
	m_bCheckLayout = FALSE;
	m_bRefreshLator = FALSE;
	m_bShowFrames = FALSE;
	m_nCreate = 0;
	m_nDefultColumns = 0;
	m_pDefultColumns = NULL;
	m_bNavigateComplete = FALSE;
	m_dwUnavailable = 0;
	VariantClear(&m_vData);

	for (int i = SB_Count; i--;) {
		m_param[i] = g_paramFV[i];
	}

	m_pTC = pTC;
	if (bNew) {
		m_ppLog = new FolderItem*[MAX_HISTORY];
	}
	for (int i = Count_SBFunc; i--;) {
		m_pDispatch[i] = NULL;
	}
	m_bVisible = FALSE;
	m_nLogCount = 0;
	m_nLogIndex = -1;
	m_nFolderSizeIndex = MAXINT;
	m_nLabelIndex = MAXINT;
	m_nSizeIndex = MAXINT;

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
	for (int i = _countof(m_pDispatch); i-- > 0;) {
		teReleaseClear(&m_pDispatch[i]);
	}
	try {
		DestroyView(0);
	} catch (...) {}
	try {
		while (--m_nLogCount >= 0) {
			teReleaseClear(&m_ppLog[m_nLogCount]);
		}
		teSysFreeString(&m_bsFilter);
		teUnadviseAndRelease(m_pDSFV, DIID_DShellFolderViewEvents, &m_dwCookie);
		m_pDSFV = NULL;
#ifdef _2000XP
		teReleaseClear(&m_pSFVCB);
		if (m_nColumns) {
			delete [] m_pColumns;
			m_nColumns = 0;
		}
#endif
		teReleaseClear(&m_pSF2);
		if (m_pDefultColumns) {
			delete [] m_pDefultColumns;
			m_pDefultColumns = NULL;
		}
		teILFreeClear(&m_pidl);
		teReleaseClear(&m_pFolderItem);
		teReleaseClear(&m_pFolderItem1);
	} catch (...) {}
	VariantClear(&m_vData);
	VariantClear(&m_vRoot);
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
		QITABENT(CteShellBrowser, IDropTarget),
		QITABENT(CteShellBrowser, IPersist),
		QITABENT(CteShellBrowser, IPersistFolder),
		QITABENT(CteShellBrowser, IPersistFolder2),
#ifdef _2000XP
		QITABENT(CteShellBrowser, IShellFolderViewCB),
#endif
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
	for (int nCount = GetMenuItemCount(hmenuShared); nCount-- > 0;) {
		DeleteMenu(hmenuShared, nCount, MF_BYPOSITION);
	}
	return S_OK;
*///
	return E_NOTIMPL;
}

STDMETHODIMP CteShellBrowser::SetStatusTextSB(LPCWSTR lpszStatusText)
{
	if (lpszStatusText) {
		lstrcpyn(g_szStatus, lpszStatusText, _countof(g_szStatus) - 1);
	} else {
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
	if (wFlags & SBSP_REDIRECT) {
		if (m_pidl) {
			if (m_nLogIndex >= 0 && m_nLogIndex < m_nLogCount && m_ppLog[m_nLogIndex]) {
				m_ppLog[m_nLogIndex]->QueryInterface(IID_PPV_ARGS(&pid));
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
	teReleaseClear(&pSB);
	return hr;
}

HRESULT CteShellBrowser::Navigate3(FolderItem *pFolderItem, UINT wFlags, DWORD *param, CteShellBrowser **ppSB, FolderItems *pFolderItems)
{
	HRESULT hr = E_FAIL;
	m_pTC->LockUpdate();
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
				pSB = GetNewShellBrowser(m_pTC);
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
					TabCtrl_SetCurSel(m_pTC->m_hwnd, m_pTC->m_nIndex + 1);
					m_pTC->m_nIndex = TabCtrl_GetCurSel(m_pTC->m_hwnd);
				} else {
					m_pTC->TabChanged(TRUE);
				}
			} else {
				pSB->Close(true);
			}
		}
	} catch (...) {
		hr = E_FAIL;
	}
	teReleaseClear(&pFolderItem);
	m_pTC->UnLockUpdate(FALSE);
	return hr;
}

BOOL CteShellBrowser::Navigate1(FolderItem *pFolderItem, UINT wFlags, FolderItems *pFolderItems, FolderItem *pPrevious, LPITEMIDLIST *ppidl, int nErrorHandleing)
{
	CteFolderItem *pid = NULL;
	if (pFolderItem) {
		if SUCCEEDED(pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid)) {
			if (!pid->m_pidl) {
				if (pid->m_v.vt == VT_BSTR) {
					if (tePathIsNetworkPath(pid->m_v.bstrVal)) {
						Navigate1Ex(pid->m_v.bstrVal, pFolderItems, wFlags, pPrevious, nErrorHandleing);
						pid->Release();
						return TRUE;
					}
					if (ppidl) {
						pid->GetCurFolder(ppidl);
						if (!pid->m_pidl && m_pShellView) {
							teILFreeClear(ppidl);
							Navigate1Ex(pid->m_v.bstrVal, pFolderItems, wFlags, pPrevious, nErrorHandleing);
							pid->Release();
							return TRUE;
						}
						pid->Release();
						return FALSE;
					}
				}
				if (pid->m_v.vt == VT_NULL) {
					return FALSE;
				}
			}
			pid->Release();
		}
		if (ppidl) {
			teGetIDListFromObject(pFolderItem, ppidl);
		}
	}
	return FALSE;
}

VOID CteShellBrowser::Navigate1Ex(LPOLESTR pstr, FolderItems *pFolderItems, UINT wFlags, FolderItem *pPrevious, int nErrorHandling)
{
	TEInvoke *pInvoke = new TEInvoke[1];
	QueryInterface(IID_PPV_ARGS(&pInvoke->pdisp));
	pInvoke->cRef = 2;
	pInvoke->dispid = 0x20000007;//Navigate2Ex
	pInvoke->cArgs = 4;
	pInvoke->pv = GetNewVARIANT(4);
	pInvoke->wErrorHandling = g_nLockUpdate || m_nUnload == 2 ? 1 : nErrorHandling;
	pInvoke->wMode = 0;
	pInvoke->pidl = NULL;
	teSetSZ(&pInvoke->pv[3], pstr);
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

	teReleaseClear(&m_pFolderItem1);
	if (!(wFlags & SBSP_RELATIVE)) {
		m_dwUnavailable = 0;
	}
	m_bRefreshLator = FALSE;
	m_nFolderSizeIndex = MAXINT;
	m_nLabelIndex = MAXINT;
	GetNewObject(&m_pDispatch[SB_TotalFileSize]);
	CteFolderItem *pid1 = NULL;
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
	if (ILIsEqual(pidl, g_pidlResultsFolder)) {
		if SUCCEEDED(pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) {
			m_dwUnavailable = pid1->m_dwUnavailable;
			pid1->Release();
		}
	} else {
		m_dwUnavailable = 0;
	}
	if (hr != S_OK) {
		if (hr == S_FALSE) {
			if (!m_pFolderItem) {
				if SUCCEEDED(pFolderItem->QueryInterface(IID_PPV_ARGS(&m_pFolderItem))) {
					CteFolderItem *pid1 = NULL;
					if (SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) && !pid1->m_pidl) {
						teILCloneReplace(&pid1->m_pidl, g_pidlResultsFolder);
						pid1->m_dwUnavailable = GetTickCount();
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
								GetFolderItemFromPidl(&m_pFolderItem1, pidl);
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
#ifdef _2000XP
	if (g_bUpperVista) {
#endif
		if (m_param[SB_Type] == 2 || m_bShowFrames || ILIsParent(g_pidlCP, pidl, FALSE)) {
			m_bShowFrames = FALSE;
			dwFrame = EBO_SHOWFRAMES;
		}
		//ExplorerBrowser
		if (m_pExplorerBrowser) {
			m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
			if (dwFrame == (m_param[SB_Options] & EBO_SHOWFRAMES)) {
				if (!ILIsEqual(pidl, g_pidlResultsFolder) || !ILIsEqual(m_pidl, g_pidlResultsFolder)) {
					if (GetShellFolder2(pidl) == S_OK) {
						IFolderViewOptions *pOptions;
						if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
							pOptions->SetFolderViewOptions(FVO_VISTALAYOUT, teGetFolderViewOptions(pidl, uViewMode));
							pOptions->Release();
						}
						m_pExplorerBrowser->SetOptions(static_cast<EXPLORER_BROWSER_OPTIONS>((m_param[SB_Options] & ~(EBO_SHOWFRAMES | EBO_NAVIGATEONCE | EBO_ALWAYSNAVIGATE)) | dwFrame | EBO_NOTRAVELLOG));
						BrowseToObject();
						teCoTaskMemFree(pidl);
						return S_OK;
					}
				}
			}
		}
#ifdef _2000XP
	}
#endif
	teCoTaskMemFree(m_pidl);
	m_pidl = pidl;
	m_pFolderItem = m_pFolderItem1;
	m_pFolderItem1 = NULL;
	if (g_nLockUpdate || !m_pTC->m_bVisible) {
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
			if (pPrevious) {
				teGetIDListFromObject(pPrevious, &m_pidl);
				pPrevious->AddRef();
				teReleaseClear(&m_pFolderItem);
				m_pFolderItem = pPrevious;
			}
			m_nLogIndex = m_nPrevLogIndex;
			return hr;
		}
		if (!teILIsEqual(m_pFolderItem, pPrevious)) {
			teSysFreeString(&m_bsFilter);
			teReleaseClear(&m_pDispatch[SB_OnIncludeObject]);
		}
	}

	SetRectEmpty(&rc);

	//History / Management
	SetHistory(pFolderItems, wFlags);
	DestroyView(m_param[SB_Type]);
	Show(FALSE, 2);
	if (dwFrame || teGetFolderViewOptions(m_pidl, m_param[SB_ViewMode]) == FVO_DEFAULT) {
		//ExplorerBrowser
		hr = NavigateEB(dwFrame);
	} else {
		//ShellBrowser
		hr = NavigateSB(pPreviousView, pPrevious);
	}
	if (hr != S_OK) {
		teReleaseClear(&m_pShellView);
		m_pShellView = pPreviousView;

		teILFreeClear(&m_pidl);
		teGetIDListFromObject(pPrevious, &m_pidl);
		if (pPrevious) {
			pPrevious->AddRef();
			teReleaseClear(&m_pFolderItem);
			m_pFolderItem = pPrevious;
		}
		m_nLogIndex = m_nPrevLogIndex;
		if (m_pSF2 && m_pTC && GetTabIndex() == m_pTC->m_nIndex) {
			Show(TRUE, 0);
		}
		return hr;
	}

	if (pPreviousView) {
//		pPreviousView->SaveViewState();
		pPreviousView->DestroyViewWindow();
		teReleaseClear(&pPreviousView);
	} else {
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
			} else {
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
	if (m_pTC && GetTabIndex() == m_pTC->m_nIndex) {
		Show(TRUE, 0);
	}
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
			if (SUCCEEDED(CoCreateInstance(CLSID_ExplorerBrowser,
			NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pExplorerBrowser)))) {
				FOLDERSETTINGS fs = { FVM_AUTO, (m_param[SB_FolderFlags] | FWF_USESEARCHFOLDER) & ~FWF_NOENUMREFRESH };
				if (SUCCEEDED(m_pExplorerBrowser->Initialize(m_pTC->m_hwndStatic, &rc, &fs))) {
					if (GetShellFolder2(m_pidl) == S_OK) {
						m_pExplorerBrowser->Advise(static_cast<IExplorerBrowserEvents *>(this), &m_dwEventCookie);
						IFolderViewOptions *pOptions;
						if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
							pOptions->SetFolderViewOptions(FVO_VISTALAYOUT, teGetFolderViewOptions(m_pidl, m_param[SB_ViewMode]));
							pOptions->Release();
						}
						m_pExplorerBrowser->SetOptions(static_cast<EXPLORER_BROWSER_OPTIONS>((m_param[SB_Options] & ~(EBO_SHOWFRAMES | EBO_NAVIGATEONCE | EBO_ALWAYSNAVIGATE)) | dwFrame | EBO_NOTRAVELLOG));
						m_pServiceProvider = new CteServiceProvider(static_cast<IShellBrowser *>(this), NULL);
						IUnknown_SetSite(m_pExplorerBrowser, m_pServiceProvider);
/*///
						IFolderFilterSite *pFolderFilterSite;
						if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pFolderFilterSite))) {
							pFolderFilterSite->SetFilter(static_cast<IFolderFilter *>(this));
						}
*///
						hr = BrowseToObject();
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
	}
	InterlockedDecrement(&m_nCreate);
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

		pMem = new CteMemory(4 * sizeof(int), &m_param[SB_ViewMode], 1, L"FOLDERSETTINGS");
		pMem->QueryInterface(IID_PPV_ARGS(&pv[2].punkVal));
		teSetObjectRelease(&pv[2], pMem);
		teSetLong(&pv[1], wFlags);
		if (!teILIsEqual(m_pFolderItem, pPrevious)) {
			teSetObject(&pv[0], pPrevious);
		}
		m_bInit = true;
		try {
			Invoke4(g_pOnFunc[TE_OnBeforeNavigate], &vResult, 4, pv);
		} catch (...) {}
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
				teILFreeClear(&pid1->m_pidlFocused);
				pFV->Item(i, &pid1->m_pidlFocused);
				pFV->ItemCount(SVGIO_SELECTION, &pid1->m_nSelected);
				pid1->Release();
			}
			pFV->Release();
		}
	}
}

VOID CteShellBrowser::FocusItem(BOOL bFree)
{
	CteFolderItem *pid;
	if (m_pFolderItem && SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid))) {
		if (pid->m_pidlFocused) {
			SelectItemEx(&pid->m_pidlFocused, pid->m_nSelected == 1 ?
				SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS | SVSI_DESELECTOTHERS | SVSI_SELECT :
				SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS | SVSI_DESELECTOTHERS, bFree);
		}
		m_dwUnavailable = pid->m_dwUnavailable;
		pid->Release();
	}
	if (!bFree) {
		int nTop = -1;
		IFolderView *pFV;
		if (m_pShellView && SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV)))) {
			if (SUCCEEDED(pFV->GetFocusedItem(&nTop)) && nTop < 0) {
				if (SUCCEEDED(pFV->ItemCount(SVGIO_ALLVIEW, &nTop)) && nTop) {
					pFV->SelectItem(0, SVSI_FOCUSED | SVSI_ENSUREVISIBLE);
				}
			}
			pFV->Release();
		}
		SetTimer(m_hwnd, 1, 32, teTimerProcFocus);
		m_nFocusItem = 0;
	}
}

HRESULT CteShellBrowser::GetAbsPidl(LPITEMIDLIST *ppidlOut, FolderItem **ppid, FolderItem *pid, UINT wFlags, FolderItems *pFolderItems, FolderItem *pPrevious, CteShellBrowser *pHistSB)
{
	if (wFlags & SBSP_PARENT) {
		if (pPrevious) {
			if (m_bsFilter || m_pDispatch[SB_OnIncludeObject]) {
				teSysFreeString(&m_bsFilter);
				teReleaseClear(&m_pDispatch[SB_OnIncludeObject]);
				Refresh(TRUE);
				return E_FAIL;
			}
			if (teILGetParent(pPrevious, ppid)) {
				if (Navigate1(*ppid, wFlags & ~SBSP_PARENT, pFolderItems, pPrevious, ppidlOut, 0)) {
					return S_FALSE;
				}
				VARIANT v;
				if SUCCEEDED(ppid[0]->QueryInterface(IID_PPV_ARGS(&v.pdispVal))) {
					v.vt = VT_DISPATCH;
					CteFolderItem *pid1 = new CteFolderItem(&v);
					VARIANT v2;
					VariantInit(&v2);
					teGetDisplayNameOf(&v, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &v2);
					if (teIsDesktopPath(v2.bstrVal) || !ILIsEmpty(pid1->GetPidl())) {
						teGetIDListFromObject(pPrevious, &pid1->m_pidlFocused);
						pid1->m_nSelected = 1;
					} else {
						pid1->Release();
						pid1 = new CteFolderItem(&v2);
						teILCloneReplace(&pid1->m_pidl, g_pidlResultsFolder);
						pid1->m_dwUnavailable = GetTickCount();
					}
					ppid[0]->Release();
					ppid[0] = pid1;
					VariantClear(&v2);
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
		if (nLogIndex > 0) {
			if (m_pFolderItem) {
				m_ppLog[nLogIndex]->Release();
				m_pFolderItem->QueryInterface(IID_PPV_ARGS(&m_ppLog[nLogIndex]));
				SaveFocusedItemToHistory();
			}
			if (teGetIDListFromObjectEx(pHistSB->m_ppLog[--nLogIndex], ppidlOut)) {
				pHistSB->m_ppLog[nLogIndex]->QueryInterface(IID_PPV_ARGS(ppid));
				if (this == pHistSB) {
					m_nLogIndex = nLogIndex;
				}
				return S_OK;
			}
		}
		return E_FAIL;
	}
	if (wFlags & SBSP_NAVIGATEFORWARD) {
		int nLogIndex = pHistSB->m_nLogIndex;
		if (nLogIndex < pHistSB->m_nLogCount - 1) {
			if (m_pFolderItem) {
				m_ppLog[nLogIndex]->Release();
				m_pFolderItem->QueryInterface(IID_PPV_ARGS(&m_ppLog[nLogIndex]));
				SaveFocusedItemToHistory();
			}
			if (teGetIDListFromObjectEx(pHistSB->m_ppLog[++nLogIndex], ppidlOut)) {
				pHistSB->m_ppLog[nLogIndex]->QueryInterface(IID_PPV_ARGS(ppid));
				if (this == pHistSB) {
					m_nLogIndex = nLogIndex;
				}
				return S_OK;
			}
		}
		return E_FAIL;
	}
	LPITEMIDLIST pidl = NULL;
	if (Navigate1(pid, wFlags, pFolderItems, pPrevious, &pidl, 2)) {
		return S_FALSE;
	}
	if (wFlags & SBSP_RELATIVE) {
		LPITEMIDLIST pidlPrevius = NULL;
		teGetIDListFromObject(pPrevious, &pidlPrevius);
		if (pidlPrevius && !ILIsEqual(pidlPrevius, g_pidlResultsFolder) && !ILIsEmpty(pidl)) {
			*ppidlOut = ILCombine(pidlPrevius, pidl);
			GetFolderItemFromPidl(ppid, *ppidlOut);
		} else {
			*ppidlOut = ILClone(pidlPrevius);
			if (pPrevious) {
				pPrevious->QueryInterface(IID_PPV_ARGS(ppid));
			} else {
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
		return teGetIDListFromObject(*ppid, ppidlOut) ? S_OK : E_FAIL;
	}
	return E_FAIL;
}

VOID CteShellBrowser::Refresh(BOOL bCheck)
{
	m_bRefreshLator = FALSE;
	if (!m_dwUnavailable) {
		teDoCommand(this, m_hwnd, WM_NULL, 0, 0);//Save folder setings
	}
	GetNewObject(&m_pDispatch[SB_TotalFileSize]);
	if (bCheck) {
		VARIANT v, vResult;
		VariantInit(&v);
		if SUCCEEDED(m_pFolderItem->QueryInterface(IID_PPV_ARGS(&v.pdispVal))) {
			v.vt = VT_DISPATCH;
			VariantInit(&vResult);
			teGetDisplayNameOf(&v, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &vResult);
			VariantClear(&v);
			if (vResult.vt == VT_BSTR) {
				if (!m_pShellView || tePathIsNetworkPath(vResult.bstrVal) || m_bShowFrames || m_dwUnavailable) {
					TEInvoke *pInvoke = new TEInvoke[1];
					QueryInterface(IID_PPV_ARGS(&pInvoke->pdisp));
					pInvoke->dispid = 0x20000206;//Refresh2Ex
					pInvoke->cRef = 2;
					pInvoke->cArgs = 1;
					pInvoke->pv = GetNewVARIANT(1);
					pInvoke->wErrorHandling = 1;
					pInvoke->wMode = 0;
					pInvoke->pidl = NULL;
					VariantInit(&pInvoke->pv[0]);
					pInvoke->pv[0].vt = VT_BSTR;
					pInvoke->pv[0].bstrVal = ::SysAllocString(vResult.bstrVal);
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
			if (m_dwUnavailable || (g_pidlLibrary && ILIsParent(g_pidlLibrary, m_pidl, FALSE))) {
				BrowseObject(NULL, SBSP_RELATIVE | SBSP_SAMEBROWSER);
				return;
			}
			IFolderView2 *pFV2;
			if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
				pFV2->SetText(FVST_EMPTYTEXT, m_bsFilter || m_pDispatch[SB_OnIncludeObject] ? L"" : NULL);
				pFV2->Release();
			}
			m_bRefreshing = TRUE;
			if (ILIsEqual(m_pidl, g_pidlResultsFolder)) {
				RemoveAll();
				m_dwSessionId = (rand() * 0x20000) + (GetTickCount() & 0x1ffff);
				SetTabName();
				DoFunc(TE_OnNavigateComplete, this, E_NOTIMPL);
			} else {
				m_pShellView->Refresh();
			}
		} else if (m_nUnload == 0) {
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
			if (IsChild(g_pWebBrowser->m_hwndBrowser, hwnd) || (m_pTV && IsChild(m_pTV->m_hwnd, hwnd))) {
				return FALSE;
			}
			WCHAR szClass[MAX_CLASS_NAME];
			GetClassName(hwnd, szClass, MAX_CLASS_NAME);
			if (lstrcmp(szClass, WC_TREEVIEW) == 0) {
				return FALSE;
			}
		}
	}
	if (m_pShellView) {
		if (IsChild(m_hwnd, hwnd)) {
			SetFocus(m_hwnd);
		}
		try {
			m_pShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
		} catch (...) {
			m_pShellView = NULL;
			g_nException = 0;
			Suspend(0);
		}
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
	teUnadviseAndRelease(m_pDSFV, DIID_DShellFolderViewEvents, &m_dwCookie);
	if SUCCEEDED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&m_pDSFV))) {
		teAdvise(m_pDSFV, DIID_DShellFolderViewEvents, static_cast<IDispatch *>(this), &m_dwCookie);
	} else {
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
			ListView_SetExtendedListViewStyleEx(m_hwndLV, LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES, HIWORD(m_param[SB_FolderFlags]));
		}
#endif
		GetViewModeAndIconSize(FALSE);
		SetLVSettings();
	}
}

VOID CteShellBrowser::InitFolderSize()
{
	m_nFolderSizeIndex = MAXINT;
	m_nLabelIndex = MAXINT;
	m_nSizeIndex = MAXINT;

	if (m_hwndLV) {
		HWND hHeader = ListView_GetHeader(m_hwndLV);
		if (hHeader) {
			UINT nHeader = Header_GetItemCount(hHeader);
			BSTR bsTotalFileSize = tePSGetNameFromPropertyKeyEx(PKEY_TotalFileSize, 0, m_pShellView);
			BSTR bsLabel = g_pOnFunc[TE_Labels] ? tePSGetNameFromPropertyKeyEx(PKEY_Contact_Label, 0, m_pShellView) : NULL;
			BSTR bsSize = tePSGetNameFromPropertyKeyEx(PKEY_Size, 0, m_pShellView);
			WCHAR szText[MAX_COLUMN_NAME_LEN];
			HD_ITEM hdi = { HDI_TEXT | HDI_FORMAT, 0, szText, NULL, MAX_COLUMN_NAME_LEN };
			for (int i = nHeader; i-- > 0;) {
				Header_GetItem(hHeader, i, &hdi);
				if (szText[0]) {
#ifdef _2000XP
					if (g_bUpperVista) {
#endif
						if (lstrcmpi(szText, bsTotalFileSize) == 0) {
							m_nFolderSizeIndex = i;
							hdi.fmt |= HDF_RIGHT;
							Header_SetItem(hHeader, i, &hdi);
							InvalidateRect(m_hwndLV, NULL, FALSE);
							continue;
						}
						if (lstrcmpi(szText, bsLabel) == 0) {
							m_nLabelIndex = i;
							continue;
						}
#ifdef _2000XP
					}
#endif
					if (lstrcmpi(szText, bsSize) == 0) {
						m_nSizeIndex = i;
					}
				}
			}
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
		if (SUCCEEDED(m_pSF2->GetDetailsEx(pidl, &PKEY_Size, &v))) {
			teStrFormatSize(m_param[SB_SizeFormat], v.ullVal, szText, cch);
		}
		VariantClear(&v);
	}
}

VOID CteShellBrowser::SetFolderSize(LPCITEMIDLIST pidl, LPWSTR szText, int cch)
{
	if (m_pSF2) {
		VARIANT v;
		VariantInit(&v);
		WIN32_FIND_DATA wfd;
		teSHGetDataFromIDList(m_pSF2, pidl, SHGDFIL_FINDDATA, &wfd, sizeof(WIN32_FIND_DATA));
		if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
			teGetProperty(m_pDispatch[SB_TotalFileSize], wfd.cFileName, &v);
			if (v.vt == VT_BSTR) {
				lstrcpyn(szText, v.bstrVal, cch);
			} else if (v.vt != VT_EMPTY) {
				teStrFormatSize(m_param[SB_SizeFormat], GetLLFromVariant(&v), szText, cch);
			} else if (m_pDispatch[SB_TotalFileSize]) {
				v.vt = VT_BSTR;
				v.bstrVal = NULL;
				tePutProperty(m_pDispatch[SB_TotalFileSize], wfd.cFileName, &v);
				TEFS *pFS = new TEFS[1];
				pFS->bsName = ::SysAllocString(wfd.cFileName);
				teGetDisplayNameBSTR(m_pSF2, pidl, SHGDN_FORPARSING, &pFS->bsPath);
				pFS->ppidl = &m_pidl;
				pFS->pidl = m_pidl;
				pFS->hwnd = m_hwndLV;
				CoMarshalInterThreadInterfaceInStream(IID_IDispatch, m_pDispatch[SB_TotalFileSize], &pFS->pStrmTotalFileSize);
				_beginthread(&threadFolderSize, 0, pFS);
			}
		} else if (SUCCEEDED(m_pSF2->GetDetailsEx(pidl, &PKEY_Size, &v))) {
			teStrFormatSize(m_param[SB_SizeFormat], v.ullVal, szText, cch);
		}
		VariantClear(&v);
	}
}

VOID CteShellBrowser::SetLabel(LPCITEMIDLIST pidl, LPWSTR szText, int cch)
{
	IShellFolder2 *pSF2;
	BSTR bs;
	VARIANT v;
	VariantInit(&v);
	if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pSF2))) {
		if SUCCEEDED(teGetDisplayNameBSTR(pSF2, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &bs)) {
			if (::SysStringLen(bs)) {
				if (g_bLabelsMode) {
					teGetProperty(g_pOnFunc[TE_Labels], bs, &v);
				} else {
					VARIANTARG *pv = GetNewVARIANT(1);
					teSetBSTR(&pv[0], bs, -1);
					Invoke4(g_pOnFunc[TE_Labels], &v, 1, pv);
				}
				if (v.vt == VT_BSTR) {
					lstrcpyn(szText, v.bstrVal, cch);
				}
				::VariantClear(&v);
			}
			::SysFreeString(bs);
		}
		pSF2->Release();
	}
}

VOID CteShellBrowser::GetViewModeAndIconSize(BOOL bGetIconSize)
{
	if (!m_pShellView) {
		return;
	}
	FOLDERSETTINGS fs;
	UINT uViewMode = 0;
	int iImageSize = m_param[SB_IconSize];
	if SUCCEEDED(m_pShellView->GetCurrentInfo(&fs)) {
		uViewMode = fs.ViewMode;
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
			} else if (uViewMode == FVM_TILE) {
				iImageSize = 48;
			} else if (uViewMode == FVM_THUMBNAIL) {
				HKEY hKey;
				if (RegOpenKeyEx(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
					DWORD dwSize = sizeof(iImageSize);
					RegQueryValueEx(hKey, L"ThumbnailSize", NULL, NULL, (LPBYTE)&iImageSize, &dwSize);
					RegCloseKey(hKey);
				}
			} else {
				iImageSize = GetSystemMetrics(SM_CXSMICON);
			}
		}
#endif
		m_param[SB_IconSize] = iImageSize;
		m_param[SB_ViewMode] = uViewMode;
		if (m_bCheckLayout) {
			FOLDERVIEWOPTIONS fvo = teGetFolderViewOptions(m_pidl, uViewMode);
			if (m_pExplorerBrowser) {
				IFolderViewOptions *pOptions;
				if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOptions))) {
					FOLDERVIEWOPTIONS fvo0 = fvo;
					pOptions->GetFolderViewOptions(&fvo0);
					m_pExplorerBrowser->GetOptions((EXPLORER_BROWSER_OPTIONS *)&m_param[SB_Options]);
					if ((fvo ^ fvo0 & FVO_VISTALAYOUT) || (m_param[SB_Options] & EBO_SHOWFRAMES) != (m_param[SB_Type] == 2 ? EBO_SHOWFRAMES : 0)) {
						m_bCheckLayout = FALSE;
					}
					pOptions->Release();
				}
			} else if (m_param[SB_Type] == 2 || fvo == FVO_DEFAULT) {
				m_bCheckLayout = FALSE;
			}
			if (!m_bCheckLayout) {
				SetRedraw(FALSE);
				Suspend(0);
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
		v.vt = VT_I4;
		while (--nCount >= 0) {
			v.lVal = nCount;
			pFolderItems->Item(v, &m_ppLog[m_nLogCount++]);
		}
		m_nLogIndex = m_nLogCount - nLogIndex - 1;
		teReleaseClear(&pFolderItems);
	} else if ((wFlags & (SBSP_NAVIGATEBACK | SBSP_NAVIGATEFORWARD | SBSP_WRITENOHISTORY)) == 0) {
		if (ILIsEqual(m_pidl, g_pidlResultsFolder)) {
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
		if (m_nLogCount > 0 && teILIsEqual(m_pFolderItem, m_ppLog[m_nLogIndex])) {
			m_ppLog[m_nLogIndex]->Release();
			m_pFolderItem->QueryInterface(IID_PPV_ARGS(&m_ppLog[m_nLogIndex]));
			return;
		}
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
		if (m_pFolderItem) {
			m_pFolderItem->QueryInterface(IID_PPV_ARGS(&m_ppLog[m_nLogCount++]));
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
	if (m_pTC) {
		CheckChangeTabTC(m_pTC->m_hwnd, false);
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
	if (m_pDispatch[SB_OnIncludeObject]) {
		VARIANT vResult;
		VariantInit(&vResult);
		VARIANTARG *pv = GetNewVARIANT(3);
		teSetObject(&pv[2], this);
		if SUCCEEDED(teGetDisplayNameBSTR(m_pSF2, pidl, SHGDN_NORMAL, &pv[1].bstrVal)) {
			pv[1].vt = VT_BSTR;
		}
		if SUCCEEDED(teGetDisplayNameBSTR(m_pSF2, pidl, SHGDN_INFOLDER | SHGDN_FORPARSING, &pv[0].bstrVal)) {
			pv[0].vt = VT_BSTR;
		}
		Invoke4(m_pDispatch[SB_OnIncludeObject], &vResult, 3, pv);
		return GetIntFromVariantClear(&vResult);
	}
	if (m_bsFilter) {
		HRESULT hr = S_OK;
		if (m_pSF2) {
			BSTR bs;
			if SUCCEEDED(teGetDisplayNameBSTR(m_pSF2, pidl, SHGDN_NORMAL, &bs)) {
				if (!tePathMatchSpec(bs, m_bsFilter)) {
					hr = S_FALSE;
					BSTR bs2;
					if SUCCEEDED(teGetDisplayNameBSTR(m_pSF2, pidl, SHGDN_INFOLDER | SHGDN_FORPARSING, &bs2)) {
						if (lstrcmpi(bs, bs2) && tePathMatchSpec(bs2, m_bsFilter)) {
							hr = S_OK;
						}
						::SysFreeString(bs2);
					}
				}
				::SysFreeString(bs);
			}
		}
		return hr;
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
	*pdwFlags = (m_param[SB_ViewFlags] & (~(CDB2GVF_NOINCLUDEITEM | CDB2GVF_NOSELECTVERB))) | ((m_bsFilter || m_pDispatch[SB_OnIncludeObject] || ILIsEqual(m_pidl, g_pidlResultsFolder)) ? 0 : CDB2GVF_NOINCLUDEITEM);
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
	HRESULT hr = teGetDispId(methodSB, _countof(methodSB), g_maps[MAP_SB], *rgszNames, rgDispId, true);
	if (hr != S_OK && m_pDSFV) {
		return m_pDSFV->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
	}
	return hr;
}

VOID CteShellBrowser::SetColumnsStr(BSTR bsColumns)
{
	int nCount = 0;
	int nAllWidth = 0;
	LPTSTR *lplpszArgs = NULL;
	if (bsColumns && bsColumns[0]) {
		lplpszArgs = CommandLineToArgvW(bsColumns, &nCount);
		nCount /= 2;
	}
	TEmethod *methodArgs = new TEmethod[nCount + 1];
	BOOL *pbAlloc = new BOOL[nCount + 1];
	BSTR bsName = tePSGetNameFromPropertyKeyEx(PKEY_ItemNameDisplay, 0, m_pShellView);
	BOOL bNoName = TRUE;
	for (int i = nCount; i-- > 0;) {
		pbAlloc[i] = FALSE;
		methodArgs[i].name = lplpszArgs[i * 2];
		PROPERTYKEY propKey;
		if SUCCEEDED(lpfnPSPropertyKeyFromStringEx(lplpszArgs[i * 2], &propKey)) {
			methodArgs[i].name = tePSGetNameFromPropertyKeyEx(propKey, 0, m_pShellView);
			pbAlloc[i] = TRUE;
		}
		if (lstrcmpi(methodArgs[i].name, bsName) == 0) {
			bNoName = FALSE;
		}
		int n;
		if (StrToIntEx(lplpszArgs[i * 2 + 1], STIF_DEFAULT, &n) && n < 32768) {
			nAllWidth += n;
		} else {
			n = -1;
			nAllWidth += 100;
		}
		methodArgs[i].id = n;
	}
	if (bNoName && nAllWidth) {
		methodArgs[nCount].name = bsName;
		methodArgs[nCount++].id = -1;
		nAllWidth += 100;
	}
	if (nCount == 0 || nAllWidth > nCount + 4) {
		IColumnManager *pColumnManager;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager))) {
			UINT uCount;
			if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_ALL, &uCount)) {
				PROPERTYKEY *pPropKey = new PROPERTYKEY[uCount];
				if SUCCEEDED(pColumnManager->GetColumns(CM_ENUM_ALL, pPropKey, uCount)) {
					VARIANT v;
					VariantInit(&v);
					IDispatch *pdispColumns = NULL;
					GetNewObject(&pdispColumns);
					CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_NAME };
					for (int i = uCount; i-- > 0;) {
						if (pColumnManager->GetColumnInfo(pPropKey[i], &cmci) == S_OK) {
							if (cmci.wszName[0]) {
								VariantClear(&v);
								teSetLong(&v, i);
								tePutProperty(pdispColumns, cmci.wszName, &v);
							}
						}
					}
					int k = 0;
					PROPERTYKEY *pPropKey2 = new PROPERTYKEY[uCount];
					int *pnWidth = new int[uCount];
					VariantClear(&v);
					for (int i = 0; i < nCount; i++) {
						if (teGetPropertyI(pdispColumns, methodArgs[i].name, &v) == S_OK) {
							pPropKey2[k] = pPropKey[GetIntFromVariantClear(&v)];
							pnWidth[k++] = methodArgs[i].id;
						}
					}
					if (bNoName && k == 1) {
						k = 0;
					}
					if (k == 0) {	//Default Columns
						while (k < (int)m_nDefultColumns && k < (int)uCount) {
							pPropKey2[k] = m_pDefultColumns[k];
							pnWidth[k++] = -1;
						}
					}
					if (k > 0) {
						if SUCCEEDED(pColumnManager->SetColumns(pPropKey2, k)) {
							CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_WIDTH };
							while (k--) {
								cmci.uWidth = pnWidth[k];
								pColumnManager->SetColumnInfo(pPropKey2[k], &cmci);
							}
						}
					}
					delete [] pnWidth;
					delete [] pPropKey2;
					pdispColumns->Release();
					VariantClear(&v);
				}
				delete [] pPropKey;
			}
			pColumnManager->Release();
		} else {
#ifdef _2000XP
			UINT uCount = m_nColumns;
			SHELLDETAILS sd;
			int nWidth;
			if (m_pSF2) {
				for (int i = uCount; i-- > 0;) {
					m_pSF2->GetDefaultColumnState(i, &m_pColumns[i].csFlags);
					m_pColumns[i].csFlags &= ~SHCOLSTATE_ONBYDEFAULT;
					m_pColumns[i].nWidth = -3;
				}
				m_pColumns[uCount - 1].csFlags = SHCOLSTATE_TYPE_STR;
				m_pColumns[uCount - 2].csFlags = SHCOLSTATE_TYPE_STR;

				int *piArgs = SortTEMethod(methodArgs, nCount);
				BSTR bs = GetColumnsStr(FALSE);
				if (bs && bs[0]) {
					int nCur;
					LPTSTR *lplpszColumns = CommandLineToArgvW(bs, &nCur);
					::SysFreeString(bs);
					nCur /= 2;
					BOOL bDiff = nCount != nCur || nCur == 0;
					if (!bDiff) {
						TEmethod *methodColumns = new TEmethod[nCur];
						for (int i = nCur; i-- > 0;) {
							methodColumns[i].name = lplpszColumns[i * 2];
						}
						int *piCur = SortTEMethod(methodColumns, nCur);
						for (int i = nCur; i-- > 0;) {
							if (lstrcmpi(methodArgs[piArgs[i]].name, methodColumns[piCur[i]].name)) {
								bDiff = true;
								break;
							}
						}
						delete [] piCur;
						delete [] methodColumns;
					}
					LocalFree(lplpszColumns);

					if (bDiff) {
						int nIndex;
						BOOL bDefault = true;
						if (nCount) {
							for (UINT i = 0; i < m_nColumns && GetDetailsOf(NULL, i, &sd) == S_OK; i++) {
								BSTR bs;
								if SUCCEEDED(StrRetToBSTR(&sd.str, NULL, &bs)) {
									nIndex = teBSearch(methodArgs, nCount, piArgs, bs);
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
							for (int i = uCount; i-- > 0;) {
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
						if SUCCEEDED(CreateViewWindowEx(pPreviousView)) {
							pPreviousView->DestroyViewWindow();
							teDelayRelease(&pPreviousView);
							GetShellFolderView();
						}
						Show(TRUE, 0);
						SetPropEx();
					}
					//Columns Order and Width;
					if (m_hwndLV) {
						HWND hHeader = ListView_GetHeader(m_hwndLV);
						if (hHeader) {
							UINT nHeader = Header_GetItemCount(hHeader);
							if (nHeader == nCount) {
								SetRedraw(false);
								int *piOrderArray = new int[nHeader];
								try {
									for (int i = nHeader; i-- > 0;) {
										piOrderArray[i] = i;
									}
									WCHAR szText[MAX_COLUMN_NAME_LEN];
									HD_ITEM hdi = { HDI_TEXT | HDI_WIDTH, 0, szText, NULL, MAX_COLUMN_NAME_LEN };
									int nIndex;
									for (int i = nHeader; i-- > 0;) {
										Header_GetItem(hHeader, i, &hdi);
										nIndex = teBSearch(methodArgs, nCount, piArgs, szText);
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
								} catch (...) {}
								delete [] piOrderArray;
								SetRedraw(true);
							}
						}
					}
				}
				delete [] piArgs;
				InitFolderSize();
			}
#endif
		}
		if (lplpszArgs) {
			for (int i = nCount; i-- > 0;) {
				if (pbAlloc[i]) {
					teSysFreeString(&methodArgs[i].name);
				}
			}
			delete [] pbAlloc;
			delete [] methodArgs;
			LocalFree(lplpszArgs);
		}
		teSysFreeString(&bsName);
	}
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
	if (m_dwUnavailable || m_nDefultColumns == 0) {
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
					CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_WIDTH | CM_MASK_NAME };
					for (UINT i = 0; i < uCount; i++) {
						if SUCCEEDED(pColumnManager->GetColumnInfo(pPropKey[i], &cmci)) {
							if (nFormat) {
								AddColumnDataEx(szColumns, tePSGetNameFromPropertyKeyEx(pPropKey[i], nFormat, NULL), cmci.uWidth);
							} else {
								AddColumnData(szColumns, cmci.wszName, cmci.uWidth);
							}
						}
					}
				}
				delete [] pPropKey;
			}
		}
		pColumnManager->Release();
	}
#ifdef _2000XP
	else {
		HWND hHeader = ListView_GetHeader(m_hwndLV);
		if (hHeader) {
			UINT nHeader = Header_GetItemCount(hHeader);
			if (nHeader) {
				int *piOrderArray = new int[nHeader];
				ListView_GetColumnOrderArray(m_hwndLV, nHeader, piOrderArray);
				WCHAR szText[MAX_COLUMN_NAME_LEN];
				HD_ITEM hdi = { HDI_TEXT | HDI_WIDTH, 0, szText, NULL, MAX_COLUMN_NAME_LEN };
				for (UINT i = 0; i < nHeader; i++) {
					Header_GetItem(hHeader, piOrderArray[i], &hdi);
					AddColumnDataXP(szColumns, szText, hdi.cxy, nFormat);
				}
				delete [] piOrderArray;
			}
		} else {
			SHCOLSTATEF csFlags;
			SHELLDETAILS sd;
			BSTR bs;
			for (UINT i = 0; i < m_nColumns && GetDefaultColumnState(i, &csFlags) == S_OK; i++) {
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
	}
#endif
	return SysAllocString(&szColumns[1]);
}

VOID CteShellBrowser::GetDefaultColumns()
{
	if (m_pDefultColumns) {
		delete [] m_pDefultColumns;
		m_pDefultColumns = NULL;
	}
	m_nDefultColumns = 0;
	IColumnManager *pColumnManager;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager))) {
		if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_VISIBLE, &m_nDefultColumns)) {
			m_pDefultColumns = new PROPERTYKEY[m_nDefultColumns];
			pColumnManager->GetColumns(CM_ENUM_VISIBLE, m_pDefultColumns, m_nDefultColumns);
		}
		pColumnManager->Release();
	}
#ifdef _2000XP
	else if (!g_bUpperVista) {
		m_nDefultColumns = -1;
	}
#endif
}

STDMETHODIMP CteShellBrowser::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		int nIndex;
		int nFormat = 0;
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
					VariantClear(&m_vData);
					VariantCopy(&m_vData, &pDispParams->rgvarg[nArg]);
				}
				if (pVarResult) {
					VariantCopy(pVarResult, &m_vData);
				}
				return S_OK;
			//hwnd
			case 0x10000002:
				teSetPtr(pVarResult, m_hwnd);
				return S_OK;
			//Type
			case 0x10000003:
				if (nArg >= 0) {
					DWORD dwType = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					if (dwType == CTRL_SB || dwType == CTRL_EB) {
						if (m_param[SB_Type] != dwType) {
							m_param[SB_Type] = dwType;
							if (!m_bEmpty) {
								m_bCheckLayout = TRUE;
								GetViewModeAndIconSize(TRUE);
							}
						}
					}
				}
				teSetLong(pVarResult, m_param[SB_Type]);
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
						} else if (pSB) {
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
							} catch (...) {}
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
							teGetIDListFromObject(punk, &pidl);
						}
						if (SUCCEEDED(Navigate2(pFolderItem, GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), param, pFolderItems, m_pFolderItem, this)) && m_bVisible) {
							Show(TRUE, 0);
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
					m_pTC->Move(GetTabIndex(), GetIntFromVariant(&pDispParams->rgvarg[nArg]), m_pTC);
				}
				if (pVarResult) {
					teSetLong(pVarResult, GetTabIndex());
				}
				return S_OK;
			//FolderFlags
			case 0x10000009:
				if (nArg >= 0) {
					m_param[SB_FolderFlags] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					SetFolderFlags();
				}
				teSetLong(pVarResult, m_param[SB_FolderFlags]);
				return S_OK;
			//CurrentViewMode
			case 0x10000010:
				IFolderView *pFV;
				IFolderView2 *pFV2;
				if (nArg >= 0) {
					m_param[SB_ViewMode] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					if (nArg >= 1) {
						m_param[SB_IconSize] = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
					}
					if (m_pShellView) {
						if (nArg >= 1 && SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2)))) {
							pFV2->SetViewModeAndIconSize(static_cast<FOLDERVIEWMODE>(m_param[SB_ViewMode]), m_param[SB_IconSize]);
							pFV2->Release();
						} else if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
							pFV->SetCurrentViewMode(m_param[SB_ViewMode]);
							pFV->Release();
						}
					}
				}
				SetFolderFlags();
				teSetLong(pVarResult, m_param[SB_ViewMode]);
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
								pFV2->SetViewModeAndIconSize(static_cast<FOLDERVIEWMODE>(m_param[SB_ViewMode]), iImageSize);
							}
							pFV2->Release();
							GetViewModeAndIconSize(TRUE);
						}
#ifdef _2000XP
						else {
							m_param[SB_IconSize] = iImageSize;
						}
#endif
					}
				}
				teSetLong(pVarResult, m_param[SB_IconSize]);
				return S_OK;
			//Options
			case 0x10000012:
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
			//SizeFormat
			case 0x10000013:
				if (nArg >= 0) {
					m_param[SB_SizeFormat] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				}
				teSetLong(pVarResult, m_param[SB_SizeFormat]);
				return S_OK;
			//ViewFlags
			case 0x10000016:
				if (nArg >= 0) {
					m_param[SB_ViewFlags] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				}
				teSetLong(pVarResult, m_param[SB_ViewFlags]);
				return S_OK;
			//Id
			case 0x10000017:
				teSetLong(pVarResult, MAX_FV - m_nSB);
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
								teSysFreeString(&m_bsFilter);
								if (lstrlen(v.bstrVal)) {
									m_bsFilter = ::SysAllocString(v.bstrVal);
								}
								Refresh(TRUE);
							}
#endif
						}
					} else if (lstrcmpi(m_bsFilter, v.bstrVal)) {
						teSysFreeString(&m_bsFilter);
						if (lstrlen(v.bstrVal)) {
							m_bsFilter = ::SysAllocString(v.bstrVal);
						}
					}
					VariantClear(&v);
				}
				teSetSZ(pVarResult, m_bsFilter);
				return S_OK;
			//FolderItem
			case 0x10000020:
				teSetObject(pVarResult, m_pFolderItem);
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
						TabCtrl_GetItem(m_pTC->m_hwnd, nIndex, &tcItem);
						int nCount = lstrlen(tcItem.pszText);
						int j = 0;
						WCHAR c = NULL;
						for (int i = 0; i < nCount; i++) {
							bsText[j] = bsText[i];
							if (c != '&' || bsText[i] != '&') {
								j++;
								c = bsText[i];
							} else {
								c = NULL;
							}
						}
						bsText[j] = NULL;
						teSetBSTR(pVarResult, bsText, -1);
					}
				}
				return S_OK;
			//Suspend
			case 0x10000033:
				Suspend(nArg >= 0 ? GetIntFromVariant(&pDispParams->rgvarg[nArg]) : 0);
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
			//DropTarget
			case 0x10000058:
				IDropTarget *pDT;
				pDT = NULL;
				if (!m_pDropTarget || FAILED(QueryInterface(IID_PPV_ARGS(&pDT)))) {
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
			//Columns
			case 0x10000059:
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
				return S_OK;
			//hwndList
			case 0x10000102:
				teSetPtr(pVarResult, m_hwndLV);
				return S_OK;
			//hwndView
			case 0x10000103:
				teSetPtr(pVarResult, m_hwndDV);
				return S_OK;
			//SortColumn
			case 0x10000104:
				if (nArg >= 0) {
					if (pDispParams->rgvarg[nArg].vt == VT_BSTR) {
						SetSort(pDispParams->rgvarg[nArg].bstrVal);
					}
					nFormat = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				}
				if (pVarResult) {
					pVarResult->vt = VT_BSTR;
					GetSort(&pVarResult->bstrVal, nFormat);
				}
				return S_OK;
			//GroupBy
			case 0x10000105:
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
							POINT pt = info.pt;
							IUIAutomationElement *pElement, *pElement2;
							if SUCCEEDED(g_pAutomation->ElementFromPoint(pt, &pElement)) {
								if SUCCEEDED(pElement->GetCurrentPropertyValue(g_PID_ItemIndex, pVarResult)) {
									pVarResult->lVal--;
								}
								if (pVarResult->lVal < 0) {
									IUIAutomationTreeWalker *pWalker = NULL;
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
			//Item
			case 0x10000111:
				if (pVarResult && m_pShellView && nArg >= 0) {
					IFolderView *pFV;
					if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
						LPITEMIDLIST pidl;
						if SUCCEEDED(pFV->Item(GetIntFromVariant(&pDispParams->rgvarg[nArg]), &pidl)) {
							LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidl);
							teSetIDList(pVarResult, pidlFull);
							CoTaskMemFree(pidlFull);
							CoTaskMemFree(pidl);
						}
						pFV->Release();
					}
				}
				return S_OK;
			//Refresh
			case 0x10000206:
				if (!m_bVisible && nArg >= 0 && GetIntFromVariant(&pDispParams->rgvarg[nArg])) {
					m_bRefreshLator = TRUE;
					return S_OK;
				}
				Refresh(TRUE);
				return S_OK;
			//Refresh2Ex
			case 0x20000206:
				if (nArg >= 0) {
					FolderItem *pid = NULL;
					GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
					if (pid && m_pFolderItem) {
						m_pFolderItem->Release();
						m_pFolderItem = pid;
						if (ILIsEqual(m_pidl, g_pidlResultsFolder)) {
							teILFreeClear(&m_pidl);
							teGetIDListFromObject(m_pFolderItem, &m_pidl);
						}
					}
					Refresh(FALSE);
				}
				return S_OK;
			//ViewMenu
			case 0x10000207:
				IContextMenu *pCM;
				CteContextMenu *pdispCM;
				pCM = NULL;
				pdispCM = NULL;
				try {
					if SUCCEEDED(m_pShellView->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&pCM))) {
						CteContextMenu *pdispCM = new CteContextMenu(pCM, NULL, static_cast<IShellBrowser *>(this));
						teSetObject(pVarResult, pdispCM);
					}
				} catch(...) {}
				teReleaseClear(&pdispCM);
				teReleaseClear(&pCM);
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
				teSetLong(pVarResult, hr);
				return S_OK;		
			case 0x10000209:// GetItemPosition
				hr = E_NOTIMPL;
				if (m_pShellView) {
					if (nArg >= 1) {
						VARIANT vMem;
						VariantInit(&vMem);
						LPPOINT ppt = (LPPOINT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1], &vMem);
						if (ppt) {
							IFolderView *pFV;
							if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
								LPITEMIDLIST pidl;
								if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
									hr = pFV->GetItemPosition(ILFindLastID(pidl), ppt);
									teCoTaskMemFree(pidl);
								}
								pFV->Release();
							}
						}
						teWriteBack(&pDispParams->rgvarg[nArg - 1], &vMem);
						VariantClear(&vMem);
					}
				}
				teSetLong(pVarResult, hr);
				return S_OK;
			//SelectAndPositionItem
			case 0x1000020A:
				hr = E_NOTIMPL;
				if (m_pShellView) {
					if (nArg >= 2) {
						VARIANT vMem;
						VariantInit(&vMem);
						LPPOINT ppt = (LPPOINT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 2], &vMem);
						if (ppt) {
							LPITEMIDLIST pidl = NULL;
							teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg]);
							UINT uFlags = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
							IShellView2 *pSV2;
							if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSV2))) {
								hr = pSV2->SelectAndPositionItem(ILFindLastID(pidl), uFlags, ppt);
								pSV2->Release();
							}
							teCoTaskMemFree(pidl);
						}
						VariantClear(&vMem);
					}
				}
				teSetLong(pVarResult, hr);
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
			//GetItemRect
			case 0x10000283:
				hr = E_NOTIMPL;
				if (m_pShellView) {
					if (nArg >= 1) {
						VARIANT vMem;
						VariantInit(&vMem);
						LPRECT prc = (LPRECT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1], &vMem);
						if (prc) {
							IFolderView *pFV;
							if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) {
								LPITEMIDLIST pidl;
								if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
									LPITEMIDLIST pidlChild = ILFindLastID(pidl);
									int i = 0;
									pFV->ItemCount(SVGIO_ALLVIEW, &i);
									while (i-- > 0) {
										LPITEMIDLIST pidlPart;
										if SUCCEEDED(pFV->Item(i, &pidlPart)) {
											if (ILIsEqual(pidlChild, pidlPart)) {
												hr = ListView_GetItemRect(m_hwndLV, i, prc, nArg >= 2 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]) : LVIR_BOUNDS) ? S_OK : E_FAIL;
												i = 0;
											}
											teCoTaskMemFree(pidlPart);
										}
									}
									teCoTaskMemFree(pidl);
								}
								pFV->Release();
							}
						}
						teWriteBack(&pDispParams->rgvarg[nArg - 1], &vMem);
						VariantClear(&vMem);
					}
				}
				teSetLong(pVarResult, hr);
				return S_OK;
			//Notify
			case 0x10000300:
				if (nArg >= 2) {
					long lEvent = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					for (int i = 1; i < (lEvent & (SHCNE_RENAMEITEM | SHCNE_RENAMEFOLDER) ? 3 : 2); i++) {
						LPITEMIDLIST pidl;
						if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg - i])) {
							if (::ILIsParent(m_pidl, pidl, false)) {
								int n = ILGetCount(pidl) - ILGetCount(m_pidl);
								while (n > 1) {
									ILRemoveLastID(pidl);
									n--;
								}
								if (n == 1) {
									WIN32_FIND_DATA wfd;
									teSHGetDataFromIDList(m_pSF2, ILFindLastID(pidl), SHGDFIL_FINDDATA, &wfd, sizeof(WIN32_FIND_DATA));
									if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
										if (nArg >= 1 && GetIntFromVariant(&pDispParams->rgvarg[nArg - 1])) {
											SetFolderSize(ILFindLastID(pidl), wfd.cAlternateFileName, 13);
										} else {
											if SUCCEEDED(teDelProperty(m_pDispatch[SB_TotalFileSize], wfd.cFileName) && m_hwndLV) {
												InvalidateRect(m_hwndLV, NULL, FALSE);
											}
										}
									}
								}
							}
							teCoTaskMemFree(pidl);
						}
					}
				}
				return S_OK;
			//AddItem
			case 0x10000501:
				if (nArg >= 0) {
					if (nArg >= 1) {
						DWORD dwSessionId = GetIntFromVariant(&pDispParams->rgvarg[0]);
						if (dwSessionId != m_dwSessionId) {
							return S_FALSE;
						}
					}
					IFolderView *pFV = NULL;
					LPITEMIDLIST pidl;
					LPITEMIDLIST pidlChild = NULL;
					if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
						if (!ILIsEmpty(pidl)) {
							HRESULT hr = S_FALSE;
							BSTR bs;
							if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_NORMAL)) {
								if (bs[0]) {
									if (m_pDispatch[SB_OnIncludeObject]) {
										VARIANT vResult;
										VariantInit(&vResult);
										VARIANTARG *pv = GetNewVARIANT(3);
										teSetObject(&pv[2], this);
										teSetSZ(&pv[1], bs);
										if (GetDisplayNameFromPidl(&pv[0].bstrVal, pidl, SHGDN_INFOLDER | SHGDN_FORPARSING) == S_OK) {
											pv[0].vt = VT_BSTR;
										}
										Invoke4(m_pDispatch[SB_OnIncludeObject], &vResult, 3, pv);
										hr = GetIntFromVariantClear(&vResult);
									}
									else if (m_bsFilter) {
										if (tePathMatchSpec(bs, m_bsFilter)) {
											hr = S_OK;
										} else {
											BSTR bs2;
											if (GetDisplayNameFromPidl(&bs2, pidl, SHGDN_INFOLDER | SHGDN_FORPARSING) == S_OK) {
												if (tePathMatchSpec(bs2, m_bsFilter)) {
													hr = S_OK;
												}
												::SysFreeString(bs2);
											}
										}
									} else {
										hr = S_OK;
									}
								}
								::SysFreeString(bs);
							}
							if (hr == S_OK) {
								IResultsFolder *pRF;
								if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) && SUCCEEDED(pFV->GetFolder(IID_PPV_ARGS(&pRF)))) {
									pRF->AddIDList(pidl, &pidlChild);
									pRF->Release();
								}
#ifdef _2000XP
								else if (ILIsEqual(m_pidl, g_pidlResultsFolder)) {
									IShellFolderView *pSFV;
									if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
										pidlChild = teILCreateResultsXP(pidl);
										UINT ui;
										pSFV->AddObject(pidlChild, &ui);
										pSFV->Release();
									}
								}
#endif
							}
							teReleaseClear(&pFV);
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
					if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
						IFolderView *pFV = NULL;
						IResultsFolder *pRF;
						if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) && SUCCEEDED(pFV->GetFolder(IID_PPV_ARGS(&pRF)))) {
							hr = pRF->RemoveIDList(pidl);
							pRF->Release();
						}
#ifdef _2000XP
						else if (ILIsEqual(m_pidl, g_pidlResultsFolder)) {
							IShellFolderView *pSFV;
							if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
								BSTR bs;
								if SUCCEEDED(GetDisplayNameFromPidl(&bs, pidl, SHGDN_FORPARSING)) {
									teILCloneReplace(&pidl, teILCreateFromPath(bs));
									LPITEMIDLIST pidlChild = teILCreateResultsXP(pidl);
									UINT ui;
									hr = pSFV->RemoveObject(pidlChild, &ui);
									teCoTaskMemFree(pidlChild);
									::SysFreeString(bs);
								}
								pSFV->Release();
							}
						}
#endif
						teReleaseClear(&pFV);
						teCoTaskMemFree(pidl);
					}
				}
				teSetLong(pVarResult, hr);
				return S_OK;
			//AddItems
			case 0x10000503:
				if (nArg >= 0) {
					if (GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
						TEAddItems *pAddItems = new TEAddItems[1];
						CoMarshalInterThreadInterfaceInStream(IID_IDispatch, static_cast<IDispatch *>(this), &pAddItems->pStrmSB);
						CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pdisp, &pAddItems->pStrmArray);
						pdisp->Release();
						pAddItems->pv = GetNewVARIANT(2);
						if (nArg >= 1) {
							VariantCopy(&pAddItems->pv[0], &pDispParams->rgvarg[0]);
						} else {
							teSetLong(&pAddItems->pv[0], m_dwSessionId);
						}
						_beginthread(&threadAddItems, 0, pAddItems);
					}
				}
				return S_OK;
			//RemoveAll
			case 0x10000504:
				teSetLong(pVarResult, RemoveAll());
				return S_OK;
			//SessionId
			case 0x10000505:
				teSetLong(pVarResult, m_dwSessionId);
				return S_OK;
			//
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
			//DIID_DShellFolderViewEvents
/*///
			case DISPID_EXPLORERWINDOWREADY:
				return E_NOTIMPL;
//*///
			case DISPID_SELECTIONCHANGED://XP+
				SetStatusTextSB(NULL);
				return DoFunc(TE_OnSelectionChanged, this, S_OK);
			case DISPID_FILELISTENUMDONE://XP+
				m_bRefreshing = FALSE;
				if (m_bNavigateComplete) {
					m_bNavigateComplete = FALSE;
					OnNavigationComplete2();
				}
				if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
					WCHAR pszFormat[MAX_STATUS];
					pszFormat[0] = NULL;
					LPWSTR lpBuf = pszFormat;
					if (m_dwUnavailable) {
						LoadString(g_hShell32, 4157, pszFormat, MAX_STATUS);
					} else if (!m_bsFilter && !m_pDispatch[SB_OnIncludeObject]) {
						lpBuf = NULL;
					}
					pFV2->SetText(FVST_EMPTYTEXT, lpBuf);
					pFV2->Release();
				}
				if (!m_hwndLV && m_pExplorerBrowser) {
					SetRedraw(TRUE);
					RedrawWindow(m_hwndDV, NULL, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
				}
				return S_OK;
			case DISPID_VIEWMODECHANGED://XP+
				SetFolderFlags();
				return DoFunc(TE_OnViewModeChanged, this, S_OK);
			case DISPID_BEGINDRAG://XP+
				DoFunc1(TE_OnBeginDrag, this, pVarResult);
				return S_OK;
			case DISPID_COLUMNSCHANGED://XP-
				InitFolderSize();
				return DoFunc(TE_OnColumnsChanged, this, S_OK);
			case DISPID_SORTDONE://XP-
				if (m_nFocusItem < 0) {
					FocusItem(FALSE);
				}
				return S_OK;
			case DISPID_INITIALENUMERATIONDONE://XP-
				SetFolderFlags();
				return S_OK;
/*			case DISPID_CONTENTSCHANGED://XP-
				return S_OK;*/
			default:
				if (dispIdMember >= START_OnFunc && dispIdMember < START_OnFunc + Count_SBFunc) {
					if (wFlags & DISPATCH_METHOD) {
						if (m_pDispatch[dispIdMember - START_OnFunc]) {
							Invoke5(m_pDispatch[dispIdMember - START_OnFunc], DISPID_VALUE, wFlags, pVarResult, - nArg - 1, pDispParams->rgvarg);
						}
						return S_OK;
					}
					if (nArg >= 0) {
						teReleaseClear(&m_pDispatch[dispIdMember - START_OnFunc]);
						IUnknown *punk;
						if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
							punk->QueryInterface(IID_PPV_ARGS(&m_pDispatch[dispIdMember - START_OnFunc]));
						}
					}
					if (m_pDispatch[dispIdMember - START_OnFunc]) {
						teSetObject(pVarResult, m_pDispatch[dispIdMember - START_OnFunc]);
					}
					return S_OK;
				}
		}
		if (m_pDSFV) {
			return m_pDSFV->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
	} catch (...) {
		return teException(pExcepInfo, L"FolderView", methodSB, dispIdMember);
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
	return m_pTC->QueryInterface(IID_PPV_ARGS(ppid));
}

STDMETHODIMP CteShellBrowser::get_Folder(Folder **ppid)
{
	HRESULT hr = E_NOTIMPL;
	IShellFolderViewDual *pSFVD;
	if SUCCEEDED(m_pDSFV->QueryInterface(IID_PPV_ARGS(&pSFVD))) {
		hr = pSFVD->get_Folder(ppid);
		pSFVD->Release();
	}
	return SUCCEEDED(hr) ? hr : GetFolderObjFromIDList(m_pidl, ppid);
}

STDMETHODIMP CteShellBrowser::SelectedItems(FolderItems **ppid)
{
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
					} else {
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
	teReleaseClear(&pDataObj);
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
				if SUCCEEDED(teGetDisplayNameBSTR(m_pSF2, pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING, &v.bstrVal)) {
					v.vt = VT_BSTR;
				}
			} else {
				FolderItem *pid;
				LPITEMIDLIST pidlFull = ILCombine(m_pidl, pidl);
				if (GetFolderItemFromPidl(&pid, pidlFull)) {
					teSetObject(&v, pid);
					pid->Release();
				}
				teCoTaskMemFree(pidlFull);
			}
			if (v.vt != VT_EMPTY) {
				pFolderItems->ItemEx(-1, NULL, &v);
				VariantClear(&v);
			}
		}
	} catch (...) {}
}

int CteShellBrowser::PSGetColumnIndexXP(LPWSTR pszName, int *pcxChar)
{
	SHELLDETAILS sd;
	BSTR bs;
	for (UINT i = 0; i < m_nColumns && GetDetailsOf(NULL, i, &sd) == S_OK; i++) {
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
			return tePSGetNameFromPropertyKeyEx(scid, nFormat, NULL);
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

HRESULT CteShellBrowser::NavigateSB(IShellView *pPreviousView, FolderItem *pPrevious)
{
	HRESULT hr = E_FAIL;
	int nCreate = 4;
	while ((nCreate -= 2) > 0) {
		if (GetShellFolder2(m_pidl) == S_OK) {
#ifdef _2000XP
			m_nFolderName = MAXINT;
			if (m_nColumns) {
				delete [] m_pColumns;
				m_nColumns = 0;
			}
			SHCOLUMNID scid;
			while (m_nColumns < MAX_COLUMNS && m_pSF2->MapColumnToSCID(m_nColumns, &scid) == S_OK) {
				m_nColumns++;
			}
			m_nColumns += 2;
			m_pColumns = new TEColumn[m_nColumns];
			BOOL bIsResultsFolder = ILIsEqual(m_pidl, g_pidlResultsFolder);
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
#endif
			try {
				hr = CreateViewWindowEx(pPreviousView);
			} catch (...) {
				hr = E_FAIL;
			}
		}
		if (hr != S_OK) {
			teILCloneReplace(&m_pidl, g_pidlResultsFolder);
			m_dwUnavailable = GetTickCount();
			nCreate++;
			hr = S_FALSE;
		}
	}
	return hr;
}

HRESULT CteShellBrowser::CreateViewWindowEx(IShellView *pPreviousView)
{
	HRESULT hr = E_FAIL;
	m_pShellView = NULL;
	RECT rc;
	if (m_pSF2) {
#ifdef _2000XP
		if SUCCEEDED(CreateViewObject(m_pTC->m_hwndStatic, IID_PPV_ARGS(&m_pShellView)) && m_pShellView) {
			if (!g_bUpperVista && m_param[SB_ViewMode] == FVM_ICON && m_param[SB_IconSize] >= 96) {
				m_param[SB_ViewMode] = FVM_THUMBNAIL;
			}
#else
		if SUCCEEDED(m_pSF2->CreateViewObject(m_pTC->m_hwndStatic, IID_PPV_ARGS(&m_pShellView)) && m_pShellView) {
#endif
			try {
				if (::InterlockedIncrement(&m_nCreate) <= 1) {
					m_hwnd = NULL;
					IEnumIDList *peidl = NULL;
					hr = ILIsEqual(m_pidl, g_pidlResultsFolder) ? S_OK : m_pSF2->EnumObjects(g_hwndMain, SHCONTF_NONFOLDERS | SHCONTF_FOLDERS, &peidl);
					if (hr == S_OK) {
#ifdef _2000XP
						FOLDERSETTINGS fs = { g_bUpperVista ? FVM_AUTO : m_param[SB_ViewMode], (m_param[SB_FolderFlags] | FWF_USESEARCHFOLDER) & ~FWF_NOENUMREFRESH };
#else
						FOLDERSETTINGS fs = { FVM_AUTO, (m_param[SB_FolderFlags] | FWF_USESEARCHFOLDER) & ~FWF_NOENUMREFRESH };
#endif
						hr = m_pShellView->CreateViewWindow(pPreviousView, &fs, static_cast<IShellBrowser *>(this), &rc, &m_hwnd);
						teReleaseClear(&peidl);
					} else {
						hr = E_FAIL;
					}
				} else {
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
	if (pvfi->vt == VT_I4) {
		HRESULT hr = E_FAIL;
		IFolderView *pFV;
		if (m_pShellView && SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV)))) {
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
			hr = pFV->SelectItem(pvfi->lVal, dwFlags);
			pFV->Release();
			if (hr == S_OK) {
				return hr;
			}
		}
	}
	teGetIDListFromVariant(&pidl, pvfi);
	return SelectItemEx(&pidl, dwFlags, TRUE);
}

HRESULT CteShellBrowser::SelectItemEx(LPITEMIDLIST *ppidl, int dwFlags, BOOL bFree)
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
/*///	No longer needed.
		if ((dwFlags & (SVSI_ENSUREVISIBLE | 2)) == SVSI_ENSUREVISIBLE && m_hwndLV) {
			int nFocused = ListView_GetNextItem(m_hwndLV, -1, LVNI_ALL | LVNI_FOCUSED);
			int nTop = ListView_GetTopIndex(m_hwndLV);
			if (nTop == 0 || nFocused < nTop || nFocused >= nTop + ListView_GetCountPerPage(m_hwndLV)) {
				ListView_EnsureVisible(m_hwndLV, 0, FALSE);
				ListView_EnsureVisible(m_hwndLV, nFocused, FALSE);
			}
		}
//*///
	}
	if (bFree) {
		teILFreeClear(ppidl);
		CteFolderItem *pid1 = NULL;
		if SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) {
			teILFreeClear(&pid1->m_pidlFocused);
			pid1->Release();
		}
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
	if (m_pTC) {
		int i;
		TC_ITEM tcItem;
		
		for (i = TabCtrl_GetItemCount(m_pTC->m_hwnd); i-- > 0;) {
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
	LPITEMIDLIST pidlPrevius = m_pidl;
	m_pidl = ::ILClone(pidlFolder);
	FolderItem *pPrevious = m_pFolderItem;
	if (m_pFolderItem1) {
		m_pFolderItem = m_pFolderItem1;
		m_pFolderItem1 = NULL;
	} else {
		GetFolderItemFromPidl(&m_pFolderItem, m_pidl);
	}

	UINT uViewMode = m_param[SB_ViewMode];
	FOLDERVIEWOPTIONS fvo = teGetFolderViewOptions(const_cast<LPITEMIDLIST>(pidlFolder), uViewMode);

	HRESULT hr = OnBeforeNavigate(pPrevious, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
	if FAILED(hr) {
		teCoTaskMemFree(m_pidl);
		m_pidl = pidlPrevius;
		FolderItem *pid = m_pFolderItem;
		m_pFolderItem = NULL;
		if (pPrevious) {
			pPrevious->QueryInterface(IID_PPV_ARGS(&m_pFolderItem));
			pPrevious->Release();
		}
		if (hr == E_ACCESSDENIED) {
			BrowseObject2(pid, SBSP_NEWBROWSER | SBSP_ABSOLUTE);
		}
		pid ->Release();
		m_bEnableSuspend = FALSE;
		return hr;
	}
	BSTR bs, bsPrev;
	GetDisplayNameFromPidl(&bs, m_pidl, SHGDN_FORPARSING);
	GetDisplayNameFromPidl(&bsPrev, pidlPrevius, SHGDN_FORPARSING);
	BOOL bDiffReal = !SysStringLen(bs) || lstrcmpi(bs, bsPrev);
	SysFreeString(bsPrev);
	SysFreeString(bs);
	teCoTaskMemFree(pidlPrevius);
	if (uViewMode != m_param[SB_ViewMode] && fvo != teGetFolderViewOptions(const_cast<LPITEMIDLIST>(pidlFolder), m_param[SB_ViewMode])) {
		m_bCheckLayout = TRUE;
		return S_OK;
	}
	if (bDiffReal) {
		teSysFreeString(&m_bsFilter);
		teReleaseClear(&m_pDispatch[SB_OnIncludeObject]);
	}
	ResetPropEx();
	//History / Management
	SetHistory(NULL, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
	teReleaseClear(&pPrevious);
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
	m_bNavigateComplete = TRUE;
	m_dwSessionId = (rand() * 0x20000) + (GetTickCount() & 0x1ffff);
	SetPropEx();
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
				g_bCharWidth = false;
			}
		}
	}
#endif
	SetTabName();
	return S_OK;
}

VOID CteShellBrowser::SetTabName()
{
	//View Tab Name
	int i;
	i = GetTabIndex();
	if (i >= 0) {
		VARIANT vResult;
		teSetLong(&vResult, S_FALSE);
		if (DoFunc(TE_OnViewCreated, this, E_NOTIMPL) != S_OK) {
			BSTR bstr;
			if (m_pFolderItem && SUCCEEDED(m_pFolderItem->get_Name(&bstr))) {
				SetTitle(bstr, i);
				::SysFreeString(bstr);
			}
		}
	}
	SetStatusTextSB(NULL);
}

STDMETHODIMP CteShellBrowser::OnNavigationComplete(PCIDLIST_ABSOLUTE pidlFolder)
{
	return S_OK;
}

VOID CteShellBrowser::OnNavigationComplete2()
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
	m_nFocusItem = 1;
	DoFunc(TE_OnNavigateComplete, this, E_NOTIMPL);
	if (m_nFocusItem > 0) {
		FocusItem(FALSE);
	}
	m_bCheckLayout = TRUE;
	SetFolderFlags();
	InitFolderSize();
	if (m_pTC->m_bRedraw) {
		SetTimer(g_hwndMain, TET_Redraw, 1, teTimerProc);
	}
}

HRESULT CteShellBrowser::BrowseToObject()
{
	m_bEnableSuspend = TRUE;
	HRESULT hr = m_pExplorerBrowser->BrowseToObject(m_pSF2, SBSP_ABSOLUTE);
	if SUCCEEDED(hr) {
		IOleWindow *pOleWindow;
		if SUCCEEDED(m_pExplorerBrowser->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
			pOleWindow->GetWindow(&m_hwnd);
			pOleWindow->Release();
		}
		if (!m_pShellView) {
			m_pExplorerBrowser->GetCurrentView(IID_PPV_ARGS(&m_pShellView));
		}
	}
	return hr;
}

HRESULT CteShellBrowser::GetShellFolder2(LPITEMIDLIST pidl)
{
	HRESULT hr = E_FAIL;
	IShellFolder *pSF = NULL;
	if (m_bsFilter && g_pidlLibrary && ILIsParent(g_pidlLibrary, pidl, FALSE)) {
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
		if (!GetShellFolder(&pSF, pidl)) {
			GetShellFolder(&pSF, g_pidlResultsFolder);
		}
	}
	if (pSF) {
		teReleaseClear(&m_pSF2);
		hr = pSF->QueryInterface(IID_PPV_ARGS(&m_pSF2));
		pSF->Release();
	}
	return hr;
}

VOID CteShellBrowser::SetLVSettings()
{
	if (m_hwndLV) {
		if (m_param[SB_ViewFlags] & CDB2GVF_NOSELECTVERB) {
			ListView_SetSelectedColumn(m_hwndLV, -1);
		}
		ListView_SetBkColor(m_hwndLV, m_clrBk);
		ListView_SetTextBkColor(m_hwndLV, m_clrTextBk);
		ListView_SetTextColor(m_hwndLV, m_clrText);
//		TreeView_SetTextColor(m_pTV->m_hwndTV, 0xff);
	}
}

HRESULT CteShellBrowser::RemoveAll()
{
	HRESULT hr = E_NOTIMPL;
	IFolderView *pFV = NULL;
	IResultsFolder *pRF;
	if (SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV))) && SUCCEEDED(pFV->GetFolder(IID_PPV_ARGS(&pRF)))) {
		hr = pRF->RemoveAll();
		pRF->Release();
	}
#ifdef _2000XP
	else {
		IShellFolderView *pSFV;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
			UINT ui;
			pSFV->GetObjectCount(&ui);
			int i = ui;
			pSFV->SetRedraw(FALSE);
			while (i-- > 0) {
				LPITEMIDLIST pidl;
				pSFV->GetObjectW(&pidl, i);
				hr = pSFV->RemoveObject(pidl, &ui);
			}
			pSFV->SetRedraw(TRUE);
			pSFV->Release();
		}
	}
#endif
	teReleaseClear(&pFV);
	return hr;
}

STDMETHODIMP CteShellBrowser::OnNavigationFailed(PCIDLIST_ABSOLUTE pidlFolder)
{
	if (m_bEnableSuspend) {
		Suspend(2);
	}
	return S_OK;
}

//IExplorerPaneVisibility
STDMETHODIMP CteShellBrowser::GetPaneState(REFEXPLORERPANE ep, EXPLORERPANESTATE *peps)
{
	if (g_pOnFunc[TE_OnGetPaneState]) {
		VARIANTARG *pv = GetNewVARIANT(3);
		CteMemory *pstEps;
		pstEps = new CteMemory(sizeof(int), peps, 1, L"DWORD");
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
		teReleaseClear(&m_pSFVCB);
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
			} else {
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
				SetFolderSize(pidl, szBuf, MAX_COLUMN_NAME_LEN);
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
		psd->str.pOleStr = (LPOLESTR)CoTaskMemAlloc((lstrlen(szBuf) + 1) * sizeof(WCHAR));
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
	} else {
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
		} else if (pt.x > rc.right - DRAG_SCROLL) {
			SendMessage(m_hwndLV, WM_HSCROLL, SB_LINERIGHT, 0);
		}
		if (pt.y < DRAG_SCROLL) {
			SendMessage(m_hwndLV, WM_VSCROLL, SB_LINEUP, 0);
		} else if (pt.y > rc.bottom - DRAG_SCROLL) {
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

VOID CteShellBrowser::Suspend(int nMode)
{
	BOOL bVisible = m_bVisible;
	if (nMode == 2) {
		if (m_dwUnavailable || ILIsEqual(m_pidl, g_pidlResultsFolder)) {
			return;
		}
		Show(FALSE, 0);
		m_dwUnavailable = GetTickCount();
		teILCloneReplace(&m_pidl, g_pidlResultsFolder);
		VARIANT v;
		if (GetVarPathFromFolderItem(m_pFolderItem, &v)) {
			teReleaseClear(&m_pFolderItem);
			CteFolderItem *pid = new CteFolderItem(&v);
			pid->m_dwUnavailable = m_dwUnavailable;
			pid->m_pidl = ILClone(g_pidlResultsFolder);
			m_pFolderItem = pid;
			VariantClear(&v);
		}
	}
	if (!m_dwUnavailable) {
		SaveFocusedItemToHistory();
		teDoCommand(this, m_hwnd, WM_NULL, 0, 0);//Save folder setings
	}
	DestroyView(0);
	if (nMode && m_pTV) {
		m_pTV->Close();
		m_pTV->m_bSetRoot = true;
	}
	CteFolderItem *pid1 = NULL;
	if SUCCEEDED(m_pFolderItem->QueryInterface(g_ClsIdFI, (LPVOID *)&pid1)) {
		if (pid1->m_v.vt == VT_BSTR) {
			teILFreeClear(&m_pidl);
			pid1->Release();
		}
	}
	if (m_pidl) {
		VARIANT v;
		if (GetVarPathFromFolderItem(m_pFolderItem, &v)) {
			teReleaseClear(&m_pFolderItem);
			CteFolderItem *pid = new CteFolderItem(&v);
			pid->m_dwUnavailable = m_dwUnavailable;
			m_pFolderItem = pid;
			teILFreeClear(&m_pidl);
			VariantClear(&v);
		}
	}
	m_nUnload = 9;
	Show(bVisible, 0);
}

VOID CteShellBrowser::SetPropEx()
{
	if (m_pShellView->GetWindow(&m_hwndDV) == S_OK) {
		if (!m_DefProc) {
			SetWindowLongPtr(m_hwndDV, GWLP_USERDATA, (LONG_PTR)this);
			m_DefProc = SetWindowLongPtr(m_hwndDV, GWLP_WNDPROC, (LONG_PTR)TELVProc);
			for (int i = WM_USER + 173; i <= WM_USER + 175; i++) {
				teChangeWindowMessageFilterEx(m_hwndDV, i, MSGFLT_ALLOW, NULL);
			}
		}
		m_hwndLV = FindWindowEx(m_hwndDV, 0, WC_LISTVIEW, NULL);
		m_hwndDT = m_hwndLV;
		if (m_hwndLV) {
			if (m_pExplorerBrowser) {
				if (m_param[SB_FolderFlags] & FWF_NOCLIENTEDGE) {
					SetWindowLong(m_hwnd, GWL_STYLE, GetWindowLong(m_hwnd, GWL_STYLE) & ~WS_BORDER);
				}
			} else if (!(m_param[SB_FolderFlags] & FWF_NOCLIENTEDGE)) {
				SetWindowLong(m_hwndLV, GWL_EXSTYLE, GetWindowLong(m_hwndLV, GWL_EXSTYLE) | WS_EX_CLIENTEDGE);
			}
		} else {
			m_hwndDT = FindWindowEx(m_hwndDV, NULL, L"DirectUIHWND", NULL);
		}
		if (!m_pDropTarget) {
			teRegisterDragDrop(m_hwndDT, this, &m_pDropTarget);
		}
	}
}

VOID CteShellBrowser::ResetPropEx()
{
	if (m_DefProc) {
		SetWindowLongPtr(m_hwndDV, GWLP_WNDPROC, (LONG_PTR)m_DefProc);
		m_DefProc = NULL;
	}
	if (m_pDropTarget) {
		RevokeDragDrop(m_hwndDT);
		RegisterDragDrop(m_hwndDT, m_pDropTarget);
		m_pDropTarget->Release();
		m_pDropTarget = NULL;
	}
}

void CteShellBrowser::Show(BOOL bShow, DWORD dwOptions)
{
	bShow &= 1;
	if (bShow ^ m_bVisible) {
		if (bShow) {
			if (!g_nLockUpdate && !m_nCreate) {
				if (m_nUnload == 9) {
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
				ShowWindow(m_hwnd, SW_SHOW);
				if (m_hwndLV) {
					ShowWindow(m_hwndLV, SW_SHOW);
				}
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
			} else {
				if ((dwOptions & 2) == 0) {
					SetRedraw(FALSE);
					m_pShellView->UIActivate(SVUIA_DEACTIVATE);
					MoveWindow(m_hwnd, -1, -1, 0, 0, FALSE);
					ShowWindow(m_hwnd, SW_HIDE);
					if (m_hwndLV) {
						ShowWindow(m_hwndLV, SW_HIDE);
					}
					if (m_pTV->m_hwnd) {
						MoveWindow(m_pTV->m_hwnd, -1, -1, 0, 0, FALSE);
						ShowWindow(m_pTV->m_hwnd, SW_HIDE);
					}
					if ((dwOptions & 1) && m_dwUnavailable && !m_nCreate && !ILIsEqual(m_pidl, g_pidlResultsFolder)) {
						Suspend(0);
					}
				}
			}
		} else {
			m_bVisible = FALSE;
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
			SendMessage(m_pTC->m_hwnd, TCM_DELETEITEM, i, 0);
		}
		ShowWindow(m_pTV->m_hwnd, SW_HIDE);
		Clear();
	}
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
			m_pServiceProvider->Release();
			m_pServiceProvider = NULL;
			Show(FALSE, 0);
			m_pExplorerBrowser->Destroy();
			m_pExplorerBrowser->Release();
		} catch (...) {
			g_nException = 0;
		}
		m_pExplorerBrowser = NULL;
	}
	if (m_pShellView) {
		try {
			if (bCloseSB && (nFlags & 1) == 0) {
				Show(FALSE, 0);
				m_pShellView->DestroyViewWindow();
			}
			if (nFlags == 0) {
				teReleaseClear(&m_pShellView);
			}
		} catch (...) {
			g_nException = 0;
		}
	}
}

VOID CteShellBrowser::GetSort(BSTR* pbs, int nFormat)
{
	*pbs = NULL;
	if (!m_pShellView) {
		return;
	}
	IFolderView2 *pFV2;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		SORTCOLUMN col;
		if SUCCEEDED(pFV2->GetSortColumns(&col, 1)) {
			*pbs = tePSGetNameFromPropertyKeyEx(col.propkey, nFormat, m_pShellView);
			if (col.direction < 0) {
				ToMinus(pbs);
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
		*pbs = tePSGetNameFromPropertyKeyEx(scid, nFormat, NULL);
		HWND hHeader = ListView_GetHeader(m_hwndLV);
		if (hHeader) {
			int nCount = Header_GetItemCount(hHeader);
			for (int i = 0; i < nCount; i++) {
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
			*pbs = tePSGetNameFromPropertyKeyEx(propKey, 1, NULL);
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
			for (int i = 0; i < nCount; i++) {
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

HRESULT CteShellBrowser::PropertyKeyFromName(BSTR bs, PROPERTYKEY *pkey)
{
	HRESULT hr = tePSPropertyKeyFromStringEx(bs, pkey);
	if SUCCEEDED(hr) {
		return hr;
	}
	IColumnManager *pColumnManager;
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pColumnManager))) {
		UINT uCount;
		if SUCCEEDED(pColumnManager->GetColumnCount(CM_ENUM_VISIBLE, &uCount)) {
			if (uCount) {
				PROPERTYKEY *pPropKey = new PROPERTYKEY[uCount];
				if SUCCEEDED(pColumnManager->GetColumns(CM_ENUM_VISIBLE, pPropKey, uCount)) {
					CM_COLUMNINFO cmci = { sizeof(CM_COLUMNINFO), CM_MASK_NAME };
					for (UINT i = 0; i < uCount; i++) {
						if SUCCEEDED(pColumnManager->GetColumnInfo(pPropKey[i], &cmci)) {
							if (lstrcmpi(bs, cmci.wszName) == 0 || teStrSameIFree(tePSGetNameFromPropertyKeyEx(pPropKey[i], 0, NULL), bs)) {
								*pkey = pPropKey[i];
								hr = S_OK;
								break;
							}						
						}
					}
				}
				delete [] pPropKey;
			}
		}
		pColumnManager->Release();
	}
	return hr;
}

FOLDERVIEWOPTIONS CteShellBrowser::teGetFolderViewOptions(LPITEMIDLIST pidl, UINT uViewMode)
{
	if (m_param[SB_FolderFlags] & FWF_CHECKSELECT) {
		return FVO_VISTALAYOUT;
	}
	UINT uFlag = 1;
	while (--uViewMode > 0) {
		uFlag *= 2;
	}
	return uFlag & g_param[TE_Layout] ? FVO_DEFAULT : FVO_VISTALAYOUT;
}

VOID CteShellBrowser::SetSort(BSTR bs)
{
	IFolderView2 *pFV2;
	SORTCOLUMN col, col2;
	if (lstrlen(bs) < 1) {
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
			SORTCOLUMN *pCol = new SORTCOLUMN[1];
			if SUCCEEDED(pFV2->GetSortColumns(pCol, 1)) {
				col.propkey = PKEY_Contact_Label;
				col.direction = SORT_ASCENDING;
				pFV2->SetSortColumns(&col, 1);
				m_nFocusItem = -1;
				SetTimer(m_hwnd, (UINT_PTR)pCol, 100, teTimerProcSort);
			} else {
				delete [] pCol;
			}
			pFV2->Release();
			return;
		}
#ifdef _2000XP
		IShellFolderView *pSFV;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pSFV))) {
			LPARAM lSort;
			if SUCCEEDED(pSFV->GetArrangeParam(&lSort)) {
				pSFV->Rearrange(lSort);
				pSFV->Rearrange(lSort);
			}
			pSFV->Release();
		}
#endif
		return;
	}
	int dir = (bs[0] == '-') ? 1 : 0;
	LPWSTR szNew = &bs[dir];
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		if SUCCEEDED(PropertyKeyFromName(szNew, &col.propkey)) {
			col.direction = 1 - dir * 2;
			if SUCCEEDED(pFV2->GetSortColumns(&col2, 1)) {
				if (col.direction != col2.direction || col.propkey.pid != col2.propkey.pid || !IsEqualFMTID(col.propkey.fmtid, col2.propkey.fmtid)) {
					m_nFocusItem = -1;
					pFV2->SetSortColumns(&col, 1);
				}
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
	if SUCCEEDED(lpfnPSPropertyKeyFromStringEx(szNew, &propKey)) {
		bsName = tePSGetNameFromPropertyKeyEx(propKey, 0, m_pShellView);
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
	IFolderView2 *pFV2;

	PROPERTYKEY propKey;
	int dir = (bs[0] == '-') ? 1 : 0;
	LPWSTR szNew = &bs[dir];
	if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
		if SUCCEEDED(PropertyKeyFromName(szNew, &propKey)) {
			pFV2->SetGroupBy(propKey, !dir);
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
			for (int i = 0; i < nCount; i++) {
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

HRESULT CteShellBrowser::SetRedraw(BOOL bRedraw)
{
	HRESULT hr = E_FAIL;
	if (m_pShellView) {
		IFolderView2 *pFV2;
		if SUCCEEDED(m_pShellView->QueryInterface(IID_PPV_ARGS(&pFV2))) {
			hr = pFV2->SetRedraw(bRedraw);
			pFV2->Release();
		} else {
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

//CTE

CTE::CTE(int nCmdShow)
{
	m_cRef = 1;
	m_pDragItems = NULL;
	::ZeroMemory(g_param, sizeof(g_param));
	g_param[TE_Type] = CTRL_TE;
	g_param[TE_CmdShow] = nCmdShow;
	g_param[TE_Layout] = 0x80;
	VariantInit(&m_vData);
}

CTE::~CTE()
{
	VariantClear(&m_vData);
}

STDMETHODIMP CTE::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CTE, IDispatch),
		QITABENT(CTE, IDropTarget),
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
					VariantClear(&m_vData);
					VariantCopy(&m_vData, &pDispParams->rgvarg[nArg]);
				}
				if (pVarResult) {
					VariantCopy(pVarResult, &m_vData);
				}
				return S_OK;
			//hwnd
			case 1002:
				teSetPtr(pVarResult, g_hwndMain);
				return S_OK;
			//About
			case 1004:
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
									int nId = GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]);
									if (nId) {
										pSB = g_ppSB[MAX_FV - nId];
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
										pTC = g_ppTC[MAX_TC - nId];
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
					int nCtrl = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					IDispatch *pArray = NULL;
					GetNewArray(&pArray);
					int i;

					switch (nCtrl) {
						case CTRL_FV:
						case CTRL_SB:
						case CTRL_EB:
						case CTRL_TV:
							CteShellBrowser *pSB;
							for (i = MAX_FV; i-- && (pSB = g_ppSB[i]);) {
								if (!pSB->m_bEmpty) {
									if (nCtrl == CTRL_TV) {
										teArrayPush(pArray, pSB->m_pTV);
									} else {
										teArrayPush(pArray, pSB);
									}
								}
							}
							break;
						case CTRL_TC:
							CteTabCtrl *pTC;
							for (i = MAX_TC; i-- && (pTC = g_ppTC[i]);) {
								if (!pTC->m_bEmpty) {
									teArrayPush(pArray, pTC);
								}
							}
							break;
						case CTRL_WB:
							if (g_pWebBrowser) {
								teArrayPush(pArray, g_pWebBrowser);
							}
							break;
					}//end_switch
					teSetObjectRelease(pVarResult, new CteDispatchEx(pArray));
					pArray->Release();
				}
				return S_OK;
			//Version
			case 1007:
				teSetLong(pVarResult, 20000000 + VER_Y * 10000 + VER_M * 100 + VER_D);
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
					g_nSize += MAXWORD;
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
				teSetObject(pVarResult, g_pAPI);
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
				GetNewObject(&pdisp);
				teSetObjectRelease(pVarResult, pdisp);
				return S_OK;
			//Array
			case TE_METHOD + 1135:
				GetNewArray(&pdisp);
				teSetObjectRelease(pVarResult, pdisp);
				return S_OK;
			//Collection
			case TE_METHOD + 1136:
				if (nArg >= 0 && GetDispatch(&pDispParams->rgvarg[nArg], &pdisp)) {
					teSetObjectRelease(pVarResult, new CteDispatch(pdisp, 1));
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
				if (nArg >= 0) {
					switch (GetIntFromVariant(&pDispParams->rgvarg[nArg])) {
						case CTRL_TC:
							if (nArg >= 4) {
								CteTabCtrl *pTC;
								pTC = GetNewTC();
								pTC->Init();
								int i = _countof(pTC->m_param) - 1;
								for (i = nArg < i ? nArg : i; i >= 0; i--) {
									pTC->m_param[i] = GetIntFromVariantPP(&pDispParams->rgvarg[nArg - i], i);
								}
								if (pTC->Create()) {
									teSetObject(pVarResult, pTC);
									pTC->Show(TRUE, !g_pTC);
								} else {
									pTC->Release();
								}
							}
							break;
						case CTRL_SW:
							if (nArg >= 5) {
								VARIANT v;
								teVariantChangeType(&v, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
								RECT rc, rcWindow;
								HWND hwnd1 = g_hwndMain;
								if (FindUnknown(&pDispParams->rgvarg[nArg - 3], &punk)) {
									IUnknown_GetWindow(punk, &hwnd1);
								}
								GetWindowRect(hwnd1, &rc);
								HWND hwndParent = hwnd1;
								while (hwnd1 = GetParent(hwndParent)) {
									hwndParent = hwnd1;
								}
								GetWindowRect(hwndParent, &rcWindow);
								int a = rcWindow.right - rc.right + rc.left - rcWindow.left;
								int b = rcWindow.bottom - rc.bottom + rc.top - rcWindow.top;
								int w = GetIntFromVariant(&pDispParams->rgvarg[nArg - 4]) + a;
								int h = GetIntFromVariant(&pDispParams->rgvarg[nArg - 5]) + b;
								int x = rc.left + (rc.right - rc.left - w) / 2;
								int y = rc.top + (rc.bottom - rc.top - h) / 2;
								if (w == rcWindow.right - rcWindow.left) {
									x += a;
								}
								if (h == rcWindow.bottom - rcWindow.top) {
									y += b / 2;
								}
								HMONITOR hMonitor = MonitorFromRect(&rcWindow, MONITOR_DEFAULTTONEAREST);
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
								MyRegisterClass(hInst, WINDOW_CLASS2, WndProc2);
								HWND hwnd = CreateWindow(WINDOW_CLASS2, g_szTE, WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
									x, y, w, h, hwndParent, NULL, hInst, NULL);
								CteWebBrowser *pWB = new CteWebBrowser(hwnd, v.bstrVal, &pDispParams->rgvarg[nArg - 2]);
								VariantClear(&v);
								teSetObject(&v, pWB);
								tePutPropertyAtLLX(g_pSubWindows, (LONGLONG)hwnd, &v);						
								VariantClear(&v);
								teSetObjectRelease(pVarResult, pWB);
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
								pSB->Close(true);
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
					for (int i = MAX_TC; i-- && g_ppTC[i];) {
						CteTabCtrl *pTC = g_ppTC[i];
						if (pTC->m_bVisible) {
							CteShellBrowser *pSB = pTC->GetShellBrowser(pTC->m_nIndex);
							if (pSB && !pSB->m_bEmpty) {
								if (pSB->m_nUnload & 5) {
									pSB->Show(TRUE, 0);
								}
								if (pTC->m_bRedraw) {
									pTC->RedrawUpdate();
								}
							}
						}
					}
					if (g_nSize >= MAXWORD) {
						g_nSize -= MAXWORD;
					}
					ArrangeWindow();
				}
				return S_OK;
			//HookDragDrop//Deprecated
			case TE_METHOD + 1100:
				return S_OK;
			//Value
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
			default:
				if (dispIdMember >= START_OnFunc && dispIdMember < START_OnFunc + Count_OnFunc) {
					if (wFlags & DISPATCH_METHOD) {
						if (g_pOnFunc[dispIdMember - START_OnFunc]) {
							Invoke5(g_pOnFunc[dispIdMember - START_OnFunc], DISPID_VALUE, wFlags, pVarResult, - nArg - 1, pDispParams->rgvarg);
						}
						return S_OK;
					}
					if (nArg >= 0) {
						teReleaseClear(&g_pOnFunc[dispIdMember - START_OnFunc]);
						if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
							punk->QueryInterface(IID_PPV_ARGS(&g_pOnFunc[dispIdMember - START_OnFunc]));
							if (dispIdMember == TE_Labels + START_OnFunc && g_pOnFunc[dispIdMember - START_OnFunc]) {
								VARIANT v;
								VariantInit(&v);
								teGetProperty(g_pOnFunc[TE_Labels], L"call", &v);
								g_bLabelsMode = v.vt != VT_DISPATCH;
								VariantClear(&v);
							}
						}
					}
					if (g_pOnFunc[dispIdMember - START_OnFunc]) {
						teSetObject(pVarResult, g_pOnFunc[dispIdMember - START_OnFunc]);
					}
					return S_OK;
				}
				//Type, offsetLeft etc.
				else if (dispIdMember >= TE_OFFSET && dispIdMember < TE_OFFSET + Count_TE_params) {
					if (nArg >= 0 && dispIdMember != TE_OFFSET) {
						g_param[dispIdMember - TE_OFFSET] = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
						if (dispIdMember >= TE_OFFSET + TE_Left && dispIdMember <= TE_Bottom + TE_OFFSET) {
							ArrangeWindow();
						}
					}
					teSetLong(pVarResult, g_param[dispIdMember - TE_OFFSET]);
				}
				return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, L"external", methodTE, dispIdMember);
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
	if (fEscapePressed || g_bDropFinished || (grfKeyState & (MK_LBUTTON | MK_RBUTTON)) == (MK_LBUTTON | MK_RBUTTON)) {
		return DRAGDROP_S_CANCEL;
	}
	if ((grfKeyState & (MK_LBUTTON | MK_RBUTTON)) == 0) {
		g_bDropFinished = TRUE;
		if (g_pOnFunc[TE_OnBeforeGetData]) {
			VARIANTARG *pv = GetNewVARIANT(3);
			teSetObject(&pv[2], g_pDraggingCtrl);
			teSetObjectRelease(&pv[1], new CteFolderItems(g_pDraggingItems, NULL, FALSE));
			teSetLong(&pv[0], 2);
			Invoke4(g_pOnFunc[TE_OnBeforeGetData], NULL, 3, pv);
		}
		return DRAGDROP_S_DROP;
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
	m_pExternal = NULL;
	VariantInit(&m_vData);
	if (pvArg) {
		VariantCopy(&m_vData, pvArg);
		VARIANT v;
		VariantInit(&v);
		GetNewObject(&m_pExternal);
		teSetObject(&v, g_pTE);
		tePutProperty(m_pExternal, L"TE", &v);
		VariantClear(&v);
		teSetObject(&v, this);
		tePutProperty(m_pExternal, L"WB", &v);
		VariantClear(&v);
		teSetLong(&v, CTRL_AR);
		tePutProperty(m_pExternal, L"Type", &v);
		VariantClear(&v);
	} else {
		g_pTE->QueryInterface(IID_PPV_ARGS(&m_pExternal));
	}
	MSG        msg;
	RECT       rc;
	IOleObject *pOleObject;
	CLSID clsid;
	if (SUCCEEDED(CLSIDFromProgID(L"Shell.Explorer", &clsid)) &&
		SUCCEEDED(CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pWebBrowser))) &&
		SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleObject)))) {
		pOleObject->SetClientSite(static_cast<IOleClientSite *>(this));
		SetRectEmpty(&rc);
		pOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, &msg, static_cast<IOleClientSite *>(this), 0, hwndParent, &rc);
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
	teReleaseClear(&m_pExternal);
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
//		QITABENT(CteWebBrowser, IDocHostShowUI),
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
					VariantClear(&m_vData);
					VariantCopy(&m_vData, &pDispParams->rgvarg[nArg]);
				}
				if (pVarResult) {
					VariantCopy(pVarResult, &m_vData);
				}
				return S_OK;
			//hwnd
			case 0x10000002:
				teSetPtr(pVarResult, get_HWND());
				return S_OK;
			//Type
			case 0x10000003:
				teSetLong(pVarResult, m_hwndParent == g_hwndMain ? CTRL_WB : CTRL_SW);
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
				teSetLong(pVarResult, hr);
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
			//Window
			case 0x10000007:
				if (m_pWebBrowser->get_Document(&pdisp) == S_OK) {
					IHTMLDocument2 *pdoc;
					if SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pdoc))) {
						IHTMLWindow2 *pwin;
						if SUCCEEDED(pdoc->get_parentWindow(&pwin)) {
							teSetObjectRelease(pVarResult, pwin);
							pwin->Release();
						}
						pdoc->Release();
					}
					pdisp->Release();
				}
				return S_OK;
			//Focus
			case 0x10000008:
				teSetForegroundWindow(m_hwndParent);
				return S_OK;
			//Close
/*			case 0x10000009:
				PostMessage(m_hwndParent, WM_CLOSE, 0, 0);
				return S_OK;*/
			//DIID_DWebBrowserEvents2
			case DISPID_DOWNLOADCOMPLETE:
				if (g_nReload == 1 && m_hwndParent == g_hwndMain) {
					g_nReload = 2;
					SetTimer(g_hwndMain, TET_Reload, 2000, teTimerProc);
				}
				return S_OK;
			case DISPID_BEFORENAVIGATE2:
				if (nArg >= 6 && m_hwndParent == g_hwndMain) {
					if (pDispParams->rgvarg[nArg - 6].vt == (VT_BYREF | VT_BOOL)) {
						VARIANT vURL;
						teVariantChangeType(&vURL, &pDispParams->rgvarg[nArg - 1], VT_BSTR);
						VARIANT v = pDispParams->rgvarg[nArg - 1];
						VARIANT_BOOL bSB = VARIANT_FALSE; 
						FolderItem *pid = new CteFolderItem(&vURL);
						pid->get_IsFolder(&bSB);
						if (bSB || StrChr(vURL.bstrVal, '/')) {
							*pDispParams->rgvarg[nArg - 6].pboolVal = VARIANT_TRUE;
							if (bSB) {
								CteShellBrowser *pSB = NULL;
								if (g_pTC) {
									pSB = g_pTC->GetShellBrowser(g_pTC->m_nIndex);
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
				if (m_bstrPath) {
					IOleWindow *pOleWindow;
					if SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
						pOleWindow->GetWindow(&m_hwndBrowser);
						pOleWindow->Release();
					}
					if (g_bsDocumentWrite) {
						g_pWebBrowser->write(g_bsDocumentWrite);
						teSysFreeString(&g_bsDocumentWrite);
						teShowWindow(m_hwndParent, SW_SHOWNORMAL);
						return S_OK;
					}
					if (g_bInit) {
						g_bInit = FALSE;
						KillTimer(g_hwndMain, TET_Create);
						SetTimer(g_hwndMain, TET_Create, 100, teTimerProc);
					}
					if (g_hwndMain != m_hwndParent) {
						teShowWindow(m_hwndParent, SW_SHOWNORMAL);
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
				if (m_hwndParent != g_hwndMain && m_bstrPath && nArg >= 0 && pDispParams->rgvarg[nArg].vt == VT_BSTR) {
					SetWindowText(m_hwndParent, pDispParams->rgvarg[nArg].bstrVal);
				}
				return S_OK;
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, L"WebBrowser", methodWB, dispIdMember);
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
	pInfo->dwFlags       = DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_SCROLL_NO;
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
		*pchKey = (LPOLESTR)CoTaskMemAlloc((lstrlen(szKey) + 1) * sizeof(WCHAR));
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
	teReleaseClear(&m_pDropTarget);
	if (pDropTarget) {
		pDropTarget->QueryInterface(IID_PPV_ARGS(&m_pDropTarget));
	}
	return QueryInterface(IID_PPV_ARGS(ppDropTarget));
}

STDMETHODIMP CteWebBrowser::GetExternal(IDispatch **ppDispatch)
{
	return m_pExternal->QueryInterface(IID_PPV_ARGS(ppDispatch));
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
		} else {
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

VOID CteWebBrowser::write(LPWSTR pszPath)
{
	IDispatch *pdisp = NULL;
	IHTMLDocument2 *pDoc = NULL;
	ClearEvents();
	do {
		Sleep(1000);
		if (g_pWebBrowser->m_pWebBrowser->get_Document(&pdisp) == S_OK && pdisp) {
			pdisp->QueryInterface(IID_PPV_ARGS(&pDoc));
			pdisp->Release();
		}
	} while (!pDoc);
	SAFEARRAY *psf = SafeArrayCreateVector(VT_VARIANT, 0, 1);
	VARIANT *pv;
	SafeArrayAccessData(psf, (LPVOID *)&pv);
	pv->vt = VT_BSTR;
	pv->bstrVal = ::SysAllocString(pszPath);
	SafeArrayUnaccessData(psf);
	HRESULT hr = pDoc->write(psf);
	VariantClear(pv);
	SafeArrayDestroy(psf);			
}

void CteWebBrowser::Close()
{
	if (m_pWebBrowser) {
		if (m_hwndParent != g_hwndMain) {
			teDelPropertyAtLLX(g_pSubWindows, (LONGLONG)m_hwndParent);
		}
		m_pWebBrowser->Quit();
		IOleObject *pOleObject;
		if SUCCEEDED(m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&pOleObject))) {
			RECT rc;
			SetRectEmpty(&rc);
			pOleObject->DoVerb(OLEIVERB_HIDE, NULL, NULL, 0, m_hwndParent, &rc);
			pOleObject->Release();
		}
		teUnadviseAndRelease(m_pWebBrowser, DIID_DWebBrowserEvents2, &m_dwCookie);
		PostMessage(get_HWND(), WM_CLOSE, 0, 0);
		teDelayRelease(&m_pWebBrowser);
	}
	teReleaseClear(&m_pDropTarget);
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
	m_pDragItems = NULL;
	m_bEmpty = false;
	VariantInit(&m_vData);
	m_param[TE_Left] = 0;
	m_param[TE_Top] = 0;
	m_param[TE_Width] = 100;
	m_param[TE_Height] = 100;
	m_param[TE_Flags] = TCS_FOCUSNEVER | TCS_HOTTRACK | TCS_MULTILINE | TCS_RAGGEDRIGHT | TCS_SCROLLOPPOSITE | TCS_TOOLTIPS;
	m_param[TC_Align] = 0;
	m_param[TC_TabWidth] = 0;
	m_param[TC_TabHeight] = 0;
	m_nScrollWidth = 0;
	m_DefTCProc = NULL;
	m_DefBTProc = NULL;
	m_DefSTProc = NULL;
	m_bRedraw = FALSE;
}

VOID CteTabCtrl::GetItem(int i, VARIANT *pVarResult)
{
	CteShellBrowser *pSB = GetShellBrowser(i);
	if (pSB) {
		teSetObject(pVarResult, pSB);
	}
}

VOID CteTabCtrl::Show(BOOL bVisible, BOOL bMain)
{
	bVisible &= 1;
	if (bVisible ^ m_bVisible) {
		m_bVisible = bVisible;
		if (bVisible) {
			if (bMain) {
				g_pTC = this;
			}
			ArrangeWindow();
		} else {
			MoveWindow(m_hwndStatic, -1, -1, 0, 0, false);
			for (int i = TabCtrl_GetItemCount(m_hwnd); i--;) {
				CteShellBrowser *pSB = GetShellBrowser(i);
				if (pSB) {
					pSB->Show(FALSE, 1);
				}
			}
		}
		DoFunc(TE_OnVisibleChanged, this, E_NOTIMPL);
	}
}

void CteTabCtrl::Close(BOOL bForce)
{
	if (CanClose(this) || bForce) {
		int nCount = TabCtrl_GetItemCount(m_hwnd);
		while (nCount--) {
			CteShellBrowser *pSB = GetShellBrowser(0);
			if (pSB) {
				pSB->Close(true);
			}
		}
		Show(FALSE, FALSE);
		if (m_bRedraw) {
			KillTimer(m_hwnd, (UINT_PTR)this);
		}
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
		if (this == g_pTC) {
			CteTabCtrl *pTC;
			for (int i = MAX_TC; i-- && (pTC = g_ppTC[i]);) {
				if (!pTC->m_bEmpty && pTC->m_bVisible) {
					g_pTC =  pTC;
					return;
				}
			}
		}
	}
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
	DWORD dwStyle = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | TCS_FOCUSNEVER | m_param[TE_Flags];
	if ((dwStyle & (TCS_SCROLLOPPOSITE | TCS_BUTTONS)) == TCS_SCROLLOPPOSITE) {
		if (m_param[TC_Align] == 4 || m_param[TC_Align] == 5) {
			dwStyle |= TCS_BUTTONS;
		}
	}
	if (dwStyle & TCS_BUTTONS) {
		if (dwStyle & TCS_SCROLLOPPOSITE) {
			dwStyle &= ~TCS_SCROLLOPPOSITE;
		}
		if (dwStyle & TCS_BOTTOM && m_param[TC_Align] > 1) {
			dwStyle &= ~TCS_BOTTOM;
		}
	}
	return dwStyle;
}

VOID CteTabCtrl::CreateTC()
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

BOOL CteTabCtrl::Create()
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

CteTabCtrl::~CteTabCtrl()
{
	Close(true);
	VariantClear(&m_vData);
}

STDMETHODIMP CteTabCtrl::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteTabCtrl, IDispatch),
		QITABENT(CteTabCtrl, IDispatchEx),
		QITABENT(CteTabCtrl, IDropTarget),
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
	return teGetDispId(methodTC, _countof(methodTC), g_maps[MAP_TC], *rgszNames, rgDispId, true);
}

STDMETHODIMP CteTabCtrl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
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
					VariantClear(&m_vData);
					VariantCopy(&m_vData, &pDispParams->rgvarg[nArg]);
				}
				if (pVarResult) {
					VariantCopy(pVarResult, &m_vData);
				}
				return S_OK;
			//hwnd
			case 2:
				teSetPtr(pVarResult, m_hwnd);
				return S_OK;
			//Type
			case 3:
				teSetLong(pVarResult, CTRL_TC);
				return S_OK;
			//HitTest (Screen coordinates)
			case 6:
				if (nArg >= 0 && pVarResult) {
					TCHITTESTINFO info;
					GetPointFormVariant(&info.pt, &pDispParams->rgvarg[nArg]);
					UINT flags = nArg >= 1 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : TCHT_ONITEM;
					teSetLong(pVarResult, static_cast<int>(DoHitTest(this, info.pt, flags)));
					if (pVarResult->lVal < 0) {
						ScreenToClient(m_hwnd, &info.pt);
						info.flags = flags;
						pVarResult->lVal = TabCtrl_HitTest(m_hwnd, &info);
						if (!(info.flags & flags)) {
							pVarResult->lVal = -1;
						}
					}
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
					teSetLong(pVarResult, TabCtrl_GetCurSel(m_hwnd));
				}
				return S_OK;
			//Visible
			case 11:
				if (nArg >= 0) {
					Show(GetIntFromVariant(&pDispParams->rgvarg[nArg]), TRUE);
				}
				teSetBool(pVarResult, m_bVisible);
				return S_OK;
			//Id
			case 12:
				teSetLong(pVarResult, MAX_TC - m_nTC);
				return S_OK;
			//LockUpdate
			case 13:
				LockUpdate();
				return S_OK;
			//UnlockUpdate
			case 14:
				UnLockUpdate(nArg >= 0 && GetIntFromVariant(&pDispParams->rgvarg[nArg]));
				return S_OK;
			//Item
			case DISPID_TE_ITEM:
				if (nArg >= 0) {
					GetItem(GetIntFromVariant(&pDispParams->rgvarg[nArg]), pVarResult);
				}
				return S_OK;
			//Count
			case DISPID_TE_COUNT:
				teSetLong(pVarResult, TabCtrl_GetItemCount(m_hwnd));
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
					} else if (dispIdMember == TE_OFFSET + TC_TabWidth || dispIdMember == TE_OFFSET + TC_TabHeight) {
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
		return teException(pExcepInfo, L"TabControl", methodTC, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteTabCtrl::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	return teGetDispId(methodTC, _countof(methodTC), g_maps[MAP_TC], bstrName, pid, true);
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
		SendMessage(m_hwnd, TCM_DELETEITEM, nSrc, 1/* Not delete SB flag */);
		CteShellBrowser *pSB = g_ppSB[tcItem.lParam];
		if (pSB) {
			teSetParent(pSB->m_hwnd, pDestTab->m_hwndStatic);
			pSB->m_pTC = pDestTab;
			TabCtrl_InsertItem(pDestTab->m_hwnd, nDest, &tcItem);
			if (nSrc == nIndex) {
				m_nIndex = -1;
				if (bOtherTab) {
					if (nSrc > 0) {
						nSrc--;
					}
					TabCtrl_SetCurSel(m_hwnd, nSrc);
				} else {
					TabCtrl_SetCurSel(m_hwnd, nDest);
					TabChanged(true);
				}
			}
		}
		if (bOtherTab) {
			pDestTab->m_nIndex = nDest;
			TabCtrl_SetCurSel(pDestTab->m_hwnd, nDest);
			ArrangeWindow();
		}
	}
}

VOID CteTabCtrl::LockUpdate()
{
	if (InterlockedIncrement(&m_nLockUpdate) == 1) {
		SendMessage(m_hwndStatic, WM_SETREDRAW, FALSE, 0);
		SendMessage(m_hwnd, WM_SETREDRAW, FALSE, 0);
		teSetRedraw(FALSE);
	}
}

VOID CteTabCtrl::UnLockUpdate(BOOL bDirect)
{
	if (::InterlockedDecrement(&m_nLockUpdate) > 0 || m_bRedraw) {
		return;
	}
	m_nLockUpdate = 0;
	m_bRedraw = TRUE;
	if (bDirect) {
		RedrawUpdate();
		return;
	}
	KillTimer(g_hwndMain, TET_Redraw);
	SetTimer(g_hwndMain, TET_Redraw, 500, teTimerProc);
}

VOID CteTabCtrl::RedrawUpdate()
{
	if (m_bRedraw) {
		if (g_nLockUpdate) {
			KillTimer(g_hwndMain, TET_Redraw);
			SetTimer(g_hwndMain, TET_Redraw, 500, teTimerProc);
			return;
		}
		m_bRedraw = FALSE;
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
STDMETHODIMP CteTabCtrl::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
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

STDMETHODIMP CteTabCtrl::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	HRESULT hr = DragSub(TE_OnDragOver, this, m_pDragItems, &grfKeyState, pt, pdwEffect);
	if (g_pDropTargetHelper) {
		g_pDropTargetHelper->DragOver((LPPOINT)&pt, *pdwEffect);
	}
	m_grfKeyState = grfKeyState;
	return hr;
}

STDMETHODIMP CteTabCtrl::DragLeave()
{
	if (m_DragLeave == E_NOT_SET) {
		if (g_pDropTargetHelper) {
			g_pDropTargetHelper->DragLeave();
		}
		m_DragLeave = DragLeaveSub(this, &m_pDragItems);
	}
	return m_DragLeave;
}

STDMETHODIMP CteTabCtrl::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	GetDragItems(&m_pDragItems, pDataObj);
	if (g_pDropTargetHelper) {
		g_pDropTargetHelper->Drop(pDataObj, (LPPOINT)&pt, *pdwEffect);
	}
	HRESULT hr = DragSub(TE_OnDrop, this, m_pDragItems, &m_grfKeyState, pt, pdwEffect);
	DragLeave();
	return hr;
}

VOID CteTabCtrl::TabChanged(BOOL bSameTC)
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

CteShellBrowser* CteTabCtrl::GetShellBrowser(int nPage)
{
	if (nPage < 0) {
		return NULL;
	}
	TC_ITEM tcItem;
	tcItem.mask = TCIF_PARAM;
	TabCtrl_GetItem(m_hwnd, nPage, &tcItem);
	if (tcItem.lParam < MAX_FV) {
		return g_ppSB[tcItem.lParam];
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
	teReleaseClear(&m_oFolderItems);
	teReleaseClear(&m_pDataObj);
	teReleaseClear(&m_pFolderItems);
	teSysFreeString(&m_bsText);
}

VOID CteFolderItems::Regenerate()
{
	if (!m_oFolderItems) {
		long nCount;
		if SUCCEEDED(get_Count(&nCount)) {
			IDispatch *oFolderItems = NULL;
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

VOID CteFolderItems::ItemEx(long nIndex, VARIANT *pVarResult, VARIANT *pVarNew)
{
	VARIANT v;
	teSetLong(&v, nIndex);
	if (pVarNew) {
		teSysFreeString(&m_bsText);
		Regenerate();
		if (m_oFolderItems) {
			IUnknown *punk;
			CteFolderItem *pid = NULL;
			if (FindUnknown(pVarNew, &punk)) {
				punk->QueryInterface(g_ClsIdFI, (LPVOID *)&pid);
			}
			if (pid == NULL) {
				pid = new CteFolderItem(pVarNew);
			}
			VARIANT vNew;
			teSetObjectRelease(&vNew, pid);
			if (nIndex >= 0) {
				tePutPropertyAt(m_oFolderItems, nIndex, &vNew);
			} else {
				teExecMethod(m_oFolderItems, L"push", NULL, -1, &vNew);
			}
			VariantClear(&vNew);
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
		teSetLong(&v, m_nCount);
		FolderItem *pid = NULL;
		while (--v.lVal >= 0) {
			if (Item(v, &pid) == S_OK) {
				teGetIDListFromObjectEx(pid, &m_pidllist[v.lVal + 1]);
				pid->Release();
			}
		}
		teReleaseClear(&m_oFolderItems);
	}
	m_bUseILF = AdjustIDList(m_pidllist, m_nCount);
}

STDMETHODIMP CteFolderItems::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENTMULTI(CteFolderItems, IDispatch, IDispatchEx),
		QITABENT(CteFolderItems, IDispatchEx),
		QITABENT(CteFolderItems, FolderItems),
		{ 0 },
	};
	*ppvObject = NULL;
	if (IsEqualIID(riid, IID_IDataObject)) {
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
		AddRef();
		return S_OK;
	}
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
					int pi[4] = { 0 };
					for (int i = nArg; i >= 0; i--) {
						pi[i] = GetIntFromVariant(&pDispParams->rgvarg[nArg - i]);
					}
					teSetPtr(pVarResult, GethDrop(pi[0], pi[1], pi[2], pi[3]));
				}
				return S_OK;
			//GetData
			case 10:
				if (pVarResult) {
					IDataObject *pDataObj;
					if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pDataObj))) {
						STGMEDIUM Medium;
						if (pDataObj->GetData(&UNICODEFormat, &Medium) == S_OK) {
							teSetSZ(pVarResult, static_cast<LPCWSTR>(GlobalLock(Medium.hGlobal)));
							GlobalUnlock(Medium.hGlobal);
							ReleaseStgMedium(&Medium);
						} else if (pDataObj->GetData(&TEXTFormat, &Medium) == S_OK) {
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
				teSetLong(pVarResult, m_nIndex);
				return S_OK;
			//dwEffect
			case 0x10000001:
				if (nArg >= 0) {
					m_dwEffect = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				}
				teSetLong(pVarResult, m_dwEffect);
				return S_OK;
			//pdwEffect
			case 0x10000002:
				punk = new CteMemory(sizeof(int), &m_dwEffect, 1, L"DWORD");
				teSetObject(pVarResult, punk);
				punk->Release();
				return S_OK;
			//this
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, L"FolderItems", methodFIs, dispIdMember);
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
			} else if (m_pDataObj->GetData(&HDROPFormat, &Medium) == S_OK) {
				m_nCount = DragQueryFile((HDROP)Medium.hGlobal, (UINT)-1, NULL, 0);
				ReleaseStgMedium(&Medium);
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
		} else {
			GetFolderItemFromPidl(ppid, m_pidllist[i + 1]);
		}
	} else if (i == -1) {
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
	*ppunk = static_cast<IDispatch *>(new CteDispatch(static_cast<FolderItems *>(this), 0));
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
				} catch (...) {}
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
			} else {
				teSysFreeString(&pbslist[i]);
			}
		} else {
			BSTR bsPath;
			pbslist[i] = NULL;
			if SUCCEEDED(GetDisplayNameFromPidl(&bsPath, pidl, SHGDN_FORPARSING)) {
				int nLen = ::SysStringByteLen(bsPath);
				if (nLen) {
					pbslist[i] = bsPath;
					uSize += nLen + sizeof(WCHAR);
				} else {
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
	} catch (...) {}
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
		} catch (...) {}
		GlobalUnlock(hGlobal);

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
		} catch (...) {}
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
			} catch (...) {}
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
			} catch (...) {}
		} else {
			hMem = GlobalAlloc(GHND, nLenA);
			try {
				LPSTR lpA = static_cast<LPSTR>(GlobalLock(hMem));
				if (lpA) {
					WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)m_bsText, nLen, lpA, nLenA, NULL, NULL);
				}
			} catch (...) {}
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
	if (IsEqualIID(riid, IID_IShellBrowser) || IsEqualIID(riid, IID_ICommDlgBrowser) || IsEqualIID(riid, IID_ICommDlgBrowser2) || IsEqualIID(riid, IID_IExplorerPaneVisibility)) {// || IsEqualIID(riid, IID_ICommDlgBrowser3)) {
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
	m_cRef = 1;
	m_pc = (char *)pc;
	m_bsStruct = NULL;
	m_nStructIndex = -1;
	if (lpStruct) {
		m_nStructIndex = teBSearchStruct(pTEStructs, _countof(pTEStructs), g_maps[MAP_SS], lpStruct);
		m_bsStruct = ::SysAllocString(lpStruct);
	}
	m_nCount = nCount;
	m_bsAlloc = NULL;
	m_nSize = 0;
	if (nSize > 0) {
		m_nSize = nSize;
		if (pc == NULL) {
			m_bsAlloc = ::SysAllocStringByteLen(NULL, nSize);
			m_pc = (char *)m_bsAlloc;
			if (m_pc) {
				::ZeroMemory(m_pc, nSize);
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
	if (m_bsAlloc) {
		::SysFreeString(m_bsAlloc);
		m_bsAlloc = NULL;
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
				teSetLong(pVarResult, m_nSize);
				return S_OK;
			//Free
			case 7:
				Free(TRUE);
				return S_OK;
			//Clone
			case 8:
				if (m_nSize) {
					CteMemory *pMem = new CteMemory(m_nSize, NULL, m_nCount, m_bsStruct);
					::CopyMemory(pMem->m_pc, m_pc, m_nSize);
					teSetObjectRelease(pVarResult, pMem);
				}
				return S_OK;
			//_BLOB
			case 9:
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
		return teException(pExcepInfo, L"Memory", methodMem, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteMemory::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	if (m_nStructIndex >= 0 && teGetDispId(pTEStructs[m_nStructIndex].pMethod, 0, NULL, bstrName, pid, false) == S_OK) {
		return S_OK;
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
						int nOffset = GetOffsetOfStruct(m_bsStruct);
						nSize -= nOffset;
						teSetObjectRelease(pVarResult, new CteMemory(nSize, m_pc + nOffset + (nSize * i), 1, GetSubStruct(m_bsStruct)));
					}
				}
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
	return m_ppbs[m_nbs++] = ::SysAllocStringByteLen((char *)bs, ::SysStringByteLen(bs));
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
				teSetLL(pVarResult, *(LONGLONG *)&m_pc[nIndex]);
				break;
			case VT_R4:
				pVarResult->fltVal = *(FLOAT *)&m_pc[nIndex];
				break;
			case VT_R8:
				pVarResult->dblVal = *(DOUBLE *)&m_pc[nIndex];
				break;
			case VT_PTR:
			case VT_BSTR:
				teSetPtr(pVarResult, *(INT_PTR *)&m_pc[nIndex]);
				break;
			case VT_LPWSTR:
				if (nLen >= 0) {
					if (nLen * (int)sizeof(WCHAR) > m_nSize) {
						nLen = m_nSize / sizeof(WCHAR);
					}
					pVarResult->bstrVal = SysAllocStringLenEx((WCHAR *)&m_pc[nIndex], nLen, lstrlen((WCHAR *)&m_pc[nIndex]));
					pVarResult->vt = VT_BSTR;
				} else {
					teSetSZ(pVarResult, (WCHAR *)&m_pc[nIndex]);
				}
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
				ppt = (POINT *)&m_pc[nIndex];
				pMem = new CteMemory(sizeof(POINT), NULL, 1, L"POINT");
				pMem->SetPoint(ppt->x, ppt->y);
				teSetObjectRelease(pVarResult, pMem);
				break;
			case VT_CARRAY:
				pMem = new CteMemory(sizeof(RECT), NULL, 1, L"RECT");
				::CopyMemory(pMem->m_pc, &m_pc[nIndex], sizeof(RECT));
				teSetObjectRelease(pVarResult, pMem);
				break;
			case VT_CLSID:
				LPOLESTR lpsz;
				StringFromCLSID(*(IID *)&m_pc[nIndex], &lpsz);
				teSetSZ(pVarResult, lpsz);
				teCoTaskMemFree(lpsz);
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
	} catch (...) {
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
				} else {
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
				VARIANT vMem;
				VariantInit(&vMem);
				pc = GetpcFromVariant(pv, &vMem);
				if (pc) {
					::CopyMemory(&m_pc[nIndex], pc, sizeof(RECT));
				} else {
					::ZeroMemory(&m_pc[nIndex], sizeof(RECT));
				}
				VariantClear(&vMem);
				break;
			case VT_CLSID:
				if (pv->vt == VT_BSTR) {
					teCLSIDFromString(pv->bstrVal, (LPCLSID)&m_pc[nIndex]);
				}
				break;
			case VT_VARIANT:
				VariantCopy((VARIANT *)&m_pc[nIndex], pv);
				break;
/*			case VT_RECORD:
				LPITEMIDLIST *ppidl;
				ppidl = (LPITEMIDLIST *)&m_pc[nIndex];
				teCoTaskMemFree(*ppidl);
				teGetIDListFromVariant(ppidl, pv);
				break;*/
		}//end_switch
	} catch (...) {
		g_nException = 0;
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
	m_pDll = NULL;
	if (punk) {
		punk->QueryInterface(IID_PPV_ARGS(&m_pContextMenu));
		if (punkSB) {
			IShellBrowser *pSB;
			if SUCCEEDED(punkSB->QueryInterface(IID_PPV_ARGS(&pSB))) {
				IShellView *pSV = NULL;
				pSB->QueryActiveShellView(&pSV);
				CteServiceProvider *pSP = new CteServiceProvider(punkSB, pSV);
				HRESULT hr = IUnknown_SetSite(punk, pSB);
				teReleaseClear(&pSP);
				teReleaseClear(&pSV);
				teReleaseClear(&pSB);
			}
		}
	}
}

CteContextMenu::~CteContextMenu()
{
	teReleaseClear(&m_pContextMenu);
	teReleaseClear(&m_pDataObj);
	teReleaseClear(&m_pDll);
}

STDMETHODIMP CteContextMenu::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteContextMenu, IDispatch),
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
	return teGetDispId(methodCM, 0, NULL, *rgszNames, rgDispId, true);
}

STDMETHODIMP CteContextMenu::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		HRESULT hr = NULL;

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
					teSetLong(pVarResult, m_pContextMenu->QueryContextMenu((HMENU)m_param[0], (UINT)m_param[1], (UINT)m_param[2], (UINT)m_param[3], (UINT)m_param[4]));
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
					CMINVOKECOMMANDINFOEX cmi = { sizeof(cmi) };
					VARIANTARG *pv = GetNewVARIANT(3);
					WCHAR **ppwc = new WCHAR*[3];
					char **ppc = new char*[3];
					BOOL bExec = TRUE;
					for (int i = 0; i <= 2; i++) {
						if (pDispParams->rgvarg[nArg - i - 2].vt == VT_I4) {
							UINT_PTR nVerb = (UINT_PTR)pDispParams->rgvarg[nArg - i - 2].lVal;
							ppwc[i] = (WCHAR *)nVerb;
							ppc[i] = (char *)nVerb;
							if (nVerb > MAXWORD) {
								bExec = FALSE;
							}
						} else {
							teVariantChangeType(&pv[i], &pDispParams->rgvarg[nArg - i - 2], VT_BSTR);
							ppwc[i] = pv[i].bstrVal;
							if (i == 2) {
								if (!tePathMatchSpec(ppwc[i], L"?:\\*;\\\\*\\*") || tePathIsDirectory(ppwc[i], 0, TRUE) != S_OK) {
									ppwc[i] = NULL;
								}
							}
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
							} else {
								ppc[i] = (char *)ppwc[i];
							}
						}
					}
					if (bExec) {
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
						} catch (...) {
							hr = E_UNEXPECTED;
						}
					}
					teSetLong(pVarResult, hr);
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
					WCHAR szName[MAX_PATH];
					szName[0] = NULL;
					UINT idCmd = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					if (idCmd <= MAXWORD) {
						m_pContextMenu->GetCommandString(idCmd,
							GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]),
							NULL, (LPSTR)szName, MAX_PATH);
					}
					teSetSZ(pVarResult, szName);
				}
				return S_OK;
			//FolderView
			case 5:
				IShellBrowser *pSB;
				if (GetFolderVew(&pSB)) {
					teSetObjectRelease(pVarResult, pSB);
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
					} else if SUCCEEDED(QueryInterface(IID_PPV_ARGS(&pCM2))) {
						pCM2->HandleMenuMsg(
							(UINT)GetIntFromVariant(&pDispParams->rgvarg[nArg]),
							(WPARAM)GetLLFromVariant(&pDispParams->rgvarg[nArg - 1]),
							(LPARAM)GetLLFromVariant(&pDispParams->rgvarg[nArg - 2])
						);
						pCM2->Release();
					}
				}
				teSetPtr(pVarResult, lResult);
				return S_OK;

			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
			default:
				if (dispIdMember >= 10 && dispIdMember <= 14) {
					teSetPtr(pVarResult, m_param[dispIdMember - 10]);
					return S_OK;
				}
				break;
		}
	} catch (...) {
		return teException(pExcepInfo, L"ContextMenu", methodCM, dispIdMember);
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
	teReleaseClear(&m_pFolderItem);
	teReleaseClear(&m_pDropTarget);
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
							GetPointFormVariant((POINT *)&pt, &pDispParams->rgvarg[nArg - 2]);
							DWORD dwEffect1 = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
							VARIANT v;
							VariantInit(&v);
							IDispatch *pEffect = NULL;
							if (nArg >= 3) {
								if (GetDispatch(&pDispParams->rgvarg[nArg - 3], &pEffect)) {
									teGetPropertyAt(pEffect, 0, &v);
									dwEffect1 = GetIntFromVariantClear(&v);
								} else {
									dwEffect1 = GetIntFromVariantClear(&pDispParams->rgvarg[nArg - 3]);
								}
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
									} else {
										dwEffect1 = 0;
									}
									if (m_pFolderItem) {
										DWORD dwEffect2 = dwEffect;
										if (DragSub(TE_OnDragEnter, this, pDragItems, &grfKeyState, pt0, &dwEffect2) == S_OK) {
											hr = S_OK;
											dwEffect1 = dwEffect2;
										}
									}
								} else {
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
							} catch (...) {}
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
								teSetLong(&v, dwEffect1);
								tePutPropertyAt(pEffect, 0, &v);
								pEffect->Release();
							}
						} catch (...) {}
						pDataObj->Release();
					}
				}
				teSetLong(pVarResult, hr);
				return S_OK;
			//DragLeave
			case 4:
				if (m_pDropTarget) {
					hr = m_pDropTarget->DragLeave();
				}
				DragLeaveSub(this, NULL);
				teSetLong(pVarResult, hr);
				return S_OK;
			//Type
			case 5:
				teSetLong(pVarResult, CTRL_DT);
				return S_OK;
			//FolderItem
			case 6:
				teSetObject(pVarResult, m_pFolderItem);
				return S_OK;

			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, L"DropTarget", methodDT, dispIdMember);
	}
	return DISP_E_MEMBERNOTFOUND;
}

//CteTreeView

CteTreeView::CteTreeView()
{
	m_cRef = 1;
#ifdef _2000XP
	m_DefProc = NULL;
#endif
	m_DefProc2 = NULL;
	m_bMain = true;
	m_pDragItems = NULL;
	m_pDropTarget = NULL;
	VariantInit(&m_vData);
#ifdef _W2000
	VariantInit(&m_vSelected);
#endif
}

CteTreeView::~CteTreeView()
{
	Close();
#ifdef _2000XP
	if (m_DefProc) {
		SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, (LONG_PTR)m_DefProc);
		m_DefProc = NULL;
	}
#endif
	if (m_DefProc2) {
		SetWindowLongPtr(m_hwndTV, GWLP_WNDPROC, (LONG_PTR)m_DefProc2);
		m_DefProc2 = NULL;
	}
	if (m_pDropTarget) {
		RevokeDragDrop(m_hwndTV);
		RegisterDragDrop(m_hwndTV, m_pDropTarget);
		teReleaseClear(&m_pDropTarget);
	}
#ifdef _W2000
	VariantClear(&m_vSelected);
#endif
	VariantClear(&m_vData);
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
		teDelayRelease(&m_pNameSpaceTreeControl);
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
	if (_Emulate_XP_ SUCCEEDED(CoCreateInstance(CLSID_NamespaceTreeControl, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pNameSpaceTreeControl)))) {
		RECT rc;
		SetRectEmpty(&rc);
		if SUCCEEDED(m_pNameSpaceTreeControl->Initialize(m_pFV->m_pTC->m_hwndStatic, &rc, m_pFV->m_param[SB_TreeFlags])) {
			m_pNameSpaceTreeControl->TreeAdvise(static_cast<INameSpaceTreeControlEvents *>(this), &m_dwCookie);
			IOleWindow *pOleWindow;
			if SUCCEEDED(m_pNameSpaceTreeControl->QueryInterface(IID_PPV_ARGS(&pOleWindow))) {
				if (pOleWindow->GetWindow(&m_hwnd) == S_OK) {
					m_hwndTV = FindTreeWindow(m_hwnd);
					if (m_hwndTV) {
						SetWindowLongPtr(m_hwndTV, GWLP_USERDATA, (LONG_PTR)this);
						m_DefProc2 = (WNDPROC)SetWindowLongPtr(m_hwndTV, GWLP_WNDPROC, (LONG_PTR)TETVProc2);
						if (!m_pDropTarget) {
							teRegisterDragDrop(m_hwndTV, this, &m_pDropTarget);
						}
					}
					BringWindowToTop(m_hwnd);
					ArrangeWindow();
				}
				pOleWindow->Release();
			}
/*// Too heavy.
			INameSpaceTreeControl2 *pNSTC2;
			if SUCCEEDED(m_pNameSpaceTreeControl->QueryInterface(IID_PPV_ARGS(&pNSTC2))) {
				pNSTC2->SetControlStyle2(NSTCS2_INTERRUPTNOTIFICATIONS, NSTCS2_INTERRUPTNOTIFICATIONS);
				pNSTC2->Release();
			}
*///
			DoFunc(TE_OnCreate, this, E_NOTIMPL);
			return TRUE;
		}
	}
#ifdef _2000XP
	m_hwnd = CreateWindowEx(
		WS_EX_CLIENTEDGE, WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | ES_READONLY,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
		m_pFV->m_pTC->m_hwndStatic, (HMENU)0, hInst, NULL);
	if FAILED(CoCreateInstance(CLSID_ShellShellNameSpace, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pShellNameSpace))) {
		if FAILED(CoCreateInstance(CLSID_ShellNameSpace, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pShellNameSpace))) {
			return FALSE;
		}
	}
	IQuickActivate *pQuickActivate;
	if (FAILED(m_pShellNameSpace->QueryInterface(IID_PPV_ARGS(&pQuickActivate)))) {
		return FALSE;
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
	pOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, static_cast<IOleClientSite *>(this), 0, m_pFV->m_pTC->m_hwndStatic, &rc);
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
				if (!m_pDropTarget) {
					teRegisterDragDrop(m_hwndTV, this, &m_pDropTarget);
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
					//NSTCS_NOREPLACEOPEN	= 0x800,
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
					if (m_pFV->m_param[SB_TreeFlags] & style[i].nstc) {
						dwStyle |= style[i].nsc;
					} else {
						dwStyle &= ~style[i].nsc;
					}
				}
				dwStyle ^= (TVS_NOHSCROLL | TVS_INFOTIP | TVS_NONEVENHEIGHT | TVS_EDITLABELS);
				SetWindowLongPtr(m_hwndTV, GWL_STYLE, dwStyle & ~WS_BORDER);
				DWORD dwExStyle = GetWindowLong(m_hwnd, GWL_EXSTYLE);
				if (dwStyle & WS_BORDER) {
					dwExStyle |= WS_EX_CLIENTEDGE;
				} else {
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
	static const QITAB qit[] =
	{
		QITABENT(CteTreeView, IDispatch),
		QITABENT(CteTreeView, IDropTarget),
		QITABENT(CteTreeView, INameSpaceTreeControlEvents),
		QITABENT(CteTreeView, INameSpaceTreeControlCustomDraw),
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
			} else if (dispIdMember != 0x20000002) {
				return S_FALSE;
			}
		}

		switch(dispIdMember) {
			//Data
			case 0x10000001:
				if (nArg >= 0) {
					VariantClear(&m_vData);
					VariantCopy(&m_vData, &pDispParams->rgvarg[nArg]);
				}
				if (pVarResult) {
					VariantCopy(pVarResult, &m_vData);
				}
				return S_OK;
			//Type
			case 0x10000002:
				teSetLong(pVarResult, CTRL_TV);
				return S_OK;
			//hwnd
			case 0x10000003:
				teSetPtr(pVarResult, m_hwnd);
				return S_OK;
			//Close
			case 0x10000004:
				m_pFV->m_param[SB_TreeAlign] = 0;
				ArrangeWindow();
				return S_OK;
			//hwnd
			case 0x10000005:
				teSetPtr(pVarResult, m_hwndTV);
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
				teSetLong(pVarResult, m_pFV->m_param[SB_TreeAlign]);
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
				teSetBool(pVarResult, m_pFV->m_param[SB_TreeAlign] & 2);
				return S_OK;
			//Focus
			case 0x10000106:
				SetFocus(m_hwndTV);
				return S_OK;
			//HitTest (Screen coordinates)
			case 0x10000107:
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
					teSetPtr(pVarResult, hItem);
				}
				return S_OK;
			//GetItemRect
			case 0x10000283:
				if (nArg >= 1) {
					VARIANT vMem;
					VariantInit(&vMem);
					LPRECT prc = (LPRECT)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1], &vMem);
					if (prc) {
						LPITEMIDLIST pidl;
						if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
							if (m_pNameSpaceTreeControl && lpfnSHCreateItemFromIDList) {
								IShellItem *pShellItem;
								if SUCCEEDED(lpfnSHCreateItemFromIDList(pidl, IID_PPV_ARGS(&pShellItem))) {
									teSetLong(pVarResult, m_pNameSpaceTreeControl->GetItemRect(pShellItem, prc));
									pShellItem->Release();
								}
							}
							teCoTaskMemFree(pidl);
						}
					}
					teWriteBack(&pDispParams->rgvarg[nArg - 1], &vMem);
					VariantClear(&vMem);
				}
				return S_OK;
			//Notify
			case 0x10000300:
				if (nArg >= 2 && m_pNameSpaceTreeControl) {
					long lEvent = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
					if (lEvent & (SHCNE_MKDIR | SHCNE_MEDIAINSERTED | SHCNE_DRIVEADD | SHCNE_NETSHARE)) {					
						LPITEMIDLIST pidl;
						if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg - 1])) {
							IShellItem *psi, *psiParent;
							if SUCCEEDED(lpfnSHCreateItemFromIDList(pidl, IID_PPV_ARGS(&psi))) {
								DWORD dwState;
								if FAILED(m_pNameSpaceTreeControl->GetItemState(psi, NSTCIS_EXPANDED, &dwState)) {
									if (ILRemoveLastID(pidl) && SUCCEEDED(lpfnSHCreateItemFromIDList(pidl, IID_PPV_ARGS(&psiParent)))) {
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
			//SelectedItems
			case 0x20000001:
				if (getSelected(&pdisp) == S_OK) {
					CteFolderItems *pFolderItems = new CteFolderItems(NULL, NULL, true);
					VARIANT v;
					VariantInit(&v);
					teSetObjectRelease(&v, pdisp);
					pFolderItems->ItemEx(-1, NULL, &v);
					VariantClear(&v);
					teSetObjectRelease(pVarResult, pFolderItems);
					return S_OK;
				}
				break;
			//Root
			case 0x20000002:
				if (pVarResult) {
					VariantCopy(pVarResult, &m_pFV->m_vRoot);
					if (pVarResult->vt == VT_NULL) {
						teSetLong(pVarResult, 0);
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
				if (nArg >= 1) {
					if (g_nLockUpdate) {
						return S_OK;
					}
					LPITEMIDLIST pidl;
					teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg]);
					if (ILIsEqual(pidl, g_pidlResultsFolder)) {
						teCoTaskMemFree(pidl);
						return S_OK;
					}
					if (m_pNameSpaceTreeControl) {
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
						return S_OK;
					}
					teCoTaskMemFree(pidl);
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
					teSetLong(pVarResult, m_pFV->m_param[dispIdMember - TE_OFFSET]);
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
	} catch (...) {
		return teException(pExcepInfo, L"TreeView", methodTV, dispIdMember);
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
		} else {
			IDispatch *pdisp;
			if (getSelected(&pdisp) == S_OK) {
				teSetObjectRelease(&pv[2], pdisp);
			}
		}
		teSetLong(&pv[1], nstceHitTest);
		teSetLong(&pv[0], nsctsFlags);
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
	if (g_pOnFunc[TE_OnItemPrePaint] && !(pnstccdItem->uItemState & CDIS_SELECTED)) {
		NMTVCUSTOMDRAW tvcd = { { { 0 }, 0, hdc, *prc, 0, pnstccdItem->uItemState }, *pclrText, *pclrTextBk, pnstccdItem->iLevel };
		VARIANTARG *pv = GetNewVARIANT(5);
		teSetObject(&pv[4], this);
		LPITEMIDLIST pidl;
		if (teGetIDListFromObject(pnstccdItem->psi, &pidl)) {
			teSetIDList(&pv[3], pidl);
			CoTaskMemFree(pidl);
		}
		teSetObjectRelease(&pv[2], new CteMemory(sizeof(NMCUSTOMDRAW), &tvcd.nmcd, 1, L"NMCUSTOMDRAW"));
		teSetObjectRelease(&pv[1], new CteMemory(sizeof(NMTVCUSTOMDRAW), &tvcd, 1, L"NMTVCUSTOMDRAW"));
		teSetObjectRelease(&pv[0], new CteMemory(sizeof(HANDLE), plres, 1, L"HANDLE"));
		Invoke4(g_pOnFunc[TE_OnItemPrePaint], NULL, 5, pv);
		*pclrText = tvcd.clrText;
		*pclrTextBk = tvcd.clrTextBk;
	}
	return S_OK;
}

STDMETHODIMP CteTreeView::ItemPostPaint(HDC hdc, RECT *prc, NSTCCUSTOMDRAW *pnstccdItem)
{
	return S_OK;
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
			if SUCCEEDED(pDispatch->GetIDsOfNames(IID_NULL, &lpVerb, 1, LOCALE_USER_DEFAULT, &dispid)) {
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
			if (!teGetIDListFromVariant(&pidl, &m_pFV->m_vRoot)) {
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
		} else {
			dwStyle |= TVS_LINESATROOT;
		}
		SetWindowLong(m_hwndTV, GWL_STYLE, dwStyle);

		//running to call DISPID_FAVSELECTIONCHANGE SetRoot(Windows 2000)
		m_pShellNameSpace->put_EnumOptions(m_pFV->m_param[SB_EnumFlags]);

		LPITEMIDLIST pidl;
		if (teGetIDListFromVariant(&pidl, &m_pFV->m_vRoot)) {
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
			} else {
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
	m_pFolderItem = NULL;
	m_bStrict = FALSE;
	m_pidlFocused = NULL;
	m_dwUnavailable = 0;
	VariantInit(&m_v);
	if (pv) {
		VariantCopy(&m_v, pv);
		if (pv->vt == VT_DISPATCH) {
			CteFolderItem *pid;
			if SUCCEEDED(pv->pdispVal->QueryInterface(g_ClsIdFI, (LPVOID *)&pid)) {
				if (pid->m_v.vt != VT_EMPTY) {
					VariantClear(&m_v);
					VariantCopy(&m_v, &pid->m_v);
				}
				pid->Release();
			}
		}
	}
}

CteFolderItem::~CteFolderItem()
{
	teCoTaskMemFree(m_pidl);
	teReleaseClear(&m_pFolderItem);
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
				} catch(...) {}
				::InterlockedDecrement(&g_nProcFI);
			}
		}
		if (teGetIDListFromVariant(&m_pidl, &vPath)) {
			BSTR bs;
			if (GetDisplayNameFromPidl(&bs, m_pidl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING) == S_OK) {
				SysFreeString(bs);
			}
			if (m_v.vt != VT_BSTR || lstrcmp(m_v.bstrVal, vPath.bstrVal) == 0) {
				::VariantClear(&m_v);
			}
		} else {
			m_dwUnavailable = GetTickCount();
			teILCloneReplace(&m_pidl, g_pidlResultsFolder);
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
	if (IsEqualIID(riid, IID_FolderItem2) && GetFolderItem()) {
		return m_pFolderItem->QueryInterface(riid, ppvObject);
	}
	return QISearch(this, qit, riid, ppvObject);
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
	int i;
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
					teGetIDListFromVariant(&m_pidl, &pDispParams->rgvarg[nArg]);
				}
				teSetIDList(pVarResult, m_pidl);
				return S_OK;
			//Unavailable
			case 5:
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
			//_BLOB ** To be necessary
			case 9:
				return S_OK;
/*/// Reserved future
			//FocusedItem
			case 0x4001004:
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
				return S_OK;
		}
		if (GetFolderItem()) {
			if (dispIdMember > 10) {
				dispIdMember -= 10;
			}
			return m_pFolderItem->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
	} catch (...) {
		return teException(pExcepInfo, L"FolderItem", methodFI, dispIdMember);
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
	if (m_v.vt == VT_BSTR) {
		*pbs = ::SysAllocString(PathFindFileName(m_v.bstrVal));
		return S_OK;
	}
	return GetDisplayNameFromPidl(pbs, GetPidl(), SHGDN_INFOLDER);
}

STDMETHODIMP CteFolderItem::put_Name(BSTR bs)
{
	return GetFolderItem() ? m_pFolderItem->put_Name(bs) : E_FAIL;
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
					} else {
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
				teSetSZ(pVarResult, m_ofn.lpstrInitialDir);
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
			teSetBool(pVarResult, bResult);
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
			teSetLong(pVarResult, *pdw);
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
			teSetSZ(pVarResult, *pbs);
			return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, L"CommonDialog", methodCD, dispIdMember);
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
	static const QITAB qit[] =
	{
		QITABENT(CteGdiplusBitmap, IDispatch),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
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
		} else if (!m_pImage) {
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
					if (nArg >= 1 && pDispParams->rgvarg[nArg - 1].vt != VT_NULL) {
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
						return S_OK;
					}
					m_pImage = Gdiplus::Bitmap::FromHICON(hicon);
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
						VARIANT vMem;
						VariantInit(&vMem);
						EncoderParameters *pEncoderParameters = NULL;
						if (nArg) {
							char *pc = GetpcFromVariant(&pDispParams->rgvarg[nArg - 1], &vMem);
							if (pc) {
								pEncoderParameters = (EncoderParameters *)pc;
							}
						}
						Result = m_pImage->Save(vText.bstrVal, &encoderClsid, pEncoderParameters);
						VariantClear(&vMem);
					}
				}
				teSetLong(pVarResult, Result);
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
						IStream *pStream = SHCreateMemStream(NULL, NULL);
						if (pStream) {
							if (m_pImage->Save(pStream, &encoderClsid) == 0) {
								ULARGE_INTEGER uliSize;
								LARGE_INTEGER liOffset;
								liOffset.QuadPart = 0;
								pStream->Seek(liOffset, STREAM_SEEK_END, &uliSize);
								pStream->Seek(liOffset, STREAM_SEEK_SET, NULL);
								UCHAR *pBuff = new UCHAR[uliSize.LowPart];
								ULONG ulBytesRead;
								if SUCCEEDED(pStream->Read(pBuff, uliSize.LowPart, &ulBytesRead)) {
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
				teSetLong(pVarResult, m_pImage->GetWidth());
				return S_OK;
			//GetHeight
			case 111:
				teSetLong(pVarResult, m_pImage->GetHeight());
				return S_OK;
			//GetPixel
			case 112:
				if (nArg >= 1) {
					Color cl;
					if (m_pImage->GetPixel(GetIntFromVariant(&pDispParams->rgvarg[nArg]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), &cl) == 0) {
						teSetLong(pVarResult, cl.GetValue());
					}
				}
				return S_OK;
			//SetPixel
			case 113:
				if (nArg >= 2) {
					Color cl;
					cl.SetValue(GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]));
					m_pImage->SetPixel(GetIntFromVariant(&pDispParams->rgvarg[nArg]), GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]), cl);
				}
				return S_OK;
			//GetThumbnailImage
			case 120:
				if (nArg >= 1) {
					CLSID encoderClsid;
					if (GetEncoderClsid(L"image/png", &encoderClsid, NULL)) {
						IStream *pStream = SHCreateMemStream(NULL, NULL);
						if (pStream) {
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
						teSetPtr(pVarResult, hBM);
						return S_OK;
					}
#endif
					if (m_pImage->GetHBITMAP(((cl & 0xff) << 16) + (cl & 0xff00) + ((cl & 0xff0000) >> 16), &hBM) == 0) {
						teSetPtr(pVarResult, hBM);
					}
				}
				return S_OK;
			//GetHICON
			case 211:
				HICON hIcon;
				if (m_pImage->GetHICON(&hIcon) == 0) {
					teSetPtr(pVarResult, hIcon);
				}
				return S_OK;
			case DISPID_VALUE:
				teSetObject(pVarResult, this);
				return S_OK;
		}
	} catch (...) {
		return teException(pExcepInfo, L"GdiplusBitmap", methodGB, dispIdMember);
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
	*ppvObject = NULL;
	
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	} else {
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
	int nIndex = teBSearchApi(dispAPI, _countof(dispAPI), g_maps[MAP_API], *rgszNames);
	if (nIndex >= 0) {
		*rgDispId = nIndex + 0x60010000;
		return S_OK;
	}
	return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP CteWindowsAPI::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (wFlags & DISPATCH_METHOD) {
		if (dispIdMember >= 0x60010000 && dispIdMember < _countof(dispAPI) + 0x60010000) {
			try {
				return teInvokeAPI(&dispAPI[dispIdMember - 0x60010000], pDispParams, pVarResult, pExcepInfo);
			} catch (...) {
				return teExceptionEx(pExcepInfo, L"WindowsAPI", m_pApi->name);
			}
		}
		if (m_pApi) {
			try {
				return teInvokeAPI(m_pApi, pDispParams, pVarResult, pExcepInfo);
			} catch (...) {
				return teExceptionEx(pExcepInfo, L"WindowsAPI", m_pApi->name);
			}
		}
	} else {
		if (dispIdMember == DISPID_VALUE) {
			teSetObject(pVarResult, this);
		} else if (dispIdMember >= 0x60010000 && dispIdMember < _countof(dispAPI) + 0x60010000) {
			teSetObjectRelease(pVarResult, new CteWindowsAPI(&dispAPI[dispIdMember - 0x60010000]));
		}
		return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}
#endif
//CteDispatch

CteDispatch::CteDispatch(IDispatch *pDispatch, int nMode)
{
	m_cRef = 1;
	m_pDispatch = pDispatch;
	m_pDispatch->AddRef();
	m_nMode = nMode;
	m_dispIdMember = DISPID_UNKNOWN;
	m_nIndex = 0;
	m_pActiveScript = NULL;
	m_pDll = NULL;
}

CteDispatch::~CteDispatch()
{
	Clear();
}

VOID CteDispatch::Clear()
{
	teReleaseClear(&m_pDispatch);
	if (m_pActiveScript) {
		m_pActiveScript->SetScriptState(SCRIPTSTATE_CLOSED);
		m_pActiveScript->Close();
		m_pActiveScript->Release();
		m_pActiveScript = NULL;
	}
	teReleaseClear(&m_pDll);
}


STDMETHODIMP CteDispatch::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteDispatch, IDispatch),
		QITABENT(CteDispatch, IEnumVARIANT),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
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
	if (m_nMode && m_pDispatch) {
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
		if (m_nMode & 3) {
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
		}
		if (wFlags & DISPATCH_PROPERTYGET && dispIdMember == DISPID_VALUE) {
			teSetObject(pVarResult, this);
			return S_OK;
		}
		if (m_pDispatch) {
			return m_pDispatch->Invoke(m_nMode ? dispIdMember : m_dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
	} catch (...) {
		return teException(pExcepInfo, L"Dispatch", NULL, dispIdMember);
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
	teReleaseClear(&m_pOnError);
	teReleaseClear(&m_pDispatchEx);
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
		if (g_pWebBrowser && lstrcmpi(pstrName, L"window") == 0) {
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
		CteMemory *pei = new CteMemory(sizeof(EXCEPINFO), NULL, 1, L"EXCEPINFO");
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
	VARIANT *pv = GetNewVARIANT(1);
	teSetLL(pv, (LONGLONG)m_hDll);
	teExecMethod(g_pFreeLibrary, L"push", NULL, 1, pv);
	SetTimer(g_hwndMain, TET_FreeLibrary, 100, teTimerProc);
}

STDMETHODIMP CteDll::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteDll, IUnknown),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteDll::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteDll::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		LPFNDllCanUnloadNow lpfnDllCanUnloadNow = (LPFNDllCanUnloadNow)GetProcAddress(m_hDll, "DllCanUnloadNow");
		if (g_pUnload && lpfnDllCanUnloadNow && lpfnDllCanUnloadNow() != S_OK) {
			if (!g_bUnload) {
				g_bUnload = TRUE;
				teArrayPush(g_pUnload, this);
				SetTimer(g_hwndMain, TET_Unload, 15 * 1000, teTimerProc);
			}
			return m_cRef;
		}
		delete this;
		return 0;
	}
	return m_cRef;
}

#ifndef _USE_BSEARCHAPI
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
		return teExceptionEx(pExcepInfo, L"WindowsAPI", m_pApi->name);
	}
	return DISP_E_MEMBERNOTFOUND;
}
#endif

///CteDispatchEx

CteDispatchEx::CteDispatchEx(IUnknown *punk)
{
	m_cRef = 1;
	punk->QueryInterface(IID_PPV_ARGS(&m_pdex));
	m_dispItem = 0;
}


CteDispatchEx::~CteDispatchEx()
{
	m_pdex->Release();
}

STDMETHODIMP CteDispatchEx::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteDispatchEx, IDispatch),
		QITABENT(CteDispatchEx, IDispatchEx),
		{ 0 },
	};
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
	return GetDispID(*rgszNames, fdexNameCaseSensitive, rgDispId);
}

STDMETHODIMP CteDispatchEx::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	return InvokeEx(dispIdMember, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, NULL);
}

//IDispatchEx
STDMETHODIMP CteDispatchEx::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	HRESULT hr = m_pdex->GetDispID(bstrName, grfdex, pid);
	if FAILED(hr) {
		if (lstrcmpi(bstrName, L"Count") == 0) {
			BSTR bs = ::SysAllocString(L"length");
			hr = m_pdex->GetDispID(bs, fdexNameCaseSensitive, pid);
			::SysFreeString(bs);
		} else if (lstrcmp(bstrName, L"Item") == 0) {
			m_dispItem = DISPID_COLLECT;
			*pid = m_dispItem;
			hr = S_OK;
		} else if (grfdex & fdexNameCaseSensitive) {
			hr = m_pdex->GetDispID(bstrName, (grfdex & fdexNameCaseSensitive) | fdexNameCaseInsensitive, pid);
		}
	}
	return hr;
}

STDMETHODIMP CteDispatchEx::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller)
{
	if (m_dispItem && id == m_dispItem && (wFlags & DISPATCH_METHOD) && pdp && pdp->cArgs > 0) {
		return teGetPropertyAt(m_pdex, GetIntFromVariant(&pdp->rgvarg[pdp->cArgs - 1]), pvarRes);
	}
	return m_pdex->InvokeEx(id, lcid, wFlags, pdp, pvarRes, pei, pspCaller);
}

STDMETHODIMP CteDispatchEx::DeleteMemberByName(BSTR bstrName, DWORD grfdex)
{
	return m_pdex->DeleteMemberByName(bstrName, grfdex);
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
//*///

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