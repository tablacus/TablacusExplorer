#include "stdafx.h"
#include "common.h"
#if defined(_WINDLL) || defined(_EXEONLY)
#include "api.h"
#include "objects.h"

extern HWND g_hwndMain;
extern TEDispatchApi dispAPI[];
extern LPITEMIDLIST g_pidls[MAX_CSIDL2];
extern HINSTANCE	g_hShell32;
extern int g_nException;
extern LPFNRtlGetVersion _RtlGetVersion;
extern long	g_nThreads;
extern BSTR	g_bsTitle;
extern IDispatch *g_pMBText;
extern BSTR	g_bsClipRoot;
extern int		g_param[Count_TE_params];
extern IUnknown	*g_pDraggingCtrl;
extern BOOL	g_bDarkMode;
extern LPFNSHRunDialog _SHRunDialog;
extern IDispatchEx *g_pCM;
extern HHOOK	g_hMenuKeyHook;
extern DWORD	g_dwMainThreadId;
extern long	g_nProcKey;
extern DWORD	g_dwTickKey;
extern LPFNGetDpiForMonitor _GetDpiForMonitor;

#ifdef _2000XP
extern LPFNIsWow64Process _IsWow64Process;
extern LPFNPSPropertyKeyFromString _PSPropertyKeyFromStringEx;
extern LPFNSetDllDirectoryW _SetDllDirectoryW;
#else
#define _PSPropertyKeyFromStringEx tePSPropertyKeyFromStringEx
#endif
#ifdef _DEBUG
extern LPWSTR	g_strException;
#endif

extern IDropSource* teFindDropSource();
extern HRESULT MessageProc(MSG *pMsg);
extern IDispatch* teCreateObj(LONG lId, VARIANT *pvArg);
extern HRESULT ParseScript(LPOLESTR lpScript, LPOLESTR lpLang, VARIANT *pv, IDispatch **ppdisp, EXCEPINFO *pExcepInfo);//CteActiveScriptSite,CteDispatch

extern VOID teApiContextMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult);//CteContextMenu
extern VOID teApiExecScript(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult);//_beginthread
extern VOID teApiMoveWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult);//CteTreeView(XP)
extern VOID teApiPSGetDisplayName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult);//CteShellBrowser
extern VOID teApiSetFocus(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult);//CteShellBrowser
extern VOID teApiSetWindowTheme(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult);//CteShellBrowser
extern VOID teApiSHFileOperation(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult);//_beginthread
extern VOID teApiGetProcAddress(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult);//GetImage

TEmethod tesNULL[] =
{
	{ 0, NULL }
};

TEmethod tesBITMAP[] =
{
	{ (VT_PTR << TE_VT) + offsetof(BITMAP, bmBits), "bmBits" },
	{ (VT_UI2 << TE_VT) + offsetof(BITMAP, bmBitsPixel), "bmBitsPixel" },
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmHeight), "bmHeight" },
	{ (VT_UI2 << TE_VT) + offsetof(BITMAP, bmPlanes), "bmPlanes" },
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmType), "bmType" },
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmWidth), "bmWidth" },
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmWidthBytes), "bmWidthBytes" },
};

TEmethod tesCHOOSECOLOR[] =
{
	{ (VT_I4 << TE_VT) + offsetof(CHOOSECOLOR, Flags), "Flags" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, hInstance), "hInstance" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, hwndOwner), "hwndOwner" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, lCustData), "lCustData" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, lpCustColors), "lpCustColors" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, lpfnHook), "lpfnHook" },
	{ (VT_BSTR << TE_VT) + offsetof(CHOOSECOLOR, lpTemplateName), "lpTemplateName" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSECOLOR, lStructSize), "lStructSize" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSECOLOR, rgbResult), "rgbResult" },
};

TEmethod tesCHOOSEFONT[] =
{
	{ (VT_UI2 << TE_VT) + offsetof(CHOOSEFONT, ___MISSING_ALIGNMENT__), "___MISSING_ALIGNMENT__" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, Flags), "Flags" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, hDC), "hDC" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, hInstance), "hInstance" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, hwndOwner), "hwndOwner" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, iPointSize), "iPointSize" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, lCustData), "lCustData" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, lpfnHook), "lpfnHook" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, lpLogFont), "lpLogFont" },
	{ (VT_BSTR << TE_VT) + offsetof(CHOOSEFONT, lpszStyle), "lpszStyle" },
	{ (VT_BSTR << TE_VT) + offsetof(CHOOSEFONT, lpTemplateName), "lpTemplateName" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, lStructSize), "lStructSize" },
	{ (VT_UI2 << TE_VT) + offsetof(CHOOSEFONT, nFontType), "nFontType" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, nSizeMax), "nSizeMax" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, nSizeMin), "nSizeMin" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, rgbColors), "rgbColors" },
};

TEmethod tesCOPYDATASTRUCT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(COPYDATASTRUCT, cbData), "cbData" },
	{ (VT_PTR << TE_VT) + offsetof(COPYDATASTRUCT, dwData), "dwData" },
	{ (VT_PTR << TE_VT) + offsetof(COPYDATASTRUCT, lpData), "lpData" },
};

TEmethod tesDIBSECTION[] =
{
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsBitfields), "dsBitfields0" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsBitfields) + sizeof(DWORD), "dsBitfields1" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsBitfields) + sizeof(DWORD) * 2, "dsBitfields2" },
	{ (VT_PTR << TE_VT) + offsetof(DIBSECTION, dsBm), "dsBm" },
	{ (VT_PTR << TE_VT) + offsetof(DIBSECTION, dsBmih), "dsBmih" },
	{ (VT_PTR << TE_VT) + offsetof(DIBSECTION, dshSection), "dshSection" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsOffset), "dsOffset" },
};

TEmethod tesEXCEPINFO[] =
{
	{ (VT_BSTR << TE_VT) + offsetof(EXCEPINFO, bstrDescription), "bstrDescription" },
	{ (VT_BSTR << TE_VT) + offsetof(EXCEPINFO, bstrHelpFile), "bstrHelpFile" },
	{ (VT_BSTR << TE_VT) + offsetof(EXCEPINFO, bstrSource), "bstrSource" },
	{ (VT_I4 << TE_VT) + offsetof(EXCEPINFO, dwHelpContext), "dwHelpContext" },
	{ (VT_PTR << TE_VT) + offsetof(EXCEPINFO, pfnDeferredFillIn), "pfnDeferredFillIn" },
	{ (VT_PTR << TE_VT) + offsetof(EXCEPINFO, pvReserved), "pvReserved" },
	{ (VT_I4 << TE_VT) + offsetof(EXCEPINFO, scode), "scode" },
	{ (VT_UI2 << TE_VT) + offsetof(EXCEPINFO, wCode), "wCode" },
	{ (VT_UI2 << TE_VT) + offsetof(EXCEPINFO, wReserved), "wReserved" },
};

TEmethod tesFINDREPLACE[] =
{
	{ (VT_I4 << TE_VT) + offsetof(FINDREPLACE, Flags), "Flags" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, hInstance), "hInstance" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, hwndOwner), "hwndOwner" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, lCustData), "lCustData" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, lpfnHook), "lpfnHook" },
	{ (VT_BSTR << TE_VT) + offsetof(FINDREPLACE, lpstrFindWhat), "lpstrFindWhat" },
	{ (VT_BSTR << TE_VT) + offsetof(FINDREPLACE, lpstrReplaceWith), "lpstrReplaceWith" },
	{ (VT_BSTR << TE_VT) + offsetof(FINDREPLACE, lpTemplateName), "lpTemplateName" },
	{ (VT_I4 << TE_VT) + offsetof(FINDREPLACE, lStructSize), "lStructSize" },
	{ (VT_UI2 << TE_VT) + offsetof(FINDREPLACE, wFindWhatLen), "wFindWhatLen" },
	{ (VT_UI2 << TE_VT) + offsetof(FINDREPLACE, wReplaceWithLen), "wReplaceWithLen" },
};

TEmethod tesFOLDERSETTINGS[] =
{
	{ (VT_I4 << TE_VT) + offsetof(FOLDERSETTINGS, fFlags), "fFlags" },
	{ (VT_I4 << TE_VT) + (SB_IconSize - 1) * 4, "ImageSize" },
	{ (VT_I4 << TE_VT) + (SB_Options - 1) * 4, "Options" },
	{ (VT_I4 << TE_VT) + (SB_ViewFlags - 1) * 4, "ViewFlags" },
	{ (VT_I4 << TE_VT) + offsetof(FOLDERSETTINGS, ViewMode), "ViewMode" },
};

TEmethod tesHDITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, cchTextMax), "cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, cxy), "cxy" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, fmt), "fmt" },
	{ (VT_PTR << TE_VT) + offsetof(HDITEM, hbm), "hbm" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, iImage), "iImage" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, iOrder), "iOrder" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, lParam), "lParam" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, mask), "mask" },
	{ (VT_BSTR << TE_VT) + offsetof(HDITEM, pszText), "pszText" },
	{ (VT_PTR << TE_VT) + offsetof(HDITEM, pvFilter), "pvFilter" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, state), "state" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, type), "type" },
};

TEmethod tesICONINFO[] =
{
	{ (VT_BOOL << TE_VT) + offsetof(ICONINFO, fIcon), "fIcon" },
	{ (VT_PTR << TE_VT) + offsetof(ICONINFO, hbmColor), "hbmColor" },
	{ (VT_PTR << TE_VT) + offsetof(ICONINFO, hbmMask), "hbmMask" },
	{ (VT_I4 << TE_VT) + offsetof(ICONINFO, xHotspot), "xHotspot" },
	{ (VT_I4 << TE_VT) + offsetof(ICONINFO, yHotspot), "yHotspot" },
};

TEmethod tesICONMETRICS[] =
{
	{ (VT_I4 << TE_VT) + offsetof(ICONMETRICS, iHorzSpacing), "iHorzSpacing" },
	{ (VT_I4 << TE_VT) + offsetof(ICONMETRICS, iTitleWrap), "iTitleWrap" },
	{ (VT_I4 << TE_VT) + offsetof(ICONMETRICS, iVertSpacing), "iVertSpacing" },
	{ (VT_PTR << TE_VT) + offsetof(ICONMETRICS, lfFont), "lfFont" },
};

TEmethod tesKEYBDINPUT[] =
{
	{ (VT_PTR << TE_VT) + offsetof(KEYBDINPUT, dwExtraInfo), "dwExtraInfo" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, dwFlags), "dwFlags" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, time), "time" },
	{ (VT_I4 << TE_VT), "type" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, wScan), "wScan" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, wVk), "wVk" },
};

TEmethod tesLOGFONT[] =
{
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfCharSet), "lfCharSet" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfClipPrecision), "lfClipPrecision" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfEscapement), "lfEscapement" },
	{ (VT_LPWSTR << TE_VT) + offsetof(LOGFONT, lfFaceName), "lfFaceName" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfHeight), "lfHeight" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfItalic), "lfItalic" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfOrientation), "lfOrientation" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfOutPrecision), "lfOutPrecision" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfPitchAndFamily), "lfPitchAndFamily" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfQuality), "lfQuality" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfStrikeOut), "lfStrikeOut" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfUnderline), "lfUnderline" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfWeight), "lfWeight" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfWidth), "lfWidth" },
};

TEmethod tesLVBKIMAGE[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, cchImageMax), "cchImageMax" },
	{ (VT_PTR << TE_VT) + offsetof(LVBKIMAGE, hbm), "hbm" },
	{ (VT_BSTR << TE_VT) + offsetof(LVBKIMAGE, pszImage), "pszImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, ulFlags), "ulFlags" },
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, xOffsetPercent), "xOffsetPercent" },
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, yOffsetPercent), "yOffsetPercent" },
};

TEmethod tesLVCOLUMN[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVCOLUMN, cchTextMax), "cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(LVCOLUMN, cx), "cx" },
	{ (VT_I4 << TE_VT) + offsetof(LVCOLUMN, cxDefault), "cxDefault" },
	{ (VT_I4 << TE_VT) + offsetof(LVCOLUMN, cxIdeal), "cxIdeal" },
	{ (VT_I4 << TE_VT) + offsetof(LVCOLUMN, cxMin), "cxMin" },
	{ (VT_I4 << TE_VT) + offsetof(LVCOLUMN, fmt), "fmt" },
	{ (VT_I4 << TE_VT) + offsetof(LVCOLUMN, iImage), "iImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVCOLUMN, iOrder), "iOrder" },
	{ (VT_I4 << TE_VT) + offsetof(LVCOLUMN, iSubItem), "iSubItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVCOLUMN, mask), "mask" },
	{ (VT_BSTR << TE_VT) + offsetof(LVCOLUMN, pszText), "pszText" },
};

TEmethod tesLVFINDINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVFINDINFO, flags), "flags" },
	{ (VT_I4 << TE_VT) + offsetof(LVFINDINFO, lParam), "lParam" },
	{ (VT_BSTR << TE_VT) + offsetof(LVFINDINFO, psz), "psz" },
	{ (VT_CY << TE_VT) + offsetof(LVFINDINFO, pt), "pt" },
	{ (VT_I4 << TE_VT) + offsetof(LVFINDINFO, vkDirection), "vkDirection" },
};

TEmethod tesLVGROUP[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchDescriptionBottom), "cchDescriptionBottom" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchDescriptionTop), "cchDescriptionTop" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchFooter), "cchFooter" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchHeader), "cchHeader" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchSubsetTitle), "cchSubsetTitle" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchSubtitle), "cchSubtitle" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchTask), "cchTask" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cItems), "cItems" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iExtendedImage), "iExtendedImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iFirstItem), "iFirstItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iGroupId), "iGroupId" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iTitleImage), "iTitleImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, mask), "mask" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszDescriptionBottom), "pszDescriptionBottom" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszDescriptionTop), "pszDescriptionTop" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszFooter), "pszFooter" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszHeader), "pszHeader" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszSubsetTitle), "pszSubsetTitle" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszSubtitle), "pszSubtitle" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszTask), "pszTask" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, state), "state" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, stateMask), "stateMask" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, uAlign), "uAlign" },
};

TEmethod tesLVHITTESTINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, flags), "flags" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, iGroup), "iGroup" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, iItem), "iItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, iSubItem ), "iSubItem" },
	{ (VT_CY << TE_VT) + offsetof(LVHITTESTINFO, pt), "pt" },
};

TEmethod tesLVITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, cchTextMax), "cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, cColumns), "cColumns" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iGroup), "iGroup" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iGroupId), "iGroupId" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iImage), "iImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iIndent), "iIndent" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iItem), "iItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iSubItem), "iSubItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, lParam), "lParam" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, mask), "mask" },
	{ (VT_PTR << TE_VT) + offsetof(LVITEM, piColFmt), "piColFmt" },
	{ (VT_BSTR << TE_VT) + offsetof(LVITEM, pszText), "pszText" },
	{ (VT_PTR << TE_VT) + offsetof(LVITEM, puColumns), "puColumns" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, state), "state" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, stateMask), "stateMask" },
};

TEmethod tesMENUITEMINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, cch), "cch" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, dwItemData), "dwItemData" },
	{ (VT_BSTR << TE_VT) + offsetof(MENUITEMINFO, dwTypeData), "dwTypeData" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, fMask), "fMask" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, fState), "fState" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, fType), "fType" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hbmpChecked), "hbmpChecked" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hbmpItem), "hbmpItem" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hbmpUnchecked), "hbmpUnchecked" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hSubMenu), "hSubMenu" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, wID), "wID" },
};

TEmethod tesMONITORINFOEX[] =
{
	{ (VT_I4 << TE_VT), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(MONITORINFOEX, dwFlags), "dwFlags" },
	{ (VT_CARRAY << TE_VT) + offsetof(MONITORINFOEX, rcMonitor), "rcMonitor" },
	{ (VT_CARRAY << TE_VT) + offsetof(MONITORINFOEX, rcWork), "rcWork" },
	{ (VT_LPWSTR << TE_VT) + offsetof(MONITORINFOEX, szDevice), "szDevice" },
};


TEmethod tesMOUSEINPUT[] =
{
	{ (VT_PTR << TE_VT) + offsetof(MOUSEINPUT, dwExtraInfo), "dwExtraInfo" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, dwFlags), "dwFlags" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, dx), "dx" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, dy), "dy" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, mouseData), "mouseData" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, time), "time" },
	{ (VT_I4 << TE_VT), "type" },
};

TEmethod tesMSG[] =
{
	{ (VT_PTR << TE_VT) + offsetof(MSG, hwnd), "hwnd" },
	{ (VT_PTR << TE_VT) + offsetof(MSG, lParam), "lParam" },
	{ (VT_I4 << TE_VT) + offsetof(MSG, message), "message" },
	{ (VT_CY << TE_VT) + offsetof(MSG, pt), "pt" },
	{ (VT_I4 << TE_VT) + offsetof(MSG, time), "time" },
	{ (VT_PTR << TE_VT) + offsetof(MSG, wParam), "wParam" },
};

TEmethod tesNONCLIENTMETRICS[] =
{
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iBorderWidth), "iBorderWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iCaptionHeight), "iCaptionHeight" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iCaptionWidth), "iCaptionWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iMenuHeight), "iMenuHeight" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iMenuWidth), "iMenuWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iPaddedBorderWidth), "iPaddedBorderWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iScrollHeight), "iScrollHeight" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iScrollWidth), "iScrollWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iSmCaptionHeight), "iSmCaptionHeight" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iSmCaptionWidth), "iSmCaptionWidth" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfCaptionFont), "lfCaptionFont" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfMenuFont), "lfMenuFont" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfMessageFont), "lfMessageFont" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfSmCaptionFont), "lfSmCaptionFont" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfStatusFont), "lfStatusFont" },
};

TEmethod tesNOTIFYICONDATA[] =
{
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwInfoFlags), "dwInfoFlags" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwState), "dwState" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwStateMask), "dwStateMask" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, guidItem), "guidItem" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, hBalloonIcon), "hBalloonIcon" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, hIcon), "hIcon" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, hWnd), "hWnd" },
	{ (VT_LPWSTR << TE_VT) + offsetof(NOTIFYICONDATA, szInfo), "szInfo" },
	{ (VT_LPWSTR << TE_VT) + offsetof(NOTIFYICONDATA, szInfoTitle), "szInfoTitle" },
	{ (VT_LPWSTR << TE_VT) + offsetof(NOTIFYICONDATA, szTip), "szTip" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uCallbackMessage), "uCallbackMessage" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uFlags), "uFlags" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uID), "uID" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uTimeout), "uTimeout" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uVersion), "uVersion" },
};

TEmethod tesNMCUSTOMDRAW[] =
{
	{ (VT_I4 << TE_VT) + offsetof(NMCUSTOMDRAW, dwDrawStage), "dwDrawStage" },
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, dwItemSpec), "dwItemSpec" },
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, hdc), "hdc" },
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, hdr), "hdr" },
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, lItemlParam), "lItemlParam" },
	{ (VT_CARRAY << TE_VT) + offsetof(NMCUSTOMDRAW, rc), "rc" },
	{ (VT_I4 << TE_VT) + offsetof(NMCUSTOMDRAW, uItemState), "uItemState" },
};

TEmethod tesNMLVCUSTOMDRAW[] =
{
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, clrFace), "clrFace" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, clrText), "clrText" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, clrTextBk), "clrTextBk" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, dwItemType), "dwItemType" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iIconEffect), "iIconEffect" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iIconPhase), "iIconPhase" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iPartId), "iPartId" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iStateId), "iStateId" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iSubItem), "iSubItem" },
	{ (VT_PTR << TE_VT) + offsetof(NMLVCUSTOMDRAW, nmcd), "nmcd" },
	{ (VT_CARRAY << TE_VT) + offsetof(NMLVCUSTOMDRAW, rcText), "rcText" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, uAlign), "uAlign" },
};

TEmethod tesNMTVCUSTOMDRAW[] =
{
	{ (VT_I4 << TE_VT) + offsetof(NMTVCUSTOMDRAW, clrText), "clrText" },
	{ (VT_I4 << TE_VT) + offsetof(NMTVCUSTOMDRAW, clrTextBk), "clrTextBk" },
	{ (VT_I4 << TE_VT) + offsetof(NMTVCUSTOMDRAW, iLevel), "iLevel" },
	{ (VT_PTR << TE_VT) + offsetof(NMTVCUSTOMDRAW, nmcd), "nmcd" },
};

TEmethod tesNMHDR[] =
{
	{ (VT_I4 << TE_VT) + offsetof(NMHDR, code), "code" },
	{ (VT_PTR << TE_VT) + offsetof(NMHDR, hwndFrom), "hwndFrom" },
	{ (VT_I4 << TE_VT) + offsetof(NMHDR, idFrom), "idFrom" },
};

TEmethod tesOSVERSIONINFOEX[] =
{
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwBuildNumber), "dwBuildNumber" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwMajorVersion), "dwMajorVersion" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwMinorVersion), "dwMinorVersion" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwOSVersionInfoSize), "dwOSVersionInfoSize" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwPlatformId), "dwPlatformId" },
	{ (VT_LPWSTR << TE_VT) + offsetof(OSVERSIONINFOEX, szCSDVersion), "szCSDVersion" },
	{ (VT_UI1 << TE_VT) + offsetof(OSVERSIONINFOEX, wProductType), "wProductType" },
	{ (VT_UI1 << TE_VT) + offsetof(OSVERSIONINFOEX, wReserved), "wReserved" },
	{ (VT_UI2 << TE_VT) + offsetof(OSVERSIONINFOEX, wServicePackMajor), "wServicePackMajor" },
	{ (VT_UI2 << TE_VT) + offsetof(OSVERSIONINFOEX, wServicePackMinor), "wServicePackMinor" },
	{ (VT_UI2 << TE_VT) + offsetof(OSVERSIONINFOEX, wSuiteMask), "wSuiteMask" },
};

TEmethod tesPAINTSTRUCT[] =
{
	{ (VT_BOOL << TE_VT) + offsetof(PAINTSTRUCT, fErase), "fErase" },
	{ (VT_BOOL << TE_VT) + offsetof(PAINTSTRUCT, fIncUpdate), "fIncUpdate" },
	{ (VT_BOOL << TE_VT) + offsetof(PAINTSTRUCT, fRestore), "fRestore" },
	{ (VT_PTR << TE_VT) + offsetof(PAINTSTRUCT, hdc), "hdc" },
	{ (VT_CARRAY << TE_VT) + offsetof(PAINTSTRUCT, rcPaint), "rcPaint" },
	{ (VT_UI1 << TE_VT) + offsetof(PAINTSTRUCT, rgbReserved), "rgbReserved" },
};

TEmethod tesPOINT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(POINT, x), "x" },
	{ (VT_I4 << TE_VT) + offsetof(POINT, y), "y" },
};

TEmethod tesRECT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(RECT, bottom), "bottom" },
	{ (VT_I4 << TE_VT) + offsetof(RECT, left), "left" },
	{ (VT_I4 << TE_VT) + offsetof(RECT, right), "right" },
	{ (VT_I4 << TE_VT) + offsetof(RECT, top), "top" },
};

TEmethod tesSHELLEXECUTEINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, dwHotKey), "dwHotKey" },
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, fMask), "fMask" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hIcon), "hIcon" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hInstApp), "hInstApp" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hkeyClass), "hkeyClass" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hMonitor), "hMonitor" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hProcess), "hProcess" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hwnd), "hwnd" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpClass), "lpClass" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpDirectory), "lpDirectory" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpFile), "lpFile" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpIDList), "lpIDList" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpParameters), "lpParameters" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpVerb), "lpVerb" },
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, nShow), "nShow" },
};

TEmethod tesSHFILEINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(SHFILEINFO, dwAttributes), "dwAttributes" },
	{ (VT_PTR << TE_VT) + offsetof(SHFILEINFO, hIcon), "hIcon" },
	{ (VT_I4 << TE_VT) + offsetof(SHFILEINFO, iIcon), "iIcon" },
	{ (VT_LPWSTR << TE_VT) + offsetof(SHFILEINFO, szDisplayName), "szDisplayName" },
	{ (VT_LPWSTR << TE_VT) + offsetof(SHFILEINFO, szTypeName), "szTypeName" },
};

TEmethod tesSHFILEOPSTRUCT[] =
{
	{ (VT_BOOL << TE_VT) + offsetof(SHFILEOPSTRUCT, fAnyOperationsAborted), "fAnyOperationsAborted" },
	{ (VT_UI2 << TE_VT) + offsetof(SHFILEOPSTRUCT, fFlags), "fFlags" },
	{ (VT_PTR << TE_VT) + offsetof(SHFILEOPSTRUCT, hNameMappings), "hNameMappings" },
	{ (VT_PTR << TE_VT) + offsetof(SHFILEOPSTRUCT, hwnd), "hwnd" },
	{ (VT_BSTR << TE_VT) + offsetof(SHFILEOPSTRUCT, lpszProgressTitle), "lpszProgressTitle" },
	{ (VT_BSTR << TE_VT) + offsetof(SHFILEOPSTRUCT, pFrom), "pFrom" },
	{ (VT_BSTR << TE_VT) + offsetof(SHFILEOPSTRUCT, pTo), "pTo" },
	{ (VT_I4 << TE_VT) + offsetof(SHFILEOPSTRUCT, wFunc), "wFunc" },
};

TEmethod tesSIZE[] =
{
	{ (VT_I4 << TE_VT) + offsetof(SIZE, cx), "cx" },
	{ (VT_I4 << TE_VT) + offsetof(SIZE, cy), "cy" },
};

TEmethod tesTCHITTESTINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(TCHITTESTINFO, flags), "flags" },
	{ (VT_CY << TE_VT) + offsetof(TCHITTESTINFO, pt), "pt" },
};

TEmethod tesTCITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, cchTextMax), "cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, dwState), "dwState" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, dwStateMask), "dwStateMask" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, iImage), "iImage" },
	{ (VT_PTR << TE_VT) + offsetof(TCITEM, lParam), "lParam" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, mask), "mask" },
	{ (VT_BSTR << TE_VT) + offsetof(TCITEM, pszText), "pszText" },
};

TEmethod tesTOOLINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(TOOLINFO, cbSize), "cbSize" },
	{ (VT_PTR << TE_VT) + offsetof(TOOLINFO, hinst), "hinst" },
	{ (VT_PTR << TE_VT) + offsetof(TOOLINFO, hwnd), "hwnd" },
	{ (VT_PTR << TE_VT) + offsetof(TOOLINFO, uId), "lParam" },
	{ (VT_PTR << TE_VT) + offsetof(TOOLINFO, uId), "lpReserved" },
	{ (VT_BSTR << TE_VT) + offsetof(TOOLINFO, uId), "lpszText" },
	{ (VT_CARRAY << TE_VT) + offsetof(TOOLINFO, rect), "rect" },
	{ (VT_I4 << TE_VT) + offsetof(TOOLINFO, uFlags), "uFlags" },
	{ (VT_PTR << TE_VT) + offsetof(TOOLINFO, uId), "uId" },
};

TEmethod tesTVHITTESTINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(TVHITTESTINFO, flags), "flags" },
	{ (VT_PTR << TE_VT) + offsetof(TVHITTESTINFO, hItem), "hItem" },
	{ (VT_CY << TE_VT) + offsetof(TVHITTESTINFO, pt), "pt" },
};

TEmethod tesTVITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, cChildren), "cChildren" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, cchTextMax), "cchTextMax" },
	{ (VT_PTR << TE_VT) + offsetof(TVITEM, hItem), "hItem" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, iImage), "iImage" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, iSelectedImage), "iSelectedImage" },
	{ (VT_PTR << TE_VT) + offsetof(TVITEM, lParam), "lParam" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, mask), "mask" },
	{ (VT_BSTR << TE_VT) + offsetof(TVITEM, pszText), "pszText" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, state), "state" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, stateMask), "stateMask" },
};

TEmethod tesWIN32_FIND_DATA[] =
{
	{ (VT_LPWSTR << TE_VT) + offsetof(WIN32_FIND_DATA, cAlternateFileName), "cAlternateFileName" },
	{ (VT_LPWSTR << TE_VT) + offsetof(WIN32_FIND_DATA, cFileName), "cFileName" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, dwFileAttributes), "dwFileAttributes" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, dwReserved0), "dwReserved0" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, dwReserved1), "dwReserved1" },
	{ (VT_FILETIME << TE_VT) + offsetof(WIN32_FIND_DATA, ftCreationTime), "ftCreationTime" },
	{ (VT_FILETIME << TE_VT) + offsetof(WIN32_FIND_DATA, ftLastAccessTime), "ftLastAccessTime" },
	{ (VT_FILETIME << TE_VT) + offsetof(WIN32_FIND_DATA, ftLastWriteTime), "ftLastWriteTime" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, nFileSizeHigh), "nFileSizeHigh" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, nFileSizeLow), "nFileSizeLow" },
};

TEmethod tesDRAWITEMSTRUCT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, CtlID), "CtlID" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, CtlType), "CtlType" },
	{ (VT_PTR << TE_VT) + offsetof(DRAWITEMSTRUCT, hDC), "hDC" },
	{ (VT_PTR << TE_VT) + offsetof(DRAWITEMSTRUCT, hwndItem), "hwndItem" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, itemAction), "itemAction" },
	{ (VT_PTR << TE_VT) + offsetof(DRAWITEMSTRUCT, itemData), "itemData" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, itemID), "itemID" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, itemState), "itemState" },
	{ (VT_CARRAY << TE_VT) + offsetof(DRAWITEMSTRUCT, rcItem), "rcItem" },
};

TEmethod tesMEASUREITEMSTRUCT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, CtlID), "CtlID" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, CtlType), "CtlType" },
	{ (VT_PTR << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemData), "itemData" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemHeight), "itemHeight" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemID), "itemID" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemWidth), "itemWidth" },
};

TEmethod tesMENUINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, cyMax), "cyMax" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, dwContextHelpID), "dwContextHelpID" },
	{ (VT_PTR << TE_VT) + offsetof(MENUINFO, dwMenuData), "dwMenuData" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, dwStyle), "dwStyle" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, fMask), "fMask" },
	{ (VT_PTR << TE_VT) + offsetof(MENUINFO, hbrBack), "hbrBack" },
};

TEmethod tesGUITHREADINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(GUITHREADINFO, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(GUITHREADINFO, flags), "flags" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndActive), "hwndActive" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndCapture), "hwndCapture" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndCaret), "hwndCaret" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndFocus), "hwndFocus" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndMenuOwner), "hwndMenuOwner" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndMoveSize), "hwndMoveSize" },
	{ (VT_CARRAY << TE_VT) + offsetof(GUITHREADINFO, hwndCaret), "rcCaret" },
};

TEStruct pTEStructs[] = {
	{ sizeof(BITMAP), FALSE, _countof(tesBITMAP), "BITMAP", tesBITMAP },
	{ sizeof(BSTR), FALSE, 0, "BSTR", tesNULL },
	{ sizeof(BYTE), FALSE, 0, "BYTE", tesNULL },
	{ sizeof(char), FALSE, 0, "char", tesNULL },
	{ sizeof(CHOOSECOLOR), TRUE, _countof(tesCHOOSECOLOR), "CHOOSECOLOR", tesCHOOSECOLOR },
	{ sizeof(CHOOSEFONT), TRUE, _countof(tesCHOOSEFONT), "CHOOSEFONT", tesCHOOSEFONT },
	{ sizeof(COPYDATASTRUCT), FALSE, _countof(tesCOPYDATASTRUCT), "COPYDATASTRUCT", tesCOPYDATASTRUCT },
	{ sizeof(DIBSECTION), FALSE, _countof(tesDIBSECTION), "DIBSECTION", tesDIBSECTION },
	{ sizeof(DRAWITEMSTRUCT), FALSE, _countof(tesDRAWITEMSTRUCT), "DRAWITEMSTRUCT", tesDRAWITEMSTRUCT },
	{ sizeof(DWORD), FALSE, 0, "DWORD", tesNULL },
	{ sizeof(EXCEPINFO), FALSE, _countof(tesEXCEPINFO), "EXCEPINFO", tesEXCEPINFO },
	{ sizeof(FINDREPLACE), TRUE, _countof(tesFINDREPLACE), "FINDREPLACE", tesFINDREPLACE },
	{ sizeof(FOLDERSETTINGS), FALSE, _countof(tesFOLDERSETTINGS), "FOLDERSETTINGS", tesFOLDERSETTINGS },
	{ sizeof(GUID), FALSE, 0, "GUID", tesNULL },
	{ sizeof(GUITHREADINFO), TRUE, _countof(tesGUITHREADINFO), "GUITHREADINFO", tesGUITHREADINFO },
	{ sizeof(HANDLE), FALSE, 0, "HANDLE", tesNULL },
	{ sizeof(HDITEM), FALSE, _countof(tesHDITEM), "HDITEM", tesHDITEM },
	{ sizeof(ICONINFO), FALSE, _countof(tesICONINFO), "ICONINFO", tesICONINFO },
	{ sizeof(ICONMETRICS), FALSE, _countof(tesICONMETRICS), "ICONMETRICS", tesICONMETRICS },
	{ sizeof(int), FALSE, 0, "int", tesNULL },
	{ sizeof(KEYBDINPUT) + sizeof(DWORD), FALSE, _countof(tesKEYBDINPUT), "KEYBDINPUT", tesKEYBDINPUT },
	{ 256, FALSE, 0, "KEYSTATE", tesNULL },
	{ sizeof(LOGFONT), FALSE, _countof(tesLOGFONT), "LOGFONT", tesLOGFONT },
	{ sizeof(LPWSTR), FALSE, 0, "LPWSTR", tesNULL },
	{ sizeof(LVBKIMAGE), FALSE, _countof(tesLVBKIMAGE), "LVBKIMAGE", tesLVBKIMAGE },
	{ sizeof(LVCOLUMN), FALSE, _countof(tesLVCOLUMN), "LVCOLUMN", tesLVCOLUMN },
	{ sizeof(LVFINDINFO), FALSE, _countof(tesLVFINDINFO), "LVFINDINFO", tesLVFINDINFO },
	{ sizeof(LVGROUP), TRUE, _countof(tesLVGROUP), "LVGROUP", tesLVGROUP },
	{ sizeof(LVHITTESTINFO), FALSE, _countof(tesLVHITTESTINFO), "LVHITTESTINFO", tesLVHITTESTINFO },
	{ sizeof(LVITEM), FALSE, _countof(tesLVITEM), "LVITEM", tesLVITEM },
	{ sizeof(MEASUREITEMSTRUCT), FALSE, _countof(tesMEASUREITEMSTRUCT), "MEASUREITEMSTRUCT", tesMEASUREITEMSTRUCT },
	{ sizeof(MENUINFO), TRUE, _countof(tesMENUINFO), "MENUINFO", tesMENUINFO },
	{ sizeof(MENUITEMINFO), TRUE, _countof(tesMENUITEMINFO), "MENUITEMINFO", tesMENUITEMINFO },
	{ sizeof(MONITORINFOEX), TRUE, _countof(tesMONITORINFOEX), "MONITORINFOEX", tesMONITORINFOEX },
	{ sizeof(MOUSEINPUT) + sizeof(DWORD), FALSE, _countof(tesMOUSEINPUT), "MOUSEINPUT", tesMOUSEINPUT },
	{ sizeof(MSG), FALSE, _countof(tesMSG), "MSG", tesMSG },
	{ sizeof(NMCUSTOMDRAW), FALSE, _countof(tesNMCUSTOMDRAW), "NMCUSTOMDRAW", tesNMCUSTOMDRAW },
	{ sizeof(NMHDR), FALSE, _countof(tesNMHDR), "NMHDR", tesNMHDR },
	{ sizeof(NMLVCUSTOMDRAW), FALSE, _countof(tesNMLVCUSTOMDRAW), "NMLVCUSTOMDRAW", tesNMLVCUSTOMDRAW },
	{ sizeof(NMTVCUSTOMDRAW), FALSE, _countof(tesNMTVCUSTOMDRAW), "NMTVCUSTOMDRAW", tesNMTVCUSTOMDRAW },
	{ sizeof(NONCLIENTMETRICS), TRUE, _countof(tesNONCLIENTMETRICS), "NONCLIENTMETRICS", tesNONCLIENTMETRICS },
	{ sizeof(NOTIFYICONDATA), TRUE, _countof(tesNOTIFYICONDATA), "NOTIFYICONDATA", tesNOTIFYICONDATA },
	{ sizeof(OSVERSIONINFO), TRUE, _countof(tesOSVERSIONINFOEX), "OSVERSIONINFO", tesOSVERSIONINFOEX },
	{ sizeof(OSVERSIONINFOEX), TRUE, _countof(tesOSVERSIONINFOEX), "OSVERSIONINFOEX", tesOSVERSIONINFOEX },
	{ sizeof(PAINTSTRUCT), FALSE, _countof(tesPAINTSTRUCT), "PAINTSTRUCT", tesPAINTSTRUCT },
	{ sizeof(POINT), FALSE, _countof(tesPOINT), "POINT", tesPOINT },
	{ sizeof(RECT), FALSE, _countof(tesRECT), "RECT", tesRECT },
	{ sizeof(SHELLEXECUTEINFO), TRUE, _countof(tesSHELLEXECUTEINFO), "SHELLEXECUTEINFO", tesSHELLEXECUTEINFO },
	{ sizeof(SHFILEINFO), FALSE, _countof(tesSHFILEINFO), "SHFILEINFO", tesSHFILEINFO },
	{ sizeof(SHFILEOPSTRUCT), FALSE, _countof(tesSHFILEOPSTRUCT), "SHFILEOPSTRUCT", tesSHFILEOPSTRUCT },
	{ sizeof(SIZE), FALSE, _countof(tesSIZE), "SIZE", tesSIZE },
	{ sizeof(TCHITTESTINFO), FALSE, _countof(tesTCHITTESTINFO), "TCHITTESTINFO", tesTCHITTESTINFO },
	{ sizeof(TCITEM), FALSE, _countof(tesTCITEM), "TCITEM", tesTCITEM },
	{ sizeof(TOOLINFO), TRUE, _countof(tesTOOLINFO), "TOOLINFO", tesTOOLINFO },
	{ sizeof(TVHITTESTINFO), FALSE, _countof(tesTVHITTESTINFO), "TVHITTESTINFO", tesTVHITTESTINFO },
	{ sizeof(TVITEM), FALSE, _countof(tesTVITEM), "TVITEM", tesTVITEM },
	{ sizeof(VARIANT), FALSE, 0, "VARIANT", tesNULL },
	{ sizeof(WCHAR), FALSE, 0, "WCHAR", tesNULL },
	{ sizeof(WIN32_FIND_DATA), FALSE, _countof(tesWIN32_FIND_DATA), "WIN32_FIND_DATA", tesWIN32_FIND_DATA },
	{ sizeof(WORD), FALSE, 0, "WORD", tesNULL },
};

TEmethod methodObject[] = {
	{ 4, "api" },
	{ 2, "Array" },
	{ 5, "CommonDialog" },
	{ 3, "Enum" },
	{ 6, "FolderItems" },
	{ 1, "Object" },
	{ 7, "ProgressDialog" },
	{ -2, "SafeArray" },
	{ 8, "WICBitmap" },
};

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
#ifdef CHECK_HANDLELEAK
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
#ifdef CHECK_HANDLELEAK
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
					MessageSub(TE_OnKeyMessage, teFindDropSource(), &msg, (HRESULT *)&lResult);
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

///Win32API Functions
VOID teApiAllowSetForegroundWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, AllowSetForegroundWindow(param[0].dword));
}

VOID teApiAlphaBlend(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, AlphaBlend(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal,
		param[5].hdc, param[6].intVal, param[7].intVal, param[8].intVal, param[9].intVal, *(BLENDFUNCTION*)&param[10].lVal));
}

VOID teApiArray(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetObjectRelease(pVarResult, teCreateObj(TE_ARRAY, NULL));
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

VOID teApiBeginPaint(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, BeginPaint(param[0].hwnd, param[1].lppaintstruct));
}

VOID teApiBitBlt(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, BitBlt(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal,
		param[5].hdc, param[6].intVal, param[7].intVal, param[8].dword));
}

VOID teApiBringWindowToTop(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, BringWindowToTop(param[0].hwnd));
}

VOID teApiCallWindowProc(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CallWindowProc(param[0].wndproc, param[1].hwnd, param[2].uintVal, param[3].wparam, param[4].lparam));
}

VOID teApiChangeWindowMessageFilterEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teChangeWindowMessageFilterEx(param[0].hwnd, param[1].uintVal, param[2].dword, param[3].pchangefilterstruct));
}

VOID teApiChooseColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ::ChooseColor(param[0].lpchoosecolor));
}

VOID teApiChooseFont(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ChooseFont(param[0].lpchoosefont));
}

VOID teApiClientToScreen(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg - 1]);
	teSetBool(pVarResult, ClientToScreen(param[0].hwnd, &pt));
	PutPointToVariant(&pt, &pDispParams->rgvarg[nArg - 1]);
}

VOID teApiCloseHandle(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, CloseHandle(param[0].handle));
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

VOID teApiCopyImage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CopyImage(param[0].handle, param[1].uintVal, param[2].intVal, param[3].intVal, param[4].uintVal));
}

VOID teApiCoTaskMemFree(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teCoTaskMemFree(param[0].lpvoid);
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

VOID teApiCreateCompatibleBitmap(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateCompatibleBitmap(param[0].hdc, param[1].intVal, param[2].intVal));
}

VOID teApiCreateCompatibleDC(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateCompatibleDC(param[0].hdc));
}

VOID teApiCreateEvent(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateEvent(param[0].lpSecurityAttributes, param[1].boolVal, param[2].boolVal, param[3].lpcwstr));
}

VOID teApiCreateFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::CreateFile(param[0].lpcwstr, param[1].dword, param[2].dword, param[3].lpSecurityAttributes, param[4].dword, param[5].dword, param[6].handle));
}

VOID teApiCreateFontIndirect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::CreateFontIndirect(param[0].plogfont));
}

VOID teApiCreateIconIndirect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::CreateIconIndirect(param[0].piconinfo));
}

VOID teApiCreateMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teCreateMenu(CreateMenu(), pVarResult);
}

VOID teApiCreateObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
#ifdef _DEBUG
	//::OutputDebugStringA("CreateObject: ");
	//::OutputDebugString(param[0].lpwstr);
	//::OutputDebugStringA("\n");
#endif
	int nIndex = teBSearch(methodObject, _countof(methodObject), param[0].lpwstr);
	if (nIndex >= 0) {
		int id = methodObject[nIndex].id;
		if (id >= 0) {
			teSetObjectRelease(pVarResult, teCreateObj(id, nArg >= 1 ? &pDispParams->rgvarg[nArg - 1] : NULL));
			return;
		}
		if (id == -2 && nArg >= 1 && pVarResult) {
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

VOID teApiCreatePen(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreatePen(param[0].intVal, param[1].intVal, param[2].colorref));
}

VOID teApiCreatePopupMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teCreateMenu(CreatePopupMenu(), pVarResult);
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
		if (pVarResult && WaitForInputIdle(pi.hProcess, dwms)) {
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

VOID teApiCreateSolidBrush(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateSolidBrush(param[0].colorref));
}

VOID teApiCreateWindowEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, CreateWindowEx(param[0].dword, param[1].lpcwstr, param[2].lpcwstr,
		param[3].dword, param[4].intVal, param[5].intVal, param[6].intVal, param[7].intVal,
		param[8].hwnd, param[9].hmenu, param[10].hinstance, param[11].lpvoid));
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

VOID teApiDataObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetObjectRelease(pVarResult, new CteFolderItems(NULL, NULL));
}

VOID teApiDeleteDC(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteDC(param[0].hdc));
}

VOID teApiDeleteFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteFile(param[0].lpcwstr));
}

VOID teApiDeleteMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteMenu(param[0].hmenu, param[1].uintVal, param[2].uintVal));
}

VOID teApiDeleteObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DeleteObject(param[0].hgdiobj));
}

VOID teApiDestroyCursor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyCursor(param[0].hcursor));
}

VOID teApiDestroyIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyIcon(param[0].hicon));
}

VOID teApiDestroyMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyMenu(param[0].hmenu));
}

VOID teApiDestroyWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DestroyWindow(param[0].hwnd));
}

VOID teApiDispatchMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::DispatchMessage(param[0].lpmsg));
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

VOID teApiDragAcceptFiles(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DragAcceptFiles(param[0].hwnd, param[1].boolVal);
}

VOID teApiDragFinish(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	DragFinish(param[0].hdrop);
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

VOID teApiDrawIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DrawIcon(param[0].hdc, param[1].intVal, param[2].intVal, param[3].hicon));
}

VOID teApiDrawIconEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, DrawIconEx(param[0].hdc, param[1].intVal, param[2].intVal, param[3].hicon,
		param[4].intVal, param[5].intVal, param[6].uintVal, param[7].hbrush, param[8].uintVal));
}

VOID teApiDrawText(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, DrawText(param[0].hdc, param[1].lpcwstr, param[2].intVal, param[3].lprect, param[4].uintVal));
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

VOID teApiEnableMenuItem(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, EnableMenuItem(param[0].hmenu, param[1].uintVal, param[2].uintVal));
}

VOID teApiEnableWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, EnableWindow(param[0].hwnd, param[1].boolVal));
}

VOID teApiEndMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, EndMenu());
}

VOID teApiEndPaint(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, EndPaint(param[0].hwnd, param[1].lppaintstruct));
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
		teStopProgressDialog(ppd);
		SafeRelease(&ppd);
	}
	teSetLong(pVarResult, hr);
}

VOID teApiExtractIconEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ExtractIconEx(param[0].lpcwstr, param[1].intVal, param[2].phicon, param[3].phicon, param[4].uintVal));
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

VOID teApiFindClose(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, FindClose(param[0].handle));
}

VOID teApiFindFirstFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, FindFirstFile(param[0].lpcwstr, param[1].lpwin32_find_data));
}

VOID teApiFindNextFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, FindNextFile(param[0].handle, param[1].lpwin32_find_data));
}

VOID teApiFindReplace(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::ReplaceText(param[0].lpfindreplace));
}

VOID teApiFindText(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ::FindText(param[0].lpfindreplace));
}

VOID teApiFindWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, FindWindow(param[0].lpcwstr, param[1].lpcwstr));
}

VOID teApiFindWindowEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, FindWindowEx(param[0].hwnd, param[1].hwnd, param[2].lpcwstr, param[3].lpcwstr));
}

VOID teApiFormatMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPWSTR lpBuf;
	DWORD dwFlags;
	LPWSTR lpSource;
	va_list *lpList;
	if (nArg >= 1) {
		dwFlags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ARGUMENT_ARRAY;
		lpSource = GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]);
		lpList = (va_list *)&param[1];
	} else {
		dwFlags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM;
		lpSource = NULL;
		lpList = NULL;
	}
	if (FormatMessage(dwFlags, lpSource, param[0].intVal, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&lpBuf, SIZE_BUFF, lpList)) {
		teSetSZ(pVarResult, lpBuf);
		LocalFree(lpBuf);
	}
}

VOID teApiFreeLibrary(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teFreeLibrary(param[0].hmodule, param[1].uintVal));
}

VOID teApiGetAsyncKeyState(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetAsyncKeyState(param[0].intVal));
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

VOID teApiGetCapture(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetCapture());
}

VOID teApiGetClassLongPtr(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetClassLongPtr(param[0].hwnd, param[1].intVal));
}

VOID teApiGetClassName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		IPersist *pPersist;
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pPersist))) {
			CLSID clsid;
			HRESULT hr = pPersist->GetClassID(&clsid);
			pPersist->Release();
			if SUCCEEDED(hr) {
				LPOLESTR lpsz;
				StringFromCLSID(clsid, &lpsz);
				teSetSZ(pVarResult, lpsz);
				teCoTaskMemFree(lpsz);
				return;
			}
		}
	}
	BSTR bsResult = ::SysAllocStringLen(NULL, MAX_CLASS_NAME);
	int nLen = GetClassName(param[0].hwnd, bsResult, MAX_CLASS_NAME);
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiGetClientRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetClientRect(param[0].hwnd, param[1].lprect));
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

VOID teApiGetCurrentProcess(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetCurrentProcess());
}

VOID teApiGetCurrentThreadId(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetCurrentThreadId());
}

VOID teApiGetCursorPos(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	teSetBool(pVarResult, GetCursorPos(&pt));
	PutPointToVariant(&pt, &pDispParams->rgvarg[nArg]);
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

VOID teApiGetDC(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetDC(param[0].hwnd));
}

VOID teApiGetDeviceCaps(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetDeviceCaps(param[0].hdc, param[1].intVal));
}

VOID teApiGetDiskFreeSpaceEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	ULARGE_INTEGER FreeBytesOfAvailable;
	ULARGE_INTEGER TotalNumberOfBytes;
	ULARGE_INTEGER TotalNumberOfFreeBytes;
	if (pVarResult && tePathIsDirectory(param[0].lpwstr, 0, 3) == S_OK) {
		if (GetDiskFreeSpaceEx(param[0].lpcwstr, &FreeBytesOfAvailable, &TotalNumberOfBytes, &TotalNumberOfFreeBytes)) {
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

VOID teApiGetDisplayNameOf(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teGetDisplayNameOf(&pDispParams->rgvarg[nArg], param[1].intVal, pVarResult);
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

VOID teApiGetFocus(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetFocus());
}

VOID teApiGetForegroundWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetForegroundWindow());
}

VOID teApiGetGUIThreadInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetGUIThreadInfo(param[0].dword, param[1].pgui));
}

VOID teApiGetIconInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetIconInfo(param[0].hicon, param[1].piconinfo));
}

VOID teApiGetKeyboardState(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetKeyboardState(param[0].pbVal));
}

VOID teApiGetKeyNameText(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = ::SysAllocStringLen(NULL, MAX_PATH);
	int nLen = GetKeyNameText(param[0].lVal, bsResult, MAX_PATH);
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiGetKeyState(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetKeyState(param[0].intVal));
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

VOID teApiGetMenuDefaultItem(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMenuDefaultItem(param[0].hmenu, param[1].uintVal, param[2].uintVal));
}

VOID teApiGetMenuInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMenuInfo(param[0].hmenu, param[1].lpmenuinfo));
}

VOID teApiGetMenuItemCount(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMenuItemCount(param[0].hmenu));
}

VOID teApiGetMenuItemID(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMenuItemID(param[0].hmenu, param[1].intVal));
}

VOID teApiGetMenuItemInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMenuItemInfo(param[0].hmenu, param[1].uintVal, param[2].boolVal, param[3].lpmenuiteminfo));
}

VOID teApiGetMenuItemRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMenuItemRect (param[0].hwnd, param[1].hmenu, param[2].uintVal, param[3].lprect));
}

VOID teApiGetMenuString(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = NULL;
	int nLen = teGetMenuString(&bsResult, param[0].hmenu, param[1].uintVal, param[2].lVal != MF_BYCOMMAND);
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiGetMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMessage(param[0].lpmsg, param[1].hwnd, param[2].uintVal, param[3].uintVal));
}

VOID teApiGetMessagePos(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetMessagePos());
}

VOID teApiGetModuleFileName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = NULL;
	int nLen = teGetModuleFileName(param[0].hmodule, &bsResult);
	teSetBSTR(pVarResult, &bsResult, nLen);
}

VOID teApiGetModuleHandle(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetModuleHandle(param[0].lpcwstr));
}

VOID teApiGetMonitorInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetMonitorInfo(param[0].hmonitor, param[1].lpmonitorinfo));
}

VOID teApiGetObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetObject(param[0].handle, param[1].intVal, param[2].lpvoid));
}

VOID teApiGetParent(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetParent(param[0].hwnd));
}

VOID teApiGetPixel(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ::GetPixel(param[0].hdc, param[1].intVal, param[2].intVal));
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

VOID teApiGetProp(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetProp(param[0].hwnd, param[1].lpcwstr));
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

VOID teApiGetShortPathName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].lpcwstr) {
		int nLen = GetShortPathName(param[0].lpcwstr, NULL, 0);
		BSTR bsResult = ::SysAllocStringLen(NULL, nLen + MAX_PATH);
		nLen = GetShortPathName(param[0].lpcwstr, bsResult, nLen);
		teSetBSTR(pVarResult, &bsResult, nLen);
	}
}

VOID teApiGetStockObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetStockObject(param[0].intVal));
}

VOID teApiGetSubMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetSubMenu(param[0].hmenu, param[1].intVal));
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

VOID teApiGetSysColorBrush(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetSysColorBrush(param[0].intVal));
}

VOID teApiGetSystemMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetSystemMenu(param[0].hwnd, param[1].boolVal));
}

VOID teApiGetSystemMetrics(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetSystemMetrics(param[0].intVal));
}

VOID teApiGetTextExtentPoint32(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetTextExtentPoint32(param[0].hdc, param[1].lpcwstr, lstrlen(param[1].lpcwstr), param[2].lpsize));
}

VOID teApiGetThreadCount(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, param[0].dword ? teGetthreadCount(param[0].dword): g_nThreads);
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

VOID teApiGetTopWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetTopWindow(param[0].hwnd));
}

VOID teApiGetUserName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	WCHAR pszName[MAX_PROP];
	DWORD dwSize = MAX_PROP;
	GetUserName(pszName, &dwSize);
	teSetSZ(pVarResult, pszName);
}

VOID teApiGetVersionEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].prtl_osversioninfoexw) {
		teSetBool(pVarResult, _RtlGetVersion ? !_RtlGetVersion(param[0].prtl_osversioninfoexw) : GetVersionEx(param[0].lposversioninfo));
	}
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

VOID teApiGetWindowDC(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetWindowDC(param[0].hwnd));
}

VOID teApiGetWindowLongPtr(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, GetWindowLongPtr(param[0].hwnd, param[1].intVal));
}

VOID teApiGetWindowRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, GetWindowRect(param[0].hwnd, param[1].lprect));
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

VOID teApiGetWindowThreadProcessId(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetWindowThreadProcessId(param[0].hwnd, param[1].lpdword));
}

VOID teApiGlobalAddAtom(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GlobalAddAtom(param[0].lpcwstr));
}

VOID teApiGlobalDeleteAtom(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GlobalDeleteAtom(param[0].atom));
}

VOID teApiGlobalFindAtom(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GlobalFindAtom(param[0].lpcwstr));
}

VOID teApiGlobalGetAtomName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	TCHAR szBuf[256];
	UINT uSize = GlobalGetAtomName(param[0].atom, szBuf, ARRAYSIZE(szBuf));
	teSetSZ(pVarResult, szBuf);
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

VOID teApiHighPart(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, param[0].li.HighPart);
}

VOID teApiILClone(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	FolderItem *pid;
	GetFolderItemFromVariant(&pid, &pDispParams->rgvarg[nArg]);
	teSetObjectRelease(pVarResult, pid);
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

VOID teApiILGetCount(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	LPITEMIDLIST pidl;
	if (teGetIDListFromVariant(&pidl, &pDispParams->rgvarg[nArg])) {
		teSetLong(pVarResult, ILGetCount(pidl));
		teCoTaskMemFree(pidl);
	}
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

VOID teApiILIsEmpty(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	FolderItem *pid1, *pid2;
	GetFolderItemFromIDList(&pid1, g_pidls[CSIDL_DESKTOP]);
	GetFolderItemFromVariant(&pid2, &pDispParams->rgvarg[nArg]);
	teSetBool(pVarResult, teILIsEqual(pid1, pid2));
	pid2->Release();
	pid1->Release();
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

VOID teApiImageList_Add(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_Add(param[0].himagelist, param[1].hbitmap, param[2].hbitmap));
}

VOID teApiImageList_AddIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_AddIcon(param[0].himagelist, param[1].hicon));
}

VOID teApiImageList_AddMasked(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_AddMasked(param[0].himagelist, param[1].hbitmap, param[2].colorref));
}

VOID teApiImageList_Copy(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_Copy(param[0].himagelist, param[1].intVal,
		param[2].himagelist, param[3].intVal, param[4].uintVal));
}

VOID teApiImageList_Create(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_Create(param[0].intVal, param[1].intVal, param[2].uintVal, param[3].intVal, param[4].intVal));
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

VOID teApiImageList_Duplicate(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_Duplicate(param[0].himagelist));
}

VOID teApiImageList_GetBkColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_GetBkColor(param[0].himagelist));
}

VOID teApiImageList_GetIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_GetIcon(param[0].himagelist, param[1].intVal, param[2].uintVal));
}

VOID teApiImageList_GetIconSize(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_GetIconSize(param[0].himagelist, (int *)&(param[1].lpsize->cx), (int *)&(param[1].lpsize->cy)));
}

VOID teApiImageList_GetImageCount(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_GetImageCount(param[0].himagelist));
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

VOID teApiImageList_LoadImage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImageList_LoadImage(param[0].hinstance, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
		param[2].intVal, param[3].intVal, param[4].colorref, param[5].uintVal, param[6].uintVal));
}

VOID teApiImageList_Remove(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_Remove(param[0].himagelist, param[1].intVal));
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

VOID teApiImageList_SetBkColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImageList_SetBkColor(param[0].himagelist, param[1].colorref));
}

VOID teApiImageList_SetIconSize(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_SetIconSize(param[0].himagelist, param[1].intVal, param[2].intVal));
}

VOID teApiImageList_SetImageCount(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_SetImageCount(param[0].himagelist, param[1].uintVal));
}

VOID teApiImageList_SetOverlayImage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImageList_SetOverlayImage(param[0].himagelist, param[1].intVal, param[2].intVal));
}

VOID teApiImmGetContext(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ImmGetContext(param[0].hwnd));
}

VOID teApiImmGetVirtualKey(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ImmGetVirtualKey(param[0].hwnd));
}

VOID teApiImmReleaseContext(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImmReleaseContext(param[0].hwnd, param[1].himc));
}

VOID teApiImmSetOpenStatus(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ImmSetOpenStatus(param[0].himc, param[1].boolVal));
}

VOID teApiInsertMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, InsertMenu(param[0].hmenu, param[1].uintVal, param[2].uintVal,
		param[3].uint_ptr, param[4].lpcwstr));
}

VOID teApiInsertMenuItem(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, InsertMenuItem(param[0].hmenu, param[1].uintVal, param[2].boolVal, param[3].lpcmenuiteminfo));
}

VOID teApiInvalidateRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, InvalidateRect(param[0].hwnd, param[1].lprect, param[2].boolVal));
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

VOID teApiIsAppThemed(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsAppThemed());
}

VOID teApiIsChild(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsChild(param[0].hwnd, param[1].hwnd));
}

VOID teApiIsClipboardFormatAvailable(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsClipboardFormatAvailable(param[0].uintVal));
}

VOID teApiIsIconic(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsIconic(param[0].hwnd));
}

VOID teApiIsMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsMenu(param[0].hmenu));
}

VOID teApiIsWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsWindow(param[0].hwnd));
}

VOID teApiIsWindowEnabled(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsWindowEnabled(param[0].hwnd));
}

VOID teApiIsWindowVisible(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsWindowVisible(param[0].hwnd));
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

VOID teApiIsZoomed(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, IsZoomed(param[0].hwnd));
}

VOID teApikeybd_event(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	keybd_event(param[0].bVal, param[1].bVal, param[2].dword, param[3].ulong_ptr);
}

VOID teApiKillTimer(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, KillTimer(param[0].hwnd, param[1].uint_ptr));
}

VOID teApiLineTo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, LineTo(param[0].hdc, param[1].intVal, param[2].intVal));
}

VOID teApiLoadCursor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadCursor(param[0].hinstance, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teApiLoadCursorFromFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadCursorFromFile(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teApiLoadIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadIcon(param[0].hinstance, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teApiLoadImage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadImage(param[0].hinstance, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
		param[2].uintVal,	param[3].intVal, param[4].intVal, param[5].uintVal));
}

VOID teApiLoadLibraryEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadLibraryEx(GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg]), param[1].handle, param[2].dword));
}

VOID teApiLoadMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, LoadMenu(param[0].hinstance, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1])));
}

VOID teApiLoadString(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = ::SysAllocStringLen(NULL, SIZE_BUFF);
	int nLen = LoadString(param[0].hinstance, param[1].uintVal, bsResult, SIZE_BUFF);
	teSetBSTR(pVarResult, &bsResult, -1);
}

VOID teApiLowPart(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, param[0].li.LowPart);
}

VOID teApiMapVirtualKey(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, MapVirtualKey(param[0].uintVal, param[1].uintVal));
}

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
			SafeRelease(&pStream);
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

VOID teApiMenuItemFromPoint(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, MenuItemFromPoint(param[0].hwnd, param[1].hmenu, *param[2].lppoint));
}

VOID teApiMessageBox(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (nArg >= 4) {
		GetDispatch(&pDispParams->rgvarg[nArg - 4], &g_pMBText);
	}
	teSetLong(pVarResult, MessageBox(param[0].hwnd, param[1].lpcwstr, param[2].lpcwstr, param[3].uintVal));
	SafeRelease(&g_pMBText);
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

VOID teApiMonitorFromWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, MonitorFromWindow(param[0].hwnd, param[1].dword));
}

VOID teApimouse_event(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	mouse_event(param[0].dword, param[1].dword, param[2].dword, param[3].dword, param[4].ulong_ptr);
}

VOID teApiMoveFileEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, MoveFileEx(param[0].lpcwstr, param[1].lpcwstr, param[2].dword));
}

VOID teApiMoveToEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, MoveToEx(param[0].hdc, param[1].intVal, param[2].intVal, param[3].lppoint));
}

VOID teApiMsgWaitForMultipleObjectsEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, MsgWaitForMultipleObjectsEx(param[0].dword, param[1].lphandle, param[2].dword, param[3].dword, param[4].dword));
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

VOID teApiObjDelete(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	IUnknown *punk;
	if (FindUnknown(&pDispParams->rgvarg[nArg], &punk)) {
		teSetLong(pVarResult, teDelProperty(punk, param[1].lpolestr));
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

VOID teApiObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetObjectRelease(pVarResult, teCreateObj(TE_OBJECT, NULL));
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

VOID teApiOpenIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, OpenIcon(param[0].hwnd));
}

VOID teApiOpenProcess(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, OpenProcess(param[0].dword, param[1].boolVal, param[2].dword));
}

VOID teApiOutputDebugString(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	::OutputDebugString(param[0].lpcwstr);
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

VOID teApiPathIsNetworkPath(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, tePathIsNetworkPath(param[0].lpcwstr));
}

VOID teApiPathIsRoot(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, PathIsRoot(param[0].lpcwstr));
}

VOID teApiPathIsSameRoot(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, PathIsSameRoot(param[0].lpcwstr, param[1].lpcwstr));
}

VOID teApiPathMatchSpec(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, tePathMatchSpec(param[0].lpcwstr, param[1].lpwstr));
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

VOID teApiPathSearchAndQualify(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (param[0].bstrVal) {
		UINT uLen = ::SysStringLen(param[0].bstrVal) + MAX_PATH;
		BSTR bsResult = teSysAllocStringLen(param[0].lpolestr, uLen);
		PathSearchAndQualify(param[0].lpcwstr, bsResult, uLen);
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

VOID teApiPeekMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, PeekMessage(param[0].lpmsg, param[1].hwnd, param[2].uintVal, param[3].uintVal, param[4].uintVal));
}

VOID teApiPlaySound(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, PlaySound(param[0].lpcwstr, param[1].hmodule, param[2].dword));
}

VOID teApiPlgBlt(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SetStretchBltMode(param[0].hdc, HALFTONE);
	teSetBool(pVarResult, PlgBlt(param[0].hdc, param[1].lppoint, param[2].hdc, param[3].intVal, param[4].intVal,
		param[5].intVal, param[6].intVal, param[7].hbitmap, param[8].intVal, param[9].intVal));
}

VOID teApiPostMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	PostMessage(param[0].hwnd, param[1].uintVal, param[2].wparam, param[3].lparam);
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

VOID teApipvData(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pDispParams->rgvarg[nArg].vt & VT_ARRAY) {
		teSetPtr(pVarResult, pDispParams->rgvarg[nArg].parray->pvData);
	}
}

VOID teApiQuadAdd(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLL(pVarResult, param[0].llVal + param[1].llVal);
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
			pStream->Release();
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

VOID teApiRectangle(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, Rectangle(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal));
}

VOID teApiRedrawWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, RedrawWindow(param[0].hwnd, param[1].lprect, param[2].hrgn, param[3].uintVal));
}

VOID teApiRegisterHotKey(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, RegisterHotKey(param[0].hwnd, param[1].intVal, param[2].uintVal, param[3].uintVal));
}

VOID teApiRegisterWindowMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, RegisterWindowMessage(param[0].lpcwstr));
}

VOID teApiReleaseCapture(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ReleaseCapture());
}

VOID teApiReleaseDC(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, ReleaseDC(param[0].hwnd, param[1].hdc));
}

VOID teApiRemoveMenu(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, RemoveMenu(param[0].hmenu, param[1].uintVal, param[2].uintVal));
}

VOID teApiResetEvent(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ResetEvent(param[0].handle));
}

VOID teApiRoundRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, RoundRect(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal, param[5].intVal, param[6].intVal));
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

VOID teApiScreenToClient(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg - 1]);
	teSetBool(pVarResult, ScreenToClient(param[0].hwnd, &pt));
	PutPointToVariant(&pt, &pDispParams->rgvarg[nArg - 1]);
}

VOID teApiSelectObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SelectObject(param[0].hdc, param[1].hgdiobj));
}

VOID teApiSendInput(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SendInput(param[0].uintVal, param[1].lpinput, param[2].intVal));
}

VOID teApiSendMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SendMessage(param[0].hwnd, param[1].uintVal, param[2].wparam, param[3].lparam));
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

VOID teApiSendNotifyMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SendNotifyMessage(param[0].hwnd, param[1].uintVal, param[2].wparam, param[3].lparam));
}

VOID teApiSetBkColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetBkColor(param[0].hdc, param[1].colorref));
}

VOID teApiSetBkMode(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetBkMode(param[0].hdc, param[1].intVal));
}

VOID teApiSetCapture(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetCapture(param[0].hwnd));
}

VOID teApiSetClassLongPtr(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetClassLongPtr(param[0].hwnd, param[1].intVal, param[2].long_ptr));
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

VOID teApiSetCurrentDirectory(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetCurrentDirectory(param[0].lpcwstr));//use wsh.CurrentDirectory
}

VOID teApiSetCursor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetCursor(param[0].hcursor));
}

VOID teApiSetCursorPos(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetCursorPos(param[0].intVal, param[1].intVal));
}

VOID teApiSetDCBrushColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetDCBrushColor(param[0].hdc, param[1].colorref));
}

VOID teApiSetDCPenColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetDCPenColor(param[0].hdc, param[1].colorref));
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

VOID teApiSetEvent(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetEvent(param[0].handle));
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
		pStream->Release();
	} else {
		LARGE_INTEGER li;
		li.HighPart = param[1].li.HighPart;
		li.LowPart = ::SetFilePointer(param[0].handle, param[1].li.LowPart, &li.HighPart, param[2].dword);
		teSetLL(pVarResult, li.QuadPart);
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

VOID teApiSetForegroundWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teSetForegroundWindow(param[0].hwnd));
}

VOID teApiSetKeyboardState(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetKeyboardState(param[0].pbVal));
}

VOID teApiSetLayeredWindowAttributes(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetLayeredWindowAttributes(param[0].hwnd, param[1].colorref, param[2].bVal, param[3].dword));
}

VOID teApiSetMenuDefaultItem(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuDefaultItem(param[0].hmenu, param[1].uintVal, param[2].uintVal));
}

VOID teApiSetMenuInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuInfo(param[0].hmenu, param[1].lpcmenuinfo));
}

VOID teApiSetMenuItemBitmaps(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuItemBitmaps(param[0].hmenu, param[1].uintVal, param[2].uintVal, param[3].hbitmap, param[4].hbitmap));
}

VOID teApiSetMenuItemInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetMenuItemInfo(param[0].hmenu, param[1].uintVal, param[2].boolVal, param[3].lpmenuiteminfo));
}

VOID teApiSetRect(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SetRect(param[0].lprect, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal);
	teVariantCopy(pVarResult, &pDispParams->rgvarg[nArg]);
}

VOID teApiSetSysColors(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetSysColors(param[0].intVal, (INT *)GetpcFromVariant(&pDispParams->rgvarg[nArg - 1], NULL), (COLORREF *)GetpcFromVariant(&pDispParams->rgvarg[nArg - 2], NULL)));
}

VOID teApiSetTextColor(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SetTextColor(param[0].hdc, param[1].colorref));
}

VOID teApiSetTimer(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetTimer(param[0].hwnd, param[1].uint_ptr, param[2].uintVal, NULL));
}

VOID teApiSetWindowLongPtr(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, SetWindowLongPtr(param[0].hwnd, param[1].intVal, param[2].long_ptr));
}

VOID teApiSetWindowPos(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SetWindowPos(param[0].hwnd, param[1].hwnd, param[2].intVal, param[3].intVal, param[4].intVal, param[5].intVal, param[6].uintVal));
}

VOID teApiSetWindowText(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, teSetWindowText(param[0].hwnd, param[1].lpcwstr));
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
						FolderItem *pPF;
						GetFolderItemFromIDList(&pPF, ppidl[i]);
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

VOID teApiSHChangeNotification_Unlock(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, TRUE);
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

VOID teApiSHChangeNotifyDeregister(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SHChangeNotifyDeregister(param[0].ulVal));
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
		SafeRelease(&pStream0);
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
	SafeRelease(&pStream0);
	SafeRelease(&pStream);
}

VOID teApiSHDefExtractIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SHDefExtractIcon(param[0].lpcwstr, param[1].intVal, param[2].uintVal, param[3].phicon, param[4].phicon, param[5].uintVal));
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

VOID teApiShell_NotifyIcon(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, Shell_NotifyIcon(param[0].dword, param[1].pnotifyicondata));
}

VOID teApiShellExecute(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetPtr(pVarResult, ShellExecute(param[0].hwnd, GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 1]),
		GetLPWSTRFromVariant(&pDispParams->rgvarg[nArg - 2]), param[3].lpcwstr, param[4].lpcwstr, param[5].intVal));
}

VOID teApiShellExecuteEx(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ::ShellExecuteEx(param[0].lpshellexecuteinfo));
}

VOID teApiSHEmptyRecycleBin(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, SHEmptyRecycleBin(param[0].hwnd, param[1].lpcwstr, param[2].dword));
}

VOID teApiSHFreeNameMappings(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SHFreeNameMappings(param[0].handle);
}

VOID teApiSHFreeShared(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SHFreeShared(param[0].handle, param[1].dword));
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

VOID teApiSHGetImageList(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HANDLE h = 0;
	if FAILED(SHGetImageList(param[0].intVal, IID_IImageList, (LPVOID *)&h)) {
		h = 0;
	}
	teSetPtr(pVarResult, h);
}

VOID teApiShouldAppsUseDarkMode(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, g_bDarkMode);
}

VOID teApiShowWindow(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ShowWindow(param[0].hwnd, param[1].intVal));
}

VOID teApiSHParseDisplayName(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teAsyncInvoke(0, nArg, pDispParams, pVarResult);
}

VOID teApiShRunDialog(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (_SHRunDialog) {
		_SHRunDialog(param[0].hwnd, param[1].hicon, param[2].lpwstr, param[3].lpwstr, param[4].lpwstr, param[5].dword);
	}
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

VOID teApiSHTestTokenMembership(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SHTestTokenMembership(param[0].handle, param[1].ulVal));
}

VOID teApisizeof(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, GetSizeOf(&pDispParams->rgvarg[nArg]));
}

VOID teApiSleep(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	Sleep(param[0].dword);
#ifdef _DEBUG
	teSetPtr(pVarResult, GetCurrentThreadId());
#endif
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

VOID teApiStrCmpI(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, teStrCmpI(param[0].lpcwstr, param[1].lpcwstr));
}

VOID teApiStrCmpLogical(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, StrCmpLogicalW(param[0].lpcwstr, param[1].lpcwstr));
}

VOID teApiStrCmpNI(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, StrCmpNI(param[0].lpcwstr, param[1].lpcwstr, param[2].intVal));
}

VOID teApiStretchBlt(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	SetStretchBltMode(param[0].hdc, HALFTONE);
	teSetBool(pVarResult, StretchBlt(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal,
		param[5].hdc, param[6].intVal, param[7].intVal, param[8].intVal, param[9].intVal, param[10].dword));
}

VOID teApiStrFormatByteSize(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	BSTR bsResult = ::SysAllocStringLen(NULL, MAX_COLUMN_NAME_LEN);
	teStrFormatSize(nArg >= 1 ? param[1].dword : 2, param[0].llVal, bsResult, MAX_COLUMN_NAME_LEN);
	teSetBSTR(pVarResult, &bsResult, -1);
}

VOID teApiStrLen(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, lstrlen(param[0].lpcwstr));
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

VOID teApiSysAllocStringLen(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	if (pVarResult) {
		VARIANT *pv = &pDispParams->rgvarg[nArg];
		pVarResult->bstrVal = (pv->vt == VT_BSTR) ? teSysAllocStringLenEx(pv->bstrVal, param[1].uintVal) : teSysAllocStringLen(GetLPWSTRFromVariant(pv), param[1].uintVal);
		pVarResult->vt = VT_BSTR;
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

VOID teApiSystemParametersInfo(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, SystemParametersInfo(param[0].uintVal, param[1].uintVal, param[2].lpvoid, param[3].uintVal));
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
	g_hMenuKeyHook = ::SetWindowsHookEx(WH_KEYBOARD, (HOOKPROC)MenuKeyProc, NULL, g_dwMainThreadId);
	RAWINPUTDEVICE rid[1];
	rid[0].usUsagePage = 0x01;
	rid[0].usUsage = 0x02;
	rid[0].dwFlags = 0;
	rid[0].hwndTarget = g_hwndMain;
	::RegisterRawInputDevices(rid, _countof(rid), sizeof(RAWINPUTDEVICE));
	try {
		teSetLong(pVarResult, TrackPopupMenuEx(param[0].hmenu, param[1].uintVal, param[2].intVal, param[3].intVal,
			param[4].hwnd, param[5].lptpmparams));
	} catch (...) {
		g_nException = 0;
#ifdef _DEBUG
		g_strException = L"TrackPopupMenuEx";
#endif
	}
	rid[0].dwFlags = RIDEV_REMOVE;
	rid[0].hwndTarget = NULL;
	::RegisterRawInputDevices(rid, _countof(rid), sizeof(RAWINPUTDEVICE));
	::UnhookWindowsHookEx(g_hMenuKeyHook);
	g_hMenuKeyHook = NULL;
	SafeRelease(&g_pCM);
}

VOID teApiTranslateMessage(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, ::TranslateMessage(param[0].lpmsg));
}

VOID teApiTransparentBlt(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetBool(pVarResult, TransparentBlt(param[0].hdc, param[1].intVal, param[2].intVal, param[3].intVal, param[4].intVal,
		param[5].hdc, param[6].intVal, param[7].intVal, param[8].intVal, param[9].intVal, param[10].uintVal));
}

VOID teApiULowPart(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetULL(pVarResult, param[0].uli.LowPart);
}

VOID teApiUnregisterHotKey(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, UnregisterHotKey(param[0].hwnd, param[1].intVal));
}

VOID teApiUQuadAdd(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetULL(pVarResult, teGetU(param[0].llVal) + teGetU(param[1].llVal));
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

VOID teApiUQuadSub(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLL(pVarResult, teGetU(param[0].llVal) - teGetU(param[1].llVal));
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

VOID teApiURLDownloadToFile(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	HRESULT hr = E_FAIL;
	IUnknown *punk = NULL;
	if (FindUnknown(&pDispParams->rgvarg[nArg - 1], &punk)) {
		IStream *pDst, *pSrc;
		hr = punk->QueryInterface(IID_PPV_ARGS(&pSrc));
		if SUCCEEDED(hr) {
			hr = SHCreateStreamOnFileEx(param[2].bstrVal, STGM_WRITE | STGM_CREATE | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_ARCHIVE, TRUE, NULL, &pDst);
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
						HANDLE hFile = CreateFile(param[2].bstrVal, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, NULL);
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
				HANDLE hFile = CreateFile(param[2].bstrVal, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, NULL);
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

VOID teApiWaitForInputIdle(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, WaitForInputIdle(param[0].handle, param[1].dword));
}

VOID teApiWaitForSingleObject(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	teSetLong(pVarResult, WaitForSingleObject(param[0].handle, param[1].dword));
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

VOID teApiWindowFromPoint(int nArg, teParam *param, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
	POINT pt;
	GetPointFormVariant(&pt, &pDispParams->rgvarg[nArg]);
	teSetPtr(pVarResult, WindowFromPoint(pt));
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
			pStream->Release();
		} else {
			teSetBool(pVarResult, ::WriteFile(param[0].handle, pc, nLen, &dwWriteByte, NULL));
		}
	}
	VariantClear(&vMem);
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
}
#endif

TEDispatchApi dispAPI[] = {
	{ 2, -1, -1, -1, "Add", teApiQuadAdd },
	{ 1, -1, -1, -1, "AllowSetForegroundWindow", teApiAllowSetForegroundWindow },
	{11, -1, -1, -1, "AlphaBlend", teApiAlphaBlend },
	{ 0, -1, -1, -1, "Array", teApiArray },
	{ 4,  2,  3, -1, "AssocQueryString", teApiAssocQueryString },
	{ 1,  0, -1, -1, "base64_decode", teApibase64_decode },
	{ 1, -1, -1, -1, "base64_encode", teApibase64_encode },
	{ 2, -1, -1, -1, "BeginPaint", teApiBeginPaint },
	{ 9, -1, -1, -1, "BitBlt", teApiBitBlt },
	{ 1, -1, -1, -1, "BringWindowToTop", teApiBringWindowToTop },
	{ 5, -1, -1, -1, "CallWindowProc", teApiCallWindowProc },
	{ 4, -1, -1, -1, "ChangeWindowMessageFilterEx", teApiChangeWindowMessageFilterEx },
	{ 1, -1, -1, -1, "ChooseColor", teApiChooseColor },
	{ 1, -1, -1, -1, "ChooseFont", teApiChooseFont },
	{ 2, -1, -1, -1, "ClientToScreen", teApiClientToScreen },
	{ 1, -1, -1, -1, "CloseHandle", teApiCloseHandle },
	{ 1,  0, -1, -1, "CommandLineToArgv", teApiCommandLineToArgv },
	{ 3, -1, -1, -1, "CompareIDs", teApiCompareIDs },
	{ 1,  5, -1, -1, "ContextMenu", teApiContextMenu },
	{ 5, -1, -1, -1, "CopyImage", teApiCopyImage },
	{ 1, -1, -1, -1, "CoTaskMemFree", teApiCoTaskMemFree },
	{ 1, -1, -1, -1, "CRC32", teApiCRC32 },
	{ 3, -1, -1, -1, "CreateCompatibleBitmap", teApiCreateCompatibleBitmap },
	{ 1, -1, -1, -1, "CreateCompatibleDC", teApiCreateCompatibleDC },
	{ 4,  3, -1, -1, "CreateEvent", teApiCreateEvent },
	{ 7,  0, -1, -1, "CreateFile", teApiCreateFile },
	{ 1, -1, -1, -1, "CreateFontIndirect", teApiCreateFontIndirect },
	{ 1, -1, -1, -1, "CreateIconIndirect", teApiCreateIconIndirect },
	{ 0, -1, -1, -1, "CreateMenu", teApiCreateMenu },
	{ 1,  0, -1, -1, "CreateObject", teApiCreateObject },
	{ 3, -1, -1, -1, "CreatePen", teApiCreatePen },
	{ 0, -1, -1, -1, "CreatePopupMenu", teApiCreatePopupMenu },
	{ 1,  0,  1, -1, "CreateProcess", teApiCreateProcess },
	{ 1, -1, -1, -1, "CreateSolidBrush", teApiCreateSolidBrush },
	{12, -1,  1,  2, "CreateWindowEx", teApiCreateWindowEx },
	{ 2,  0,  1, -1, "CryptProtectData", teApiCryptProtectData },
	{ 2, -1,  1, -1, "CryptUnprotectData", teApiCryptUnprotectData },
	{ 0, -1, -1, -1, "DataObject", teApiDataObject },
	{ 1, -1, -1, -1, "DeleteDC", teApiDeleteDC },
	{ 1,  0, -1, -1, "DeleteFile", teApiDeleteFile },
	{ 3, -1, -1, -1, "DeleteMenu", teApiDeleteMenu },
	{ 1, -1, -1, -1, "DeleteObject", teApiDeleteObject },
	{ 2, -1, -1, -1, "DestroyCursor", teApiDestroyCursor },
	{ 1, -1, -1, -1, "DestroyIcon", teApiDestroyIcon },
	{ 1, -1, -1, -1, "DestroyMenu", teApiDestroyMenu },
	{ 1, -1, -1, -1, "DestroyWindow", teApiDestroyWindow },
	{ 1, -1, -1, -1, "DispatchMessage", teApiDispatchMessage },
	{ 2, -1, -1, -1, "DllGetClassObject", teApiDllGetClassObject },
	{ 2, -1, -1, -1, "DoDragDrop", teApiDoDragDrop },//Deprecated
	{ 0, -1, -1, -1, "DoEvents", teApiDoEvents },
	{ 2, -1, -1, -1, "DragAcceptFiles", teApiDragAcceptFiles },
	{ 1, -1, -1, -1, "DragFinish", teApiDragFinish },
	{ 2, -1, -1, -1, "DragQueryFile", teApiDragQueryFile },
	{ 4, -1, -1, -1, "DrawIcon", teApiDrawIcon },
	{ 9, -1, -1, -1, "DrawIconEx", teApiDrawIconEx },
	{ 5,  1, -1, -1, "DrawText", teApiDrawText },
	{ 1, -1, -1, -1, "DropTarget", teApiDropTarget },
	{ 3, -1, -1, -1, "EnableMenuItem", teApiEnableMenuItem },
	{ 2, -1, -1, -1, "EnableWindow", teApiEnableWindow },
	{ 0, -1, -1, -1, "EndMenu", teApiEndMenu },
	{ 2, -1, -1, -1, "EndPaint", teApiEndPaint },
	{ 3,  1, -1, -1, "ExecMethod", teApiExecMethod },
	{ 2,  0,  1, -1, "ExecScript", teApiExecScript },
	{ 4,  2,  3, -1, "Extract", teApiExtract },
	{ 5,  0, -1, -1, "ExtractIconEx", teApiExtractIconEx },
	{ 3, -1, -1, -1, "FillRect", teApiFillRect },
	{ 1, -1, -1, -1, "FindClose", teApiFindClose },
	{ 2, 10, -1, -1, "FindFirstFile", teApiFindFirstFile },
	{ 2, -1, -1, -1, "FindNextFile", teApiFindNextFile },
	{ 1, -1, -1, -1, "FindReplace", teApiFindReplace },
	{ 1, -1, -1, -1, "FindText", teApiFindText },
	{ 2,  0,  1, -1, "FindWindow", teApiFindWindow },
	{ 4,  2,  3, -1, "FindWindowEx", teApiFindWindowEx },
	{ 1,  1,  2,  3, "FormatMessage", teApiFormatMessage },
	{ 1, -1, -1, -1, "FreeLibrary", teApiFreeLibrary },
	{ 1, -1, -1, -1, "GetAsyncKeyState", teApiGetAsyncKeyState },
	{ 2, -1, -1, -1, "GetAttributesOf", teApiGetAttributesOf },
	{ 0, -1, -1, -1, "GetCapture", teApiGetCapture },
	{ 2, -1, -1, -1, "GetClassLongPtr", teApiGetClassLongPtr },
	{ 1, -1, -1, -1, "GetClassName", teApiGetClassName },
	{ 2, -1, -1, -1, "GetClientRect", teApiGetClientRect },
	{ 0, -1, -1, -1, "GetClipboardData", teApiGetClipboardData },
	{ 0, -1, -1, -1, "GetCommandLine", teApiGetCommandLine },
	{ 0, -1, -1, -1, "GetCurrentDirectory", teApiGetCurrentDirectory },//use wsh.CurrentDirectory
	{ 0, -1, -1, -1, "GetCurrentProcess", teApiGetCurrentProcess },
	{ 0, -1, -1, -1, "GetCurrentThreadId", teApiGetCurrentThreadId },
	{ 1, -1, -1, -1, "GetCursorPos", teApiGetCursorPos },
	{ 4,  3, -1, -1, "GetDateFormat", teApiGetDateFormat },
	{ 1, -1, -1, -1, "GetDC", teApiGetDC },
	{ 2, -1, -1, -1, "GetDeviceCaps", teApiGetDeviceCaps },
	{ 1,  0, -1, -1, "GetDiskFreeSpaceEx", teApiGetDiskFreeSpaceEx },
	{ 2,  1, -1, -1, "GetDispatch", teApiGetDispatch },
	{ 2, -1, -1, -1, "GetDisplayNameOf", teApiGetDisplayNameOf },
	{ 0, -1, -1, -1, "GetDpiForMonitor", teApiGetDpiForMonitor },
	{ 0, -1, -1, -1, "GetFocus", teApiGetFocus },
	{ 0, -1, -1, -1, "GetForegroundWindow", teApiGetForegroundWindow },
	{ 2, -1, -1, -1, "GetGUIThreadInfo", teApiGetGUIThreadInfo },
	{ 2, -1, -1, -1, "GetIconInfo", teApiGetIconInfo },
	{ 1, -1, -1, -1, "GetKeyboardState", teApiGetKeyboardState },
	{ 1, -1, -1, -1, "GetKeyNameText", teApiGetKeyNameText },
	{ 1, -1, -1, -1, "GetKeyState", teApiGetKeyState },
	{ 2, -1, -1, -1, "GetLocaleInfo", teApiGetLocaleInfo },
	{ 3, -1, -1, -1, "GetMenuDefaultItem", teApiGetMenuDefaultItem },
	{ 2, -1, -1, -1, "GetMenuInfo", teApiGetMenuInfo },
	{ 1, -1, -1, -1, "GetMenuItemCount", teApiGetMenuItemCount },
	{ 2, -1, -1, -1, "GetMenuItemID", teApiGetMenuItemID },
	{ 4, -1, -1, -1, "GetMenuItemInfo", teApiGetMenuItemInfo },
	{ 4, -1, -1, -1, "GetMenuItemRect", teApiGetMenuItemRect },
	{ 3, -1, -1, -1, "GetMenuString", teApiGetMenuString },
	{ 4, -1, -1, -1, "GetMessage", teApiGetMessage },
	{ 0, -1, -1, -1, "GetMessagePos", teApiGetMessagePos },
	{ 1, -1, -1, -1, "GetModuleFileName", teApiGetModuleFileName },
	{ 1, -1,  0, -1, "GetModuleHandle", teApiGetModuleHandle },
	{ 2, -1, -1, -1, "GetMonitorInfo", teApiGetMonitorInfo },
	{ 3, -1, -1, -1, "GetObject", teApiGetObject },
	{ 1, -1, -1, -1, "GetParent", teApiGetParent },
	{ 3, -1, -1, -1, "GetPixel", teApiGetPixel },
	{ 2, -1, -1, -1, "GetProcAddress", teApiGetProcAddress },
	{ 2,  0, -1, -1, "GetProcObject", teApiGetProcObject },
	{ 2,  1, -1, -1, "GetProp", teApiGetProp },
	{ 2,  0,  1, -1, "GetScriptDispatch", teApiGetScriptDispatch },
	{ 1,  0, -1, -1, "GetShortPathName", teApiGetShortPathName },
	{ 1, -1, -1, -1, "GetStockObject", teApiGetStockObject },
	{ 2, -1, -1, -1, "GetSubMenu", teApiGetSubMenu },
	{ 1, -1, -1, -1, "GetSysColor", teApiGetSysColor },
	{ 1, -1, -1, -1, "GetSysColorBrush", teApiGetSysColorBrush },
	{ 2, -1, -1, -1, "GetSystemMenu", teApiGetSystemMenu },
	{ 1, -1, -1, -1, "GetSystemMetrics", teApiGetSystemMetrics },
	{ 3,  1, -1, -1, "GetTextExtentPoint32", teApiGetTextExtentPoint32 },
	{ 0, -1, -1, -1, "GetThreadCount", teApiGetThreadCount },
	{ 4,  3, -1, -1, "GetTimeFormat", teApiGetTimeFormat },
	{ 1, -1, -1, -1, "GetTopWindow", teApiGetTopWindow },
	{ 0, -1, -1, -1, "GetUserName", teApiGetUserName },
	{ 1, -1, -1, -1, "GetVersionEx", teApiGetVersionEx },
	{ 1, -1, -1, -1, "GetWindow", teApiGetWindow },
	{ 1, -1, -1, -1, "GetWindowDC", teApiGetWindowDC },
	{ 2, -1, -1, -1, "GetWindowLongPtr", teApiGetWindowLongPtr },
	{ 2, -1, -1, -1, "GetWindowRect", teApiGetWindowRect },
	{ 1, -1, -1, -1, "GetWindowText", teApiGetWindowText },
	{ 2, -1, -1, -1, "GetWindowThreadProcessId", teApiGetWindowThreadProcessId },
	{ 1,  0, -1, -1, "GlobalAddAtom", teApiGlobalAddAtom },
	{ 1, -1, -1, -1, "GlobalDeleteAtom", teApiGlobalDeleteAtom },
	{ 1,  0, -1, -1, "GlobalFindAtom", teApiGlobalFindAtom },
	{ 1, -1, -1, -1, "GlobalGetAtomName", teApiGlobalGetAtomName },
	{ 2, -1, -1, -1, "HashData", teApiHashData },
	{ 1, -1, -1, -1, "HasThumbnail", teApiHasThumbnail },
	{ 1, -1, -1, -1, "HighPart", teApiHighPart },
	{ 1, -1, -1, -1, "ILClone", teApiILClone },
	{ 1, -1, -1, -1, "ILCreateFromPath", teApiILCreateFromPath },
	{ 1, -1, -1, -1, "ILFindLastID", teApiILFindLastID },
	{ 1, -1, -1, -1, "ILGetCount", teApiILGetCount },
	{ 1, -1, -1, -1, "ILGetParent", teApiILGetParent },
	{ 1, -1, -1, -1, "ILIsEmpty", teApiILIsEmpty },
	{ 2, -1, -1, -1, "ILIsEqual", teApiILIsEqual },
	{ 3, -1, -1, -1, "ILIsParent", teApiILIsParent },
	{ 1, -1, -1, -1, "ILRemoveLastID", teApiILRemoveLastID },
	{ 3, -1, -1, -1, "ImageList_Add", teApiImageList_Add },
	{ 2, -1, -1, -1, "ImageList_AddIcon", teApiImageList_AddIcon },
	{ 3, -1, -1, -1, "ImageList_AddMasked", teApiImageList_AddMasked },
	{ 5, -1, -1, -1, "ImageList_Copy", teApiImageList_Copy },
	{ 5, -1, -1, -1, "ImageList_Create", teApiImageList_Create },
	{ 1, -1, -1, -1, "ImageList_Destroy", teApiImageList_Destroy },
	{ 6, -1, -1, -1, "ImageList_Draw", teApiImageList_Draw },
	{10, -1, -1, -1, "ImageList_DrawEx", teApiImageList_DrawEx },
	{ 1, -1, -1, -1, "ImageList_Duplicate", teApiImageList_Duplicate },
	{ 1, -1, -1, -1, "ImageList_GetBkColor", teApiImageList_GetBkColor },
	{ 3, -1, -1, -1, "ImageList_GetIcon", teApiImageList_GetIcon },
	{ 2, -1, -1, -1, "ImageList_GetIconSize", teApiImageList_GetIconSize },
	{ 1, -1, -1, -1, "ImageList_GetImageCount", teApiImageList_GetImageCount },
	{ 2, -1, -1, -1, "ImageList_GetOverlayImage", teApiImageList_GetOverlayImage },
	{ 7, -1, -1, -1, "ImageList_LoadImage", teApiImageList_LoadImage },
	{ 2, -1, -1, -1, "ImageList_Remove", teApiImageList_Remove },
	{ 4, -1, -1, -1, "ImageList_Replace", teApiImageList_Replace },
	{ 3, -1, -1, -1, "ImageList_ReplaceIcon", teApiImageList_ReplaceIcon },
	{ 2, -1, -1, -1, "ImageList_SetBkColor", teApiImageList_SetBkColor },
	{ 3, -1, -1, -1, "ImageList_SetIconSize", teApiImageList_SetIconSize },
	{ 2, -1, -1, -1, "ImageList_SetImageCount", teApiImageList_SetImageCount },
	{ 3, -1, -1, -1, "ImageList_SetOverlayImage", teApiImageList_SetOverlayImage },
	{ 1, -1, -1, -1, "ImmGetContext", teApiImmGetContext },
	{ 1, -1, -1, -1, "ImmGetVirtualKey", teApiImmGetVirtualKey },
	{ 2, -1, -1, -1, "ImmReleaseContext", teApiImmReleaseContext },
	{ 2, -1, -1, -1, "ImmSetOpenStatus", teApiImmSetOpenStatus },
	{ 5,  4, -1, -1, "InsertMenu", teApiInsertMenu },
	{ 4, -1, -1, -1, "InsertMenuItem", teApiInsertMenuItem },
	{ 3, -1, -1, -1, "InvalidateRect", teApiInvalidateRect },
	{ 1, -1, -1, -1, "Invoke", teApiInvoke },
	{ 0, -1, -1, -1, "IsAppThemed", teApiIsAppThemed },
	{ 2, -1, -1, -1, "IsChild", teApiIsChild },
	{ 1, -1, -1, -1, "IsClipboardFormatAvailable", teApiIsClipboardFormatAvailable },
	{ 1, -1, -1, -1, "IsIconic", teApiIsIconic },
	{ 1, -1, -1, -1, "IsMenu", teApiIsMenu },
	{ 1, -1, -1, -1, "IsWindow", teApiIsWindow },
	{ 1, -1, -1, -1, "IsWindowEnabled", teApiIsWindowEnabled },
	{ 1, -1, -1, -1, "IsWindowVisible", teApiIsWindowVisible },
	{ 1, -1, -1, -1, "IsWow64Process", teApiIsWow64Process },
	{ 1, -1, -1, -1, "IsZoomed", teApiIsZoomed },
	{ 4, -1, -1, -1, "keybd_event", teApikeybd_event },
	{ 2, -1, -1, -1, "KillTimer", teApiKillTimer },
	{ 3, -1, -1, -1, "LineTo", teApiLineTo },
	{ 2, -1, -1, -1, "LoadCursor", teApiLoadCursor },
	{ 2, -1, -1, -1, "LoadCursorFromFile", teApiLoadCursorFromFile },
	{ 2, -1, -1, -1, "LoadIcon", teApiLoadIcon },
	{ 6, -1, -1, -1, "LoadImage", teApiLoadImage },
	{ 3, -1, -1, -1, "LoadLibraryEx", teApiLoadLibraryEx },
	{ 2, -1, -1, -1, "LoadMenu", teApiLoadMenu },
	{ 2, -1, -1, -1, "LoadString", teApiLoadString },
	{ 1, -1, -1, -1, "LowPart", teApiLowPart },
	{ 2, -1, -1, -1, "MapVirtualKey", teApiMapVirtualKey },
	{ 1, -1, -1, -1, "Memory", teApiMemory },
	{ 3, -1, -1, -1, "MenuItemFromPoint", teApiMenuItemFromPoint },
	{ 4,  1,  2, -1, "MessageBox", teApiMessageBox },
	{ 2, -1, -1, -1, "MonitorFromPoint", teApiMonitorFromPoint },
	{ 2, -1, -1, -1, "MonitorFromRect", teApiMonitorFromRect },
	{ 2, -1, -1, -1, "MonitorFromWindow", teApiMonitorFromWindow },
	{ 2, -1, -1, -1, "mouse_event", teApimouse_event },
	{ 3,  0,  1, -1, "MoveFileEx", teApiMoveFileEx },
	{ 4, -1, -1, -1, "MoveToEx", teApiMoveToEx },
	{ 6, -1, -1, -1, "MoveWindow", teApiMoveWindow },
	{ 5, -1, -1, -1, "MsgWaitForMultipleObjectsEx", teApiMsgWaitForMultipleObjectsEx },
	{ 2, -1, -1, -1, "MultiByteToWideChar", teApiMultiByteToWideChar },
	{ 2,  1, -1, -1, "ObjDelete", teApiObjDelete },
	{ 2,  1, -1, -1, "ObjDispId", teApiObjDispId },
	{ 0, -1, -1, -1, "Object", teApiObject },
	{ 2,  1, -1, -1, "ObjGetI", teApiObjGetI },
	{ 3,  1, -1, -1, "ObjPutI", teApiObjPutI },
	{ 5,  1, -1, -1, "OleCmdExec", teApiOleCmdExec },
	{ 0, -1, -1, -1, "OleGetClipboard", teApiOleGetClipboard },
	{ 1, -1, -1, -1, "OleSetClipboard", teApiOleSetClipboard },
	{ 1, -1, -1, -1, "OpenIcon", teApiOpenIcon },
	{ 3, -1, -1, -1, "OpenProcess", teApiOpenProcess },
	{ 1,  0, -1, -1, "OutputDebugString", teApiOutputDebugString },
	{ 1,  0, -1, -1, "PathCreateFromUrl", teApiPathCreateFromUrl },
	{ 1, 10, -1, -1, "PathIsDirectory", teApiPathIsDirectory },
	{ 1,  0, -1, -1, "PathIsNetworkPath", teApiPathIsNetworkPath },
	{ 1,  0, -1, -1, "PathIsRoot", teApiPathIsRoot },
	{ 2,  0,  1, -1, "PathIsSameRoot", teApiPathIsSameRoot },
	{ 2,  0,  1, -1, "PathMatchSpec", teApiPathMatchSpec },
	{ 1,  0, -1, -1, "PathQuoteSpaces", teApiPathQuoteSpaces },
	{ 1,  0, -1, -1, "PathSearchAndQualify", teApiPathSearchAndQualify },
	{ 1,  0, -1, -1, "PathUnquoteSpaces", teApiPathUnquoteSpaces },
	{ 5, -1, -1, -1, "PeekMessage", teApiPeekMessage },
	{ 3,  0, -1, -1, "PlaySound", teApiPlaySound },
	{10, -1, -1, -1, "PlgBlt", teApiPlgBlt },
	{ 1, -1, -1, -1, "PostMessage", teApiPostMessage },
	{ 3,  0, -1, -1, "PSFormatForDisplay", teApiPSFormatForDisplay },
	{ 1,  0, -1, -1, "PSGetDisplayName", teApiPSGetDisplayName },
	{ 1, -1, -1, -1, "pvData", teApipvData },
	{ 1, -1, -1, -1, "QuadPart", teApiQuadPart },
	{ 2, -1, -1, -1, "ReadFile", teApiReadFile },
	{ 5, -1, -1, -1, "Rectangle", teApiRectangle },
	{ 4, -1, -1, -1, "RedrawWindow", teApiRedrawWindow },
	{ 4, -1, -1, -1, "RegisterHotKey", teApiRegisterHotKey },
	{ 1,  0, -1, -1, "RegisterWindowMessage", teApiRegisterWindowMessage },
	{ 0, -1, -1, -1, "ReleaseCapture", teApiReleaseCapture },
	{ 2, -1, -1, -1, "ReleaseDC", teApiReleaseDC },
	{ 3, -1, -1, -1, "RemoveMenu", teApiRemoveMenu },
	{ 1, -1, -1, -1, "ResetEvent", teApiResetEvent },
	{ 7, -1, -1, -1, "RoundRect", teApiRoundRect },
	{ 6,  4, -1, -1, "RunDLL", teApiRunDLL },
	{ 2, -1, -1, -1, "ScreenToClient", teApiScreenToClient },
	{ 2, -1, -1, -1, "SelectObject", teApiSelectObject },
	{ 3, -1, -1, -1, "SendInput", teApiSendInput },
	{ 2, -1, -1, -1, "SendMessage", teApiSendMessage },
	{ 6, -1, -1, -1, "SendMessageTimeout", teApiSendMessageTimeout },
	{ 4, -1, -1, -1, "SendNotifyMessage", teApiSendNotifyMessage },
	{ 2, -1, -1, -1, "SetBkColor", teApiSetBkColor },
	{ 2, -1, -1, -1, "SetBkMode", teApiSetBkMode },
	{ 1, -1, -1, -1, "SetCapture", teApiSetCapture },
	{ 3, -1, -1, -1, "SetClassLongPtr", teApiSetClassLongPtr },
	{ 1,  0, -1, -1, "SetClipboardData", teApiSetClipboardData },
	{ 1,  0, -1, -1, "SetCurrentDirectory", teApiSetCurrentDirectory },//use wsh.CurrentDirectory
	{ 1, -1, -1, -1, "SetCursor", teApiSetCursor },
	{ 2, -1, -1, -1, "SetCursorPos", teApiSetCursorPos },
	{ 2, -1, -1, -1, "SetDCBrushColor", teApiSetDCBrushColor },
	{ 2, -1, -1, -1, "SetDCPenColor", teApiSetDCPenColor },
	{ 1,  0, -1, -1, "SetDllDirectory", teApiSetDllDirectory },
	{ 1, -1, -1, -1, "SetEvent", teApiSetEvent },
	{ 3, -1, -1, -1, "SetFilePointer", teApiSetFilePointer },
	{ 4, -1, -1, -1, "SetFileTime", teApiSetFileTime },
	{ 1, -1, -1, -1, "SetFocus", teApiSetFocus },
	{ 1, -1, -1, -1, "SetForegroundWindow", teApiSetForegroundWindow },
	{ 1, -1, -1, -1, "SetKeyboardState", teApiSetKeyboardState },
	{ 4, -1, -1, -1, "SetLayeredWindowAttributes", teApiSetLayeredWindowAttributes },
	{ 3, -1, -1, -1, "SetMenuDefaultItem", teApiSetMenuDefaultItem },
	{ 2, -1, -1, -1, "SetMenuInfo", teApiSetMenuInfo },
	{ 5, -1, -1, -1, "SetMenuItemBitmaps", teApiSetMenuItemBitmaps },
	{ 4, -1, -1, -1, "SetMenuItemInfo", teApiSetMenuItemInfo },
	{ 1, -1, -1, -1, "SetRect", teApiSetRect },
	{ 3, -1, -1, -1, "SetSysColors", teApiSetSysColors },
	{ 2, -1, -1, -1, "SetTextColor", teApiSetTextColor },
	{ 3, -1, -1, -1, "SetTimer", teApiSetTimer },
	{ 3, -1, -1, -1, "SetWindowLongPtr", teApiSetWindowLongPtr },
	{ 7, -1, -1, -1, "SetWindowPos", teApiSetWindowPos },
	{ 2, -1,  1, -1, "SetWindowText", teApiSetWindowText },
	{ 3,  1,  2, -1, "SetWindowTheme", teApiSetWindowTheme },
	{ 3, -1, -1, -1, "SHChangeNotification_Lock", teApiSHChangeNotification_Lock },
	{ 1, -1, -1, -1, "SHChangeNotification_Unlock", teApiSHChangeNotification_Unlock },
	{ 4, -1, -1, -1, "SHChangeNotify", teApiSHChangeNotify },
	{ 1, -1, -1, -1, "SHChangeNotifyDeregister", teApiSHChangeNotifyDeregister },
	{ 6, -1, -1, -1, "SHChangeNotifyRegister", teApiSHChangeNotifyRegister },
	{ 4,  0, -1, -1, "SHCreateStreamOnFileEx", teApiSHCreateStreamOnFileEx },
	{ 6,  0, -1, -1, "SHDefExtractIcon", teApiSHDefExtractIcon },
	{ 4, -1, -1, -1, "SHDoDragDrop", teApiSHDoDragDrop },
	{ 2, -1, -1, -1, "Shell_NotifyIcon", teApiShell_NotifyIcon },
	{ 6, -1,  3,  4, "ShellExecute", teApiShellExecute },
	{ 1, -1, -1, -1, "ShellExecuteEx", teApiShellExecuteEx },
	{ 3,  1, -1, -1, "SHEmptyRecycleBin", teApiSHEmptyRecycleBin },
	{ 1, -1, -1, -1, "SHFileOperation", teApiSHFileOperation },
	{ 1, -1, -1, -1, "SHFreeNameMappings", teApiSHFreeNameMappings },
	{ 2, -1, -1, -1, "SHFreeShared", teApiSHFreeShared },
	{ 4, -1, -1, -1, "SHGetDataFromIDList", teApiSHGetDataFromIDList },
	{ 5, 10, -1, -1, "SHGetFileInfo", teApiSHGetFileInfo },
	{ 1, -1, -1, -1, "SHGetImageList", teApiSHGetImageList },
	{ 0, -1, -1, -1, "ShouldAppsUseDarkMode", teApiShouldAppsUseDarkMode },
	{ 2, -1, -1, -1, "ShowWindow", teApiShowWindow },
	{ 3, -1, -1, -1, "SHParseDisplayName", teApiSHParseDisplayName },
	{ 6,  2,  3,  4, "ShRunDialog", teApiShRunDialog },
	{ 1,  0, -1, -1, "SHSimpleIDListFromPath", teApiSHSimpleIDListFromPath },
	{ 2, -1, -1, -1, "SHTestTokenMembership", teApiSHTestTokenMembership },
	{ 1, -1, -1, -1, "sizeof", teApisizeof },
	{ 1, -1, -1, -1, "Sleep", teApiSleep },
	{ 3, -1,  1, -1, "sprintf", teApisprintf },
	{ 2,  0,  1, -1, "sscanf", teApisscanf },
	{ 2,  0,  1, -1, "StrCmpI", teApiStrCmpI },
	{ 2,  0,  1, -1, "StrCmpLogical", teApiStrCmpLogical },
	{ 3,  0,  1, -1, "StrCmpNI", teApiStrCmpNI },
	{11, -1, -1, -1, "StretchBlt", teApiStretchBlt },
	{ 1, -1, -1, -1, "StrFormatByteSize", teApiStrFormatByteSize },
	{ 1,  0, -1, -1, "StrLen", teApiStrLen },
	{ 1, -1, -1, -1, "SysAllocString", teApiSysAllocString },
	{ 2, -1, -1, -1, "SysAllocStringByteLen", teApiSysAllocStringByteLen },
	{ 2, -1, -1, -1, "SysAllocStringLen", teApiSysAllocStringLen },
	{ 1, -1, -1, -1, "SysStringByteLen", teApiSysStringByteLen },
	{ 4, -1, -1, -1, "SystemParametersInfo", teApiSystemParametersInfo },
	{ 6, -1, -1, -1, "TrackPopupMenuEx", teApiTrackPopupMenuEx },
	{ 1, -1, -1, -1, "TranslateMessage", teApiTranslateMessage },
	{11, -1, -1, -1, "TransparentBlt", teApiTransparentBlt },
	{ 1, -1, -1, -1, "ULowPart", teApiULowPart },
	{ 2, -1, -1, -1, "UnregisterHotKey", teApiUnregisterHotKey },
	{ 2, -1, -1, -1, "UQuadAdd", teApiUQuadAdd },
	{ 2, -1, -1, -1, "UQuadCmp", teApiUQuadCmp },
	{ 1, -1, -1, -1, "UQuadPart", teApiUQuadPart },
	{ 2, -1, -1, -1, "UQuadSub", teApiUQuadSub },
	{ 1,  0, -1, -1, "UrlCreateFromPath", teApiUrlCreateFromPath },
	{ 3,  1,  2, -1, "URLDownloadToFile", teApiURLDownloadToFile },
	{ 2, -1, -1, -1, "WaitForInputIdle", teApiWaitForInputIdle },
	{ 2, -1, -1, -1, "WaitForSingleObject", teApiWaitForSingleObject },
	{ 2, -1, -1, -1, "WideCharToMultiByte", teApiWideCharToMultiByte },
	{ 1, -1, -1, -1, "WindowFromPoint", teApiWindowFromPoint },
	{ 2, -1, -1, -1, "WriteFile", teApiWriteFile },
#ifdef _DEBUG
	{ 0, -1, -1, -1, "Test", teApiTest },
	{ 4,  1, -1, -1, "GetGlyphIndices", teApiGetGlyphIndices },
	{ 4,  0, -1, -1, "mciSendString", teApimciSendString },
#endif
};

VOID InitAPI()
{
	for (int i = 0; i < _countof(pTEStructs); ++i) {
		teSetUM(MAP_SS, pTEStructs[i].name, i);
	}
	for (int i = 0; i < _countof(dispAPI); ++i) {
		teSetUM(MAP_API, dispAPI[i].name, i);
	}
#ifdef _DEBUG
	//Check Otder
	for (int i = 0; i < _countof(pTEStructs) - 1; ++i) {
		if (lstrcmpiA(pTEStructs[i].name, pTEStructs[i + 1].name) >= 0) {
			MessageBoxA(NULL, pTEStructs[i].name, "pTEStructs", MB_OK);
		}
		teCheckMethod(pTEStructs[i].name, pTEStructs[i].pMethod, pTEStructs[i].nCount);
	}
	teCheckMethod("methodObject", methodObject, _countof(methodObject));
#endif
}

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
#pragma warning( push )
#pragma warning( disable: 4838 )
	}
#pragma warning( pop )
	;
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
	int nIndex = teUMSearch(MAP_API, *rgszNames);
	if (nIndex >= 0) {
		*rgDispId = nIndex + TE_METHOD;
		return S_OK;
	}
	if (teStrCmpIWA(*rgszNames, "ADBSTRM") == 0) {
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
#ifdef _DEBUG
			//::OutputDebugStringA("API: ");
			//::OutputDebugStringA(pApi->name);
			//::OutputDebugStringA("\n");
#endif
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

#ifdef USE_OBJECTAPI
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

#endif
