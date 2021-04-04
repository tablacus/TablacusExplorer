#pragma once

#include "common.h"

TEmethod tesNULL[] =
{
	{ 0, NULL }
};

TEmethod tesBITMAP[] =
{
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmType), "bmType" },
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmWidth), "bmWidth" },
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmHeight), "bmHeight" },
	{ (VT_I4 << TE_VT) + offsetof(BITMAP, bmWidthBytes), "bmWidthBytes" },
	{ (VT_UI2 << TE_VT) + offsetof(BITMAP, bmPlanes), "bmPlanes" },
	{ (VT_UI2 << TE_VT) + offsetof(BITMAP, bmBitsPixel), "bmBitsPixel" },
	{ (VT_PTR << TE_VT) + offsetof(BITMAP, bmBits), "bmBits" },
	{ 0, NULL }
};

TEmethod tesCHOOSECOLOR[] =
{
	{ (VT_I4 << TE_VT) + offsetof(CHOOSECOLOR, lStructSize), "lStructSize" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, hwndOwner), "hwndOwner" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, hInstance), "hInstance" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSECOLOR, rgbResult), "rgbResult" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, lpCustColors), "lpCustColors" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSECOLOR, Flags), "Flags" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, lCustData), "lCustData" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSECOLOR, lpfnHook), "lpfnHook" },
	{ (VT_BSTR << TE_VT) + offsetof(CHOOSECOLOR, lpTemplateName), "lpTemplateName" },
	{ 0, NULL }
};

TEmethod tesCHOOSEFONT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, lStructSize), "lStructSize" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, hwndOwner), "hwndOwner" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, hDC), "hDC" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, lpLogFont), "lpLogFont" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, iPointSize), "iPointSize" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, Flags), "Flags" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, rgbColors), "rgbColors" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, lCustData), "lCustData" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, lpfnHook), "lpfnHook" },
	{ (VT_BSTR << TE_VT) + offsetof(CHOOSEFONT, lpTemplateName), "lpTemplateName" },
	{ (VT_PTR << TE_VT) + offsetof(CHOOSEFONT, hInstance), "hInstance" },
	{ (VT_BSTR << TE_VT) + offsetof(CHOOSEFONT, lpszStyle), "lpszStyle" },
	{ (VT_UI2 << TE_VT) + offsetof(CHOOSEFONT, nFontType), "nFontType" },
	{ (VT_UI2 << TE_VT) + offsetof(CHOOSEFONT, ___MISSING_ALIGNMENT__), "___MISSING_ALIGNMENT__" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, nSizeMin), "nSizeMin" },
	{ (VT_I4 << TE_VT) + offsetof(CHOOSEFONT, nSizeMax), "nSizeMax" },
	{ 0, NULL }
};

TEmethod tesCOPYDATASTRUCT[] =
{
	{ (VT_PTR << TE_VT) + offsetof(COPYDATASTRUCT, dwData), "dwData" },
	{ (VT_I4 << TE_VT) + offsetof(COPYDATASTRUCT, cbData), "cbData" },
	{ (VT_PTR << TE_VT) + offsetof(COPYDATASTRUCT, lpData), "lpData" },
	{ 0, NULL }
};

TEmethod tesDIBSECTION[] =
{
	{ (VT_PTR << TE_VT) + offsetof(DIBSECTION, dsBm), "dsBm" },
	{ (VT_PTR << TE_VT) + offsetof(DIBSECTION, dsBmih), "dsBmih" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsBitfields), "dsBitfields0" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsBitfields) + sizeof(DWORD), "dsBitfields1" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsBitfields) + sizeof(DWORD) * 2, "dsBitfields2" },
	{ (VT_PTR << TE_VT) + offsetof(DIBSECTION, dshSection), "dshSection" },
	{ (VT_I4 << TE_VT) + offsetof(DIBSECTION, dsOffset), "dsOffset" },
	{ 0, NULL }
};

TEmethod tesEXCEPINFO[] =
{
	{ (VT_UI2 << TE_VT) + offsetof(EXCEPINFO, wCode), "wCode" },
	{ (VT_UI2 << TE_VT) + offsetof(EXCEPINFO, wReserved), "wReserved" },
	{ (VT_BSTR << TE_VT) + offsetof(EXCEPINFO, bstrSource), "bstrSource" },
	{ (VT_BSTR << TE_VT) + offsetof(EXCEPINFO, bstrDescription), "bstrDescription" },
	{ (VT_BSTR << TE_VT) + offsetof(EXCEPINFO, bstrHelpFile), "bstrHelpFile" },
	{ (VT_I4 << TE_VT) + offsetof(EXCEPINFO, dwHelpContext), "dwHelpContext" },
	{ (VT_PTR << TE_VT) + offsetof(EXCEPINFO, pvReserved), "pvReserved" },
	{ (VT_PTR << TE_VT) + offsetof(EXCEPINFO, pfnDeferredFillIn), "pfnDeferredFillIn" },
	{ (VT_I4 << TE_VT) + offsetof(EXCEPINFO, scode), "scode" },
	{ 0, NULL }
};

TEmethod tesFINDREPLACE[] =
{
	{ (VT_I4 << TE_VT) + offsetof(FINDREPLACE, lStructSize), "lStructSize" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, hwndOwner), "hwndOwner" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, hInstance), "hInstance" },
	{ (VT_I4 << TE_VT) + offsetof(FINDREPLACE, Flags), "Flags" },
	{ (VT_BSTR << TE_VT) + offsetof(FINDREPLACE, lpstrFindWhat), "lpstrFindWhat" },
	{ (VT_BSTR << TE_VT) + offsetof(FINDREPLACE, lpstrReplaceWith), "lpstrReplaceWith" },
	{ (VT_UI2 << TE_VT) + offsetof(FINDREPLACE, wFindWhatLen), "wFindWhatLen" },
	{ (VT_UI2 << TE_VT) + offsetof(FINDREPLACE, wReplaceWithLen), "wReplaceWithLen" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, lCustData), "lCustData" },
	{ (VT_PTR << TE_VT) + offsetof(FINDREPLACE, lpfnHook), "lpfnHook" },
	{ (VT_BSTR << TE_VT) + offsetof(FINDREPLACE, lpTemplateName), "lpTemplateName" },
	{ 0, NULL }
};

TEmethod tesFOLDERSETTINGS[] =
{
	{ (VT_I4 << TE_VT) + offsetof(FOLDERSETTINGS, ViewMode), "ViewMode" },
	{ (VT_I4 << TE_VT) + offsetof(FOLDERSETTINGS, fFlags), "fFlags" },
	{ (VT_I4 << TE_VT) + (SB_Options - 1) * 4, "Options" },
	{ (VT_I4 << TE_VT) + (SB_ViewFlags - 1) * 4, "ViewFlags" },
	{ (VT_I4 << TE_VT) + (SB_IconSize - 1) * 4, "ImageSize" },
	{ 0, NULL }
};

TEmethod tesHDITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, mask), "mask" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, cxy), "cxy" },
	{ (VT_BSTR << TE_VT) + offsetof(HDITEM, pszText), "pszText" },
	{ (VT_PTR << TE_VT) + offsetof(HDITEM, hbm), "hbm" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, cchTextMax), "cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, fmt), "fmt" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, lParam), "lParam" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, iImage), "iImage" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, iOrder), "iOrder" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, type), "type" },
	{ (VT_PTR << TE_VT) + offsetof(HDITEM, pvFilter), "pvFilter" },
	{ (VT_I4 << TE_VT) + offsetof(HDITEM, state), "state" },
	{ 0, NULL }
};

TEmethod tesICONINFO[] =
{
	{ (VT_BOOL << TE_VT) + offsetof(ICONINFO, fIcon), "fIcon" },
	{ (VT_I4 << TE_VT) + offsetof(ICONINFO, xHotspot), "xHotspot" },
	{ (VT_I4 << TE_VT) + offsetof(ICONINFO, yHotspot), "yHotspot" },
	{ (VT_PTR << TE_VT) + offsetof(ICONINFO, hbmMask), "hbmMask" },
	{ (VT_PTR << TE_VT) + offsetof(ICONINFO, hbmColor), "hbmColor" },
	{ 0, NULL }
};

TEmethod tesICONMETRICS[] =
{
	{ (VT_I4 << TE_VT) + offsetof(ICONMETRICS, iHorzSpacing), "iHorzSpacing" },
	{ (VT_I4 << TE_VT) + offsetof(ICONMETRICS, iVertSpacing), "iVertSpacing" },
	{ (VT_I4 << TE_VT) + offsetof(ICONMETRICS, iTitleWrap), "iTitleWrap" },
	{ (VT_PTR << TE_VT) + offsetof(ICONMETRICS, lfFont), "lfFont" },
	{ 0, NULL }
};

TEmethod tesKEYBDINPUT[] =
{
	{ (VT_I4 << TE_VT), "type" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, wVk), "wVk" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, wScan), "wScan" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, dwFlags), "dwFlags" },
	{ (VT_I4 << TE_VT) + offsetof(KEYBDINPUT, time), "time" },
	{ (VT_PTR << TE_VT) + offsetof(KEYBDINPUT, dwExtraInfo), "dwExtraInfo" },
	{ 0, NULL }
};

TEmethod tesLOGFONT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfHeight), "lfHeight" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfWidth), "lfWidth" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfEscapement), "lfEscapement" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfOrientation), "lfOrientation" },
	{ (VT_I4 << TE_VT) + offsetof(LOGFONT, lfWeight), "lfWeight" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfItalic), "lfItalic" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfUnderline), "lfUnderline" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfStrikeOut), "lfStrikeOut" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfCharSet), "lfCharSet" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfOutPrecision), "lfOutPrecision" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfClipPrecision), "lfClipPrecision" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfQuality), "lfQuality" },
	{ (VT_UI1 << TE_VT) + offsetof(LOGFONT, lfPitchAndFamily), "lfPitchAndFamily" },
	{ (VT_LPWSTR << TE_VT) + offsetof(LOGFONT, lfFaceName), "lfFaceName" },
	{ 0, NULL }
};

TEmethod tesLVBKIMAGE[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, ulFlags), "ulFlags" },
	{ (VT_PTR << TE_VT) + offsetof(LVBKIMAGE, hbm), "hbm" },
	{ (VT_BSTR << TE_VT) + offsetof(LVBKIMAGE, pszImage), "pszImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, cchImageMax), "cchImageMax" },
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, xOffsetPercent), "xOffsetPercent" },
	{ (VT_I4 << TE_VT) + offsetof(LVBKIMAGE, yOffsetPercent), "yOffsetPercent" },
	{ 0, NULL }
};

TEmethod tesLVFINDINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVFINDINFO, flags), "flags" },
	{ (VT_BSTR << TE_VT) + offsetof(LVFINDINFO, psz), "psz" },
	{ (VT_I4 << TE_VT) + offsetof(LVFINDINFO, lParam), "lParam" },
	{ (VT_CY << TE_VT) + offsetof(LVFINDINFO, pt), "pt" },
	{ (VT_I4 << TE_VT) + offsetof(LVFINDINFO, vkDirection), "vkDirection" },
	{ 0, NULL }
};

TEmethod tesLVGROUP[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, mask), "mask" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszHeader), "pszHeader" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchHeader), "cchHeader" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszFooter), "pszFooter" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchFooter), "cchFooter" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iGroupId), "iGroupId" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, stateMask), "stateMask" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, state), "state" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, uAlign), "uAlign" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszSubtitle), "pszSubtitle" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchSubtitle), "cchSubtitle" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszTask), "pszTask" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchTask), "cchTask" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszDescriptionTop), "pszDescriptionTop" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchDescriptionTop), "cchDescriptionTop" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszDescriptionBottom), "pszDescriptionBottom" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchDescriptionBottom), "cchDescriptionBottom" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iTitleImage), "iTitleImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iExtendedImage), "iExtendedImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, iFirstItem), "iFirstItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cItems), "cItems" },
	{ (VT_BSTR << TE_VT) + offsetof(LVGROUP, pszSubsetTitle), "pszSubsetTitle" },
	{ (VT_I4 << TE_VT) + offsetof(LVGROUP, cchSubsetTitle), "cchSubsetTitle" },
	{ 0, NULL }
};

TEmethod tesLVHITTESTINFO[] =
{
	{ (VT_CY << TE_VT) + offsetof(LVHITTESTINFO, pt), "pt" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, flags), "flags" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, iItem), "iItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, iSubItem ), "iSubItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVHITTESTINFO, iGroup), "iGroup" },
	{ 0, NULL }
};

TEmethod tesLVITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, mask), "mask" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iItem), "iItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iSubItem), "iSubItem" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, state), "state" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, stateMask), "stateMask" },
	{ (VT_BSTR << TE_VT) + offsetof(LVITEM, pszText), "pszText" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, cchTextMax), "cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iImage), "iImage" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, lParam), "lParam" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iIndent), "iIndent" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iGroupId), "iGroupId" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, cColumns), "cColumns" },
	{ (VT_PTR << TE_VT) + offsetof(LVITEM, puColumns), "puColumns" },
	{ (VT_PTR << TE_VT) + offsetof(LVITEM, piColFmt), "piColFmt" },
	{ (VT_I4 << TE_VT) + offsetof(LVITEM, iGroup), "iGroup" },
	{ 0, NULL }
};

TEmethod tesMENUITEMINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, fMask), "fMask" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, fType), "fType" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, fState), "fState" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, wID), "wID" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hSubMenu), "hSubMenu" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hbmpChecked), "hbmpChecked" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hbmpUnchecked), "hbmpUnchecked" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, dwItemData), "dwItemData" },
	{ (VT_BSTR << TE_VT) + offsetof(MENUITEMINFO, dwTypeData), "dwTypeData" },
	{ (VT_I4 << TE_VT) + offsetof(MENUITEMINFO, cch), "cch" },
	{ (VT_PTR << TE_VT) + offsetof(MENUITEMINFO, hbmpItem), "hbmpItem" },
	{ 0, NULL }
};

TEmethod tesMONITORINFOEX[] =
{
	{ (VT_I4 << TE_VT), "cbSize" },
	{ (VT_CARRAY << TE_VT) + offsetof(MONITORINFOEX, rcMonitor), "rcMonitor" },
	{ (VT_CARRAY << TE_VT) + offsetof(MONITORINFOEX, rcWork), "rcWork" },
	{ (VT_I4 << TE_VT) + offsetof(MONITORINFOEX, dwFlags), "dwFlags" },
	{ (VT_LPWSTR << TE_VT) + offsetof(MONITORINFOEX, szDevice), "szDevice" },
	{ 0, NULL }
};


TEmethod tesMOUSEINPUT[] =
{
	{ (VT_I4 << TE_VT), "type" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, dx), "dx" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, dy), "dy" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, mouseData), "mouseData" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, dwFlags), "dwFlags" },
	{ (VT_I4 << TE_VT) + offsetof(MOUSEINPUT, time), "time" },
	{ (VT_PTR << TE_VT) + offsetof(MOUSEINPUT, dwExtraInfo), "dwExtraInfo" },
	{ 0, NULL }
};

TEmethod tesMSG[] =
{
	{ (VT_PTR << TE_VT) + offsetof(MSG, hwnd), "hwnd" },
	{ (VT_I4 << TE_VT) + offsetof(MSG, message), "message" },
	{ (VT_PTR << TE_VT) + offsetof(MSG, wParam), "wParam" },
	{ (VT_PTR << TE_VT) + offsetof(MSG, lParam), "lParam" },
	{ (VT_I4 << TE_VT) + offsetof(MSG, time), "time" },
	{ (VT_CY << TE_VT) + offsetof(MSG, pt), "pt" },
	{ 0, NULL }
};

TEmethod tesNONCLIENTMETRICS[] =
{
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iBorderWidth), "iBorderWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iScrollWidth), "iScrollWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iScrollHeight), "iScrollHeight" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iCaptionWidth), "iCaptionWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iCaptionHeight), "iCaptionHeight" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfCaptionFont), "lfCaptionFont" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iSmCaptionWidth), "iSmCaptionWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iSmCaptionHeight), "iSmCaptionHeight" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfSmCaptionFont), "lfSmCaptionFont" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iMenuWidth), "iMenuWidth" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iMenuHeight), "iMenuHeight" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfMenuFont), "lfMenuFont" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfStatusFont), "lfStatusFont" },
	{ (VT_PTR << TE_VT) + offsetof(NONCLIENTMETRICS, lfMessageFont), "lfMessageFont" },
	{ (VT_I4 << TE_VT) + offsetof(NONCLIENTMETRICS, iPaddedBorderWidth), "iPaddedBorderWidth" },
	{ 0, NULL }
};

TEmethod tesNOTIFYICONDATA[] =
{
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, cbSize), "cbSize" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, hWnd), "hWnd" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uID), "uID" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uFlags), "uFlags" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uCallbackMessage), "uCallbackMessage" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, hIcon), "hIcon" },
	{ (VT_LPWSTR << TE_VT) + offsetof(NOTIFYICONDATA, szTip), "szTip" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwState), "dwState" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwState), "dwState" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwStateMask), "dwStateMask" },
	{ (VT_LPWSTR << TE_VT) + offsetof(NOTIFYICONDATA, szInfo), "szInfo" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uTimeout), "uTimeout" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, uVersion), "uVersion" },
	{ (VT_LPWSTR << TE_VT) + offsetof(NOTIFYICONDATA, szInfoTitle), "szInfoTitle" },
	{ (VT_I4 << TE_VT) + offsetof(NOTIFYICONDATA, dwInfoFlags), "dwInfoFlags" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, guidItem), "guidItem" },
	{ (VT_PTR << TE_VT) + offsetof(NOTIFYICONDATA, hBalloonIcon), "hBalloonIcon" },
	{ 0, NULL }
};

TEmethod tesNMCUSTOMDRAW[] =
{
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, hdr), "hdr" },
	{ (VT_I4 << TE_VT) + offsetof(NMCUSTOMDRAW, dwDrawStage), "dwDrawStage" },
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, hdc), "hdc" },
	{ (VT_CARRAY << TE_VT) + offsetof(NMCUSTOMDRAW, rc), "rc" },
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, dwItemSpec), "dwItemSpec" },
	{ (VT_I4 << TE_VT) + offsetof(NMCUSTOMDRAW, uItemState), "uItemState" },
	{ (VT_PTR << TE_VT) + offsetof(NMCUSTOMDRAW, lItemlParam), "lItemlParam" },
	{ 0, NULL }
};

TEmethod tesNMLVCUSTOMDRAW[] =
{
	{ (VT_PTR << TE_VT) + offsetof(NMLVCUSTOMDRAW, nmcd), "nmcd" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, clrText), "clrText" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, clrTextBk), "clrTextBk" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iSubItem), "iSubItem" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, dwItemType), "dwItemType" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, clrFace), "clrFace" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iIconEffect), "iIconEffect" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iIconPhase), "iIconPhase" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iPartId), "iPartId" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, iStateId), "iStateId" },
	{ (VT_CARRAY << TE_VT) + offsetof(NMLVCUSTOMDRAW, rcText), "rcText" },
	{ (VT_I4 << TE_VT) + offsetof(NMLVCUSTOMDRAW, uAlign), "uAlign" },
	{ 0, NULL }
};

TEmethod tesNMTVCUSTOMDRAW[] =
{
	{ (VT_PTR << TE_VT) + offsetof(NMTVCUSTOMDRAW, nmcd), "nmcd" },
	{ (VT_I4 << TE_VT) + offsetof(NMTVCUSTOMDRAW, clrText), "clrText" },
	{ (VT_I4 << TE_VT) + offsetof(NMTVCUSTOMDRAW, clrTextBk), "clrTextBk" },
	{ (VT_I4 << TE_VT) + offsetof(NMTVCUSTOMDRAW, iLevel), "iLevel" },
	{ 0, NULL }
};

TEmethod tesNMHDR[] =
{
	{ (VT_PTR << TE_VT) + offsetof(NMHDR, hwndFrom), "hwndFrom" },
	{ (VT_I4 << TE_VT) + offsetof(NMHDR, idFrom), "idFrom" },
	{ (VT_I4 << TE_VT) + offsetof(NMHDR, code), "code" },
	{ 0, NULL }
};

TEmethod tesOSVERSIONINFOEX[] =
{
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwOSVersionInfoSize), "dwOSVersionInfoSize" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwMajorVersion), "dwMajorVersion" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwMinorVersion), "dwMinorVersion" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwBuildNumber), "dwBuildNumber" },
	{ (VT_I4 << TE_VT) + offsetof(OSVERSIONINFOEX, dwPlatformId), "dwPlatformId" },
	{ (VT_LPWSTR << TE_VT) + offsetof(OSVERSIONINFOEX, szCSDVersion), "szCSDVersion" },
	{ (VT_UI2 << TE_VT) + offsetof(OSVERSIONINFOEX, wServicePackMajor), "wServicePackMajor" },
	{ (VT_UI2 << TE_VT) + offsetof(OSVERSIONINFOEX, wServicePackMinor), "wServicePackMinor" },
	{ (VT_UI2 << TE_VT) + offsetof(OSVERSIONINFOEX, wSuiteMask), "wSuiteMask" },
	{ (VT_UI1 << TE_VT) + offsetof(OSVERSIONINFOEX, wProductType), "wProductType" },
	{ (VT_UI1 << TE_VT) + offsetof(OSVERSIONINFOEX, wReserved), "wReserved" },
	{ 0, NULL }
};

TEmethod tesPAINTSTRUCT[] =
{
	{ (VT_PTR << TE_VT) + offsetof(PAINTSTRUCT, hdc), "hdc" },
	{ (VT_BOOL << TE_VT) + offsetof(PAINTSTRUCT, fErase), "fErase" },
	{ (VT_CARRAY << TE_VT) + offsetof(PAINTSTRUCT, rcPaint), "rcPaint" },
	{ (VT_BOOL << TE_VT) + offsetof(PAINTSTRUCT, fRestore), "fRestore" },
	{ (VT_BOOL << TE_VT) + offsetof(PAINTSTRUCT, fIncUpdate), "fIncUpdate" },
	{ (VT_UI1 << TE_VT) + offsetof(PAINTSTRUCT, rgbReserved), "rgbReserved" },
	{ 0, NULL }
};

TEmethod tesPOINT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(POINT, x), "x" },
	{ (VT_I4 << TE_VT) + offsetof(POINT, y), "y" },
	{ 0, NULL }
};

TEmethod tesRECT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(RECT, left), "left" },
	{ (VT_I4 << TE_VT) + offsetof(RECT, top), "top" },
	{ (VT_I4 << TE_VT) + offsetof(RECT, right), "right" },
	{ (VT_I4 << TE_VT) + offsetof(RECT, bottom), "bottom" },
	{ 0, NULL }
};

TEmethod tesSHELLEXECUTEINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, fMask), "fMask" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hwnd), "hwnd" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpVerb), "lpVerb" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpFile), "lpFile" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpParameters), "lpParameters" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpDirectory), "lpDirectory" },
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, nShow), "nShow" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hInstApp), "hInstApp" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpIDList), "lpIDList" },
	{ (VT_BSTR << TE_VT) + offsetof(SHELLEXECUTEINFO, lpClass), "lpClass" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hkeyClass), "hkeyClass" },
	{ (VT_I4 << TE_VT) + offsetof(SHELLEXECUTEINFO, dwHotKey), "dwHotKey" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hIcon), "hIcon" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hMonitor), "hMonitor" },
	{ (VT_PTR << TE_VT) + offsetof(SHELLEXECUTEINFO, hProcess), "hProcess" },
	{ 0, NULL }
};

TEmethod tesSHFILEINFO[] =
{
	{ (VT_PTR << TE_VT) + offsetof(SHFILEINFO, hIcon), "hIcon" },
	{ (VT_I4 << TE_VT) + offsetof(SHFILEINFO, iIcon), "iIcon" },
	{ (VT_I4 << TE_VT) + offsetof(SHFILEINFO, dwAttributes), "dwAttributes" },
	{ (VT_LPWSTR << TE_VT) + offsetof(SHFILEINFO, szDisplayName), "szDisplayName" },
	{ (VT_LPWSTR << TE_VT) + offsetof(SHFILEINFO, szTypeName), "szTypeName" },
	{ 0, NULL }
};

TEmethod tesSHFILEOPSTRUCT[] =
{
	{ (VT_PTR << TE_VT) + offsetof(SHFILEOPSTRUCT, hwnd), "hwnd" },
	{ (VT_I4 << TE_VT) + offsetof(SHFILEOPSTRUCT, wFunc), "wFunc" },
	{ (VT_BSTR << TE_VT) + offsetof(SHFILEOPSTRUCT, pFrom), "pFrom" },
	{ (VT_BSTR << TE_VT) + offsetof(SHFILEOPSTRUCT, pTo), "pTo" },
	{ (VT_UI2 << TE_VT) + offsetof(SHFILEOPSTRUCT, fFlags), "fFlags" },
	{ (VT_BOOL << TE_VT) + offsetof(SHFILEOPSTRUCT, fAnyOperationsAborted), "fAnyOperationsAborted" },
	{ (VT_PTR << TE_VT) + offsetof(SHFILEOPSTRUCT, hNameMappings), "hNameMappings" },
	{ (VT_BSTR << TE_VT) + offsetof(SHFILEOPSTRUCT, lpszProgressTitle), "lpszProgressTitle" },
	{ 0, NULL }
};

TEmethod tesSIZE[] =
{
	{ (VT_I4 << TE_VT) + offsetof(SIZE, cx), "cx" },
	{ (VT_I4 << TE_VT) + offsetof(SIZE, cy), "cy" },
	{ 0, NULL }
};

TEmethod tesTCHITTESTINFO[] =
{
	{ (VT_CY << TE_VT) + offsetof(TCHITTESTINFO, pt), "pt" },
	{ (VT_I4 << TE_VT) + offsetof(TCHITTESTINFO, flags), "flags" },
	{ 0, NULL }
};

TEmethod tesTCITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, mask), "mask" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, dwState), "dwState" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, dwStateMask), "dwStateMask" },
	{ (VT_BSTR << TE_VT) + offsetof(TCITEM, pszText), "pszText" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, cchTextMax), "cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(TCITEM, iImage), "iImage" },
	{ (VT_PTR << TE_VT) + offsetof(TCITEM, lParam), "lParam" },
	{ 0, NULL }
};

TEmethod tesTOOLINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(TOOLINFO, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(TOOLINFO, uFlags), "uFlags" },
	{ (VT_PTR << TE_VT) + offsetof(TOOLINFO, hwnd), "hwnd" },
	{ (VT_PTR << TE_VT) + offsetof(TOOLINFO, uId), "uId" },
	{ (VT_CARRAY << TE_VT) + offsetof(TOOLINFO, rect), "rect" },
	{ (VT_PTR << TE_VT) + offsetof(TOOLINFO, hinst), "hinst" },
	{ (VT_BSTR << TE_VT) + offsetof(TOOLINFO, uId), "lpszText" },
	{ (VT_PTR << TE_VT) + offsetof(TOOLINFO, uId), "lParam" },
	{ (VT_PTR << TE_VT) + offsetof(TOOLINFO, uId), "lpReserved" },
	{ 0, NULL }
};

TEmethod tesTVHITTESTINFO[] =
{
	{ (VT_CY << TE_VT) + offsetof(TVHITTESTINFO, pt), "pt" },
	{ (VT_I4 << TE_VT) + offsetof(TVHITTESTINFO, flags), "flags" },
	{ (VT_PTR << TE_VT) + offsetof(TVHITTESTINFO, hItem), "hItem" },
	{ 0, NULL }
};

TEmethod tesTVITEM[] =
{
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, mask), "mask" },
	{ (VT_PTR << TE_VT) + offsetof(TVITEM, hItem), "hItem" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, state), "state" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, stateMask), "stateMask" },
	{ (VT_BSTR << TE_VT) + offsetof(TVITEM, pszText), "pszText" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, cchTextMax), "cchTextMax" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, iImage), "iImage" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, iSelectedImage), "iSelectedImage" },
	{ (VT_I4 << TE_VT) + offsetof(TVITEM, cChildren), "cChildren" },
	{ (VT_PTR << TE_VT) + offsetof(TVITEM, lParam), "lParam" },
	{ 0, NULL }
};

TEmethod tesWIN32_FIND_DATA[] =
{
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, dwFileAttributes), "dwFileAttributes" },
	{ (VT_FILETIME << TE_VT) + offsetof(WIN32_FIND_DATA, ftCreationTime), "ftCreationTime" },
	{ (VT_FILETIME << TE_VT) + offsetof(WIN32_FIND_DATA, ftLastAccessTime), "ftLastAccessTime" },
	{ (VT_FILETIME << TE_VT) + offsetof(WIN32_FIND_DATA, ftLastWriteTime), "ftLastWriteTime" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, nFileSizeHigh), "nFileSizeHigh" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, nFileSizeLow), "nFileSizeLow" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, dwReserved0), "dwReserved0" },
	{ (VT_I4 << TE_VT) + offsetof(WIN32_FIND_DATA, dwReserved1), "dwReserved1" },
	{ (VT_LPWSTR << TE_VT) + offsetof(WIN32_FIND_DATA, cFileName), "cFileName" },
	{ (VT_LPWSTR << TE_VT) + offsetof(WIN32_FIND_DATA, cAlternateFileName), "cAlternateFileName" },
	{ 0, NULL }
};

TEmethod tesDRAWITEMSTRUCT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, CtlType), "CtlType" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, CtlID), "CtlID" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, itemID), "itemID" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, itemAction), "itemAction" },
	{ (VT_I4 << TE_VT) + offsetof(DRAWITEMSTRUCT, itemState), "itemState" },
	{ (VT_PTR << TE_VT) + offsetof(DRAWITEMSTRUCT, hwndItem), "hwndItem" },
	{ (VT_PTR << TE_VT) + offsetof(DRAWITEMSTRUCT, hDC), "hDC" },
	{ (VT_CARRAY << TE_VT) + offsetof(DRAWITEMSTRUCT, rcItem), "rcItem" },
	{ (VT_PTR << TE_VT) + offsetof(DRAWITEMSTRUCT, itemData), "itemData" },
	{ 0, NULL }
};

TEmethod tesMEASUREITEMSTRUCT[] =
{
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, CtlType), "CtlType" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, CtlID), "CtlID" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemID), "itemID" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemWidth), "itemWidth" },
	{ (VT_I4 << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemHeight), "itemHeight" },
	{ (VT_PTR << TE_VT) + offsetof(MEASUREITEMSTRUCT, itemData), "itemData" },
	{ 0, NULL }
};

TEmethod tesMENUINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, fMask), "fMask" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, dwStyle), "dwStyle" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, cyMax), "cyMax" },
	{ (VT_PTR << TE_VT) + offsetof(MENUINFO, hbrBack), "hbrBack" },
	{ (VT_I4 << TE_VT) + offsetof(MENUINFO, dwContextHelpID), "dwContextHelpID" },
	{ (VT_PTR << TE_VT) + offsetof(MENUINFO, dwMenuData), "dwMenuData" },
	{ 0, NULL }
};

TEmethod tesGUITHREADINFO[] =
{
	{ (VT_I4 << TE_VT) + offsetof(GUITHREADINFO, cbSize), "cbSize" },
	{ (VT_I4 << TE_VT) + offsetof(GUITHREADINFO, flags), "flags" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndActive), "hwndActive" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndFocus), "hwndFocus" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndCapture), "hwndCapture" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndMenuOwner), "hwndMenuOwner" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndMoveSize), "hwndMoveSize" },
	{ (VT_PTR << TE_VT) + offsetof(GUITHREADINFO, hwndCaret), "hwndCaret" },
	{ (VT_CARRAY << TE_VT) + offsetof(GUITHREADINFO, hwndCaret), "rcCaret" },
	{ 0, NULL }
};

TEStruct pTEStructs[] = {
	{ sizeof(BITMAP), "BITMAP", tesBITMAP },
	{ sizeof(BSTR), "BSTR", tesNULL },
	{ sizeof(BYTE), "BYTE", tesNULL },
	{ sizeof(char), "char", tesNULL },
	{ sizeof(CHOOSECOLOR), "CHOOSECOLOR", tesCHOOSECOLOR },
	{ sizeof(CHOOSEFONT), "CHOOSEFONT", tesCHOOSEFONT },
	{ sizeof(COPYDATASTRUCT), "COPYDATASTRUCT", tesCOPYDATASTRUCT },
	{ sizeof(DIBSECTION), "DIBSECTION", tesDIBSECTION },
	{ sizeof(DRAWITEMSTRUCT), "DRAWITEMSTRUCT", tesDRAWITEMSTRUCT },
	{ sizeof(DWORD), "DWORD", tesNULL },
	{ sizeof(EXCEPINFO), "EXCEPINFO", tesEXCEPINFO },
	{ sizeof(FINDREPLACE), "FINDREPLACE", tesFINDREPLACE },
	{ sizeof(FOLDERSETTINGS), "FOLDERSETTINGS", tesFOLDERSETTINGS },
	{ sizeof(GUID), "GUID", tesNULL },
	{ sizeof(GUITHREADINFO), "GUITHREADINFO", tesGUITHREADINFO },
	{ sizeof(HANDLE), "HANDLE", tesNULL },
	{ sizeof(HDITEM), "HDITEM", tesHDITEM },
	{ sizeof(ICONINFO), "ICONINFO", tesICONINFO },
	{ sizeof(ICONMETRICS), "ICONMETRICS", tesICONMETRICS },
	{ sizeof(int), "int", tesNULL },
	{ sizeof(KEYBDINPUT) + sizeof(DWORD), "KEYBDINPUT", tesKEYBDINPUT },
	{ 256, "KEYSTATE", tesNULL },
	{ sizeof(LOGFONT), "LOGFONT", tesLOGFONT },
	{ sizeof(LPWSTR), "LPWSTR", tesNULL },
	{ sizeof(LVBKIMAGE), "LVBKIMAGE", tesLVBKIMAGE },
	{ sizeof(LVFINDINFO), "LVFINDINFO", tesLVFINDINFO },
	{ sizeof(LVGROUP), "LVGROUP", tesLVGROUP },
	{ sizeof(LVHITTESTINFO), "LVHITTESTINFO", tesLVHITTESTINFO },
	{ sizeof(LVITEM), "LVITEM", tesLVITEM },
	{ sizeof(MEASUREITEMSTRUCT), "MEASUREITEMSTRUCT", tesMEASUREITEMSTRUCT },
	{ sizeof(MENUINFO), "MENUINFO", tesMENUINFO },
	{ sizeof(MENUITEMINFO), "MENUITEMINFO", tesMENUITEMINFO },
	{ sizeof(MONITORINFOEX), "MONITORINFOEX", tesMONITORINFOEX },
	{ sizeof(MOUSEINPUT) + sizeof(DWORD), "MOUSEINPUT", tesMOUSEINPUT },
	{ sizeof(MSG), "MSG", tesMSG },
	{ sizeof(NMCUSTOMDRAW), "NMCUSTOMDRAW", tesNMCUSTOMDRAW },
	{ sizeof(NMLVCUSTOMDRAW), "NMLVCUSTOMDRAW", tesNMLVCUSTOMDRAW },
	{ sizeof(NMTVCUSTOMDRAW), "NMTVCUSTOMDRAW", tesNMTVCUSTOMDRAW },
	{ sizeof(NMHDR), "NMHDR", tesNMHDR },
	{ sizeof(NONCLIENTMETRICS), "NONCLIENTMETRICS", tesNONCLIENTMETRICS },
	{ sizeof(NOTIFYICONDATA), "NOTIFYICONDATA", tesNOTIFYICONDATA },
	{ sizeof(OSVERSIONINFO), "OSVERSIONINFO", tesOSVERSIONINFOEX },
	{ sizeof(OSVERSIONINFOEX), "OSVERSIONINFOEX", tesOSVERSIONINFOEX },
	{ sizeof(PAINTSTRUCT), "PAINTSTRUCT", tesPAINTSTRUCT },
	{ sizeof(POINT), "POINT", tesPOINT },
	{ sizeof(RECT), "RECT", tesRECT },
	{ sizeof(SHELLEXECUTEINFO), "SHELLEXECUTEINFO", tesSHELLEXECUTEINFO },
	{ sizeof(SHFILEINFO), "SHFILEINFO", tesSHFILEINFO },
	{ sizeof(SHFILEOPSTRUCT), "SHFILEOPSTRUCT", tesSHFILEOPSTRUCT },
	{ sizeof(SIZE), "SIZE", tesSIZE },
	{ sizeof(TCHITTESTINFO), "TCHITTESTINFO", tesTCHITTESTINFO },
	{ sizeof(TCITEM), "TCITEM", tesTCITEM },
	{ sizeof(TOOLINFO), "TOOLINFO", tesTOOLINFO },
	{ sizeof(TVHITTESTINFO), "TVHITTESTINFO", tesTVHITTESTINFO },
	{ sizeof(TVITEM), "TVITEM", tesTVITEM },
	{ sizeof(VARIANT), "VARIANT", tesNULL },
	{ sizeof(WCHAR), "WCHAR", tesNULL },
	{ sizeof(WIN32_FIND_DATA), "WIN32_FIND_DATA", tesWIN32_FIND_DATA },
	{ sizeof(WORD), "WORD", tesNULL },
};

TEmethod methodMem2[] = {
	{ VT_I4  << TE_VT, "int" },
	{ VT_UI4 << TE_VT, "DWORD" },
	{ VT_UI1 << TE_VT, "BYTE" },
	{ VT_UI2 << TE_VT, "WORD" },
	{ VT_UI2 << TE_VT, "WCHAR" },
	{ VT_PTR << TE_VT, "HANDLE" },
	{ VT_PTR << TE_VT, "LPWSTR" },
	{ 0, NULL }
};

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
#ifdef _USE_TESTOBJECT
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
	{ 0, NULL }
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
	{ 0, NULL }
};

TEmethod methodWB[] = {
	{ TE_PROPERTY + 1, "Data" },
	{ TE_PROPERTY + 2, "hwnd" },
	{ TE_PROPERTY + 3, "Type" },
	{ TE_METHOD + 4, "TranslateAccelerator" },
	{ TE_PROPERTY + 5, "Application" },
	{ TE_PROPERTY + 6, "Document" },
	{ TE_PROPERTY + 7, "Window" },
	{ TE_METHOD + 8, "Focus" },
	{ TE_METHOD + 9, "Close" },
	{ TE_METHOD + 10, "PreventClose" },
	{ TE_PROPERTY + 11, "Visible" },
	{ TE_PROPERTY + 12, "DropMode" },
	{ START_OnFunc + WB_OnClose, "OnClose" },
	{ 0, NULL }
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
	{ 0, NULL }
};

TEmethod methodFIs[] = {
	{ TE_PROPERTY + 2, "Application" },
	{ TE_PROPERTY + 3, "Parent" },
	{ TE_METHOD + 8, "AddItem" },
	{ TE_PROPERTY + 9, "hDrop" },
	{ TE_METHOD + 10, "GetData" },
	{ TE_METHOD + 11, "SetData" },
	{ DISPID_NEWENUM, "_NewEnum" },
	{ DISPID_TE_ITEM, "Item" },
	{ DISPID_TE_COUNT, "Count" },
	{ DISPID_TE_COUNT, "length" },
	{ DISPID_TE_INDEX, "Index" },
	{ TE_PROPERTY + 0x1001, "dwEffect" },
	{ TE_PROPERTY + 0x1002, "pdwEffect" },
	{ TE_PROPERTY + 0x1003, "Data" },
	{ TE_PROPERTY + 0x1004, "UseText" },
	{ 0, NULL }
};

TEmethod methodDT[] = {
	{ TE_METHOD + 1, "DragEnter" },
	{ TE_METHOD + 2, "DragOver" },
	{ TE_METHOD + 3, "Drop" },
	{ TE_METHOD + 4, "DragLeave" },
	{ TE_PROPERTY + 5, "Type" },
	{ TE_PROPERTY + 6, "FolderItem" },
	{ 0, NULL }
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
	{ 0, NULL }
};

TEmethod methodFI[] = {
	{ 1, "Name" },
	{ 2, "Path" },
	{ 3, "Alt" },
//	{ 4, "FocusedItem" },
	{ 5, "Unavailable" },
	{ 6, "Enum" },
	{ 9, "_BLOB" },	//To be necessary
	{ 9, "FolderItem" }, // Reserved future
	{ 0, NULL }
};

TEmethod methodMem[] = {
	{ TE_PROPERTY + 0xfff1, "P" },
	{ TE_METHOD + 4, "Read" },
	{ TE_METHOD + 5, "Write" },
	{ TE_PROPERTY + 0xfff6, "Size" },
	{ TE_METHOD + 7, "Free" },
	{ TE_METHOD + 8, "Clone" },
	{ TE_PROPERTY + 0xfff9, "_BLOB" },
	{ DISPID_NEWENUM, "_NewEnum" },
	{ DISPID_TE_ITEM, "Item" },
	{ DISPID_TE_COUNT, "Count" },
	{ DISPID_TE_COUNT, "length" },
	{ 0, NULL }
};

TEmethod methodCM[] = {
	{ TE_METHOD + 1, "QueryContextMenu" },
	{ TE_METHOD + 2, "InvokeCommand" },
	{ TE_METHOD + 3, "Items" },
	{ TE_METHOD + 4, "GetCommandString" },
	{ TE_PROPERTY + 5, "FolderView" },
	{ TE_METHOD + 6, "HandleMenuMsg" },
	{ TE_PROPERTY + 10, "hmenu" },
	{ TE_PROPERTY + 11, "indexMenu" },
	{ TE_PROPERTY + 12, "idCmdFirst" },
	{ TE_PROPERTY + 13, "idCmdLast" },
	{ TE_PROPERTY + 14, "uFlags" },
	{ 0, NULL }
};

TEmethod methodCD[] = {
	{ TE_METHOD + 40, "ShowOpen" },
	{ TE_METHOD + 41, "ShowSave" },
//	{ TE_METHOD + 42, "ShowFolder" },
	{ TE_PROPERTY + 10, "FileName" },
	{ TE_PROPERTY + 13, "Filter" },
	{ TE_PROPERTY + 20, "InitDir" },
	{ TE_PROPERTY + 21, "DefExt" },
	{ TE_PROPERTY + 22, "Title" },
	{ TE_PROPERTY + 30, "MaxFileSize" },
	{ TE_PROPERTY + 31, "Flags" },
	{ TE_PROPERTY + 32, "FilterIndex" },
	{ TE_PROPERTY + 31, "FlagsEx" },
	{ 0, NULL }
};

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
	{ 0, NULL }
};

TEmethod methodPD[] = {
	{ TE_METHOD + 1, "HasUserCancelled" },
	{ TE_METHOD + 2, "SetCancelMsg" },
	{ TE_METHOD + 3, "SetLine" },
	{ TE_METHOD + 4, "SetProgress" },
	{ TE_METHOD + 5, "SetTitle" },
	{ TE_METHOD + 6, "StartProgressDialog" },
	{ TE_METHOD + 7, "StopProgressDialog" },
	{ TE_METHOD + 8, "Timer" },
	{ TE_METHOD + 9, "SetAnimation" },
	{ 0, NULL }
};

TEmethod methodCO[] = {
	{ TE_METHOD + 1, "Free" },
	{ 0, NULL }
};

TEmethod methodObject[] = {
	{ 1, "Object" },
	{ 2, "Array" },
	{ 3, "Enum" },
	{ 4, "api" },
	{ 5, "CommonDialog" },
	{ 6, "FolderItems" },
	{ 7, "ProgressDialog" },
	{ 8, "WICBitmap" },
{ 0, NULL }
};

#ifdef _USE_TEOBJ
TEmethod methodArray[] = {
	{ DISPID_NEWENUM, "_NewEnum" },
	{ DISPID_TE_ITEM, "Item" },
	{ DISPID_TE_COUNT, "Count" },
	{ DISPID_TE_COUNT, "length" },
	{ TE_METHOD + 1, "push" },
	{ TE_METHOD + 2, "pop" },
	{ TE_METHOD + 3, "shift" },
	{ TE_METHOD + 4, "unshift" },
	{ TE_METHOD + 5, "join" },
	{ TE_METHOD + 6, "slice" },
	{ TE_METHOD + 7, "splice" },
	{ 0, NULL }
};
#endif

TEmethod methodEN[] = {
	{ TE_METHOD + 1, "atEnd" },
	{ TE_METHOD + 2, "item" },
	{ TE_METHOD + 3, "moveFirst" },
	{ TE_METHOD + 4, "moveNext" },
	{ 0, NULL }
};
