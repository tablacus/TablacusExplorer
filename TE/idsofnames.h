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
	{ 1132, "GdiplusBitmap" },
	{ 1137, "ProgressDialog" },
	{ 1138, "DateTimeFormat" },
	{ 1139, "HiddenFilter" },
//	{ 1140, "Background" },,//Deprecated
//	{ 1150, "ThumbnailProvider" },//Deprecated
	{ 1160, "DragIcon" },
	{ 1180, "ExplorerBrowserFilter" },
	{ TE_METHOD + 1133, "FolderItems" },
	{ TE_METHOD + 1134, "Object" },
	{ TE_METHOD + 1135, "Array" },
	{ TE_METHOD + 1136, "Collection" },
	{ TE_METHOD + 1050, "CreateCtrl" },
	{ TE_METHOD + 1040, "CtrlFromPoint" },
	{ TE_METHOD + 1060, "MainMenu" },
	{ TE_METHOD + 1070, "CtrlFromWindow" },
	{ TE_METHOD + 1080, "LockUpdate" },
	{ TE_METHOD + 1090, "UnlockUpdate" },
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
	{ 0, NULL }
};

TEmethod methodSB[] = {
	{ 0x10000001, "Data" },
	{ 0x10000002, "hwnd" },
	{ 0x10000003, "Type" },
	{ 0x10000004, "Navigate" },
	{ 0x10000007, "Navigate2" },
	{ 0x10000008, "Index" },
	{ 0x10000009, "FolderFlags" },
	{ 0x1000000B, "History" },
	{ 0x10000010, "CurrentViewMode" },
	{ 0x10000011, "IconSize" },
	{ 0x10000012, "Options" },
	{ 0x10000013, "SizeFormat" },
	{ 0x10000014, "NameFormat" }, //Deprecated
	{ 0x10000016, "ViewFlags" },
	{ 0x10000017, "Id" },
	{ 0x10000018, "FilterView" },
	{ 0x10000020, "FolderItem" },
	{ 0x10000021, "TreeView" },
	{ 0x10000024, "Parent" },
	{ 0x10000031, "Close" },
	{ 0x10000032, "Title" },
	{ 0x10000033, "Suspend" },
	{ 0x10000040, "Items" },
	{ 0x10000041, "SelectedItems" },
	{ 0x10000050, "ShellFolderView" },
	{ 0x10000058, "Droptarget" },
	{ 0x10000059, "Columns" },
//	{ 0x1000005A, "Searches" },
	{ 0x1000005B, "MapColumnToSCID" },
	{ 0x10000102, "hwndList" },
	{ 0x10000103, "hwndView" },
	{ 0x10000104, "SortColumn" },
	{ 0x10000105, "GroupBy" },
	{ 0x10000106, "Focus" },
	{ 0x10000107, "HitTest" },
	{ 0x10000108, "hwndAlt" },
	{ 0x10000110, "ItemCount" },
	{ 0x10000111, "Item" },
	{ 0x10000206, "Refresh" },
	{ 0x10000207, "ViewMenu" },
	{ 0x10000208, "TranslateAccelerator" },
	{ 0x10000209, "GetItemPosition" },
	{ 0x1000020A, "SelectAndPositionItem" },
	{ 0x10000280, "SelectItem" },
	{ 0x10000281, "FocusedItem" },
	{ 0x10000282, "GetFocusedItem" },
	{ 0x10000283, "GetItemRect" },
	{ 0x10000300, "Notify" },
	{ 0x10000400, "NavigateComplete" },
	{ 0x10000501, "AddItem" },
	{ 0x10000502, "RemoveItem" },
	{ 0x10000503, "AddItems" },
	{ 0x10000504, "RemoveAll" },
	{ 0x10000505, "SessionId" },
	{ START_OnFunc + SB_TotalFileSize, "TotalFileSize" },
	{ START_OnFunc + SB_ColumnsReplace, "ColumnsReplace" },
	{ START_OnFunc + SB_OnIncludeObject, "OnIncludeObject" },
	{ START_OnFunc + SB_AltSelectedItems, "AltSelectedItems" },
	{ START_OnFunc + SB_VirtualName, "VirtualName" }, //Deprecated
	{ 0, NULL }
};

TEmethod methodWB[] = {
	{ 0x10000001, "Data" },
	{ 0x10000002, "hwnd" },
	{ 0x10000003, "Type" },
	{ 0x10000004, "TranslateAccelerator" },
	{ 0x10000005, "Application" },
	{ 0x10000006, "Document" },
	{ 0x10000007, "Window" },
	{ 0x10000008, "Focus" },
//	{ 0x10000009, "Close" },
	{ 0, NULL }
};

TEmethod methodTC[] = {
	{ 1, "Data" },
	{ 2, "hwnd" },
	{ 3, "Type" },
	{ 6, "HitTest" },
	{ 7, "Move" },
	{ 8, "Selected" },
	{ 9, "Close" },
	{ 10, "SelectedIndex" },
	{ 11, "Visible" },
	{ 12, "Id" },
	{ 13, "LockUpdate" },
	{ 14, "UnlockUpdate" },
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
	{ 2, "Application" },
	{ 3, "Parent" },
	{ 8, "AddItem" },
	{ 9, "hDrop" },
	{ 10, "GetData" },
	{ 11, "SetData" },
	{ DISPID_NEWENUM, "_NewEnum" },
	{ DISPID_TE_ITEM, "Item" },
	{ DISPID_TE_COUNT, "Count" },
	{ DISPID_TE_COUNT, "length" },
	{ DISPID_TE_INDEX, "Index" },
//	{ 0x10000001, "lEvent" },
	{ 0x10000001, "dwEffect" },
	{ 0x10000002, "pdwEffect" },
	{ 0x10000003, "Data" },
	{ 0x10000004, "UseText" },
	{ 0, NULL }
};

TEmethod methodDT[] = {
	{ 1, "DragEnter" },
	{ 2, "DragOver" },
	{ 3, "Drop" },
	{ 4, "DragLeave" },
	{ 5, "Type" },
	{ 6, "FolderItem" },
	{ 0, NULL }
};

TEmethod methodTV[] = {
	{ 0x10000001, "Data" },
	{ 0x10000002, "Type" },
	{ 0x10000003, "hwnd" },
	{ 0x10000004, "Close" },
	{ 0x10000005, "hwndTree" },
	{ 0x10000007, "FolderView" },
	{ 0x10000008, "Align" },
	{ 0x10000009, "Visible" },
	{ 0x10000106, "Focus" },
	{ 0x10000107, "HitTest" },
	{ 0x10000206, "Refresh" },
	{ 0x10000283, "GetItemRect" },
	{ 0x10000300, "Notify" },
	{ TE_OFFSET + SB_TreeWidth, "Width" },
	{ TE_OFFSET + SB_TreeFlags, "Style" },
	{ TE_OFFSET + SB_EnumFlags, "EnumFlags" },
	{ TE_OFFSET + SB_RootStyle, "RootStyle" },
	{ 0x20000000, "SelectedItem" },
	{ 0x20000001, "SelectedItems" },
	{ 0x20000002, "Root" },
	{ 0x20000003, "SetRoot" },
	{ 0x20000004, "Expand" },
	{ 0x20000005, "Columns" },
	{ 0x20000006, "CountViewTypes" },
	{ 0x20000007, "Depth" },
	{ 0x20000008, "EnumOptions" },
	{ 0x20000009, "Export" },
	{ 0x2000000a, "Flags" },
	{ 0x2000000b, "Import" },
	{ 0x2000000c, "Mode" },
	{ 0x2000000d, "ResetSort" },
	{ 0x2000000e, "SetViewType" },
	{ 0x2000000f, "Synchronize" },
	{ 0x20000010, "TVFlags" },
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
	{ 9, "FolderItem" },
	{ 0, NULL }
};

TEmethod methodMem[] = {
	{ 1, "P" },
	{ 4, "Read" },
	{ 5, "Write" },
	{ 6, "Size" },
	{ 7, "Free" },
	{ 8, "Clone" },
	{ 9, "_BLOB" },
	{ DISPID_NEWENUM, "_NewEnum" },
	{ DISPID_TE_ITEM, "Item" },
	{ DISPID_TE_COUNT, "Count" },
	{ DISPID_TE_COUNT, "length" },
	{ 0, NULL }
};

TEmethod methodCM[] = {
	{ 1, "QueryContextMenu" },
	{ 2, "InvokeCommand" },
	{ 3, "Items" },
	{ 4, "GetCommandString" },
	{ 5, "FolderView" },
	{ 6, "HandleMenuMsg" },
	{ 10, "hmenu" },
	{ 11, "indexMenu" },
	{ 12, "idCmdFirst" },
	{ 13, "idCmdLast" },
	{ 14, "uFlags" },
	{ 0, NULL }
};

TEmethod methodCD[] = {
	{ 40, "ShowOpen" },
	{ 41, "ShowSave" },
//	{ 42, "ShowFolder" },
	{ 10, "FileName" },
	{ 13, "Filter" },
	{ 20, "InitDir" },
	{ 21, "DefExt" },
	{ 22, "Title" },
	{ 30, "MaxFileSize" },
	{ 31, "Flags" },
	{ 32, "FilterIndex" },
	{ 31, "FlagsEx" },
	{ 0, NULL }
};

TEmethod methodGB[] = {
	{ 1, "FromHBITMAP" },
	{ 2, "FromHICON" },
	{ 3, "FromResource" },
	{ 4, "FromFile" },
	{ 5, "FromStream" },
	{ 6, "FromArchive" },
	{ 7, "FromItem" },//Deprecated
	{ 8, "FromClipboard" },
	{ 9, "FromSource" },
	{ 90, "Create" },
	{ 99, "Free" },

	{ 100, "Save" },
	{ 101, "Base64" },
	{ 102, "DataURI" },
	{ 103, "GetStream" },
	{ 110, "GetWidth" },
	{ 111, "GetHeight" },
	{ 112, "GetPixel" },
	{ 113, "SetPixel" },
	{ 114, "GetPixelFormat" },
	{ 115, "FillRect" },
	{ 120, "GetThumbnailImage" },
	{ 130, "RotateFlip" },
	{ 140, "GetFrameCount" },
	{ 150, "Frame" },
	{ 160, "GetMetadata" },
	{ 161, "GetFrameMetadata" },

	{ 210, "GetHBITMAP" },
	{ 211, "GetHICON" },
	{ 212, "DrawEx" },

	{ 900, "GetCodecInfo" },
	{ START_OnFunc + WIC_OnGetAlt, "OnGetAlt" },
	{ 0, NULL }
};

TEmethod methodPD[] = {
	{ 0x60010001, "HasUserCancelled" },
	{ 0x60010002, "SetCancelMsg" },
	{ 0x60010003, "SetLine" },
	{ 0x60010004, "SetProgress" },
	{ 0x60010005, "SetTitle" },
	{ 0x60010006, "StartProgressDialog" },
	{ 0x60010007, "StopProgressDialog" },
	{ 0x60010008, "Timer" },
	{ 0x60010009, "SetAnimation" },
	{ 0, NULL }
};

TEmethod methodCO[] = {
	{ 0x60010001, "Free" },
	{ 0, NULL }
};
