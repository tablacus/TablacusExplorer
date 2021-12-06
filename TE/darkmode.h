#pragma once

VOID teSetDarkMode(HWND hwnd);
VOID teGetDarkMode();
VOID teSetTreeTheme(HWND hwnd, COLORREF cl);
BOOL teIsDarkColor(COLORREF cl);
VOID teFixGroup(LPNMLVCUSTOMDRAW lplvcd, COLORREF clrBk);
LRESULT CALLBACK CBTProc(int nCode, WPARAM wParam, LPARAM lParam);

