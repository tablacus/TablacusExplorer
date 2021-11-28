#pragma once

VOID teSetDarkMode(HWND hwnd);
VOID teGetDarkMode();
VOID teSetTreeTheme(HWND hwnd, COLORREF cl);
BOOL teIsDarkColor(COLORREF cl);
LRESULT CALLBACK CBTProc(int nCode, WPARAM wParam, LPARAM lParam);

