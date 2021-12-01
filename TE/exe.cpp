#include "stdafx.h"
#include "common.h"
#if !defined(_WINDLL) && !defined(_EXEONLY)

int APIENTRY _tWinMain(HINSTANCE hInstance,
	HINSTANCE hPrevInstance,
	LPTSTR    lpCmdLine,
	int       nCmdShow)
{
	__security_init_cookie();
	HINSTANCE hDll = ::GetModuleHandleA("kernel32.dll");
	LPFNSetDefaultDllDirectories _SetDefaultDllDirectories = (LPFNSetDefaultDllDirectories)::GetProcAddress(hDll, "SetDefaultDllDirectories");
	if (_SetDefaultDllDirectories) {
		_SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_SYSTEM32);
	}
	lpCmdLine = ::GetCommandLine();
	if (lpCmdLine = ::StrChr(&lpCmdLine[1], lpCmdLine[0] == '"' ? '"' : 0x20)) {
		do {
			++lpCmdLine;
		} while (*lpCmdLine == 0x20);
	}
	WCHAR pszPath[MAX_PATHEX];
	::GetModuleFileName(NULL, pszPath, MAX_PATHEX);
	::PathRemoveFileSpec(pszPath);
#ifdef _WIN64
#ifdef _DEBUG
	::PathAppend(pszPath, L"lib\\TEd64.dll");
#else
	::PathAppend(pszPath, L"lib\\TE64.dll");
#endif
#else
#ifdef _DEBUG
	::PathAppend(pszPath, L"lib\\TEd32.dll");
#else
	::PathAppend(pszPath, L"lib\\TE32.dll");
#endif
#endif
	LPWSTR pszError = L"404 File Not Found";
	hDll = ::LoadLibrary(pszPath);
	if (hDll) {
		pszError = L"501 Not Implemented";
		LPFNEntryPointW _RunDLLW = (LPFNEntryPointW)::GetProcAddress(hDll, "RunDLLW");
		if (_RunDLLW) {
			STARTUPINFO si;
			::GetStartupInfo(&si);
			_RunDLLW(NULL, hDll, lpCmdLine, (si.dwFlags & STARTF_USESHOWWINDOW) ? si.wShowWindow : SW_SHOWDEFAULT);
			pszError = NULL;
		}
		::FreeLibrary(hDll);
	}
	if (pszError) {
		hDll = LoadLibrary(L"user32.dll");
		LPFNMessageBoxW _MessageBoxW = (LPFNMessageBoxW)::GetProcAddress(hDll, "MessageBoxW");
		if (_MessageBoxW) {
			_MessageBoxW(NULL, pszPath, pszError, MB_OK | MB_ICONERROR);
		}
		::FreeLibrary(hDll);
		::ExitProcess(-1);
	}
	::ExitProcess(0);
}
#endif
