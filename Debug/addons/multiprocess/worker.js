importScripts("..\\..\\script\\consts.js");
ex = {};

if (MainWindow.Exchange) {
	te.OnSystemMessage = function (Ctrl, hwnd, msg, wParam, lParam) {
		try {
			if (msg == WM_TIMER && wParam == 1) {
				let bClose = ex.bClose;;
				let hwnd1 = null;
				const pid = api.Memory("DWORD");
				api.GetWindowThreadProcessId(hwnd, pid);
				const pid0 = pid[0];
				while (hwnd1 = api.FindWindowEx(null, hwnd1, null, null)) {
					const sClass = api.GetClassName(hwnd1);
					if (hwnd1 != hwnd && api.IsWindowVisible(hwnd1) && sClass != "Internet Explorer_Hidden") {
						api.GetWindowThreadProcessId(hwnd1, pid);
						if (pid[0] == pid0) {
							if (sClass == "#32770") {
								if (!(api.GetWindowLongPtr(hwnd1, GWL_EXSTYLE) & 8)) {
									if (ex) {
										if (!ex.NoFront) {
											api.SetForegroundWindow(hwnd1);
											api.SetWindowPos(hwnd1, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
										}
										if (!ex.bClose) {
											api.SetTimer(te.hwnd, 1, 99999, null);
										}
									}
								}
							}
							bClose = false;
							break;
						}
						//Dialog(UAC)
						if (api.GetWindowLongPtr(hwnd1, GWL_EXSTYLE) & 8) {
							if (sClass == "#32770") {
								if (ex) {
									ex.time = new Date().getTime();
								}
								bClose = false;
								break;
							}
						}
					}
				}
				if (bClose) {
					api.KillTimer(hwnd, wParam);
					try {
						if (ex && ex.TimeOver && new Date().getTime() - ex.time > ex.Sec * 1000 + 500) {
							ex.TimeOver = 0;
							ex.Callback(true);
						}
					} catch (e) { }
					delete ex.pids[pid0];
					api.PostMessage(hwnd, WM_CLOSE, 0, 0);
				}
			}
		} catch (e) {
			api.PostMessage(hwnd, WM_CLOSE, 0, 0);
		}
		return 0;
	}
	const pid = api.Memory("DWORD");
	api.GetWindowThreadProcessId(te.hwnd, pid);
	pid0 = pid[0];
	api.GetWindowThreadProcessId(api.GetForegroundWindow(), pid);

	ex.bClose = false;
	api.SetTimer(te.hwnd, 1, 999, null);
	const ex0 = MainWindow.Exchange[arg[3]];
	if (ex0) {
		for (let s in ex0) {
			ex[s] = ex0[s];
		}
		ex.State = [];
		for (let i = ex.State.length; i-- > 0;) {
			ex.State[i] = ex0.State[i];
		}
		delete MainWindow.Exchange[arg[3]];
		const KeyState0 = api.Memory("KEYSTATE");
		api.GetKeyboardState(KeyState0);
		if (ex.State.length) {
			const KeyState = api.Memory("KEYSTATE");
			api.GetKeyboardState(KeyState);
			for (let i = ex.State.length; i--;) {
				KeyState.Write(ex.State[i], VT_UI1, 0x80);
			}
			api.SetKeyboardState(KeyState);
		}
		ex.time = new Date().getTime();
		ex.pids[pid0] = 1;
		switch (ex.Mode) {
			case 0:
				api.DropTarget(ex.Dest).Drop(ex.Items, ex.grfKeyState, ex.pt, ex.dwEffect);
				break;
			case 1:
				InvokeCommandMP("delete", ex.Items);
				break;
			case 2:
				InvokeCommandMP("paste", ex.Dest);
				break;
		}
		if (ex.State.length) {
			api.SetKeyboardState(KeyState0);
		}
	}
	ex.bClose = true;
	api.SetTimer(te.hwnd, 1, 999, null);
	return "wait";
}

function InvokeCommandMP(sCommand, Items) {
	const ContextMenu = api.ContextMenu(Items);
	if (ContextMenu) {
		const hMenu = api.CreatePopupMenu();
		ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_DEFAULTONLY);
		ContextMenu.InvokeCommand(0, te.hwnd, sCommand, null, null, SW_SHOWNORMAL, null, null);
		api.DestroyMenu(hMenu);
	}
}
