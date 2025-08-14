const Addon_Id = "multiprocess";
let item = GetAddonElement(Addon_Id);

Sync.MultiProcess = {
	NoDelete: item.getAttribute("NoDelete"),
	NoPaste: item.getAttribute("NoPaste"),
	NoDrop: item.getAttribute("NoDrop"),
	NoRDrop: item.getAttribute("NoRDrop"),
	NoTemp: item.getAttribute("NoTemp"),
	NoFront: item.getAttribute("NoFront"),
	TimeOver: item.getAttribute("TimeOver"),
	Sec: item.getAttribute("Sec"),

	pids: {},

	FO: function (Ctrl, Items, Dest, grfKeyState, pt, pdwEffect, nMode) {
		if (Items.Count == 0) {
			return false;
		}
		const dwEffect = pdwEffect[0];
		if (Dest !== null) {
			const path = api.GetDisplayNameOf(Dest, SHGDN_FORPARSING);
			if (/^::{/.test(path) || (/^[A-Z]:\\|^\\/i.test(path) && !api.PathIsDirectory(path))) {
				return false;
			}
			if (nMode == 0) {
				if (!(grfKeyState & (MK_CONTROL | MK_RBUTTON)) && api.ILIsEqual(Dest, Items.Item(-1))) {
					return false;
				}
				const DropTarget = api.DropTarget(Dest);
				DropTarget.DragOver(Items, grfKeyState, pt, pdwEffect);
			}
		}
		if (nMode == 0) {
			if (grfKeyState & MK_RBUTTON) {
				if (Sync.MultiProcess.NoRDrop) {
					return false;
				}
			} else {
				if (Sync.MultiProcess.NoDrop) {
					return false;
				}
			}
		} else if (nMode == 1) {
			if (Sync.MultiProcess.NoDelete) {
				return false;
			}
		} else if (nMode == 2) {
			if (Sync.MultiProcess.NoPaste) {
				return false;
			}
		}
		if (Dest !== null) {
			const strTemp = GetTempPath(4);
			if (Items.Count == 1 && api.PathMatchSpec(Items.Item(0).Path, strTemp + "*.bmp")) {
				return false;
			}
			let strTemp2;
			const Items2 = api.CreateObject("FolderItems");
			for (let i = 0; i < Items.Count; ++i) {
				let path1 = Items.Item(i).Path;
				if (api.PathFileExists(path1)) {
					if (!api.StrCmpNI(path1, strTemp, strTemp.length)) {
						if (!strTemp2) {
							if (Sync.MultiProcess.NoTemp) {
								return false;
							}
							strTemp2 = GetTempPath(7);
							CreateFolder(strTemp2);
						}
						const path2 = strTemp2 + GetFileName(path1);
						if (!SameText(path1, path2)) {
							if (api.MoveFileEx(path1, path2, 0)) {
								path1 = path2;
							}
						}
					}
					Items2.AddItem(path1);
				} else {
					delete strTemp2;
					break;
				}
			}
			if (strTemp2) {
				Items = Items2;
			}
		}
		if (!Sync.MultiProcess.NoFront) {
			setTimeout(function () {
				let hwnd = null;
				while (hwnd = api.FindWindowEx(null, hwnd, null, null)) {
					if (api.GetClassName(hwnd) == "OperationStatusWindow") {
						if (!(api.GetWindowLongPtr(hwnd, GWL_EXSTYLE) & 8)) {
							api.SetForegroundWindow(hwnd);
							api.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
						}
					}
				}
			}, 5000);
		}

		const State = [VK_SHIFT, VK_CONTROL, VK_MENU];
		for (let i = State.length; i--;) {
			if (api.GetKeyState(State[i]) >= 0) {
				State.splice(i, 1);
			}
		}
		OpenNewProcess("addons\\multiprocess\\worker.js", {
			Items: Items,
			Dest: Dest,
			Mode: nMode,
			grfKeyState: grfKeyState,
			pt: pt,
			dwEffect: dwEffect,
			TimeOver: Sync.MultiProcess.TimeOver,
			Sec: Sync.MultiProcess.Sec,
			NoFront: Sync.MultiProcess.NoFront,
			State: State,
			pids: Sync.MultiProcess.pids,
			Callback: Sync.MultiProcess.Player
		});
		return true;
	},

	Player: function (autoplay) {
		InvokeUI("Addons.MultiProcess.Player", autoplay);
	}
};
delete item;

AddEvent("Drop", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	switch (Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
		case CTRL_TV:
			const Dest = GetDropTargetItem(Ctrl, Ctrl.hwndList, pt);
			if (Dest && Sync.MultiProcess.FO(Ctrl, dataObj, Dest, grfKeyState, pt, pdwEffect, 0)) {
				return S_OK
			}
			break;
		case CTRL_DT:
			if (Sync.MultiProcess.FO(null, dataObj, Ctrl.FolderItem, grfKeyState, pt, pdwEffect, 0)) {
				return S_OK
			}
			break;
	}
});

AddEvent("Command", function (Ctrl, hwnd, msg, wParam, lParam) {
	if (Ctrl.Type == CTRL_SB || Ctrl.Type == CTRL_EB) {
		switch ((wParam & 0xfff) + 1) {
			case CommandID_PASTE:
				let Items = api.OleGetClipboard();
				if (Items && !api.ILIsEmpty(Items.Item(-1)) && Sync.MultiProcess.FO(null, Items, Ctrl.FolderItem, MK_LBUTTON, null, Items.pdwEffect, 2)) {
					return S_OK;
				}
				break;
			case CommandID_DELETE:
				Items = Ctrl.SelectedItems();
				if (Sync.MultiProcess.FO(null, Items, null, MK_LBUTTON, null, Items.pdwEffect, 1)) {
					return S_OK;
				}
				break;
		}
	}
});

AddEvent("InvokeCommand", function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon) {
	const sVerb = ContextMenu.GetCommandString(Verb, GCS_VERB);
	if (SameText(sVerb, "paste")) {
		const Target = ContextMenu.Items();
		if (Target.Count) {
			const Items = api.OleGetClipboard()
			if (Sync.MultiProcess.FO(null, Items, Target.Item(0), MK_LBUTTON, null, Items.pdwEffect, 2)) {
				return S_OK;
			}
		}
	} else if (SameText(sVerb, "delete")) {
		const Items = ContextMenu.Items();
		if (Sync.MultiProcess.FO(null, Items, null, MK_LBUTTON, null, Items.pdwEffect, 1)) {
			return S_OK;
		}
	}
});

AddEvent("CanClose", function (Ctrl) {
	if (Ctrl.Type == CTRL_TE) {
		let hwnd1 = null;
		const pid = api.Memory("DWORD");
		while (hwnd1 = api.FindWindowEx(null, hwnd1, null, null)) {
			if (!/tablacus|_hidden/i.test(api.GetClassName(hwnd1)) && api.IsWindowVisible(hwnd1)) {
				api.GetWindowThreadProcessId(hwnd1, pid);
				for (let pid0 in Sync.MultiProcess.pids) {
					if (pid[0] == pid0) {
						return S_FALSE;
					}
				}
			}
		}
	}
});
