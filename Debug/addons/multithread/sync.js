const Addon_Id = "multithread";
const item = GetAddonElement(Addon_Id);

Sync.MultiThread = {
	Copy: GetNum(item.getAttribute("Copy")),
	Move: GetNum(item.getAttribute("Move")),
	Delete: GetNum(item.getAttribute("Delete")),
	NoTemp: item.getAttribute("NoTemp"),

	FO: function (Ctrl, Items, Dest, grfKeyState, pt, pdwEffect, bOver, bDelete) {
		let path;
		if (!(grfKeyState & MK_LBUTTON) || Items.Count == 0) {
			return false;
		}
		try {
			path = Dest.ExtendedProperty("linktarget") || Dest.Path || Dest;
		} catch (e) {
			path = Dest.Path || Dest;
		}
		const wfd = api.Memory("WIN32_FIND_DATA");
		if (bDelete || (path && fso.FolderExists(path))) {
			const arFrom = [];
			const strTemp = GetTempPath(4);
			if (Items.Count == 1 && api.PathMatchSpec(Items.Item(0).Path, strTemp + "*.bmp")) {
				return false;
			}
			let strTemp2;
			for (let i = 0; i < Items.Count; ++i) {
				let path1 = Items.Item(i).Path;
				const hFind = api.FindFirstFile(path1, wfd);
				if (hFind != INVALID_HANDLE_VALUE) {
					api.FindClose(hFind);
					if (!bDelete && !api.StrCmpNI(path1, strTemp, strTemp.length)) {
						if (!strTemp2) {
							if (Sync.MultiThread.NoTemp) {
								return false;
							}
							strTemp2 = GetTempPath(7);
							CreateFolder(strTemp2);
						}
						path1 = strTemp2 + GetFileName(path1);
					}
					arFrom.push(path1);
				} else {
					pdwEffect[0] = DROPEFFECT_NONE;
					break;
				}
			}
			if (pdwEffect[0]) {
				let wFunc = 0;
				if (bDelete) {
					if (!Sync.MultiThread.Delete) {
						return false;
					}
					wFunc = FO_DELETE;
				} else {
					if (bOver) {
						api.DropTarget(path).DragOver(Items, grfKeyState, pt, pdwEffect);
					}
					if (pdwEffect[0] & DROPEFFECT_COPY) {
						if (!Sync.MultiThread.Copy) {
							return false;
						}
						wFunc = FO_COPY;
					} else if (pdwEffect[0] & DROPEFFECT_MOVE) {
						if (!Sync.MultiThread.Move) {
							return false;
						}
						wFunc = FO_MOVE;
					}
				}
				if (wFunc) {
					let fFlags = FOF_ALLOWUNDO;
					if (bDelete) {
						if (api.GetKeyState(VK_SHIFT) < 0) {
							fFlags = 0;
						}
					} else if (api.StrCmpI(path, Items.Item(-1).Path) == 0) {
						if (wFunc == FO_MOVE) {
							return false;
						}
						fFlags |= FOF_RENAMEONCOLLISION;
					}
					if (strTemp2) {
						for (let i = 0; i < Items.Count; ++i) {
							const path1 = Items.Item(i).Path;
							const hFind = api.FindFirstFile(path1, wfd);
							if (hFind != INVALID_HANDLE_VALUE) {
								api.FindClose(hFind);
								if (!api.StrCmpNI(path1, strTemp, strTemp.length)) {
									if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
										fso.MoveFolder(path1, strTemp2);
									} else {
										fso.MoveFile(path1, strTemp2);
									}
								}
							}
						}
					}
					api.SHFileOperation(wFunc, arFrom, path, fFlags, true);
					return true;
				}
			}
		}
		return false;
	}
};

AddEvent("Drop", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	switch (Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
		case CTRL_TV:
			const Dest = GetDropTargetItem(Ctrl, Ctrl.hwndList, pt);
			if (Dest && Sync.MultiThread.FO(Ctrl, dataObj, Dest, grfKeyState, pt, pdwEffect, true)) {
				return S_OK
			}
			break;
		case CTRL_DT:
			if (Sync.MultiThread.FO(null, dataObj, Ctrl.FolderItem, grfKeyState, pt, pdwEffect, true)) {
				return S_OK
			}
			break;
	}
});

AddEvent("Command", function (Ctrl, hwnd, msg, wParam, lParam) {
	if (Ctrl.Type == CTRL_SB || Ctrl.Type == CTRL_EB) {
		switch ((wParam & 0xfff) + 1) {
			case CommandID_PASTE:
				let Items = api.OleGetClipboard()
				if (Sync.MultiThread.FO(null, Items, Ctrl.FolderItem, MK_LBUTTON, null, Items.pdwEffect, false)) {
					return S_OK;
				}
				break;
			case CommandID_DELETE:
				Items = Ctrl.SelectedItems();
				if (Sync.MultiThread.FO(null, Items, "", MK_LBUTTON, null, Items.pdwEffect, false, true)) {
					return S_OK;
				}
				break;
		}
	}
});

AddEvent("InvokeCommand", function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon) {
	switch (Verb + 1) {
		case CommandID_PASTE:
			const Target = ContextMenu.Items();
			if (Target.Count) {
				const Items = api.OleGetClipboard()
				if (Sync.MultiThread.FO(null, Items, Target.Item(0), MK_LBUTTON, null, Items.pdwEffect, false)) {
					return S_OK;
				}
			}
			break;
		case CommandID_DELETE:
			const Items = ContextMenu.Items();
			if (Sync.MultiThread.FO(null, Items, "", MK_LBUTTON, null, Items.pdwEffect, false, true)) {
				return S_OK;
			}
			break;
	}
});
