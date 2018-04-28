var Addon_Id = "multithread";
var item = GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("Copy", 1);
	item.setAttribute("Move", 1);
	item.setAttribute("Delete", 1);
}

if (window.Addon == 1) {
	Addons.MultiThread =
	{
		Copy: api.LowPart(item.getAttribute("Copy")),
		Move: api.LowPart(item.getAttribute("Move")),
		Delete: api.LowPart(item.getAttribute("Delete")),

		FO: function (Ctrl, Items, Dest, grfKeyState, pt, pdwEffect, bOver, bDelete)
		{
			var path;
			if (!(grfKeyState & MK_LBUTTON) || Items.Count == 0) {
				return false;
			}
			var Parent = Items.Item(-1);
			if (!bDelete && api.ILIsParent(wsh.ExpandEnvironmentStrings("%TEMP%"), Parent, false)) {
				return false;
			}
			try {
				path = Dest.ExtendedProperty("linktarget") || Dest.Path || Dest;
			} catch (e) {
				path = Dest.Path || Dest;
			}
			if (bDelete || (path && fso.FolderExists(path))) {
				var arFrom = [];
				for (i = Items.Count - 1; i >= 0; i--) {
					var path1 = Items.Item(i).Path;
					if (IsExists(path1)) {
						arFrom.unshift(path1);
					} else {
						pdwEffect[0] = DROPEFFECT_NONE;
						break;
					}
				}
				if (pdwEffect[0]) {
					var wFunc = 0;
					if (bDelete) {
						if (!Addons.MultiThread.Delete) {
							return false;
						}
						wFunc = FO_DELETE;
					} else {
						if (bOver) {
							var DropTarget = api.DropTarget(path);
							DropTarget.DragOver(Items, grfKeyState, pt, pdwEffect);
						}
						if (pdwEffect[0] & DROPEFFECT_COPY) {
							if (!Addons.MultiThread.Copy) {
								return false;
							}
							wFunc = FO_COPY;
						} else if (pdwEffect[0] & DROPEFFECT_MOVE) {
							if (!Addons.MultiThread.Move) {
								return false;
							}
							wFunc = FO_MOVE;
						}
					}
					if (wFunc) {
						var fFlags = FOF_ALLOWUNDO;
						if (bDelete) {
							if (api.GetKeyState(VK_SHIFT) < 0) {
								fFlags = 0;
							}
						} else if (api.ILIsEqual(path, Parent)) {
							fFlags |= FOF_RENAMEONCOLLISION;
						}
						api.SHFileOperation(wFunc, arFrom.join("\0"), path, fFlags, true);
						return true;
					}
				}
			}
			return false;
		}
	};

	AddEvent("Drop", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect)
	{
		switch (Ctrl.Type) {
			case CTRL_SB:
			case CTRL_EB:
			case CTRL_TV:
				var Dest = Ctrl.HitTest(pt);
				if (Dest) {
					if (!fso.FolderExists(Dest.Path)) {
						if (api.DropTarget(Dest)) {
							return E_FAIL;
						}
						Dest = Ctrl.FolderItem;
					}
				} else {
					Dest = Ctrl.FolderItem;
				}
				if (Dest && Addons.MultiThread.FO(Ctrl, dataObj, Dest, grfKeyState, pt, pdwEffect, true)) {
					return S_OK
				}
				break;
			case CTRL_DT:
				if (Addons.MultiThread.FO(null, dataObj, Ctrl.FolderItem, grfKeyState, pt, pdwEffect, true)) {
					return S_OK
				}
				break;
		}
	});

	AddEvent("Command", function (Ctrl, hwnd, msg, wParam, lParam)
	{
		if (Ctrl.Type == CTRL_SB || Ctrl.Type == CTRL_EB) {
			switch ((wParam & 0xfff) + 1) {
				case CommandID_PASTE:
					var Items = api.OleGetClipboard()
					if (Addons.MultiThread.FO(null, Items, Ctrl.FolderItem, MK_LBUTTON, null, Items.pdwEffect, false)) {
						return S_OK;
					}
					break;
				case CommandID_DELETE:
					var Items = Ctrl.SelectedItems();
					if (Addons.MultiThread.FO(null, Items, "", MK_LBUTTON, null, Items.pdwEffect, false, true)) {
						return S_OK;
					}
					break;
			}
		}
	});

	AddEvent("InvokeCommand", function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon)
	{
		switch (Verb + 1) {
			case CommandID_PASTE:
				var Target = ContextMenu.Items();
				if (Target.Count) {
					var Items = api.OleGetClipboard()
					if (Addons.MultiThread.FO(null, Items, Target.Item(0), MK_LBUTTON, null, Items.pdwEffect, false)) {
						return S_OK;
					}
				}
				break;
			case CommandID_DELETE:
				var Items = ContextMenu.Items();
				if (Addons.MultiThread.FO(null, Items, "", MK_LBUTTON, null, Items.pdwEffect, false, true)) {
					return S_OK;
				}
				break;
		}
	});
} else {
	SetTabContents(0, "General", '<input type="checkbox" id="Copy" /><label for="Copy">Copy</label><br /><input type="checkbox" id="Move" /><label for="Move">Move</label><br /><input type="checkbox" id="Delete" /><label for="Delete">Delete</label><br />');
}
