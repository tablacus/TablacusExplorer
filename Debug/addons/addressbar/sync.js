Sync.AddressBar = {
	GetPath: function (n) {
		let FolderItem = 0;
		const FV = te.Ctrl(CTRL_FV);
		if (FV) {
			for (FolderItem = FV.FolderItem; n > 0; n--) {
				FolderItem = api.ILGetParent(FolderItem);
			}
		}
		return FolderItem;
	},

	SplitPath: function (FolderItem) {
		const Items = [];
		let n = 0;
		do {
			Items.push({
				next: n || api.GetAttributesOf(FolderItem, SFGAO_HASSUBFOLDER),
				name: GetFolderItemName(FolderItem)
			});
			FolderItem = api.ILGetParent(FolderItem);
			n++;
		} while (!api.ILIsEmpty(FolderItem) && n < 99);
		return JSON.stringify(Items);
	}
};

AddEvent("MouseMessage", function (Ctrl, hwnd, msg, mouseData, pt, wHitTestCode, dwExtraInfo) {
	if (msg == WM_MOUSEMOVE && Ctrl.Type == CTRL_TE && Common.AddressBar.rcItem) {
		const Ctrl2 = te.CtrlFromPoint(pt);
		if (Ctrl2 && Ctrl2.Type == CTRL_WB) {
			const ptc = pt.Clone();
			api.ScreenToClient(WebBrowser.hwnd, ptc);
			for (let i = Common.AddressBar.rcItem.length; i-- > 0;) {
				if (PtInRect(Common.AddressBar.rcItem[i], ptc)) {
					api.PostMessage(hwnd, WM_KEYDOWN, VK_ESCAPE, 0);
					InvokeUI("Addons.AddressBar.ChangeMenu", i);
					break;
				}
			}
		}
	}
});

AddEvent("DragEnter", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	if (Ctrl.Type == CTRL_WB) {
		InvokeUI("Addons.AddressBar.SetRects");
		return S_OK;
	}
});

AddEvent("DragOver", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	if (Ctrl.Type == CTRL_WB && dataObj.Count && Common.AddressBar.rcDrop) {
		const ptc = pt.Clone();
		api.ScreenToClient(WebBrowser.hwnd, ptc);
		for (let i = Common.AddressBar.rcDrop.length; i-- > 0;) {
			if (PtInRect(Common.AddressBar.rcDrop[i], ptc)) {
				const Target = Sync.AddressBar.GetPath(i);
				if (!api.ILIsEqual(dataObj.Item(-1), Target)) {
					const DropTarget = api.DropTarget(Target);
					if (DropTarget) {
						MouseOver("addressbar" + i + "_");
						return DropTarget.DragOver(dataObj, grfKeyState, pt, pdwEffect);
					}
				}
				pdwEffect[0] = DROPEFFECT_NONE;
				return S_OK;
			}
		}
	}
});

AddEvent("Drop", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	if (Ctrl.Type == CTRL_WB && dataObj.Count && Common.AddressBar.rcDrop) {
		const ptc = pt.Clone();
		api.ScreenToClient(WebBrowser.hwnd, ptc);
		for (let i = Common.AddressBar.rcDrop.length; i-- > 0;) {
			if (PtInRect(Common.AddressBar.rcDrop[i], ptc)) {
				let hr = S_FALSE;
				const Target = Sync.AddressBar.GetPath(i);
				if (!api.ILIsEqual(dataObj.Item(-1), Target)) {
					const DropTarget = api.DropTarget(Target);
					if (DropTarget) {
						hr = DropTarget.Drop(dataObj, grfKeyState, pt, pdwEffect);
					}
				}
				return hr;
			}
		}
	}
});

AddEvent("DragLeave", function (Ctrl) {
	MouseOut("addressbar");
	return S_OK;
});
