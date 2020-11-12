var Addon_Id = "treeview";
var item = GetAddonElement(Addon_Id);

Sync.TreeView = {
	strName: "Tree",
	List: item.getAttribute("List"),
	nPos: 0,
	WM: TWM_APP++,
	Depth: GetNum(item.getAttribute("Depth")),

	Exec: function (Ctrl, pt) {
		var FV = GetFolderView(Ctrl, pt);
		if (FV) {
			FV.Focus();
			var TV = FV.TreeView;
			if (TV) {
				TV.Visible = !TV.Visible;
				if (TV.Visible) {
					if (!TV.Width) {
						TV.Width = 200;
					}
					Sync.TreeView.Expand(FV);
				}
			}
		}
		return S_OK;
	},

	Expand: function (Ctrl) {
		if (Sync.TreeView.List && Ctrl.FolderItem) {
			var TV = Ctrl.TreeView;
			if (TV) {
				TV.Expand(Ctrl.FolderItem, Sync.TreeView.Depth);
			}
		}
	},

	Refresh: function (Ctrl, pt) {
		var FV = GetFolderView(Ctrl, pt);
		FV.TreeView.Refresh();
		Sync.TreeView.Expand(FV);
	}
};

if (GetNum(item.getAttribute("Refresh"))) {
	AddEvent("Refresh", Sync.TreeView.Refresh);
} else {
	SetKeyExec("Tree", "$3f,Ctrl+R", function (Ctrl, pt) {
		Sync.TreeView.Refresh(Ctrl, pt);
		return S_OK;
	}, "Func", true);
}

if (Sync.TreeView.List) {
	AddEvent("ChangeView", Sync.TreeView.Expand);
}

//Menu
if (item.getAttribute("MenuExec")) {
	Sync.TreeView.nPos = api.LowPart(item.getAttribute("MenuPos"));
	Sync.TreeView.strName = item.getAttribute("MenuName") || Sync.TreeView.strName;
	AddEvent(item.getAttribute("Menu"), function (Ctrl, hMenu, nPos) {
		api.InsertMenu(hMenu, Sync.TreeView.nPos, MF_BYPOSITION | MF_STRING, ++nPos, GetText(Sync.TreeView.strName));
		ExtraMenuCommand[nPos] = Sync.TreeView.Exec;
		return nPos;
	});
}
//Key
if (item.getAttribute("KeyExec")) {
	SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Sync.TreeView.Exec, "Func");
}
//Mouse
if (item.getAttribute("MouseExec")) {
	SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Sync.TreeView.Exec, "Func");
}

SetGestureExec("Tree", "1", function (Ctrl, pt) {
	var Item = Ctrl.HitTest(pt);
	if (Item) {
		var FV = Ctrl.FolderView;
		if (!api.ILIsEqual(FV.FolderItem, Item)) {
			setTimeout(function () {
				FV.Navigate(Item, GetNavigateFlags(FV));
			}, 99);
		}
	}
	return S_OK;
}, "Func", true);

SetGestureExec("Tree", "3", function (Ctrl, pt) {
	var Item = Ctrl.SelectedItem;
	if (Item) {
		setTimeout(function () {
			Ctrl.FolderView.Navigate(Item, SBSP_NEWBROWSER);
		}, 99);
	}
	return S_OK;
}, "Func", true);

//Tab
SetKeyExec("Tree", "$f", function (Ctrl, pt) {
	var FV = GetFolderView(Ctrl, pt);
	FV.focus();
	return S_OK;
}, "Func", true);

//Enter
SetKeyExec("Tree", "$1c", function (Ctrl, pt) {
	var FV = GetFolderView(Ctrl, pt);
	FV.Navigate(Ctrl.SelectedItem, GetNavigateFlags(FV));
	return S_OK;
}, "Func", true);

AddTypeEx("Add-ons", "Tree", Sync.TreeView.Exec);

if (WINVER >= 0x600) {
	AddEvent("AppMessage", function (Ctrl, hwnd, msg, wParam, lParam) {
		if (msg == Sync.TreeView.WM) {
			var pidls = {};
			var hLock = api.SHChangeNotification_Lock(wParam, lParam, pidls);
			if (hLock) {
				api.SHChangeNotification_Unlock(hLock);
				var cFV = te.Ctrls(CTRL_FV);
				for (var i in cFV) {
					cFV[i].TreeView.Notify(pidls.lEvent, pidls[0], pidls[1], wParam, lParam);
				}
			}
			return S_OK;
		}
	});

	AddEvent("Finalize", function () {
		api.SHChangeNotifyDeregister(Sync.TreeView.uRegisterId);
	});

	Sync.TreeView.uRegisterId = api.SHChangeNotifyRegister(te.hwnd, SHCNRF_InterruptLevel | SHCNRF_NewDelivery, SHCNE_MKDIR | SHCNE_MEDIAINSERTED | SHCNE_DRIVEADD | SHCNE_NETSHARE | SHCNE_DRIVEREMOVED | SHCNE_MEDIAREMOVED | SHCNE_NETUNSHARE | SHCNE_RENAMEFOLDER | SHCNE_RMDIR | SHCNE_SERVERDISCONNECT | SHCNE_UPDATEDIR, Sync.TreeView.WM, ssfDESKTOP, true);
}
