const Addon_Id = "treeview";
const item = GetAddonElement(Addon_Id);

Sync.TreeView = {
	strName: "Tree",
	List: item.getAttribute("List"),
	nPos: 0,
	WM: TWM_APP++,
	Depth: GetNum(item.getAttribute("Depth")),
	Collapse: GetNum(item.getAttribute("Collapse")),

	Exec: function (Ctrl, pt) {
		const FV = GetFolderView(Ctrl, pt);
		if (FV) {
			FV.Focus();
			const TV = FV.TreeView;
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
		if (Sync.TreeView.List && Ctrl.FolderItem && IsWitness(Ctrl.FolderItem)) {
			const TV = Ctrl.TreeView;
			if (TV) {
				if (Sync.TreeView.Collapse) {
					const hwnd = TV.hwndTree;
					let hItem = api.SendMessage(hwnd, TVM_GETNEXTITEM, 9, null);
					let Now = TV.SelectedItem;
					let New = Ctrl.FolderItem;
					let nUp = Sync.TreeView.Depth ? 0 : 1;
					while (api.ILGetCount(New) > api.ILGetCount(Now)) {
						New = api.ILRemoveLastID(New);
					}
					while (api.ILGetCount(Now) > api.ILGetCount(New)) {
						Now = api.ILRemoveLastID(Now);
						++nUp;
					}
					while (!api.ILIsEqual(Now, New) && api.ILGetCount(Now) > 1) {
						New = api.ILRemoveLastID(New);
						Now = api.ILRemoveLastID(Now);
						++nUp;
					}
					while (--nUp > 0) {
						hItem = api.SendMessage(hwnd, TVM_GETNEXTITEM, 3, hItem);
					}
					do {
						api.PostMessage(hwnd, TVM_EXPAND, 0x8001, hItem);
					} while (hItem = api.SendMessage(hwnd, TVM_GETNEXTITEM, 1, hItem));
				}
				TV.Expand(Ctrl.FolderItem, Sync.TreeView.Depth);
			}
		}
	},

	Refresh: function (Ctrl, pt) {
		const FV = GetFolderView(Ctrl, pt);
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

if (Sync.TreeView.List) {
	SetGestureExec("Tree", "1", function (Ctrl, pt) {
		const Item = Ctrl.HitTest(pt);
		if (Item) {
			const FV = Ctrl.FolderView;
			if (!api.ILIsEqual(FV.FolderItem, Item) && Item.IsFolder) {
				setTimeout(function () {
					FV.Navigate(Item, GetNavigateFlags(FV));
				}, 99);
			}
		}
		return S_OK;
	}, "Func", true);
}

SetGestureExec("Tree", "11", function (Ctrl, pt) {
	const Item = Ctrl.HitTest(pt);
	if (Item) {
		const FV = Ctrl.FolderView;
		if (!api.ILIsEqual(FV.FolderItem, Item)) {
			setTimeout(function () {
				FolderMenu.Invoke(Item, void 0, FV);
			}, 99);
		}
	}
	return S_OK;
}, "Func", true);

SetGestureExec("Tree", "3", function (Ctrl, pt) {
	const Item = Ctrl.SelectedItem;
	if (Item) {
		setTimeout(function () {
			Ctrl.FolderView.Navigate(Item, SBSP_NEWBROWSER);
		}, 99);
	}
	return S_OK;
}, "Func", true);

//Tab
SetKeyExec("Tree", "$f", function (Ctrl, pt) {
	const FV = GetFolderView(Ctrl, pt);
	FV.focus();
	return S_OK;
}, "Func", true);

//Enter
SetKeyExec("Tree", "$1c", function (Ctrl, pt) {
	const FV = GetFolderView(Ctrl, pt);
	FV.Navigate(Ctrl.SelectedItem, GetNavigateFlags(FV));
	return S_OK;
}, "Func", true);

AddTypeEx("Add-ons", "Tree", Sync.TreeView.Exec);

if (WINVER >= 0x600) {
	AddEvent("AppMessage", function (Ctrl, hwnd, msg, wParam, lParam) {
		if (msg == Sync.TreeView.WM) {
			const pidls = {};
			const hLock = api.SHChangeNotification_Lock(wParam, lParam, pidls);
			if (hLock) {
				api.SHChangeNotification_Unlock(hLock);
				if (pidls[0] && /^[A-Z]:\\|^\\\\\w/i.test(pidls[0].Path) && IsWitness(pidls[0])) {
					const cFV = te.Ctrls(CTRL_FV);
					for (let i in cFV) {
						cFV[i].TreeView.Notify(pidls.lEvent, pidls[0], pidls[1], wParam, lParam);
					}
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
