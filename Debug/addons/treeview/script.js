var Addon_Id = "treeview";
var Default = "ToolBar2Left";

var item = GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("MenuPos", -1);
	item.setAttribute("List", 1);
}

if (window.Addon == 1) {
	Addons.TreeView =
	{
		strName: "Tree",
		nPos: 0,
		WM: TWM_APP++,
		Depth: api.LowPart(item.getAttribute("Depth")),
		tid: {},

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
						Addons.TreeView.Expand(TV.FolderView);
					}
				}
			}
			return S_OK;
		},

		Expand: function (Ctrl) {
			if (Ctrl.FolderItem && !IsSearchPath(Ctrl.FolderItem)) {
				var TV = Ctrl.TreeView;
				if (TV) {
					if (Addons.TreeView.tid[TV.Id]) {
						clearTimeout(Addons.TreeView.tid[TV.Id]);
						delete Addons.TreeView.tid[TV.Id];
					}
					TV.Expand(Ctrl.FolderItem, Addons.TreeView.Depth);
					Addons.TreeView.tid[TV.Id] = setTimeout(function () {
						delete Addons.TreeView.tid[TV.Id];
						TV.Expand(Ctrl.FolderItem, 0);
					}, 99);
				}
			}
		},

		Popup: function (o) {
			var FV = GetFolderView(o);
			if (FV) {
				FV.Focus();
				var TV = FV.TreeView;
				if (TV) {
					var n = InputDialog(GetText("Width"), TV.Width);
					if (n) {
						TV.Width = n;
						TV.Align = true;
					}
				}
			}
			return false;
		}
	};

	if (item.getAttribute("List")) {
		AddEvent("ChangeView", Addons.TreeView.Expand);
	}

	//Menu
	if (item.getAttribute("MenuExec")) {
		Addons.TreeView.nPos = api.LowPart(item.getAttribute("MenuPos"));
		Addons.TreeView.strName = item.getAttribute("MenuName") || Addons.TreeView.strName;
		AddEvent(item.getAttribute("Menu"), function (Ctrl, hMenu, nPos) {
			api.InsertMenu(hMenu, Addons.TreeView.nPos, MF_BYPOSITION | MF_STRING, ++nPos, GetText(Addons.TreeView.strName));
			ExtraMenuCommand[nPos] = Addons.TreeView.Exec;
			return nPos;
		});
	}
	//Key
	if (item.getAttribute("KeyExec")) {
		SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Addons.TreeView.Exec, "Func");
	}
	//Mouse
	if (item.getAttribute("MouseExec")) {
		SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Addons.TreeView.Exec, "Func");
	}
	var h = GetIconSize(item.getAttribute("IconSize"), item.getAttribute("Location") == "Inner" && 16);
	var src = item.getAttribute("Icon") || (h <= 16 ? "bitmap:ieframe.dll,216,16,43" : "bitmap:ieframe.dll,214,24,43");
	var s = ['<span class="button" onclick="Addons.TreeView.Exec(this)" oncontextmenu="return Addons.TreeView.Popup(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', GetImgTag({ title: "Tree", src: src }, h), '</span>'];
	SetAddon(Addon_Id, Default, s);

	SetGestureExec("Tree", "1", function (Ctrl, pt) {
		var Item = Ctrl.SelectedItem;
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

	AddTypeEx("Add-ons", "Tree", Addons.TreeView.Exec);

	if (WINVER >= 0x600) {
		AddEvent("AppMessage", function (Ctrl, hwnd, msg, wParam, lParam) {
			if (msg == Addons.TreeView.WM) {
				var pidls = {};
				var hLock = api.SHChangeNotification_Lock(wParam, lParam, pidls);
				if (hLock) {
					api.SHChangeNotification_Unlock(hLock);
					var cTV = te.Ctrls(CTRL_TV);
					for (var i in cTV) {
						cTV[i].Notify(pidls.lEvent, pidls[0], pidls[1], wParam, lParam);
					}
				}
				return S_OK;
			}
		});

		AddEvent("Finalize", function () {
			api.SHChangeNotifyDeregister(Addons.TreeView.uRegisterId);
		});

		Addons.TreeView.uRegisterId = api.SHChangeNotifyRegister(te.hwnd, SHCNRF_InterruptLevel | SHCNRF_NewDelivery, SHCNE_MKDIR | SHCNE_MEDIAINSERTED | SHCNE_DRIVEADD | SHCNE_NETSHARE | SHCNE_DRIVEREMOVED | SHCNE_MEDIAREMOVED | SHCNE_NETUNSHARE | SHCNE_RENAMEFOLDER | SHCNE_RMDIR | SHCNE_SERVERDISCONNECT | SHCNE_UPDATEDIR, Addons.TreeView.WM, ssfDESKTOP, true);
	}
} else {
	EnableInner();
	SetTabContents(0, "General", '<input type="checkbox" id="Depth" value="1"><label for="Depth">Expanded</label><br><input type="checkbox" id="List" value="1"><label for="List">List</label>');
}
