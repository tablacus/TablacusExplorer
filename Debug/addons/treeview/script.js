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
		List: item.getAttribute("List"),
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
			if (Ctrl.FolderItem && Addons.TreeView.List) {
				var TV = Ctrl.TreeView;
				if (TV) {
					TV.Expand(Ctrl.FolderItem, Addons.TreeView.Depth);
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
		},

		Refresh: function (Ctrl, pt) {
			var FV = GetFolderView(Ctrl, pt);
			FV.TreeView.Refresh();
			Addons.TreeView.Expand(FV);
		}
	};

	if (api.LowPart(item.getAttribute("Refresh"))) {
		AddEvent("Refresh", Addons.TreeView.Refresh);
	} else {
		SetKeyExec("Tree", "$3f,Ctrl+R", function (Ctrl, pt) {
			Addons.TreeView.Refresh(Ctrl, pt);
			return S_OK;
		}, "Func", true);
	}

	if (Addons.TreeView.List) {
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

	AddTypeEx("Add-ons", "Tree", Addons.TreeView.Exec);

	if (WINVER >= 0x600) {
		AddEvent("AppMessage", function (Ctrl, hwnd, msg, wParam, lParam) {
			if (msg == Addons.TreeView.WM) {
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
			api.SHChangeNotifyDeregister(Addons.TreeView.uRegisterId);
		});

		Addons.TreeView.uRegisterId = api.SHChangeNotifyRegister(te.hwnd, SHCNRF_InterruptLevel | SHCNRF_NewDelivery, SHCNE_MKDIR | SHCNE_MEDIAINSERTED | SHCNE_DRIVEADD | SHCNE_NETSHARE | SHCNE_DRIVEREMOVED | SHCNE_MEDIAREMOVED | SHCNE_NETUNSHARE | SHCNE_RENAMEFOLDER | SHCNE_RMDIR | SHCNE_SERVERDISCONNECT | SHCNE_UPDATEDIR, Addons.TreeView.WM, ssfDESKTOP, true);
	}
} else {
	EnableInner();
	var ado = OpenAdodbFromTextFile("addons\\" + Addon_Id + "\\options.html");
	if (ado) {
		SetTabContents(0, "General", ado.ReadText(adReadAll));
		ado.Close();
	}
}
