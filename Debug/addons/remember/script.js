var Addon_Id = "remember";

if (window.Addon == 1) {
	var item = GetAddonElement(Addon_Id);

	Addons.Remember =
	{
		db: {},
		ID: ["Time", "ViewMode", "IconSize", "Columns", "SortColumn", "Group", "SortColumns", "Path"],
		nFormat: api.QuadPart(GetAddonOption(Addon_Id, "Format")),
		Filter: (GetAddonOption(Addon_Id, "Filter") || "*").replace(/\s+$/, "").replace(/\r\n/g, ";"),
		Disable: (GetAddonOption(Addon_Id, "Disable") || "-").replace(/\s+$/, "").replace(/\r\n/g, ";"),
		nIcon: api.GetSystemMetrics(SM_CYICON) * 96 / screen.deviceYDPI,
		nSM: api.GetSystemMetrics(SM_CYSMICON) * 96 / screen.deviceYDPI,
		nSave: item.getAttribute("Save") || 1000,

		RememberFolder: function (FV) {
			if (FV && FV.FolderItem && !FV.FolderItem.Unavailable && FV.Data && FV.Data.Remember) {
				var path = Addons.Remember.GetPath(FV);
				if (path == FV.Data.Remember && PathMatchEx(path, Addons.Remember.Filter) && !PathMatchEx(path, Addons.Remember.Disable)) {
					var col = FV.Columns(Addons.Remember.nFormat);
					if (col) {
						Addons.Remember.db[path] = [new Date().getTime(), FV.CurrentViewMode, FV.IconSize, col, FV.SortColumn(Addons.Remember.nFormat), FV.GroupBy, FV.SortColumns];
					}
				}
			}
		},

		GetPath: function (pid) {
			var path = ("string" === typeof pid ? pid : String(api.GetDisplayNameOf(pid, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX))).toLowerCase();
			var res = /^([a-z][a-z\-_]+:).*/i.exec(path);
			return res ? res[1] : path;
		}
	};

	var xml = OpenXml("remember.xml", true, false);
	if (xml) {
		var items = xml.getElementsByTagName('Item');
		for (var i = items.length; i-- > 0;) {
			var item = items[i];
			var ar = [];
			for (var j = Addons.Remember.ID.length; j--;) {
				ar[j] = item.getAttribute(Addons.Remember.ID[j]);
			}
			if (ar[1]) {
				var path = Addons.Remember.GetPath(String(ar.pop()));
				if (PathMatchEx(path, Addons.Remember.Filter) && !PathMatchEx(path, Addons.Remember.Disable)) {
					Addons.Remember.db[path] = ar;
				}
			}
		}
		xml = null;
	}

	AddEvent("BeforeNavigate", function (Ctrl, fs, wFlags, Prev) {
		if (Ctrl.Data && !Ctrl.Data.Setting) {
			if (Prev && !Prev.Unavailable) {
				var path = Addons.Remember.GetPath(Prev);
				var col = Ctrl.Columns(Addons.Remember.nFormat);
				if (col && PathMatchEx(path, Addons.Remember.Filter) && !PathMatchEx(path, Addons.Remember.Disable)) {
					Addons.Remember.db[path] = [new Date().getTime(), Ctrl.CurrentViewMode, Ctrl.IconSize, col, Ctrl.SortColumn(Addons.Remember.nFormat), Ctrl.GroupBy, Ctrl.SortColumns];
				}
			}
			var path = Addons.Remember.GetPath(Ctrl);
			var ar = Addons.Remember.db[path];
			if (ar) {
				if (PathMatchEx(path, Addons.Remember.Filter) && !PathMatchEx(path, Addons.Remember.Disable)) {
					fs.ViewMode = ar[1];
					if (ar[2] > Addons.Remember.nIcon && (ar[1] > FVM_ICON && ar[1] <= FVM_DETAILS)) {
						ar[2] = Addons.Remember.nSM;
					}
					fs.ImageSize = ar[2];
					Ctrl.Data.Setting = 'Remember';
				} else {
					delete Addons.Remember.db[path];
				}
			}
			Ctrl.Data.Remember = "";
		}
	});

	AddEvent("NavigateComplete", function (Ctrl) {
		var path = Addons.Remember.GetPath(Ctrl);
		if (path && PathMatchEx(path, Addons.Remember.Filter) && !PathMatchEx(path, Addons.Remember.Disable)) {
			Ctrl.Data.Remember = path;
			if (Ctrl.Data && Ctrl.Data.Setting == 'Remember') {
				var ar = Addons.Remember.db[Ctrl.Data.Remember];
				if (ar) {
					if (ar[2] > Addons.Remember.nIcon && (ar[1] > FVM_ICON && ar[1] <= FVM_DETAILS)) {
						ar[2] = Addons.Remember.nSM;
					}
					Ctrl.CurrentViewMode(ar[1], ar[2]);
					Ctrl.Columns = ar[3];
					if (Ctrl.GroupBy && ar[5]) {
						Ctrl.GroupBy = ar[5];
					}
					if ((ar[6] || "").split(/;/).length > 2 && Ctrl.SortColumns) {
						Ctrl.SortColumns = ar[6];
					} else {
						Ctrl.SortColumn = ar[4];
					}
					ar[0] = new Date().getTime();
				}
			}
		}
	});

	AddEvent("ChangeView", Addons.Remember.RememberFolder);
	AddEvent("CloseView", Addons.Remember.RememberFolder);
	AddEvent("Command", Addons.Remember.RememberFolder);
	AddEvent("ViewModeChanged", Addons.Remember.RememberFolder);
	AddEvent("ColumnsChanged", Addons.Remember.RememberFolder);

	AddEvent("InvokeCommand", function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon) {
		var Ctrl = te.Ctrl(CTRL_FV);
		if (Ctrl) {
			Addons.Remember.RememberFolder(Ctrl);
		}
	});

	AddEvent("SaveConfig", function () {
		Addons.Remember.RememberFolder(te.Ctrl(CTRL_FV));

		var arFV = [];
		for (var path in Addons.Remember.db) {
			if (path && PathMatchEx(path, Addons.Remember.Filter) && !PathMatchEx(path, Addons.Remember.Disable)) {
				var ar = Addons.Remember.db[path];
				ar.push(path);
				arFV.push(ar);
			}
		}

		arFV.sort(
			function (a, b) {
				return b[0] - a[0];
			}
		);
		arFV.splice(Addons.Remember.nSave, arFV.length);
		var xml = CreateXml();
		var root = xml.createElement("TablacusExplorer");

		while (arFV.length) {
			var ar = arFV.shift();
			var item = xml.createElement("Item");
			for (var j in Addons.Remember.ID) {
				item.setAttribute(Addons.Remember.ID[j], ar[j] || "");
			}
			root.appendChild(item);
		}
		xml.appendChild(root);
		SaveXmlEx("remember.xml", xml, true);
	});
} else {
	var ado = OpenAdodbFromTextFile("addons\\" + Addon_Id + "\\options.html");
	if (ado) {
		SetTabContents(0, "General", ado.ReadText(adReadAll));
		ado.Close();
	}
}