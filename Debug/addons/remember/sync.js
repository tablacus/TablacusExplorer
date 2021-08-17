var Addon_Id = "remember";
var item = GetAddonElement(Addon_Id);

Common.Remember = api.CreateObject("Object");
Common.Remember.db = api.CreateObject("Object");

Sync.Remember = {
	ID: ["Time", "ViewMode", "IconSize", "Columns", "SortColumn", "Group", "SortColumns", "Path"],
	nFormat: api.LowPart(GetAddonOption(Addon_Id, "Format")),
	Filter: ExtractFilter(GetAddonOption(Addon_Id, "Filter") || "*"),
	Disable: ExtractFilter(GetAddonOption(Addon_Id, "Disable") || "-"),
	nIcon: api.GetSystemMetrics(SM_CYICON) * 96 / screen.deviceYDPI,
	nSM: api.GetSystemMetrics(SM_CYSMICON) * 96 / screen.deviceYDPI,
	nSave: item.getAttribute("Save") || 1000,

	RememberFolder: function (FV) {
		if (FV && FV.FolderItem && !FV.FolderItem.Unavailable && FV.Data && FV.Data.Remember) {
			var path = Sync.Remember.GetPath(FV);
			if (path == FV.Data.Remember && PathMatchEx(path, Sync.Remember.Filter) && !PathMatchEx(path, Sync.Remember.Disable)) {
				var col = FV.Columns(Sync.Remember.nFormat);
				if (col) {
					var ar = api.CreateObject("Array");
					ar.push(new Date().getTime(), FV.CurrentViewMode, FV.IconSize, col, FV.GetSortColumn(Sync.Remember.nFormat), FV.GroupBy, FV.SortColumns);
					Common.Remember.db[path] = ar;
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
		var ar = api.CreateObject("Array");
		for (var j = Sync.Remember.ID.length; j--;) {
			ar[j] = item.getAttribute(Sync.Remember.ID[j]);
		}
		if (ar[1]) {
			var path = Sync.Remember.GetPath(String(ar.pop()));
			if (PathMatchEx(path, Sync.Remember.Filter) && !PathMatchEx(path, Sync.Remember.Disable)) {
				Common.Remember.db[path] = ar;
			}
		}
	}
	xml = null;
}

AddEvent("BeforeNavigate", function (Ctrl, fs, wFlags, Prev) {
	if (Ctrl.Data && !Ctrl.Data.Setting) {
		if (Prev && !Prev.Unavailable) {
			var path = Sync.Remember.GetPath(Prev);
			var col = Ctrl.Columns(Sync.Remember.nFormat);
			if (col && PathMatchEx(path, Sync.Remember.Filter) && !PathMatchEx(path, Sync.Remember.Disable)) {
				var ar = api.CreateObject("Array");
				ar.push(new Date().getTime(), Ctrl.CurrentViewMode, Ctrl.IconSize, col, Ctrl.GetSortColumn(Sync.Remember.nFormat), Ctrl.GroupBy, Ctrl.SortColumns);
				Common.Remember.db[path] = ar;
			}
		}
		var path = Sync.Remember.GetPath(Ctrl);
		var ar = Common.Remember.db[path];
		if (ar) {
			if (PathMatchEx(path, Sync.Remember.Filter) && !PathMatchEx(path, Sync.Remember.Disable)) {
				fs.ViewMode = ar[1];
				if (ar[2] > Sync.Remember.nIcon && (ar[1] > FVM_ICON && ar[1] <= FVM_DETAILS)) {
					ar[2] = Sync.Remember.nSM;
				}
				fs.ImageSize = ar[2];
				Ctrl.Data.Setting = 'Remember';
			} else {
				delete Common.Remember.db[path];
			}
		}
		Ctrl.Data.Remember = "";
	}
});

AddEvent("NavigateComplete", function (Ctrl) {
	var path = Sync.Remember.GetPath(Ctrl);
	if (path && PathMatchEx(path, Sync.Remember.Filter) && !PathMatchEx(path, Sync.Remember.Disable)) {
		Ctrl.Data.Remember = path;
		if (Ctrl.Data && Ctrl.Data.Setting == 'Remember') {
			var ar = Common.Remember.db[Ctrl.Data.Remember];
			if (ar) {
				if (ar[2] > Sync.Remember.nIcon && (ar[1] > FVM_ICON && ar[1] <= FVM_DETAILS)) {
					ar[2] = Sync.Remember.nSM;
				}
				Ctrl.SetViewMode(ar[1], ar[2]);
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

AddEvent("ChangeView", Sync.Remember.RememberFolder);
AddEvent("CloseView", Sync.Remember.RememberFolder);
AddEvent("Command", Sync.Remember.RememberFolder);
AddEvent("ViewModeChanged", Sync.Remember.RememberFolder);
AddEvent("ColumnsChanged", Sync.Remember.RememberFolder);

AddEvent("InvokeCommand", function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon) {
	var Ctrl = te.Ctrl(CTRL_FV);
	if (Ctrl) {
		Sync.Remember.RememberFolder(Ctrl);
	}
});

AddEvent("SaveConfig", function () {
	Sync.Remember.RememberFolder(te.Ctrl(CTRL_FV));

	var arFV = [];
	for (var path in Common.Remember.db) {
		if (path && PathMatchEx(path, Sync.Remember.Filter) && !PathMatchEx(path, Sync.Remember.Disable)) {
			var ar = Common.Remember.db[path];
			if (ar) {
				ar.push(path);
				arFV.push(ar);
			}
		}
	}

	arFV.sort(
		function (a, b) {
			return b[0] - a[0];
		}
	);
	arFV.splice(Sync.Remember.nSave, arFV.length);
	var xml = CreateXml();
	var root = xml.createElement("TablacusExplorer");

	while (arFV.length) {
		var ar = arFV.shift();
		var item = xml.createElement("Item");
		for (var j in Sync.Remember.ID) {
			item.setAttribute(Sync.Remember.ID[j], ar[j] || "");
		}
		root.appendChild(item);
	}
	xml.appendChild(root);
	SaveXmlEx("remember.xml", xml, true);
});
