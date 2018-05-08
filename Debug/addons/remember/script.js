var Addon_Id = "remember";

var item = GetAddonElement(Addon_Id);
if (!item.getAttribute("Save")) {
	item.setAttribute("Save", 1000);
}
if (window.Addon == 1) {
	Addons.Remember =
	{
		db: {},
		ID: ["Time", "ViewMode", "IconSize", "Columns", "SortColumn", "Group", "SortColumns", "Path"],
		nFormat: api.QuadPart(GetAddonOption(Addon_Id, "Format")),

		RememberFolder: function (FV)
		{
			if (FV && FV.FolderItem && !FV.FolderItem.Unavailable && FV.Data && FV.Data.Remember) {
				var path = Addons.Remember.GetPath(FV);
				if (path == FV.Data.Remember) {
					var col = FV.Columns(Addons.Remember.nFormat);
					if (col) {
						Addons.Remember.db[path] = [new Date().getTime(), FV.CurrentViewMode, FV.IconSize, col, FV.SortColumn(Addons.Remember.nFormat), FV.GroupBy, FV.SortColumns];
					}
				}
			}
		},

		GetPath: function (pid)
		{
			var path = (typeof(pid) == "string" ? pid : String(api.GetDisplayNameOf(pid, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX))).toLowerCase();
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
			for (var j in Addons.Remember.ID) {
				ar[j] = item.getAttribute(Addons.Remember.ID[j]);
			}
			if (ar[1]) {
				Addons.Remember.db[Addons.Remember.GetPath(String(ar.pop()))] = ar;
			}
		}
		xml = null;
	}

	AddEvent("BeforeNavigate", function (Ctrl, fs, wFlags, Prev)
	{
		if (Ctrl.Data && !Ctrl.Data.Setting) {
			if (Prev && !Prev.Unavailable) {
				var path = Addons.Remember.GetPath(Prev);
				var col = Ctrl.Columns(Addons.Remember.nFormat);
				if (col) {
					Addons.Remember.db[path] = [new Date().getTime(), Ctrl.CurrentViewMode, Ctrl.IconSize, col, Ctrl.SortColumn(Addons.Remember.nFormat), Ctrl.GroupBy, Ctrl.SortColumns];
				}
			}
			var ar = Addons.Remember.db[Addons.Remember.GetPath(Ctrl)];
			if (ar) {
				fs.ViewMode = ar[1];
				fs.ImageSize = ar[2];
				Ctrl.Data.Setting = 'Remember';
			}
			Ctrl.Data.Remember = "";
		}
	});

	AddEvent("NavigateComplete", function (Ctrl)
	{
		var path = Addons.Remember.GetPath(Ctrl);
		if (path) {
			Ctrl.Data.Remember = path;
			if (Ctrl.Data && Ctrl.Data.Setting == 'Remember') {
				var ar = Addons.Remember.db[Ctrl.Data.Remember];
				if (ar) {
					Ctrl.CurrentViewMode(ar[1], ar[2]);
					Ctrl.Columns = ar[3];
					var col = Ctrl.Columns(Addons.Remember.nFormat);
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

	AddEvent("InvokeCommand", function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon)
	{
		var Ctrl = te.Ctrl(CTRL_FV);
		if (Ctrl) {
			Addons.Remember.RememberFolder(Ctrl);
		}
	});

	AddEvent("SaveConfig", function ()
	{
		Addons.Remember.RememberFolder(te.Ctrl(CTRL_FV));

		var arFV = [];
		for (var path in Addons.Remember.db) {
			if (path) {
				var ar = Addons.Remember.db[path];
				ar.push(path);
				arFV.push(ar);
			}
		}

		arFV.sort(
			function(a, b) {
				return b[0] - a[0];
			}
		);
		var items = te.Data.Addons.getElementsByTagName("remember");
		if (items.length) {
			arFV.splice(items[0].getAttribute("Save"), arFV.length);
		}
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
	importScript("addons\\" + Addon_Id + "\\options.js");
}