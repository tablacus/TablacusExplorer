var Addon_Id = "treeview";
var Default = "ToolBar2Left";

var item = await GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("MenuPos", -1);
	item.setAttribute("List", 1);
}

if (window.Addon == 1) {
	Addons.TreeView = {
		strName: "Tree",
		List: item.getAttribute("List"),
		nPos: 0,
		WM: TWM_APP++,
		Depth: GetNum(item.getAttribute("Depth")),
		tid: {},

		Exec: async function (o) {
			Sync.TreeView.Exec(await GetFolderViewEx(o));
			return S_OK;
		},

		Popup: async function (o) {
			var FV = await GetFolderViewEx(o);
			if (FV) {
				FV.Focus();
				var TV = await FV.TreeView;
				if (TV) {
					var n = await InputDialog(await GetText("Width"), await TV.Width);
					if (n) {
						TV.Width = n;
						TV.Align = true;
					}
				}
			}
			return false;
		}
	};

	var h = GetIconSize(item.getAttribute("IconSize"), item.getAttribute("Location") == "Inner" && 16);
	var src = item.getAttribute("Icon") || (h <= 16 ? "bitmap:ieframe.dll,216,16,43" : "bitmap:ieframe.dll,214,24,43");
	var s = ['<span class="button" onclick="Addons.TreeView.Exec(this)" oncontextmenu="return Addons.TreeView.Popup(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({ title: "Tree", src: src }, h), '</span>'];
	SetAddon(Addon_Id, Default, s);
	importJScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	EnableInner();
	SetTabContents(0, "General", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
}
