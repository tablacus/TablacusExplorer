const Addon_Id = "treeview";
const Default = "ToolBar2Left";
const item = await GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("MenuPos", -1);
	item.setAttribute("List", 1);
}

if (window.Addon == 1) {
	Addons.TreeView = {
		Popup: async function (o) {
			const FV = await GetFolderViewEx(o);
			if (FV) {
				FV.Focus();
				const TV = await FV.TreeView;
				if (TV) {
					InputDialog(await GetText("Width"), await TV.Width, function (n) {
						if (n) {
							TV.Width = n;
							TV.Align = true;
						}
					});
				}
			}
			return false;
		}
	};

	const h = GetIconSize(item.getAttribute("IconSize"), item.getAttribute("Location") == "Inner" && 16);
	const src = item.getAttribute("Icon") || (h <= 16 ? "bitmap:ieframe.dll,216,16,43" : "bitmap:ieframe.dll,214,24,43");
	SetAddon(Addon_Id, Default, ['<span class="button" onclick="SyncExec(Sync.TreeView.Exec, this)" oncontextmenu="return Addons.TreeView.Popup(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({ title: "Tree", src: src }, h), '</span>']);
	$.importScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	EnableInner();
	SetTabContents(0, "General", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
}
