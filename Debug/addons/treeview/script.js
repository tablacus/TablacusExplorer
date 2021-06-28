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
			const FV = await GetFolderView(o);
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

	AddEvent("Layout", async function () {
		const h = GetIconSizeEx(item);
		SetAddon(Addon_Id, Default, ['<span class="button" onclick="SyncExec(Sync.TreeView.Exec, this)" oncontextmenu="return Addons.TreeView.Popup(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({
			title: await GetText("Tree"),
			src: item.getAttribute("Icon") || "bitmap:ieframe.dll,214,24,43"
		}, h), '</span>']);
	});
	$.importScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	EnableInner();
	SetTabContents(0, "General", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
}
