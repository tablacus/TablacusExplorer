const Addon_Id = "titlebar";
if (window.Addon == 1) {
	const item = GetAddonElement(Addon_Id);
	const show_fullpath = GetNum(item.getAttribute("ShowFullPath"));

	AddEvent("ChangeView1", async function (Ctrl) {
		if (show_fullpath) {
			const pid = await Ctrl.FolderItem;
			let fullpath;
			if (pid && (fullpath = await pid.Path) && fullpath.length > 0) {
				document.title = fullpath + ' - ' + TITLE;
			} else {
				document.title = await Ctrl.Title + ' - ' + TITLE;
			}
		} else {
			document.title = await Ctrl.Title + ' - ' + TITLE;
		}
	});
} else {
	SetTabContents(0, "General", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
}
