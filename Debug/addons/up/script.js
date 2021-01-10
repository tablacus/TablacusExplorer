const Addon_Id = "up";
const Default = "ToolBar2Left";
const item = await GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("Menu", "View");
	item.setAttribute("MenuPos", -1);

	item.setAttribute("KeyOn", "List");
	item.setAttribute("Key", "$e");

	item.setAttribute("MouseOn", "List");
	item.setAttribute("Mouse", "2U");
}
if (window.Addon == 1) {
	Addons.Up = {
		Exec: async function (Ctrl, pt) {
			const FV = await GetFolderViewEx(Ctrl, pt);
			FV.Focus();
			Exec(FV, "Up", "Tabs", 0, pt);
		},

		Popup: async function (ev, el) {
			const FV = await GetFolderViewEx(el);
			if (FV) {
				FV.Focus();
				await FolderMenu.Clear();
				const hMenu = await api.CreatePopupMenu();
				let FolderItem = await FV.FolderItem;
				if (await api.ILIsEmpty(FolderItem)) {
					FolderItem = ssfDRIVES;
				}
				while (!await api.ILIsEmpty(FolderItem)) {
					FolderItem = await api.ILRemoveLastID(FolderItem);
					FolderMenu.AddMenuItem(hMenu, FolderItem);
				}
				const x = ev.screenX * ui_.Zoom, y = ev.screenY * ui_.Zoom;
				const nVerb = await FolderMenu.TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, x, y);
				api.DestroyMenu(hMenu);
				if (nVerb) {
					await FolderMenu.Invoke(await FolderMenu.Items[nVerb - 1]);
				}
				FolderMenu.Clear();
			}
			return false;
		}
	};
	AddEvent("ChangeView", async function (Ctrl) {
		if (await Ctrl.Id == await Ctrl.Parent.Selected.Id) {
			DisableImage(document.getElementById("ImgUp"), await api.ILIsEmpty(Ctrl));
		}
	});
	//Menu
	const strName = item.getAttribute("MenuName") || await GetText("&Up One Level");
	if (item.getAttribute("MenuExec")) {
		SetMenuExec("Up", strName, item.getAttribute("Menu"), item.getAttribute("MenuPos"));
	}
	//Key
	if (item.getAttribute("KeyExec")) {
		SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Addons.Up.Exec, "Async", true);
	}
	//Mouse
	if (item.getAttribute("MouseExec")) {
		SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Addons.Up.Exec, "Async", true);
	}
	const h = GetIconSize(item.getAttribute("IconSize"), item.getAttribute("Location") == "Inner" && 16);
	const s = item.getAttribute("Icon") || (h <= 16 ? "bitmap:ieframe.dll,216,16,28" : "bitmap:ieframe.dll,214,24,28");
	SetAddon(Addon_Id, Default, ['<span class="button" id="UpButton" onclick="Addons.Up.Exec(this)" oncontextmenu="return Addons.Up.Popup(event, this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({ id: "ImgUp", title: strName, src: s }, h), '</span>']);
} else {
	EnableInner();
}
