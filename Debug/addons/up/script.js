const Addon_Id = "up";
const Default = "ToolBar2Left";
let item = await GetAddonElement(Addon_Id);
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
			const FV = await GetFolderView(Ctrl, pt);
			FV.Focus();
			Exec(FV, "Up", "Tabs", 0, pt);
		},

		Popup: async function (el) {
			const FV = await GetFolderView(el);
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
				const pt = GetPos(el, 9);
				const nVerb = await FolderMenu.TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y);
				api.DestroyMenu(hMenu);
				if (nVerb) {
					await FolderMenu.Invoke(await FolderMenu.Items[nVerb - 1]);
				}
				FolderMenu.Clear();
			}
			return false;
		}
	};
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

	if (item.getAttribute("Location") == "Inner") {
		AddEvent("ChangeView2", async function (Ctrl) {
			DisableImage(document.getElementById("ImgUp_" + await Ctrl.Parent.Id), await api.ILIsEmpty(Ctrl));
		});
	} else {
		AddEvent("ChangeView1", async function (Ctrl) {
			DisableImage(document.getElementById("ImgUp_$"), await api.ILIsEmpty(Ctrl));
		});
	}

	AddEvent("Layout", async function () {
		SetAddon(Addon_Id, Default, ['<span class="button" id="UpButton" onclick="Addons.Up.Exec(this)" oncontextmenu="return Addons.Up.Popup(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({
			id: "ImgUp_$",
			title: strName,
			src: item.getAttribute("Icon") || "icon:general,28"
		}, GetIconSizeEx(item)), '</span>']);
		delete item;
	});
} else {
	EnableInner();
}
