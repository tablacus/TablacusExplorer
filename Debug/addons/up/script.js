var Addon_Id = "up";
var Default = "ToolBar2Left";

var item = await GetAddonElement(Addon_Id);
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
		nPos: GetNum(item.getAttribute("MenuPos")),
		strName: item.getAttribute("MenuName") || GetText("&Up One Level"),

		Exec: async function (Ctrl, pt) {
			var FV = await GetFolderViewEx(Ctrl, pt);
			FV.Focus();
			Exec(FV, "Up", "Tabs", 0, pt);
			return S_OK;
		},

		Popup: async function (ev, el) {
			var FV = await GetFolderViewEx(el);
			if (FV) {
				FV.Focus();
				await $.FolderMenu.Clear();
				var hMenu = await api.CreatePopupMenu();
				var FolderItem = await FV.FolderItem;
				if (await api.ILIsEmpty(FolderItem)) {
					FolderItem = ssfDRIVES;
				}
				while (!await api.ILIsEmpty(FolderItem)) {
					FolderItem = await api.ILRemoveLastID(FolderItem);
					$.FolderMenu.AddMenuItem(hMenu, FolderItem);
				}
				var x = ev.screenX * ui_.Zoom;
				var y = ev.screenY * ui_.Zoom;
				var nVerb = await $.FolderMenu.TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, x, y);
				api.DestroyMenu(hMenu);
				if (nVerb) {
					await $.FolderMenu.Invoke(await $.FolderMenu.Items[nVerb - 1]);
				}
				$.FolderMenu.Clear();
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
	if (item.getAttribute("MenuExec")) {
		Common.Up = await api.CreateObject("Object");
		Common.Up.strMenu = item.getAttribute("Menu");
		Common.Up.strName = item.getAttribute("MenuName") || await GetAddonInfo(Addon_Id).Name;
		Common.Up.nPos = GetNum(item.getAttribute("MenuPos"));
		importJScript("addons\\" + Addon_Id + "\\sync.js");
	}
	//Key
	if (item.getAttribute("KeyExec")) {
		SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Addons.Up.Exec, "Async", true);
	}
	//Mouse
	if (item.getAttribute("MouseExec")) {
		SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Addons.Up.Exec, "Async", true);
	}
	var h = await GetIconSize(item.getAttribute("IconSize"), item.getAttribute("Location") == "Inner" && 16);
	var s = item.getAttribute("Icon") || (h <= 16 ? "bitmap:ieframe.dll,216,16,28" : "bitmap:ieframe.dll,214,24,28");
	SetAddon(Addon_Id, Default, ['<span class="button" id="UpButton" onclick="Addons.Up.Exec(this)" oncontextmenu="return Addons.Up.Popup(event, this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({ id: "ImgUp", title: Addons.Up.strName, src: s }, h), '</span>']);
} else {
	EnableInner();
}
