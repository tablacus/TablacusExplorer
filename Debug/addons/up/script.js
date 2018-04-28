var Addon_Id = "up";
var Default = "ToolBar2Left";

var item = GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("Menu", "View");
	item.setAttribute("MenuPos", -1);
	item.setAttribute("MenuName", "&Up One Level");

	item.setAttribute("KeyOn", "List");
	item.setAttribute("Key", "$e");

	item.setAttribute("MouseOn", "List");
	item.setAttribute("Mouse", "2U");
}
if (window.Addon == 1) {
	Addons.Up =
	{
		nPos: 0,
		strName: "&Up One Level",

		Exec: function ()
		{
			Navigate(null, SBSP_PARENT | OpenMode);
			return S_OK;
		},

		Popup: function ()
		{
			var o = document.getElementById("UpButton");
			var FV = external.Ctrl(CTRL_FV);
			if (FV) {
				FolderMenu.Clear();
				var hMenu = api.CreatePopupMenu();
				var FolderItem = FV.FolderItem;
				if (api.ILIsEmpty(FolderItem)) {
					FolderItem = ssfDRIVES;
				}
				while (!api.ILIsEmpty(FolderItem)) {
					FolderItem = api.ILRemoveLastID(FolderItem);
					FolderMenu.AddMenuItem(hMenu, FolderItem);
				}
				var pt = api.Memory("POINT");
				api.GetCursorPos(pt);
				window.g_menu_click = true;
				var nVerb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, null);
				api.DestroyMenu(hMenu);
				if (nVerb) {
					FolderMenu.Invoke(FolderMenu.Items[nVerb - 1]);
				}
				FolderMenu.Clear();
			}
		}
	};
	//Menu
	if (item.getAttribute("MenuExec")) {
		Addons.Up.nPos = api.LowPart(item.getAttribute("MenuPos"));
		var s = item.getAttribute("MenuName");
		if (s && s != "") {
			Addons.Up.strName = s;
		}
		AddEvent(item.getAttribute("Menu"), function (Ctrl, hMenu, nPos)
		{
			api.InsertMenu(hMenu, Addons.Up.nPos, MF_BYPOSITION | MF_STRING, ++nPos, Addons.Up.strName);
			ExtraMenuCommand[nPos] = Addons.Up.Exec;
			return nPos;
		});
	}
	//Key
	if (item.getAttribute("KeyExec")) {
		SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Addons.Up.Exec, "Func", true);
	}
	//Mouse
	if (item.getAttribute("MouseExec")) {
		SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Addons.Up.Exec, "Func", true);
	}
	var h = item.getAttribute("IconSize") || window.IconSize || 24;
	var s = item.getAttribute("Icon") || (h <= 16 ? "bitmap:ieframe.dll,216,16,28" : "bitmap:ieframe.dll,214,24,28");
	SetAddon(Addon_Id, Default, '<span class="button" id="UpButton" onclick="Addons.Up.Exec();" oncontextmenu="Addons.Up.Popup(); return false;" onmouseover="MouseOver(this)" onmouseout="MouseOut()"><img title="Up" src="' + EncodeSC(s) + '" width="' + h + 'px" height="' + h + 'px"></span>');
}
