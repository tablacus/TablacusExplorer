var Addon_Id = "back";
var Default = "ToolBar2Left";

if (window.Addon == 1) {
	var item = GetAddonElement(Addon_Id);

	Addons.Back =
	{
		Exec: function (Ctrl, pt) {
			Exec(Ctrl, "Back", "Tabs", 0, pt);
		},

		Popup: function (o) {
			var FV = te.Ctrl(CTRL_FV);
			if (FV) {
				var Log = FV.History;
				var hMenu = api.CreatePopupMenu();
				var mii = api.Memory("MENUITEMINFO");
				mii.cbSize = mii.Size;
				mii.fMask = MIIM_ID | MIIM_STRING | MIIM_BITMAP;
				for (var i = Log.Index + 1; i < Log.Count; i++) {
					var FolderItem = Log.Item(i);
					AddMenuIconFolderItem(mii, FolderItem);
					mii.dwTypeData = api.GetDisplayNameOf(FolderItem, SHGDN_INFOLDER | SHGDN_ORIGINAL);
					mii.wID = i;
					api.InsertMenuItem(hMenu, MAXINT, false, mii);
				}
				var pt = api.Memory("POINT");
				api.GetCursorPos(pt);
				var nVerb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, null);
				api.DestroyMenu(hMenu);
				if (nVerb) {
					var wFlags = GetNavigateFlags(FV);
					if (wFlags & SBSP_NEWBROWSER) {
						FV.Navigate(Log[nVerb], wFlags);
					} else {
						Log.Index = nVerb;
						FV.History = Log;
					}
				}
			}
			return false;
		}
	};

	AddEvent("ChangeView", function (Ctrl) {
		if (Ctrl.Id == Ctrl.Parent.Selected.Id) {
			var Log = Ctrl.History;
			DisableImage(document.getElementById("ImgBack"), Log && Log.Index >= Log.Count - 1);
		}
	});
	var h = GetIconSize(item.getAttribute("IconSize"));
	var src = item.getAttribute("Icon") || (h <= 16 ? "bitmap:ieframe.dll,206,16,0" : "bitmap:ieframe.dll,214,24,0");
	SetAddon(Addon_Id, Default, ['<span class="button" onclick="Addons.Back.Exec(this)" oncontextmenu="return Addons.Back.Popup(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', GetImgTag({ id: "ImgBack", title: "Back", src: src }, h), '</span>']);
}
