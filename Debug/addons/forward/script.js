var Addon_Id = "forward";
var Default = "ToolBar2Left";

if (window.Addon == 1) {
	Addons.Forward =
	{
		Popup: function (o)
		{
			var FV = te.Ctrl(CTRL_FV);
			if (FV) {
				var Log = FV.History;
				var hMenu = api.CreatePopupMenu();
				var mii = api.Memory("MENUITEMINFO");
				mii.cbSize = mii.Size;
				mii.fMask  = MIIM_ID | MIIM_STRING | MIIM_BITMAP;
				var arBM = [];
				for (var i = Log.Index; i-- > 0;) {
					var FolderItem = Log.Item(i);
					AddMenuIconFolderItem(mii, FolderItem);
					mii.dwTypeData = api.GetDisplayNameOf(FolderItem, SHGDN_INFOLDER | SHGDN_ORIGINAL);
					mii.wID = i + 1;
					api.InsertMenuItem(hMenu, MAXINT, false, mii);
				}
				var pt = api.Memory("POINT");
				api.GetCursorPos(pt);
				var nVerb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, null);
				api.DestroyMenu(hMenu);
				if (nVerb) {
					Log.Index = nVerb - 1;
					FV.History = Log;
				}
			}
		}
	};

	AddEvent("ChangeView", function (Ctrl)
	{
		if (Ctrl.Id == Ctrl.Parent.Selected.Id) {
			var Log = Ctrl.History;
			DisableImage(document.getElementById("ImgForward"), Log && Log.Index == 0);
		}
	});

	var h = GetAddonOption(Addon_Id, "IconSize") || window.IconSize || 24;
	var src = GetAddonOption(Addon_Id, "Icon") || (h <= 16 ? "bitmap:ieframe.dll,206,16,1" : "bitmap:ieframe.dll,214,24,1");
	var s = ['<span class="button" onclick="Navigate(null, SBSP_NAVIGATEFORWARD | SBSP_SAMEBROWSER); return false;" oncontextmenu="Addons.Forward.Popup(this); return false;" onmouseover="MouseOver(this)" onmouseout="MouseOut()"><img id="ImgForward" title="Forward" src="', src.replace(/"/g, ""), '" width="', h, 'px" height="', h, 'px"></span>'];
	SetAddon(Addon_Id, Default, s);
}
