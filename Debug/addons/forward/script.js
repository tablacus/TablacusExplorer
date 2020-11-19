var Addon_Id = "forward";
var Default = "ToolBar2Left";

if (window.Addon == 1) {
	var item = await GetAddonElement(Addon_Id);

	Addons.Forward = {
		Exec: function (Ctrl, pt) {
			Exec(Ctrl, "Forward", "Tabs", 0, pt);
		},

		ExecEx: async function (el) {
			Exec(await GetFolderViewEx(el), "Forward", "Tabs", 0);
		},

		Popup: async function (ev) {
			var FV = await te.Ctrl(CTRL_FV);
			if (FV) {
				var Log = await FV.History;
				var hMenu = await api.CreatePopupMenu();
				var mii = await api.Memory("MENUITEMINFO");
				mii.cbSize = await mii.Size;
				mii.fMask = MIIM_ID | MIIM_STRING | MIIM_BITMAP;
				for (var i = await Log.Index; i-- > 0;) {
					var FolderItem = await Log[i];
					AddMenuIconFolderItem(mii, FolderItem);
					mii.dwTypeData = await FolderItem.Name;
					mii.wID = i + 1;
					await api.InsertMenuItem(hMenu, MAXINT, false, mii);
				}
				var x = ev.screenX * ui_.Zoom;
				var y = ev.screenY * ui_.Zoom;
				var nVerb = await api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, x, y, await te.hwnd, null, null);
				api.DestroyMenu(hMenu);
				if (nVerb) {
					var wFlags = await GetNavigateFlags(FV);
					if (wFlags & SBSP_NEWBROWSER) {
						FV.Navigate(await Log[nVerb - 1], wFlags);
					} else {
						Log.Index = nVerb - 1;
						FV.History = Log;
					}
				}
			}
		}
	};

	AddEvent("ChangeView", async function (Ctrl) {
		if (await Ctrl.Id == await Ctrl.Parent.Selected.Id) {
			var Log = await Ctrl.History;
			DisableImage(document.getElementById("ImgForward"), Log && await Log.Index < 1);
		}
	});
	var h = GetIconSize(item.getAttribute("IconSize"));
	var src = item.getAttribute("Icon") || (h <= 16 ? "bitmap:ieframe.dll,206,16,1" : "bitmap:ieframe.dll,214,24,1");
	SetAddon(Addon_Id, Default, ['<span class="button" onclick="Addons.Forward.ExecEx(this)" oncontextmenu="Addons.Forward.Popup(event); return false;" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({ id: "ImgForward", title: "Forward", src: src }, h), '</span>']);
}
