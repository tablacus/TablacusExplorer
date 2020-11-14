var Addon_Id = "back";
var Default = "ToolBar2Left";

if (window.Addon == 1) {
	var item = await GetAddonElement(Addon_Id);

	Addons.Back = {
		Exec: function (Ctrl, pt) {
			Exec(Ctrl, "Back", "Tabs", 0, pt);
		},

		ExecEx: async function (el) {
			Exec(await GetFolderViewEx(el), "Back", "Tabs", 0);
		},

		Popup: async function (ev) {
			var FV = await te.Ctrl(CTRL_FV);
			if (FV) {
				var Log = await FV.History;
				var hMenu = await api.CreatePopupMenu();
				var mii = await api.Memory("MENUITEMINFO");
				mii.cbSize = await mii.Size;
				mii.fMask = MIIM_ID | MIIM_STRING | MIIM_BITMAP;
				var nCount = await Log.Count;
				for (var i = await Log.Index + 1; i < nCount; i++) {
					var FolderItem = await Log[i];
					AddMenuIconFolderItem(mii, FolderItem);
					mii.dwTypeData = await FolderItem.Name;
					mii.wID = i;
					await api.InsertMenuItem(hMenu, MAXINT, false, mii);
				}
				var x = ev.screenX * ui_.Zoom;
				var y = ev.screenX * ui_.Zoom;
				var nVerb = await api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, x, y, await te.hwnd, null, null);
				api.DestroyMenu(hMenu);
				if (nVerb) {
					var wFlags = await GetNavigateFlags(FV);
					if (wFlags & SBSP_NEWBROWSER) {
						FV.Navigate(await Log[nVerb], wFlags);
					} else {
						Log.Index = nVerb;
						FV.History = Log;
					}
				}
			}
		}
	};

	AddEvent("ChangeView", async function (Ctrl) {
		if (await Ctrl.Id == await Ctrl.Parent.Selected.Id) {
			var Log = await Ctrl.History;
			DisableImage(document.getElementById("ImgBack"), Log && await Log.Index >= await Log.Count - 1);
		}
	});
	var h = GetIconSize(item.getAttribute("IconSize"));
	var src = item.getAttribute("Icon") || (h <= 16 ? "bitmap:ieframe.dll,206,16,0" : "bitmap:ieframe.dll,214,24,0");
	SetAddon(Addon_Id, Default, ['<span class="button" onclick="Addons.Back.ExecEx(this)" oncontextmenu="Addons.Back.Popup(event); return false;" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({ id: "ImgBack", title: "Back", src: src }, h), '</span>']);
}
