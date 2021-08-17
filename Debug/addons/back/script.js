const Addon_Id = "back";
const Default = "ToolBar2Left";
if (window.Addon == 1) {
	Addons.Back = {
		Exec: function (Ctrl, pt) {
			Exec(Ctrl, "Back", "Tabs", 0, pt);
		},

		ExecEx: async function (el) {
			Exec(await GetFolderView(el), "Back", "Tabs", 0);
		},

		Popup: async function (el) {
			const FV = await te.Ctrl(CTRL_FV);
			if (FV) {
				const Log = await FV.History;
				const hMenu = await api.CreatePopupMenu();
				const mii = await api.Memory("MENUITEMINFO");
				mii.fMask = MIIM_ID | MIIM_STRING | MIIM_BITMAP;
				const nCount = await Log.Count;
				for (let i = await Log.Index + 1; i < nCount; i++) {
					const FolderItem = await Log[i];
					AddMenuIconFolderItem(mii, FolderItem);
					mii.dwTypeData = await FolderItem.Name;
					mii.wID = i;
					await api.InsertMenuItem(hMenu, MAXINT, false, mii);
				}
				const pt = GetPos(el, 9);
				const nVerb = await api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, ui_.hwnd, null, null);
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

	AddEvent("Layout", async function () {
		const item = await GetAddonElement(Addon_Id);
		const h = GetIconSize(item.getAttribute("IconSize"));
		SetAddon(Addon_Id, Default, ['<span class="button" onclick="Addons.Back.ExecEx(this)" oncontextmenu="Addons.Back.Popup(this); return false;" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({ id: "ImgBack",
			title: await GetText("Back"),
			src: item.getAttribute("Icon") || "icon:general,0"
		}, h), '</span>']);
	});

	AddEvent("ChangeView1", async function (Ctrl) {
		const Log = await Ctrl.History;
		DisableImage(document.getElementById("ImgBack"), Log && await Log.Index >= await Log.Count - 1);
	});
}
