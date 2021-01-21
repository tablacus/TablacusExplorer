Sync.Key = {
	Menus: [],

	OpenMenu: function (Ctrl, pt, nIndex) {
		const ar = this.Menus[nIndex];
		const items = ar[0];
		if (items) {
			const arMenu = ar[1];
			const hMenu = api.CreatePopupMenu();
			MakeMenus(hMenu, null, arMenu, items, Ctrl, pt);
			AdjustMenuBreak(hMenu);
			window.g_menu_click = 2;
			const nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null);
			api.DestroyMenu(hMenu);
			if (nVerb == 0) {
				return S_OK;
			}
			const item = items[nVerb - 1];
			const s = item.getAttribute("Type");
			Exec(Ctrl, item.text, window.g_menu_button == 3 && Sametext(s, "Open") ? "Open in new tab" : s, Ctrl.hwnd, pt);
		}
		return S_OK;
	}
}

const xml = OpenXml("key.xml", false, true);
for (let mode in eventTE.Key) {
	const items = xml.getElementsByTagName(mode);
	for (let i = 0; i < items.length; i++) {
		let item = items[i];
		const strKey = item.getAttribute("Key");
		let strType = item.getAttribute("Type");
		let strOpt = item.text;
		if (SameText(strType, "menus") && SameText(strOpt, "open")) {
			const arMenu = [];
			let nLevel = 0;
			while (++i < items.length) {
				item = items[i];
				if (SameText(item.getAttribute("Type"), "menus")) {
					if (SameText(item.text, "open")) {
						nLevel++;
					}
					if (SameText(item.text, "close")) {
						if (--nLevel < 0) {
							break;
						}
					}
				}
				arMenu.push(i);
			}
			strType = "JScript";
			strOpt = ["Sync.Key.OpenMenu(Ctrl, pt,", Sync.Key.Menus.length, ");"].join("");
			Sync.Key.Menus.push([items, arMenu]);
		}
		SetKeyExec(mode, strKey, strOpt, strType);
	}
}
