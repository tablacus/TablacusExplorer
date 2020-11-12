Sync.Key = {
	Menus: [],

	OpenMenu: function (Ctrl, pt, nIndex) {
		var ar = this.Menus[nIndex];
		var items = ar[0];
		if (items) {
			var arMenu = ar[1];
			var hMenu = api.CreatePopupMenu();
			MakeMenus(hMenu, null, arMenu, items, Ctrl, pt);
			AdjustMenuBreak(hMenu);
			window.g_menu_click = 2;
			var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null);
			api.DestroyMenu(hMenu);
			if (nVerb == 0) {
				return S_OK;
			}
			item = items[nVerb - 1];
			var s = item.getAttribute("Type");
			Exec(Ctrl, item.text, window.g_menu_button == 3 && Sametext(s, "Open") ? "Open in new tab" : s, Ctrl.hwnd, pt);
		}
		return S_OK;
	}
}

var xml = OpenXml("key.xml", false, true);
for (var mode in eventTE.Key) {
	var items = xml.getElementsByTagName(mode);
	for (var i = 0; i < items.length; i++) {
		var item = items[i];
		var strKey = item.getAttribute("Key");
		var strType = item.getAttribute("Type");
		var strOpt = item.text;
		if (SameText(strType, "menus") && SameText(strOpt, "open")) {
			var arMenu = [];
			var nLevel = 0;
			while (++i < items.length) {
				item = items[i];
				strType = item.getAttribute("Type");
				strOpt = item.text;
				if (SameText(strType, "menus")) {
					if (SameText(strOpt, "open")) {
						nLevel++;
					}
					if (SameText(strOpt, "close")) {
						if (--nLevel < 0) {
							break;
						}
					}
					continue;
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
