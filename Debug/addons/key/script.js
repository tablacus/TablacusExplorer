if (window.Addon == 1) {
	Addons.Key =
	{
		Menus: [],

		OpenMenu: function (Ctrl, pt, nIndex)
		{
			var ar =this.Menus[nIndex];
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
				Exec(Ctrl, item.text, window.g_menu_button == 3 && s == "Open" ? "Open in New Tab" : s, Ctrl.hwnd, pt);
			}
			return S_OK;
		}
	}

	var xml = OpenXml("key.xml", false, true);
	for (var mode in eventTE.Key) {
		var items = xml.getElementsByTagName(mode);
		for (var i = 0; i < items.length; i++) {
			var item = items[i];
			var strKey = String(item.getAttribute("Key"));
			var strType = item.getAttribute("Type");
			var strOpt = item.text;
			if (String(strType).toLowerCase() == "menus" && String(strOpt.toLowerCase()) == "open") {
				var arMenu = [];
				var nLevel = 0;
				while (++i < items.length) {
					item = items[i];
					strType = String(item.getAttribute("Type")).toLowerCase();
					strOpt = String(item.text);
					if (strType == "menus") {
						if (strOpt.toLowerCase() == "open") {
							nLevel++;
						}
						if (strOpt.toLowerCase() == "close") {
							if (--nLevel < 0) {
								break;
							}
						}
						continue;
					}
					arMenu.push(i);
				}
				strType = "JScript";
				strOpt = ["Addons.Key.OpenMenu(Ctrl, pt,", Addons.Key.Menus.length, ");"].join("");
				Addons.Key.Menus.push([items, arMenu]);
			}
			SetKeyExec(mode, strKey, strOpt, strType);
		}
	}
} else {
	importScript("addons\\" + Addon_Id + "\\options.js");
}