Sync.Mouse = {
	Menus: [],

	OpenMenu: function (Ctrl, pt, nIndex) {
		var ar = this.Menus[nIndex];
		var items = ar[0];
		if (items) {
			setTimeout(function () {
				var arMenu = ar[1];
				var hMenu = api.CreatePopupMenu();
				MakeMenus(hMenu, null, arMenu, items, Ctrl, pt);
				AdjustMenuBreak(hMenu);
				window.g_menu_click = 2;
				var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null);
				api.DestroyMenu(hMenu);
				if (nVerb == 0) {
					return;
				}
				item = items[nVerb - 1];
				var s = item.getAttribute("Type");
				Exec(Ctrl, item.text, window.g_menu_button == 3 && s == "Open" ? "Open in New Tab" : s, Ctrl.hwnd, pt);
			}, 99);
		}
		return S_OK;
	}
}

var xml = OpenXml("mouse.xml", false, true);
for (var mode in eventTE.Mouse) {
	var items = xml.getElementsByTagName(mode);
	for (var i = 0; i < items.length; ++i) {
		var item = items[i];
		var strMouse = item.getAttribute("Mouse");
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
			strOpt = ["Sync.Mouse.OpenMenu(Ctrl, pt,", Sync.Mouse.Menus.length, ");"].join("");
			Sync.Mouse.Menus.push([items, arMenu]);
		}
		SetGestureExec(mode, strMouse, strOpt, strType, false, item.getAttribute("Name"));
	}
}
