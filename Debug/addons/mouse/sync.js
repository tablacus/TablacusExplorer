Sync.Mouse = {
	Menus: [],

	OpenMenu: function (Ctrl, pt, nIndex) {
		const ar = this.Menus[nIndex];
		const items = ar[0];
		if (items) {
			setTimeout(function () {
				const arMenu = ar[1];
				const hMenu = api.CreatePopupMenu();
				MakeMenus(hMenu, null, arMenu, items, Ctrl, pt);
				AdjustMenuBreak(hMenu);
				window.g_menu_click = 2;
				const nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null);
				api.DestroyMenu(hMenu);
				if (nVerb == 0) {
					return;
				}
				const item = items[nVerb - 1];
				const s = item.getAttribute("Type");
				Exec(Ctrl, item.text, window.g_menu_button == 3 && s == "Open" ? "Open in New Tab" : s, Ctrl.hwnd, pt);
			}, 99);
		}
		return S_OK;
	}
}

const xml = OpenXml("mouse.xml", false, true);
for (let mode in eventTE.Mouse) {
	const items = xml.getElementsByTagName(mode);
	for (let i = 0; i < items.length; ++i) {
		let item = items[i];
		const strMouse = item.getAttribute("Mouse");
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
			strOpt = ["Sync.Mouse.OpenMenu(Ctrl, pt,", Sync.Mouse.Menus.length, ");"].join("");
			Sync.Mouse.Menus.push([items, arMenu]);
		}
		SetGestureExec(mode, strMouse, strOpt, strType, false, item.getAttribute("Name"));
	}
}
