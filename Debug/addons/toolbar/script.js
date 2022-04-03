const Addon_Id = "toolbar";
const Default = "ToolBar2Left";
if (window.Addon == 1) {
	Addons.ToolBar = {
		Click: async function (i, bNew) {
			const items = await GetXmlItems(await te.Data.xmlToolBar.getElementsByTagName("Item"));
			const item = items[i];
			if (item) {
				Exec(te, item.text, (bNew && /^Open$|^Open in background$/i.test(item.type)) ? "Open in new tab" : item.Type, ui_.hwnd, null);
			}
			return false;
		},

		Down: function (ev, i) {
			if ((ev.buttons != null ? ev.buttons : ev.button) == 4) {
				return this.Click(i, true);
			}
		},

		Open: async function (ev, i) {
			if (Addons.ToolBar.bClose) {
				return S_OK;
			}
			if ((ev.buttons != null ? ev.buttons : ev.button) == 1) {
				const items = await te.Data.xmlToolBar.getElementsByTagName("Item");
				let item = await items[i];
				const hMenu = await api.CreatePopupMenu();
				const arMenu = await api.CreateObject("Array");
				for (let j = await GetLength(items); --j > i;) {
					await arMenu.unshift(j);
				}
				const o = document.getElementById("_toolbar" + i);
				const pt = await GetPosEx(o, 9);
				await MakeMenus(hMenu, null, arMenu, items, te, pt);
				await AdjustMenuBreak(hMenu);
				AddEvent("ExitMenuLoop", function () {
					Addons.ToolBar.bClose = true;
					setTimeout(function () {
						Addons.ToolBar.bClose = false;
					}, 99);
				});
				const nVerb = await api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, await pt.x, await pt.y, ui_.hwnd, null);
				api.DestroyMenu(hMenu);
				if (nVerb > 0) {
					item = await items[nVerb - 1];
					Exec(te, await item.text, await item.getAttribute("Type"), ui_.hwnd, null);
				}
				return S_OK;
			}
		},

		Popup: async function (ev, i) {
			if (i >= 0) {
				const hMenu = await api.CreatePopupMenu();
				await api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 1, await GetText("&Edit"));
				await api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 2, await GetText("Add"));
				const nVerb = await api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, ev.screenX * ui_.Zoom, ev.screenY * ui_.Zoom, ui_.hwnd, null, null);
				if (nVerb == 1) {
					this.ShowOptions(i + 1);
				}
				if (nVerb == 2) {
					this.ShowOptions();
				}
				api.DestroyMenu(hMenu);
			}
		},

		ShowOptions: async function (nEdit) {
			const opt = await api.CreateObject("Object");
			opt.nEdit = nEdit;
			AddonOptions("toolbar", "Addons.ToolBar.Changed", opt);
		},

		Changed: function () {
			Addons.ToolBar.Arrange();
			ApplyLang(document);
		},

		FromPt: function (n, pt) {
			while (--n >= 0) {
				if (HitTest(document.getElementById("_toolbar" + n), pt)) {
					return n;
				}
			}
			return -1;
		},

		Arrange: async function () {
			const s = [];
			const items = await GetXmlItems(await te.Data.xmlToolBar.getElementsByTagName("Item"));
			let menus = 0;
			const nLen = items.length;
			for (let i = 0; i < nLen; ++i) {
				const item = items[i];
				const strType = item.Type;
				const strFlag = SameText(strType, "Menus") ? item.text : "";
				if (SameText(strFlag, "close") && menus) {
					menus--;
					continue;
				}
				if (SameText(strFlag, "open")) {
					if (menus++) {
						continue;
					}
				} else if (menus) {
					continue;
				}
				let img = EncodeSC(await ExtractMacro(te, item.Name));
				if (img == "/" || SameText(strFlag, "break")) {
					s.push('<br class="break">');
				} else if (img == "//" || SameText(strFlag, "barbreak")) {
					s.push('<hr class="barbreak">');
				} else if (img == "-" || SameText(strFlag, "separator")) {
					s.push('<span class="separator">|</span>');
				} else {
					let icon = item.Icon;
					if (icon) {
						let h = (item.Height * screen.deviceYDPI / 96);
						let sh = {
							src: await ExtractPath(te, icon)
						};
						if (h && isFinite(h)) {
							sh.style = 'width:' + h + 'px; height:' + h + 'px';
						}
						img = await GetImgTag(sh, h);
					}
					s.push('<span id="_toolbar', i, '" ', !SameText(strType, "Menus") || !SameText(await item.text, "Open") ? 'onclick="Addons.ToolBar.Click(' + i + ')" onmousedown="Addons.ToolBar.Down(event, ' : 'onmousedown="Addons.ToolBar.Open(event, ');
					s.push(i, ')" oncontextmenu="Addons.ToolBar.Popup(event, ', i, '); return false" onmouseover="MouseOver(this)" onmouseout="MouseOut()" class="button" title="', EncodeSC(await ExtractMacro(te, await item.Name)), '">', img, '</span>');
				}
			}
			if (nLen == 0) {
				s.push('<label id="_toolbar+" title="', await GetText("Add"), '" onclick="Addons.ToolBar.ShowOptions()" oncontextmenu="Addons.ToolBar.ShowOptions(); return false" onmouseover="MouseOver(this)" onmouseout="MouseOut()" class="button">+</label>');
			}
			document.getElementById('_toolbar').innerHTML = s.join("");
			Resize();
		},

		Changed: function () {
			Addons.ToolBar.Arrange();
			ApplyLang(document.getElementById("_toolbar"));
		},

		SetRects: async function () {
			const items = await GetXmlItems(await te.Data.xmlToolBar.getElementsByTagName("Item"));
			Common.ToolBar.Count = items.length;
			for (let i = 0; i < items.length; ++i) {
				const el = document.getElementById("_toolbar" + i);
				if (el) {
					Common.ToolBar.Items[i] = await GetRect(el);
				}
			}
			Common.ToolBar.Append = await GetRect(document.getElementById('_toolbar'));
		}
	}

	AddEvent("Layout", async function () {
		await SetAddon(Addon_Id, Default, '<span id="_' + Addon_Id + '"></span>');
		await Addons.ToolBar.Arrange();
	});

	te.Data.xmlToolBar = await OpenXml("toolbar.xml", false, true);
	$.importScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	AddonName = "ToolBar";
	importScript("addons\\" + Addon_Id + "\\options.js");
}
