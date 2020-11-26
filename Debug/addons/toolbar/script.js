var Addon_Id = "toolbar";
var Default = "ToolBar2Left";
var AddonName = "ToolBar";

if (window.Addon == 1) {
	Addons.ToolBar = {
		Click: async function (i, bNew) {
			var items = await GetXmlItems(await te.Data.xmlToolBar.getElementsByTagName("Item"));
			var item = items[i];
			if (item) {
				Exec(te, item.text, (bNew && /^Open$|^Open in background$/i.test(type)) ? "Open in new tab" : item.Type, ui_.hwnd, null);
			}
			return false;
		},

		Down: function (ev, i) {
			if (ev.button == 1) {
				return this.Click(i, true);
			}
		},

		Open: async function (ev, i) {
			if (Addons.ToolBar.bClose) {
				return S_OK;
			}
			if (ev.button == 0) {
				var items = await te.Data.xmlToolBar.getElementsByTagName("Item");
				var item = await items[i];
				var hMenu = await api.CreatePopupMenu();
				var arMenu = await api.CreateObject("Array");
				for (var j = GetLength(items); --j > i;) {
					await arMenu.unshift(j);
				}
				var o = document.getElementById("_toolbar" + i);
				var pt = GetPosEx(o, 9);
				await MakeMenus(hMenu, null, arMenu, items, te, pt);
				await AdjustMenuBreak(hMenu);
				AddEvent("ExitMenuLoop", function () {
					Addons.ToolBar.bClose = true;
					setTimeout(function () {
						Addons.ToolBar.bClose = false;
					}, 99);
				});
				var nVerb = await api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, await pt.x, await pt.y, ui_.hwnd, null);
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
				var hMenu = await api.CreatePopupMenu();
				await api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 1, await GetText("&Edit"));
				await api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 2, await GetText("Add"));
				var nVerb = await api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, ev.screenX * ui_.Zoom, ev.screenY * ui_.Zoom, ui_.hwnd, null, null);
				if (nVerb == 1) {
					this.ShowOptions(i + 1);
				}
				if (nVerb == 2) {
					this.ShowOptions();
				}
				api.DestroyMenu(hMenu);
			}
		},

		ShowOptions: function (nEdit) {
			AddonOptions("toolbar", function () {
				Addons.ToolBar.Arrange();
				ApplyLang(document);
			}, { nEdit: nEdit });
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
			var s = [];
			var items = await GetXmlItems(await te.Data.xmlToolBar.getElementsByTagName("Item"));
			var menus = 0;
			var nLen = items.length;
			for (var i = 0; i < nLen; ++i) {
				var item = items[i];
				var strType = item.Type;
				var strFlag = (SameText(strType, "Menus") ? item.text : "").toLowerCase();
				if (strFlag == "close" && menus) {
					menus--;
					continue;
				}
				if (strFlag == "open") {
					if (menus++) {
						continue;
					}
				} else if (menus) {
					continue;
				}
				var img = EncodeSC(await ExtractMacro(null, item.Name));
				if (img == "/" || strFlag == "break") {
					s.push('<br class="break">');
				} else if (img == "//" || strFlag == "barbreak") {
					s.push('<hr class="barbreak">');
				} else if (img == "-" || strFlag == "separator") {
					s.push('<span class="separator">|</span>');
				} else {
					var icon = item.Icon;
					if (icon) {
						var h = EncodeSC(item.Height);
						var sh = {
							src: await api.PathUnquoteSpaces(await ExtractMacro(null, icon))
						};
						if (h != "") {
							sh.style = 'height:' + h + 'px';
						}
						h -= 0;
						img = await GetImgTag(sh, h);
					}
					s.push('<span id="_toolbar', i, '" ', !SameText(strType, "Menus") || !SameText(await item.text, "Open") ? 'onclick="Addons.ToolBar.Click(' + i + ')" onmousedown="Addons.ToolBar.Down(event, ' : 'onmousedown="Addons.ToolBar.Open(event, ');
					s.push(i, ')" oncontextmenu="Addons.ToolBar.Popup(event, ', i, '); return false" onmouseover="MouseOver(this)" onmouseout="MouseOut()" class="button" title="', EncodeSC(await ExtractMacro(te, await item.Name)), '">', img, '</span>');
				}
			}
			if (nLen == 0) {
				s.push('<label id="_toolbar', nLen, '" title="Edit" onclick="Addons.ToolBar.ShowOptions()" oncontextmenu="Addons.ToolBar.ShowOptions(); return false" onmouseover="MouseOver(this)" onmouseout="MouseOut()" class="button">+</label>');
			}
			document.getElementById('_toolbar').innerHTML = s.join("");
			Resize();
		}
	}
	te.Data.xmlToolBar = await OpenXml("toolbar.xml", false, true);
	SetAddon(Addon_Id, Default, '<span id="_' + Addon_Id + '"></span>');
	Addons.ToolBar.Arrange();

	if (!window.chrome) {
		AddEvent("DragEnter", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
			if (Ctrl.Type == CTRL_WB) {
				return S_OK;
			}
		});

		AddEvent("DragOver", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
			if (Ctrl.Type == CTRL_WB) {
				var items = te.Data.xmlToolBar.getElementsByTagName("Item");
				var i = Addons.ToolBar.FromPt(items.length + 1, pt);
				if (i >= 0) {
					if (i == items.length) {
						pdwEffect[0] = DROPEFFECT_LINK;
						MouseOver(document.getElementById("_toolbar" + i));
						return S_OK;
					}
					hr = Exec(te, items[i].text, items[i].getAttribute("Type"), te.hwnd, pt, dataObj, grfKeyState, pdwEffect);
					if (hr == S_OK && pdwEffect[0]) {
						MouseOver(document.getElementById("_toolbar" + i));
					}
					return S_OK;
				}
			}
			MouseOut("_toolbar");
		});

		AddEvent("Drop", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
			MouseOut();
			if (Ctrl.Type == CTRL_WB) {
				var items = te.Data.xmlToolBar.getElementsByTagName("Item");
				var i = Addons.ToolBar.FromPt(items.length + 1, pt);
				if (i >= 0) {
					if (i == items.length) {
						var xml = te.Data.xmlToolBar;
						var root = xml.documentElement;
						if (!root) {
							xml.appendChild(xml.createProcessingInstruction("xml", 'version="1.0" encoding="UTF-8"'));
							root = xml.createElement("TablacusExplorer");
							xml.appendChild(root);
						}
						if (root) {
							for (i = 0; i < dataObj.Count; i++) {
								var FolderItem = dataObj.Item(i);
								var item = xml.createElement("Item");
								item.setAttribute("Name", api.GetDisplayNameOf(FolderItem, SHGDN_INFOLDER));
								item.text = api.GetDisplayNameOf(FolderItem, SHGDN_FORPARSINGEX | SHGDN_FORPARSING);
								if (fso.FileExists(item.text)) {
									item.text = api.PathQuoteSpaces(item.text);
									item.setAttribute("Type", "Exec");
								} else {
									item.setAttribute("Type", "Open");
								}
								root.appendChild(item);
							}
							SaveXmlEx("toolbar.xml", xml);
							Addons.ToolBar.Arrange();
							ApplyLang(document);
						}
						return S_OK;
					}
					return Exec(te, items[i].text, items[i].getAttribute("Type"), te.hwnd, pt, dataObj, grfKeyState, pdwEffect, true);
				}
			}
		});

		AddEvent("DragLeave", function (Ctrl) {
			MouseOut();
			return S_OK;
		});
	}
} else {
	importScript("addons\\" + Addon_Id + "\\options.js");
}

