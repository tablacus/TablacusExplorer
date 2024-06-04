const Addon_Id = "addressbar";
const Default = "ToolBar2Center";
const item = GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("Menu", "Edit");
	item.setAttribute("MenuPos", -1);

	item.setAttribute("KeyExec", 1);
	item.setAttribute("KeyOn", "All");
	item.setAttribute("Key", "Alt+D");
}

if (window.Addon == 1) {
	Addons.AddressBar = {
		Item: null,
		bLoop: false,
		nLevel: 0,
		bClose: false,
		XP: item.getAttribute("XP"),
		nPos: 0,
		nWidth: 0,

		KeyDown: function (ev, o) {
			if (ev.keyCode ? ev.keyCode == VK_RETURN : /^Enter/i.test(ev.key)) {
				setTimeout(async function (o, str) {
					let uri = str.trim().replace(/\//g, "\\");

					if (await fso.FileExists(uri)) {
						let pos = uri.lastIndexOf("\\");
						uri = uri.substr(0, pos)
					}

					if (await fso.FolderExists(uri)) {
						// If it is a file or folder path, navigate to the folder.
						Navigate(uri, SBSP_NEWBROWSER);
					} else if (uri == o.value) {
						const pt = await GetPosEx(o, 9);
						$.Input = uri;
						if (await ExecMenu(await te.Ctrl(CTRL_WB), "Alias", pt, 2) != S_OK) {
							NavigateFV(await te.Ctrl(CTRL_FV), uri, await GetNavigateFlags(), true);
						}
					}
				}, 99, o, o.value);
			}
			return true;
		},

		Resize: async function () {
			Addons.AddressBar.tid = setTimeout(Addons.AddressBar.Arrange, 500);
		},

		Arrange: async function () {
			const FV = await te.Ctrl(CTRL_FV);
			clearTimeout(Addons.AddressBar.tid);
			const FolderItem = FV && await FV.FolderItem;
			if (FolderItem) {
				const o = document.getElementById("breadcrumbbuttons");
				const oAddr = document.F.addressbar;
				const oImg = document.getElementById("addr_img");
				const oPopup = document.getElementById("addressbarselect");
				const width = oAddr.offsetWidth - oImg.offsetWidth + oPopup.offsetWidth - 2;
				const height = oAddr.offsetHeight - 11;
				if (Addons.AddressBar.XP) {
					oAddr.style.color = "";
				} else {
					const arHTML = [];
					o.style.height = "auto";
					const bRoot = api.ILIsEmpty(FolderItem);
					const Items = JSON.parse(await Sync.AddressBar.SplitPath(FolderItem));
					o.style.width = "auto";
					let bEmpty = true, n;
					o.innerHTML = "";
					for (n = 0; n < Items.length; ++n) {
						if (Items[n].next) {
							arHTML.unshift('<span id="addressbar' + n + '" class="button" style="line-height: ' + height + 'px" onclick="Addons.AddressBar.Popup(this,' + n + ')" onmouseover="MouseOver(this)" onmouseout="MouseOut()" oncontextmenu="Addons.AddressBar.Exec(); return false;">' + BUTTONS.next + '</span>');
							o.insertAdjacentHTML("afterbegin", arHTML[0]);
						}
						arHTML.unshift('<span id="addressbar' + n + '_" class="button" style="line-height: ' + height + 'px" onmouseover="MouseOver(this)" onmouseout="MouseOut()" oncontextmenu="Addons.AddressBar.Exec(); return false;" ondragstart="Addons.AddressBar.Drag(event,' + n + '); return false" draggable="true">' + EncodeSC(Items[n].name) + '</span>');
						const nBefore = o.offsetWidth;
						o.insertAdjacentHTML("afterbegin", arHTML[0]);
						if (nBefore != o.offsetWidth && o.offsetWidth > width && n > 0) {
							arHTML.splice(0, 2);
							o.innerHTML = arHTML.join("");
							bEmpty = false;
							break;
						}
					}
					o.style.width = (oAddr.offsetWidth - 2) + "px";
					if (bEmpty) {
						if (!await bRoot) {
							o.insertAdjacentHTML("afterbegin", '<span id="addressbar' + n + '" class="button" style="line-height: ' + height + 'px" onclick="Addons.AddressBar.Popup(this, ' + n + ')" onmouseover="MouseOver(this)" onmouseout="MouseOut()">' + BUTTONS.next + '</span>');
						}
					} else {
						o.insertAdjacentHTML("afterbegin", '<span id="addressbar' + n + '" class="button" style="line-height: ' + height + 'px" onclick="Addons.AddressBar.Popup2(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">' + BUTTONS.parent + '</span>');
					}
					Addons.AddressBar.nLevel = n;
				}
				oPopup.style.left = (oAddr.offsetWidth - oPopup.offsetWidth - 2) + "px";
				oPopup.style.lineHeight = Math.abs(height) + "px";
			}
		},

		Exec: function () {
			WebBrowser.Focus();
			document.F.addressbar.focus();
		},

		Focus: function () {
			if (Addons.AddressBar.tm && new Date().getTime() - Addons.AddressBar.tm < 999) {
				delete Addons.AddressBar.tm;
				return;
			}
			const o = document.getElementById("addressbar");
			if (Addons.AddressBar.bClose) {
				o.blur();
			} else {
				if (o.selectionEnd == o.selectionStart) {
					o.select()
				}
				o.focus();
				o.style.color = "";
				document.getElementById("breadcrumbbuttons").style.display = "none";
			}
		},

		Blur: function () {
			if (!Addons.AddressBar.XP) {
				document.getElementById("breadcrumbbuttons").style.display = "inline-block";
				ClearAutocomplete();
				Addons.AddressBar.Arrange();
			}
		},

		Click1: function (ev) {
			const el = document.elementFromPoint(ev.clientX, ev.clientY);
			const res = el && /^addressbar(\d+)_$/.exec(el.id);
			if (res) {
				Promise.all([Sync.AddressBar.GetPath(res[1]), GetNavigateFlags()]).then(function (r) {
					Navigate(r[0], r[1]);
				});
				return;
			}
			delete Addons.AddressBar.tm;
			Addons.AddressBar.Focus();
		},

		Popup1: async function (ev) {
			const el = document.elementFromPoint(ev.clientX, ev.clientY);
			const res = el && /^addressbar(\d+)_$/.exec(el.id)
			if (res) {
				Addons.AddressBar.tm = new Date().getTime();
				const el = ev.target || ev.srcElement;
				const pt = GetPos(el, 9);
				MouseOver(el);
				const hMenu = await api.CreatePopupMenu();
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 1, await api.LoadString(hShell32, 33561));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 2, await GetText("Copy full path"));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 3, await GetText("Open in new &tab"));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 4, await GetText("Open in background"));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 5, await GetText("&Edit"));
				Addons.AddressBar.tm = new Date().getTime();
				const nVerb = await api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, ui_.hwnd, null, null);
				api.DestroyMenu(hMenu);
				switch (nVerb) {
					case 1:
						const Items = await api.CreateObject("FolderItems");
						Items.AddItem(await Sync.AddressBar.GetPath(res[1]));
						api.OleSetClipboard(Items);
						break;
					case 2:
						api.SetclipboardData(await (await Sync.AddressBar.GetPath(res[1])).Path);
						break;
					case 3:
						Navigate(await Sync.AddressBar.GetPath(res[1]), SBSP_NEWBROWSER);
						break;
					case 4:
						Navigate(await Sync.AddressBar.GetPath(res[1]), SBSP_NEWBROWSER | SBSP_ACTIVATE_NOFOCUS);
						break;
					case 5:
						Addons.AddressBar.Exec();
						Addons.AddressBar.Focus();
						break;
				}
			}
		},

		Drag: async function (ev, n) {
			const DataObj = await api.CreateObject("FolderItems");
			DataObj.AddItem(await Sync.AddressBar.GetPath(n));
			DataObj.dwEffect = DROPEFFECT_LINK;
			DoDragDrop(DataObj, DROPEFFECT_LINK | DROPEFFECT_COPY | DROPEFFECT_MOVE);
		},

		Down1: function (ev) {
			const el = document.elementFromPoint(ev.clientX, ev.clientY);
			if (el && /^addressbar(\d+)_$/.test(el.id)) {
				Addons.AddressBar.tm = new Date().getTime();
				Addons.AddressBar.ev = ev;
				return;
			}
			const buttons = ev.buttons != null ? ev.buttons : ev.button;
			if (buttons == 2) {
				Addons.AddressBar.Focus();
			}
		},

		Up1: function (ev) {
			const ev1 = Addons.AddressBar.ev || {};
			const buttons = ev1.buttons != null ? ev1.buttons : ev1.button;
			if ((buttons & 4) && Math.abs(ev.screenX - ev1.screenX) < 4 && Math.abs(ev.screenY - ev1.screenY) < 4) {
				const el = document.elementFromPoint(ev.clientX, ev.clientY);
				const res = el && /^addressbar(\d+)_$/.exec(el.id);
				if (res) {
					Promise.all([Sync.AddressBar.GetPath(res[1]), GetNavigateFlags()]).then(function (r) {
						Navigate(r[0], r[1] | SBSP_NEWBROWSER);
					});
					return;
				}
			}
			setTimeout(function () {
				delete Addons.AddressBar.ev;
			}, 99);
		},

		SavePos: async function (o) {
			const rc = await api.CreateObject("Array");
			Common.AddressBar.rcItem = rc;
			for (let i = Addons.AddressBar.nLevel; i >= 0; --i) {
				const el = document.getElementById("addressbar" + i);
				if (el && o.id != el.id) {
					rc[i] = await GetRect(el);
				}
			}
		},

		SetRects: async function () {
			const rc = await api.CreateObject("Array");
			Common.AddressBar.rcDrop = rc;
			for (let i = Addons.AddressBar.nLevel; --i >= 0;) {
				const el = document.getElementById("addressbar" + i + "_");
				if (el) {
					rc[i] = await GetRect(el);
				}
			}
		},

		Popup: function (o, n) {
			if (Addons.AddressBar.CanPopup()) {
				setTimeout(async function (o, n) {
					await Addons.AddressBar.SavePos(o);
					const pt = GetPos(o, 9);
					MouseOver(o);
					FolderMenu.Invoke(await FolderMenu.Open(await Sync.AddressBar.GetPath(n), pt.x, pt.y, null, 1));
				}, 9, o, n);
			}
		},

		Popup2: async function (o) {
			const FV = await te.Ctrl(CTRL_FV);
			if (FV) {
				let FolderItem = await FV.FolderItem;
				await FolderMenu.Clear();
				const hMenu = await api.CreatePopupMenu();
				for (let n = 99; !await api.ILIsEmpty(FolderItem) && n--;) {
					FolderItem = await api.ILGetParent(FolderItem);
					await FolderMenu.AddMenuItem(hMenu, FolderItem);
				}
				await Addons.AddressBar.SavePos(o);
				AddEvent("ExitMenuLoop", function () {
					Addons.AddressBar.bLoop = false;
					Common.AddressBar.rcItem = void 0;
				});
				MouseOver(o);
				const pt = GetPos(o, 9);
				const nVerb = await FolderMenu.TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y);
				FolderItem = nVerb ? await FolderMenu.Items[nVerb - 1] : null;
				await FolderMenu.Clear();
				FolderMenu.Invoke(FolderItem);
			}
		},

		Popup3: async function (o) {
			if (Addons.AddressBar.CanPopup()) {
				const pt = GetPos(o, 9);
				FolderMenu.LocationEx(pt.x + o.offsetWidth, pt.y);
			}
		},

		CanPopup: function () {
			if (!Addons.AddressBar.bClose) {
				Addons.AddressBar.bLoop = true;
				Addons.AddressBar.bClose = true;
				AddEvent("ExitMenuLoop", function () {
					Addons.AddressBar.bLoop = false;
					Common.AddressBar.rcItem = void 0;
					Addons.AddressBar.bClose = true;
					clearTimeout(Addons.AddressBar.tid2);
					Addons.AddressBar.tid2 = setTimeout(function () {
						Addons.AddressBar.bClose = false;
					}, 500);

				});
				return true;
			}
			return false;
		},

		ContextMenu: function (o) {
			if (!window.chrome && o.selectionEnd == o.selectionStart) {
				o.select();
			}
		},

		ChangeMenu: function (i) {
			setTimeout(function (i) {
				Addons.AddressBar.bClose = false;
				let el = document.getElementById("addressbar" + i);
				if (el) {
					el.click();
				}
			}, 99, i);
		}
	};

	AddEvent("ChangeView1", async function (Ctrl) {
		await Addons.AddressBar.Arrange();
		document.F.addressbar.value = await Ctrl.FolderItem.Path;
		document.getElementById("addr_img").src = await GetIconImage(Ctrl, CLR_DEFAULT | COLOR_WINDOW);
	});

	AddEvent("Resize", Addons.AddressBar.Resize);

	AddEvent("SetAddress", function (s) {
		document.F.addressbar.value = s;
	});

	GetAddress = function () {
		return document.F.addressbar.value;
	}

	//Menu
	const strName = item.getAttribute("MenuName") || await GetAddonInfo(Addon_Id).Name;
	if (item.getAttribute("MenuExec")) {
		SetMenuExec("AddressBar", strName, item.getAttribute("Menu"), item.getAttribute("MenuPos"));
	}
	//Key
	if (item.getAttribute("KeyExec")) {
		SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Addons.AddressBar.Exec, "Async");
	}
	//Mouse
	if (item.getAttribute("MouseExec")) {
		SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Addons.AddressBar.Exec, "Async");
	}

	AddEvent("Layout", function () {
		let s = item.getAttribute("Width");
		if (s) {
			if (GetNum(s) == s) {
				s += "px";
			}
		} else {
			s = "100%";
		}
		const z = screen.deviceYDPI / 96;
		s = ['<div style="position: relative; overflow: hidden">',
		'<div id="breadcrumbbuttons" class="breadcrumb" style="vertical-align: middle; position: absolute; margin-top:1px; margin-bottom:1px; top:1px; bottom: 1px; left: 1px; padding-left: ', 16 * z + 4, 'px" onclick="Addons.AddressBar.Click1(event)" oncontextmenu="Addons.AddressBar.Popup1(event); return false;" onmousedown="Addons.AddressBar.Down1(event)" onmouseup="return Addons.AddressBar.Up1(event); return false"></div>',
		'<input id="addressbar" type="text" autocomplate="on" list="AddressList" onkeydown="return Addons.AddressBar.KeyDown(event, this)" onfocus="Addons.AddressBar.Focus()" onblur="Addons.AddressBar.Blur()" onresize="Addons.AddressBar.Resize()" oninput="AdjustAutocomplete(this.value)" oncontextmenu="Addons.AddressBar.ContextMenu(this)" style="width: ', EncodeSC(s), '; vertical-align: middle; padding-left: ', 16 * z + 4, 'px; padding-top:2px;padding-bottom:2px; padding-right: 16px">',
		'<div class="breadcrumb"><div id="addressbarselect" class="button" style="vertical-align: middle; position: absolute; margin-top:1px; margin-bottom:1px; top: 1px; bottom: 1px; " onmouseover="MouseOver(this);" onmouseout="MouseOut()" onclick="Addons.AddressBar.Popup3(this)">', BUTTONS.dropdown, '</div></div>'];

		s.push('<img id="addr_img"');
		if (ui_.IEVer < 8) {
			s.push(' src="', MakeImgSrc("folder:closed"), '"');
		}
		s.push(' onclick="return Addons.AddressBar.Exec();"');
		s.push(' oncontextmenu="Addons.AddressBar.Exec(); return false;"');
		s.push(' ondragstart="Addons.AddressBar.Drag(event, 0); return false;" draggable="true"');
		s.push(' style="position: absolute; left: 4px; top:', 2 * z, 'px; width: ', 16 * z, 'px; height: ', 16 * z, 'px; z-index: 3; border: 0px"></div>');

		SetAddon(Addon_Id, Default, s, "middle");
	});

	Common.AddressBar = await api.CreateObject("Object");
	Common.AddressBar.MenuExec = item.getAttribute("MenuExec");
	Common.AddressBar.strName = item.getAttribute("MenuName") || await GetAddonInfo(Addon_Id).Name;
	Common.AddressBar.nPos = GetNum(item.getAttribute("MenuPos"));

	AddTypeEx("Add-ons", "Address Bar", Addons.AddressBar.Exec);

	if (s = item.getAttribute("MenuName")) {
		Common.AddressBar.strName = s;
	}
	$.importScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	SetTabContents(0, "General", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
}
