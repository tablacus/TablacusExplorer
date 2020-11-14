var Addon_Id = "addressbar";
var Default = "ToolBar2Center";

var item = await GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("Menu", "Edit");
	item.setAttribute("MenuPos", -1);

	item.setAttribute("KeyExec", 1);
	item.setAttribute("KeyOn", "All");
	item.setAttribute("Key", "Alt+D");
}

if (window.Addon == 1) {
	Addons.AddressBar = {
		tid: null,
		Item: null,
		bLoop: false,
		nLevel: 0,
		tid2: null,
		bClose: false,
		XP: item.getAttribute("XP"),
		nPos: 0,
		nWidth: 0,
		strName: "Address bar",

		KeyDown: function (ev, o) {
			if (ev.keyCode == VK_RETURN || /^Enter/i.test(ev.key)) {
				setTimeout(async function (o, str) {
					if (str == o.value) {
						var pt = GetPos(o, 9);
						$.Input = str;
						if (await ExecMenu(await te.Ctrl(CTRL_WB), "Alias", pt, 2) != S_OK) {
							NavigateFV(await te.Ctrl(CTRL_FV), str, await GetNavigateFlags(), true);
						}
					}
				}, 99, o, o.value);
			}
			return true;
		},

		Resize: function () {
			clearTimeout(Addons.AddressBar.tid);
			Addons.AddressBar.tid = setTimeout(Addons.AddressBar.Arrange, 500);
		},

		Arrange: async function (FV) {
			Addons.AddressBar.tid = null;
			var FolderItem = FV && await FV.FolderItem;
			if (FolderItem) {
				var bRoot = await api.ILIsEmpty(FolderItem);
				var o = document.getElementById("breadcrumbbuttons");
				var oAddr = document.F.addressbar;
				var oImg = document.getElementById("addr_img");
				var oPopup = document.getElementById("addressbarselect");
				var width = oAddr.offsetWidth - oImg.offsetWidth + oPopup.offsetWidth - 2;
				var height = oAddr.offsetHeight - 6;
				if (Addons.AddressBar.XP) {
					oAddr.style.color = "";
				} else {
					var SessionId = await FV.SessionId;
					var arHTML = [];
					o.style.width = "auto";
					var n = 0;

					var Items = [];
					do {
						Items.push({
							next: n || await api.GetAttributesOf(FolderItem, SFGAO_HASSUBFOLDER),
							name: await GetFolderItemName(FolderItem)
						});
						FolderItem = await api.ILGetParent(FolderItem);
						n++;
					} while (!await api.ILIsEmpty(FolderItem) && n < 99);
					var bEmpty = true;
					if (o.innerHTML == "" || SessionId == await te.Ctrl(CTRL_FV).SessionId) {
						o.innerHTML = "";
						for (n = 0; n < Items.length; ++n) {
							if (Items[n].next) {
								arHTML.unshift('<span id="addressbar' + n + '" class="button" style="line-height: ' + height + 'px; vertical-align: middle" onclick="Addons.AddressBar.Popup(this,' + n + ')" onmouseover="MouseOver(this)" onmouseout="MouseOut()" oncontextmenu="Addons.AddressBar.Exec(); return false;">' + BUTTONS.next + '</span>');
								o.insertAdjacentHTML("afterbegin", arHTML[0]);
							}
							arHTML.unshift('<span id="addressbar' + n + '_" class="button" style="line-height: ' + height + 'px" onclick="Addons.AddressBar.Go(this, ' + n + ')" onmousedown="return Addons.AddressBar.GoEx(event, this, ' + n + ')" onmouseover="MouseOver(this)" onmouseout="MouseOut()" oncontextmenu="Addons.AddressBar.Exec(); return false;">' + EncodeSC(Items[n].name) + '</span>');
							o.insertAdjacentHTML("afterbegin", arHTML[0]);
							if (o.offsetWidth > width && n > 0) {
								arHTML.splice(0, 2);
								o.innerHTML = arHTML.join("");
								bEmpty = false;
								break;
							}
						}
						o.style.width = (oAddr.offsetWidth - 2) + "px";
						if (bEmpty) {
							if (!bRoot) {
								o.insertAdjacentHTML("afterbegin", '<span id="addressbar' + n + '" class="button" style="line-height: ' + height + 'px" onclick="Addons.AddressBar.Popup(this, ' + n + ')" onmouseover="MouseOver(this)" onmouseout="MouseOut()">' + BUTTONS.next + '</span>');
							}
						} else {
							o.insertAdjacentHTML("afterbegin", '<span id="addressbar' + n + '" class="button" style="line-height: ' + height + 'px" onclick="Addons.AddressBar.Popup2(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">' + BUTTONS.parent + '</span>');
						}
					}
					Addons.AddressBar.nLevel = n;
				}
				oPopup.style.left = (oAddr.offsetWidth - oPopup.offsetWidth - 1) + "px";
				oPopup.style.lineHeight = Math.abs(oAddr.offsetHeight - 6) + "px";
				oImg.style.top = Math.abs(oAddr.offsetHeight - oImg.offsetHeight) / 2 + "px";
			}
		},

		Exec: function () {
			WebBrowser.Focus();
			document.F.addressbar.focus();
		},

		Focus: function () {
			var o = document.getElementById("addressbar");
			if (Addons.AddressBar.bClose) {
				o.blur();
			} else {
				o.select();
				document.getElementById("breadcrumbbuttons").style.display = "none";
			}
		},

		Blur: function () {
			if (!Addons.AddressBar.XP) {
				document.getElementById("breadcrumbbuttons").style.display = "inline-block";
				ClearAutocomplete();
			}
		},

		Go: async function (o, n) {
			Navigate(await Addons.AddressBar.GetPath(n), await GetNavigateFlags());
			api.OutputDebugString("Go:" + document.activeElement.tagName + "\n");

		},

		GoEx: async function (ev, o, n) {
			if (ev.button == 1) {
				Addons.AddressBar.Go(o, n);
				return false;
			} else if (ev.button == 2) {
				var pt = GetPos(o, 9);
				MouseOver(o);
				var hMenu = await api.CreatePopupMenu();
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 1, await api.LoadString(hShell32, 33561));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 2, await GetText("Copy full path"));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 3, await GetText("Open in new &tab"));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 4, await GetText("Open in background"));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 5, await GetText("&Edit"));
				var nVerb = await api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, await te.hwnd, null, null);
				api.DestroyMenu(hMenu);
				switch (nVerb) {
					case 1:
						var Items = await api.CreateObject("FolderItems");
						Items.AddItem(await Addons.AddressBar.GetPath(n));
						api.OleSetClipboard(Items);
						break;
					case 2:
						api.SetclipboardData(await Addons.AddressBar.GetPath(n).Path);
						break;
					case 3:
						Navigate(await Addons.AddressBar.GetPath(n), SBSP_NEWBROWSER);
						break;
					case 4:
						Navigate(await Addons.AddressBar.GetPath(n), SBSP_NEWBROWSER | SBSP_ACTIVATE_NOFOCUS);
						break;
					case 5:
						Addons.AddressBar.Focus();
						break;
				}
				return false;
			}
		},

		GetPath: async function (n) {
			var FolderItem = 0;
			var FV = await te.Ctrl(CTRL_FV);
			if (FV) {
				for (FolderItem = await FV.FolderItem; n > 0; n--) {
					FolderItem = await api.ILGetParent(FolderItem);
				}
			}
			return FolderItem;
		},

		SavePos: async function (o) {
			var rc = await api.CreateObject("Array");
			Common.AddressBar.rcItem = rc;
			for (var i = Addons.AddressBar.nLevel; i >= 0; --i) {
				var el = document.getElementById("addressbar" + i);
				if (el && o.id != el.id) {
					rc[i] = await GetRect(el);
				}
			}
		},

		Popup: async function (o, n) {
			if (Addons.AddressBar.CanPopup()) {
				await Addons.AddressBar.SavePos(o);
				var pt = GetPos(o, 9);
				MouseOver(o);
				$.FolderMenu.Invoke(await $.FolderMenu.Open(await Addons.AddressBar.GetPath(n), pt.x, pt.y, null, 1));
			}
		},

		Popup2: async function (o) {
			var FV = await te.Ctrl(CTRL_FV);
			if (FV) {
				var FolderItem = await FV.FolderItem;
				await $.FolderMenu.Clear();
				var hMenu = await api.CreatePopupMenu();
				var n = 99;
				while (!await api.ILIsEmpty(FolderItem) && n--) {
					FolderItem = await api.ILGetParent(FolderItem);
					await $.FolderMenu.AddMenuItem(hMenu, FolderItem);
				}
				await Addons.AddressBar.SavePos(o);
				AddEvent("ExitMenuLoop", function () {
					Addons.AddressBar.bLoop = false;
					Common.AddressBar.rcItem = void 0;
				});
				MouseOver(o);
				var pt = GetPos(o, 9);
				var nVerb = await $.FolderMenu.TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y);
				FolderItem = nVerb ? await $.FolderMenu.Items[nVerb - 1] : null;
				await $.FolderMenu.Clear();
				$.FolderMenu.Invoke(FolderItem);
			}
		},

		Popup3: async function (o) {
			if (Addons.AddressBar.CanPopup()) {
				var pt = GetPos(o, 9);
				$.FolderMenu.LocationEx(pt.x + o.offsetWidth, pt.y);
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
			if (!window.chrome) {
				o.select();
			}
		},

		ChangeMenu: function (i) {
			setTimeout(function (i) {
				Addons.AddressBar.bClose = false;
				var el = document.getElementById("addressbar" + i);
				if (el) {
					el.click();
				}
			}, 99, i);
		}
	};

	AddEvent("ChangeView", async function (Ctrl) {
		if (await Ctrl.FolderItem && await Ctrl.Id == await Ctrl.Parent.Selected.Id && await Ctrl.Parent.Id == await te.Ctrl(CTRL_TC).Id) {
			document.F.addressbar.value = await Ctrl.FolderItem.Path;
			Addons.AddressBar.Arrange(Ctrl);
			document.getElementById("addr_img").src = await GetIconImage(Ctrl, await GetSysColor(COLOR_WINDOW));
		}
	});

	AddEvent("Resize", async function () {
		Addons.AddressBar.Arrange(await te.Ctrl(CTRL_FV));
	});

	AddEvent("SetAddress", function (s) {
		document.F.addressbar.value = s;
	});

	GetAddress = function () {
		return document.F.addressbar.value;
	}

	//Key
	if (item.getAttribute("KeyExec")) {
		SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Addons.AddressBar.Exec, "Async");
	}
	//Mouse
	if (item.getAttribute("MouseExec")) {
		SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Addons.AddressBar.Exec, "Async");
	}

	AddTypeEx("Add-ons", "Address Bar", Addons.AddressBar.Exec);

	var s = item.getAttribute("Width");
	if (s) {
		if (GetNum(s) == s) {
			s += "px";
		}
	} else {
		s = "100%";
	}
	var nSize = await api.GetSystemMetrics(SM_CYSMICON);
	s = ['<div style="position: relative; overflow: hidden"><div id="breadcrumbbuttons" class="breadcrumb" style="position: absolute; left: 1px; top: 1px; padding-left: ', nSize + 4, 'px" onfocus="Addons.AddressBar.Focus()" onclick="return Addons.AddressBar.Exec();"></div><input id="addressbar" type="text" autocomplate="on" list="AddressList" onkeydown="return Addons.AddressBar.KeyDown(event, this)" onfocus="Addons.AddressBar.Focus()" onblur="Addons.AddressBar.Blur()" onresize="Addons.AddressBar.Resize()" oninput="AdjustAutocomplete(this.value)" oncontextmenu="Addons.AddressBar.ContextMenu(this)" style="width: ', EncodeSC(s), '; vertical-align: middle; padding-left: ', nSize + 4, 'px; padding-right: 16px"><div class="breadcrumb"><div id="addressbarselect" class="button" style="position: absolute; top: 1px" onmouseover="MouseOver(this);" onmouseout="MouseOut()" onclick="Addons.AddressBar.Popup3(this)">', BUTTONS.dropdown, '</div></div>'];

	s.push('<img id="addr_img" src="', await MakeImgSrc("folder:closed"), '"');
	s.push(' onclick="return Addons.AddressBar.Exec();"');
	s.push(' oncontextmenu="Addons.AddressBar.Exec(); return false;"');
	s.push(' style="position: absolute; left: 4px; top: 1.5pt; width: ', nSize, 'px; height: ', nSize, 'px; z-index: 3; border: 0px"></div>');

	SetAddon(Addon_Id, Default, s, "middle");
	Addons.AddressBar.Resize();

	Common.AddressBar = await api.CreateObject("Object");
	Common.AddressBar.MenuExec = item.getAttribute("MenuExec");
	Common.AddressBar.strName = item.getAttribute("MenuName") || await GetAddonInfo(Addon_Id).Name;
	Common.AddressBar.nPos = GetNum(item.getAttribute("MenuPos"));
	var s = item.getAttribute("MenuName");
	if (s && s != "") {
		Common.AddressBar.strName = s;
	}


	importJScript("addons\\" + Addon_Id + "\\sync.js");


} else {
	SetTabContents(0, "General", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
}
