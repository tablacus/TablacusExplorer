var Addon_Id = "addressbar";
var Default = "ToolBar2Center";

var item = GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("Menu", "Edit");
	item.setAttribute("MenuPos", -1);

	item.setAttribute("KeyExec", 1);
	item.setAttribute("KeyOn", "All");
	item.setAttribute("Key", "Alt+D");
}

if (window.Addon == 1) {
	Addons.AddressBar =
	{
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

		KeyDown: function (o)
		{
			if (event.keyCode == VK_RETURN) {
				(function (o, str) {
					setTimeout(function () {
						if (str == o.value) {
							var p = GetPos(o);
							var pt = api.Memory("POINT");
							pt.x = screenLeft + p.x;
							pt.y = screenTop + p.y + o.offsetHeight;
							window.Input = str;
							if (ExecMenu(te.Ctrl(CTRL_WB), "Alias", pt, 2) != S_OK) {
								NavigateFV(te.Ctrl(CTRL_FV), str, GetNavigateFlags(), true);
							}
						}
					}, 99);
				})(o, o.value);
			}
			return true;
		},

		Resize: function ()
		{
			clearTimeout(this.tid);
			this.tid = setTimeout(this.Arrange, 500);
		},

		Arrange: function (FolderItem)
		{
			this.tid = null;

			if (!FolderItem) {
				var FV = te.Ctrl(CTRL_FV);
				if (FV) {
					FolderItem = FV.FolderItem;
				}
			}
			if (FolderItem) {
				var bRoot = api.ILIsEmpty(FolderItem);
				var s = [];
				var o = document.getElementById("breadcrumbbuttons");
				var oAddr = document.F.addressbar;
				var oImg = document.getElementById("addr_img");
				var oPopup = document.getElementById("addressbarselect");
				var width = oAddr.offsetWidth - oImg.offsetWidth + oPopup.offsetWidth - 2;
				var height = oAddr.offsetHeight - (6 * screen.deviceYDPI / 96);
				if (Addons.AddressBar.XP) {
					oAddr.style.color = "";
				} else {
					o.style.width = "auto";
					var n = 0;
					do {
						if (n || api.GetAttributesOf(FolderItem, SFGAO_HASSUBFOLDER)) {
							s.unshift('<span id="addressbar' + n + '" class="button" style="line-height: ' + height + 'px; vertical-align: middle" onclick="Addons.AddressBar.Popup(this,' + n + ')" onmouseover="MouseOver(this)" onmouseout="MouseOut()" oncontextmenu="Addons.AddressBar.Exec(); return false;">' + BUTTONS.next + '</span>');
						}
						s.unshift('<span id="addressbar' + n + '_" class="button" style="line-height: ' + height + 'px" onclick="Addons.AddressBar.Go(this, ' + n + ')" onmousedown="return Addons.AddressBar.GoEx(this, ' + n + ')" onmouseover="MouseOver(this)" onmouseout="MouseOut()" oncontextmenu="Addons.AddressBar.Exec(); return false;">' + EncodeSC(GetFolderItemName(FolderItem)) + '</span>');
						FolderItem = api.ILGetParent(FolderItem);
						o.innerHTML = s.join("");
						if (o.offsetWidth > width && n > 0) {
							s.splice(0, 2);
							o.innerHTML = s.join("");
							break;
						}
						n++;
					} while (!api.ILIsEmpty(FolderItem) && n < 99);
					o.style.width = (oAddr.offsetWidth - 2) + "px";
					if (api.ILIsEmpty(FolderItem)) {
						if (!bRoot) {
							o.insertAdjacentHTML("AfterBegin", '<span id="addressbar' + n + '" class="button" style="line-height: ' + height + 'px" onclick="Addons.AddressBar.Popup(this, ' + n + ')" onmouseover="MouseOver(this)" onmouseout="MouseOut()">' + BUTTONS.next + '</span>');
						}
					} else {
						o.insertAdjacentHTML("AfterBegin", '<span id="addressbar' + n + '" class="button" style="line-height: ' + height + 'px" onclick="Addons.AddressBar.Popup2(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()">' + BUTTONS.parent + '</span>');
					}
					this.nLevel = n;
				}
				oPopup.style.left = (oAddr.offsetWidth - oPopup.offsetWidth - 1) + "px";
				oPopup.style.lineHeight = Math.abs(oAddr.offsetHeight - 6) + "px";
				oImg.style.top = Math.abs(oAddr.offsetHeight - oImg.offsetHeight) / 2 + "px";
			}
		},

		Exec: function ()
		{
			document.F.addressbar.focus();
			return S_OK;
		},

		Focus: function ()
		{
			var o = document.getElementById("addressbar");
			if (Addons.AddressBar.bClose) {
				o.blur();
			} else {
				o.select();
				document.getElementById("breadcrumbbuttons").style.display = "none";
			}
		},

		Blur: function ()
		{
			if (!Addons.AddressBar.XP) {
				document.getElementById("breadcrumbbuttons").style.display = "inline-block";
			}
		},

		Go: function (o, n)
		{
			Navigate(this.GetPath(n), GetNavigateFlags());
		},

		GoEx: function (o, n)
		{
			if (event.button == 1) {
				this.Go(o, n);
				return false;
			} else if (event.button == 2) {
				var pt = GetPos(o, true);
				MouseOver(o);
				var hMenu = api.CreatePopupMenu();
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 1, api.LoadString(hShell32, 33561));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 2, GetText("Copy full path"));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 3, GetText("Open in new &tab"));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 4, GetText("Open in background"));
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 5,  GetText("&Edit"));
				var nVerb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y + o.offsetHeight, te.hwnd, null, null);
				api.DestroyMenu(hMenu);
				switch (nVerb) {
					case 1:
						var Items = te.FolderItems();
						Items.AddItem(this.GetPath(n));
						api.OleSetClipboard(Items);
						break;
					case 2:
						clipboardData.setData("text", this.GetPath(n).Path);
						break;
					case 3:
						Navigate(this.GetPath(n), SBSP_NEWBROWSER);
						break;
					case 4:
						Navigate(this.GetPath(n), SBSP_NEWBROWSER | SBSP_ACTIVATE_NOFOCUS);
						break;
					case 5:
						this.Focus();
						break;
				}
				return false;
			}
		},

		GetPath: function(n)
		{
			var FolderItem = 0;
			var FV = te.Ctrl(CTRL_FV);
			if (FV) {
				for (FolderItem = FV.FolderItem; n > 0; n--) {
					FolderItem = api.ILGetParent(FolderItem);
				}
			}
			return FolderItem;
		},

		Popup: function (o, n)
		{
			if (Addons.AddressBar.CanPopup()) {
				Addons.AddressBar.Item = o;
				var pt = GetPos(o, true);
				MouseOver(o);
				FolderMenu.Invoke(FolderMenu.Open(this.GetPath(n), pt.x, pt.y + o.offsetHeight, null, 1));
			}
		},

		Popup2: function (o)
		{
			var FV = te.Ctrl(CTRL_FV);
			if (FV) {
				var FolderItem = FV.FolderItem;
				FolderMenu.Clear();
				var hMenu = api.CreatePopupMenu();
				var n = 99;
				while (!api.ILIsEmpty(FolderItem) && n--) {
					FolderItem = api.ILGetParent(FolderItem);
					FolderMenu.AddMenuItem(hMenu, FolderItem);
				}
				Addons.AddressBar.Item = o;
				Addons.AddressBar.bLoop = true;
				ExitMenuLoop = function () {
					Addons.AddressBar.bLoop = false;
				};
				MouseOver(o);
				var pt = GetPos(o, true);
				window.g_menu_click = true;
				var nVerb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y + o.offsetHeight, te.hwnd, null, null);
				api.DestroyMenu(hMenu);
				FolderItem = null;
				if (nVerb) {
					FolderItem = FolderMenu.Items[nVerb - 1];
				}
				FolderMenu.Clear();
				FolderMenu.Invoke(FolderItem);
			}
		},

		Popup3: function (o)
		{
			if (Addons.AddressBar.CanPopup()) {
				FolderMenu.Location(o);
			}
		},

		CanPopup: function ()
		{
			if (!Addons.AddressBar.bClose) {
				Addons.AddressBar.bLoop = true;
				AddEvent("ExitMenuLoop", function () {
					Addons.AddressBar.bLoop = false;
					Addons.AddressBar.bClose = true;
					clearTimeout(Addons.AddressBar.tid2);
					Addons.AddressBar.tid2 = setTimeout("Addons.AddressBar.bClose = false;", 500);

				});
				return true;
			}
			return false;
		}
	};


	AddEvent("ChangeView", function (Ctrl)
	{
		if (Ctrl.FolderItem && Ctrl.Id == Ctrl.Parent.Selected.Id && Ctrl.Parent.Id == te.Ctrl(CTRL_TC).Id) {
			document.F.addressbar.value = api.GetDisplayNameOf(Ctrl.FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
			Addons.AddressBar.Arrange(Ctrl.FolderItem);
			document.getElementById("addr_img").src = GetIconImage(Ctrl, api.GetSysColor(COLOR_WINDOW));
			setTimeout("Addons.AddressBar.Blur()", 99);
		}
	});

	AddEvent("Resize", function ()
	{
		Addons.AddressBar.Arrange();
	});

	AddEvent("MouseMessage", function (Ctrl, hwnd, msg, mouseData, pt, wHitTestCode, dwExtraInfo)
	{
		if (msg == WM_MOUSEMOVE && Ctrl.Type == CTRL_TE && Addons.AddressBar.bLoop) {
			var Ctrl2 = te.CtrlFromPoint(pt);
			if (Ctrl2 && Ctrl2.Type == CTRL_WB && !HitTest(Addons.AddressBar.Item, pt)) {
				for (var i = Addons.AddressBar.nLevel; i >= 0; i--) {
					var o = document.getElementById("addressbar" + i);
					if (o) {
						if (HitTest(o, pt)) {
							api.PostMessage(hwnd, WM_KEYDOWN, VK_ESCAPE, 0);
							(function (o) { setTimeout(function () {
								Addons.AddressBar.bClose = false;
								o.click();
							}, 99);}) (o);
						}
					}
				}
			}
		}
	});

	AddEvent("SetAddress", function (s)
	{
		document.F.addressbar.value = s;
	});

	AddEvent("DragEnter", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
		if (Ctrl.Type == CTRL_WB) {
			return S_OK;
		}
	});

	AddEvent("DragOver", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
		if (Ctrl.Type == CTRL_WB && dataObj.Count) {
			var ptc = pt.Clone();
			api.ScreenToClient(api.GetWindow(document), ptc);
			var el = document.elementFromPoint(ptc.x, ptc.y);
			if (el) {
				var res = /^addressbar(\d+)_$/.exec(el.id);
				if (res) {
					var Target = Addons.AddressBar.GetPath(res[1] - 0);
					if (!api.ILIsEqual(dataObj.Item(-1), Target)) {
						var DropTarget = api.DropTarget(Target);
						if (DropTarget) {
							return DropTarget.DragOver(dataObj, grfKeyState, pt, pdwEffect);
						}
					}
					pdwEffect[0] = DROPEFFECT_NONE;
					return S_OK;
				}
			}
		}
	});

	AddEvent("Drop", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
		if (Ctrl.Type == CTRL_WB && dataObj.Count) {
			var ptc = pt.Clone();
			api.ScreenToClient(api.GetWindow(document), ptc);
			var el = document.elementFromPoint(ptc.x, ptc.y);
			if (el) {
				var res = /^addressbar(\d+)_$/.exec(el.id);
				if (res) {
					var hr = S_FALSE;
					var Target = Addons.AddressBar.GetPath(res[1] - 0);
					if (!api.ILIsEqual(dataObj.Item(-1), Target)) {
						var DropTarget = api.DropTarget(Target);
						if (DropTarget) {
							hr = DropTarget.Drop(dataObj, grfKeyState, pt, pdwEffect);
						}
					}
					return hr;
				}
			}
		}
	});

	AddEvent("DragLeave", function (Ctrl) {
		return S_OK;
	});

	GetAddress = function ()
	{
		return document.F.addressbar.value;
	}

	//Menu
	if (item.getAttribute("MenuExec")) {
		Addons.AddressBar.nPos = api.LowPart(item.getAttribute("MenuPos"));
		var s = item.getAttribute("MenuName");
		if (s && s != "") {
			Addons.AddressBar.strName = s;
		}
		AddEvent(item.getAttribute("Menu"), function (Ctrl, hMenu, nPos)
		{
			api.InsertMenu(hMenu, Addons.AddressBar.nPos, MF_BYPOSITION | MF_STRING, ++nPos, GetText(Addons.AddressBar.strName));
			ExtraMenuCommand[nPos] = Addons.AddressBar.Exec;
			return nPos;
		});
	}
	//Key
	if (item.getAttribute("KeyExec")) {
		SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Addons.AddressBar.Exec, "Func");
	}
	//Mouse
	if (item.getAttribute("MouseExec")) {
		SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Addons.AddressBar.Exec, "Func");
	}

	AddTypeEx("Add-ons", "Address Bar", Addons.AddressBar.Exec);

	var s = item.getAttribute("Width");
	if (s) {
		s = (api.QuadPart(s) == s) ? (s + "px") : s;
	} else {
		s = "100%";
	}
	var nSize = api.GetSystemMetrics(SM_CYSMICON);
	s = ['<div style="position: relative; overflow: hidden"><div id="breadcrumbbuttons" class="breadcrumb" style="position: absolute; left: 1px; top: 1px; padding-left: ', nSize + 4, 'px" onfocus="Addons.AddressBar.Focus()" onclick="return Addons.AddressBar.Exec();"></div><input id="addressbar" type="text" autocomplate="on" list="AddressList" onkeydown="return Addons.AddressBar.KeyDown(this)" onfocus="Addons.AddressBar.Focus()" onblur="Addons.AddressBar.Blur()" onresize="Addons.AddressBar.Resize()" oninput="AdjustAutocomplete(this.value)" style="width: ', s.replace(/;"<>/g, ''), '; vertical-align: middle; padding-left: ', nSize + 4, 'px; padding-right: 16px"><div class="breadcrumb"><div id="addressbarselect" class="button" style="position: absolute; top: 1px" onmouseover="MouseOver(this);" onmouseout="MouseOut()" onclick="Addons.AddressBar.Popup3(this)">', BUTTONS.dropdown,'</div></div>'];

	s.push('<img id="addr_img" src="folder:closed"');
	s.push(' onclick="return Addons.AddressBar.Exec();"');
	s.push(' oncontextmenu="Addons.AddressBar.Exec(); return false;"');
	s.push(' style="position: absolute; left: 4px; top: 1.5pt; width: ', nSize, 'px; height: ', nSize, 'px; z-index: 3; border: 0px"></div>');

	SetAddon(Addon_Id, Default, s, "middle");
	Addons.AddressBar.Resize();
} else {
	var ado = OpenAdodbFromTextFile("addons\\" + Addon_Id + "\\options.html");
	if (ado) {
		SetTabContents(0, "General", ado.ReadText(adReadAll));
		ado.Close();
	}
}
