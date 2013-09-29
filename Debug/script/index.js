//Tablacus Explorer

te.ClearEvents();
var g_Bar = "";
var Addon = 1;
var Addons = {"_stack": []};
var Init = false;
var OpenMode = SBSP_SAMEBROWSER;
var ExtraMenus = {};
var ExtraMenuCommand = {};

var GetAddress = null;
var ShowContextMenu = null;

var g_tidResize = null;
var xmlWindow = null;
var g_Panels = [];
var g_KeyCode = {};
var g_KeyState = [
	[0x1d0000, 0x2000],
	[0x2a0000, 0x1000],
	[0x380000, 0x4000],
	["Win",    0x8000],
	["Ctrl",   0x2000],
	["Shift",  0x1000],
	["Alt",    0x4000]
];
var g_dlgs = {};

PanelCreated = function (Ctrl)
{
	for (var i in eventTE.PanelCreated) {
		eventTE.PanelCreated[i](Ctrl);
	}
}

ChangeView = function (Ctrl)
{
	te.Data.bSaveConfig = true;
	ChangeTabName(Ctrl);
	if (Ctrl.Data.bSuspend) {
		Ctrl.Data.bSuspend = false;
		setTimeout(function () {
			if (Ctrl.Items && Ctrl.Items.Count == 0) {
				Ctrl.Refresh();
			}
		}, 500);
	}
	for (var i in eventTE.ChangeView) {
		eventTE.ChangeView[i](Ctrl);
	}
}

ChangeTabName = function (Ctrl)
{
	Ctrl.Title = GetTabName(Ctrl);
}

GetTabName = function (Ctrl)
{
	if (Ctrl.FolderItem) {
		for (var i in eventTE.GetTabName) {
			var s = eventTE.GetTabName[i](Ctrl);
			if (s) {
				return s;
			}
		}
		return Ctrl.FolderItem.Name;
	}
}

CloseView = function (Ctrl)
{
	for (var i in eventTE.CloseView) {
		var hr = eventTE.CloseView[i](Ctrl);
		if (isFinite(hr) && hr != S_OK) {
			return hr; 
		}
	}
	return S_OK; 
}

DeviceChanged = function (Ctrl)
{
	for (var i in eventTE.DeviceChanged) {
		eventTE.DeviceChanged[i](Ctrl);
	}
}

ListViewCreated = function (Ctrl)
{
	for (var i in eventTE.ListViewCreated) {
		eventTE.ListViewCreated[i](Ctrl);
	}
}

TabViewCreated = function (Ctrl)
{
	for (var i in eventTE.TabViewCreated) {
		eventTE.TabViewCreated[i](Ctrl);
	}
}

TreeViewCreated = function (Ctrl)
{
	for (var i in eventTE.TreeViewCreated) {
		eventTE.TreeViewCreated[i](Ctrl);
	}
}

SetAddrss = function (s)
{
	for (var i in eventTE.SetAddrss) {
		eventTE.SetAddrss[i](Ctrl);
	}
}

RestoreFromTray = function ()
{
	api.ShowWindow(te.hwnd, api.IsIconic(te.hwnd) ? SW_RESTORE : SW_SHOW);
	for (var i in eventTE.RestoreFromTray) {
		eventTE.RestoreFromTray[i]();
	}
}

Finalize = function ()
{
	for (var i in eventTE.Finalize) {
		eventTE.Finalize[i]();
	}
	SaveConfig();
}

SetGestureText = function (Ctrl, Text)
{
	for (var i in eventTE.SetGestureText) {
		if (isFinite(eventTE.SetGestureText[i](Ctrl, Text))) {
			return;
		}
	}
}

Lock = function (Ctrl, nIndex, turn)
{
	var FV = Ctrl[nIndex];
	if (FV) {
		if (turn) {
			FV.Data.Lock = !FV.Data.Lock;
		}
		for (var i in eventTE.Lock) {
			eventTE.Lock[i](Ctrl, nIndex, FV.Data.Lock);
		}
	}
}

LoadConfig = function ()
{
	var xml = OpenXml("window.xml", true, false);
	if (xml) {
		xmlWindow = xml;
		arKey = ["Conf", "Tab", "Tree", "View"]
		for (var j in arKey) {
			var key = arKey[j];
			var items = xml.getElementsByTagName(key);
			if (items.length == 0 && j == 0) {
				items = xml.getElementsByTagName('Config');
			}
			for (var i = 0; i < items.length; i++) {
				var item = items[i];
				var s = item.text;
				if (s == "0") {
					s = 0;
				}
				te.Data[key + "_" + item.getAttribute("Id")] = s;
			}
		}
		var items = xml.getElementsByTagName('Window');
		for (var i = 0; i < items.length; i++) {
			var item = items[i];
			var x = item.getAttribute("Left");
			var y = item.getAttribute("Top")
			if (x > -30000 && y > -30000) {
				api.MoveWindow(te.hwnd, x, y, item.getAttribute("Width"), item.getAttribute("Height"), 0);
				var nCmdShow = item.getAttribute("CmdShow");
				if (nCmdShow != SW_SHOWNORMAL) {
					api.ShowWindow(te.hwnd, nCmdShow);
				}
			}
		}
	}
	else {
		xmlWindow = "Init";
	}
}

SaveConfig = function ()
{
	if (te.Data.bSaveMenus) {
		te.Data.bSaveMenus = false;
		SaveXmlEx("menus.xml", te.Data.xmlMenus);
	}
	if (te.Data.bSaveAddons) {
		te.Data.bSaveAddons = false;
		SaveXmlEx("addons.xml", te.Data.Addons);
	}
	if (te.Data.bSaveConfig) {
		te.Data.bChanged = false;
		for (var i in eventTE.SaveConfig) {
			eventTE.SaveConfig[i]();
		}
		SaveXml(fso.BuildPath(te.Data.DataFolder, "config\\window.xml"), true);
	}
}

Resize = function ()
{
	if (!g_tidResize) {
		clearTimeout(g_tidResize);
		g_tidResize = setTimeout(Resize2, 500);
	}
}

Resize2 = function ()
{
	if (g_tidResize) {
		clearTimeout(g_tidResize);
		g_tidResize = null;
	}
	o = document.getElementById("toolbar");
	if (o) {
		te.offsetTop = o.offsetHeight;
	}
	else {
		te.offsetTop = 0;
	}

	var h = 0;
	o = document.getElementById("bottombar");
	if (o) {
		te.offsetBottom = o.offsetHeight;
		o = document.getElementById("client");
		if (o) {
			h = (document.documentElement ? document.documentElement.offsetHeight : document.body.offsetHeight) - te.offsetBottom - te.offsetTop;
			o.style.height = ((h >= 0) ? h : 0) + "px";
		}
	}
	o = document.getElementById("leftbarT");
	if (o) {
		var i = h;
		o.style.height = ((i >= 0) ? i : 0) + "px";
		i = o.clientHeight - o.style.height.replace(/\D/g, "");

		var h2 = o.clientHeight - document.getElementById("LeftBar1").offsetHeight - document.getElementById("LeftBar3").offsetHeight;
		o = document.getElementById("LeftBar2");
		o.style.height = h2 - i + "px";
	}
	o = document.getElementById("rightbarT");
	if (o) {
		var i = h;
		o.style.height = ((i >= 0) ? i : 0) + "px";
	}

	var w2 = 0;
	var w = 0;
	var o = te.Data.Locations;
	if (o.LeftBar1 || o.LeftBar2 || o.LeftBar3) {
		w = te.Data.Conf_LeftBarWidth;
	}
	var o = document.getElementById("leftbar");
	if (o) {
		w = (w > 0) ? w : 0;
		o.style.width = w + "px";
		for (var i = 1; i <= 3; i++) {
			var ob = document.getElementById("LeftBar" + i);
			if (ob && api.StrCmpI(ob.style.display, "none")) {
				pt = GetPos(ob);
				o.style.width = w + "px";
				w2 = w;
				break;
			}
		}
	}
	te.offsetLeft = w2;

	w2 = 0;
	w = 0;
	var o = te.Data.Locations;
	if (o.RightBar1 || o.RightBar2 || o.RightBar3) {
		w = te.Data.Conf_RightBarWidth;
	}
	var o = document.getElementById("rightbar");
	if (o) {
		w = (w > 0) ? w : 0;
		o.style.width = w + "px";
		for (var i = 1; i <= 3; i++) {
			var ob = document.getElementById("RightBar" + i);
			if (ob && api.StrCmpI(ob.style.display, "none")) {
				o.style.width = w + "px";
				w2 = w;
				break;
			}
		}
	}
	te.offsetRight = w2;
	o = document.getElementById("Background");
	if (o) {
		w = (document.documentElement ? document.documentElement.offsetWidth : document.body.offsetWidth) - te.offsetLeft - te.offsetRight;
		if (w > 0) {
			o.style.width = w + "px";
		}
	}

	for (var i in eventTE.Resize) {
		eventTE.Resize[i]();
	}
}

LoadLang = function (bAppend)
{
	if (!bAppend) {
		MainWindow.Lang = {};
	}
	var filename = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "lang\\" + GetLangId() + ".xml");
	LoadLang2(filename);
}

Refresh = function ()
{
	var FV = te.Ctrl(CTRL_FV);
	if (FV) {
		FV.Refresh();
	}
}

AddFavorite = function (FolderItem)
{
	var xml = te.Data.xmlMenus;
	var menus = xml.getElementsByTagName("Favorites");
	if (menus && menus.length) {
		var item = xml.createElement("Item");
		if (!FolderItem) {
			var FV = te.Ctrl(CTRL_FV);
			if (FV) {
				FolderItem = FV.FolderItem;
			}
		}
		if (!FolderItem) {
			return false;
		}
		var s = InputDialog(GetText("Add Favorite"), api.GetDisplayNameOf(FolderItem, SHGDN_INFOLDER));
		if (s) {
			item.setAttribute("Name", s);
			item.setAttribute("Filter", "");
			item.text = api.GetDisplayNameOf(FolderItem, SHGDN_FORPARSINGEX | SHGDN_FORPARSING);
			if (fso.FileExists(item.text)) {
				item.text = api.PathQuoteSpaces(item.text);
				item.setAttribute("Type", "Exec");
			}
			else {
				item.setAttribute("Type", "Open");
			}
			menus[0].appendChild(item);
			SaveXmlEx("menus.xml", xml);
			return true;
		}
	}
	return false;
}

CreateNewFolder = function (Ctrl, pt)
{
	var path = InputDialog(GetText("New Folder"), "");
	if (path) {
		if (!path.match(/^[A-Z]:\\|^\\/i)) {
			var FV = GetFolderView(Ctrl, pt);
			path = fso.BuildPath(FV.FolderItem.Path, path);
		}
		CreateFolder(path);
	}
}

CreateNewFile = function (Ctrl, pt)
{
	var path = InputDialog(GetText("New File"), "");
	if (path) {
		if (!path.match(/^[A-Z]:\\|^\\/i)) {
			var FV = GetFolderView(Ctrl, pt);
			path = fso.BuildPath(FV.FolderItem.Path, path);
		}
		CreateFile(path);
	}
}

GetCommandId = function (hMenu, s)
{
	var wId = 0;
	if (s) {
		var mii = api.Memory("MENUITEMINFO");
		mii.cbSize = mii.Size;
		mii.fMask = MIIM_SUBMENU | MIIM_ID;
		for (var i = api.GetMenuItemCount(hMenu); i--;) {
			var title = api.GetMenuString(hMenu, i, MF_BYPOSITION);
			if (title) {
				api.GetMenuItemInfo(hMenu, i, true, mii);
				if (mii.hSubMenu) {
					wId = GetCommandId(mii.hSubMenu, s);
					if (wId) {
						break;
					}
				}
				else {
					var a = title.split(/\t/);
					if (api.PathMatchSpec(a[0], s) || api.PathMatchSpec(a[1], s)) {
						wId = mii.wId;
						break;
					}
				}
			}
		}
	}
	api.DestroyMenu(hMenu);
	return wId;
};

OpenSelected = function (Ctrl, NewTab)
{
	if (Ctrl.Type <= CTRL_EB) {
		var Exec = [];
		var Selected = Ctrl.SelectedItems();
		for (var i in Selected) {
			var Item = Selected[i];
			var path = Item.Path;
			if (Item.IsFolder) {
			 	Ctrl.Navigate(Item, NewTab);
			 	NewTab |= SBSP_NEWBROWSER;
			}
			else if (api.PathMatchSpec(path, "*.lnk")) {
				var lnk = wsh.CreateShortcut(path)
				path = lnk.TargetPath;
				if (fso.FolderExists(path)) {
					Ctrl.Navigate(path, NewTab);
					NewTab |= SBSP_NEWBROWSER;
				}
				else {
					Exec.push(Item);
				}
			}
			else {
				Exec.push(Item);
			}
		}
		if (Exec.length) {
			if (Selected.Count != Exec.length) {
				Selected = te.FolderItems();
				for (var i = 0; i < Exec.length; i++) {
					Selected.AddItem(Exec[i]);
				}
			}
			var ContextMenu = api.ContextMenu(Selected);
			if (ContextMenu) {
				var hMenu = api.CreatePopupMenu();
				ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_NORMAL);
				var nCommand = api.GetMenuDefaultItem(hMenu, MF_BYCOMMAND, GMDI_USEDISABLED);
				ContextMenu.InvokeCommand(0, external.hwnd, nCommand - 1, null, null, SW_SHOWNORMAL, 0, 0);
				api.DestroyMenu(hMenu);
			}

		}
	}
	return S_OK;
}

SendCommand = function (Ctrl, nVerb)
{
	if (nVerb) {
		var hwnd = Ctrl.hwndView;
		if (!hwnd) {
			var FV = te.Ctrl(CTRL_FV);
			if (FV) {
				hwnd = FV.hwndView;
			}
		}
		api.SendMessage(hwnd, WM_COMMAND, nVerb, 0);
	}
}

IncludeObject = function (FV, Item)
{
	if (api.PathMatchSpec(Item.Name, FV.Data.Filter)) {
		return S_OK;
	}
	return S_FALSE;
}

EnableDragDrop = function ()
{
}

DisableImage = function (img, bDisable)
{
	if (img) {
		if (api.QuadPart(document.documentMode) < 10) {
			img.style.filter = bDisable ? "gray(); alpha(style=0,opacity=48);": "";
		}
		else {
			var s = img.src;
			if (bDisable) {
				if (s.match(/^data:image\/png/i)) {
					img.src = "data:image/svg+xml," + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ' + img.offsetWidth + ' ' + img.offsetHeight + '"><filter id="G"><feColorMatrix type="saturate" values="0.1" /></filter><image width="' + img.width + '" height="' + img.height + '" xlink:href="' + img.src + '" filter="url(#G)" opacity=".48"></image></svg>');
				}
			}
			else if (s.match(/^data:image\/svg/i) && decodeURIComponent(s).match(/href="([^"]*)/i)) {
				img.src = RegExp.$1
			}
		}
	}
}

//Events

te.OnCreate = function (Ctrl)
{
	if (!Ctrl.Data && Ctrl.Type != CTRL_TE) {
		Ctrl.Data = te.Object();
	}
	for (var i in eventTE.Create) {
		eventTE.Create[i](Ctrl);
	}
	if (Ctrl.Type == CTRL_TE) {
		RunCommandLine(api.GetCommandLine());
	}
}

te.OnClose = function (Ctrl)
{
	if (Ctrl.Type == CTRL_TC) {
		var o = document.getElementById("Panel_" + Ctrl.Id);
		if (o) {
			o.style.display = "none";
		}
	}
	for (var i in eventTE.Close) {
		var hr = eventTE.Close[i](Ctrl);
		if (isFinite(hr) && hr != S_OK) {
			return hr; 
		}
	}
	return S_OK;
}

AddEvent("Close", function (Ctrl)
{
	switch (Ctrl.Type) {
		case CTRL_TE:
			Finalize();
			if (api.GetThreadCount() && wsh.Popup(GetText("File is in operation."), 0, TITLE, MB_ICONSTOP | MB_ABORTRETRYIGNORE) != IDIGNORE) {
				return S_FALSE;
			}
			break;
		case CTRL_WB:
			break;
		case CTRL_SB:
		case CTRL_EB:
			if (Ctrl.Data.Lock) {
				return S_FALSE;
			}
			return CloseView(Ctrl);
		case CTRL_TC:
			break;
	}
});

te.OnViewCreated = function (Ctrl)
{
	switch (Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
			ListViewCreated(Ctrl);
			ChangeView(Ctrl);
			break;
		case CTRL_TC:
			TabViewCreated(Ctrl);
			break;
		case CTRL_TV:
			TreeViewCreated(Ctrl);
			break;
	}
	for (var i in eventTE.ViewCreated) {
		eventTE.ViewCreated[i](Ctrl);
	}
}

te.OnBeforeNavigate = function (Ctrl, fs, wFlags, Prev)
{
	if (Ctrl.Data.Lock && (wFlags & SBSP_NEWBROWSER) == 0) {
		return E_ACCESSDENIED;
	}
	Ctrl.OnIncludeObject = null;
	Ctrl.FilterView = null;

	for (var i in eventTE.BeforeNavigate) {
		var hr = eventTE.BeforeNavigate[i](Ctrl, fs, wFlags, Prev);
		if (isFinite(hr) && hr != S_OK) {
			return hr; 
		}
	}
	return S_OK;
}

te.OnStatusText = function (Ctrl, Text, iPart)
{
	for (var i in eventTE.StatusText) {
		eventTE.StatusText[i](Ctrl, Text, iPart);
	}
	return S_OK; 
}

te.OnKeyMessage = function (Ctrl, hwnd, msg, key, keydata)
{
	for (var i in eventTE.KeyMessage) {
		var hr = eventTE.KeyMessage[i](Ctrl, hwnd, msg, key, keydata);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN) {
		var nKey = ((keydata >> 16) & 0x17f) | GetKeyShift();
		te.Data.cmdKey = nKey;
		switch (Ctrl.Type) {
			case CTRL_SB:
			case CTRL_EB:
				var strClass = api.GetClassName(hwnd);
				if (api.strcmpi(strClass, WC_LISTVIEW) == 0 || api.strcmpi(strClass, "DirectUIHWND") == 0) {
					var cmd = eventTE.Key.List[nKey];
					if (cmd) {
						return Exec(Ctrl, cmd[0], cmd[1], hwnd, null);
					}
				}
				break;
			case CTRL_TV:
				var strClass = api.GetClassName(hwnd);
				if (api.strcmpi(strClass, WC_TREEVIEW) == 0) {
					var cmd = eventTE.Key.Tree[nKey];
					if (cmd) {
						return Exec(Ctrl, cmd[0], cmd[1], hwnd, null);
					}
				}
				break;
			case CTRL_WB:
				var cmd = eventTE.Key.Browser[nKey];
				if (cmd) {
					return Exec(Ctrl, cmd[0], cmd[1], hwnd, null);
				}
				break;
			default:
				if (window.g_menu_click) {
					if (key == VK_RETURN) {
						var hSubMenu = api.GetSubMenu(window.g_menu_handle, window.g_menu_pos);
						if (hSubMenu) {
							var mii = api.Memory("MENUITEMINFO");
							mii.cbSize = mii.Size;
							mii.fMask = MIIM_SUBMENU;
							api.SetMenuItemInfo(window.g_menu_handle, window.g_menu_pos, true, mii);
							api.DestroyMenu(hSubMenu);
							api.PostMessage(hwnd, WM_CHAR, VK_LBUTTON, 0);
						}
						window.g_menu_button = api.GetKeyState(VK_SHIFT) >= 0 ? 1 : 2;
						if (api.GetKeyState(VK_CONTROL) < 0) {
							window.g_menu_button = 3;
						}
					}
				}
				break;
		}
		if (Ctrl.Type != CTRL_TE) {
			var cmd = eventTE.Key.All[nKey];
			if (cmd) {
				return Exec(Ctrl, cmd[0], cmd[1], hwnd, null);
			}
		}
	}
	return S_FALSE; 
}

te.OnMouseMessage = function (Ctrl, hwnd, msg, mouseData, pt, wHitTestCode, dwExtraInfo)
{
	for (var i in eventTE.MouseMessage) {
		var hr = eventTE.MouseMessage[i](Ctrl, hwnd, msg, mouseData, pt, wHitTestCode, dwExtraInfo);
		if (isFinite(hr)) {
			return hr; 
		}
	}

	if (msg == WM_MOUSEWHEEL) {
		var Ctrl2 = te.CtrlFromPoint(pt);
		if (Ctrl2) {
			g_mouse.str = GetGestureKey() + GetGestureButton() + (mouseData > 0 ? "8" : "9");
			if (api.GetKeyState(VK_RBUTTON) < 0) {
				g_mouse.CancelButton = true;
			}
			var hr = g_mouse.Exec(Ctrl2, hwnd, pt);
			g_mouse.EndGesture(false);
			if (hr == S_OK) {
				return hr;
			}
			if (Ctrl2.Type == CTRL_TC) {
				ChangeTab(Ctrl2, mouseData > 0 ? -1 : 1)
				return S_OK;
			}
			var hwnd2 = api.WindowFromPoint(pt);
			if (hwnd2 && hwnd != hwnd2) {
				api.SetFocus(hwnd2);
				api.SendMessage(hwnd2, msg, mouseData & 0xffff0000, pt.x + (pt.y << 16));
				return S_OK;
			}
		}
	}
	if (msg == WM_LBUTTONUP || msg == WM_RBUTTONUP || msg == WM_MBUTTONUP || msg == WM_XBUTTONUP) {
		if (g_mouse.GetButton(msg, mouseData) == g_mouse.str.charAt(0)) {
			var hr = S_FALSE;
			var bButton = false;
			if (g_mouse.RButton >= 0 && msg == WM_RBUTTONUP) {
				g_mouse.RButtonDown(true);
				bButton = api.strcmpi(g_mouse.str, "2") == 0;
			}
			if (g_mouse.str.length >= 2 || !IsDrag(pt, te.Data.pt)) {
				hr = g_mouse.Exec(te.CtrlFromWindow(g_mouse.hwndGesture), g_mouse.hwndGesture, pt);
			}
			var s = g_mouse.str;
			g_mouse.EndGesture(false);
			if (g_mouse.bCapture) {
				api.ReleaseCapture();
				g_mouse.bCapture = false;
			}
			if (g_mouse.CancelButton || hr == S_OK) {
				g_mouse.CancelButton = false;
				return S_OK;
			}
			if (bButton) {
				api.PostMessage(g_mouse.hwndGesture, WM_CONTEXTMENU, g_mouse.hwndGesture, pt.x + (pt.y << 16));
				return S_OK;
			}
		}
	}
	if (msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN || msg == WM_MBUTTONDOWN || msg == WM_XBUTTONDOWN) {
		clearTimeout(g_mouse.tidDblClk);
		if (g_mouse.tidDblClk && msg == WM_LBUTTONDOWN && Ctrl.Type == CTRL_EB) {
			g_mouse.tidDblClk = false;
			msg = WM_LBUTTONDBLCLK;
		}
		if (g_mouse.str.length == 0) {
			te.Data.pt = pt;
			g_mouse.ptGesture.x = pt.x;
			g_mouse.ptGesture.y = pt.y;
			g_mouse.hwndGesture = hwnd;
			if (msg == WM_LBUTTONDOWN && Ctrl.Type == CTRL_EB) {
				g_mouse.tidDblClk = true;
				g_mouse.tidDblClk = setTimeout("g_mouse.tidDblClk = false", sha.GetSystemInformation("DoubleClickTime"));
			}
		}
		g_mouse.str += g_mouse.GetButton(msg, mouseData);
		if (g_mouse.str.length > 1) {
			g_mouse.StartGestureTimer();
		}
		SetGestureText(Ctrl, GetGestureKey() + g_mouse.str);
		if (msg == WM_RBUTTONDOWN) {
			if (te.Data.Conf_Gestures >= 2) {
				var iItem = -1;
				if (Ctrl.Type == CTRL_SB || Ctrl.Type == CTRL_EB) {
					iItem = Ctrl.HitTest(pt, LVHT_ONITEM);
					if (iItem < 0) {
						return S_OK;
					}
				}
				if (te.Data.Conf_Gestures == 3) {
					g_mouse.RButton = iItem;
					g_mouse.StartGestureTimer();
					return S_OK;
				}
			}
		}
	}
	if (msg == WM_LBUTTONDBLCLK || msg == WM_RBUTTONDBLCLK || msg == WM_MBUTTONDBLCLK || msg == WM_XBUTTONDBLCLK) {
		var strClass = api.GetClassName(hwnd);
		if (api.strcmpi(strClass, WC_HEADER)) {
			te.Data.pt = pt;
			g_mouse.str = g_mouse.GetButton(msg, mouseData);
			g_mouse.str += g_mouse.str;
			var hr = g_mouse.Exec(Ctrl, hwnd, pt);
			g_mouse.EndGesture(false);
			if (hr == S_OK) {
				return hr;
			}
		}
	}

	if (msg == WM_MOUSEMOVE) {
		if (g_mouse.str.length && (te.Data.Conf_Gestures > 1 && api.GetKeyState(VK_RBUTTON) < 0) || (te.Data.Conf_Gestures && (api.GetKeyState(VK_MBUTTON) < 0 || api.GetKeyState(VK_XBUTTON1) < 0 || api.GetKeyState(VK_XBUTTON2) < 0))) {
			var x = (pt.x - g_mouse.ptGesture.x);
			var y = (pt.y - g_mouse.ptGesture.y);
			if (Math.abs(x) + Math.abs(y) >= 20) {
				if (te.Data.Conf_TrailSize) {
					var hdc = api.GetWindowDC(te.hwnd);
					if (hdc) {
						var rc = api.Memory("RECT");
						api.GetWindowRect(te.hwnd, rc);
						api.MoveToEx(hdc, g_mouse.ptGesture.x - rc.left, g_mouse.ptGesture.y - rc.top, null);
						var pen1 = api.CreatePen(PS_SOLID, te.Data.Conf_TrailSize, te.Data.Conf_TrailColor);
						var hOld = api.SelectObject(hdc, pen1);
						api.LineTo(hdc, pt.x - rc.left, pt.y - rc.top);
						api.SelectObject(hdc, hOld);
						api.DeleteObject(pen1);
						g_mouse.bTrail = true;
						api.ReleaseDC(te.hwnd, hdc);
					}
				}
				g_mouse.ptGesture.x = pt.x;
				g_mouse.ptGesture.y = pt.y;
				var s = (Math.abs(x) >= Math.abs(y)) ? ((x < 0) ? "L" : "R") :  ((y < 0) ? "U" : "D");

				if (s != g_mouse.str.charAt(g_mouse.str.length - 1)) {
					g_mouse.str += s;
					SetGestureText(Ctrl, GetGestureKey() + g_mouse.str);
				}
				if (!g_mouse.bCapture) {
					api.SetCapture(g_mouse.hwndGesture);
					g_mouse.bCapture = true;
				}
				g_mouse.StartGestureTimer();
			}
		}
	}
	return g_mouse.str.length >= 2 ? S_OK : S_FALSE;
}

te.OnCommand = function (Ctrl, hwnd, msg, wParam, lParam)
{
	te.Data.bSaveConfig = true;
	for (var i in eventTE.Command) {
		var hr = eventTE.Command[i](Ctrl, hwnd, msg, wParam, lParam);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	return S_FALSE; 
}

te.OnInvokeCommand = function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon)
{
	te.Data.bSaveConfig = true;
	for (var i in eventTE.InvokeCommand) {
		var hr = eventTE.InvokeCommand[i](ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	return S_FALSE; 
}

te.OnDragEnter = function (Ctrl, dataObj, grfKeyState, pt, pdwEffect)
{
	var hr = E_NOTIMPL;
	var dwEffect = pdwEffect[0];
	for (var i in eventTE.DragEnter) {
		pdwEffect[0] = dwEffect;
		var hr2 = eventTE.DragEnter[i](Ctrl, dataObj, grfKeyState, pt, pdwEffect);
		if (isFinite(hr2) && hr != S_OK) {
			hr = hr2;
		}
	}
	g_mouse.str = "";
	return hr; 
}

te.OnDragOver = function (Ctrl, dataObj, grfKeyState, pt, pdwEffect)
{
	var dwEffect = pdwEffect[0];
	for (var i in eventTE.DragOver) {
		pdwEffect[0] = dwEffect;
		var hr = eventTE.DragOver[i](Ctrl, dataObj, grfKeyState, pt, pdwEffect);
		if (isFinite(hr)) {
			return hr;
			break; 
		}
	}
	return E_NOTIMPL; 
}

te.OnDrop = function (Ctrl, dataObj, grfKeyState, pt, pdwEffect)
{
	var dwEffect = pdwEffect[0];
	for (var i in eventTE.Drop) {
		pdwEffect[0] = dwEffect;
		hr = eventTE.Drop[i](Ctrl, dataObj, grfKeyState, pt, pdwEffect);
		if (isFinite(hr)) {
			return hr;
		}
	}
	return E_NOTIMPL; 
}

te.OnDragleave = function (Ctrl, dataObj, grfKeyState, pt, pdwEffect)
{
	var hr = E_NOTIMPL;
	for (var i in eventTE.Dragleave) {
		var hr2 = eventTE.Dragleave[i](Ctrl, dataObj, grfKeyState, pt, pdwEffect);
		if (isFinite(hr2) && hr != S_OK) {
			hr = hr2;
		}
	}
	g_mouse.str = "";
	return hr; 
}

te.OnSelectionChanging = function (Ctrl)
{
	for (var i in eventTE.SelectionChanging) {
		var hr = eventTE.SelectionChanging[i](Ctrl);
		if (isFinite(hr)) {
			return hr;
		}
	}
	return S_OK;
}

te.OnSelectionChanged = function (Ctrl, uChange)
{
	if (Ctrl.Type == CTRL_TC && Ctrl.SelectedIndex >= 0) {
		ChangeView(Ctrl.Selected);
	}
	for (var i in eventTE.SelectionChanged) {
		var hr = eventTE.SelectionChanged[i](Ctrl, uChange);
		if (isFinite(hr)) {
			return hr;
		}
	}
	return S_OK;
}

te.OnShowContextMenu = function (Ctrl, hwnd, msg, wParam, pt)
{
	for (var i in eventTE.ShowContextMenu) {
		var hr = eventTE.ShowContextMenu[i](Ctrl, hwnd, msg, wParam, pt);
		if (isFinite(hr)) {
			return hr; 
		}
	}

	switch (Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
			if (Ctrl.SelectedItems.Count) {
				if (ExecMenu(Ctrl, "Context", pt, 1) == S_OK) {
					return S_OK;
				}
			}
			else {
				if (ExecMenu(Ctrl, "ViewContext", pt, 1) == S_OK) {
					return S_OK;
				}
			}
			break;
		case CTRL_TV:
			if (ExecMenu(Ctrl, "Tree", pt, 1) == S_OK) {
				return S_OK;
			}
			break;
		case CTRL_TC:
			if (ExecMenu(Ctrl, "Tabs", pt, 1) == S_OK) {
				return S_OK;
			}
			break;
		case CTRL_WB:
			if (ShowContextMenu) {
				return ShowContextMenu(Ctrl, hwnd, msg, wParam, pt);
			}
			if (wParam == CONTEXT_MENU_DEFAULT) {
				return S_OK;
			}
			break;
		case CTRL_TE:
			api.GetSystemMenu(te.hwnd, true);
			var hMenu = api.GetSystemMenu(te.hwnd, false);
			var menus = te.Data.xmlMenus.getElementsByTagName("System");
			if (menus && menus.length) {
				var items = menus[0].getElementsByTagName("Item");
				var arMenu = OpenMenu(items, null);
				MakeMenus(hMenu, menus, arMenu, items);
			}
			break;
	}
	return S_FALSE;
}

te.OnDefaultCommand = function (Ctrl)
{
	for (var i in eventTE.DefaultCommand) {
		var hr = eventTE.DefaultCommand[i](Ctrl);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	return ExecMenu(Ctrl, "Default", null, 2);
}

te.OnItemClick = function (Ctrl, Item, HitTest, Flags)
{
	for (var i in eventTE.ItemClick) {
		var hr = eventTE.ItemClick[i](Ctrl, Item, HitTest, Flags);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	return S_FALSE;
}

te.OnSystemMessage = function (Ctrl, hwnd, msg, wParam, lParam)
{
	if (msg == WM_KILLFOCUS && Ctrl.Type == CTRL_TE) {
		g_mouse.str = "";
		SetGestureText(Ctrl, "");
	}
	for (var i in eventTE.SystemMessage) {
		var hr = eventTE.SystemMessage[i](Ctrl, hwnd, msg, wParam, lParam);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	switch (Ctrl.Type) {
		case CTRL_WB:
			if (msg == WM_KILLFOCUS) {
				try {
					document.selection.empty();
				} catch (e) {}
			}
			break;
		case CTRL_TE:
			switch (msg) {
				case WM_DESTROY:
					api.ImageList_Destroy(Ctrl.Data.himlTC);
					break;
				case WM_DEVICECHANGE:
					if (wParam == DBT_DEVICEARRIVAL || wParam == DBT_DEVICEREMOVECOMPLETE) {
						DeviceChanged();
					}
					break;
				case WM_SETFOCUS:
					var FV = te.Ctrl(CTRL_FV);
					if (FV) {
						api.SetFocus(FV.hwnd);
					}
					break;
				case WM_ACTIVATE:
					if (te.Data.bSaveMenus) {
						te.Data.bSaveMenus = false;
						SaveXmlEx("menus.xml", te.Data.xmlMenus);
					}
					if (te.Data.bSaveAddons) {
						te.Data.bSaveAddons = false;
						SaveXmlEx("addons.xml", te.Data.Addons);
					}
					if (te.Data.bReload) {
						te.Data.bReload = false;
						te.Reload();
					}
					break;
				case WM_SYSCOMMAND:
					if (wParam < 0xf000) {
						var menus = te.Data.xmlMenus.getElementsByTagName("System");
						if (menus && menus.length) {
							var items = menus[0].getElementsByTagName("Item");
							if (items) {
								var item = items[wParam - 1];
								if (item) {
									Exec(Ctrl, item.text, item.getAttribute("Type"), null);
									return 1;
								}
							}
						}
					}
					break;
				case WM_COPYDATA:
					var cd = api.Memory("COPYDATASTRUCT", 1, lParam);
					for (var i in eventTE.CopyData) {
						var hr = eventTE.CopyData[i](Ctrl, cd, wParam);
						if (isFinite(hr)) {
							return hr; 
						}
					}
					if (cd.dwData == 0 && cd.cbData) {
						var strData = api.SysAllocStringByteLen(cd.lpData, cd.cbData);
						RestoreFromTray();
						RunCommandLine(strData);
						return S_OK;
					}
					break;
			}
			break;
		case CTRL_TC:
			if (msg == WM_SHOWWINDOW && wParam == 0) {
				var o = document.getElementById("Panel_" + Ctrl.Id);
				if (o) {
					o.style.display = "none";
				}
			}
			break;
	}
	if (msg == WM_POWERBROADCAST && wParam == 18) {
		setTimeout(function ()
		{
			var cFV = te.Ctrls(CTRL_FV);
			for (var i in cFV) {
				if (api.PathIsNetworkPath(cFV[i].Path)) {
					if (cFV[i].Items && cFV[i].Items.Count == 0) {
						cFV[i].Refresh();
					}
					else {
						cFV[i].Data.bSuspend = true;
					}
				}
			}
		}, 500);
	}
	return 0; 
};

te.OnMenuMessage = function (Ctrl, hwnd, msg, wParam, lParam)
{
	for (var i in eventTE.MenuMessage) {
		var hr = eventTE.MenuMessage[i](Ctrl, hwnd, msg, wParam, lParam);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	switch (msg) {
		case WM_MENUSELECT:
			if ((wParam >> 16) & MF_POPUP) {
				var hSubMenu = api.GetSubMenu(lParam, (wParam & 0xffff));
				if (hSubMenu) {
					var s = api.GetMenuString(hSubMenu, 0, MF_BYPOSITION);
					if (s && s.match(/^\tJScript\t|\tVBScript\t/i)) {
						var ar = s.split("\t");
						api.DeleteMenu(hSubMenu, 0, MF_BYPOSITION);
						ExecScriptEx(Ctrl, ar[2], ar[1], hwnd);
					}
					window.g_menu_handle = lParam;
					window.g_menu_pos = (wParam & 0xffff);
				}
			}
			break;
		case WM_EXITMENULOOP:
			window.g_menu_click = false;
			while (eventTE.ExitMenuLoop && eventTE.ExitMenuLoop.length) {
				eventTE.ExitMenuLoop.shift()();
			}
			break;
		case WM_MENUCHAR:
			if (window.g_menu_click && (wParam & 0xffff) == VK_LBUTTON) {
				return MNC_EXECUTE << 16 + window.g_menu_pos;
			}
			break;
	}
	return S_FALSE; 
};

te.OnAppMessage = function (Ctrl, hwnd, msg, wParam, lParam)
{
	for (var i in eventTE.AppMessage) {
		var hr = eventTE.AppMessage[i](Ctrl, hwnd, msg, wParam, lParam);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	return S_FALSE;
}

te.OnNewWindow = function (Ctrl, dwFlags, UrlContext, Url)
{
	for (var i in eventTE.NewWindow) {
		var hr = eventTE.NewWindow[i](Ctrl, dwFlags, UrlContext, Url);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	var Path = api.PathCreateFromUrl(Url);
	var FolderItem = api.ILCreateFromPath(Path);
	if (FolderItem.IsFolder) {
		Navigate(FolderItem, SBSP_NEWBROWSER);
		return S_OK;
	}
	return S_FALSE;
}

te.OnClipboardText = function (Items)
{
	for (var i in eventTE.ClipboardText) {
		var r = eventTE.ClipboardText[i](Items);
		if (r !== undefined) {
			return r; 
		}
	}
	var s = [];
	for (var i = Items.Count; i > 0; s.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(Items.Item(--i), SHGDN_FORPARSING)))) {
	}
	return s.join(" ");
}

te.OnArrange = function (Ctrl, rc)
{
	if (Ctrl.Type == CTRL_TC) {
		var o = g_Panels[Ctrl.Id];
		if (!o) {
			var s = ['<table id="Panel_$" class="layout" style="position: absolute; z-index: 1;">'];
			s.push('<tr><td id="InnerLeft_$" class="sidebar" style="width: 0px; display: none"></td><td><div id="InnerTop_$" style="display: none"></div>');
			s.push('<table id="InnerTop2_$" class="layout" style="width: 100%">');
			s.push('<tr><td id="Inner1Left_$" class="toolbar1"></td><td id="Inner1Center_$" class="toolbar2" style="white-space: nowrap;"></td><td id="Inner1Right_$" class="toolbar3"></td></tr></table>');
			s.push('<table id="InnerView_$" class="layout" style="width: 100%"><tr><td id="Inner2Left_$" style="width: 0px"></td><td id="Inner2Center_$"></td><td id="Inner2Right_$" style="width: 0px"></td></tr></table>');
			s.push('<div id="InnerBottom_$"></div></td><td id="InnerRight_$" class="sidebar" style="width: 0px; display: none"></td></tr></table>');
			document.getElementById("Panel").insertAdjacentHTML("BeforeEnd", s.join("").replace(/\$/g, Ctrl.Id));
			PanelCreated(Ctrl);
			o = document.getElementById("Panel_" + Ctrl.Id);
			g_Panels[Ctrl.Id] = o;
			ApplyLang(o);
		}
		o.style.left = rc.Left + "px";
		o.style.top = rc.Top + "px";
		if (Ctrl.Visible) {
			o.style.display = !document.documentMode || api.strcmpi(o.tagName, "td") ? "block" : "table-cell";
		}
		else {
			o.style.display = "none";
		}
		var i = rc.Right - rc.Left
		o.style.width = (i >=0 ? i : 0) + "px";
		i = rc.Bottom - rc.Top;
		o.style.height = (i >=0 ? i : 0) + "px";
		rc.Top += document.getElementById("InnerTop_" + Ctrl.Id).offsetHeight + document.getElementById("InnerTop2_" + Ctrl.Id).offsetHeight;
		var w1 = 0;
		var w2 = 0;
		var x = '';
		for (i = 0; i <= 1; i++) {
			w1 += api.QuadPart(document.getElementById("Inner" + x + "Left_" + Ctrl.Id).style.width.replace(/\D/g, ""));
			w2 += api.QuadPart(document.getElementById("Inner" + x + "Right_" + Ctrl.Id).style.width.replace(/\D/g, ""));
			x = '2';
		}
		rc.Left += w1;
		rc.Right -= w2;
		rc.Bottom -= document.getElementById("InnerBottom_" + Ctrl.Id).offsetHeight;
		o = document.getElementById("Inner2Center_" + Ctrl.Id).style;
		try {
			o.width = (rc.Right - rc.Left) + "px";
			o.height = (rc.Bottom - rc.Top) + "px";
		} catch (e) {}
	}
	for (var i in eventTE.Arrange) {
		eventTE.Arrange[i](Ctrl, rc);
	}
}

te.OnVisibleChanged = function (Ctrl)
{
	if (Ctrl.Type == CTRL_TC) {
		var o = g_Panels[Ctrl.Id];
		if (o) {
			if (Ctrl.Visible) {
				o.style.display = !document.documentMode || api.strcmpi(o.tagName, "td") ? "block" : "table-cell";
			}
			else {
				o.style.display = "none";
			}
		}
		if (Ctrl.Visible && Ctrl.SelectedIndex >= 0) {
			ChangeView(Ctrl.Selected);
			for (var i = Ctrl.Count; i--;) {
				ChangeTabName(Ctrl.Item(i))
			}
		}
	}
	for (var i in eventTE.VisibleChanged) {
		eventTE.VisibleChanged[i](Ctrl);
	}
}

te.OnWindowRegistered = function (Ctrl)
{
	for (var i in eventTE.WindowRegistered) {
		eventTE.WindowRegistered[i](Ctrl);
	}
}

te.OnToolTip = function (Ctrl, Index)
{
	for (var i in eventTE.ToolTip) {
		var r = eventTE.ToolTip[i](Ctrl, Index);
		if (r !== undefined) {
			return r; 
		}
	}
	return "";
}

te.OnHitTest = function (Ctrl, pt, flags)
{
	for (var i in eventTE.HitTest) {
		var hr = eventTE.HitTest[i](Ctrl, pt, flags);
		if (isFinite(hr)) {
			return hr;
		}
	}
	return -1;
}

te.OnGetPaneState = function (Ctrl, ep, peps)
{
	for (var i in eventTE.GetPaneState) {
		var hr = eventTE.GetPaneState[i](Ctrl, ep, peps);
		if (isFinite(hr)) {
			return hr;
		}
	}
	return -1;
}

// Browser Events

AddEventEx(window, "load", function ()
{
	ApplyLang(document);

	setTimeout(function ()
	{
		if (api.strcmpi(typeof(xmlWindow), "string") == 0) {
			te.LockUpdate();
			var TC = te.CreateCtrl(CTRL_TC, 0, 0, "100%", "100%", te.Data.Tab_Style, te.Data.Tab_Align, te.Data.Tab_TabWidth, te.Data.Tab_TabHeight);
			TC.Selected.Navigate2(HOME_PATH, SBSP_NEWBROWSER, te.Data.View_Type, te.Data.View_ViewMode, te.Data.View_fFlags, te.Data.View_Options, te.Data.View_ViewFlags, te.Data.View_IconSize, te.Data.Tree_Align, te.Data.Tree_Width, te.Data.Tree_Style, te.Data.Tree_EnumFlags, te.Data.Tree_RootStyle, te.Data.Tree_Root);
			te.UnlockUpdate();
		}
		else if (xmlWindow) {
			LoadXml(xmlWindow);
			xmlWindow = null;
		}
		var a, i, j, root, menus, items;
		root = te.Data.xmlMenus.documentElement;
		if (root) {
			menus = root.childNodes;
			for (i = menus.length; i--;) {
				items = menus[i].getElementsByTagName("Item");
				for (j = items.length; j--;) {
					a = items[j].getAttribute("Name").split(/\\t/);
					if (a.length > 1) {
						SetKeyExec("List", a[1], items[j].text, items[j].getAttribute("Type"), true);
					}
				}
			}
		}
		DeviceChanged();
		Resize();
		setTimeout(function ()
		{
			var cTC = te.Ctrls(CTRL_TC);
			for (var i in cTC) {
				if (cTC[i].SelectedIndex >= 0) {
					ChangeView(cTC[i].Selected);
				}
			}
		}, 500);
	}, 1);
	if (api.GetKeyState(VK_SHIFT) < 0 && api.GetKeyState(VK_CONTROL) < 0) {
		ShowOptions("Tab=Add-ons");
	}
});

AddEventEx(window, "resize", Resize);

AddEventEx(window, "beforeunload", Finalize);

document.body.onselectstart = function (e)
{
	var s = (e || event).srcElement.tagName;
	return api.strcmpi(s, "input") == 0 || api.strcmpi(s, "textarea") == 0;
};

//

function InitMenus()
{
	te.Data.xmlMenus = OpenXml("menus.xml", false, true);
}

function ArrangeAddons()
{
	te.Data.Locations = te.Object();
	window.IconSize = te.Data.Conf_IconSize;
	var xml = OpenXml("addons.xml", false, true);
	te.Data.Addons = xml;
	if (api.GetKeyState(VK_SHIFT) < 0 && api.GetKeyState(VK_CONTROL) < 0) {
		IsSavePath = function (path, mode)
		{
			return false;
		}
		return;
	}
	var AddonId = new Array();
	var root = xml.documentElement;
	if (root) {
		var items = root.childNodes;
		if (items) {
			var arError = [];
			for (var i = 0; i < items.length; i++) {
				var item = items[i];
				var Id = item.nodeName;
				if (!AddonId[Id]) {
					var Enabled = api.QuadPart(item.getAttribute("Enabled"));
					if (Enabled & 6) {
						LoadLang2(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\" + Id + "\\lang\\" + GetLangId() + ".xml"));
					}
					if (Enabled & 8) {
						LoadAddon("vbs", Id, arError);
					}
					if (Enabled & 1) {
						LoadAddon("js", Id, arError);
					}
					AddonId[Id] = true;
				}
			}
			if (arError.length) {
				setTimeout(function () {
					wsh.Popup(arError.join("\n\n"), 9, TITLE, MB_ICONSTOP);
				}, 500);
			}
		}
	}
}

LoadAddon = function(ext, Addon_Id, arError)
{
	try {
		var ado = te.CreateObject("Adodb.Stream");
		ado.CharSet = "utf-8";
		ado.Open();
		var fname = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons") + "\\" + Addon_Id + "\\script." + ext;
		ado.LoadFromFile(fname);
		var s = ado.ReadText();
		ado.Close();
		if (api.strcmpi(ext, "js") == 0) {
			(new Function(s))(Addon_Id);
		}
		else if (api.strcmpi(ext, "vbs") == 0) {
			var fn = api.GetScriptDispatch(s, "VBScript", {"_Addon_Id": {"Addon_Id": Addon_Id}, window: window},
				function (ei, SourceLineText, dwSourceContext, lLineNumber, CharacterPosition)
				{
					arError.push(api.SysAllocString(ei.bstrDescription) + "\n" + fname);
				}
			);
			if (fn) {
				Addons["_stack"].push(fn);
			}
		}
	}
	catch (e) {
		arError.push(e.description + "\n" + fname);
	}
}

function GetAddonLocation(strName)
{
	var items = te.Data.Addons.getElementsByTagName(strName);
	return (items.length ? items[0].getAttribute("Location") : null);
}

function SetAddon(strName, Location, Tag)
{
	if (strName) {
		var s = GetAddonLocation(strName);
		if (s) {
			Location = s;
		}
	}
	if (Tag) {
		if (strName) {
			if (!te.Data.Locations[Location]) {
				te.Data.Locations[Location] = te.Array();
			}
			te.Data.Locations[Location].push(strName);
		}
		var o = document.getElementById(Location);
		if (api.strcmpi(typeof(Tag), "string") == 0) {
			o.insertAdjacentHTML("BeforeEnd", Tag);
		}
		else if (Tag.join) {
			o.insertAdjacentHTML("BeforeEnd", Tag.join(""));
		}
		else {
			o.appendChild(Tag);
		}
		o.style.display = !document.documentMode || api.strcmpi(o.tagName, "td") ? "block" : "table-cell";
	}
	return Location;
}

function InitCode()
{
	var types = 
	{
		Key:   ["All", "List", "Tree", "Browser"],
		Mouse: ["All", "List", "Tree", "Tabs", "Browser"]
	};
	var i;
	for (i = 0; i < 3; i++) {
		g_KeyState[i][0] = api.GetKeyNameText(g_KeyState[i][0]);
	}
	i = g_KeyState.length;
	while (i-- > 4 && g_KeyState[i][0] == g_KeyState[i - 4][0]) {
		g_KeyState.pop();
	}
	for (var j = 256; j >= 0; j -= 256) {
		for (var i = 128; i > 0; i--) {
			var v = api.GetKeyNameText((i + j) * 0x10000);
			if (v && v.charCodeAt(0) > 32) {
				g_KeyCode[v.toUpperCase()] = i + j;
			}
		}
	}
	for (var i in types) {
		eventTE[i] = {};
		for (var j in types[i]) {
			eventTE[i][types[i][j]] = {};
		}
	}
}

SetKeyExec = function (mode, strKey, path, type, bNew)
{
	if (strKey) {
		strKey = GetKeyKey(strKey);
		if (!bNew || !eventTE.Key[mode][strKey]) {
			eventTE.Key[mode][strKey] = [path, type];
		}
	}
}

SetGestureExec = function (mode, strGesture, path, type, bNew)
{
	if (strGesture) {
		if (!bNew || !eventTE.Mouse[mode][strGesture.toUpperCase()]) {
			eventTE.Mouse[mode][strGesture.toUpperCase()] = [path, type];
		}
	}
}

GestureExec = function (Ctrl, mode, str, pt)
{
	var cmd = eventTE.Mouse[mode][str];
	if (cmd) {
		return Exec(Ctrl, cmd[0], cmd[1], hwnd, pt);
	}
	return S_FALSE;
}

KeyExec = function (Ctrl, mode, str)
{
	var cmd = eventTE.Key[mode][GetKeyKey(str)];
	if (cmd) {
		return Exec(Ctrl, cmd[0], cmd[1], hwnd, null);
	}
	return S_FALSE;
}

function InitMouse()
{
	if (!te.Data.Conf_Gestures) {
		te.Data.Conf_Gestures = 2;
	}
	if (!te.Data.Conf_TrailColor) {
		te.Data.Conf_TrailColor = 0xff00;
	}
	if (!te.Data.Conf_TrailSize) {
		te.Data.Conf_TrailSize = 2;
	}
	if (!te.Data.Conf_GestureTimeout) {
		te.Data.Conf_GestureTimeout = 3000;
	}
}

g_mouse = 
{
	str: "",
	CancelButton: false,
	ptGesture: api.Memory("POINT"),
	hwndGesture: null,
	tidGesture: null,
	bCapture: false,
	RButton: -1,
	bTrail: false,
	tidDblClk: null,
	bDblClk: false,

	StartGestureTimer: function ()
	{
		var i = te.Data.Conf_GestureTimeout;
		if (i) {
			clearTimeout(this.tidGesture);
			this.tidGesture = setTimeout("g_mouse.EndGesture(true)", i);
		}
	},

	EndGesture: function (button)
	{
		clearTimeout(this.tidGesture);
		if (this.bCapture) {
			api.ReleaseCapture();
			this.bCapture = false;
		}
		if (this.RButton >= 0) {
			this.RButtonDown(false)
		}
		this.str = "";
		g_bRButton = false;
		SetGestureText(Ctrl, "");
		if (this.bTrail) {
			api.RedrawWindow(te.hwnd, null, 0, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
			this.bTrail = false;
		}
	},

	RButtonDown: function (mode)
	{
		if (api.strcmpi(this.str, "2") == 0) {
			var item = api.Memory("LVITEM");
			item.iItem = this.RButton;
			item.mask = LVIF_STATE;
			item.stateMask = LVIS_SELECTED;
			api.SendMessage(this.hwndGesture, LVM_GETITEM, 0, item);
			if (!(item.state & LVIS_SELECTED)) {
				if (mode) {
					var Ctrl = te.CtrlFromWindow(this.hwndGesture);
					Ctrl.SelectItem(Ctrl.Items.Item(this.RButton), SVSI_SELECT | SVSI_DESELECTOTHERS);
				}
				else {
					var ptc = api.Memory("POINT");
					ptc = te.Data.pt.clone();
					api.ScreenToClient(this.hwndGesture, ptc);
					api.SendMessage(this.hwndGesture, WM_RBUTTONDOWN, 0, ptc.x + (ptc.y << 16));
				}
			}
		}
		this.RButton = -1;
	},

	GetButton: function (msg, wParam)
	{
		if (msg >= WM_LBUTTONDOWN && msg <= WM_LBUTTONDBLCLK) {
			return "1";
		}
		if (msg >= WM_RBUTTONDOWN && msg <= WM_RBUTTONDBLCLK) {
			return "2";
		}
		if (msg >= WM_MBUTTONDOWN && msg <= WM_MBUTTONDBLCLK) {
			return "3";
		}
		if (msg >= WM_XBUTTONDOWN && msg <= WM_XBUTTONDBLCLK) {
			switch (wParam >> 16) {
				case XBUTTON1:
					return "4";
				case XBUTTON2:
					return "5";
			}
		}
		return "";
	},

	Exec: function (Ctrl, hwnd, pt)
	{
		this.str = GetGestureKey() + this.str;
		te.Data.cmdMouse = this.str;
		if (!Ctrl) {
			return S_FALSE;
		}
		var s = null;
		switch (Ctrl.Type) {
			case CTRL_SB:
			case CTRL_EB:
				s = "List";
				break;
			case CTRL_TV:
				s = "Tree";
				break;
			case CTRL_TC:
				s = "Tabs";
				break;
			case CTRL_WB:
				s = "Browser";
				break;
			default:
				if (this.str.length) {
					if (window.g_menu_click != 2) {
						window.g_menu_button = this.str;
					}
					if (window.g_menu_click) {
						var hSubMenu = api.GetSubMenu(window.g_menu_handle, window.g_menu_pos);
						if (hSubMenu) {
							var mii = api.Memory("MENUITEMINFO");
							mii.cbSize = mii.Size;
							mii.fMask = MIIM_SUBMENU;
							if (api.SetMenuItemInfo(window.g_menu_handle, window.g_menu_pos, true, mii)) {
								api.DestroyMenu(hSubMenu);
							}
						}
						if (this.str > 2) {
							window.g_menu_click = 2;
							var lParam = pt.x + (pt.y << 16);
							api.PostMessage(hwnd, WM_LBUTTONDOWN, 0, lParam);
							api.PostMessage(hwnd, WM_LBUTTONUP, 0, lParam);
						}
					}
				}
				break;
		}
		if (s) {
			var cmd = eventTE.Mouse[s][this.str];
			if (cmd) {
				return Exec(Ctrl, cmd[0], cmd[1], hwnd, pt);
			}
		}
		if (Ctrl.Type != CTRL_TE) {
			var cmd = eventTE.Mouse.All[this.str];
			if (cmd) {
				return Exec(Ctrl, cmd[0], cmd[1], hwnd, pt);
			}
		}
		return S_FALSE;
	}
};

g_basic =
{
	Func:
	{
		"Open":
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				return ExecOpen(Ctrl, s, type, hwnd, pt, OpenMode);
			},

			Drop: DropOpen
		},

		"Open in New Tab":
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				return ExecOpen(Ctrl, s, type, hwnd, pt, SBSP_NEWBROWSER);
			},

			Drop: DropOpen
		},

		"Open in Background":
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				return ExecOpen(Ctrl, s, type, hwnd, pt, SBSP_NEWBROWSER | SBSP_ACTIVATE_NOFOCUS);
			},

			Drop: DropOpen
		},

		"Exec":
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				try {
					wsh.Run(ExtractMacro(Ctrl, s));
				} catch (e) {}
				return S_OK;
			},

			Drop: function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
			{
				if (!pdwEffect) {
					pdwEffect = api.Memory("DWORD");
				}
				var re = /%Selected%/i;
				if (s.match(re)) {
					pdwEffect[0] = DROPEFFECT_LINK;
					if (bDrop) {
						var ar = [];
						for (var i = dataObj.Count; i > 0; ar.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(dataObj.Item(--i), SHGDN_FORPARSING)))) {
						}
						s = s.replace(re, ar.join(" "));
					}
					else {
						return S_OK;
					}
				}
				else {
					pdwEffect[0] = DROPEFFECT_NONE;
					return S_OK;
				}
			}
		},

		"JScript":
		{
			Exec: ExecScriptEx,
			Drop: DropScript
		},

		"VBScript":
		{
			Exec: ExecScriptEx,
			Drop: DropScript
		},

		"Selected Items": 
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				var fn = g_basic.Func["Selected Items"].Cmd[s];
				if (fn) {
					return fn(Ctrl);
				}
				else {
					return Exec(Ctrl, s + " %Selected%", "Exec", hwnd, pt);
				}
			},

			Ref: function (s, pt)
			{
				var r = g_basic.Popup(g_basic.Func["Selected Items"].Cmd, s, pt);
				if (api.strcmpi(r, GetText("Send to...")) == 0) {
					var Folder = sha.NameSpace(ssfSENDTO);
					if (Folder) {
						var Items = Folder.Items();
						var hMenu = api.CreatePopupMenu();
						for (i = 0; i < Items.Count; i++) {
							api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, i + 1, Items.Item(i).Name);
						}
						var nVerb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, null);
						api.DestroyMenu(hMenu);
						if (nVerb) {
							return api.PathQuoteSpaces(Items.Item(nVerb - 1).Path);
						}
					}
					return 1;
				}
				return api.strcmpi(r, GetText("Open with...")) ? r : 2;
			},

			Cmd:
			{
				Open: function (Ctrl)
				{
					return OpenSelected(Ctrl, OpenMode);
				},
				"Open in New Tab": function (Ctrl)
				{
					return OpenSelected(Ctrl, SBSP_NEWBROWSER);
				},
				"Open in Background": function (Ctrl)
				{
					return OpenSelected(Ctrl, SBSP_NEWBROWSER | SBSP_ACTIVATE_NOFOCUS);
				},
				"Open with...": undefined,
				"Send to...": undefined
			}

		},

		Tabs:
		{
			Cmd:
			{
				"Close Tab": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt, true);
					FV && FV.Close();
				},
				"Close Other Tabs": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						var TC = FV.Parent;
						var nIndex = GetFolderView(Ctrl, pt, true) ? FV.Index : -1;
						for (var i = TC.Count; i--;) {
							if (i != nIndex) {
								TC[i].Close();
							}
						}
					}
				},
				"Close Tabs on Left": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt, true);
					if (FV) {
						var TC = FV.Parent;
						for (var i = FV.Index; i--;) {
							TC[i].Close();
						}
					}
				},
				"Close Tabs on Right": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt, true);
					if (FV) {
						var TC = FV.Parent;
						var nIndex = FV.Index;
						for (var i = TC.Count; --i > nIndex;) {
							TC[i].Close();
						}
					}
				},
				"New Tab": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					FV && FV.Navigate(null, SBSP_RELATIVE | SBSP_NEWBROWSER);
				},
				Lock: function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					FV && Lock(FV.Parent, FV.Index, true);
				},
				"Previous Tab": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					FV && ChangeTab(FV.Parent, -1);
				},
				"Next Tab": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					FV && ChangeTab(FV.Parent, 1);
				},
				"Up": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					FV && FV.Navigate(null, SBSP_PARENT | OpenMode);
				},
				"Back": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					FV && FV.Navigate(null, SBSP_NAVIGATEBACK);
				},
				"Forward": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					FV && FV.Navigate(null, SBSP_NAVIGATEFORWARD);
				},
				"Refresh": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					FV && FV.Refresh();
				},
				"Switch Explorer Engine": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						FV.Type = (FV.Type == CTRL_SB) ? CTRL_EB : CTRL_SB;
					}
				},
				"Open in Explorer": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					FV && OpenInExplorer(FV);
				}
			}
		},

		Edit: 
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				var nVerb = GetCommandId(te.MainMenu(FCIDM_MENU_EDIT), s);
				SendCommand(Ctrl, nVerb);
				return S_OK;
			},

			Ref: function (s, pt)
			{
				return g_basic.PopupMenu(te.MainMenu(FCIDM_MENU_EDIT), null, pt);
			}
		},

		View: 
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				var nVerb = GetCommandId(te.MainMenu(FCIDM_MENU_VIEW), s);
				SendCommand(Ctrl, nVerb);
				return S_OK;
			},

			Ref: function (s, pt)
			{
				return g_basic.PopupMenu(te.MainMenu(FCIDM_MENU_VIEW), null, pt);
			}
		},

		Context: 
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				var Selected;
				if (Ctrl.Type <= CTRL_EB || Ctrl.Type == CTRL_TV) {
					Selected = Ctrl.SelectedItems();
				}
				else if (Ctrl.Type == CTRL_TV ) {
					Selected = Ctrl.SelectedItem;
				}
				else {
					var FV = te.Ctrl(CTRL_FV);
					Selected = FV.SelectedItems();
				}
				if (Selected && Selected.Count) {
					var ContextMenu = api.ContextMenu(Selected);
					if (ContextMenu) {
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_NORMAL);
						var nVerb = GetCommandId(hMenu, s);
						if (nVerb) {
							ContextMenu.InvokeCommand(0, te.hwnd, nVerb - 1, null, null, SW_SHOWNORMAL, 0, 0);
						}
						else {
							ContextMenu.InvokeCommand(0, te.hwnd, s, null, null, SW_SHOWNORMAL, 0, 0);
						}
					}
				}
				return S_OK;
			},

			Ref: function (s, pt)
			{
				var FV = te.Ctrl(CTRL_FV);
				if (FV) {
					var Selected = FV.SelectedItems();
					if (!Selected.Count) {
						Selected = api.GetModuleFileName(null);
					}
					var ContextMenu = api.ContextMenu(Selected);
					if (ContextMenu) {
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_NORMAL);
						return g_basic.PopupMenu(hMenu, ContextMenu, pt);
					}
				}
			}
		},

		ViewContext: 
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				var FV = te.Ctrl(CTRL_FV);
				if (FV) {
					var ContextMenu = FV.ViewMenu();
					if (ContextMenu) {
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_NORMAL);
						var nVerb = GetCommandId(hMenu, s);
						if (nVerb) {
							ContextMenu.InvokeCommand(0, te.hwnd, nVerb - 1, null, null, SW_SHOWNORMAL, 0, 0);
						}
					}
				}
				return S_OK;
			},

			Ref: function (s, pt)
			{
				var FV = te.Ctrl(CTRL_FV);
				if (FV) {
					var ContextMenu = FV.ViewMenu();
					if (ContextMenu) {
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_NORMAL);
						return g_basic.PopupMenu(hMenu, null, pt);
					}
				}
			}
		},

		Tools:
		{
			Cmd:
			{
				"New Folder": CreateNewFolder,
				"New File": CreateNewFile,
				"Copy Full Path": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					var Selected = FV.SelectedItems();
					var s = [];
					var nCount = Selected.Count;
					if (nCount) {
						while (--nCount >= 0) {
							s.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(Selected.Item(nCount), SHGDN_FORPARSING)));
						}
					}
					else {
						s.push(api.PathQuoteSpaces(api.GetDisplayNameOf(FV, SHGDN_FORPARSING)));
					}
					clipboardData.setData("text", s.join(" "));
				},
				"Run Dialog": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					api.ShRunDialog(te.hwnd, 0, FV ? FV.FolderItem.Path : null, null, null, 0);
				},
				"Add to Favorites": AddFavorite,
				"Reload Customize": function ()
				{
					te.reload();
				},
				"Load Layout": LoadLayout,
				"Save Layout": SaveLayout,
				"Close Application": function ()
				{
					api.PostMessage(te.hwnd, WM_CLOSE, 0, 0);
				}
			}
		},
		Options:
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				ShowOptions("Tab=" + s);
				return S_OK;
			},

			Ref: function (s, pt)
			{
				return g_basic.Popup(g_basic.Func.Options.List, s, pt);
			},

			List: ["General", "Add-ons", "Menus", "Tabs", "Tree", "List", "Init"]
		},
		"Add-ons":
		{
			Cmd:{}
		},
		Menus:
		{
			Ref: function (s, pt)
			{
				return g_basic.Popup(g_basic.Func.Menus.List, s, pt);
			},

			List: ["Open", "Close", "Separator", "Break", "BarBreak"]
		}
	},

	Exec: function (Ctrl, s, type, hwnd, pt)
	{
		var fn = g_basic.Func[type].Cmd[s];
		if (!pt) {
			var pt = api.Memory("POINT");
			api.GetCursorPos(pt);
		}
		fn && fn(Ctrl, pt);
		return S_OK;
	},

	Popup: function (Cmd, strDefault, pt)
	{
		var i, j, s;
		var ar = [];
		if (Cmd.length) {
			ar = Cmd;
		}
		else {
			for (i in Cmd) {
				ar.push(i);
			}
		}
		var hMenu = api.CreatePopupMenu();
		for (i = 0; i < ar.length; i++) {
			api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, i + 1, GetText(ar[i]));
		}
		var nVerb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, null);
		s = api.GetMenuString(hMenu, nVerb, MF_BYCOMMAND);
		api.DestroyMenu(hMenu);
		if (nVerb == 0) {
			return 1;
		}
		return s;
	},

	PopupMenu: function (hMenu, ContextMenu, pt)
	{
		var Verb;
		for (var i = api.GetMenuItemCount(hMenu); i--;) {
			if (api.GetMenuString(hMenu, i, MF_BYPOSITION)) {
				api.EnableMenuItem(hMenu, i, MF_ENABLED | MF_BYPOSITION);
			}
			else {
				api.DeleteMenu(hMenu, i, MF_BYPOSITION);
			}
		}
		var nVerb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, null);
		if (nVerb == 0) {
			api.DestroyMenu(hMenu);
			return 1;
		}
		if (ContextMenu) {
			Verb = ContextMenu.GetCommandString(nVerb - 1, GCS_VERB);
		}
		if (!Verb) {
			Verb = api.GetMenuString(hMenu, nVerb, MF_BYCOMMAND).replace(/\t.*$/, "");
		}
		api.DestroyMenu(hMenu);
		return Verb;
	}
};

AddEvent("Exec", function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
{
	var fn = g_basic.Func[type];
	if (fn) {
		if (dataObj) {
			if (fn.Drop) {
				return fn.Drop(Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop);
			}
			pdwEffect[0] = DROPEFFECT_NONE;
			return E_NOTIMPL;
		}
		if (fn.Exec) {
			return fn.Exec(Ctrl, s, type, hwnd, pt);
		}
		return g_basic.Exec(Ctrl, s, type, hwnd, pt);
	}
});

AddEvent("AddType", function (arFunc)
{
	for (var i in g_basic.Func) {
		arFunc.push(i);
	}
});

AddType = function (strType, o)
{
	g_basic.Func[strType] = o;
};

AddTypeEx = function (strType, strTitle, fn)
{
	var type = g_basic.Func[strType];
	if (type && type.Cmd) {
		type.Cmd[strTitle] = fn;
	}
};

AddEvent("OptionRef", function (Id, s, pt)
{
	var fn = g_basic.Func[Id];
	if (fn) {
		if (fn.Ref) {
			return fn.Ref(s, pt);
		}
		if (fn.Cmd) {
			return g_basic.Popup(fn.Cmd, s, pt);
		}
	}
});

AddEvent("OptionEncode", function (Id, p)
{
	if (g_basic.Func[Id]) {
		p.s = GetSourceText(p.s);
		return S_OK;
	}
});

AddEvent("OptionDecode", function (Id, p)
{
	if (g_basic.Func[Id]) {
		var s = GetText(p.s);
		if (GetSourceText(s) == p.s) {
			p.s = GetText(p.s);
			return S_OK;
		}
		return S_OK;
	}
});

//Init

if (!te.Data) {
	te.Data = te.Object();
	te.Data.CustColors = api.Memory("int", 16);
	//Default Value
	te.Data.Tab_Style = TCS_HOTTRACK | TCS_MULTILINE | TCS_RAGGEDRIGHT | TCS_SCROLLOPPOSITE | TCS_HOTTRACK | TCS_TOOLTIPS;
	te.Data.Tab_Align = TCA_TOP;
	te.Data.Tab_TabWidth = 96;
	te.Data.Tab_TabHeight = 0;

	te.Data.View_Type = CTRL_SB;
	te.Data.View_ViewMode = FVM_DETAILS;
	te.Data.View_fFlags = FWF_SHOWSELALWAYS | FWF_NOWEBVIEW;
	te.Data.View_Options = EBO_SHOWFRAMES | EBO_ALWAYSNAVIGATE;
	te.Data.View_ViewFlags = CDB2GVF_SHOWALLFILES;
	te.Data.View_IconSize = 0;

	te.Data.Tree_Align = 0;
	te.Data.Tree_Width = 200;
	te.Data.Tree_Style = NSTCS_HASEXPANDOS | NSTCS_SHOWSELECTIONALWAYS | NSTCS_HASLINES | NSTCS_BORDER | NSTCS_HORIZONTALSCROLL;
	te.Data.Tree_EnumFlags = SHCONTF_FOLDERS;
	te.Data.Tree_RootStyle = NSTCRS_VISIBLE | NSTCRS_EXPANDED;
	te.Data.Tree_Root = 0;

	te.Data.Installed = fso.GetParentFolderName(api.GetModuleFileName(null));
	var DataFolder = te.Data.Installed;

	var pf = [ssfPROGRAMFILES, ssfPROGRAMFILESx86];
	var x = api.sizeof("HANDLE") / 4;
	for (var i = 0; i < x; i++) {
		var s = api.GetDisplayNameOf(pf[i], SHGDN_FORPARSING);
		var l = s.replace(/\s*\(x86\)$/i, "").length;
		if (api.StrCmpI(s.substr(0, l), DataFolder.substr(0, l)) == 0) {
			DataFolder = fso.BuildPath(api.GetDisplayNameOf(ssfAPPDATA, SHGDN_FORPARSING), "Tablacus\\Explorer");
			var ParentFolder = fso.GetParentFolderName(DataFolder);
			if (!fso.FolderExists(ParentFolder)) {
				if (fso.CreateFolder(ParentFolder)) {
					if (!fso.FolderExists(DataFolder)) {
						fso.CreateFolder(DataFolder);
					}
				}
			}
			break;
		}
	}
	CreateFolder2(fso.BuildPath(DataFolder, "config"));
	if (!document.documentMode) {
		var s = fso.BuildPath(DataFolder, "cache");
		CreateFolder2(s);
		CreateFolder2(fso.BuildPath(s, "bitmap"));
		CreateFolder2(fso.BuildPath(s, "icon"));
		CreateFolder2(fso.BuildPath(s, "file"));
	}

	te.Data.DataFolder = DataFolder;
	te.Data.Conf_Lang = GetLangId();

	if (api.GetKeyState(VK_SHIFT) < 0 && api.GetKeyState(VK_CONTROL) < 0) {
		xmlWindow = "Init";
	}
	else {
		LoadConfig();
	}
}

InitCode();
InitMouse();
InitMenus();
LoadLang();
ArrangeAddons();
