//Tablacus Explorer

te.ClearEvents();
te.LockUpdate();
g_Bar = "";
Addon = 1;
Init = false;
OpenMode = SBSP_SAMEBROWSER;
ExtraMenuCommand = [];
g_arBM = [];
Exchange = {};

GetAddress = null;
ShowContextMenu = null;

g_tidResize = null;
g_tidWindowRegistered = null;
xmlWindow = null;
g_Panels = [];
g_KeyCode = {};
g_KeyState = [
	[0x1d0000, 0x2000],
	[0x2a0000, 0x1000],
	[0x380000, 0x4000],
	["Win",    0x8000],
	["Ctrl",   0x2000],
	["Shift",  0x1000],
	["Alt",    0x4000]
];
g_dlgs = {};
g_bWindowRegistered = true;
Addon_Id = "";
g_stack_TC = [];
g_pidlCP = api.ILRemoveLastID(ssfCONTROLS);
if (api.ILIsEmpty(g_pidlCP) || api.ILIsEqual(g_pidlCP, ssfDRIVES)) {
	g_pidlCP = ssfCONTROLS;
}

RunEvent1 = function (en, a1, a2, a3)
{
	var eo = eventTE[en];
	for (var i in eo) {
		try {
			eo[i](a1, a2, a3);
		} catch (e) {
			ShowError(e, en, i);
		}
	}
}

RunEvent2 = function (en, a1, a2, a3, a4)
{
	var eo = eventTE[en];
	for (var i in eo) {
		try {
			var hr = eo[i](a1, a2, a3, a4);
			if (isFinite(hr) && hr != S_OK) {
				return hr; 
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	return S_OK; 
}

RunEvent3 = function (en, a1, a2, a3, a4, a5, a6, a7, a8, a9)
{
	var eo = eventTE[en];
	for (var i in eo) {
		try {
			var hr = eo[i](a1, a2, a3, a4, a5, a6, a7, a8, a9);
			if (isFinite(hr)) {
				return hr; 
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
}

RunEvent4 = function (en, a1, a2)
{
	var eo = eventTE[en];
	for (var i in eo) {
		try {
			var r = eo[i](a1, a2);
			if (r !== undefined) {
				return r; 
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
}

PanelCreated = function (Ctrl)
{
	RunEvent1("PanelCreated", Ctrl);
}

ChangeView = function (Ctrl)
{
	if (Ctrl && Ctrl.FolderItem) {
		ChangeTabName(Ctrl);
		RunEvent1("ChangeView", Ctrl);
	}
	RunEvent1("ConfigChanged", "Config");
}

SetAddress = function (s)
{
	RunEvent1("SetAddress", s);
}

ChangeTabName = function (Ctrl)
{
	if (Ctrl.FolderItem) {
		Ctrl.Title = GetTabName(Ctrl);
	}
}

GetTabName = function (Ctrl)
{
	if (Ctrl.FolderItem) {
		var en = "GetTabName";
		var eo = eventTE[en];
		for (var i in eo) {
			try {
				var s = eo[i](Ctrl);
				if (s) {
					return s;
				}
			} catch (e) {
				ShowError(e, en, i);
			}
		}
		return Ctrl.FolderItem.Name;
	}
}

CloseView = function (Ctrl)
{
	return RunEvent2("CloseView", Ctrl);
}

DeviceChanged = function (Ctrl)
{
	g_tidDevice = null;
	RunEvent1("DeviceChanged", Ctrl);
}

ListViewCreated = function (Ctrl)
{
	RunEvent1("ListViewCreated", Ctrl);
}

TabViewCreated = function (Ctrl)
{
	RunEvent1("TabViewCreated", Ctrl);
}

TreeViewCreated = function (Ctrl)
{
	RunEvent1("TreeViewCreated", Ctrl);
}

SetAddrss = function (s)
{
	RunEvent1("SetAddrss", s);
}

RestoreFromTray = function ()
{
	api.ShowWindow(te.hwnd, api.IsIconic(te.hwnd) ? SW_RESTORE : SW_SHOW);
	api.SetForegroundWindow(te.hwnd);
	RunEvent1("RestoreFromTray");
}

Finalize = function ()
{
	RunEvent1("Finalize");
	SaveConfig();

	for (var i in g_dlgs) {
		var dlg = g_dlgs[i];
		try {
			if (dlg) {
				if (dlg.oExec) {
					if (dlg.oExec.Status == 0) {
						dlg.oExec.Terminate();
					}
				} else if (dlg.window) {
					dlg.close();
				}
			}
		} catch (e) {}
	}
}

SetGestureText = function (Ctrl, Text)
{
	RunEvent3("SetGestureText", Ctrl, Text);
	if (g_mouse.str.length > 1 && te.Data.Conf_Gestures > 1) {
		g_mouse.bTrail = true;
		var hdc = api.GetWindowDC(te.hwnd);
		if (hdc) {
			var rc = api.Memory("RECT");
			if (!g_mouse.ptText) {
				g_mouse.ptText = g_mouse.ptGesture.Clone();
				api.ScreenToClient(te.hwnd, g_mouse.ptText);
				g_mouse.right = -32767;
			}
			rc.left = g_mouse.ptText.x;
			rc.top = g_mouse.ptText.y;
			var hOld = api.SelectObject(hdc, CreateFont(DefaultFont));
			api.DrawText(hdc, Text, -1, rc, DT_CALCRECT);
			if (g_mouse.right < rc.right) {
				g_mouse.right = rc.right;
			}
			api.Rectangle(hdc, rc.left - 2, rc.top - 1, g_mouse.right + 2, rc.bottom + 1);
			api.DrawText(hdc, Text, -1, rc, DT_TOP);
			api.SelectObject(hdc, hOld);
			api.ReleaseDC(te.hwnd, hdc);
		}
	}
}

IsSavePath = function (path)
{
	var en = "IsSavePath";
	var eo = eventTE[en];
	for (var i in eo) {
		try {
			if (!eo[i](path)) {
				return false;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	return !api.PathMatchSpec(path, "search-ms:*");
}

Lock = function (Ctrl, nIndex, turn)
{
	var FV = Ctrl[nIndex];
	if (FV) {
		if (turn) {
			FV.Data.Lock = !FV.Data.Lock;
		}
		RunEvent1("Lock", Ctrl, nIndex, FV.Data.Lock);
	}
}

FontChanged = function ()
{
	RunEvent1("FontChanged");
}

FavoriteChanged = function ()
{
	RunEvent1("FavoriteChanged");
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
		if  (items.length) {
			var item = items[0];
			var x = api.QuadPart(item.getAttribute("Left"));
			var y = api.QuadPart(item.getAttribute("Top"));
			if (x > -30000 && y > -30000) {
				var w = api.QuadPart(item.getAttribute("Width"));
				var h = api.QuadPart(item.getAttribute("Height"));
				var pt = {x: x + w / 2, y: y};
				if (!api.MonitorFromPoint(pt, MONITOR_DEFAULTTONULL)) {
					var hMonitor = api.MonitorFromPoint(pt, MONITOR_DEFAULTTOPRIMARY);
					var mi = api.Memory("MONITORINFOEX");
					mi.cbSize = mi.Size;
					api.GetMonitorInfo(hMonitor, mi);
					x = mi.rcWork.left + (mi.rcWork.right - mi.rcWork.left - w) / 2;
					y = mi.rcWork.top + (mi.rcWork.bottom - mi.rcWork.top - h) / 2;
					if (y < mi.rcWork.top) {
						y = mi.rcWork.top;
					}
				}
				api.ShowWindow(te.hwnd, SW_SHOWNORMAL);
				api.MoveWindow(te.hwnd, x, y, w, h, true);
			}
			te.CmdShow = item.getAttribute("CmdShow");
		}
		return;
	}
	xmlWindow = "Init";
}

SaveConfig = function ()
{
	RunEvent1("SaveConfig");
	if (te.Data.bSaveMenus) {
		te.Data.bSaveMenus = false;
		SaveXmlEx("menus.xml", te.Data.xmlMenus);
	}
	if (te.Data.bSaveAddons) {
		te.Data.bSaveAddons = false;
		SaveXmlEx("addons.xml", te.Data.Addons);
	}
	if (te.Data.bSaveConfig) {
		te.Data.bSaveConfig = false;
		if (document.msFullscreenElement) {
			while (g_stack_TC.length) {
				g_stack_TC.pop().Visible = true;
			}
			document.msExitFullscreen();
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
	var o = document.getElementById("toolbar");
	var offsetTop = o ? o.offsetHeight : 0;
	te.offsetTop = Math.ceil(offsetTop * screen.deviceYDPI / screen.logicalYDPI);

	var h = 0;
	o = document.getElementById("bottombar");
	if (o) {
		var offsetBottom = o.offsetHeight;
		o = document.getElementById("client");
		if (o) {
			h = (document.documentElement ? document.documentElement.offsetHeight : document.body.offsetHeight) - offsetBottom - offsetTop;
			o.style.height = ((h >= 0) ? h : 0) + "px";
		}
		te.offsetBottom = Math.ceil(offsetBottom * screen.deviceYDPI / screen.logicalYDPI);
	}
	o = document.getElementById("leftbarT");
	if (o) {
		var i = h;
		o.style.height = ((i >= 0) ? i : 0) + "px";
		i = o.clientHeight - o.style.height.replace(/\D/g, "");

		var h2 = o.clientHeight - document.getElementById("LeftBar1").offsetHeight - document.getElementById("LeftBar3").offsetHeight;
		document.getElementById("LeftBar2").style.height = h2 - i + "px";
	}
	o = document.getElementById("rightbarT");
	if (o) {
		var i = h;
		o.style.height = ((i >= 0) ? i : 0) + "px";
		i = o.clientHeight - o.style.height.replace(/\D/g, "");

		var h2 = o.clientHeight - document.getElementById("RightBar1").offsetHeight - document.getElementById("RightBar3").offsetHeight;
		document.getElementById("RightBar2").style.height = h2 - i + "px";
	}

	var w2 = 0;
	var w = 0;
	o = te.Data.Locations;
	if (o.LeftBar1 || o.LeftBar2 || o.LeftBar3) {
		w = te.Data.Conf_LeftBarWidth;
	}
	o = document.getElementById("leftbar");
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
	te.offsetLeft = Math.ceil(w2 * screen.deviceXDPI / screen.logicalXDPI);

	w2 = 0;
	w = 0;
	o = te.Data.Locations;
	if (o.RightBar1 || o.RightBar2 || o.RightBar3) {
		w = te.Data.Conf_RightBarWidth;
	}
	o = document.getElementById("rightbar");
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
	te.offsetRight = Math.ceil(w2 * screen.deviceXDPI / screen.logicalXDPI);
	o = document.getElementById("Background");
	if (o) {
		w = (document.documentElement ? document.documentElement.offsetWidth : document.body.offsetWidth) - te.offsetLeft - te.offsetRight;
		if (w > 0) {
			o.style.width = w + "px";
		}
	}
	RunEvent1("Resize");
}

LoadLang = function (bAppend)
{
	if (!bAppend) {
		MainWindow.Lang = {};
		MainWindow.LangSrc = {};
	}
	var filename = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "lang\\" + GetLangId() + ".xml");
	LoadLang2(filename);
}

Refresh = function (Ctrl, pt)
{
	var FV = GetFolderView(Ctrl, pt);
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
		var s = InputDialog("Add Favorite", api.GetDisplayNameOf(FolderItem, SHGDN_INFOLDER));
		if (s) {
			item.setAttribute("Name", s);
			item.setAttribute("Filter", "");
			item.text = api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
			if (fso.FileExists(item.text)) {
				item.text = api.PathQuoteSpaces(item.text);
				item.setAttribute("Type", "Exec");
			} else {
				item.setAttribute("Type", "Open");
			}
			menus[0].appendChild(item);
			SaveXmlEx("menus.xml", xml);
			FavoriteChanged();
			return true;
		}
	}
	return false;
}

AddFavoriteEx = function (Ctrl, pt)
{
	AddFavorite();
	return S_OK
}

CancelFilterView = function (FV)
{
	if (IsSearchPath(FV) || FV.FilterView) {
		FV.Navigate(null, SBSP_PARENT);
		return S_OK;
	}
	return S_FALSE;
}

IsSearchPath = function (FI)
{
	return api.PathMatchSpec(api.GetDisplayNameOf(FI, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING), "search-ms:*");
}

GetCommandId = function (hMenu, s, ContextMenu)
{
	var arMenu = [hMenu];
	var wId = 0;
	if (s) {
		var sl = GetText(s);
		var mii = api.Memory("MENUITEMINFO");
		mii.cbSize = mii.Size;
		mii.fMask = MIIM_SUBMENU | MIIM_ID | MIIM_FTYPE | MIIM_STATE;
		while (arMenu.length) {
			hMenu = arMenu.pop();
			if (ContextMenu) {
				ContextMenu.HandleMenuMsg(WM_INITMENUPOPUP, hMenu, 0);
			}
			var i = api.GetMenuItemCount(hMenu);
			if (i == 1 && ContextMenu && WINVER >= 0x600) {
				setTimeout('api.EndMenu()', 200);
				api.TrackPopupMenuEx(hMenu, TPM_RETURNCMD, 0, 0, te.hwnd, null, ContextMenu);
				i = api.GetMenuItemCount(hMenu);
			}
			while (i-- > 0) {
				if (api.GetMenuItemInfo(hMenu, i, true, mii) && !(mii.fType & MFT_SEPARATOR) && !(mii.fState & MFS_DISABLED)) {
					var title = api.GetMenuString(hMenu, i, MF_BYPOSITION);
					if (title) {
						var a = title.split(/\t/);
						if (api.PathMatchSpec(a[0], s) || api.PathMatchSpec(a[0], sl) || api.PathMatchSpec(a[1], s) || (ContextMenu && api.PathMatchSpec(ContextMenu.GetCommandString(mii.wId - 1, GCS_VERB), s))) {
							wId = mii.hSubMenu ? api.TrackPopupMenuEx(mii.hSubMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, ContextMenu) : mii.wId;
							arMenu.length = 0;
							break;
						}
						if (mii.hSubMenu) {
							arMenu.push(mii.hSubMenu);
						}
					}
				}
			}
			if (ContextMenu) {
				ContextMenu.HandleMenuMsg(WM_UNINITMENUPOPUP, hMenu, 0);
			}
		}
	}
	return wId;
};

OpenSelected = function (Ctrl, NewTab, pt)
{
	var ar = GetSelectedArray(Ctrl, pt, true);
	var Selected = ar.shift();
	var SelItem = ar.shift();
	var FV = ar.shift();
	if (Selected) {
		var Exec = [];
		for (var i in Selected) {
			var Item = Selected.Item(i);
			var bFolder = Item.IsFolder;
			if (!bFolder) {
				var path = Item.ExtendedProperty("linktarget");
				if (path) {
					bFolder = api.PathIsDirectory(path) === true;
				}
			}
			if (bFolder) {
			 	FV.Navigate(Item, NewTab);
			 	NewTab |= SBSP_NEWBROWSER;
			} else {
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
			InvokeCommand(Selected, 0, te.hwnd, null, null, null, SW_SHOWNORMAL, 0, 0, FV, CMF_DEFAULTONLY);
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
		} else {
			var s = img.src;
			if (bDisable) {
				if (/^data:image\/png/i.test(s)) {
					img.src = "data:image/svg+xml," + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ' + img.offsetWidth + ' ' + img.offsetHeight + '"><filter id="G"><feColorMatrix type="saturate" values="0.1" /></filter><image width="' + img.width + '" height="' + img.height + '" xlink:href="' + img.src + '" filter="url(#G)" opacity=".48"></image></svg>');
				}
			} else if (/^data:image\/svg/i.test(s) && /href="([^"]*)/i.test(decodeURIComponent(s))) {
				img.src = RegExp.$1
			}
		}
	}
}

//Events

te.OnCreate = function (Ctrl)
{
	if (Ctrl.Type == CTRL_TE) {
		if (xmlWindow && typeof(xmlWindow) != "string") {
			LoadXml(xmlWindow);
		}
		if (te.Ctrls(CTRL_TC).length == 0) {
			var TC = te.CreateCtrl(CTRL_TC, 0, 0, "100%", "100%", te.Data.Tab_Style, te.Data.Tab_Align, te.Data.Tab_TabWidth, te.Data.Tab_TabHeight);
			TC.Selected.Navigate2(HOME_PATH, SBSP_NEWBROWSER, te.Data.View_Type, te.Data.View_ViewMode, te.Data.View_fFlags, te.Data.View_Options, te.Data.View_ViewFlags, te.Data.View_IconSize, te.Data.Tree_Align, te.Data.Tree_Width, te.Data.Tree_Style, te.Data.Tree_EnumFlags, te.Data.Tree_RootStyle, te.Data.Tree_Root, te.Data.View_SizeFormat);
		}
		xmlWindow = null;
		setTimeout(function ()
		{
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
			Resize();
			var cTC = te.Ctrls(CTRL_TC);
			for (var i in cTC) {
				if (cTC[i].SelectedIndex >= 0) {
					ChangeView(cTC[i].Selected);
				}
			}
			api.ShowWindow(te.hwnd, te.CmdShow);
			te.UnlockUpdate();
			setTimeout(function ()
			{
				RunCommandLine(api.GetCommandLine());
			}, 500);
		}, 99);
		RunEvent1("Create", Ctrl);
	} else {
		if (!Ctrl.Data) {
			Ctrl.Data = te.Object();
		}
		RunEvent1("Create", Ctrl);
	}
}

te.OnClose = function (Ctrl)
{
	return RunEvent2("Close", Ctrl);
}

AddEvent("Close", function (Ctrl)
{
	switch (Ctrl.Type) {
		case CTRL_TE:
			Finalize();
			if (api.GetThreadCount() && MessageBox("File is in operation.", TITLE, MB_ABORTRETRYIGNORE) != IDIGNORE) {
				return S_FALSE;
			}
			api.SHChangeNotifyDeregister(te.Data.uRegisterId);
			break;
		case CTRL_WB:
			break;
		case CTRL_SB:
		case CTRL_EB:
			if (Ctrl.Data.Lock) {
				return S_FALSE;
			}
			var hr = CloseView(Ctrl);
			if (hr == S_OK && Ctrl.Parent.Count <= 1) {
				Ctrl.Navigate(HOME_PATH, SBSP_SAMEBROWSER);
				hr = S_FALSE;
			}
			return hr;
		case CTRL_TC:
			var o = document.getElementById("Panel_" + Ctrl.Id);
			if (o) {
				o.style.display = "none";
			}
			break;
	}
});

te.OnViewCreated = function (Ctrl)
{
	switch (Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
			ListViewCreated(Ctrl);
			break;
		case CTRL_TC:
			TabViewCreated(Ctrl);
			break;
		case CTRL_TV:
			TreeViewCreated(Ctrl);
			break;
	}
	RunEvent1("ViewCreated", Ctrl);
}

te.OnBeforeNavigate = function (Ctrl, fs, wFlags, Prev)
{
	if (Ctrl.Type <= CTRL_EB) {
		if (api.ILIsParent(g_pidlCP, Ctrl, false) && !api.ILIsEqual(Ctrl, g_pidlCP) && !api.ILIsEqual(Ctrl, ssfCONTROLS)) {
			(function (Path) { setTimeout(function () {
				OpenInExplorer(Path);
			}, 99);}) (Ctrl.FolderItem);
			return E_NOTIMPL;
		}
	}
	if (Ctrl.Data.Lock && (wFlags & SBSP_NEWBROWSER) == 0) {
		return E_ACCESSDENIED;
	}
	return RunEvent2("BeforeNavigate", Ctrl, fs, wFlags, Prev);
}

te.OnNavigateComplete = function (Ctrl)
{
	RunEvent1("NavigateComplete", Ctrl);
	ChangeView(Ctrl);
	return S_OK;
}

ShowStatusText = function (Ctrl, Text, iPart)
{
	RunEvent1("StatusText", Ctrl, Text, iPart);
	return S_OK; 
}

te.OnStatusText = ShowStatusText;

te.OnKeyMessage = function (Ctrl, hwnd, msg, key, keydata)
{
	var hr = RunEvent3("KeyMessage", Ctrl, hwnd, msg, key, keydata);
	if (isFinite(hr)) {
		return hr; 
	}
	if (g_mouse.str.length > 1) {
		SetGestureText(Ctrl, GetGestureKey() + g_mouse.str);
	}
	if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN) {
		var nKey = ((keydata >> 16) & 0x17f) | GetKeyShift();
		if (nKey == 0x15d) {
			g_mouse.CancelContextMenu = false;
		}
		te.Data.cmdKey = nKey;
		switch (Ctrl.Type) {
			case CTRL_SB:
			case CTRL_EB:
				var strClass = api.GetClassName(hwnd);
				if (api.PathMatchSpec(strClass, WC_LISTVIEW + ";DirectUIHWND")) {
					if (KeyExecEx(Ctrl, "List", nKey, hwnd) === S_OK) {
						return S_OK;
					}
				}
				if (strClass == WC_EDIT) {
					return S_FALSE;
				}
				break;
			case CTRL_TV:
				var strClass = api.GetClassName(hwnd);
				if (api.PathMatchSpec(strClass, WC_TREEVIEW)) {
					if (KeyExecEx(Ctrl, "Tree", nKey, hwnd) === S_OK) {
						return S_OK;
					}
					if (key == VK_DELETE) {
						InvokeCommand(Ctrl.SelectedItem, 0, te.hwnd, CommandID_DELETE - 1, null, null, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
						return S_OK;
					}
				}
				if (strClass == WC_EDIT) {
					return S_FALSE;
				}
				break;
			case CTRL_WB:
				if (KeyExecEx(Ctrl, "Browser", nKey, hwnd) === S_OK) {
					return S_OK;
				}
				break;
			default:
				if (window.g_menu_click) {
					if (key == VK_RETURN) {
						if (window.g_menu_click === true) {
							var hSubMenu = api.GetSubMenu(window.g_menu_handle, window.g_menu_pos);
							if (hSubMenu) {
								var mii = api.Memory("MENUITEMINFO");
								mii.cbSize = mii.Size;
								mii.fMask = MIIM_SUBMENU;
								api.SetMenuItemInfo(window.g_menu_handle, window.g_menu_pos, true, mii);
								api.DestroyMenu(hSubMenu);
								api.PostMessage(hwnd, WM_CHAR, VK_LBUTTON, 0);
							}
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
			if (KeyExecEx(Ctrl, "All", nKey, hwnd) === S_OK) {
				return S_OK;
			}
		}
	}
	return S_FALSE; 
}

te.OnMouseMessage = function (Ctrl, hwnd, msg, wParam, pt)
{
	var hr = RunEvent3("MouseMessage", Ctrl, hwnd, msg, wParam, pt);
	if (isFinite(hr)) {
		return hr; 
	}
	var strClass = api.GetClassName(hwnd);
	var bLV = Ctrl.Type <= CTRL_EB && api.PathMatchSpec(strClass, WC_LISTVIEW + ";DirectUIHWND");
	if (msg == WM_MOUSEWHEEL) {
		var Ctrl2 = te.CtrlFromPoint(pt);
		if (Ctrl2) {
			g_mouse.str = GetGestureButton() + (wParam > 0 ? "8" : "9");
			if (api.GetKeyState(VK_RBUTTON) < 0) {
				g_mouse.CancelContextMenu = true;
			}
			if (g_mouse.Exec(Ctrl2, hwnd, pt) == S_OK) {
				return S_OK;
			}
			var hwnd2 = api.WindowFromPoint(pt);
			if (hwnd2 && hwnd != hwnd2) {
				api.SetFocus(hwnd2);
				api.SendMessage(hwnd2, msg, wParam & 0xffff0000, pt.x + (pt.y << 16));
				return S_OK;
			}
		}
	}
	if (msg == WM_LBUTTONUP || msg == WM_RBUTTONUP || msg == WM_MBUTTONUP || msg == WM_XBUTTONUP) {
		if (g_mouse.str.length) {
			var hr = S_FALSE;
			var bButton = false;
			if (msg == WM_RBUTTONUP) {
				if (g_mouse.RButton >= 0) {
					g_mouse.RButtonDown(true);
					bButton = (g_mouse.str == "2");
				} else if (bLV) {
					var iItem = Ctrl.HitTest(pt);
					if (iItem < 0 && !IsDrag(pt, te.Data.pt)) {
						Ctrl.SelectItem(null, SVSI_DESELECTOTHERS);
					}
				}
			}
			if (msg == WM_MBUTTONUP) {
				if (bLV && api.GetKeyState(VK_SHIFT) >= 0 && api.GetKeyState(VK_CONTROL) >= 0) {
					var ar = eventTE.Mouse.List[g_mouse.str];
					if (ar && !api.StrCmpI(ar[0][1], "Selected Items")) {
						var iItem = Ctrl.HitTest(pt);
						if (iItem >= 0) {
							Ctrl.SelectItem(Ctrl.Item(iItem), SVSI_SELECT | SVSI_FOCUSED | SVSI_DESELECTOTHERS);
						} else {
							Ctrl.SelectItem(null, SVSI_DESELECTOTHERS);
						}
					}
				}
			}
			if (g_mouse.str.length >= 2 || (!IsDrag(pt, te.Data.pt) && strClass != WC_HEADER)) {
				if (msg != WM_RBUTTONUP || g_mouse.str.length < 2) {
					hr = g_mouse.Exec(te.CtrlFromWindow(g_mouse.hwndGesture), g_mouse.hwndGesture, pt);
					if (msg == WM_LBUTTONUP) {
						hr = S_FALSE;
					}
				} else {
					(function (Ctrl, hwnd, pt, str) { setTimeout(function () {
						hr = g_mouse.Exec(Ctrl, hwnd, pt, str);
					}, 99);}) (te.CtrlFromWindow(g_mouse.hwndGesture), g_mouse.hwndGesture, pt, g_mouse.str);
					hr = S_OK;
				}
			}
			g_mouse.EndGesture(false);
			if (g_mouse.bCapture) {
				api.ReleaseCapture();
				g_mouse.bCapture = false;
			}
			if (bButton) {
				api.PostMessage(g_mouse.hwndGesture, WM_CONTEXTMENU, g_mouse.hwndGesture, pt.x + (pt.y << 16));
				return S_OK;
			}
			if (hr === S_OK) {
				return S_OK;
			}
		}
	}
	if (msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN || msg == WM_MBUTTONDOWN || msg == WM_XBUTTONDOWN) {
		var s = g_mouse.GetButton(msg, wParam);
		if (g_mouse.str.indexOf(s) >= 0) {
			g_mouse.str = "";
		}
		if (g_mouse.str.length == 0) {
			te.Data.pt = pt.Clone();
			g_mouse.ptGesture.x = pt.x;
			g_mouse.ptGesture.y = pt.y;
			g_mouse.hwndGesture = hwnd;
		}
		g_mouse.str += s;
		g_mouse.StartGestureTimer();
		SetGestureText(Ctrl, GetGestureKey() + g_mouse.str);
		if (msg == WM_RBUTTONDOWN) {
			g_mouse.CancelContextMenu = api.GetKeyState(VK_LBUTTON) < 0 || api.GetKeyState(VK_MBUTTON) < 0 || api.GetKeyState(VK_XBUTTON1) < 0 || api.GetKeyState(VK_XBUTTON2) < 0;
			if (te.Data.Conf_Gestures >= 2) {
				var iItem = -1;
				if (bLV) {
					iItem = Ctrl.HitTest(pt);
					if (iItem < 0) {
						api.ScreenToClient(hwnd, pt);
						return pt.y < 32 ? S_FALSE : S_OK;
					}
				}
				if (te.Data.Conf_Gestures == 3 && Ctrl.Type != CTRL_WB) {
					g_mouse.RButton = iItem;
					g_mouse.StartGestureTimer();
					return S_OK;
				}
			}
		}
	}
	if (msg == WM_LBUTTONDBLCLK || msg == WM_RBUTTONDBLCLK || msg == WM_MBUTTONDBLCLK || msg == WM_XBUTTONDBLCLK) {
		if (strClass != WC_HEADER) {
			te.Data.pt = pt.Clone();
			g_mouse.str = g_mouse.GetButton(msg, wParam);
			g_mouse.str += g_mouse.str;
			if (g_mouse.Exec(Ctrl, hwnd, pt) == S_OK) {
				return S_OK;
			}
		}
	}

	if (msg == WM_MOUSEMOVE) {
		if (api.GetKeyState(VK_ESCAPE) < 0) {
			g_mouse.EndGesture(false);
		}
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
				var s = (Math.abs(x) >= Math.abs(y)) ? ((x < 0) ? "L" : "R") : ((y < 0) ? "U" : "D");

				if (s != g_mouse.str.charAt(g_mouse.str.length - 1)) {
					g_mouse.str += s;
					if (api.GetKeyState(VK_RBUTTON) < 0) {
						g_mouse.CancelContextMenu = true;
					}
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
	if (Ctrl.Type <= CTRL_EB) {
		if ((wParam & 0xfff) + 1 == CommandID_DELETE) {
			var Items = Ctrl.SelectedItems();
			for (var i = Items.Count; i--;) {
				UnlockFV(Items[i]);
			}
		}
	}
	var hr = RunEvent3("Command", Ctrl, hwnd, msg, wParam, lParam);
	RunEvent1("ConfigChanged", "Config");
	return isFinite(hr) ? hr : S_FALSE; 
}

te.OnInvokeCommand = function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon)
{
	var path;
	var hr = RunEvent3("InvokeCommand", ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon);
	RunEvent1("ConfigChanged", "Config");
	if (isFinite(hr)) {
		return hr; 
	}
	if (typeof(Directory) == "string" && !fso.FolderExists(Directory)) {
		return S_FALSE;
	}
	var Items = ContextMenu.Items();
	var Exec = [];
	if (isFinite(Verb)) {
		Verb = ContextMenu.GetCommandString(Verb, GCS_VERB);
	}
	var strVerb = String(Verb).toLowerCase();
	if (api.PathMatchSpec(strVerb, "opennewwindow;opennewprocess")) {
		CancelWindowRegistered();
	}

	NewTab = GetNavigateFlags();
	for (var i = 0; i < Items.Count; i++) {
		if (Verb && strVerb != "runas") {
			path = Items.Item(i).ExtendedProperty("linktarget") || Items.Item(i).Path;
			var cmd = api.AssocQueryString(ASSOCF_NONE, ASSOCSTR_COMMAND, path, strVerb == "default" ? null : Verb).replace(/"?%1"?|%L/g, api.PathQuoteSpaces(path)).replace(/%\*|%I/g, "");
			if (cmd) {
				ShowStatusText(te, Verb + ":" + cmd, 1);
				if (strVerb == "open" && api.PathMatchSpec(cmd, "?:\\Windows\\Explorer.exe;*\\Explorer.exe /idlist,*;rundll32.exe *fldr.dll,RouteTheCall*")) {
					Navigate(Items.Item(i), NewTab);
				 	NewTab |= SBSP_NEWBROWSER;
					continue;
				}
				var cmd2 = ExtractMacro(te, cmd);
				if (api.StrCmpI(cmd, cmd2)) {
					ShellExecute(cmd2, null, nShow, Directory);
					continue;
				}
			}
			if (strVerb == "open" && IsFolderEx(Items.Item(i))) {
				Navigate(Items.Item(i), NewTab);
				NewTab |= SBSP_NEWBROWSER;
				continue;
			}
		}
		Exec.push(Items.Item(i));
	}
	if (Items.Count != Exec.length) {
		if (Exec.length) {
			var Selected = te.FolderItems();
			for (var i = 0; i < Exec.length; i++) {
				Selected.AddItem(Exec[i]);
			}
			InvokeCommand(Selected, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon, ContextMenu.FolderView);
		}
		return S_OK;
	}
	ShowStatusText(te, [Verb || "", Items.Count == 1 ? Items.Item(0).Path : Items.Count].join(":"), 1);
	return S_FALSE; 
}

AddEvent("InvokeCommand", function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon)
{
	var strVerb = String(isFinite(Verb) ? ContextMenu.GetCommandString(Verb, GCS_VERB) : Verb).toLowerCase();
	if (strVerb == "opencontaining") {
		var Items = ContextMenu.Items();
		for (var j in Items) {
			var Item = Items.Item(j);
			var path = Item.ExtendedProperty("linktarget");
			if (path) {
			 	api.PathIsDirectory(function (hr, path)
			 	{
					if (hr >= 0) {
						Navigate(path, SBSP_NEWBROWSER);
					} else {
						Navigate(fso.GetParentFolderName(path), SBSP_NEWBROWSER);
						setTimeout(function ()
						{
							var FV = te.Ctrl(CTRL_FV);
							FV.SelectItem(path, SVSI_SELECT | SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS);
						}, 99);
					}
				}, -1, path, path);
			}
			return S_OK;
		}
	}
	if (strVerb == "cmd") {
		ShellExecute(ExtractMacro(ContextMenu.FolderView, "%ComSpec% /k cd %Current%"), null, SW_SHOWNORMAL);
		return S_OK;
	}
	if (strVerb == "delete") {
		var Items = ContextMenu.Items();
		for (var i = Items.Count; i--;) {
			UnlockFV(Items.Item(i));
		}
	}
});

te.OnDragEnter = function (Ctrl, dataObj, pgrfKeyState, pt, pdwEffect)
{
	var hr = E_NOTIMPL;
	var dwEffect = pdwEffect[0];
	var en = "DragEnter";
	var eo = eventTE[en];
	for (var i in eo) {
		try {
			pdwEffect[0] = dwEffect;
			var hr2 = eo[i](Ctrl, dataObj, pgrfKeyState[0], pt, pdwEffect, pgrfKeyState);
			if (isFinite(hr2) && hr != S_OK) {
				hr = hr2;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	g_mouse.str = "";
	return hr; 
}

te.OnDragOver = function (Ctrl, dataObj, pgrfKeyState, pt, pdwEffect)
{
	var dwEffect = pdwEffect[0];
	var en = "DragOver";
	var eo = eventTE[en];
	for (var i in eo) {
		try {
			pdwEffect[0] = dwEffect;
			var hr = eo[i](Ctrl, dataObj, pgrfKeyState[0], pt, pdwEffect, pgrfKeyState);
			if (isFinite(hr)) {
				return hr; 
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	return E_NOTIMPL; 
}

te.OnDrop = function (Ctrl, dataObj, pgrfKeyState, pt, pdwEffect)
{
	var dwEffect = pdwEffect[0];
	var en = "Drop";
	var eo = eventTE[en];
	for (var i in eo) {
		try {
			pdwEffect[0] = dwEffect;
			var hr = eo[i](Ctrl, dataObj, pgrfKeyState[0], pt, pdwEffect, pgrfKeyState);
			if (isFinite(hr)) {
				return hr; 
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	return E_NOTIMPL; 
}

te.OnDragLeave = function (Ctrl)
{
	var hr = E_NOTIMPL;
	var en = "DragLeave";
	var eo = eventTE[en];
	for (var i in eo) {
		try {
			var hr2 = eo[i](Ctrl);
			if (isFinite(hr2) && hr != S_OK) {
				hr = hr2;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	g_mouse.str = "";
	return hr;
}

te.OnSelectionChanging = function (Ctrl)
{
	var hr = RunEvent3("SelectionChanging", Ctrl);
	return isFinite(hr) ? hr : S_OK;
}

te.OnSelectionChanged = function (Ctrl, uChange)
{
	if (Ctrl.Type == CTRL_TC && Ctrl.SelectedIndex >= 0) {
		ChangeView(Ctrl.Selected);
	}
	var hr = RunEvent3("SelectionChanged", Ctrl, uChange);
	return isFinite(hr) ? hr : S_OK;
}

te.OnViewModeChanged = function (Ctrl)
{
	RunEvent1("ViewModeChanged", Ctrl);
}

te.OnColumnsChanged = function (Ctrl)
{
	RunEvent1("ColumnsChanged", Ctrl);
}

te.OnShowContextMenu = function (Ctrl, hwnd, msg, wParam, pt)
{
	if (g_mouse.CancelContextMenu) {
		return S_OK;
	}
	var hr = RunEvent3("ShowContextMenu", Ctrl, hwnd, msg, wParam, pt);
	if (isFinite(hr)) {
		return hr; 
	}
	switch (Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
			var Selected = Ctrl.SelectedItems();
			if (Selected.Count) {
				if (ExecMenu(Ctrl, "Context", pt, 1) == S_OK) {
					return S_OK;
				}
			} else {
				if (ExecMenu(Ctrl, "Background", pt, 1) == S_OK) {
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
				MakeMenus(hMenu, menus, arMenu, items, Ctrl, pt);
			}
			api.DestroyMenu(hMenu);
			break;
	}
	return S_FALSE;
}

te.OnDefaultCommand = function (Ctrl)
{
	if (api.GetKeyState(VK_MENU) < 0) {
		return S_FALSE;
	}
	var Selected = Ctrl.SelectedItems();
	var hr = RunEvent3("DefaultCommand", Ctrl, Selected);
	if (isFinite(hr)) {
		return hr; 
	}
	if (ExecMenu(Ctrl, "Default", null, 2) != S_OK) {
		InvokeCommand(Selected, 0, te.hwnd, null, null, null, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
	}
	return S_OK;
}

te.OnItemClick = function (Ctrl, Item, HitTest, Flags)
{
	var hr = RunEvent3("ItemClick", Ctrl, Item, HitTest, Flags);
	return isFinite(hr) ? hr : S_FALSE;
}

te.OnColumnClick = function (Ctrl, iItem)
{
	var hr = RunEvent3("ColumnClick", Ctrl, iItem);
	return isFinite(hr) ? hr : S_FALSE;
}

te.OnSystemMessage = function (Ctrl, hwnd, msg, wParam, lParam)
{
	var hr = RunEvent3("SystemMessage", Ctrl, hwnd, msg, wParam, lParam);
	if (isFinite(hr)) {
		return hr; 
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
					var locator = te.CreateObject("WbemScripting.SWbemLocator");
					if (locator) {
						var server = locator.ConnectServer();
						if (server) {
							var cols = server.ExecQuery("SELECT * FROM Win32_Process WHERE ExecutablePath = '" + api.GetModuleFileName(null).replace(/\\/g, "\\\\") + "'");
							for (var list = new Enumerator(cols); !list.atEnd(); list.moveNext()) {
								var item = list.item();
								if (/\/run/i.test(item.CommandLine)) {
									var hwnd = GethwndFromPid(item.ProcessId);
									api.SetWindowLongPtr(hwnd, GWL_EXSTYLE, api.GetWindowLongPtr(hwnd, GWL_EXSTYLE) & ~0x80);
									api.ShowWindow(hwnd, SW_SHOWNORMAL);
									api.SetForegroundWindow(hwnd);
								}
							}
						}
					}
					for (var i in te.Data.Fonts) {
						api.DeleteObject(te.Data.Fonts[i]);
					}
					te.Data.Fonts = null;
					break;
				case WM_DEVICECHANGE:
					if (wParam == DBT_DEVICEARRIVAL || wParam == DBT_DEVICEREMOVECOMPLETE) {
						DeviceChanged(Ctrl);
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
					if (wParam & 0xffff) {
						if (g_mouse.str == "") {
							setTimeout(function ()
							{
								var FV = te.Ctrl(CTRL_FV);
								if (FV) {
									FV.Focus();
								}
							}, 99);
						}
					} else  {
						g_mouse.str = "";
						SetGestureText(Ctrl, "");
					}
					break;
				case WM_QUERYENDSESSION:
					SaveConfig();
					return true;
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
				case WM_POWERBROADCAST:
					if (wParam == PBT_APMRESUMEAUTOMATIC) {
						var cFV = te.Ctrls(CTRL_FV);
						for (var i in cFV) {
							var FV = cFV[i];
							if (FV.hwndView) {
								if (api.PathIsNetworkPath(api.GetDisplayNameOf(FV.FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING))) {
									FV.Suspend();
								}
							}
						}
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
	if (msg == WM_COPYDATA) {
		var cd = api.Memory("COPYDATASTRUCT", 1, lParam);
		var hr = RunEvent3("CopyData", Ctrl, cd, wParam);
		if (isFinite(hr)) {
			return hr; 
		}
		if (Ctrl.Type == CTRL_TE && cd.dwData == 0 && cd.cbData) {
			var strData = api.SysAllocStringByteLen(cd.lpData, cd.cbData);
			RestoreFromTray();
			RunCommandLine(strData);
			return S_OK;
		}
	}
	return 0; 
};

te.OnMenuMessage = function (Ctrl, hwnd, msg, wParam, lParam)
{
	var hr = RunEvent3("MenuMessage", Ctrl, hwnd, msg, wParam, lParam);
	if (isFinite(hr)) {
		return hr; 
	}
	switch (msg) {
		case WM_INITMENUPOPUP:
			var s = api.GetMenuString(wParam, 0, MF_BYPOSITION);
			if (api.PathMatchSpec(s, "\t*Script\t*")) {
				var ar = s.split("\t");
				api.DeleteMenu(wParam, 0, MF_BYPOSITION);
				ExecScriptEx(Ctrl, ar[2], ar[1], hwnd);
			}
			break;
		case WM_MENUSELECT:
			if (lParam) {
				var s = wParam & 0xffff;
				var hSubMenu = api.GetSubMenu(lParam, s);
				if ((wParam >> 16) & MF_POPUP) {
					if (hSubMenu) {
						window.g_menu_handle = lParam;
						window.g_menu_pos = s;
					}
				}
				window.g_menu_string = api.GetMenuString(lParam, s, hSubMenu ? MF_BYPOSITION : MF_BYCOMMAND);
			}
			break;
		case WM_EXITMENULOOP:
			window.g_menu_click = false;
			var en = "ExitMenuLoop";
			var eo = eventTE[en];
			try {
				while (eo && eo.length) {
					eo.shift()();
				}
			} catch (e) {
				ShowError(e, en);
			}
			while (g_arBM.length) {
				api.DeleteObject(g_arBM.pop());
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
	var hr = RunEvent3("AppMessage", Ctrl, hwnd, msg, wParam, lParam);
	if (isFinite(hr)) {
		return hr; 
	}
	if (msg == TWM_CHANGENOTIFY) {
		var pidls = {};
		var hLock = api.SHChangeNotification_Lock(wParam, lParam, pidls);
		if (hLock) {
			api.SHChangeNotification_Unlock(hLock);
			ChangeNotifyFV(pidls.lEvent, pidls[0], pidls[1]);
			RunEvent1("ChangeNotify", Ctrl, pidls);
		}
		return S_OK;
	}
	return S_FALSE;
}

te.OnNewWindow = function (Ctrl, dwFlags, UrlContext, Url)
{
	var hr = RunEvent3("NewWindow", Ctrl, dwFlags, UrlContext, Url);
	if (isFinite(hr)) {
		return hr; 
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
	var r = RunEvent4("ClipboardText", Items);
	if (r !== undefined) {
		return r; 
	}
	var s = [];
	for (var i = Items.Count; i-- > 0;) {
		s.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(Items.Item(i), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)))
	}
	return s.join(" ");
}

te.OnArrange = function (Ctrl, rc)
{
	if (Ctrl.Type == CTRL_TC) {
		var o = g_Panels[Ctrl.Id];
		if (!o) {
			var s = ['<table id="Panel_$" class="layout" style="position: absolute; z-index: 1;">'];
			s.push('<tr><td id="InnerLeft_$" class="sidebar" style="width: 0px; display: none; overflow: auto"></td><td><div id="InnerTop_$" style="display: none"></div>');
			s.push('<table id="InnerTop2_$" class="layout" style="width: 100%">');
			s.push('<tr><td id="Inner1Left_$" class="toolbar1"></td><td id="Inner1Center_$" class="toolbar2" style="white-space: nowrap;"></td><td id="Inner1Right_$" class="toolbar3"></td></tr></table>');
			s.push('<table id="InnerView_$" class="layout" style="width: 100%"><tr><td id="Inner2Left_$" style="width: 0px"></td><td id="Inner2Center_$"></td><td id="Inner2Right_$" style="width: 0px; overflow: auto"></td></tr></table>');
			s.push('<div id="InnerBottom_$"></div></td><td id="InnerRight_$" class="sidebar" style="width: 0px; display: none"></td></tr></table>');
			document.getElementById("Panel").insertAdjacentHTML("BeforeEnd", s.join("").replace(/\$/g, Ctrl.Id));
			PanelCreated(Ctrl);
			o = document.getElementById("Panel_" + Ctrl.Id);
			g_Panels[Ctrl.Id] = o;
			ApplyLang(o);
			ChangeView(Ctrl.Selected);
		}
		o.style.left = (rc.Left * screen.logicalXDPI / screen.deviceXDPI) + "px";
		o.style.top = (rc.Top * screen.logicalYDPI / screen.deviceYDPI) + "px";
		if (Ctrl.Visible) {
			o.style.display = (document.documentMode && o.tagName.toLowerCase() == "td") ? "table-cell" : "block";
		} else {
			o.style.display = "none";
		}
		var i = rc.Right - rc.Left
		o.style.width = (i >= 0 ? i : 0) + "px";
		i = rc.Bottom - rc.Top;
		o.style.height = (i >= 0 ? i : 0) + "px";
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
		i = rc.Right - rc.Left;
		o.width = i > 0 ? i + "px" : "0px";
		i = rc.Bottom - rc.Top;
		o.height = i > 0 ? i + "px" : "0px";
	}
	RunEvent1("Arrange", Ctrl, rc);
}

te.OnVisibleChanged = function (Ctrl)
{
	if (Ctrl.Type == CTRL_TC) {
		var o = g_Panels[Ctrl.Id];
		if (o) {
			if (Ctrl.Visible) {
				o.style.display = (document.documentMode && o.tagName.toLowerCase() == "td") ? "table-cell" : "block";
			} else {
				o.style.display = "none";
			}
		}
		if (Ctrl.Visible && Ctrl.SelectedIndex >= 0) {
			ChangeView(Ctrl.Selected);
			for (var i = Ctrl.Count; i--;) {
				ChangeTabName(Ctrl[i])
			}
		}
	}
	RunEvent1("VisibleChanged", Ctrl);
}

te.OnWindowRegistered = function (Ctrl)
{
	if (g_bWindowRegistered) {
		RunEvent1("WindowRegistered", Ctrl);
	}
	g_bWindowRegistered = true;
}

te.OnToolTip = function (Ctrl, Index)
{
	return RunEvent4("ToolTip", Ctrl, Index);
}

te.OnHitTest = function (Ctrl, pt, flags)
{
	var hr = RunEvent3("HitTest", Ctrl, pt, flags);
	return isFinite(hr) ? hr : -1;
}

te.OnGetPaneState = function (Ctrl, ep, peps)
{
	var hr = RunEvent3("GetPaneState", Ctrl, ep, peps);
	return isFinite(hr) ? hr : E_NOTIMPL;
}

te.OnTranslatePath = function (Ctrl, Path)
{
	return RunEvent4("TranslatePath", Ctrl, Path);
}

te.OnILGetParent = function (FolderItem)
{
	var r = RunEvent4("ILGetParent", FolderItem);
	if (r !== undefined) {
		return r; 
	}
	if (/search\-ms:.*?&crumb=location:([^&]*)/.test(api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING))) {
		return api.PathCreateFromUrl("file:" + RegExp.$1);
	}
	if (api.ILIsEqual(FolderItem, ssfRESULTSFOLDER)) {
		var ar = api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX).split("\\");
		if (ar.length > 1 && ar.pop()) {
			return ar.join("\\");
		}
	}
}

//Tablacus Events

GetIconImage = function (Ctrl, BGColor)
{
	var img = RunEvent4("GetIconImage", Ctrl, BGColor);
	if (img) {
		return img;
	}
	var FolderItem = Ctrl.FolderItem || Ctrl;
	if (!FolderItem || FolderItem.Unavailable) {
		return MakeImgSrc("icon:shell32.dll,234,16", 0, false, 16);
	}
	var path = FolderItem.Path;
	if (api.PathIsNetworkPath(path)) {
		if (fso.GetDriveName(path) != path.replace(/\\$/, "")) {
			return MakeImgSrc(WINVER >= 0x600 ? "icon:shell32.dll,275,16" : "icon:shell32.dll,85,16", 0, false, 16);
		}
		return MakeImgSrc(WINVER >= 0x600 ? "icon:shell32.dll,273,16" : "icon:shell32.dll,9,16", 0, false, 16);
	}
	if (document.documentMode) {
		var info = api.Memory("SHFILEINFO");
		api.SHGetFileInfo(FolderItem, 0, info, info.Size, SHGFI_ICON | SHGFI_SMALLICON | SHGFI_PIDL);
		var image = te.GdiplusBitmap();
		image.FromHICON(info.hIcon, BGColor);
		api.DestroyIcon(info.hIcon);
		return image.DataURI("image/png");
	}
	return MakeImgSrc("icon:shell32.dll,3,16", 0, false, 16);
}

// Browser Events

AddEventEx(window, "load", function ()
{
	ApplyLang(document);
	if (api.GetKeyState(VK_SHIFT) < 0 && api.GetKeyState(VK_CONTROL) < 0) {
		ShowOptions("Tab=Add-ons");
	}
});

AddEventEx(window, "resize", Resize);

AddEventEx(window, "beforeunload", Finalize);

AddEventEx(document, "MSFullscreenChange", function ()
{
	if (document.msFullscreenElement) {
		var cTC = te.Ctrls(CTRL_TC);
		for (var i in cTC) {
			var TC = cTC[i];
			if (TC.Visible) {
				g_stack_TC.push(TC);
				TC.Visible = false;
			}
		}
	} else {
		while (g_stack_TC.length) {
			g_stack_TC.pop().Visible = true;
		}
	}
});

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
		IsSavePath = function (path)
		{
			return false;
		}
		return;
	}
	var AddonId = [];
	var root = xml.documentElement;
	if (root) {
		var items = root.childNodes;
		if (items) {
			var arError = [];
			for (var i = 0; i < items.length; i++) {
				var item = items[i];
				var Id = item.nodeName;
				window.Error_source = Id;
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
				window.Error_source = "";
			}
			if (arError.length) {
				(function (arError) { setTimeout(function () {
					if (MessageBox(arError.join("\n\n"), TITLE, MB_OKCANCEL) != IDCANCEL) {
						te.Data.bErrorAddons = true;
						ShowOptions("Tab=Add-ons");
					}
				}, 500);}) (arError);
			}
		}
	}
	DeviceChanged(te);
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
		if (Tag.join) {
			Tag = Tag.join("");
		}
		var o = document.getElementById(Location);
		if (o) {
			if (typeof(Tag) == "string") {
				o.insertAdjacentHTML("BeforeEnd", Tag);
			} else {
				o.appendChild(Tag);
			}
			o.style.display = (document.documentMode && o.tagName.toLowerCase() == "td") ? "table-cell" : "block";
		} else if (Location == "Inner") {
			AddEvent("PanelCreated", function (Ctrl)
			{
				SetAddon(null, "Inner1Left_" + Ctrl.Id, Tag);
			});
		}
	}
	return Location;
}

function InitCode()
{
	var types = 
	{
		Key:   ["All", "List", "Tree", "Browser"],
		Mouse: ["All", "List", "List_Background", "Tree", "Tabs", "Tabs_Background", "Browser"]
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
	DefaultFont = api.Memory("LOGFONT");
	api.SystemParametersInfo(SPI_GETICONTITLELOGFONT, DefaultFont.Size, DefaultFont, 0);
	HOME_PATH = te.Data.Conf_NewTab || HOME_PATH;
}

function UnlockFV(Item)
{
	var cFV = te.Ctrls(CTRL_FV);
	for (var i in cFV) {
		var FV = cFV[i];
		if (FV.hwndView && api.ILIsParent(Item, FV, false)) {
			FV.Suspend();
		}
	}
}

function ChangeNotifyFV(lEvent, item1, item2)
{
	var fAdd = SHCNE_DRIVEADD | SHCNE_MEDIAINSERTED | SHCNE_NETSHARE | SHCNE_MKDIR;
	var fRemove = SHCNE_DRIVEREMOVED | SHCNE_MEDIAREMOVED | SHCNE_NETUNSHARE | SHCNE_RENAMEFOLDER | SHCNE_RMDIR | SHCNE_SERVERDISCONNECT;
	if (lEvent & (SHCNE_DISKEVENTS | fAdd | fRemove)) {
		var path1 = String(api.GetDisplayNameOf(item1, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING));
		var cFV = te.Ctrls(CTRL_FV);
		for (var i in cFV) {
			var FV = cFV[i];
			var path = String(api.GetDisplayNameOf(FV.FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING));
			if (lEvent == SHCNE_RENAMEFOLDER && !FV.Data.Lock) {
				if (api.PathMatchSpec(path, [path1.replace(/\\$/, ""), path1].join("\\*;"))) {
					FV.Navigate(path.replace(path1, api.GetDisplayNameOf(item2, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)), SBSP_SAMEBROWSER);
					continue;
				}
			}
			if (FV.hwndView) {
				if (lEvent & fRemove) {
					if (api.ILIsParent(item1, FV, false)) {
						FV.Suspend();
						continue;
					}
				}
				FV.Notify(lEvent, item1, item2);
				if (FV.hwndList) {
					if (api.ILIsParent(FV, item1, true)) {
						var item = api.Memory("LVITEM");
						item.stateMask = LVIS_CUT;
						api.SendMessage(FV.hwndList, LVM_SETITEMSTATE, -1, item);
					}
				}
				if ((lEvent & fAdd) && FV.FolderItem.Unavailable) {
					FV.Refresh();
				}
				if (lEvent & SHCNE_UPDATEDIR) {
					if (/^[A-Z]:\\|^\\/i.test(path) && api.PathMatchSpec(path, [path1.replace(/\\$/, ""), path1].join("\\*;"))) {
					 	if (!FV.FolderItem.Unavailable) {
						 	api.PathIsDirectory(function (hr, FV, FolderItem)
						 	{
								if (hr < 0 && api.ILIsEqual(FV.FolderItem, FolderItem)) {
									FV.Suspend(2);
								}
							}, 0, FV.FolderItem, FV, FV.FolderItem);
						} else if (FV.FolderItem.Unavailable > 3000) {
							FV.FolderItem.Unavailable = 1;
						 	api.PathIsDirectory(function (hr, FV)
						 	{
								if (hr >= 0 && FV.FolderItem.Unavailable) {
									FV.Refresh();
								}
							}, -1, path, FV);
						}
					}
				}
			}
		}
	}
}

SetKeyExec = function (mode, strKey, path, type, bLast)
{
	if (/,/.test(strKey) && !/,$/.test(strKey)) {
		var ar = strKey.split(",");
		for (var i in ar) {
			SetKeyExec(mode, ar[i], path, type, bLast);
		}
		return;
	}
	if (strKey) {
		strKey = GetKeyKey(strKey);
		var KeyMode = eventTE.Key[mode];
		if (KeyMode) {
			if (!KeyMode[strKey]) {
				KeyMode[strKey] = [];
			}
			if (bLast) {
				KeyMode[strKey].push([path, type]);
			} else {
				KeyMode[strKey].unshift([path, type]);
			}
		}
	}
}

SetGestureExec = function (mode, strGesture, path, type, bLast)
{
	if (/,/.test(strGesture) && !/,$/.test(strGesture)) {
		var ar = strGesture.split(",");
		for (var i in ar) {
			SetGestureExec(mode, ar[i], path, type, bLast);
		}
		return;
	}
	if (strGesture) {
		strGesture = strGesture.toUpperCase();
		var MouseMode = eventTE.Mouse[mode];
		if (MouseMode) {
			if (!MouseMode[strGesture]) {
				MouseMode[strGesture] = [];
			}
			if (bLast) {
				MouseMode[strGesture].push([path, type]);
			} else {
				MouseMode[strGesture].unshift([path, type]);
			}
		}
	}
}

ArExec = function (Ctrl, ar, pt, hwnd)
{
	for (var i in ar) {
		var cmd = ar[i];
		if (Exec(Ctrl, cmd[0], cmd[1], hwnd, pt) === S_OK) {
			return S_OK;
		}
	}
	return S_FALSE;
}

GestureExec = function (Ctrl, mode, str, pt, hwnd)
{
	if (hwnd && Ctrl.Type != CTRL_TC && Ctrl.Type != CTRL_WB) {
		api.SetFocus(hwnd);
	}
	return ArExec(Ctrl, eventTE.Mouse[mode][str], pt, hwnd || Ctrl.hwnd);
}

KeyExec = function (Ctrl, mode, str, hwnd)
{
	return KeyExecEx(Ctrl, mode, GetKeyKey(str), hwnd || Ctrl.hwnd);
}

KeyExecEx = function (Ctrl, mode, nKey, hwnd)
{
	var pt = api.Memory("POINT");
	if (Ctrl.Type <= CTRL_EB || Ctrl.Type == CTRL_TV) {
		var rc = api.Memory("RECT");
		Ctrl.GetItemRect(Ctrl.FocusedItem || Ctrl.SelectedItem, rc);
		pt.x = rc.Left;
		pt.y = rc.Top;
	}
	api.ClientToScreen(Ctrl.hwnd, pt);
	return ArExec(Ctrl, eventTE.Key[mode][nKey], pt, hwnd);
}

function InitMouse()
{
	te.Data.Conf_Gestures = isFinite(te.Data.Conf_Gestures) ? Number(te.Data.Conf_Gestures) : 2;
	if (typeof(te.Data.Conf_TrailColor) == "string") {
		te.Data.Conf_TrailColor = GetWinColor(te.Data.Conf_TrailColor);
	}
	if (!isFinite(te.Data.Conf_TrailColor)) {
		te.Data.Conf_TrailColor = 0xff00;
	}
	te.Data.Conf_TrailSize = isFinite(te.Data.Conf_TrailSize) ? Number(te.Data.Conf_TrailSize) : 2;
	te.Data.Conf_GestureTimeout = isFinite(te.Data.Conf_GestureTimeout) ? Number(te.Data.Conf_GestureTimeout) : 3000;
	te.Data.Conf_Layout = isFinite(te.Data.Conf_Layout) ? Number(te.Data.Conf_Layout) : 0x80;
	te.Data.Conf_NetworkTimeout = isFinite(te.Data.Conf_NetworkTimeout) ? Number(te.Data.Conf_NetworkTimeout) : 2000;
	te.Layout = te.Data.Conf_Layout;
	te.NetworkTimeout = te.Data.Conf_NetworkTimeout;
}

g_mouse = 
{
	str: "",
	CancelContextMenu: false,
	ptGesture: api.Memory("POINT"),
	hwndGesture: null,
	tidGesture: null,
	bCapture: false,
	RButton: -1,
	bTrail: false,
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
			api.RedrawWindow(te.hwnd, null, 0, RDW_INVALIDATE | RDW_FRAME | RDW_ALLCHILDREN);
			this.bTrail = false;
			this.ptText = null;
		}
	},

	RButtonDown: function (mode)
	{
		if (this.str == "2") {
			var item = api.Memory("LVITEM");
			item.iItem = this.RButton;
			item.mask = LVIF_STATE;
			item.stateMask = LVIS_SELECTED;
			api.SendMessage(this.hwndGesture, LVM_GETITEM, 0, item);
			if (!(item.state & LVIS_SELECTED)) {
				if (mode) {
					var Ctrl = te.CtrlFromWindow(this.hwndGesture);
					Ctrl.SelectItem(Ctrl.Item(this.RButton), SVSI_SELECT | SVSI_FOCUSED | SVSI_DESELECTOTHERS);
				} else {
					var ptc = api.Memory("POINT");
					ptc = te.Data.pt.Clone();
					api.ScreenToClient(this.hwndGesture, ptc);
					api.SendMessage(this.hwndGesture, WM_RBUTTONDOWN, 0, ptc.x + (ptc.y << 16));
				}
			}
		}
		this.RButton = -1;
	},

	GetButton: function (msg, wParam)
	{
		var s = "";
		if (msg >= WM_LBUTTONDOWN && msg <= WM_LBUTTONDBLCLK) {
			if (api.GetKeyState(VK_RBUTTON) < 0) {
				g_mouse.CancelContextMenu = true;
			}
			s = "1";
		}
		if (msg >= WM_RBUTTONDOWN && msg <= WM_RBUTTONDBLCLK) {
			s = "2";
		}
		if (msg >= WM_MBUTTONDOWN && msg <= WM_MBUTTONDBLCLK) {
			if (api.GetKeyState(VK_RBUTTON) < 0) {
				g_mouse.CancelContextMenu = true;
			}
			s = "3";
		}
		if (msg >= WM_XBUTTONDOWN && msg <= WM_XBUTTONDBLCLK) {
			if (api.GetKeyState(VK_RBUTTON) < 0) {
				g_mouse.CancelContextMenu = true;
			}
			switch (wParam >> 16) {
				case XBUTTON1:
					s = "4";
					break;
				case XBUTTON2:
					s = "5";
					break;
			}
		}
		return this.str.length ? s : GetGestureButton().replace(s, "") + s;
	},

	Exec: function (Ctrl, hwnd, pt, str)
	{
		var str = GetGestureKey() + (str || this.str);
		this.EndGesture(false);
		te.Data.cmdMouse = str;
		if (!Ctrl) {
			return S_FALSE;
		}
		var s = null;
		switch (Ctrl.Type) {
			case CTRL_SB:
			case CTRL_EB:
				if (!Ctrl.SelectedItems.Count) {
					if (GestureExec(Ctrl, "List_Background", str, pt, hwnd) === S_OK) {
						return S_OK;
					}
				}
				s = "List";
				break;
			case CTRL_TV:
				s = "Tree";
				break;
			case CTRL_TC:
				if (Ctrl.HitTest(pt, TCHT_ONITEM) < 0) {
					if (GestureExec(Ctrl, "Tabs_Background", str, pt, hwnd) === S_OK) {
						return S_OK;
					}
				}
				s = "Tabs";
				break;
			case CTRL_WB:
				s = "Browser";
				break;
			default:
				if (str.length) {
					window.g_menu_button = str;
					if (window.g_menu_click) {
						if (window.g_menu_click === true) {
							var hSubMenu = api.GetSubMenu(window.g_menu_handle, window.g_menu_pos);
							if (hSubMenu) {
								var mii = api.Memory("MENUITEMINFO");
								mii.cbSize = mii.Size;
								mii.fMask = MIIM_SUBMENU;
								if (api.SetMenuItemInfo(window.g_menu_handle, window.g_menu_pos, true, mii)) {
									api.DestroyMenu(hSubMenu);
								}
							}
						}
						if (str > 2) {
							window.g_menu_click = 4;
							var lParam = pt.x + (pt.y << 16);
							api.PostMessage(hwnd, WM_LBUTTONDOWN, 0, lParam);
							api.PostMessage(hwnd, WM_LBUTTONUP, 0, lParam);
						}
					}
				}
				break;
		}
		if (s) {
			if (GestureExec(Ctrl, s, str, pt, hwnd) === S_OK) {
				return S_OK;
			}
		}
		if (Ctrl.Type != CTRL_TE) {
			if (GestureExec(Ctrl, "All", str, pt, hwnd) === S_OK) {
				return S_OK;
			}
		}
		return S_FALSE;
	}
};

g_basic =
{
	FuncI: function (s)
	{
		return this.Func[s] || api.ObjGetI(this.Func, s);
	},

	CmdI: function (s, s2)
	{
		var type = this.Func[s] || api.ObjGetI(this.Func, s);
		if (type) {
			return type.Cmd[s2] || api.ObjGetI(type.Cmd, s2);
		}
	},

	Func:
	{
		"":
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				var lines = s.split(/\r?\n/);
				for (var i in lines) {
					var cmd = lines[i].split(",");
					var Id = cmd.shift();
					var hr = Exec(Ctrl, cmd.join(","), Id, hwnd, pt);
					if (hr != S_OK) {
						break;
					}
				}
				return S_OK;
			},

			Ref: function (s, pt)
			{
				var lines = s.split(/\r?\n/);
				var last = lines.length ? lines[lines.length - 1] : "";
				if (/^([^,]+),$/.test(last)) {
					var Id = GetSourceText(RegExp.$1);
					var r = OptionRef(Id, "", pt);
					if (typeof r == "string") {
						return s + r + "\n";
					}
					return r;
				} else {
					var arFunc = [];
					RunEvent1("AddType", arFunc);
					var r = g_basic.Popup(arFunc, s, pt);
					return r == 1 ? 1 : s + (s.length && !/\n$/.test(s) ? "\n" : "") + r + ",";
				}
			}
		},

		Open:
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				return ExecOpen(Ctrl, s, type, hwnd, pt, GetNavigateFlags(GetFolderView(Ctrl, pt)));
			},

			Drop: DropOpen,
			Ref: BrowseForFolder
		},

		"Open in New Tab":
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				return ExecOpen(Ctrl, s, type, hwnd, pt, SBSP_NEWBROWSER);
			},

			Drop: DropOpen,
			Ref: BrowseForFolder
		},

		"Open in Background":
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				return ExecOpen(Ctrl, s, type, hwnd, pt, SBSP_NEWBROWSER | SBSP_ACTIVATE_NOFOCUS);
			},

			Drop: DropOpen,
			Ref: BrowseForFolder
		},

		Exec:
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				s = ExtractMacro(Ctrl, s);
				try {
					ShellExecute(s, null, SW_SHOWNORMAL, Ctrl, pt);
				} catch (e) {
					ShowError(e, s);
				}
				return S_OK;
			},

			Drop: function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
			{
				if (!pdwEffect) {
					pdwEffect = api.Memory("DWORD");
				}
				var re = /%Selected%/i;
				if (re.test(s)) {
					pdwEffect[0] = DROPEFFECT_LINK;
					if (bDrop) {
						var ar = [];
						for (var i = dataObj.Count; i > 0; ar.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(dataObj.Item(--i), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)))) {
						}
						s = s.replace(re, ar.join(" "));
					} else {
						return S_OK;
					}
				} else {
					pdwEffect[0] = DROPEFFECT_NONE;
					return S_OK;
				}
			},

			Ref: OpenDialog
		},

		RunAs:
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				s = ExtractMacro(Ctrl, s);
				try {
					ShellExecute(s, "RunAs", SW_SHOWNORMAL, Ctrl, pt);
				} catch (e) {
					ShowError(e, s);
				}
				return S_OK;
			},

			Ref: OpenDialog
		},

		JScript:
		{
			Exec: ExecScriptEx,
			Drop: DropScript,
			Ref: OpenDialog
		},

		VBScript:
		{
			Exec: ExecScriptEx,
			Drop: DropScript,
			Ref: OpenDialog
		},

		"Selected Items": 
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				var fn = g_basic.CmdI(type, s);
				if (fn) {
					return fn(Ctrl, pt);
				} else {
					return Exec(Ctrl, s + " %Selected%", "Exec", hwnd, pt);
				}
			},

			Ref: function (s, pt)
			{
				var r = g_basic.Popup(g_basic.Func["Selected Items"].Cmd, s, pt);
				if (api.StrCmpI(r, GetText("Send to...")) == 0) {
					var Folder = sha.NameSpace(ssfSENDTO);
					if (Folder) {
						var Items = Folder.Items();
						var hMenu = api.CreatePopupMenu();
						for (i = 0; i < Items.Count; i++) {
							api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, i + 1, Items.Item(i).Name);
						}
						var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, null);
						api.DestroyMenu(hMenu);
						if (nVerb) {
							return api.PathQuoteSpaces(Items.Item(nVerb - 1).Path);
						}
					}
					return 1;
				}
				if (api.StrCmpI(r, GetText("Open with...")) == 0) {
					r = OpenDialog(s);
					if (!r) {
						r = 1;
					}
				}
				return r;
			},

			Cmd:
			{
				Open: function (Ctrl, pt)
				{
					return OpenSelected(Ctrl, GetNavigateFlags(GetFolderView(Ctrl)), pt);
				},
				"Open in New Tab": function (Ctrl, pt)
				{
					return OpenSelected(Ctrl, SBSP_NEWBROWSER, pt);
				},
				"Open in Background": function (Ctrl, pt)
				{
					return OpenSelected(Ctrl, SBSP_NEWBROWSER | SBSP_ACTIVATE_NOFOCUS, pt);
				},
				Exec: function (Ctrl, pt)
				{
					var Selected = GetSelectedArray(Ctrl, pt, true).shift();
					if (Selected) {
						InvokeCommand(Selected, 0, te.hwnd, null, null, null, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
					}
					return S_OK;
				},
				"Open with...": undefined,
				"Send to...": undefined
			},
			Enc: true
		},

		Tabs:
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				var fn = g_basic.CmdI(type, s);
				if (fn) {
					fn(Ctrl, pt);
					return S_OK;
				}
				var FV = GetFolderView(Ctrl, pt, true);
				if (FV) {
					var TC = FV.Parent;
					if (/^\d/.test(s)) {
						TC.SelectedIndex = api.QuadPart(s);
					} else if (/^\-/.test(s)) {
						TC.SelectedIndex = TC.Count + api.QuadPart(s);
					}
				}
				return S_OK;
			},
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
					if (HOME_PATH) {
						NavigateFV(FV, HOME_PATH, SBSP_NEWBROWSER);
					} else {
						NavigateFV(FV, null, SBSP_RELATIVE | SBSP_NEWBROWSER);
					}
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
					FV && FV.Navigate(null, SBSP_PARENT | GetNavigateFlags(FV));
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
				"Show frames": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						FV.Type = (FV.Type == CTRL_SB) ? CTRL_EB : CTRL_SB;
					}
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
			},

			Enc: true
		},

		Edit: 
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				var hMenu = te.MainMenu(FCIDM_MENU_EDIT);
				var nVerb = GetCommandId(hMenu, s);
				SendCommand(Ctrl, nVerb);
				api.DestroyMenu(hMenu);
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
				var hMenu = te.MainMenu(FCIDM_MENU_VIEW);
				var nVerb = GetCommandId(hMenu, s);
				SendCommand(Ctrl, nVerb);
				api.DestroyMenu(hMenu);
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
				} else {
					var FV = te.Ctrl(CTRL_FV);
					Selected = FV.SelectedItems();
				}
				if (Selected && Selected.Count) {
					var ContextMenu = api.ContextMenu(Selected, FV);
					if (ContextMenu) {
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS);
						var nVerb = GetCommandId(hMenu, s, ContextMenu);
						ContextMenu.InvokeCommand(0, te.hwnd, nVerb ? nVerb - 1 : s, null, null, SW_SHOWNORMAL, 0, 0);
						api.DestroyMenu(hMenu);
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
					var ContextMenu = api.ContextMenu(Selected, FV);
					if (ContextMenu) {
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS);
						return g_basic.PopupMenu(hMenu, ContextMenu, pt);
					}
				}
			}
		},

		Background: 
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				var FV = te.Ctrl(CTRL_FV);
				if (FV) {
					var bNoCmd = String(s).toLowerCase() != "cmd";
					var ContextMenu = bNoCmd ? FV.ViewMenu() : api.ContextMenu(FV, FV);
					if (ContextMenu) {
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS);
						var nVerb = bNoCmd ? GetCommandId(hMenu, s, ContextMenu) : 0;
						ContextMenu.InvokeCommand(0, te.hwnd, nVerb ? nVerb - 1 : s, null, null, SW_SHOWNORMAL, 0, 0);
						api.DestroyMenu(hMenu);
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
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS);
						return g_basic.PopupMenu(hMenu, ContextMenu, pt);
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
							s.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(Selected.Item(nCount), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)));
						}
					} else {
						s.push(api.PathQuoteSpaces(api.GetDisplayNameOf(FV, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)));
					}
					clipboardData.setData("text", s.join(" "));
					return S_OK;
				},
				"Run Dialog": function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					api.ShRunDialog(te.hwnd, 0, FV ? FV.FolderItem.Path : null, null, null, 0);
					return S_OK;
				},
				Search: function (Ctrl, pt)
				{
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						var s = InputDialog("Search", IsSearchPath(FV) ? api.GetDisplayNameOf(FV, SHGDN_INFOLDER) : "");
						if (s) {
							FV.FilterView(s);
						} else if (s === "") {
							CancelFilterView(FV);
						}
					}
					return S_OK;
				},
				"Add to Favorites": AddFavoriteEx,
				"Reload Customize": function ()
				{
					te.Reload();
					return S_OK;
				},
				"Load Layout": LoadLayout,
				"Save Layout": SaveLayout,
				"Close Application": function ()
				{
					api.PostMessage(te.hwnd, WM_CLOSE, 0, 0);
					return S_OK;
				}
			},

			Enc: true
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

			List: ["General", "Add-ons", "Menus", "Tabs", "Tree", "List"],

			Enc: true
		},

		Key:
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				api.SetFocus(Ctrl.hwnd);
				wsh.SendKeys(s);
				return S_OK;
			}
		},

		"&Load Layout...":
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				if (s) {
					LoadXml(s);
				} else {
					LoadLayout();
				}
				return S_OK;
			},

			Ref: OpenDialog
		},

		"&Save Layout...":
		{
			Exec: function (Ctrl, s, type, hwnd, pt)
			{
				if (s) {
					SaveXml(s);
				} else {
					SaveLayout();
				}
				return S_OK;
			},

			Ref: OpenDialog
		},

		"Add-ons":
		{
			Cmd: {},
			Enc: true
		},

		Menus:
		{
			Ref: function (s, pt)
			{
				return g_basic.Popup(g_basic.Func.Menus.List, s, pt);
			},

			List: ["Open", "Close", "Separator", "Break", "BarBreak"],

			Enc: true
		}
	},

	Exec: function (Ctrl, s, type, hwnd, pt)
	{
		var fn = g_basic.CmdI(type, s);
		if (!pt) {
			pt = api.Memory("POINT");
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
		} else {
			for (i in Cmd) {
				ar.push(i);
			}
		}
		var hMenu = api.CreatePopupMenu();
		for (i = 0; i < ar.length; i++) {
			if (ar[i]) {
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, i + 1, GetText(ar[i]));
			}
		}
		var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, null);
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
			} else {
				api.DeleteMenu(hMenu, i, MF_BYPOSITION);
			}
		}
		window.g_menu_click = true;
		var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, ContextMenu);
		if (nVerb == 0) {
			api.DestroyMenu(hMenu);
			return 1;
		}
		if (ContextMenu) {
			Verb = ContextMenu.GetCommandString(nVerb - 1, GCS_VERB);
		}
		if (!Verb) {
			Verb = window.g_menu_string;
			if (/\t(.*)$/.test(Verb)) {
				Verb = RegExp.$1;
			}
		}
		api.DestroyMenu(hMenu);
		return Verb;
	}
};

AddEvent("Exec", function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
{
	var fn = g_basic.FuncI(type);
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

AddEvent("MenuState:Tabs:Close Tab", function (Ctrl, pt, mii)
{
	var FV = GetFolderView(Ctrl, pt);
	if (FV && FV.Data.Lock) {
		mii.fMask |= MIIM_STATE;
		mii.fState |= MFS_DISABLED;
	}
});

AddEvent("MenuState:Tabs:Close Tabs on Left", function (Ctrl, pt, mii)
{
	var FV = GetFolderView(Ctrl, pt, true);
	if (FV && FV.Index == 0) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_DISABLED;
	}
});

AddEvent("MenuState:Tabs:Close Tabs on Right", function (Ctrl, pt, mii)
{
	var FV = GetFolderView(Ctrl, pt, true);
	if (FV) {
		var TC = FV.Parent;
		if (FV.Index >= TC.Count - 1) {
			mii.fMask |= MIIM_STATE;
			mii.fState = MFS_DISABLED;
		}
	}
});

AddEvent("MenuState:Tabs:Up", function (Ctrl, pt, mii)
{
	var FV = GetFolderView(Ctrl, pt);
	if (!FV || api.ILIsEmpty(FV)) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_DISABLED;
	}
});

AddEvent("MenuState:Tabs:Lock", function (Ctrl, pt, mii)
{
	var FV = GetFolderView(Ctrl, pt);
	if (FV && FV.Data.Lock) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_CHECKED;
	}
});

AddEvent("MenuState:Tabs:Show frames", function (Ctrl, pt, mii)
{
	if (WINVER < 0x600) {
		mii.fMask = 0;
		return S_OK;
	}
	var FV = GetFolderView(Ctrl, pt);
	if (!FV || FV.Type == CTRL_EB) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_CHECKED;
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
	api.ObjPutI(g_basic.Func, strType, o);
};

AddTypeEx = function (strType, strTitle, fn)
{
	var type = g_basic.FuncI(strType);
	if (type && type.Cmd) {
		api.ObjPutI(type.Cmd, strTitle, fn);
	}
};

AddEvent("OptionRef", function (Id, s, pt)
{
	var fn = g_basic.FuncI(Id);
	if (fn) {
		var r;
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
	if (Id === "") {
		var lines = p.s.split(/\r?\n/);
		for (var i in lines) {
			if (/^([^,]+),(.*)$/.test(lines[i])) {
				var p2 = { s: RegExp.$2 };
				Id = GetSourceText(RegExp.$1);
				OptionEncode(Id, p2);
				lines[i] = [Id, p2.s].join(",");
			}
		}
		p.s = lines.join("\n");
		return S_OK;
	}
	var type = g_basic.FuncI(Id);
	if (type && type.Enc) {
		p.s = GetSourceText(p.s);
		return S_OK;
	}
});

AddEvent("OptionDecode", function (Id, p)
{
	if (Id === "") {
		var lines = p.s.split(/\r?\n/);
		for (var i in lines) {
			if (/^([^,]+),(.*)$/.test(lines[i])) {
				var p2 = { s: RegExp.$2 };
				Id = RegExp.$1;
				OptionDecode(Id, p2);
				lines[i] = [GetText(Id), p2.s].join(",");
			}
		}
		p.s = lines.join("\n");
		return S_OK;
	}
	var type = g_basic.FuncI(Id);
	if (type && type.Enc) {
		var s = GetText(p.s);
		if (GetSourceText(s) == p.s) {
			p.s = GetText(p.s);
			return S_OK;
		}
		return S_OK;
	}
});

AddEvent("Addons.OpenInstead", function (path)
{
	if (api.ILIsParent(g_pidlCP, path, false) && !api.ILIsEqual(path, g_pidlCP) && !api.ILIsEqual(path, ssfCONTROLS)) {
		return S_FALSE;
	}
});

//Init

if (!te.Data) {
	te.Data = te.Object();
	te.Data.CustColors = api.Memory("int", 16);
	te.Data.AddonsData = te.Object();
	te.Data.Fonts = te.Object();
	//Default Value
	te.Data.Tab_Style = TCS_HOTTRACK | TCS_MULTILINE | TCS_RAGGEDRIGHT | TCS_SCROLLOPPOSITE | TCS_HOTTRACK | TCS_TOOLTIPS;
	te.Data.Tab_Align = TCA_TOP;
	te.Data.Tab_TabWidth = 96;
	te.Data.Tab_TabHeight = 0;

	te.Data.View_Type = CTRL_SB;
	te.Data.View_ViewMode = FVM_DETAILS;
	te.Data.View_fFlags = FWF_SHOWSELALWAYS;
	te.Data.View_Options = EBO_SHOWFRAMES | EBO_ALWAYSNAVIGATE;
	te.Data.View_ViewFlags = CDB2GVF_SHOWALLFILES;
	te.Data.View_IconSize = 0;
	te.Data.View_SizeFormat = 0;

	te.Data.Tree_Align = 0;
	te.Data.Tree_Width = 200;
	te.Data.Tree_Style = NSTCS_HASEXPANDOS | NSTCS_SHOWSELECTIONALWAYS | NSTCS_HASLINES | NSTCS_BORDER | NSTCS_NOINFOTIP;
	te.Data.Tree_EnumFlags = SHCONTF_FOLDERS;
	te.Data.Tree_RootStyle = NSTCRS_VISIBLE | NSTCRS_EXPANDED;
	te.Data.Tree_Root = 0;

	te.Data.Conf_TabDefault = true;
	te.Data.Conf_TreeDefault = true;
	te.Data.Conf_ListDefault = true;

	te.Data.Installed = fso.GetParentFolderName(api.GetModuleFileName(null));
	var DataFolder = te.Data.Installed;

	var pf = [ssfPROGRAMFILES, ssfPROGRAMFILESx86];
	var x = api.sizeof("HANDLE") / 4;
	for (var i = 0; i < x; i++) {
		var s = api.GetDisplayNameOf(pf[i], SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
		var l = s.replace(/\s*\(x86\)$/i, "").length;
		if (api.StrCmpNI(s, DataFolder, l) == 0) {
			DataFolder = fso.BuildPath(api.GetDisplayNameOf(ssfAPPDATA, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING), "Tablacus\\Explorer");
			var ParentFolder = fso.GetParentFolderName(DataFolder);
			if (!fso.FolderExists(ParentFolder)) {
				if (fso.CreateFolder(ParentFolder)) {
					CreateFolder2(DataFolder);
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
	} else {
		LoadConfig();
	}
	te.Data.uRegisterId = api.SHChangeNotifyRegister(te.hwnd, SHCNRF_InterruptLevel | SHCNRF_ShellLevel | SHCNRF_NewDelivery, SHCNE_ALLEVENTS, TWM_CHANGENOTIFY, ssfDESKTOP, true);
} else {
	setTimeout(function ()
	{	
		te.UnlockUpdate();
		setTimeout(function ()
		{
			Resize();
		}, 99);
	}, 500);
}

InitCode();
InitMouse();
InitMenus();
LoadLang();
ArrangeAddons();
