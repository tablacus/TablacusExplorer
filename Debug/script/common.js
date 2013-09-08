//Tablacus Explorer

var Ctrl;
var g_temp;
var g_sep = "` ~";
var Handled;
var hwnd;
var pt;
var dataObj = null;
var grfKeyState;
var pdwEffect;
var bDrop;
var Input;
var g_tidNew = null;
var g_dlgOptions;
var eventTE = {};
var g_ptDrag = api.Memory("POINT");
var objHover = null;

FolderMenu = 
{
	arBM: [],
	Items: [],

	Open: function (FolderItem, x, y)
	{
		this.Clear();
		var hMenu = api.CreatePopupMenu();
		this.OpenMenu(hMenu, FolderItem);
		window.g_menu_click = true;
		var Verb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, x, y, te.hwnd, null, null);
		g_popup = null;
		api.DestroyMenu(hMenu);
		Verb = Verb ? this.Items[Verb - 1] : null;
		this.Clear();
		return Verb;
	},

	OpenSubMenu: function (hMenu, wID, hSubMenu)
	{
		this.OpenMenu(api.sscanf(hSubMenu, "%llx"), this.Items[wID - 1], api.sscanf(hMenu, "%llx"), wID);
	},

	OpenMenu: function (hMenu, FolderItem, hParent, wID)
	{
		if (!FolderItem || FolderItem.IsBrowsable) {
			return;
		}
		var bSep = false;
		if (!api.ILIsEmpty(FolderItem)) {
			this.AddMenuItem(hMenu, api.ILRemoveLastID(FolderItem), "../");
			bSep = true;
		}
		var Folder = FolderItem.GetFolder;
		if (Folder) {
			var Items = Folder.Items();
			if (Items) {
				var nCount = Items.Count;
				for (var i = 0; i < nCount; i++) {
					var Item = Items.Item(i);
					if (Item.IsFolder) {
						if (bSep) {
							api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
							bSep = false;
						}
						this.AddMenuItem(hMenu, Item);
						wID = null;
					}
				}
			}
		}
		if (hParent && wID) {
			var mii = api.Memory("MENUITEMINFO");
			mii.cbSize = mii.Size;
			mii.fMask = MIIM_SUBMENU | MIIM_FTYPE;
			api.GetMenuItemInfo(hParent, wID, false, mii);
			mii.hSubMenu = 0;
			mii.fType = mii.fType & ~MF_POPUP;
			api.SetMenuItemInfo(hParent, wID, false, mii);
			api.DestroyMenu(hMenu);
		}
	},

	AddMenuItem: function (hMenu, FolderItem, Name, bSelect)
	{
		var mii = api.Memory("MENUITEMINFO");
		mii.cbSize = mii.Size;
		mii.fMask  = MIIM_ID | MIIM_STRING | MIIM_BITMAP | MIIM_SUBMENU;
		if (bSelect) {
			mii.dwTypeData = ' ' + Name;
		}
		else {
			mii.dwTypeData = ' ' + (Name ? Name + api.GetDisplayNameOf(FolderItem, SHGDN_INFOLDER) : api.GetDisplayNameOf(FolderItem, SHGDN_INFOLDER));
		}
		var image = te.GdiplusBitmap();
		var info = api.Memory("SHFILEINFO");
		api.ShGetFileInfo(FolderItem, 0, info, info.Size, SHGFI_ICON | SHGFI_SMALLICON | SHGFI_PIDL);
		var hIcon = info.hIcon;
		var cl = api.GetSysColor(COLOR_MENU);
		image.FromHICON(hIcon, cl);
		api.DestroyIcon(hIcon);
		mii.hbmpItem = image.GetHBITMAP(cl);
		this.arBM.push(mii.hbmpItem);
		this.Items.push(FolderItem);
		mii.wID = this.Items.length;
		if (!bSelect && api.GetAttributesOf(FolderItem, SFGAO_HASSUBFOLDER | SFGAO_BROWSABLE) == SFGAO_HASSUBFOLDER) {
			mii.hSubMenu = api.CreatePopupMenu();
			api.InsertMenu(mii.hSubMenu, 0, MF_BYPOSITION | MF_STRING, 0, api.sprintf(100, '\tJScript\tFolderMenu.OpenSubMenu("%llx",%d,"%llx")', hMenu, mii.wID, mii.hSubMenu));
		}
		api.InsertMenuItem(hMenu, MAXINT, false, mii);
	},

	Clear: function ()
	{
		while (this.arBM.length) {
			api.DeleteObject(this.arBM.pop());
		}
		this.Items = [];
	},

	Invoke: function (FolderItem)
	{
		if (FolderItem) {
			switch (window.g_menu_button - 0) {
				case 2:
					PopupContextMenu(FolderItem);
					break;
				case 3:
					Navigate(FolderItem, SBSP_NEWBROWSER);
					break;
				default:
					Navigate(FolderItem, OpenMode);
					break;
			}
		}
	}
};

AddEvent = function (Name, fn, priority)
{
	if (Name) {
		if (!eventTE[Name]) {
			eventTE[Name] = [];
		}
		if (priority) {
			eventTE[Name].unshift(fn);
		}
		else {
			eventTE[Name].push(fn);
		}
	}
}

AddEventEx = function (w, Name, fn)
{
	if (w.addEventListener) {
		w.addEventListener(Name, fn, false);
	}
	else if (w.attachEvent){
		w.attachEvent("on" + Name, fn);
	}
}

function ApplyLang(doc)
{
	var LogFont = api.Memory("LOGFONT");
	api.SystemParametersInfo(SPI_GETICONTITLELOGFONT, LogFont.Size, LogFont, 0);
	var FaceName = LogFont.lfFaceName;
	if (doc.body) {
		doc.body.style.fontFamily = FaceName;
		doc.body.style.backgroundColor = 'buttonface';
	}

	var i;
	var Lang = MainWindow.Lang;
	var o = doc.getElementsByTagName("a");
	if (o) {
		for (i = o.length; i--;) {
			var s = Lang[o[i].innerHTML.replace(/&amp;/ig, "&")];
			if (!s) {
				s = o[i].innerHTML;
			}
			o[i].innerHTML = amp2ul(s);
			var s = Lang[o[i].title];
			if (s) {
				o[i].title = s;
			}
			var s = Lang[o[i].alt];
			if (s) {
				o[i].alt = s;
			}
		}
	}
	var h = 0;
	var o = doc.getElementsByTagName("input");
	if (o) {
		for (i = o.length; i--;) {
			if (!h && o[i].type == "text") {
				h = o[i].offsetHeight;
			}
			var s = Lang[o[i].title];
			if (s) {
				o[i].title = s;
			}
			var s = Lang[o[i].alt];
			if (s) {
				o[i].alt = s;
			}
			if (o[i].type == "button") {
				s = Lang[o[i].value];
				if (s) {
					o[i].value = s;
				}
			}
			var s = ImgBase64(o[i], 0);
			if (s != "") {
				o[i].src = s;
				if (o[i].type == "text" && s != "") {
					o[i].style.backgroundImage = "url('" + s + "')";
				}
			}
			o[i].style.fontFamily = FaceName;
		}
	}
	var o = doc.getElementsByTagName("img");
	if (o) {
		for (i = o.length; i--;) {
			var s = Lang[o[i].title];
			if (s) {
				o[i].title = delamp(s);
			}
			var s = Lang[o[i].alt];
			if (s) {
				o[i].alt = delamp(s);
			}
			var s = ImgBase64(o[i], 0);
			if (s != "") {
				o[i].src = s;
			}
		}
	}
	var o = doc.getElementsByTagName("select");
	if (o) {
		for (i = o.length; i--;) {
			o[i].style.fontFamily = FaceName;
			for (var j = 0; j < o[i].length; j++) {
				var s = Lang[o[i][j].text.replace(/^\n/, "").replace(/\n$/, "")];
				if (s) {
					o[i][j].text = s;
				}
			}
		}
	}
	var o = doc.getElementsByTagName("label");
	if (o) {
		for (i = o.length; i--;) {
			var s = Lang[o[i].innerHTML.replace(/&amp;/ig, "&")];
			if (!s) {
				s = o[i].innerHTML;
			}
			o[i].innerHTML = amp2ul(s);
			var s = Lang[o[i].title];
			if (s) {
				o[i].title = s;
			}
			var s = Lang[o[i].alt];
			if (s) {
				o[i].alt = s;
			}
		}
	}
	var o = doc.getElementsByTagName("button");
	if (o) {
		for (i = o.length; i--;) {
			o[i].style.fontFamily = FaceName;
			var s = Lang[o[i].innerHTML.replace(/&amp;/ig, "&")];
			if (!s) {
				s = o[i].innerHTML;
			}
			o[i].innerHTML = amp2ul(s);
			var s = Lang[o[i].title];
			if (s) {
				o[i].title = s;
			}
			var s = Lang[o[i].alt];
			if (s) {
				o[i].alt = s;
			}
		}
	}
	var o = doc.getElementsByTagName("textarea");
	if (o) {
		for (i = o.length; i--;) {
			o[i].style.fontFamily = FaceName;
			o[i].onkeydown = InsertTab;
		}
	}

	var o = doc.getElementsByTagName("form");
	if (o) {
		for (i = o.length; i--;) {
			o[i].onsubmit = function () { return false };
		}
	}

	setTimeout(function ()
	{
		var hwnd = 0;
		do {
			hwnd = api.FindWindowEx(null, hwnd, "Internet Explorer_TridentDlgFrame", null);
			if (hwnd) {
				var hwnd2 = hwnd;
				do {
					hwnd2 = api.GetWindow(hwnd2, GW_OWNER);
					if (te.hwnd == hwnd2) {
						api.SetWindowText(hwnd, GetText(api.GetWindowText(hwnd).replace(/ \-\- .*$/, "")));
						break;
					}
				} while (hwnd2);
			}
		} while (hwnd);
	}, 500);
}

function amp2ul(s)
{
	s = s.replace(/&amp;/ig, "&");
	if (s.match(";")) {
		return s;
	}
	else {
		return s.replace(/&(.)/ig, "<u>$1</u>");
	}
}

function delamp(s)
{
	s = s.replace(/&amp;/ig, "&");
	if (s.match(";")) {
		return s;
	}
	else {
		return s.replace(/&/ig, "");
	}
}

function ImgBase64(o, index)
{
	var s = MakeImgSrc(o.src, index, false, o.height, o.getAttribute("bitmap"), o.getAttribute("icon"));
	if (s) {
		o.removeAttribute("bitmap");
		o.removeAttribute("icon");
	}
	return s;
}

function MakeImgSrc(src, index, bSrc, h, strBitmap, strIcon)
{
	var fn;
	if (!document.documentMode) {
		var value = src.match(/^bitmap:(.*)/i) ? RegExp.$1 : strBitmap;
		if (value) {
			fn = fso.BuildPath(te.Data.DataFolder, "cache\\bitmap\\" + value.replace(/[:\\\/]/g, "$") + ".png");
		}
		else {
			value = src.match(/^icon:(.*)/i) ? RegExp.$1 : strIcon;
			if (value) {
				fn = fso.BuildPath(te.Data.DataFolder, "cache\\icon\\" + value.replace(/[:\\\/]/g, "$") + ".png");
			}
			else if (src && !api.PathMatchSpec(src, "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.ico;data:*")) {
				src = src.replace(/^file:\/\/\//i, "").replace(/\//g, "\\");
				fn = fso.BuildPath(te.Data.DataFolder, "cache\\file\\" + src.replace(/[:\\\/]/g, "$") + ".png");
			}
		}
		if (fn && fso.FileExists(fn)) {
			return fn;
		}
	}
	var image = MakeImgData(src, index, h, strBitmap, strIcon);
	if (image) {
		if (document.documentMode) {
			return image.DataURI("image/png");
		}
		if (fn) {
			try {
				image.Save(fn);
			} catch (e) {
			}
			return fn;
		}
	}
	return bSrc ? src : "";
}

function MakeImgData(src, index, h, strBitmap, strIcon)
{
	var Result = null;
	var image = te.GdiplusBitmap();
	var value = src.match(/^bitmap:(.*)/i) ? RegExp.$1 : strBitmap;
	if (value) {
		var icon = value.split(",");
		var hModule = LoadImgDll(icon, index);
		if (hModule) {
			var himl = api.ImageList_LoadImage(hModule, isFinite(icon[index * 4 + 1]) ? api.LowPart(icon[index * 4 + 1]) : icon[index * 4 + 1], icon[index * 4 + 2], 0, CLR_DEFAULT, IMAGE_BITMAP, LR_CREATEDIBSECTION);
			if (himl) {
				var hIcon = api.ImageList_GetIcon(himl, icon[index * 4 + 3], ILD_NORMAL);
				if (hIcon) {
					image.FromHICON(hIcon, api.GetSysColor(COLOR_BTNFACE));
					Result = image;
					api.DestroyIcon(hIcon);
				}
				api.ImageList_Destroy(himl);
			}
			api.FreeLibrary(hModule);
			return Result;
		}
	}
	value = src.match(/^icon:(.*)/i) ? RegExp.$1 : strIcon;
	if (value) {
		var icon = value.split(",");
		var phIcon = api.Memory("HANDLE");
		if (icon[index * 4 + 2] > 16) {
			api.ExtractIconEx(icon[index * 4], icon[index * 4 + 1], phIcon, null, 1);
		}
		else {
			api.ExtractIconEx(icon[index * 4], icon[index * 4 + 1], null, phIcon, 1);
		}
		var hIcon = phIcon[0];
		if (hIcon) {
			image.FromHICON(hIcon, api.GetSysColor(COLOR_BTNFACE));
			api.DestroyIcon(hIcon);
			return image;
		}
	}
	if (src && !api.PathMatchSpec(src, "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.ico;data:*")) {
		var info = api.Memory("SHFILEINFO");
		var pidl = api.ILCreateFromPath(api.PathUnquoteSpaces(src));
		if (pidl) {
			var uFlags = SHGFI_PIDL | SHGFI_ICON;
			if (h && h <= 16) {
				uFlags |= SHGFI_SMALLICON;
			}
			api.ShGetFileInfo(pidl, 0, info, info.Size, uFlags);
			image.FromHICON(info.hIcon, api.GetSysColor(COLOR_BTNFACE));
			Result = image;
			api.DestroyIcon(info.hIcon);
		}
	}
	return Result;
}

LoadImgDll = function (icon, index)
{
	var hModule = api.LoadLibraryEx(fso.BuildPath(system32, icon[index * 4]), 0, LOAD_LIBRARY_AS_DATAFILE);
	if (!hModule && api.strcmpi(icon[index * 4], "ieframe.dll") == 0) {
		if (osInfo.dwMinorVersion || osInfo.dwMajorVersion > 5) {
			hModule = api.LoadLibraryEx(fso.BuildPath(system32, "shell32.dll"), 0, LOAD_LIBRARY_AS_DATAFILE);
		}
		else {
			hModule = api.LoadLibraryEx(fso.BuildPath(system32, "browseui.dll"), 0, LOAD_LIBRARY_AS_DATAFILE);
			icon[index * 4 + 1] = (icon[index * 4 + 1] < 210 ? 62 : 63) + (icon[index * 4 + 1] & ~1);
			if (icon[index * 4 + 2] > 20) {
				icon[index * 4 + 2] = 20;
			}
		}
	}
	return hModule;
}

GetText = function (id)
{
	try {
		id = id.replace(/&amp;/g, "&");
		var s = MainWindow.Lang[id];
		if (s) {
			return s;
		}
	} catch (e) {}
	return id;
}

function LoadLang2(filename)
{
	var xml = te.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	xml.load(filename);
	var items = xml.getElementsByTagName('text');
	for (var i = 0; i < items.length; i++) {
		var item = items[i];
		var s = item.getAttribute("s").replace("\\t", "\t").replace("\\n", "\n");
		var v = item.text.replace("\\t", "\t").replace("\\n", "\n");
		MainWindow.Lang[s] = v;
		if (s.match(/&|\.\.\.$/)) {
			MainWindow.Lang[StripAmp(s)] = StripAmp(v);
		}
	}
}

LoadXml = function (filename)
{
	te.LockUpdate();
	var cTC = te.Ctrls(CTRL_TC);
	for (i in cTC) {
		cTC[i].Close();
	}
	var xml = filename;
	if (typeof(filename) == "string") {
		var xml = te.CreateObject("Msxml2.DOMDocument");
		xml.async = false;
		xml.load(filename);
	}
	var items = xml.getElementsByTagName('Ctrl');
	for (var i = 0; i < items.length; i++) {
		var item = items[i];
		switch(item.getAttribute("Type") - 0) {
			case CTRL_TC:
				var TC = te.CreateCtrl(CTRL_TC, item.getAttribute("Left"), item.getAttribute("Top"), item.getAttribute("Width"), item.getAttribute("Height"), item.getAttribute("Style"), item.getAttribute("Align"), item.getAttribute("TabWidth"), item.getAttribute("TabHeight"));
				TC.Data.Group = api.LowPart(item.getAttribute("Group"));
				var tabs = item.getElementsByTagName('Ctrl');
				for (var i2 = 0; i2 < tabs.length; i2++) {
					var tab = tabs[i2];
					var Path = tab.getAttribute("Path");
					var logs = tab.getElementsByTagName('Log');
					var nLogCount = logs.length;
					if (nLogCount > 1) {
						Path = te.FolderItems();
						for (var i3 = 0; i3 < nLogCount; i3++) {
							Path.AddItem(logs[i3].getAttribute("Path"));
						}
						Path.Index = tab.getAttribute("LogIndex");
					}
					var FV = TC.Selected.Navigate2(Path, SBSP_NEWBROWSER, tab.getAttribute("Type"), tab.getAttribute("ViewMode"), tab.getAttribute("FolderFlags"), tab.getAttribute("Options"), tab.getAttribute("ViewFlags"), tab.getAttribute("IconSize"), tab.getAttribute("Align"), tab.getAttribute("Width"), tab.getAttribute("Flags"), tab.getAttribute("EnumFlags"), tab.getAttribute("RootStyle"), tab.getAttribute("Root"));
					ChangeTabName(FV);
					FV.Data.Lock = api.LowPart(tab.getAttribute("Lock")) != 0;
					Lock(TC, i2, false);
				}
				TC.SelectedIndex = item.getAttribute("SelectedIndex");
				TC.Visible = api.LowPart(item.getAttribute("Visible"));
				break;
		}
	}
	te.UnlockUpdate();
}

SaveXml = function (filename, all)
{
	var xml = CreateXml();
	var root = xml.createElement("TablacusExplorer");

	if (all) {
		var item = xml.createElement("Window");
		var CmdShow = SW_SHOWNORMAL;
		var hwnd = te.hwnd;
		if (api.IsZoomed(hwnd)) {
			CmdShow = SW_SHOWMAXIMIZED;
		}
		api.ShowWindow(hwnd, SW_SHOWNORMAL);
		rc = api.Memory("RECT");
		api.GetWindowRect(hwnd, rc);
		item.setAttribute("Left", rc.left);
		item.setAttribute("Top", rc.top);
		item.setAttribute("Width", rc.right - rc.left);
		item.setAttribute("Height", rc.bottom - rc.top);
		item.setAttribute("CmdShow", CmdShow);
		root.appendChild(item);
		item = null;
	}
	var cTC = te.Ctrls(CTRL_TC);
	for (var i in cTC) {
		var Ctrl = cTC[i];
		var item = xml.createElement("Ctrl");
		item.setAttribute("Type", Ctrl.Type);
		item.setAttribute("Left", Ctrl.Left);
		item.setAttribute("Top", Ctrl.Top);
		item.setAttribute("Width", Ctrl.Width);
		item.setAttribute("Height", Ctrl.Height);
		item.setAttribute("Style", Ctrl.Style);
		item.setAttribute("Align", Ctrl.Align);
		item.setAttribute("TabWidth", Ctrl.TabWidth);
		item.setAttribute("TabHeight", Ctrl.TabHeight);
		item.setAttribute("SelectedIndex", Ctrl.SelectedIndex);
		item.setAttribute("Visible", api.LowPart(Ctrl.Visible));
		item.setAttribute("Group", api.LowPart(Ctrl.Data.Group));

		var bEmpty = true;
		var nCount2 = Ctrl.Count;
		for (var i2 in Ctrl) {
			var FV = Ctrl[i2];
			var path = api.GetDisplayNameOf(FV, SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
			var bSave = !all || IsSavePath(path);
			if (bSave || (bEmpty && i2 == nCount2 - 1)) {
				if (!bSave) {
					path = HOME_PATH;
				}
				var item2 = xml.createElement("Ctrl");
				item2.setAttribute("Type", FV.Type);
				item2.setAttribute("Path", path);
				item2.setAttribute("FolderFlags", FV.FolderFlags);
				item2.setAttribute("ViewMode", FV.CurrentViewMode);
				item2.setAttribute("IconSize", FV.IconSize);
				item2.setAttribute("Options", FV.Options);
				item2.setAttribute("ViewFlags", FV.ViewFlags);
				item2.setAttribute("Lock", api.LowPart(FV.Data.Lock));
				var TV = FV.TreeView;
				if (TV) {
					item2.setAttribute("Align", FV.TreeView.Align);
					item2.setAttribute("Width", FV.TreeView.Width);
					item2.setAttribute("Flags", FV.TreeView.Style);
					item2.setAttribute("EnumFlags", FV.TreeView.EnumFlags);
					item2.setAttribute("RootStyle", FV.TreeView.RootStyle);
					item2.setAttribute("Root", String(FV.TreeView.Root));
				}
				var TL = FV.History;
				if (TL) {
					if (TL.Count > 1) {
						var bLogSaved = false;
						var nLogIndex = TL.Index;
						for (var i3 in TL) {
							path = api.GetDisplayNameOf(TL[i3], SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
							if (IsSavePath(path)) {
								var item3 = xml.createElement("Log");
								item3.setAttribute("Path", path);
								item2.appendChild(item3);
								bLogSaved = true;
							}
							else if (i3 < nLogIndex) {
								nLogIndex--;
							}
						}
						if (bLogSaved) {
							item2.setAttribute("LogIndex", nLogIndex);
						}
					}
				}
				item.appendChild(item2);
				bEmpty = false;
			}
		}
		root.appendChild(item);
	}
	if (all) {
		for (var i in te.Data) {
			if (i.match(/^(Tab|Tree|View|Conf)_(.*)/)) {
				var item = xml.createElement(RegExp.$1);
				item.setAttribute("Id", RegExp.$2);
				item.text = te.Data[i];
				if (item.text != "") {
					root.appendChild(item);
				}
			}
		}
	}
	xml.appendChild(root);
	xml.save(filename);
}

GetKeyKey = function (strKey)
{
	var nShift = api.sscanf(strKey, "$%x");
	if (nShift) {
		return nShift;
	}
	strKey = strKey.toUpperCase();
	for (var j in MainWindow.g_KeyState) {
		var s = MainWindow.g_KeyState[j][0].toUpperCase() + "+";
		if (strKey.match(s)) {
			strKey = strKey.replace(s, "");
			nShift |= MainWindow.g_KeyState[j][1];
		}
	}
	return nShift | MainWindow.g_KeyCode[strKey];
}

GetKeyName = function (strKey)
{
	var nKey = api.sscanf(strKey, "$%x");
	if (nKey) {
		strKey = "";
		for (var j in MainWindow.g_KeyState) {
			if (nKey & MainWindow.g_KeyState[j][1]) {
				nKey -= MainWindow.g_KeyState[j][1];
				strKey += MainWindow.g_KeyState[j][0] + "+";
			}
		}
		strKey += api.GetKeyNameText((nKey & 0x17f) << 16);
	}
	return strKey;
}

GetKeyShift = function ()
{
	var nShift = 0;
	var n = 0x1000;
	var vka = [VK_SHIFT, VK_CONTROL, VK_MENU, VK_LWIN];
	for (var i in vka) {
		if (api.GetKeyState(vka[i]) < 0) {
			nShift += n;
		}
		n *= 2;
	}
	return nShift;
}

function SetKeyData(mode, strKey, path, type, km, o)
{
	var s = "";
	if (!o) {
		o = te.Data;
		s = km;
	}
	if (km == "Key") {
		o[s + mode][GetKeyKey(strKey)] = [path, type];
	}
	else {
		o[s + mode][strKey] = [path, type];
	}
}

function SendShortcutKeyFV(Key)
{
	var FV = te.Ctrl(CTRL_FV);
	if (FV) {
		var KeyState = api.Memory("KEYSTATE");
		api.GetKeyboardState(KeyState);
		var KeyCtrl = KeyState.Read(VK_CONTROL, VT_UI1);
		KeyState.Write(VK_CONTROL, VT_UI1, 0x80);
		api.SetKeyboardState(KeyState);
		FV.TranslateAccelerator(0, WM_KEYDOWN, Key.charCodeAt(0), 0);
		FV.TranslateAccelerator(0, WM_KEYUP, Key.charCodeAt(0), 0);
		KeyState.Write(VK_CONTROL, VT_UI1, KeyCtrl);
		api.SetKeyboardState(KeyState);
	}
}

CreateTab = function ()
{
	var FV = te.Ctrl(CTRL_FV);
	Navigate(FV ? FV.FolderItem : HOME_PATH, SBSP_NEWBROWSER);
}

Navigate = function (Path, wFlags)
{
	var FV = te.Ctrl(CTRL_FV);
	if (!FV) {
		var TC = te.CreateCtrl(CTRL_TC, 0, 0, "100%", "100%", te.Data.Tab_Style, te.Data.Tab_Align, te.Data.Tab_TabWidth, te.Data.Tab_TabHeight);
		FV = TC.Selected;
	}
	NavigateFV(FV, Path, wFlags);
}

NavigateFV = function (FV, Path, wFlags)
{
	if (FV) {
		var Focus = null;
		if (typeof(Path) == "string") {
			if (Path.match(/\?|\*/) && !Path.match(/^::{/)) {
				FV.FilterView = Path;
				FV.Refresh();
				return;
			}
			if (Path.match(/%([^%]+)%/)) {
				Path = wsh.ExpandEnvironmentStrings(Path);
			}
			if (Path.length > 3 && Path.match(/\\$/)) {
				Focus = Path.replace(/\\$/, '');
				Path = fso.GetParentFolderName(Path);
			}
		}
		FV.Data.Filter = null;
		FV.OnIncludeObject = null;
		FV.Navigate(Path, wFlags);
		if (Focus) {
			setTimeout(function () {
				var FV = te.Ctrl(CTRL_FV);
				FV.SelectItem(Focus, SVSI_FOCUSED | SVSI_ENSUREVISIBLE);
			}, 100);
		}
	}
}

IsDrag = function (pt1, pt2)
{
	if (pt1 && pt2) {
		return (Math.abs(pt1.x - pt2.x) > api.GetSystemMetrics(SM_CXDRAG) | Math.abs(pt1.y - pt2.y) > api.GetSystemMetrics(SM_CYDRAG));
	}
	return false;
}

ChangeTab = function (TC, nMove)
{
	var nCount = TC.Count;
	TC.SelectedIndex = (TC.SelectedIndex + nCount + nMove) % nCount;
}

ShowOptions = function (s)
{
	try {
		if (g_dlgOptions && g_dlgOptions.window) {
			g_dlgOptions.SetTab(s);
			return;
		}
	}
	catch (e) {
		g_dlgOptions = null;
	}
	g_dlgOptions = showModelessDialog("options.html", {MainWindow: MainWindow, Data: s} , "dialogWidth: 640px; dialogHeight: 480px; resizable: yes; status: 0");
}

LoadLayout = function ()
{
	var commdlg = te.CommonDialog();
	commdlg.InitDir = fso.BuildPath(te.Data.DataFolder, "layout");
	commdlg.Filter = "XML Files|*.xml|All Files|*.*";
	commdlg.Flags = OFN_FILEMUSTEXIST;
	if (commdlg.ShowOpen()) {
		LoadXml(commdlg.FileName);
	}
}

SaveLayout = function ()
{
	var commdlg = te.CommonDialog();
	commdlg.InitDir = fso.BuildPath(te.Data.DataFolder, "layout");
	commdlg.Filter = "XML Files|*.xml|All Files|*.*";
	commdlg.DefExt = "xml";
	commdlg.Flags = OFN_OVERWRITEPROMPT;
	if (commdlg.ShowSave()) {
		SaveXml(commdlg.FileName);
	}
}

GetPos = function (o, bScreen, bAbs, bPanel)
{
	var x = (bScreen ? screenLeft : 0);
	var y = (bScreen ? screenTop : 0);

	while (o) {
		if (bAbs || !bPanel || api.strcmpi(o.style.position, "absolute")) {
			x += o.offsetLeft - (bAbs ? 0 : o.scrollLeft);
			y += o.offsetTop - (bAbs ? 0 : o.scrollTop);
			o = o.offsetParent;
		}
		else {
			break;
		}
	}
	var pt = api.Memory("POINT");
	pt.x = x;
	pt.y = y;
	return pt;
}

HitTest = function (o, pt)
{
	if (o) {
		var p = GetPos(o, true);
		if (pt.x >= p.x && pt.x < p.x + o.offsetWidth && pt.y >= p.y && pt.y < p.y + o.offsetHeight) {
			o = o.offsetParent;
			p = GetPos(o, true, true);
			return pt.x >= p.x && pt.x < p.x + o.offsetWidth && pt.y >= p.y && pt.y < p.y + o.offsetHeight;
		}
	}
	return false;
}

DeleteItem = function (path)
{
	api.SHFileOperation(FO_DELETE, path, null, FOF_SILENT | FOF_ALLOWUNDO | FOF_NOCONFIRMATION, false);
}

IsExists = function (path)
{
	var FindData = api.Memory("WIN32_FIND_DATA");
	var hFind = api.FindFirstFile(path, FindData);
	api.FindClose(hFind);
	return hFind != INVALID_HANDLE_VALUE;
}

CreateNew = function (path, fn)
{
	clearTimeout(g_tidNew);
	if (fn && !IsExists(path)) {
		try {
			fn(path);
		} catch (e) {
			if (path.match(/^[A-Z]:\\|^\\/i)) {
				var s = fso.BuildPath(wsh.ExpandEnvironmentStrings("%TEMP%"), fso.GetFileName(path));
				DeleteItem(s);
				fn(s);
				sha.NameSpace(fso.GetParentFolderName(path)).MoveHere(s, FOF_SILENT | FOF_NOCONFIRMATION);
			}
		}
	}
	g_tidNew = setTimeout(function ()
	{
		var FV = te.Ctrl(CTRL_FV);
		if (FV) {
			if (api.ILIsEqual(FV, fso.GetParentFolderName(path))) {
				var FolderItem = api.ILCreateFromPath(path);
				FV.SelectItem(FolderItem, SVSI_SELECT | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_FOCUSED);
			}
		}
	}, 1000);
}

CreateFolder = function (path)
{
	CreateNew(path, function (strPath)
	{
		fso.CreateFolder(strPath);
	});
}

CreateFile = function (path)
{
	CreateNew(path, function (strPath)
	{
		fso.CreateTextFile(strPath).Close();
	});
}

CreateFolder2 = function (path)
{
	if (!fso.FolderExists(path)) {
		CreateFolder(path);
	}
}

GetConsts = function (s)
{
	var Result = window[s.replace(/\s/, "")];
	if (Result !== undefined) {
		return Result;
	}
	return s;
}

Navigate2 = function (path, NewTab)
{
	var a = path.toString().split("\n");
	for (var i in a) {
		var s = a[i].replace(/^\s+/, "");
		if (s != "") {
			Navigate(s, NewTab);
			NewTab |= SBSP_NEWBROWSER;
		}
	}
}

ExecOpen = function (Ctrl, s, type, hwnd, pt, NewTab)
{
	var line = s.split("\n");
	for (var i = 0; i < line.length; i++) {
		if (line[i] != "") {
			Navigate2(ExtractPath(Ctrl, line[i]), NewTab);
			NewTab |= SBSP_NEWBROWSER;
		}
	}
	return S_OK;
}

DropOpen = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
{
	var line = s.split("\n");
	var hr = E_FAIL;
	var path = ExtractPath(Ctrl, line[0]);
	if (!api.ILIsEqual(dataObj.Item(-1), path)) {
		var DropTarget = api.DropTarget(path);
		if (DropTarget) {
			if (!pdwEffect) {
				pdwEffect = dataObj.pdwEffect;
			}
			pdwEffect[0] = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
			hr = bDrop ? DropTarget.Drop(dataObj, grfKeyState, pt, pdwEffect) : DropTarget.DragOver(dataObj, grfKeyState, pt, pdwEffect);
		}
	}
	return hr;
}

Exec = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
{
	window.Ctrl = Ctrl;
	window.hwnd = hwnd;
	window.dataObj = dataObj;
	window.grfKeyState = grfKeyState;
	window.pdwEffect = pdwEffect;
	window.bDrop = bDrop;
	if (pt) {
		window.pt = pt;
		te.Data.pt = pt;
	}
	else {
		window.pt = te.Data.pt;
	}
	window.Handled = S_OK;

	if (api.StrCmpI(type, "Func") == 0) {
		return s(Ctrl, pt, hwnd, dataObj, grfKeyState, pdwEffect, bDrop);
	}
	for (var i in eventTE.Exec) {
		var hr = eventTE.Exec[i](Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	return window.Handled;
}

ExecScriptEx = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
{
	if (api.StrCmpI(type, "JScript") == 0) {
		var r = (new Function(s))(Ctrl, pt, hwnd, dataObj, grfKeyState, pdwEffect, bDrop);
		return isFinite(r) ? r : window.Handled;
	}
	execScript(s, type);
	return window.Handled;
}

DropScript = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
{
	if (!pdwEffect) {
		pdwEffect = api.Memory("DWORD");
	}
	if (s.match("EnableDragDrop")) {
		return ExecScriptEx(Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop);
	}
	pdwEffect[0] = DROPEFFECT_NONE;
	return E_NOTIMPL;
}

ExtractPath = function (Ctrl, s)
{
	return api.PathUnquoteSpaces(ExtractMacro(Ctrl, GetConsts(s)));
}

ExtractMacro = function (Ctrl, s)
{
	if (typeof(s) == "string") {
		for (var i in eventTE.ReplaceMacro) {
			var re = eventTE.ReplaceMacro[i][0];
			if (s.match(re)) {
				s = s.replace(re, eventTE.ReplaceMacro[i][1](Ctrl));
			}
		}
		for (var i in eventTE.ExtractMacro) {
			var re = eventTE.ExtractMacro[i][0];
			if (s.match(re)) {
				s = eventTE.ExtractMacro[i][1](Ctrl, s, re);
			}
		}
	}
	return s;
}

AddEvent("ReplaceMacro", [/%Selected%/ig, function(Ctrl)
{
	var ar = [];
	if (!Ctrl || (Ctrl.Type != CTRL_SB && Ctrl.Type != CTRL_EB)) {
		Ctrl = te.Ctrl(CTRL_FV);
	}
	if (Ctrl) {
		var Selected = Ctrl.SelectedItems();
		if (Selected) {
			for (var i = Selected.Count; i > 0; ar.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(Selected.Item(--i), SHGDN_FORPARSING)))) {
			}
		}
	}
	return ar.join(" ");
}]);

AddEvent("ReplaceMacro", [/%Current%/ig, function(Ctrl)
{
	var strSel = "";
	if (!Ctrl || (Ctrl.Type != CTRL_SB && o.Type != CTRL_EB)) {
		Ctrl = te.Ctrl(CTRL_FV);
	}
	if (Ctrl) {
		strSel = api.PathQuoteSpaces(api.GetDisplayNameOf(Ctrl, SHGDN_FORPARSING));
	}
	return strSel;
}]);

AddEvent("ReplaceMacro", [/%TreeSelected%/ig, function(Ctrl)
{
	var strSel = "";
	if (!Ctrl || Ctrl.Type != CTRL_TV) {
		Ctrl = te.Ctrl(CTRL_TV);
	}
	if (Ctrl) {
		strSel = api.PathQuoteSpaces(api.GetDisplayNameOf(Ctrl.SelectedItem, SHGDN_FORPARSING));
	}
	return strSel;
}]);

AddEvent("ReplaceMacro", [/%Installed%/ig, function(Ctrl)
{
	return fso.GetDriveName(api.GetModuleFileName(null));
}]);

OpenMenu = function (items, SelItem)
{
	var arMenu;
	var path = "";
	if (SelItem) {
		if (typeof(SelItem) != "object") {
			path = SelItem;	
		}
		else {
			path = api.GetDisplayNameOf(SelItem, SHGDN_FORPARSINGEX | SHGDN_FORPARSING) + "";
			arMenu = OpenMenu(items, path);
			if (arMenu.length) {
				return arMenu;
			}
			if (!SelItem.IsFolder) {
				var FindData = api.Memory("WIN32_FIND_DATA");
				api.SHGetDataFromIDList(SelItem, SHGDFIL_FINDDATA, FindData, FindData.Size);
				if ((FindData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
					return arMenu;
				}
			}
			path += ".folder";
		}
	}
	arMenu = [];
	var i = items.length;
	while (--i >= 0) {
		if (SelItem) {
			var s = items[i].getAttribute("Filter");
			if (s.charAt(0) != '/') {
				if (api.PathMatchSpec(path, s)) {
					arMenu.unshift(i);
				}
			}
			else {
				var j = s.lastIndexOf("/");
				if (j > 1 && path.match(new RegExp(s.substr(1, j - 1), s.substr(j + 1)))) {
					arMenu.unshift(i);
				}
			}
		}
		else if (items[i].getAttribute("Filter") == "") {
			arMenu.unshift(i);
		}
	}
	return arMenu;
}

ExecMenu3 = function (Ctrl, Name, x, y)
{
	window.Ctrl = Ctrl;
	setTimeout(function () {
		ExecMenu2(Name, x, y);
	}, 100);;
}

ExecMenu2 = function (Name, x, y)
{
	pt = api.Memory("POINT");
	pt.x = x;
	pt.y = y;
	ExecMenu(Ctrl, Name, pt, 0);
}

ExecMenu = function (Ctrl, Name, pt, Mode)
{
	var items = null;
	var menus = te.Data.xmlMenus.getElementsByTagName(Name);
	if (menus && menus.length) {
		items = menus[0].getElementsByTagName("Item");
	}
	var uCMF = Ctrl.Type != CTRL_TV ? CMF_NORMAL | CMF_CANRENAME : CMF_EXPLORE | CMF_CANRENAME;
	if (api.GetKeyState(VK_SHIFT) < 0) {
		uCMF |= CMF_EXTENDEDVERBS;
	}
	var ar = GetSelectedArray(Ctrl, pt);
	var Selected = ar.shift();
	var SelItem = ar.shift();
	var FV = ar.shift();

	var bSel = true;
	switch(Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
			FV = Ctrl;
			break;
		case CTRL_TC:
			FV = Ctrl.Item(Ctrl.HitTest(pt, TCHT_ONITEM));
			bSel = false;
			break;
		case CTRL_TV:
			SelItem = Ctrl.SelectedItem;
			uCMF = CMF_EXPLORE | CMF_CANRENAME;
			break;
		case CTRL_WB:
			SelItem = window.Input;
			break;
		default:
			FV = te.Ctrl(CTRL_FV);
			break;
	}
	if (FV) {
		if (bSel) {
	 		Selected = FV.SelectedItems();
		}
		if (Selected && Selected.Count) {
			SelItem = Selected.Item(0);
		}
		else {
			SelItem = FV.FolderItem;
		}
	}
	ExtraMenuCommand = [];
	var arMenu;
	var item;
	if (items) {
		arMenu = OpenMenu(items, SelItem);
		if (arMenu.length) {
			item = items[arMenu[0]];
		}
		var nBase = api.LowPart(menus[0].getAttribute("Base"));
		if (arMenu.length > 1 || nBase != 1) {
			var hMenu;
			var ContextMenu = null;
			switch (nBase) {
				case 2:
				case 4:
					var Items = Selected;
					if (!FV && (!Items || !Items.Count)) {
						Items = SelItem;
					}
					hMenu = api.CreatePopupMenu();
					if (nBase == 2 || Items) {
						ContextMenu = api.ContextMenu(Items);
						if (ContextMenu) {
							ContextMenu.QueryContextMenu(hMenu, 0, 0x1001, 0x6FFF, uCMF);
						}
					}
					else if (FV) {
						ContextMenu = FV.ViewMenu();
						if (ContextMenu) {
							ContextMenu.QueryContextMenu(hMenu, 0, 0x1001, 0x6FFF, uCMF);
							var mii = api.Memory("MENUITEMINFO");
							mii.cbSize = mii.Size;
							mii.fMask = MIIM_FTYPE | MIIM_SUBMENU;
							for (var i = api.GetMenuItemCount(hMenu); i--;) {
								api.GetMenuItemInfo(hMenu, 0, true, mii);
								if (mii.hSubMenu || mii.fType & MFT_SEPARATOR) {
									api.DeleteMenu(hMenu, 0, MF_BYPOSITION);
								}
								else {
									break;
								}
							}
						}
					}
					break;
				case 3:
					hMenu = api.CreatePopupMenu();
					if (FV) {
						ContextMenu = FV.ViewMenu();
						if (ContextMenu) {
							ContextMenu.QueryContextMenu(hMenu, 0, 0x1001, 0x6FFF, uCMF);
						}
					}
					break;
				case 5:
				case 6:
					var id = nBase == 5 ? FCIDM_MENU_EDIT : FCIDM_MENU_VIEW;
					if (FV) {
						ContextMenu = FV.ViewMenu();
						if (ContextMenu) {
							hMenu = api.CreatePopupMenu();
							ContextMenu.QueryContextMenu(hMenu, 0, 0x1001, 0x6FFF, uCMF);
							var hMenu2 = te.MainMenu(id);
							var oMenu = {};
							var oMenu2 = {};

							for (var i = api.GetMenuItemCount(hMenu2); i--;) {
								var mii = api.Memory("MENUITEMINFO");
								mii.cbSize = mii.Size;
								mii.fMask  = MIIM_SUBMENU;
								var s = api.GetMenuString(hMenu2, i, MF_BYPOSITION);
								if (s) {
									api.GetMenuItemInfo(hMenu2, i, true, mii);
									oMenu2[s] = mii.hSubMenu;
								}
							}
							MenuDbInit(hMenu, oMenu, oMenu2);
							MenuDbReplace(hMenu, oMenu, hMenu2);
						}
					}
					else {
						hMenu = te.MainMenu(id);
					}
					break;
				case 7:
					hMenu = api.CreatePopupMenu();
					var dir = [te.About, -1, fso.BuildPath(te.Data.DataFolder, "config") , null, ssfDRIVES, ssfNETHOOD, ssfWINDOWS, ssfSYSTEM, ssfPROGRAMFILES];
					if (api.sizeof("HANDLE") > 4) {
						dir.push(ssfPROGRAMFILESx86);
					}
					dir = dir.concat(["%TEMP%", ssfPERSONAL, ssfSTARTMENU, ssfPROGRAMS, ssfSTARTUP, ssfSENDTO, ssfAPPDATA, ssfFAVORITES, ssfRECENT, ssfHISTORY, ssfDESKTOPDIRECTORY, ssfCONTROLS, ssfTEMPLATES, ssfFONTS, ssfPRINTERS, ssfBITBUCKET]);

					for (var i = 0; i < dir.length; i++) {
						var s = dir[i];
						if (api.strcmpni(s, '%', 1) == 0) {
							s = wsh.ExpandEnvironmentStrings(s);
							dir[i] = s;
						}
						if (s === null) {
							api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
						}
						else {
							if (i) {
								s = (i > 1) ? api.GetDisplayNameOf(s, SHGDN_INFOLDER) : GetText("Check for updates");
							}
							if (s) {
								api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, i + 0x1001, s);
							}
							if (i == 0) {
								dir[0] = api.GetModuleFileName(null) + "\\";
							}
						}
					}
					ExtraMenuCommand[0x1002] = CheckUpdate;
					break;
				case 8:
					hMenu = api.CreatePopupMenu();
					api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 0x1001, GetText("&Add to Favorites..."));
					ExtraMenuCommand[0x1001] = function (Ctrl, pt) { AddFavorite(); };
					api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
					break;
				default:
					hMenu = api.CreatePopupMenu();
					break;
			}
			if (nBase < 5) {
				var mii = api.Memory("MENUITEMINFO");
				mii.cbSize = mii.Size;
				mii.fMask = MIIM_FTYPE;
				for (var i = api.GetMenuItemCount(hMenu); i--;) {
					api.GetMenuItemInfo(hMenu, i, true, mii);
					if (mii.fType & MFT_SEPARATOR) {
						api.RemoveMenu(hMenu, i, MF_BYPOSITION);
					}
					else {
						break;
					}
				}
			}
			MakeMenus(hMenu, menus, arMenu, items);
			var nPos = arMenu.length;
			for (var i in eventTE[Name]) {
				nPos = eventTE[Name][i](Ctrl, hMenu, nPos, Selected, SelItem);
			}
			if (ExtraMenus[Name]) {
				ExtraMenus[Name](Ctrl, hMenu, nPos, Selected, SelItem);
			}
			if (!pt) {
				pt = api.Memory("POINT");
				pt.x = -1;
				pt.y = -1;
			}
			if (pt.x == -1 && pt.y == -1) {
				switch (Ctrl.Type) {
					case CTRL_SB:
					case CTRL_EB:
						Ctrl.GetItemPosition(SelItem, pt);
						api.ClientToScreen(Ctrl.hwnd, pt);
						break;
					default:
						api.ClientToScreen(te.hwnd, pt);
						break;
				}
			}
			var nVerb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, ContextMenu);
			if (ExtraMenuCommand[nVerb]) {
				ExtraMenuCommand[nVerb](Ctrl, pt);
				nVerb = 0;
			}
			if (nVerb > 0x1000 && nVerb < 0x7000) {
				if (ContextMenu) {
					var hr = S_FALSE;
					if (ContextMenu.InvokeCommand(0, te.hwnd, nVerb - 0x1001, null, null, SW_SHOWNORMAL, 0, 0) == S_OK) {
						nVerb = 0;
					}
					else if (nVerb == CommandID_RENAME + 0x1000) {
						setTimeout(function () {
							if (FV) {
								FV.SelectItem(FV.FocusedItem, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_EDIT);
							}
							else {
								wsh.SendKeys("{F2}");
							}
						}, 100);
					}
				}
				if (nBase == 7) {
					Navigate(dir[nVerb - 0x1001], SBSP_NEWBROWSER);
				}
			}
			api.DestroyMenu(hMenu);
			if (nVerb == 0) {
				return S_OK;
			}
			if (FV && nBase >= 4) {
				if ((nVerb & 0xfff) == (CommandID_RENAME - 1)) {
					setTimeout(function () {
						FV.SelectItem(FV.FocusedItem, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_EDIT);
					}, 100);
					return S_OK;
				}
				if (api.SendMessage(FV.hwndView, WM_COMMAND, nVerb, 0) == S_OK) {
					return S_OK;
				}
			}
			if (items) {
				item = items[nVerb - 1];
			}
		}
		if (item) {
			Exec(Ctrl, item.text, item.getAttribute("Type"), Ctrl.hwnd, pt);
			return S_OK;
		}
		if (Mode != 2) {
			return S_OK;
		}
	}
	return S_FALSE;
}

MenuDbInit = function (hMenu, oMenu, oMenu2)
{
	for (var i = api.GetMenuItemCount(hMenu); i--;) {
		var mii = api.Memory("MENUITEMINFO");
		mii.cbSize = mii.Size;
		mii.fMask  = MIIM_ID | MIIM_BITMAP | MIIM_SUBMENU | MIIM_DATA | MIIM_FTYPE | MIIM_STATE;
		var s = api.GetMenuString(hMenu, i, MF_BYPOSITION);
		api.GetMenuItemInfo(hMenu, i, true, mii);
		if (s) {
			oMenu[s] = mii;
			if (s.match(/[^\t]+(\t.*)/)) {
				oMenu[RegExp.$1] = mii;
			}
			api.RemoveMenu(hMenu, i, MF_BYPOSITION);
			if (oMenu2 && mii.hSubMenu && !oMenu2[s]) {
				MenuDbInit(mii.hSubMenu, oMenu, null)
			}
		}
		else {
			api.DeleteMenu(hMenu, i, MF_BYPOSITION);
		}
	}
}

MenuDbReplace = function (hMenu, oMenu, hMenu2)
{
	var mii;
	var nCount = api.GetMenuItemCount(hMenu2);
	for (var i = 0; i < nCount; i++) {
		var s = api.GetMenuString(hMenu2, 0, MF_BYPOSITION);
		if (s) {
			mii = oMenu[s] || oMenu[s.replace(/\t.*/, "")];
		}
		if (mii) {
			api.DeleteMenu(hMenu2, 0, MF_BYPOSITION);
		}
		else {
			delete oMenu[s];
			mii = api.Memory("MENUITEMINFO");
			mii.cbSize = mii.Size;
			mii.fMask = MIIM_ID | MIIM_BITMAP | MIIM_SUBMENU | MIIM_DATA | MIIM_FTYPE | MIIM_STATE;
			api.GetMenuItemInfo(hMenu2, 0, true, mii);
			if (mii.hSubMenu) {
				api.DeleteMenu(hMenu2, 0, MF_BYPOSITION);
				continue;
			}
			else {
				api.RemoveMenu(hMenu2, 0, MF_BYPOSITION);
			}
		}
		mii.fMask = MIIM_ID | MIIM_BITMAP | MIIM_SUBMENU | MIIM_DATA | MIIM_FTYPE | MIIM_STATE;
		if (s) {
			mii.dwTypeData = s;
			mii.fMask |= MIIM_STRING;
		}
		api.InsertMenuItem(hMenu, MAXINT, false, mii);
	}
	for (var s in oMenu) {
		if (!s.match(/^\t/)) {
			api.InsertMenuItem(hMenu2, MAXINT, false, oMenu[s]);
		}
	}
	api.DestroyMenu(hMenu2);
}

GetAccelerator = function (s)
{
	if (s.match(/&(.)/)) {
		return RegExp.$1;
	}
	return "";
}

MakeMenus = function (hMenu, menus, arMenu, items)
{
	var hMenus = [hMenu];
	var nPos = menus[0].getAttribute("Pos");
	var nLen = api.GetMenuItemCount(hMenu);
	if (nPos > nLen || nPos < 0) {
		nPos = nLen;
	}
	nLen = arMenu.length;
	var uFlags = 0;
	for (var i = 0; i < nLen; i++) {
		var item = items[arMenu[i]];
		var s = item.getAttribute("Name").replace(/\\t/i, "\t");
		var strFlag = api.strcmpi(item.getAttribute("Type"), "Menus") ? "" : item.text;

		if (s == "/" || api.strcmpi(strFlag, "Break") == 0) {
			uFlags = MF_MENUBREAK;
		}
		else if (s == "//" || api.strcmpi(strFlag, "BarBreak") == 0) {
			uFlags = MF_MENUBARBREAK;
		}
		else if (api.strcmpi(strFlag, "Close") == 0) {
			hMenus.pop();
			if (!hMenus.length) {
				break;
			}
		}
		else {
			var ar = s.split(/\t/);
			ar[0] = GetText(ar[0]);
			if (ar.length > 1) {
				ar[1] = GetKeyName(ar[1]);
			}
			if (api.strcmpi(strFlag, "Open") == 0) {
				var mii = api.Memory("MENUITEMINFO");
				mii.cbSize = mii.Size;
				mii.fMask = MIIM_STRING | MIIM_SUBMENU | MIIM_FTYPE;
				mii.fType = uFlags;
				mii.dwTypeData = ar.join("\t");
				mii.hSubMenu = api.CreatePopupMenu();
				api.InsertMenuItem(hMenus[hMenus.length - 1], nPos++, true, mii);
				hMenus.push(mii.hSubMenu);
			}
			else {
				if (s == "-" || api.strcmpi(strFlag, "Separator") == 0) {
					api.InsertMenu(hMenus[hMenus.length - 1], nPos++, uFlags | MF_BYPOSITION | MF_SEPARATOR, 0, null);
				}
				else {
					api.InsertMenu(hMenus[hMenus.length - 1], nPos++, uFlags | MF_BYPOSITION | MF_STRING, arMenu[i] + 1, ar.join("\t"));
				}
			}
			uFlags = 0;
		}
	}
}

SaveXmlEx = function (filename, xml, bAppData)
{
	xml.save(fso.BuildPath(te.Data.DataFolder, "config\\" + filename));
}

BlurId = function (Id)
{
	document.getElementById(Id).blur();
}

RunCommandLine = function (s)
{
	var arg = api.CommandLineToArgv(s.replace(/\/[^,\s]*/g, ""));
	for (var i = 1; i < arg.Count; i++) {
		Navigate(arg.Item(i), SBSP_NEWBROWSER);
	}
}

GetAddonInfo = function (Id)
{
	var info = new Array();

	var path = fso.GetParentFolderName(api.GetModuleFileName(null));
	var xml = te.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	xml.load(fso.BuildPath(path, "addons\\" + Id + "\\config.xml"));

	GetAddonInfo2(xml, info, "General");
	GetAddonInfo2(xml, info, "en");
	GetAddonInfo2(xml, info, GetLangId());
	if (!info["Name"]) {
		info["Name"] = Id;
	}
	return info;
}

GetAddonInfo2 = function (xml, info, Tag)
{
	var items = xml.getElementsByTagName(Tag);
	if (items.length) {
		var item = items[0].childNodes;
		for (var i = 0; i < item.length; i++) {
			info[item[i].tagName] = item[i].text;
		}
	}
}

OpenXml = function (strFile, bAppData, bEmpty)
{
	var xml = te.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	if (xml.load(fso.BuildPath(te.Data.DataFolder, "config\\" + strFile))) {
		return xml;
	}
	if (!bAppData) {
		var path = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "config\\" + strFile);
		if (xml.load(path)) {
			var Dest = sha.NameSpace(fso.BuildPath(te.Data.DataFolder, "config"));
			Dest.MoveHere(path, FOF_SILENT | FOF_NOCONFIRMATION);
			return xml;
		}
	}
	if (xml.load(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "init\\" + strFile))) {
		return xml;
	}
	if (bEmpty) {
		return xml;
	}
	return null;
}

CreateXml = function ()
{
	var xml = te.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	xml.appendChild(xml.createProcessingInstruction("xml", 'version="1.0" encoding="UTF-8"'));
	return xml;
}

Extract = function (Src, Dest)
{
	for (var i in eventTE.Extract) {
		var hr = eventTE.Extract[i](Src, Dest);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	try {
		var oSrc = sha.NameSpace(Src)
		if (oSrc) {
			var oDest = sha.NameSpace(Dest);
			if (oDest) {
				oDest.CopyHere(oSrc.Items(), FOF_NOCONFIRMATION);
				return S_OK;
			}
		}
	}
	catch (e) {
		wsh.Popup(GetText("Extract Error"), 0, "Tablacus Explorer", MB_ICONSTOP);
	}
	return E_FAIL;
}

OptionRef = function (Id, s, pt)
{
	for (var i in eventTE.OptionRef) {
		var r = eventTE.OptionRef[i](Id, s, pt);
		if (r !== undefined) {
			return r; 
		}
	}
}

OptionDecode = function (Id, p)
{
	for (var i in eventTE.OptionDecode) {
		var hr = eventTE.OptionDecode[i](Id, p);
		if (isFinite(hr)) {
			return hr; 
		}
	}
}

OptionEncode = function (Id, p)
{
	for (var i in eventTE.OptionEncode) {
		var hr = eventTE.OptionEncode[i](Id, p);
		if (isFinite(hr)) {
			return hr; 
		}
	}
}

function CheckUpdate()
{
	var url = "http://www.eonet.ne.jp/~gakana/tablacus/";
	var xhr = createHttpRequest();
	xhr.open("GET", url + "explorer_en.html?" + Math.floor(new Date().getTime() / 60000), false);
	xhr.setRequestHeader('Pragma', 'no-cache');
	xhr.setRequestHeader('Cache-Control', 'no-cache');
	xhr.setRequestHeader('If-Modified-Since', 'Thu, 01 Jun 1970 00:00:00 GMT');
	xhr.send(null);
	if (!xhr.responseText.match(/<td id="te">(.*?)<\/td>/i)) {
		return;
	}
	var s = RegExp.$1;
	if (!s.match(/<a href="dl\/([^"]*)/i)) {
		return;
	}
	var file = RegExp.$1;
	s = s.replace(/Download/i, "").replace(/<[^>]*>/ig, "");
	var ver = 0;
	if (file.match(/(\d+)/)) {
		ver = 20000000 + api.LowPart(RegExp.$1)
	}
	if (ver <= te.Version) {
		wsh.Popup(te.About + "\n" + GetText("the latest version"), 0, TITLE, MB_ICONINFORMATION);
		return;
	}
	if (!confirmYN(GetText("Update available") + "\n" + s + "\n" + GetText("Do you want to install it now?"), 0, TITLE, MB_ICONQUESTION | MB_YESNO)) {
		return;
	}
	var temp = fso.BuildPath(wsh.ExpandEnvironmentStrings("%TEMP%"), "tablacus");
	if (!IsExists(temp)) {
		CreateFolder(temp);
	}
	api.SetCurrentDirectory(temp);
	var InstalledFolder = fso.GetParentFolderName(api.GetModuleFileName(null));
	var zipfile = fso.BuildPath(temp, file);
	temp += "\\explorer";
	DeleteItem(temp);
	CreateFolder(temp);

	xhr.open("GET", url + "dl/" + file, false);
	xhr.send(null);

	var ado = te.CreateObject("Adodb.Stream");
	ado.Type = adTypeBinary;
	ado.Open();
	ado.Write(xhr.responseBody);
	ado.Savetofile(zipfile, adSaveCreateOverWrite);

	if (Extract(zipfile, temp) != S_OK) {
		return;
	}
	var te64exe = temp + "\\te64.exe";
	var nDog = 300;
	while (!fso.FileExists(te64exe)) {
		if (wsh.Popup(GetText("Please wait."), 1, TITLE, MB_ICONINFORMATION | MB_OKCANCEL) == IDCANCEL || nDog-- == 0) {
			return;
		}
	}

	var update = api.sprintf(2000, "\
		T='Tablacus Explorer';\
		F='%s';\
		Q='\\x22';\
		A=new ActiveXObject('Shell.Application');\
		W=new ActiveXObject('WScript.Shell');\
		W.Popup('%s',9,T,%d);\
		A.NameSpace(F).MoveHere(A.NameSpace('%s').Items(),%d);\
		W.Popup('%s',0,T,%d);\
		W.Run(Q+F+'\\\\%s'+Q);\
		close()", EscapeUpdateFile(InstalledFolder), GetText("Please wait."), MB_ICONINFORMATION, EscapeUpdateFile(temp), FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR, GetText("Completed."), MB_ICONINFORMATION, 
EscapeUpdateFile(fso.GetFileName(api.GetModuleFileName(null)))).replace(/[\t\n]/g, "");

	wsh.CurrentDirectory = temp;
	var exe = "mshta.exe";
	var s1 = ' "javascript:';
	if (update.length >= 500) {
		exe = "wscript.exe";
		s1 = fso.GetParentFolderName(temp) + "\\update.js";
		DeleteItem(s1);
		var a = fso.CreateTextFile(s1, true);
		a.WriteLine(update.replace(/close\(\)$/, ""));
		a.Close();
		update = s1;
		s1 = ' "';
	}
	var mshta = wsh.ExpandEnvironmentStrings("%windir%\\Sysnative\\" + exe);
	if (!fso.FileExists(mshta)) {
		mshta = fso.BuildPath(system32, exe);
 	}
	wsh.Run(api.PathQuoteSpaces(mshta) + s1 + update + '"', SW_SHOWNORMAL, false);
	api.PostMessage(te.hwnd, WM_CLOSE, 0, 0);
}

function EscapeUpdateFile(s)
{
	return s.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

confirmYN = function (s)
{
	return wsh.Popup(s, 0, TITLE, MB_ICONQUESTION | MB_YESNO) == IDYES;
}

createHttpRequest = function ()
{
	try {
		return te.CreateObject("Msxml2.XMLHTTP");
	} catch(e) {
		return te.CreateObject("Microsoft.XMLHTTP");
	}
}

InputDialog = function (text, defaultText)
{
	return prompt(text, defaultText);
}

AddonOptions = function (Id, fn)
{
	LoadLang2(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\" + Id + "\\lang\\" + GetLangId() + ".xml"));
	var items = te.Data.Addons.getElementsByTagName(Id);
	if (!items.length) {
		var root = te.Data.Addons.documentElement;
		if (root) {
			root.appendChild(te.Data.Addons.createElement(Id));
		}
	}
	var info = GetAddonInfo(Id);
	var sURL = "../addons/" + Id + "/options.html";
	var Data = {id: Id};
	var sFeatures = info.Options;
	if (sFeatures.match(/Common:([\d,]+):(\d)/i)) {
		sURL = "location.html";
		Data.show = RegExp.$1;
		Data.index = RegExp.$2;
		sFeatures = 'Default';
	}
	if (api.strcmpi(sFeatures, "Location") == 0) {
		sURL = "location.html";
		Data.show = "6";
		Data.index = "6";
		sFeatures = 'Default';
	}
	if (api.strcmpi(sFeatures, "Default") == 0) {
		sFeatures = 'dialogWidth: 640px; dialogHeight: 480px; resizable: yes; status: 0';
	}
	try {
		if (MainWindow.g_dlgs[Id] && MainWindow.g_dlgs[Id].window) {
			MainWindow.g_dlgs[Id].focus();
			return;
		}
	}
	catch (e) {
		MainWindow.g_dlgs[Id] = null;
	}
	var dlg = showModelessDialog(sURL, {MainWindow: MainWindow, Data: Data}, sFeatures);
	MainWindow.g_dlgs[Id] = dlg;
	if (fn || g_Chg) {
		try {
			while (!dlg.window.document.body) {
				api.Sleep(100);
			}
			if (fn) {
				dlg.window.TEOk = fn;
			}
			else if (dlg.window.returnValue === undefined) {
				dlg.window.TEOk = function ()
				{
					g_Chg.Addons = true;
				}
			}
			else {
				g_Chg.Addons = true;
			}
		} catch (e) {}
	}
}

IsSavePath = function (path, mode)
{
	return true;
}

function CalcVersion(s)
{
	var r = 0;
	if (s.match(/(\d+)\.(\d+)\.(\d+)/)) {
		r =  api.LowPart(RegExp.$1) * 10000 + api.LowPart(RegExp.$2) * 100 + api.LowPart(RegExp.$3);
	}
	if (r < 2000 * 10000) {
		r += 2000 * 10000;
	}
	return r;
}

GethwndFromPid = function (ProcessId, bDT)
{
	var hProcess = api.OpenProcess(PROCESS_QUERY_INFORMATION, false, ProcessId);
	if (hProcess) {
		api.WaitForInputIdle(hProcess, 10000);
		api.CloseHandle(hProcess);
	}
	var nIndex = bDT ? GWL_EXSTYLE : GWLP_HWNDPARENT;
	var nFilter = bDT ? 16 : -1;
	var nValue = bDT ? 16 : 0;
	var hwnd = api.GetTopWindow(null);
	do {
		if ((api.GetWindowLongPtr(hwnd, nIndex) & nFilter) == nValue && api.IsWindowVisible(hwnd)) {
			var pProcessId = api.Memory("DWORD");
			api.GetWindowThreadProcessId(hwnd, pProcessId);
			if (ProcessId == pProcessId[0]) {
				return hwnd;
			}
		}
	} while(hwnd = api.GetWindow(hwnd, GW_HWNDNEXT));
	return null;
}

PopupContextMenu = function (Item)
{
	var hMenu = api.CreatePopupMenu();
	var ContextMenu = api.ContextMenu(Item);
	if (ContextMenu) {
		ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, (api.GetKeyState(VK_SHIFT) < 0) ? CMF_NORMAL | CMF_EXTENDEDVERBS : CMF_NORMAL);
		var pt = api.Memory("POINT");
		api.GetCursorPos(pt);
		var nVerb = api.TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, ContextMenu);
		g_popup = null;
		if (nVerb) {
			ContextMenu.InvokeCommand(0, te.hwnd, nVerb - 1, null, null, SW_SHOWNORMAL, 0, 0);
		}
	}
	api.DestroyMenu(hMenu);
}

GetAddonOption = function (strAddon, strTag)
{
	var items = te.Data.Addons.getElementsByTagName(strAddon);
	if (items.length) {
		var item = items[0];
		return item.getAttribute(strTag);
	}
}

GetAddonOptionEx = function (strAddon, strTag)
{
	return api.LowPart(GetAddonOption (strAddon, strTag));
}

GetInnerFV = function (id)
{
	var TC = te.Ctrl(CTRL_TC, id);
	if (TC && TC.SelectedIndex >= 0) {
		return TC.Selected;
	}
	return null;
}

OpenInExplorer = function (FV)
{
	if (FV) {
		var exp = te.CreateObject("new:{C08AFD90-F2A1-11D1-8455-00A0C91F3880}");
		exp.Navigate2(FV.FolderItem);
    	exp.Visible = true;
		exp.Document.CurrentViewMode = FV.CurrentViewMode;
		do {
			api.Sleep(100);
		} while (exp.Busy || exp.ReadyState < 4);
		var doc = exp.Document;
		doc.CurrentViewMode = FV.CurrentViewMode;
		try {
			if (doc.IconSize) {
				doc.IconSize = FV.IconSize;
			}
		} catch (e) {
		}
		try {
			if (doc.SortColumns) {
				doc.SortColumns = FV.SortColumns;
			}
		} catch (e) {
		}
		if (FV.TreeView.Align & 2) {
			exp.ShowBrowserBar("{EFA24E64-B078-11D0-89E4-00C04FC9E26E}", true);
		}
	}
}

InputMouse = function ()
{
	var s = showModalDialog(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "script\\dialog.html"), {MainWindow: MainWindow, Query: "mouse"}, 'dialogWidth: 500px; dialogHeight: 420px; resizable: yes; status: 0;');
	if (s) {
		(document.F.MouseMouse || document.F.Mouse).value = s;
	}
}

InputKey = function()
{
	var s = showModalDialog(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "script\\dialog.html"), {MainWindow: MainWindow, Query: "key"}, 'dialogWidth: 320px; dialogHeight: 120px; resizable: yes; status: 0;');
	if (s) {
		(document.F.KeyKey || document.F.Key).value = s;
		SetKeyShift();
	}
}

ShowLocationEx = function (s)
{
	showModelessDialog("../../script/location.html", {MainWindow: MainWindow, Data: s}, 'dialogWidth: 640px; dialogHeight: 480px; resizable: yes; status=0;');
}

ShowIconEx = function ()
{
	return showModalDialog("../../script/dialog.html", {MainWindow: MainWindow, Query: "icon"}, 'dialogWidth: 640px; dialogHeight: 480px; resizable: yes; status=0;');
}

function MakeKeySelect()
{
	var oa = document.getElementById("_KeyState");
	if (oa) {
		var ar = [];
		for (var i = 0; i < 4; i++) {
			var s = MainWindow.g_KeyState[i][0];
			ar.push('<input type="checkbox" onclick="KeyShift(this)" id="_Key' + s + '"><label for="_Key' + s + '">' + s + '&nbsp;</label>');
		}
		oa.insertAdjacentHTML("AfterBegin", ar.join(""));
	}
	oa = document.getElementById("_KeySelect");
	oa.length = 0;
	oa[++oa.length - 1].value = "";
	oa[oa.length - 1].text = GetText("Select");
	var s = [];
	for (var j = 256; j >= 0; j -= 256) {
		for (var i = 128; i > 0; i--) {
			var v = api.GetKeyNameText((i + j) * 0x10000);
			if (v && v.charCodeAt(0) > 32) {
				s.push(v);
			}
		}
	}
	s.sort(function (a,b) {
		if (a.length != b.length && (a.length == 1 || b.length == 1)) {
			return a.length - b.length;
		}
		return api.StrCmpLogical(a, b);
	});
	var j = "";
	for (i in s) {
		if (j != s[i]) {
			j = s[i];
			oa[++oa.length - 1].value = j;
			oa[oa.length - 1].text = j;
		}
	}
}

function SetKeyShift()
{
	var key = (document.F.elements.KeyKey || document.F.elements.Key).value;
	for (var i = 0; i < MainWindow.g_KeyState.length; i++) {
		var s = MainWindow.g_KeyState[i][0];
		var o = document.getElementById("_Key" + s);
		if (o) {
			o.checked = key.match(s + "+");
		}
		key = key.replace(s + "+", "");
	}
	o = document.getElementById("_KeySelect");
	for (var i = o.length; i--;) {
		if (api.strcmpi(key, o[i].value) == 0) {
			o.selectedIndex = i;
			break;
		}
	}
}

function KeyShift(o)
{
	var oKey = document.F.elements.KeyKey || document.F.elements.Key;
	var key = oKey.value;
	var shift = o.id.replace(/^_Key(.*)/, "$1+");
	key = key.replace(shift, "");
	if (o.checked) {
		key = shift + key;
	}
	oKey.value = key;
}

function KeySelect(o)
{
	var oKey = document.F.elements.KeyKey || document.F.elements.Key;
	oKey.value = oKey.value.replace(/(\+)[^\+]*$|^[^\+]*$/, "$1") + o[o.selectedIndex].value;
}

GetLangId = function ()
{
	return te.Data.Conf_Lang || navigator.userLanguage.replace(/\-.*/,"");
}

GetSourceText = function (s)
{
	var Lang = MainWindow.Lang;
	for (var i in Lang) {
		if (s == Lang[i]) {
			return i;
		}
	}
	return s;
}

GetFolderView = function (Ctrl, pt, bStrict)
{
	if (!Ctrl) {
		return te.Ctrl(CTRL_FV);
	}
	if (Ctrl.Type == CTRL_SB || Ctrl.Type == CTRL_EB) {
		return Ctrl;
	}
	if (Ctrl.Type == CTRL_TV) {
		return Ctrl.FolderView;
	}
	if (Ctrl.Type != CTRL_TC) {
		return te.Ctrl(CTRL_FV);
	}
	if (pt) {
		var i = Ctrl.HitTest(pt, TCHT_ONITEM);
		if (i >= 0) {
			return Ctrl.Item(i);
		}
	}
	if (!bStrict || !pt) {
		return Ctrl.Selected;
	}
}

GetSelectedArray = function (Ctrl, pt, bPlus)
{
	var Selected = null;
	var SelItem = null;
	var FV = null;
	var bSel = true;
	switch(Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
			FV = Ctrl;
			break;
		case CTRL_TC:
			FV = Ctrl.Item(Ctrl.HitTest(pt, TCHT_ONITEM));
			bSel = false;
			break;
		case CTRL_TV:
			SelItem = Ctrl.SelectedItem;
			break;
		case CTRL_WB:
			SelItem = window.Input;
			break;
		default:
			FV = te.Ctrl(CTRL_FV);
			break;
	}
	if (FV) {
		if (bSel) {
	 		Selected = FV.SelectedItems();
		}
		if (Selected && Selected.Count) {
			SelItem = Selected.Item(0);
		}
		else {
			SelItem = FV.FolderItem;
		}
	}
	if (!Selected || Selected.Count == 0) {
		Selected = te.FolderItems();
		if (bPlus) {
			Selected.AddItem(SelItem);
		}
	}
	return [Selected, SelItem, FV];
}

StripAmp = function (s)
{
	return s.replace(/\(&\w\)|&/, "").replace(/\.\.\.$/, "");
}

GetGestureKey = function ()
{
	var s = "";
	if (api.GetKeyState(VK_SHIFT) < 0) {
		s += "S";
	}
	if (api.GetKeyState(VK_CONTROL) < 0) {
		s += "C";
	}
	if (api.GetKeyState(VK_MENU) < 0) {
		s += "A";
	}
	return s;
}

GetGestureButton = function ()
{
	var s = "";
	if (api.GetKeyState(VK_LBUTTON) < 0) {
		s = "1";
	}
	if (api.GetKeyState(VK_RBUTTON) < 0) {
		s += "2";
	}
	if (api.GetKeyState(VK_MBUTTON) < 0) {
		s += "3";
	}
	if (api.GetKeyState(VK_XBUTTON1) < 0) {
		s += "4";
	}
	if (api.GetKeyState(VK_XBUTTON2) < 0) {
		s += "5";
	}
	return s;
}

GetWebColor = function (c)
{
	return api.sprintf(8, "#%06x", ((c & 0xff) << 16) | (c & 0xff00) | ((c & 0xff0000) >> 16));
}

GetWinColor = function (c)
{
	var c2 = api.sscanf(c, "#%06x");
	if (c2) {
		return ((c2 & 0xff) << 16) | (c2 & 0xff00) | ((c2 & 0xff0000) >> 16);
	}
	c2 = api.sscanf(c, "#%03x");
	return c2 ? ((c2 & 0xf) << 20) | ((c2 & 0xf) << 16) | ((c2 & 0xf0) << 8) | ((c2 & 0xf0) << 4) | ((c2 & 0xf00) >> 4) | ((c2 & 0xf00) >> 8) : c;
}

ChooseColor = function (c)
{
	var cc = api.Memory("CHOOSECOLOR");
	cc.lStructSize = cc.Size;
	cc.hwndOwner = te.hwnd;
	cc.Flags = CC_FULLOPEN | CC_RGBINIT;
	cc.rgbResult = c;
	cc.lpCustColors = te.Data.CustColors;
	if (api.ChooseColor(cc)) {
		return cc.rgbResult;
	}
}

ChooseWebColor = function (c)
{
	c = ChooseColor(GetWinColor(c));
	if (c) {
		return GetWebColor(c);
	}
}

SetCursor = function (o, s)
{
	if (o) {
		if (o.style) {
			o.style.cursor = s;
		}
		if (o.getElementsByTagName) {
			var e = o.getElementsByTagName("*");
			for (var i in e) {
				SetCursor(e[i], s);
			}
		}
	}
}

function MouseOver(o)
{
	if (o.className == 'button' || o.className == 'menu') {
		if (objHover && o != objHover) {
			MouseOut();
		}
		var pt = api.Memory("POINT");
		api.GetCursorPos(pt, true);
		if (HitTest(o, pt)) {
			objHover = o;
			o.className = 'hover' + o.className;
		}
	}
}

function MouseOut(s)
{
	if (objHover) {
		if (!s || objHover.id.match(s)) {
			if (objHover.className == 'hoverbutton') {
				objHover.className = 'button';
			}
			else if (objHover.className == 'hovermenu') {
				objHover.className = 'menu';
			}
			objHover = null;
		}
	}
}


InsertTab = function(e)
{
	var ot = (e || event).srcElement;
	if (event.keyCode == VK_TAB) {
		ot.focus();
		if (document.all && document.selection) {
			var selection = document.selection.createRange();
			if (selection) {
				selection.text += "\t";
				return false;
			}
		}
		var i = ot.selectionEnd || s.length;
		var s = ot.value;
		ot.value = s.substr(0, i) + "\t" + s.substr(i, s.length);
		ot.selectionStart = ++i;
		ot.selectionEnd = i;
		return false;
	}
	return true;
}

RegEnumKey = function(hKey, Name)
{
	var locator = te.CreateObject("WbemScripting.SWbemLocator");
	var server = locator.ConnectServer(null, "root\\default");
	var reg = server.Get("StdRegProv");

	var method = reg.Methods_.Item("EnumKey");
	var iParams = method.InParameters.SpawnInstance_();

	iParams.hDefKey = hKey;
	iParams.sSubKeyName = Name; 

	var result = reg.ExecMethod_(method.Name, iParams);
	return result.sNames.toArray();
}
