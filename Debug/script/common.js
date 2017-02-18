//Tablacus Explorer

Ctrl = null;
g_temp = null;
g_sep = "` ~";
Handled = null;
hwnd = null;
pt = api.Memory("POINT");
dataObj = null;
grfKeyState = null;
pdwEffect = 0;
bDrop = null;
Input = null;
eventTE = { Environment: {} };
eventTA = {};
g_ptDrag = api.Memory("POINT");
objHover = null;
Addons = {"_stack": []};
g_ = {
	rcWindow: api.Memory("RECT"),
	Colors: {},
	nFind: 0,
	Panels: {},
	KeyCode: {},
	KeyState: [
		[0x1d0000, 0x2000],
		[0x2a0000, 0x1000],
		[0x380000, 0x4000],
		["Win",    0x8000],
		["Ctrl",   0x2000],
		["Shift",  0x1000],
		["Alt",    0x4000]
	],
	stack_TC: [],
	dlgs: {},
	tidWindowRegistered: null,
	bWindowRegistered: true,
	xmlWindow: null
};

FolderMenu =
{
	Items: [],

	Clear: function ()
	{
		this.Items.length = 0;
	},

	Open: function (FolderItem, x, y, filter)
	{
		this.Clear();
		this.Filter = filter;
		var hMenu = api.CreatePopupMenu();
		this.OpenMenu(hMenu, FolderItem);
		window.g_menu_click = true;
		var Verb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, x, y, te.hwnd, null, null);
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
		if (!FolderItem) {
			return;
		}
		var path = FolderItem.Path || FolderItem;
		if (path === FolderItem || /^[A-Z]:\\$/i.test(path)) {
			FolderItem = api.ILCreateFromPath(path);
		}
		try {
			if (!FolderItem || FolderItem.IsBrowsable) {
				FolderItem = {};
			}
			var bSep = false;
			if (!api.ILIsEmpty(FolderItem)) {
				this.AddMenuItem(hMenu, api.ILRemoveLastID(FolderItem, true), "../");
				bSep = true;
			}
			var Folder = FolderItem.GetFolder;
			if (Folder) {
				var Items = Folder.Items();
				if (Items) {
					var nCount = Items.Count;
					for (var i = 0; i < nCount; i++) {
						var Item = Items.Item(i);
						var bMatch = IsFolderEx(Item);
						if (this.Filter) {
							bMatch = api.PathMatchSpec(bMatch ? Item.Name + ".folder" : Item.Name, this.Filter);
						}
						if (bMatch) {
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
		} catch (e) {}
	},

	AddMenuItem: function (hMenu, FolderItem, Name, bSelect)
	{
		var mii = api.Memory("MENUITEMINFO");
		mii.cbSize = mii.Size;
		mii.fMask = MIIM_ID | MIIM_STRING | MIIM_BITMAP | MIIM_SUBMENU;
		if (bSelect && Name) {
			mii.dwTypeData = Name;
		} else {
			mii.dwTypeData = (Name ? Name + api.GetDisplayNameOf(FolderItem, SHGDN_INFOLDER) : api.GetDisplayNameOf(FolderItem, SHGDN_INFOLDER));
		}
		AddMenuIconFolderItem(mii, FolderItem);
		this.Items.push(FolderItem);
		mii.wID = this.Items.length;
		if (!bSelect && api.GetAttributesOf(FolderItem, SFGAO_HASSUBFOLDER | SFGAO_BROWSABLE) == SFGAO_HASSUBFOLDER) {
			var path = api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
			mii.hSubMenu = api.CreateMenu();
			api.InsertMenu(mii.hSubMenu, 0, MF_BYPOSITION | MF_STRING, 0, api.sprintf(99, '\tJScript\tFolderMenu.OpenSubMenu("%llx",%d,"%llx")', hMenu, mii.wID, mii.hSubMenu));
		}
		api.InsertMenuItem(hMenu, MAXINT, false, mii);
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
		Name = Name.replace("Dragleave", "DragLeave");
		if (Name == "ItemPrePaint" && !te.OnItemPrePaint) {
			te.OnItemPrePaint = function (Ctrl, pid, nmcd, vcd, plRes)
			{
				RunEvent3("ItemPrePaint", Ctrl, pid, nmcd, vcd, plRes);
			}
		}

		if (!eventTE[Name]) {
			eventTE[Name] = [];
		}
		if (!eventTA[Name]) {
			eventTA[Name] = [];
		}
		if (priority) {
			eventTE[Name].unshift(fn);
			eventTA[Name].unshift(window.Error_source);
		} else {
			eventTE[Name].push(fn);
			eventTA[Name].push(window.Error_source);
		}
	}
}

AddEnv = function (Name, fn)
{
	eventTE.Environment[Name.toLowerCase()] = fn;
}

AddEventId = function (Name, Id, fn)
{
	if (!eventTE[Name]) {
		eventTE[Name] = {};
	}
	eventTE[Name][Id.toLowerCase()] = fn;
}

AddonDisabled = function (Id)
{
	for (var i in eventTE.AddonDisabled) {
		eventTE.AddonDisabled[i](Id);
	}
	if (eventTE.AddonDisabledEx) {
		var fn = eventTE.AddonDisabledEx[Id.toLowerCase()];
		if (fn) {
			delete eventTE.AddonDisabledEx[Id.toLowerCase()];
			AddEventEx(window, "beforeunload", function ()
			{
				fn();
			});
		}
	}
	CollectGarbage();
}

function ApplyLang(doc)
{
	var FaceName = MainWindow.DefaultFont.lfFaceName;
	if (doc.body) {
		doc.body.style.fontFamily = FaceName;
		doc.body.style.fontSize = Math.abs(MainWindow.DefaultFont.lfHeight) + "px";
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
			if (o[i].className == "treebutton" && o[i].innerHTML == "") {
				o[i].innerHTML = BUTTONS.opened;
			}
		}
	}
	var h = 0;
	var o = doc.getElementsByTagName("input");
	if (o) {
		for (i = o.length; i--;) {
			if (!h && o[i].type == "text") {
				h = o[i].offsetHeight * screen.deviceYDPI / screen.logicalYDPI;
			}
			var s = Lang[o[i].placeholder];
			if (s) {
				o[i].placeholder = s;
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
			if (!o[i].ondragstart) {
				o[i].draggable = false;
			}
		}
	}
	var o = doc.getElementsByTagName("select");
	if (o) {
		for (i = o.length; i--;) {
			var s = Lang[o[i].title];
			if (s) {
				o[i].title = delamp(s);
			}
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
			o[i].onkeydown = InsertTab;
		}
	}

	var o = doc.getElementsByTagName("li");
	if (o) {
		for (i = o.length; i--;) {
			var s = Lang[o[i].innerHTML.replace(/&amp;/ig, "&")];
			if (!s) {
				s = o[i].innerHTML;
			}
			o[i].innerHTML = amp2ul(s);
		}
	}

	var o = doc.getElementsByTagName("form");
	if (o) {
		for (i = o.length; i--;) {
			o[i].onsubmit = function () { return false };
		}
	}

	doc.title = GetText(doc.title);
	setTimeout(function ()
	{
		var hwnd = api.GetParent(api.GetWindow(doc));
		var s = api.GetWindowText(hwnd);
		if (/ \-+ .*$/.test(s)) {
			api.SetWindowText(hwnd, s.replace(/ \-+ .*$/, ""));
		}
	}, 500);
}

function amp2ul(s)
{
	s = s.replace(/&amp;/ig, "&");
	if (/@.*\..*,\-?\d+/.test(s)) {
		try {
			var lk = wsh.CreateShortCut(".lnk");
			lk.Description = s;
			s = lk.Description;
		} catch (e) {}
	}
	return /;/.test(s) ? s : s.replace(/&(.)/ig, "<u>$1</u>");
}

function delamp(s)
{
	s = s.replace(/&amp;/ig, "&");
	return /;/.test(s) ? s : s.replace(/&/ig, "");
}

function ImgBase64(o, index)
{
	var src = ExtractMacro(te, o.src);
	var s = MakeImgSrc(src, index, false, o.height, o.getAttribute("bitmap"), o.getAttribute("icon"));
	if (s) {
		o.removeAttribute("bitmap");
		o.removeAttribute("icon");
	} else if (o.src.toLowerCase() != src.toLowerCase()) {
		return src.replace(location.href.replace(/[^\/]*$/, ""), "file:///");
	}
	return s;
}

function MakeImgSrc(src, index, bSrc, h, strBitmap, strIcon)
{
	var fn;
	if (!document.documentMode) {
		var res = /^bitmap:(.*)/i.exec(src);
		var value = res ? res[1] : strBitmap;
		if (value) {
			fn = fso.BuildPath(te.Data.DataFolder, "cache\\bitmap\\" + value.replace(/[:\\\/]/g, "$") + ".png");
		} else {
			res = /^icon:(.*)/i.exec(src);
			value = res ? res[1] : strIcon;
			if (value) {
				fn = fso.BuildPath(te.Data.DataFolder, "cache\\icon\\" + value.replace(/[:\\\/]/g, "$") + ".png");
			} else if (src && !REGEXP_IMAGE.test(src)) {
				src = src.replace(/^file:\/\/\//i, "").replace(/\//g, "\\");
				fn = fso.BuildPath(te.Data.DataFolder, "cache\\file\\" + src.replace(/[:\\\/]/g, "$") + ".png");
			}
		}
		if (fn && fso.FileExists(fn)) {
			return fn;
		}
	}
	src = ExtractMacro(te, src);
	var image = MakeImgData(src, index, h, strBitmap, strIcon);
	if (image) {
		if (document.documentMode) {
			return image.DataURI("image/png");
		}
		if (fn) {
			try {
				image.Save(fn);
			} catch (e) {}
			return fn;
		}
	}
	return bSrc ? src : "";
}

function MakeImgData(src, index, h, strBitmap, strIcon)
{
	var hIcon = MakeImgIcon(src, index, h, strBitmap, strIcon);
	if (hIcon) {
		var image = te.GdiplusBitmap();
		image.FromHICON(hIcon);
		api.DestroyIcon(hIcon);
		return image;
	}
	return null;
}

function MakeImgIcon(src, index, h, strBitmap, strIcon)
{
	var hIcon = null;
	var res = /^bitmap:(.*)/i.exec(src);
	var value = res ? res[1] : strBitmap;
	if (value) {
		var icon = value.split(",");
		var hModule = LoadImgDll(icon, index);
		if (hModule) {
			var himl = api.ImageList_LoadImage(hModule, isFinite(icon[index * 4 + 1]) ? Number(icon[index * 4 + 1]) : icon[index * 4 + 1], icon[index * 4 + 2], 0, CLR_DEFAULT, IMAGE_BITMAP, LR_CREATEDIBSECTION);
			if (himl) {
				hIcon = api.ImageList_GetIcon(himl, icon[index * 4 + 3], ILD_NORMAL);
				api.ImageList_Destroy(himl);
			}
			api.FreeLibrary(hModule);
			return hIcon;
		}
	}
	res = /^icon:(.*)/i.exec(src);
	value = res ? res[1] : strIcon;
	if (value) {
		var icon = value.split(",");
		var phIcon = api.Memory("HANDLE");
		if (icon[index * 4 + 2] > 16) {
			api.SHDefExtractIcon(icon[index * 4], icon[index * 4 + 1], 0, phIcon, null, icon[index * 4 + 2]);
		} else {
			api.SHDefExtractIcon(icon[index * 4], icon[index * 4 + 1], 0, null, phIcon, icon[index * 4 + 2] << 16);
		}
		if (phIcon[0]) {
			return phIcon[0];
		}
	}
	if (src && !REGEXP_IMAGE.test(src)) {
		var sfi = api.Memory("SHFILEINFO");
		var pidl = api.ILCreateFromPath(api.PathUnquoteSpaces(src));
		if (pidl) {
			api.SHGetFileInfo(pidl, 0, sfi, sfi.Size, (h && h <= 16) ? SHGFI_PIDL | SHGFI_ICON | SHGFI_SMALLICON : SHGFI_PIDL | SHGFI_ICON);
			return sfi.hIcon;
		}
	}
	return null;
}

LoadImgDll = function (icon, index)
{
	var hModule = api.LoadLibraryEx(fso.BuildPath(system32, icon[index * 4]), 0, LOAD_LIBRARY_AS_DATAFILE);
	if (!hModule && icon[index * 4].toLowerCase() == "ieframe.dll") {
		if (icon[index * 4 + 1] >= 500) {
			hModule = api.LoadLibraryEx(fso.BuildPath(system32, "browseui.dll"), 0, LOAD_LIBRARY_AS_DATAFILE);
		} else if (WINVER > 0x500) {
			hModule = api.LoadLibraryEx(fso.BuildPath(system32, "shell32.dll"), 0, LOAD_LIBRARY_AS_DATAFILE);
		} else {
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
	if (!fso.FileExists(filename)) {
		if (/_\w+\.xml$/.test(filename)) {
			filename = filename.replace(/_\w+\.xml$/, ".xml");
			if (!fso.FileExists(filename)) {
				return;
			}
		} else {
			return;
		}
	}
	xml.load(filename);
	var items = xml.getElementsByTagName('text');
	for (var i = 0; i < items.length; i++) {
		var item = items[i];
		SetLang2(item.getAttribute("s").replace("\\t", "\t").replace("\\n", "\n"), item.text.replace("\\t", "\t").replace("\\n", "\n"));
	}
}

SetLang2 = function(s, v)
{
	if (!MainWindow.Lang[s] && !MainWindow.LangSrc[v]) {
		MainWindow.Lang[s] = v;
		MainWindow.LangSrc[v] = s;
		if (/&/.test(s)) {
			SetLang2(s.replace(/\(&\w\)|&/, ""), v.replace(/\(&\w\)|&/, ""));
		}
		if (/\.\.\.$/.test(s)) {
			SetLang2(StripAmp(s), StripAmp(v));
		}
	}
}

LoadXml = function (filename, nGroup)
{
	te.LockUpdate();
	if (!nGroup) {
		var cTC = te.Ctrls(CTRL_TC);
		for (i in cTC) {
			cTC[i].Close();
		}
	}
	var xml = filename;
	if (/string/i.test(typeof filename)) {
		filename = api.PathUnquoteSpaces(filename);
		if (fso.FileExists(filename)) {
			xml = te.CreateObject("Msxml2.DOMDocument");
			xml.async = false;
			xml.load(filename);
		}
	}
	var items = xml.getElementsByTagName('Ctrl');
	for (var i = 0; i < items.length; i++) {
		var item = items[i];
		switch(item.getAttribute("Type") - 0) {
			case CTRL_TC:
				var TC = te.CreateCtrl(CTRL_TC, item.getAttribute("Left"), item.getAttribute("Top"), item.getAttribute("Width"), item.getAttribute("Height"), item.getAttribute("Style"), item.getAttribute("Align"), item.getAttribute("TabWidth"), item.getAttribute("TabHeight"));
				TC.Data.Group = nGroup || Number(item.getAttribute("Group")) || 0;
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
					var FV = TC.Selected.Navigate2(Path, SBSP_NEWBROWSER, tab.getAttribute("Type"), tab.getAttribute("ViewMode"), tab.getAttribute("FolderFlags"), tab.getAttribute("Options"), tab.getAttribute("ViewFlags"), tab.getAttribute("IconSize"), tab.getAttribute("Align"), tab.getAttribute("Width"), tab.getAttribute("Flags"), tab.getAttribute("EnumFlags"), tab.getAttribute("RootStyle"), tab.getAttribute("Root"), tab.getAttribute("SizeFormat"));
					FV.FilterView = tab.getAttribute("FilterView");
					FV.Data.Lock = api.LowPart(tab.getAttribute("Lock")) != 0;
					Lock(TC, i2, false);
				}
				TC.SelectedIndex = item.getAttribute("SelectedIndex");
				TC.Visible = api.LowPart(item.getAttribute("Visible"));
				break;
		}
	}
	if (!nGroup) {
		MainWindow.RunEvent1("LoadWindow", xml);
	}
	te.UnlockUpdate();
}

SaveXmlTC = function (Ctrl, xml, nGroup)
{
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
	item.setAttribute("Group", api.LowPart(nGroup || Ctrl.Data.Group));

	var bEmpty = true;
	var nCount2 = Ctrl.Count;
	for (var i2 in Ctrl) {
		var FV = Ctrl[i2];
		var path = GetSavePath(FV.FolderItem);
		var bSave = IsSavePath(path);
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
			item2.setAttribute("SizeFormat", FV.SizeFormat);
			item2.setAttribute("ViewFlags", FV.ViewFlags);
			item2.setAttribute("FilterView", FV.FilterView);
			item2.setAttribute("Lock", api.LowPart(FV.Data.Lock));
			var TV = FV.TreeView;
			item2.setAttribute("Align", TV.Align);
			item2.setAttribute("Width", TV.Width);
			item2.setAttribute("Flags", TV.Style);
			item2.setAttribute("EnumFlags", TV.EnumFlags);
			item2.setAttribute("RootStyle", TV.RootStyle);
			item2.setAttribute("Root", String(TV.Root));
			var TL = FV.History;
			if (TL) {
				if (TL.Count > 1) {
					var bLogSaved = false;
					var nLogIndex = TL.Index;
					for (var i3 in TL) {
						path = GetSavePath(TL[i3]);
						if (IsSavePath(path)) {
							var item3 = xml.createElement("Log");
							item3.setAttribute("Path", path);
							item2.appendChild(item3);
							bLogSaved = true;
						} else if (i3 < nLogIndex) {
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
	xml.documentElement.appendChild(item);
}

SaveXml = function (filename, all)
{
	var xml = CreateXml(true);
	var root = xml.documentElement;
	if (all) {
		var item = xml.createElement("Window");
		if (!api.IsZoomed(te.hwnd) && !api.IsIconic(te.hwnd)) {
			api.GetWindowRect(te.hwnd, g_.rcWindow);
		}
		item.setAttribute("Left", g_.rcWindow.left);
		item.setAttribute("Top", g_.rcWindow.top);
		item.setAttribute("Width", g_.rcWindow.right - g_.rcWindow.left);
		item.setAttribute("Height", g_.rcWindow.bottom - g_.rcWindow.top);
		item.setAttribute("CmdShow", api.IsZoomed(te.hwnd) ? SW_SHOWMAXIMIZED : SW_SHOWNORMAL);
		root.appendChild(item);
	}
	var cTC = te.Ctrls(CTRL_TC);
	for (var i in cTC) {
		SaveXmlTC(cTC[i], xml);
	}
	if (all) {
		for (var i in te.Data) {
			var res = /^(Tab|Tree|View|Conf)_(.*)/.exec(i);
			if (res) {
				if (isFinite(te.Data[i]) || te.Data[i] != "") {
					var item = xml.createElement(res[1]);
					item.setAttribute("Id", res[2]);
					item.text = te.Data[i];
					root.appendChild(item);
				}
			}
		}
	}
	MainWindow.RunEvent1("SaveWindow", xml, all);
	try {
		xml.save(api.PathUnquoteSpaces(filename));
	} catch (e) {
		if (e.number != E_ACCESSDENIED) {
			ShowError(e, [GetText("Save"), filename].join(": "));
		}
	}
}

GetKeyKey = function (strKey)
{
	var nShift = api.sscanf(strKey, "$%x");
	if (nShift) {
		return nShift;
	}
	strKey = strKey.toUpperCase();
	for (var j in MainWindow.g_.KeyState) {
		var s = MainWindow.g_.KeyState[j][0].toUpperCase() + "+";
		if (strKey.match(s)) {
			strKey = strKey.replace(s, "");
			nShift |= MainWindow.g_.KeyState[j][1];
		}
	}
	return nShift | MainWindow.g_.KeyCode[strKey];
}

GetKeyName = function (strKey)
{
	var nKey = api.sscanf(strKey, "$%x");
	if (nKey) {
		var s = api.GetKeyNameText((nKey & 0x17f) << 16);
		if (s) {
			var arKey = [];
			for (var j in MainWindow.g_.KeyState) {
				if (nKey & MainWindow.g_.KeyState[j][1]) {
					nKey -= MainWindow.g_.KeyState[j][1];
					arKey.push(MainWindow.g_.KeyState[j][0]);
				}
			}
			if (GetKeyKey(s) == nKey) {
				arKey.push(s);
				return arKey.join("+");
			}
		}
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
	} else {
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
	Navigate(HOME_PATH ? HOME_PATH : FV, SBSP_NEWBROWSER);
}

Navigate = function (Path, wFlags)
{
	NavigateFV(te.Ctrl(CTRL_FV), Path, wFlags);
}

NavigateFV = function (FV, Path, wFlags)
{
	if (!FV) {
		var TC = te.CreateCtrl(CTRL_TC, 0, 0, "100%", "100%", te.Data.Tab_Style, te.Data.Tab_Align, te.Data.Tab_TabWidth, te.Data.Tab_TabHeight);
		FV = TC.Selected;
	}
	if (/string/i.test(typeof Path)) {
		if (/%([^%]+)%/.test(Path)) {
			Path = ExtractMacro(FV, Path);
		}
		if (/\?|\*/.test(Path) && !/\\\\\?\\|:/.test(Path)) {
			FV.FilterView = Path;
			FV.Refresh();
			return;
		}
		Path = ExtractPath(FV, Path);
	}
	if (FV.Data.Lock) {
		wFlags |= SBSP_NEWBROWSER;
	}
	FV.Navigate(Path, wFlags);
	FV.Focus();
}

GetOpenMode = function ()
{
	return window.g_menu_button == 3 || api.GetKeyState(VK_CONTROL) < 0 ? SBSP_NEWBROWSER : OpenMode;
}

IsDrag = function (pt1, pt2)
{
	if (pt1 && pt2) {
		try {
			return (Math.abs(pt1.x - pt2.x) > api.GetSystemMetrics(SM_CXDRAG) | Math.abs(pt1.y - pt2.y) > api.GetSystemMetrics(SM_CYDRAG));
		} catch (e) {}
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
		var dlg = g_.dlgs.Options;
		if (dlg) {
			dlg.Window.SetTab(s);
			dlg.Focus();
			return;
		}
	} catch (e) {}
	g_.dlgs.Options = ShowDialog("options.html",
	{
		Data: s, event:
		{
			onbeforeunload: function () 
			{
				delete MainWindow.g_.dlgs.Options;
			}
		}
	})
}

ShowDialog = function (fn, opt)
{
	opt.opener = window;
	if (!/:/.test(fn)) {
		fn = location.href.replace(/[^\/]*$/, fn);
	}
	var r = Math.abs(MainWindow.DefaultFont.lfHeight) / 12;
	return te.CreateCtrl(CTRL_SW, fn, opt, document, (opt.width || 640) * r, (opt.height || 480) * r);
}

LoadLayout = function ()
{
	var commdlg = te.CommonDialog();
	commdlg.InitDir = fso.BuildPath(te.Data.DataFolder, "layout");
	commdlg.Filter = "XML Files|*.xml|" + (api.LoadString(hShell32, 34193) || "All Files") + "|*.*";
	commdlg.Flags = OFN_FILEMUSTEXIST;
	if (commdlg.ShowOpen()) {
		LoadXml(commdlg.FileName);
	}
	return S_OK;
}

SaveLayout = function ()
{
	var commdlg = te.CommonDialog();
	commdlg.InitDir = fso.BuildPath(te.Data.DataFolder, "layout");
	commdlg.Filter = "XML Files|*.xml|" + (api.LoadString(hShell32, 34193) || "All Files") + "|*.*";
	commdlg.DefExt = "xml";
	commdlg.Flags = OFN_OVERWRITEPROMPT;
	if (commdlg.ShowSave()) {
		SaveXml(commdlg.FileName);
	}
	return S_OK;
}

GetPos = function (o, bScreen, bAbs, bPanel, bBottom)
{
	if (/number/i.test(typeof bScreen)) {
		bAbs = bScreen & 2;
		bPanel = bScreen & 4;
		bBottom = bScreen & 8;
		bScreen &= 1;
	}
	var x = (bScreen ? screenLeft : 0);
	var y = (bScreen ? screenTop : 0);
	if (bBottom) {
		y += o.offsetHeight;
	}
	while (o) {
		if (bAbs || !bPanel || String(o.style.position).toLowerCase() != "absolute") {
			x += o.offsetLeft - (bAbs ? 0 : o.scrollLeft);
			y += o.offsetTop - (bAbs ? 0 : o.scrollTop);
			o = o.offsetParent;
		} else {
			break;
		}
	}
	var pt = api.Memory("POINT");
	pt.x = x * screen.deviceXDPI / screen.logicalXDPI;
	pt.y = y * screen.deviceYDPI / screen.logicalYDPI;
	return pt;
}

HitTest = function (o, pt)
{
	if (o) {
		var p = GetPos(o, 1);
		if (pt.x >= p.x && pt.x < p.x + o.offsetWidth && pt.y >= p.y && pt.y < p.y + o.offsetHeight * screen.deviceYDPI / screen.logicalYDPI) {
			o = o.offsetParent;
			p = GetPos(o, 3);
			return pt.x >= p.x && pt.x < p.x + o.offsetWidth * screen.deviceXDPI / screen.logicalXDPI && pt.y >= p.y && pt.y < p.y + o.offsetHeight * screen.deviceYDPI / screen.logicalYDPI;
		}
	}
	return false;
}

PtInRect = function (rc, pt)
{
	return pt.x >= rc.Left && pt.x < rc.Right && pt.y >= rc.Top && pt.y < rc.Bottom;
}

DeleteItem = function (path, fFlags)
{
	if (IsExists(path)) {
		api.SHFileOperation(FO_DELETE, path, null, fFlags || FOF_SILENT | FOF_NOCONFIRMATION | FOF_NOERRORUI, false);
	}
}

IsExists = function (path)
{
	var wfd = api.Memory("WIN32_FIND_DATA");
	var hFind = api.FindFirstFile(path, wfd);
	api.FindClose(hFind);
	return hFind != INVALID_HANDLE_VALUE;
}

CreateNew = function (path, fn)
{
	if (fn && !IsExists(path)) {
		try {
			fn(path);
		} catch (e) {
			if (/^[A-Z]:\\|^\\\\[A-Z]/i.test(path)) {
				var path1, path2, path3, path4;
				path1 = path;
				path2 = "";
				do {
					path2 = fso.BuildPath(fso.GetFileName(path1), path2);
					path1 = fso.GetParentFolderName(path1);
				} while (path1 && !fso.FolderExists(path1));
				var ar = path2.split("\\");
				if (ar[0]) {
					path = fso.BuildPath(path1, ar[0]);
					path3 = fso.BuildPath(fso.GetSpecialFolder(2).Path, ar[0]);
					DeleteItem(path3);
					path4 = path3;
					for (var i = 1; i < ar.length; i++) {
						fso.CreateFolder(path4);
						path4 = fso.BuildPath(path4, ar[i]);
					}
					fn(path4);
					api.SHFileOperation(FO_MOVE, path3, fso.GetParentFolderName(path), FOF_SILENT | FOF_NOCONFIRMATION, false);
				}
			}
		}
	}
	MainWindow.setTimeout(
		'path="' + path.replace(/\\/g, "\\\\") + '";\
		var FV = te.Ctrl(CTRL_FV);\
		if (FV) {\
			if (!api.StrCmpI(FV.FolderItem.Path, fso.GetParentFolderName(path))) {\
				FV.SelectItem(path, SVSI_SELECT | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_FOCUSED | SVSI_NOTAKEFOCUS);\
			}\
		}', 99
	);
}

SetFileTime = function (path, ctime, atime, mtime)
{
	var b = MainWindow.RunEvent3("SetFileTime", path, ctime, atime, mtime);
	if (isFinite(b)) {
		return b; 
	}
	return api.SetFileTime(path, ctime, atime, mtime);
}

SetFileAttributes = function (path, attr)
{
	var b = MainWindow.RunEvent3("SetFileAttributes", path, attr);
	if (isFinite(b)) {
		return b; 
	}
	try {
		fso.GetFile(path).Attributes = attr;
	} catch (e) {
		try {
			fso.GetFolder(path).Attributes = attr;
		} catch (e) {
			return false;
		}
	}
	return true;
}

CreateFolder = function (path)
{
	var r = MainWindow.RunEvent4("CreateFolder", path);
	if (r !== undefined) {
		return r; 
	}
	CreateNew(path, function (strPath)
	{
		fso.CreateFolder(strPath.replace(/^\s*/, ""));
	});
}

CreateFile = function (path)
{
	var r = MainWindow.RunEvent4("CreateFile", path);
	if (r !== undefined) {
		return r; 
	}
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

FormatDateTime = function (s)
{
	return new Date(s).getTime() > 0 ? (api.GetDateFormat(LOCALE_USER_DEFAULT, 0, s, api.GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE)) + " " + api.GetTimeFormat(LOCALE_USER_DEFAULT, 0, s, api.GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STIMEFORMAT))) : "";
};

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
			NavigateFV(GetFolderView(Ctrl, pt), ExtractPath(Ctrl, line[i], pt), NewTab);
			NewTab |= SBSP_NEWBROWSER;
		}
	}
	return S_OK;
}

DropOpen = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
{
	var line = s.split("\n");
	var hr = E_FAIL;
	var path = ExtractPath(Ctrl, line[0], pt);
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
	if (!s) {
		return S_FALSE;
	}
	window.Ctrl = Ctrl;
	window.hwnd = hwnd;
	window.dataObj = dataObj;
	window.grfKeyState = grfKeyState;
	window.pdwEffect = pdwEffect;
	window.bDrop = bDrop;
	if (pt) {
		window.pt = pt;
		te.Data.pt = pt;
	} else {
		window.pt = te.Data.pt;
	}
	window.Handled = S_OK;
	window.FV = GetFolderView(Ctrl, pt);

	if (/^Func$/i.test(type)) {
		return s(Ctrl, pt, hwnd, dataObj, grfKeyState, pdwEffect, bDrop, window.FV);
	}
	for (var i in eventTE.Exec) {
		var hr = eventTE.Exec[i](Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop, window.FV);
		if (isFinite(hr)) {
			return hr; 
		}
	}
	return window.Handled;
}

ExecScriptEx = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop, FV)
{
	var fn = null;
	try {
		if (/J.*Script/i.test(type)) {
			fn = {Handled: new Function(s)};
		} else if (/VBScript/i.test(type)) {
			fn = api.GetScriptDispatch('Function Handled(Ctrl, pt, hwnd, dataObj, grfKeyState, pdwEffect, bDrop, FV)\n' + s + '\nEnd Function', type, true);
		}
		if (fn) {
			var r = fn.Handled(Ctrl, pt, hwnd, dataObj, grfKeyState, pdwEffect, bDrop, FV);
			return isFinite(r) ? r : window.Handled;
		}
		api.ExecScript(s, type,
			{
				window: window,
				Ctrl: Ctrl,
				pt: pt,
				hwnd: hwnd,
				dataObj: dataObj,
				grfKeyState: grfKeyState,
				pdwEffect: pdwEffect,
				bDrop: bDrop,
				FV: FV
			}
		);
	} catch (e) {
		ShowError(e, s);
	}
	return window.Handled;
}

DropScript = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop, FV)
{
	if (!pdwEffect) {
		pdwEffect = api.Memory("DWORD");
	}
	if (/EnableDragDrop/.test(s)) {
		return ExecScriptEx(Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop, FV);
	}
	pdwEffect[0] = DROPEFFECT_NONE;
	return E_NOTIMPL;
}

ExtractPath = function (Ctrl, s, pt)
{
	s = api.PathUnquoteSpaces(ExtractMacro(Ctrl, GetConsts(s)));
	if (/^\.|^\\$/.test(s)) {
		var FV = GetFolderView(Ctrl, pt);
		if (FV) {
			if (s == "\\") {
				return fso.GetDriveName(FV.FolderItem.Path) + s;
			}
			if (s == "..") {
				return api.GetDisplayNameOf(api.ILGetParent(FV), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
			}
			var res = /\.\.\\(.*)/.exec(s);
			if (res) {
				return fso.BuildPath(api.GetDisplayNameOf(api.ILGetParent(FV), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING), res[1]);
			}
			res = /\.\\(.*)/.exec(s);
			if (res) {
				return fso.BuildPath(api.GetDisplayNameOf(FV, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING), res[1]);
			}
		}
	}
	return s;
}

ExtractMacro2 = function (Ctrl, s)
{
	for (var j = 99; j--;) {
		var s1 = s;
		for (var i in eventTE.ReplaceMacroEx) {
			s = s.replace(eventTE.ReplaceMacroEx[i][0], eventTE.ReplaceMacroEx[i][1]);
		}
		for (var i in eventTE.ReplaceMacro) {
			var re = eventTE.ReplaceMacro[i][0];
			var res = re.exec(s);
			if (res) {
				var r = eventTE.ReplaceMacro[i][1](Ctrl, re, res);
				if (/string|number/i.test(typeof r)) {
					s = s.replace(re, r);
				}
			}
		}
		for (var i in eventTE.ExtractMacro) {
			var re = eventTE.ExtractMacro[i][0];
			if (re.test(s)) {
				s = eventTE.ExtractMacro[i][1](Ctrl, s, re);
			}
		}
		s = s.replace(/%([\w\-_]+)%/g, function (strMatch, ref)
		{
			var fn = eventTE.Environment[ref.toLowerCase()];
			if (/string|number/i.test(typeof fn)) {
				return fn;
			} else if (fn) {
				try {
					var r = fn(Ctrl);
					if (/string|number/i.test(typeof r)) {
						return r;
					}
				} catch (e) {}
			}
			return strMatch;
		});
		s = wsh.ExpandEnvironmentStrings(s);
		if (s == s1) {
			break;
		}
	}
	return s;
}

ExtractMacro = function (Ctrl, s)
{
	if (/string/i.test(typeof s)) {
		s = ExtractMacro2(Ctrl, s);
		if (!/\t/.test(s) && /%/.test(s)) {
			do {
				s = ExtractMacro2(Ctrl, s.replace(/%/, "\t"));
			} while (/%/.test(s));
			s = s.replace(/\t/g, "%");
		}
	}
	return s;
}

AddEnv("Selected", function(Ctrl)
{
	var ar = [];
	var FV = GetFolderView(Ctrl);
	if (FV) {
		var Selected = FV.SelectedItems();
		if (Selected) {
			for (var i = Selected.Count; i > 0; ar.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(Selected.Item(--i), SHGDN_FORPARSING)))) {
			}
		}
	}
	return ar.join(" ");
});

AddEnv("Current", function(Ctrl)
{
	var strSel = "";
	var FV = GetFolderView(Ctrl);
	if (FV) {
		strSel = api.PathQuoteSpaces(api.GetDisplayNameOf(FV, SHGDN_FORPARSING));
	}
	return strSel;
});

AddEnv("TreeSelected", function(Ctrl)
{
	var strSel = "";
	if (!Ctrl || Ctrl.Type != CTRL_TV) {
		Ctrl = te.Ctrl(CTRL_TV);
	}
	if (Ctrl) {
		strSel = api.PathQuoteSpaces(api.GetDisplayNameOf(Ctrl.SelectedItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING));
	}
	return strSel;
});

AddEnv("Installed", fso.GetDriveName(api.GetModuleFileName(null)));

AddEnv("TE_Config", function ()
{
	return fso.BuildPath(te.Data.DataFolder, "config");
});

AddEvent("ReplaceMacroEx", [/%res:(\d+)%/ig, function (strMatch, ref1)
{
	return api.LoadString(hShell32, ref1) || "";
}]);

AddEvent("ReplaceMacroEx", [/%AddonStatus:([^%]*)%/ig, function (strMatch, ref1)
{
	return api.LowPart(GetAddonElement(ref1).getAttribute("Enabled")) ? "on" : "off";
}]);

PathMatchEx = function (path, s)
{
	var res = /^\/(.*)\/(.*)/.exec(s);
	return res ? new RegExp(res[1], res[2]).test(path) : api.PathMatchSpec(path, s);
}

IsFolderEx = function (Item)
{
	if (Item) {
		var wfd = api.Memory("WIN32_FIND_DATA");
		var hr = api.SHGetDataFromIDList(Item, SHGDFIL_FINDDATA, wfd, wfd.Size);
		if (Item.IsFolder) {
			return (hr < 0) || Boolean(wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY);
		}
	}
	return false;
}

OpenMenu = function (items, SelItem)
{
	var arMenu;
	var path = "";
	if (SelItem) {
		if (/object/i.test(typeof SelItem)) {
			var link = SelItem.ExtendedProperty("linktarget");
			path = link || String(api.GetDisplayNameOf(SelItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX));
			arMenu = OpenMenu(items, path);
			if (!IsFolderEx(SelItem) && (!link || !api.PathIsDirectory(path))) {
				return arMenu;
			}
			path += ".folder";
		} else {
			path = SelItem;	
		}
	}
	arMenu = [];
	var arLevel = [];
	for (var i = 0; i < items.length; i++) {
		var item = items[i];
		var strType = String(item.getAttribute("Type")).toLowerCase();
		var strFlag = strType == "menus" ? item.text.toLowerCase() : "";
		var bAdd = SelItem ? PathMatchEx(path, item.getAttribute("Filter")) : /^$|^\/\^\$\//.test(item.getAttribute("Filter"));
		if (strFlag == "close") {
			bAdd = arLevel.pop();
		}
		if (strFlag == "open") {
			arLevel.push(bAdd);
		}
		if (bAdd && (arLevel.length == 0 || arLevel[arLevel.length - 1])) {
			arMenu.push(i);
		}
	}
	return arMenu;
}

ExecMenu3 = function (Ctrl, Name, x, y)
{
	window.Ctrl = Ctrl;
	setTimeout(function () {
		ExecMenu2(Name, x, y);
	}, 99);;
}

ExecMenu2 = function (Name, x, y)
{
	if (!pt) {
		pt = api.Memory("POINT");
	}
	pt.x = x;
	pt.y = y;
	ExecMenu(Ctrl, Name, pt, 0);
}

AdjustMenuBreak = function (hMenu)
{
	var mii = api.Memory("MENUITEMINFO");
	mii.cbSize = mii.Size;
	var uFlags = 0;
	for (var i = api.GetMenuItemCount(hMenu); i-- > 0;) {
		mii.fMask = MIIM_FTYPE | MIIM_SUBMENU;
		api.GetMenuItemInfo(hMenu, i, true, mii);
		if (mii.hSubMenu) {
			AdjustMenuBreak(mii.hSubMenu);
			continue;
		}
		if (api.GetMenuString(hMenu, i, MF_BYPOSITION) != "") {
			mii.fType |= uFlags;
			api.SetMenuItemInfo(hMenu, i, true, mii);
			uFlags = 0;
			continue;
		}
		var u = mii.fType & (MFT_MENUBREAK | MFT_MENUBARBREAK);
		if (u) {
			uFlags = u;
			api.DeleteMenu(hMenu, i++, MF_BYPOSITION);
		}
	}
	for (var i = api.GetMenuItemCount(hMenu); i--;) {
		mii.fMask = MIIM_FTYPE;
		api.GetMenuItemInfo(hMenu, i, true, mii);
		if ((mii.fType & MFT_SEPARATOR) || api.GetMenuString(hMenu, i, MF_BYPOSITION).charAt(0) == '{') {
			api.DeleteMenu(hMenu, i, MF_BYPOSITION);
			continue;
		}
		break;
	}
	uFlags = 0;
	for (var i = api.GetMenuItemCount(hMenu); i--;) {
		mii.fMask = MIIM_FTYPE;
		api.GetMenuItemInfo(hMenu, i, true, mii);
		if (uFlags & mii.fType & MFT_SEPARATOR) {
			api.DeleteMenu(hMenu, i, MF_BYPOSITION);
		}
		uFlags = mii.fType;
	}
}

teMenuGetElementsByTagName = function (Name)
{
	var menus = te.Data.xmlMenus.getElementsByTagName(Name);
	if (!menus || !menus.length) {
		var altMenu = {
			"ViewContext": "Background",
			"Background": "ViewContext",
			"TaskTray": "Systray",
			"Systray": "TaskTray"
		}
		menus = te.Data.xmlMenus.getElementsByTagName(altMenu[Name]);
	}
	return menus;
}

ExecMenu = function (Ctrl, Name, pt, Mode, bNoExec)
{
	var items = null;
	var menus = teMenuGetElementsByTagName(Name);
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
	ExtraMenuCommand = [];
	ExtraMenuData = [];
	eventTE.MenuCommand = [];
	var arMenu;
	var item;
	if (items) {
		arMenu = OpenMenu(items, SelItem);
		if (arMenu.length) {
			item = items[arMenu[0]];
		}
		var nBase = api.LowPart(menus[0].getAttribute("Base"));
		if (nBase == 1) {
			if (api.LowPart(menus[0].getAttribute("Pos")) < 0) {
				item = items[arMenu[arMenu.length - 1]];
				if (arMenu.length > 1) {
					for (var i = arMenu.length; i--;) {
						var nLevel = 0;
						if (String(items[arMenu[i]].getAttribute("Type")).toLowerCase() == "menus") {
		 					var s = String(items[arMenu[i]].text).toLowerCase();
		 					if (s == "close") {
		 						nLevel++;
		 					}
		 					if (s == "open") {
		 						if (--nLevel < 0) {
		 							arMenu.splice(0, i + 1);
									nBase = 0;
									break;
								}
							}
						}
					}
				}
			}
		}
		if (nBase != 1) {
			var hMenu = api.CreatePopupMenu();
			var ContextMenu = GetBaseMenuEx(hMenu, nBase, FV, Selected, uCMF, Mode, SelItem);
			if (nBase < 5) {
				AdjustMenuBreak(hMenu);
			}
			g_nPos = MakeMenus(hMenu, menus, arMenu, items, Ctrl, pt);
			for (var i in eventTE[Name]) {
				g_nPos = eventTE[Name][i](Ctrl, hMenu, g_nPos, Selected, SelItem, ContextMenu, Name, pt);
			}
			for (var i in eventTE.Menus) {
				g_nPos = eventTE.Menus[i](Ctrl, hMenu, g_nPos, Selected, SelItem, ContextMenu, Name, pt);
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
					case CTRL_TV:
						var rc = api.Memory("RECT");
						if (Ctrl.GetItemRect(SelItem, rc) != S_OK) {
							api.GetClientRect(Ctrl.hwnd, rc);
						}
						api.GetCursorPos(pt);
						api.ScreenToClient(Ctrl.hwnd, pt);
						if (!PtInRect(rc, pt)) {
							pt.x = rc.Left;
							pt.y = rc.Top;
						}
						api.ClientToScreen(Ctrl.hwnd, pt);
						break;
					default:
						api.ClientToScreen(te.hwnd, pt);
						break;
				}
			}
			AdjustMenuBreak(hMenu);
			window.g_menu_click = bNoExec ? true : 2;
			var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, ContextMenu);
			if (bNoExec) {
				return nVerb > 0 ? S_OK : S_FALSE;
			} else {
				var hr = ExecMenu4(Ctrl, Name, pt, hMenu, [ContextMenu], nVerb, FV);
				if (isFinite(hr)) {
					return hr;
				}
				item = items[nVerb - 1];
				Mode = 0;
			}
		}
		if (item && !bNoExec) {
			var s = item.getAttribute("Type");
			if (window.g_menu_button == 2 && api.PathMatchSpec(s, "Open;Open in New Tab;Open in Background")) {
				PopupContextMenu(item.text);
				return S_OK;
			}
			Exec(Ctrl, item.text, window.g_menu_button == 3 && s == "Open" ? "Open in New Tab" : s, Ctrl.hwnd, pt);
			return S_OK;
		}
		if (Mode != 2) {
			return (nBase != 2 || ContextMenu) ? S_OK : S_FALSE;
		}
	}
	return S_FALSE;
}

ExecMenu4 = function (Ctrl, Name, pt, hMenu, arContextMenu, nVerb, FV)
{
	if (ExtraMenuCommand[nVerb]) {
		if (ExtraMenuCommand[nVerb](Ctrl, pt, Name, nVerb) != S_FALSE) {
			api.DestroyMenu(hMenu);
			return S_OK;
		}
	}
	for (var i in eventTE.MenuCommand) {
		var hr = eventTE.MenuCommand[i](Ctrl, pt, Name, nVerb, hMenu);
		if (isFinite(hr) && hr == S_OK) {
			api.DestroyMenu(hMenu);
			return S_OK;
		}
	}
	for (var i in arContextMenu) {
		var ContextMenu = arContextMenu[i];
		if (ContextMenu && nVerb >= ContextMenu.idCmdFirst && nVerb <= ContextMenu.idCmdLast) {
			var FolderView = ContextMenu.FolderView;
			if (FolderView) {
				FolderView.Focus();
			}
			if (ContextMenu.InvokeCommand(0, te.hwnd, nVerb - ContextMenu.idCmdFirst, null, null, SW_SHOWNORMAL, 0, 0) == S_OK) {
				api.DestroyMenu(hMenu);
				return S_OK;
			}
		}
	}
	api.DestroyMenu(hMenu);
	if (FV && nVerb > 0x7000) {
		if (api.SendMessage(FV.hwndView, WM_COMMAND, nVerb, 0) == S_OK) {
			return S_OK;
		}
	}
}

CopyMenu = function (hSrc, hDest)
{
	var mii = api.Memory("MENUITEMINFO");
	mii.cbSize = mii.Size;
	mii.fMask = MIIM_ID | MIIM_TYPE | MIIM_SUBMENU | MIIM_STATE;
	var n = api.GetMenuItemCount(hSrc);
	while (--n >= 0) {
		api.GetMenuItemInfo(hSrc, n, true, mii);
		var hSubMenu = mii.hSubMenu;
		if (hSubMenu) {
			mii.hSubMenu = api.CreateMenu();
		}
		api.InsertMenuItem(hDest, 0, true, mii);
		if (hSubMenu) {
			CopyMenu(hSubMenu, mii.hSubMenu);
		}
	}
}

GetViewMenu = function (arContextMenu, FV, hMenu, uCMF)
{
	var ContextMenu = arContextMenu && arContextMenu[0];
	if (!ContextMenu) {
		ContextMenu = FV.ViewMenu();
		if (arContextMenu) {
			arContextMenu[0] = ContextMenu;
		}
	}
	if (ContextMenu) {
		ContextMenu.QueryContextMenu(hMenu, 0, 0x5001, 0x5fff, uCMF);
	}
	return ContextMenu;
}

GetBaseMenuEx = function (hMenu, nBase, FV, Selected, uCMF, Mode, SelItem, arContextMenu)
{
	var ContextMenu;
	for (var i in eventTE.GetBaseMenuEx) {
		ContextMenu = eventTE.GetBaseMenuEx[i](hMenu, nBase, FV, Selected, uCMF, Mode, SelItem, arContextMenu);
		if (ContextMenu !== undefined) {
			return ContextMenu; 
		}
	}
	switch (nBase) {
		case 2:
		case 4:
			var Items = Selected;
			if (!Items || !Items.Count) {
				Items = SelItem;
			}
			if (nBase == 2 || Items && Items.Count) {
				ContextMenu = arContextMenu && arContextMenu[1];
				if (!ContextMenu) {
					ContextMenu = api.ContextMenu(Items, FV);
					if (arContextMenu) {
						arContextMenu[1] = ContextMenu;
					}
				}
				if (ContextMenu) {
					ContextMenu.QueryContextMenu(hMenu, 0, 0x6001, 0x6fff, uCMF | CMF_DONOTPICKDEFAULT);
					if (!Items.Count) {
						SetRenameMenu(ContextMenu.idCmdFirst);
					}
				}
			} else if (FV) {
				ContextMenu = GetViewMenu(arContextMenu, FV, hMenu, uCMF | CMF_DONOTPICKDEFAULT);
				var mii = api.Memory("MENUITEMINFO");
				mii.cbSize = mii.Size;
				mii.fMask = MIIM_FTYPE | MIIM_SUBMENU;
				for (var i = api.GetMenuItemCount(hMenu); i--;) {
					api.GetMenuItemInfo(hMenu, 0, true, mii);
					if (mii.hSubMenu || (mii.fType & MFT_SEPARATOR)) {
						api.DeleteMenu(hMenu, 0, MF_BYPOSITION);
						continue;
					}
					break;
				}
			}
			break;
		case 3:
			if (FV) {
				ContextMenu = GetViewMenu(arContextMenu, FV, hMenu, uCMF | CMF_DONOTPICKDEFAULT);
			}
			break;
		case 5:
		case 6:
			var id = nBase == 5 ? FCIDM_MENU_EDIT : FCIDM_MENU_VIEW;
			if (FV) {
				ContextMenu = GetViewMenu(arContextMenu, FV, hMenu, CMF_DEFAULTONLY);
				var hMenu2 = te.MainMenu(id);
				var oMenu = {};
				var oMenu2 = {};
				var mii = api.Memory("MENUITEMINFO");
				mii.cbSize = mii.Size;
				mii.fMask = MIIM_SUBMENU;
				for (var i = api.GetMenuItemCount(hMenu2); i-- > 0;) {
					var s = api.GetMenuString(hMenu2, i, MF_BYPOSITION);
					if (s) {
						s = s.toLowerCase().replace(/[&\(\)]/g, "");
						api.GetMenuItemInfo(hMenu2, i, true, mii);
						oMenu2[s] = mii.hSubMenu;
					}
				}
				MenuDbInit(hMenu, oMenu, oMenu2);
				MenuDbReplace(hMenu, oMenu, hMenu2);
			} else {
				var hMenu1 = te.MainMenu(id);
				CopyMenu(hMenu1, hMenu);
				api.DestroyMenu(hMenu1);
			}
			break;
		case 7:
			var dir = GetHelpMenu(true);
			for (var i = 0; i < dir.length; i++) {
				var s = dir[i];
				if (s === null) {
					api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
				} else {
					if (s) {
						api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, i + 0x4011, s);
					}
				}
			}
			AddEvent("MenuCommand", function (Ctrl, pt, Name, nVerb)
			{
				var s = GetHelpMenu(false)[nVerb - 0x4011];
				if (s) {
					if (/^function$/i.test(typeof s)) {
						s(Ctrl, pt, Name, nVerb);
					} else {
						Navigate(s, SBSP_NEWBROWSER);
					}
					return S_OK;
				}
			});
			break;
		case 8:
			api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 0x4001, GetText("&Add to Favorites..."));
			ExtraMenuCommand[0x4001] = AddFavoriteEx;
			api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 0x4002, GetText("&Edit"));
			ExtraMenuCommand[0x4002] = function ()
			{
				ShowOptions("Tab=Menus&Menus=Favorites");
			};
			api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
			break;
		default:
			break;
	}
	return ContextMenu;
}

GetHelpMenu = function (bTitle)
{
	var dir = [fso.BuildPath(te.Data.DataFolder, "config") , null, ssfDRIVES, ssfNETHOOD, ssfWINDOWS, ssfSYSTEM, ssfPROGRAMFILES];
	if (api.sizeof("HANDLE") > 4) {
		dir.push(ssfPROGRAMFILESx86);
	} else if (api.IsWow64Process(api.GetCurrentProcess())) {
		dir.push(api.GetDisplayNameOf(ssfPROGRAMFILES, SHGDN_FORPARSING).replace(/\s*\(x86\)$/i, ""));
	}
	dir.push(fso.GetSpecialFolder(2).Path);
	if (WINVER >= 0x600) {
		dir.push("shell:libraries");
	}
	dir = dir.concat([ssfPERSONAL, ssfSTARTMENU, ssfPROGRAMS, ssfSTARTUP, ssfSENDTO, ssfAPPDATA, ssfFAVORITES, ssfRECENT, ssfHISTORY, ssfDESKTOPDIRECTORY, ssfCONTROLS, ssfTEMPLATES, ssfFONTS, ssfPRINTERS, ssfBITBUCKET]);
	if (bTitle) {
		for (var i = dir.length; i--;) {
			dir[i] = api.GetDisplayNameOf(dir[i], SHGDN_INFOLDER);
		}
		return [te.About, GetText("Check for updates"), GetText("Get Add-ons")].concat(dir);
	}
	return [api.GetModuleFileName(null) + "\\..", CheckUpdate, GetAddons].concat(dir);
}

MenuDbInit = function (hMenu, oMenu, oMenu2)
{
	for (var i = api.GetMenuItemCount(hMenu); i--;) {
		var mii = api.Memory("MENUITEMINFO");
		mii.cbSize = mii.Size;
		mii.fMask = MIIM_ID | MIIM_BITMAP | MIIM_SUBMENU | MIIM_DATA | MIIM_FTYPE | MIIM_STATE;
		var s = api.GetMenuString(hMenu, i, MF_BYPOSITION);
		api.GetMenuItemInfo(hMenu, i, true, mii);
		if (s) {
			s = s.toLowerCase().replace(/[&\(\)]/g, "");
			oMenu[s] = mii;
			api.RemoveMenu(hMenu, i, MF_BYPOSITION);
			if (oMenu2 && mii.hSubMenu && !oMenu2[s]) {
				MenuDbInit(mii.hSubMenu, oMenu, null)
			}
		} else {
			api.DeleteMenu(hMenu, i, MF_BYPOSITION);
		}
	}
}

MenuDbReplace = function (hMenu, oMenu, hMenu2)
{
	for (var i = api.GetMenuItemCount(hMenu2); i-- > 0;) {
		var s = api.GetMenuString(hMenu2, 0, MF_BYPOSITION);
		var mii = null;
		var s2 = null;
		if (s) {
			s2 = s.toLowerCase().replace(/[&\(\)]/g, "");
			mii = oMenu[s2];
			if (!mii) {
				s2 = s2.replace(/\t.*/, "");
				mii = oMenu[s2];
			}
		}
		if (mii) {
			delete oMenu[s2];
			api.DeleteMenu(hMenu2, 0, MF_BYPOSITION);
		} else {
			mii = api.Memory("MENUITEMINFO");
			mii.cbSize = mii.Size;
			mii.fMask = MIIM_ID | MIIM_BITMAP | MIIM_SUBMENU | MIIM_DATA | MIIM_FTYPE | MIIM_STATE;
			api.GetMenuItemInfo(hMenu2, 0, true, mii);
			if (mii.hSubMenu) {
				api.DeleteMenu(hMenu2, 0, MF_BYPOSITION);
				continue;
			} else {
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
		if (!/^\t/.test(s)) {
			api.InsertMenuItem(hMenu2, MAXINT, false, oMenu[s]);
		}
	}
	api.DestroyMenu(hMenu2);
}

GetAccelerator = function (s)
{
	var res = /&(.)/.exec(s);
	return res ? res[1] : "";
}

AddMenuIconFolderItem = function (mii, FolderItem)
{
	var image = te.GdiplusBitmap();
	var sfi = api.Memory("SHFILEINFO");
	var path = /string/i.test(typeof FolderItem) ? FolderItem : api.GetDisplayNameOf(FolderItem, SHGDN_FORPARSING);
	if (api.PathIsNetworkPath(path)) {
		if (fso.GetDriveName(path) != path.replace(/\\$/, "")) {
			MenusIcon(mii, WINVER >= 0x600 ? "icon:shell32.dll,275,16" : "icon:shell32.dll,85,16");
			return;
		}
		MenusIcon(mii, WINVER >= 0x600 ? "icon:shell32.dll,273,16" : "icon:shell32.dll,9,16");
		return;
	}
	api.SHGetFileInfo(FolderItem, 0, sfi, sfi.Size, SHGFI_SYSICONINDEX | SHGFI_PIDL);
	var id = sfi.iIcon;
	mii.hbmpItem = MainWindow.g_arBM['U' + id];
	if (mii.hbmpItem) {
		mii.fMask = mii.fMask | MIIM_BITMAP;
		return;
	}
	var hIcon = api.ImageList_GetIcon(te.Data.SHIL[SHIL_SMALL], id, ILD_NORMAL);
	image.FromHICON(hIcon);
	api.DestroyIcon(hIcon);
	AddMenuImage(mii, image, id);
}

AddMenuImage = function (mii, image, id)
{
	mii.hbmpItem = image.GetHBITMAP(WINVER >= 0x600 ? null : GetSysColor(COLOR_MENU));
	if (mii.hbmpItem) {
		mii.fMask = mii.fMask | MIIM_BITMAP;
		if (id) {
			MainWindow.g_arBM['U' + id] = mii.hbmpItem;
		} else {
			MainWindow.g_arBM.push(mii.hbmpItem);
		}
	}
}

MenusIcon = function (mii, src)
{
	mii.cbSize = mii.Size;
	if (src) {
		src = ExtractMacro(te, src);
		var image = te.GdiplusBitmap();
		mii.hbmpItem = MainWindow.g_arBM['U' + src];
		if (mii.hbmpItem) {
			mii.fMask = mii.fMask | MIIM_BITMAP;
			return;
		}
		if (REGEXP_IMAGE.test(src)) {
			image.FromFile(src);
		} else {
			var hIcon = MakeImgIcon(src, 0, 16);
			image.FromHICON(hIcon);
			api.DestroyIcon(hIcon);
		}
		AddMenuImage(mii, image, src);
	}
}

MakeMenus = function (hMenu, menus, arMenu, items, Ctrl, pt, nMin, arItem)
{
	var hMenus = [hMenu];
	var nPos = menus ? Number(menus[0].getAttribute("Pos")) : 0;
	var nLen = api.GetMenuItemCount(hMenu);
	var nResult = 0;
	nMin = nMin || 0;
	if (nPos < 0) {
		nPos += nLen + 1;
	}
	if (nPos > nLen || nPos < 0) {
		nPos = nLen;
	}
	nLen = arMenu.length;
	for (var i = 0; i < nLen; i++) {
		var item = items[arMenu[i]];
		var s = (item.getAttribute("Name") || item.getAttribute("Mouse") || GetKeyName(item.getAttribute("Key")) || "").replace(/\\t/i, "\t");
		var strFlag = String(item.getAttribute("Type")).toLowerCase() == "menus" ? item.text.toLowerCase() : "";
		if (strFlag == "close") {
			hMenus.pop();
			if (!hMenus.length) {
				break;
			}
		} else {
			var ar = s.split(/\t/);
			if (!item.getAttribute("Org")) {
				ar[0] = GetText(ar[0]);
			}
			if (ar.length > 1) {
				ar[1] = GetKeyName(ar[1]);
			}
			if (strFlag == "open") {
				var mii = api.Memory("MENUITEMINFO");
				mii.fMask = MIIM_STRING | MIIM_SUBMENU | MIIM_FTYPE;
				mii.fType = 0;
				mii.dwTypeData = ar.join("\t");
				mii.hSubMenu = api.CreateMenu();
				MenusIcon(mii, item.getAttribute("Icon"));
				api.InsertMenuItem(hMenus[hMenus.length - 1], nPos++, true, mii);
				hMenus.push(mii.hSubMenu);
			} else {
				nResult = arMenu[i] + nMin + 1;
				if (s == "/" || strFlag == "break") {
					api.InsertMenu(hMenus[hMenus.length - 1], nPos++, MF_BYPOSITION | MF_MENUBREAK | MF_DISABLED, 0, "");
				} else if (s == "//" || strFlag == "barbreak") {
					api.InsertMenu(hMenus[hMenus.length - 1], nPos++, MF_BYPOSITION | MF_MENUBARBREAK | MF_DISABLED, 0, "");
				} else if (s == "-" || strFlag == "separator") {
					api.InsertMenu(hMenus[hMenus.length - 1], nPos++, MF_BYPOSITION | MF_SEPARATOR, 0, null);
				} else if (s) {
					var mii = api.Memory("MENUITEMINFO");
					mii.fMask = MIIM_STRING | MIIM_ID;
					mii.wID = nResult;
					mii.dwTypeData = ar.join("\t");
					MenusIcon(mii, item.getAttribute("Icon"));
					RunEvent3(["MenuState", item.getAttribute("Type"), item.text].join(":"), Ctrl, pt, mii);
					if (arItem) {
						arItem[nResult - 1] = items[arMenu[i]];
					}
					api.InsertMenuItem(hMenus[hMenus.length - 1], nPos++, true, mii);
				}
			}
		}
	}
	return nResult > nMin ? nResult : nMin;
}

SaveXmlEx = function (filename, xml)
{
	try {
		filename = fso.BuildPath(te.Data.DataFolder, "config\\" + filename);
		xml.save(filename);
	} catch (e) {
		if (e.number != E_ACCESSDENIED) {
			ShowError(e, [GetText("Save"), filename].join(": "));
		}
	}
}

BlurId = function (Id)
{
	document.getElementById(Id).blur();
}

RunCommandLine = function (s)
{
	var re = /\/select,([^,]+)/i.exec(s);
	if (re) {
		var arg = api.CommandLineToArgv(re[1]);
		Navigate(fso.GetParentFolderName(arg[0]), SBSP_NEWBROWSER);
		(function (Item) { setTimeout(function ()
		{
			var FV = te.Ctrl(CTRL_FV);
			FV.SelectItem(Item, SVSI_SELECT | SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS);
		}, 99);}) (arg[0]);
		return;
	}
	var arg = api.CommandLineToArgv(s.replace(/^\/e,|^\/n,|^\/root,/ig, ""));
	arg.shift();
	var s  = arg.join(" ");
	if (/^[A-Z]:\\|^\\\\/i.test(s) && IsExists(s)) {
		Navigate(s, SBSP_NEWBROWSER);
		return;
	}
	while (s = arg.shift()) {
		var ar = s.split(",");
		if (ar.length > 1) {
			Exec(te, GetSourceText(ar[1]), GetSourceText(ar[0]), te.hwnd, api.Memory("POINT"))
			continue;
		}
		Navigate(s, SBSP_NEWBROWSER);
	}
}

OpenNewProcess = function (fn, ex, mode, vOperation)
{
	var uid;
	do {
		uid = String(Math.random()).replace(/^0?\./, "");
	} while (MainWindow.Exchange[uid]);
	MainWindow.Exchange[uid] = ex;
	return ShellExecute([api.PathQuoteSpaces(api.GetModuleFileName(null)), mode ? '/open' : '/run', fn, uid].join(" "), vOperation, SW_SHOWNORMAL);
}

GetAddonInfo = function (Id)
{
	var info = [];

	var path = fso.GetParentFolderName(api.GetModuleFileName(null));
	var xml = te.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	var xmlfile = fso.BuildPath(path, "addons\\" + Id + "\\config.xml");
	if (fso.FileExists(xmlfile)) {
		xml.load(xmlfile);

		GetAddonInfo2(xml, info, "General", true);
		var lang = GetLangId();
		if (!/^en/.test(lang)) {
			GetAddonInfo2(xml, info, "en", true);
		}
		var res = /(\w+)_/.exec(lang);
		if (res) {
			GetAddonInfo2(xml, info, res[1]);
		}
		GetAddonInfo2(xml, info, lang);
		if (!info.Name) {
			info.Name = Id;
		}
	}
	return info;
}

GetAddonInfo2 = function (xml, info, Tag, bTrans)
{
	var items = xml.getElementsByTagName(Tag);
	if (items.length) {
		var item = items[0].childNodes;
		for (var i = 0; i < item.length; i++) {
			var n = item[i].tagName;
			info[n] = bTrans && /Name|Description/i.test(n) ? GetText(item[i].text) : item[i].text;
		}
	}
}

OpenXml = function (strFile, bAppData, bEmpty, strInit)
{
	var xml = te.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	var path = fso.BuildPath(te.Data.DataFolder, "config\\" + strFile);
	if (fso.FileExists(path) && xml.load(path)) {
		return xml;
	}
	if (!bAppData) {
		path = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "config\\" + strFile);
		if (fso.FileExists(path) && xml.load(path)) {
			api.SHFileOperation(FO_MOVE, path, fso.BuildPath(te.Data.DataFolder, "config"), FOF_SILENT | FOF_NOCONFIRMATION, false);
			return xml;
		}
	}
	if (strInit) {
		path = fso.BuildPath(strInit, strFile);
		if (fso.FileExists(path) && xml.load(path)) {
			return xml;
		}
	}
	path = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "init\\" + strFile);
	if (fso.FileExists(path) && xml.load(path)) {
		return xml;
	}
	return bEmpty ? xml : null;
}

CreateXml = function (bRoot)
{
	var xml = te.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	xml.appendChild(xml.createProcessingInstruction("xml", 'version="1.0" encoding="UTF-8"'));
	if (bRoot) {
		xml.appendChild(xml.createElement("TablacusExplorer"));
	}
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
	} catch (e) {
		if (api.Extract(fso.BuildPath(system32, "zipfldr.dll"), "{E88DCCE0-B7B3-11d1-A9F0-00AA0060FA31}", Src, Dest) != S_OK) {
			MessageBox(GetText("Extract Error"), TITLE, MB_OK);
		}
	}
	return S_OK;
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

function GetAddons()
{
	ShowOptions("Tab=Get Addons");
}

function CheckUpdate()
{
	OpenHttpRequest("http://www.eonet.ne.jp/~gakana/tablacus/explorer_en.html", "text/html", CheckUpdate2);
}

function CheckUpdate2(xhr)
{
	var res = /<td id="te">(.*?)<\/td>/i.exec(xhr.responseText);
	if (!res) {
		return;
	}
	var s = res[1];
	res = /<a href="dl\/([^"]*)/i.exec(s);
	if (!res) {
		return;
	}
	var file = res[1];
	s = s.replace(/Download/i, "").replace(/<[^>]*>/ig, "");
	var ver = 0;
	res = /(\d+)/.exec(file);
	if (res) {
		ver = api.Add(20000000, res[1]);
	}
	if (ver <= te.Version) {
		MessageBox(te.About + "\n" + GetText("the latest version"), TITLE, MB_ICONINFORMATION);
		if (api.GetKeyState(VK_SHIFT) >= 0 || api.GetKeyState(VK_CONTROL) >= 0) {
			return;
		}
	}
	if (!confirmOk(GetText("Update available") + "\n" + s + "\n" + GetText("Do you want to install it now?"))) {
		return;
	}
	var temp = fso.BuildPath(fso.GetSpecialFolder(2).Path, "tablacus");
	CreateFolder2(temp);
	wsh.CurrentDirectory = temp;
	var InstalledFolder = fso.GetParentFolderName(api.GetModuleFileName(null));
	var zipfile = fso.BuildPath(temp, file);
	temp += "\\explorer";
	DeleteItem(temp);
	CreateFolder2(temp);

	if (DownloadFile("http://www.eonet.ne.jp/~gakana/tablacus/dl/" + file, zipfile) != S_OK || Extract(zipfile, temp) != S_OK) {
		return;
	}
	var te64exe = temp + "\\te64.exe";
	var nDog = 300;
	while (!fso.FileExists(te64exe)) {
		if (wsh.Popup(GetText("Please wait."), 1, TITLE, MB_ICONINFORMATION | MB_OKCANCEL) == IDCANCEL || nDog-- == 0) {
			return;
		}
	}
	var arDel = [];
	var addons = temp + "\\addons";

	for (var list = new Enumerator(fso.GetFolder(addons).SubFolders); !list.atEnd(); list.moveNext()) {
		var n = list.item().Name;
		var items = te.Data.Addons.getElementsByTagName(n);
		if (!items || items.length == 0) {
			arDel.push(fso.BuildPath(addons, n));
		}
	}
	if (arDel.length) {
		api.SHFileOperation(FO_DELETE, arDel.join("\0"), null, FOF_SILENT | FOF_NOCONFIRMATION, false);
	}
	var Taskkill = "";
	var pid = api.Memory("DWORD");
	api.GetWindowThreadProcessId(te.hwnd, pid);
	if (fso.FileExists(fso.BuildPath(system32, "taskkill.exe"))) {
		Taskkill = "W.Run('taskkill /pid " + pid[0] + " /f',2,true);";
	}
	var update = api.sprintf(2000, "\
F='%s';\
Q='\\x22';\
A=new ActiveXObject('Shell.Application');\
W=new ActiveXObject('WScript.Shell');\
W.Popup('%s',9,'Tablacus Explorer',%d);\
%s\
A.NameSpace(F).MoveHere(A.NameSpace('%s').Items(),%d);\
W.Run(Q+F+'\\\\%s'+Q);\
close()", EscapeUpdateFile(InstalledFolder), GetText("Please wait."), MB_ICONINFORMATION, Taskkill, EscapeUpdateFile(temp), FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR,
EscapeUpdateFile(fso.GetFileName(api.GetModuleFileName(null)))).replace(/[\t\n]/g, "");
	wsh.CurrentDirectory = temp;
	var exe = "mshta.exe";
	var s1 = '"javascript:';
	if (update.length >= 500 || !fso.FileExists(fso.BuildPath(system32, exe))) {
		exe = "wscript.exe";
		s1 = fso.GetParentFolderName(temp) + "\\update.js";
		DeleteItem(s1);
		var a = fso.CreateTextFile(s1, true);
		a.WriteLine(update.replace(/close\(\)$/, ""));
		a.Close();
		update = s1;
		s1 = '"';
	}
	MainWindow.g_strUpdate = ['"', api.IsWow64Process(api.GetCurrentProcess()) ? wsh.ExpandEnvironmentStrings("%windir%\\Sysnative") : system32, "\\", exe, '" ', s1, update, '"'].join("");
	MainWindow.DeleteTempFolder = MainWindow.UpdateSelf;
	WmiProcess("WHERE ExecutablePath = '" + api.GetModuleFileName(null).replace(/\\/g, "\\\\") + "' AND ProcessId!=" + pid[0], function (item)
	{
		item.Terminate();
	});
	api.PostMessage(te.hwnd, WM_CLOSE, 0, 0);
}

function EscapeUpdateFile(s)
{
	return s.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

confirmYN = function (s, title)
{
	return MessageBox(s, title, MB_ICONQUESTION | MB_YESNO) == IDYES;
}

confirmOk = function (s, title)
{
	return MessageBox(s || "Are you sure?", title, MB_ICONQUESTION | MB_OKCANCEL) == IDOK;
}

MessageBox = function (s, title, uType)
{
	return api.MessageBox(api.GetForegroundWindow(), GetText(s), GetText(title) || TITLE, uType);
}

createHttpRequest = function ()
{
	try {
		return te.CreateObject("Msxml2.XMLHTTP");
	} catch (e) {
		return te.CreateObject("Microsoft.XMLHTTP");
	}
}

OpenHttpRequest = function (url, ct, fn)
{
	var xhr = createHttpRequest();
	xhr.onreadystatechange = function()
	{
		if (xhr.readyState == 4) {
			if (xhr.status == 200) {
				fn(xhr);
			}
		}
	}
	xhr.open("GET", url + "?" + Math.floor(new Date().getTime() / 60000), false);
	xhr.setRequestHeader('Content-Type', ct);
	xhr.setRequestHeader('Pragma', 'no-cache');
	xhr.setRequestHeader('Cache-Control', 'no-cache');
	xhr.setRequestHeader('If-Modified-Since', 'Thu, 01 Jun 1970 00:00:00 GMT');
	try {
		xhr.send(null);
	} catch (e) {
		ShowError(e);
	}
}

InputDialog = function (text, defaultText)
{
	return prompt(GetText(text), defaultText);
}

AddonOptions = function (Id, fn, Data)
{
	var sParent = fso.GetParentFolderName(api.GetModuleFileName(null));
	LoadLang2(fso.BuildPath(sParent, "addons\\" + Id + "\\lang\\" + GetLangId() + ".xml"));
	var items = te.Data.Addons.getElementsByTagName(Id);
	if (!items.length) {
		var root = te.Data.Addons.documentElement;
		if (root) {
			root.appendChild(te.Data.Addons.createElement(Id));
		}
	}
	var info = GetAddonInfo(Id);
	var sURL = "../addons/" + Id + "/options.html";
	if (!Data) {
		Data = {};
	}
	Data.id = Id;
	var sFeatures = info.Options;
	var res = /Common:([\d,]+):(\d)/i.exec(sFeatures);
	if (res) {
		sURL = "location.html";
		Data.show = res[1];
		Data.index = res[2];
		sFeatures = 'Default';
	}
	if (/^Location$/i.test(sFeatures)) {
		sURL = "location.html";
		Data.show = "6";
		Data.index = "6";
		sFeatures = 'Default';
	}
	if (/^Default$/i.test(sFeatures)) {
		sFeatures = 'Width: 640; Height: 480';
	}
	try {
		var dlg = MainWindow.g_.dlgs[Id];
		if (dlg) {
			dlg.Focus();
			return;
		}
	} catch (e) {
		delete MainWindow.g_.dlgs[Id];
	}
	var opt = {MainWindow: MainWindow, Data: Data, event: {}};
	if (fn) {
		opt.event.TEOk = fn;
	} else if (window.g_Chg) {
		opt.event.TEOk = function ()
		{
			g_Chg.Addons = true;
		}
	}
	res = /width: *([0-9]+)/i.exec(sFeatures);
	if (res) {
		opt.width = res[1] - 0;
		res = /height: *([0-9]+)/i.exec(sFeatures);
		if (res) {
			opt.height = res[1] - 0;
		}
	}
	opt.event.onbeforeunload = function () {
		delete MainWindow.g_.dlgs[Id];
	}
	MainWindow.g_.dlgs[Id] = ShowDialog(sURL, opt);
}

function CalcVersion(s)
{
	var r = 0;
	var res = /(\d+)\.(\d+)\.(\d+)\.(\d+)/.exec(s);
	if (res) {
		return api.sprintf(99, "%04x%04x%04x%04x", res[1], res[2], res[3], res[4]);
	}
	res = /(\d+)\.(\d+)\.(\d+)/.exec(s);
	if (res) {
		r = res[1] * 10000 + res[2] * 100 + (res[3] - 0);
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
		api.WaitForInputIdle(hProcess, 9999);
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
	} while (hwnd = api.GetWindow(hwnd, GW_HWNDNEXT));
	return null;
}

PopupContextMenu = function (Item, FV)
{
	if (/string/i.test(typeof Item)) {
		var arg = api.CommandLineToArgv(Item);
		Item = te.FolderItems();
		for (var i in arg) {
			Item.AddItem(arg[i]);
		}
	}
	var hMenu = api.CreatePopupMenu();
	var ContextMenu = api.ContextMenu(Item, FV);
	if (ContextMenu) {
		var uCMF = (api.GetKeyState(VK_SHIFT) < 0) ? CMF_EXTENDEDVERBS : CMF_NORMAL;
		ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, uCMF);
		RemoveCommand(hMenu, ContextMenu, "delete");
		var pt = api.Memory("POINT");
		api.GetCursorPos(pt);
		var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, ContextMenu);
		g_popup = null;
		if (nVerb) {
			ContextMenu.InvokeCommand(0, te.hwnd, nVerb - 1, null, null, SW_SHOWNORMAL, 0, 0);
		}
	}
	api.DestroyMenu(hMenu);
}

GetAddonElement = function (id)
{
	var items = te.Data.Addons.getElementsByTagName(id);
	if (items.length) {
		return items[0];
	}
	return {
		getAttribute: function (s) {
			return "";
		},
		setAttribute: function () {}
	}
}

GetAddonOption = function (id, strTag)
{
	return GetAddonElement(id).getAttribute(strTag);
}

GetAddonOptionEx = function (id, strTag)
{
	return api.LowPart(GetAddonOption(id, strTag));
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
		CancelWindowRegistered();
		ShellExecute(api.PathQuoteSpaces(api.GetDisplayNameOf(FV, SHGDN_FORPARSING)), "Explore", SW_SHOWNORMAL);
	}
}

CancelWindowRegistered = function ()
{
	clearTimeout(g_.tidWindowRegistered);
	g_.bWindowRegistered = false;
	g_.tidWindowRegistered = setTimeout(function ()
	{
		g_.bWindowRegistered = true;
	}, 9999);
}

ShowDialogEx = function (mode, w, h, ele)
{
	ShowDialog(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "script\\dialog.html"), { MainWindow: MainWindow, Query: mode, width: w, height: h, element: ele});
}
ShowNew = function (Ctrl, pt, Mode)
{
	var FV = GetFolderView(Ctrl, pt);
	var path = api.GetDisplayNameOf(FV, SHGDN_FORPARSING);
	if (/^[A-Z]:\\|^\\\\/i.test(path)) {
		ShowDialog(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "script\\dialog.html"), { MainWindow: MainWindow, Query: "new", Mode: Mode, path: path, FV: FV, Modal: false, width: 480, height: 120});
	}
}

CreateNewFolder = function (Ctrl, pt)
{
	ShowNew(Ctrl, pt, "folder");
	return S_OK;
}

CreateNewFile = function (Ctrl, pt)
{
	ShowNew(Ctrl, pt, "file");
	return S_OK;
}

InputMouse = function (o)
{
	ShowDialogEx("mouse", 500, 420, o || document.F.MouseMouse || document.F.Mouse);
}

InputKey = function(o)
{
	ShowDialogEx("key", 320, 120, o || document.F.KeyKey || document.F.Key);
}

ShowIconEx = function (ele)
{
	if (!ele) {
		ele = document.F.Icon;
		ShowDialogEx("icon", 640, 480, ele);
	}
}

ShowLocationEx = function (s)
{
	ShowDialog(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "script\\location.html"), {MainWindow: MainWindow, Data: s});
}

function MakeKeySelect()
{
	var oa = document.getElementById("_KeyState");
	if (oa) {
		var ar = [];
		for (var i = 0; i < 4; i++) {
			var s = MainWindow.g_.KeyState[i][0];
			ar.push('<input type="checkbox" onclick="KeyShift(this)" id="_Key', s, '"><label for="_Key', s, '">', s, '&nbsp;</label>');
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
			var o = oa[++oa.length - 1];
			o.value = j;
			o.text = j + " ";
		}
	}
}

function SetKeyShift()
{
	var key = (document.F.elements.KeyKey || document.F.elements.Key).value;
	for (var i = 0; i < MainWindow.g_.KeyState.length; i++) {
		var s = MainWindow.g_.KeyState[i][0];
		var o = document.getElementById("_Key" + s);
		if (o) {
			o.checked = key.match(s + "+");
		}
		key = key.replace(s + "+", "");
	}
	o = document.getElementById("_KeySelect");
	for (var i = o.length; i--;) {
		if (api.StrCmpI(key, o[i].value) == 0) {
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

GetLangId = function (nDefault)
{
	if (!nDefault && te.Data.Conf_Lang) {
		return te.Data.Conf_Lang;
	}
	var lang = navigator.userLanguage.replace(/\-/, '_').toLowerCase();
	if (nDefault != 2) {
		if (!fso.FileExists(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "lang\\" + lang + ".xml"))) {
			lang = lang.replace(/_.*/,"");
		}
	}
	if (!te.Data.Conf_Lang) {
		te.Data.Conf_Lang = lang;
	}
	return lang;
}

GetSourceText = function (s)
{
	try {
		return (MainWindow.LangSrc || LangSrc)[s] || s;
	} catch (e) {
		return s;
	}
}

GetFolderView = function (Ctrl, pt, bStrict)
{
	if (!Ctrl) {
		return te.Ctrl(CTRL_FV);
	}
	if (!Ctrl.Type) {
		var o = Ctrl.offsetParent;
		while (o) {
			var res = /^Panel_(\d+)$/.exec(o.id);
			if (res) {
				return te.Ctrl(CTRL_TC, res[1]).Selected;
			}
			o = o.offsetParent
		}
		return te.Ctrl(CTRL_FV);
	}
	if (Ctrl.Type <= CTRL_EB) {
		return Ctrl;
	}
	if (Ctrl.Type == CTRL_TV) {
		return Ctrl.FolderView;
	}
	if (Ctrl.Type != CTRL_TC) {
		return te.Ctrl(CTRL_FV);
	}
	if (pt) {
		var FV = Ctrl.HitTest(pt);
		if (FV) {
			return FV;
		}
	}
	if (!bStrict || !pt) {
		return Ctrl.Selected;
	}
}

GetSelectedArray = function (Ctrl, pt, bPlus)
{
	var Selected, SelItem;
	var FV = null;
	var bSel = true;
	switch(Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
			FV = Ctrl;
			break;
		case CTRL_TC:
			FV = Ctrl.HitTest(pt);
			bSel = false;
			break;
		case CTRL_TV:
			FV = Ctrl.FolderView;
			SelItem = Ctrl.SelectedItem;
			break;
		case CTRL_WB:
			FV = te.Ctrl(CTRL_FV);
			SelItem = window.Input;
			break;
		default:
			FV = te.Ctrl(CTRL_FV);
			break;
	}
	if (FV && !SelItem) {
		if (bSel) {
	 		Selected = FV.SelectedItems();
		}
		if (Selected && Selected.Count) {
			SelItem = Selected.Item(0);
		} else {
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
	return isNaN(c) && /^#[0-9a-f]{3,6}$/i ? c : api.sprintf(8, "#%06x", ((c & 0xff) << 16) | (c & 0xff00) | ((c & 0xff0000) >> 16));
}

GetWinColor = function (c)
{
	var res = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(c);
	if (res) {
		return Number(["0x", res[3], res[2], res[1]].join(""));
	}
	res = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(c);
	if (res) {
		return Number(["0x", res[3], res[3], res[2], res[2], res[1], res[1]].join(""));
	}
	return c;
}

ChooseColor = function (c)
{
	var cc = api.Memory("CHOOSECOLOR");
	cc.lStructSize = cc.Size;
	cc.hwndOwner = api.GetForegroundWindow();
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
	if (isFinite(c)) {
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
				if (e[i].style) {
					e[i].style.cursor = s;
				}
			}
		}
	}
}

function MouseOver(o)
{
	if (api.PathMatchSpec(o.className, 'button;menu')) {
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
			} else if (objHover.className == 'hovermenu') {
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
	try {
		var locator = te.CreateObject("WbemScripting.SWbemLocator");
		var server = locator.ConnectServer(null, "root\\default");
		var reg = server.Get("StdRegProv");
		var Params = [hKey, Name, null];
		api.ExecMethod(reg, "EnumKey", Params);
		return new VBArray(Params[2]).toArray();
	} catch (e) {}
	return [];
}

FindText = function ()
{
	api.OleCmdExec(document, null, 32, 0, 0);
}

OpenDialogEx = function (path, filter, bFilesOnly)
{
	var commdlg = te.CommonDialog;
	var te_path = fso.GetParentFolderName(api.GetModuleFileName(null));
	var res = /^\.\.(\/.*)/.exec(path);
	if (res) {
		path = te_path + (res[1].replace(/\//g, "\\"));
	}
	path = api.PathUnquoteSpaces(ExtractMacro(te, path));
	if (!fso.FolderExists(path)) {
		path = fso.GetParentFolderName(path);
		if (!fso.FolderExists(path)) {
			path = fso.GetDriveName(te_path);
		}
	}
	commdlg.InitDir = path;
	var s = (api.LoadString(hShell32, 34193) || "All Files") + "|*.*";
	commdlg.Filter = filter ? ExtractMacro(te, filter).replace(/#/g, "|") + "|" + s : s;
	commdlg.Flags = OFN_FILEMUSTEXIST | OFN_EXPLORER | OFN_ENABLESIZING | (bFilesOnly ? 0 : OFN_ENABLEHOOK);
	if (commdlg.ShowOpen()) {
		return api.PathQuoteSpaces(commdlg.FileName);
	}
}

OpenDialog = function (path, bFilesOnly)
{
	return OpenDialogEx(path, null, bFilesOnly);
}

ChooseFolder = function (path, pt)
{
	if (!pt) {
		pt = api.Memory("POINT");
		api.GetCursorPos(pt);
	}
	var FolderItem = api.ILCreateFromPath(path);
	FolderItem = FolderMenu.Open(FolderItem.IsFolder ? FolderItem : ssfDRIVES, pt.x, pt.y);
	if (FolderItem) {
		return api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
	}
}

BrowseForFolder = function (path)
{
	return OpenDialogEx(path, api.LoadString(hShell32, 4131) + "|<Folder>");
}

InvokeCommand = function (Items, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon, FV, uCMF)
{
	if (Items) {
		var ContextMenu = api.ContextMenu(Items, FV);
		if (ContextMenu) {
			var hMenu = api.CreatePopupMenu();
			ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, uCMF);
			if (Verb === null) {
				Verb = api.GetMenuDefaultItem(hMenu, MF_BYCOMMAND, GMDI_USEDISABLED) - 1;
				if (Verb == -2) {
					api.DestroyMenu(hMenu);
					return S_FALSE;
				}
			}
			ContextMenu.InvokeCommand(fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon);
			api.DestroyMenu(hMenu);
		}
	}
	return S_OK;
}

SetRenameMenu = function (n)
{
	ExtraMenuCommand[CommandID_RENAME + n - 1] = function (Ctrl, pt, Name, nVerb)
	{
		if (Ctrl.Type == CTRL_TV) {
			setTimeout(function ()
			{
				Ctrl.Focus();
				wsh.SendKeys("{F2}");
			}, 99);
		} else {
			return S_FALSE;
		}
	};
}

ShowError = function (e, s, i)
{
	if (isFinite(i)) {
		if (eventTA[s][i]) {
			s = eventTA[s][i] + " : " + s;
		}
	}
	MessageBox([e.stack || e.description || e.toString(), s, GetTEInfo()].join("\n\n"), TITLE, MB_OK);
}

ApiStruct = function (oTypedef, nAli, oMemory)
{
	this.Size = 0;
	this.Typedef = oTypedef;
	for (var i in oTypedef) {
		var ar = oTypedef[i];
		var n = ar[1];
		this.Size += (n - (this.Size % n)) % n;
		ar[3] = this.Size;
		this.Size += n * (ar[2] || 1);
	}
	n = api.LowPart(nAli);
	this.Size += (n - (this.Size % n)) % n;
	this.Memory = /^object$/i.test(typeof oMemory) ? oMemory : api.Memory("BYTE", this.Size);
	this.Read = function (Id)
	{
		var ar = this.Typedef[Id];
		if (ar) {
			return this.Memory.Read(ar[3], ar[0]);
		}
	};
	this.Write = function (Id, Data)
	{
		var ar = this.Typedef[Id];
		if (ar) {
			this.Memory.Write(ar[3], ar[0], Data);
		}
	};
}

FindChildByClass = function (hwnd, s)
{
	var hwnd1, hwnd2;
	while (hwnd1 = api.FindWindowEx(hwnd, hwnd1, null, null)) {
		if (api.GetClassName(hwnd1) == s) {
			return hwnd1;
		}
		if (hwnd2 = FindChildByClass(hwnd1, s)) {
			return hwnd2;
		}
	}
	return null;
}

DownloadFile = function (url, fn)
{
	var hr = api.URLDownloadToFile(null, url, fn);
	if (hr && /^https:/.test(url)) {
		hr = api.URLDownloadToFile(null, url.replace(/^https/, "http"), fn);
	}
	return hr;
}

GetNavigateFlags = function (FV)
{
	if (!FV && OpenMode != SBSP_NEWBROWSER) {
		FV = te.Ctrl(CTRL_FV);
	}
	return api.GetKeyState(VK_MBUTTON) < 0 || api.GetKeyState(VK_CONTROL) < 0 || (FV && FV.Data.Lock) ? SBSP_NEWBROWSER : OpenMode;
}

AddEvent("ConfigChanged", function (s)
{
	te.Data["bSave" + s] = true;
});

GetSysColor = function (i)
{
	var c = g_.Colors[i];
	return c !== undefined ? c : api.GetSysColor(i);
}

SetSysColor = function (i, color)
{
	g_.Colors[i] = color;
}

ShellExecute = function (s, vOperation, nShow, vDir2, pt)
{
	var arg = api.CommandLineToArgv(s);
	var s = arg.shift();
	var vDir = fso.GetParentFolderName(s) || vDir2;
	if (pt && vDir.Type) {
		vDir = (GetFolderView(Ctrl, pt) || {FolderItem: {}}).FolderItem.Path;
	}
	for (var i = arg.length; i-- > 0;) {
		arg[i] = api.PathQuoteSpaces(arg[i]);
	}
	return sha.ShellExecute(s, arg.join(" "), vDir, vOperation, nShow);
}

CreateFont = function (LogFont)
{
	var key = [LogFont.lfFaceName, LogFont.lfHeight, LogFont.lfCharSet, LogFont.lfWeight, LogFont.lfItalic, LogFont.lfUnderline].join("\t");
	var hFont = te.Data.Fonts[key];
	if (!hFont) {
		hFont = api.CreateFontIndirect(LogFont);
		te.Data.Fonts[key] = hFont;
	}
	return hFont;
}

Activate = function (o, id)
{
	var TC = te.Ctrl(CTRL_TC);
	if (TC && TC.Id != id) {
		var FV = GetInnerFV(id);
		if (FV) {
			FV.Focus();
			if (o) {
				o.focus();
			}
		}
	}
}

function DetectProcessTag(e)
{
	return /input|textarea/i.test((e || event).srcElement.tagName);
}

AddEventEx(window, "load", function ()
{
	document.body.onselectstart = DetectProcessTag;
	document.body.oncontextmenu = DetectProcessTag;
	document.body.onmousewheel = function ()
	{
		return api.GetKeyState(VK_CONTROL) >= 0;
	};
});

Alt = function ()
{
	return S_OK;
}

GetSavePath = function (FolderItem)
{
	var path = api.GetDisplayNameOf(FolderItem, SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
	if (!/^[A-Z]:\\|^\\\\[A-Z]/i.test(path)) {
		var res = /search\-ms:.*?&crumb=location:([^&]*)/.exec(api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING));
		if (res) {
			return api.PathCreateFromUrl("file:" + res[1]);
		}
	}
	if (/\?/.test(path)) {
		var nCount = api.ILGetCount(FolderItem);
		path = [];
		while (nCount-- > 0) {
			path.unshift(api.GetDisplayNameOf(FolderItem, (nCount > 0 ? SHGDN_FORADDRESSBAR : 0) | SHGDN_FORPARSING | SHGDN_INFOLDER));
			FolderItem = api.ILRemoveLastID(FolderItem);
		}
		return path.join("\\")
	}
	return path;
}

ExecAddonScript = function (type, s, fn, arError, o, arStack)
{
	var sc = api.GetScriptDispatch(s, type, o,
		function (ei, SourceLineText, dwSourceContext, lLineNumber, CharacterPosition)
		{
			arError.push([api.SysAllocString(ei.bstrDescription), fn, api.sprintf(16, "Line: %d", lLineNumber)].join("\n"));
		}
	);
	if (sc && arStack) {
		arStack.push(sc);
	}
	return sc;
}

LoadAddon = function (ext, Id, arError, param)
{
	var r;
	try {
		var sc;
		var ar = ext.split(".");
		if (ar.length == 1) {
			ar.unshift("script");
		}
		var fn = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons") + "\\" + Id + "\\" + ar.join(".");
		var ado = OpenAdodbFromTextFile(fn);
		if (ado) {
			var s = ado.ReadText();
			ado.Close();
			if (s) {
				if (ar[1] == "js") {
					try {
						sc = new Function(s);
					} catch (e) {
						sc = ExecAddonScript("JScript", s, fn, arError);
					}
				} else if (ar[1] == "vbs") {
					sc = ExecAddonScript("VBScript", s, fn, arError, {"_Addon_Id": {"Addon_Id": Id}, window: window}, Addons["_stack"]);
				}
				if (sc) {
					r = sc(Id);
					if (param) {
						var res = /[\r\n\s]Default\s*=\s*["'](.*)["'];/.exec(s);
						if (res) {
							param.Default = res[1];
						}
					}
				}
			}
		}
	} catch (e) {
		arError.push([(e.description || e.toString()), fn].join("\n"));
	}
	return r;
}

AddEventEx(window, "beforeunload", function ()
{
	var hwnd = api.GetWindow(document);
	var hwnd1 = hwnd;
	while (hwnd1 = api.GetParent(hwnd)) {
		hwnd = hwnd1;
	}
	while (hwnd1 = api.FindWindowEx(null, hwnd1, null, null)) {
		if (hwnd == api.GetWindowLongPtr(hwnd1, GWLP_HWNDPARENT)) {
			api.PostMessage(hwnd1, WM_CLOSE, 0, 0);
		}
	}
});

GetTEInfo = function ()
{
	return api.sprintf(99, "TE%d %d.%d.%d Win %d.%d.%d%s %s %x%s IE %d %s", api.sizeof("HANDLE") * 8, (te.Version / 10000) % 100, (te.Version / 100) % 100, te.Version % 100, osInfo.dwMajorVersion, osInfo.dwMinorVersion, osInfo.dwBuildNumber, api.IsWow64Process(api.GetCurrentProcess()) ? " Wow64" : "", ["WS", "DC", "SV"][osInfo.wProductType - 1] || osInfo.wProductType, osInfo.wSuiteMask, api.SHTestTokenMembership(null, 0x220) ? " Admin" : "", document.documentMode || (document.body.style.maxHeight === undefined ? 6 : 7), GetLangId(2));
}

FireEvent = function (o, event)
{

	if (o) {
		if (o.fireEvent){
			return o.fireEvent('on' + event);
		} else if (document.createEvent) {
			var evt = document.createEvent("HTMLEvents");
			evt.initEvent(event, true, true);
			return !o.dispatchEvent(evt);
		}
	}
}

RemoveCommand = function (hMenu, ContextMenu, strDelete)
{
	var mii = api.Memory("MENUITEMINFO");
	mii.cbSize = mii.Size;
	mii.fMask = MIIM_ID;
	for (var i = api.GetMenuItemCount(hMenu); i-- > 0;) {
		if (api.GetMenuItemInfo(hMenu, i, true, mii)) {
			if (api.PathMatchSpec(ContextMenu.GetCommandString(mii.wID - ContextMenu.idCmdFirst, GCS_VERB), strDelete)) {
				api.DeleteMenu(hMenu, i, MF_BYPOSITION);
			}
		}
	}
}

DeleteTempFolder = function ()
{
	DeleteItem(fso.BuildPath(fso.GetSpecialFolder(2).Path, "tablacus"));
}

UpdateSelf = function ()
{
	var oExec = wsh.Exec(g_strUpdate);
	wsh.AppActivate(oExec.ProcessID);
}

OpenContains = function (Ctrl, pt)
{ 
	var FV = GetFolderView(Ctrl. pt);
	var Items = FV.SelectedItems();
	for (var j in Items) {
		var Item = Items.Item(j);
		var path = Item.Path;
		Navigate(fso.GetParentFolderName(path), SBSP_NEWBROWSER);
		(function (Item) { setTimeout(function ()
		{
			var FV = te.Ctrl(CTRL_FV);
			FV.SelectItem(Item, SVSI_SELECT | SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS);
		}, 99);}) (Item);
	}
}

OpenAdodbFromTextFile = function (fn)
{
	var ado = te.CreateObject(api.ADBSTRM);
	var charset = "_autodetect_all";
	try {
		ado.CharSet = "iso-8859-1";
		ado.Open();
		ado.LoadFromFile(fn);
		var s = ado.ReadText(999);
	} catch (e) {
		ado.Close();
		return;
	}
	if (/^\xEF\xBB\xBF/.test(s)) {
		charset = 'utf-8';
	} else if (/^\xFF\xFE|^\xFE\xFF/.test(s)) {
		charset = 'unicode';
	} else {
		var res = /<meta[^>]*charset\s*=([\w_\-]+)|\@charset.*?([\w_\-]+)|<\?xml[^>]*encoding\s*=[^\w_\->]*([\w_\-]+)/i.exec(s);
		if (res) {
			charset = res[1] || res[2] || res[3];
		}
	}
	ado.Position = 0;
	ado.CharSet = charset;
	return ado;
}

WmiProcess = function(arg, fn)
{
	try {
		var locator = te.CreateObject("WbemScripting.SWbemLocator");
		if (locator) {
			var server = locator.ConnectServer();
			if (server) {
				var cols = server.ExecQuery("SELECT * FROM Win32_Process " + arg);
				for (var list = new Enumerator(cols); !list.atEnd(); list.moveNext()) {
					fn(list.item());
				}
			}
		}
	} catch (e) {}
}

function CalcElementHeight(o, em)
{
	if (o) {
		if (document.documentMode >= 9) {
			o.style.height = "calc(100vh - " + em + "em)";
		} else {
			var h = document.documentElement.clientHeight || document.body.clientHeight;
			h += MainWindow.DefaultFont.lfHeight * em;
			if (h > 0) {
				o.style.height = h + 'px';
				o.style.height = 2 * h - o.offsetHeight + "px";
			}
		}
	}
}

AddFavoriteEx = function (Ctrl, pt)
{
	AddFavorite();
	return S_OK
}

importScript = function (fn)
{
	if (/"/.test(fn)) {
		fn = api.api.PathUnquoteSpaces(fn);
	}
	fn = ExtractMacro(te, fn);
	if (!api.PathMatchSpec(fn, '?:\\*;\\\\*')) {
		fn = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), fn);
	}
	var hr = E_FAIL;
	var ado = OpenAdodbFromTextFile(fn);
	if (ado) {
		ExecScriptEx(window.Ctrl, ado.ReadText(), /\.vbs/i.test(fn) ? "VBScript" : "JScript", window.pt, window.dataObj, window.grfKeyState, window.pdwEffect, window.bDrop);
		ado.Close();
	}
	return hr;
}
