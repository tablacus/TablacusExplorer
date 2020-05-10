//Tablacus Explorer

function AboutTE(n) {
	if (n == 0) {
		return te.Version < 20200510 ? te.Version : 20200510
	}
	if (n == 1) {
		var v = AboutTE(0);
		return api.sprintf(99, "%d.%d.%d", (v / 10000) % 100, (v / 100) % 100, v % 100);
	}
	if (n == 2) {
		return "Tablacus Explorer " + AboutTE(1) + " Gaku";
	}
	var ar = [g_.IEVer, GetLangId(2), screen.deviceYDPI];
	var server = te.GetObject("winmgmts:\\\\.\\root\\SecurityCenter" + (WINVER >= 0x600 ? "2" : ""));
	if (server) {
		var cols = server.ExecQuery("SELECT * FROM AntiVirusProduct");
		for (var list = new Enumerator(cols); !list.atEnd(); list.moveNext()) {
			ar.push(list.item().displayName);
		}
	}
	return api.sprintf(99, "TE%d %s Win %d.%d.%d%s %s %x%s%s IE", api.sizeof("HANDLE") * 8, AboutTE(1), osInfo.dwMajorVersion, osInfo.dwMinorVersion, osInfo.dwBuildNumber, api.IsWow64Process(api.GetCurrentProcess()) ? " Wow64" : "", ["WS", "DC", "SV"][osInfo.wProductType - 1] || osInfo.wProductType, osInfo.wSuiteMask, api.SHTestTokenMembership(null, 0x220) ? " Admin" : "", api.ShouldAppsUseDarkMode() ? " Dark" : "") + ar.join(" ");
}

Ctrl = null;
g_temp = null;
g_sep = "` ~";
Handled = null;
hwnd = null;
pt = api.Memory("POINT");
dataObj = null;
grfKeyState = null;
pdwEffect = [0];
bDrop = null;
Input = null;
eventTE = {
	Environment: {}
};
eventTA = {};
g_ptDrag = api.Memory("POINT");
Addons = {
	"_stack": []
};
g_ = {
	Colors: {},
	Panels: {},
	KeyCode: {},
	KeyState: [
		[0x1d0000, 0x2000],
		[0x2a0000, 0x1000],
		[0x380000, 0x4000],
		["Win", 0x8000],
		["Ctrl", 0x2000],
		["Shift", 0x1000],
		["Alt", 0x4000]
	],
	stack_TC: [],
	dlgs: {},
	tidWindowRegistered: null,
	bWindowRegistered: true,
	xmlWindow: null,
	elAddons: {},
	event: {},
	tid_rf: [],
	Autocomplete: {},
	LockUpdate: 0,
	IEVer: window.document && (document.documentMode || (/MSIE 6/.test(navigator.appVersion) ? 6 : 7))
};

FolderMenu =
{
	Items: [],
	SortMode: 0,
	SortReverse: false,

	Clear: function () {
		this.Items.length = 0;
		delete this.Filter;
	},

	Open: function (FolderItem, x, y, filter, nParent, hParent, wID) {
		this.Clear();
		this.Filter = filter;
		var hMenu = api.CreatePopupMenu();
		this.OpenMenu(hMenu, FolderItem, hParent, wID, nParent);
		window.g_menu_click = true;
		this.MenuLoop = true;
		var Verb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, x, y, te.hwnd, null, null);
		delete this.MenuLoop;
		g_popup = null;
		api.DestroyMenu(hMenu);
		Verb = Verb ? this.Items[Verb - 1] : null;
		this.Clear();
		return Verb;
	},

	OpenSubMenu: function (hMenu, wID, hSubMenu, nParent) {
		this.OpenMenu(api.sscanf(hSubMenu, "%llx"), this.Items[wID - 1], api.sscanf(hMenu, "%llx"), wID, nParent);
	},

	OpenMenu: function (hMenu, FolderItem, hParent, wID, nParent) {
		var Item;
		if (!FolderItem) {
			return;
		}
		if ("object" !== typeof FolderItem) {
			FolderItem = api.ILCreateFromPath(FolderItem);
		}
		var bSep = false;
		if (!nParent && !api.ILIsEmpty(FolderItem) && !api.ILIsParent(1, FolderItem, false)) {
			Item = api.ILRemoveLastID(FolderItem);
			var bMatch = IsFolderEx(Item);
			if (this.Filter) {
				bMatch = PathMatchEx(bMatch ? Item.Name + ".folder" : Item.Name, this.Filter);
			}
			if (bMatch) {
				this.AddMenuItem(hMenu, Item, "../", false, true);
				bSep = true;
			}
		}
		var Items = this.Enum(FolderItem);
		if (Items) {
			var nCount = Items.Count;
			var ar = new Array(nCount);
			for (var i = nCount; i--;) {
				ar[i] = i;
			}
			if (this.SortMode >= 0) {
				try {
					var d = fso.GetDrive(fso.GetDriveName(FolderItem.Path));
				} catch (e) {
					d = { FileSystem: "NTFS" };
				}
				if (!/NTFS/i.test(d.FileSystem) || this.SortMode || this.Filter || FolderItem.IsBrowsable) {
					this.Sort(Items, ar);
				}
			}
			if (this.SortReverse) {
				ar = ar.reverse();
			}
			for (var i = 0; i < nCount; ++i) {
				Item = Items.Item(ar[i]);
				var bMatch = IsFolderEx(Item) || api.ILIsParent(MainWindow.g_pidlCP, Item, false);
				if (this.Filter) {
					bMatch = PathMatchEx(bMatch ? Item.Name + ".folder" : Item.Name, this.Filter);
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
		RemoveSubMenu(hParent, wID);
		MainWindow.RunEvent1("FolderMenuCreated", hMenu, FolderItem, hParent);
	},

	Enum: function (FolderItem) {
		var Items = GetEnum(FolderItem, te.Data.Conf_MenuHidden);
		if (Items) {
			MainWindow.RunEvent1("AddItems", Items, FolderItem);
			if (!Items.Count) {
				Items.Item = function (i) {
					return api.ILCreateFromPath(Items[i]);
				}
				Items.Count = Items.length;
			}
		}
		return Items;
	},

	Sort: function (Items, ar) {
		ar.sort(function (a, b) {
			var r = api.CompareIDs(FolderMenu.SortMode, Items.Item(a), Items.Item(b));
			return r ? r > 32767 ? - 1 : 1 : 0;
		});
	},

	AddMenuItem: function (hMenu, FolderItem, Name, bSelect, bParent) {
		var mii = api.Memory("MENUITEMINFO");
		mii.cbSize = mii.Size;
		mii.fMask = MIIM_ID | MIIM_STRING | MIIM_BITMAP | MIIM_SUBMENU;
		if (bSelect && Name) {
			mii.dwTypeData = Name;
		} else {
			var s = (Name || "") + MainWindow.GetFolderItemName(FolderItem);
			if (!s) {
				return;
			}
			mii.dwTypeData = s;
		}
		AddMenuIconFolderItem(mii, FolderItem);
		this.Items.push(FolderItem);
		mii.wID = this.Items.length;
		var cc = this.Filter ? SFGAO_FOLDER : SFGAO_HASSUBFOLDER;
		if (!bSelect && api.GetAttributesOf(FolderItem, cc | SFGAO_BROWSABLE | SFGAO_LINK) == cc) {
			try {
				var o = fso.GetDrive(fso.GetDriveName(FolderItem.Path));
			} catch (e) {
				o = { IsReady: FolderItem && !FolderItem.Unavailable };
			}
			if (o.IsReady) {
				mii.hSubMenu = api.CreateMenu();
				api.InsertMenu(mii.hSubMenu, 0, MF_BYPOSITION | MF_STRING, 0, api.sprintf(99, '\tJScript\tFolderMenu.OpenSubMenu("%llx",%d,"%llx",%d)', hMenu, mii.wID, mii.hSubMenu, !bParent));
			}
		}
		MainWindow.RunEvent1("FolderMenuAddMenuItem", hMenu, mii, FolderItem, bSelect);
		api.InsertMenuItem(hMenu, MAXINT, false, mii);
	},

	Invoke: function (FolderItem, wFlags) {
		if (FolderItem) {
			var path = api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL);
			var bVirtual = api.ILIsParent(1, FolderItem, false) || FolderItem.Unavailable;
			if (window.g_menu_button == 4) {
				if (!bVirtual) {
					var pdwEffect = [DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK];
					api.SHDoDragDrop(null, FolderItem, te, pdwEffect[0], pdwEffect, true);
				}
				return;
			}
			if (window.g_menu_button == 2) {
				var pt = api.Memory("POINT");
				api.GetCursorPos(pt);
				if (bVirtual) {
					if (!confirmOk(path, TITLE, MB_OK | MB_ICONINFORMATION)) {
						return;
					}
				} else {
					var FV = te.Ctrl(CTRL_FV);
					var AltSelectedItems = FV.AltSelectedItems;
					var Items = api.CreateObject("FolderItems");
					Items.AddItem(FolderItem);
					FV.AltSelectedItems = Items;
					if (ExecMenu(FV, "Context", pt, 1) != S_OK) {
						PopupContextMenu(FolderItem);
					}
					FV.AltSelectedItems = AltSelectedItems;
					return;
				}
			}
			if (res = /^`(.*)`$/.exec(path)) {
				ShellExecute(res[1], null, SW_SHOWNORMAL);
				return;
			}
			if (res = /^javascript:(.*)$/i.exec(path)) {
				try {
					new Function(res[1])();
				} catch (e) {
					ShowError(e, res[1]);
				}
				return;
			}
			if (FolderItem.Enum || ((window.g_menu_button == 3 || isFinite(wFlags)) && (FolderItem.IsFolder || (!FolderItem.IsFileSystem && FolderItem.IsBrowsable)))) {
				Navigate(FolderItem, isFinite(wFlags) ? wFlags : GetOpenMode());
				return;
			}
			var FV = te.Ctrl(CTRL_FV);
			var AltSelectedItems = FV.AltSelectedItems;
			var Items = api.CreateObject("FolderItems");
			Items.AddItem(FolderItem);
			FV.AltSelectedItems = Items;
			if (ExecMenu(FV, "Default", null, 2) != S_OK) {
				ShellExecute(api.PathQuoteSpaces(api.GetDisplayNameOf(FolderItem, SHGDN_ORIGINAL | SHGDN_FORPARSING)), null, SW_SHOWNORMAL);
			}
			FV.AltSelectedItems = AltSelectedItems;
		}
	},

	Location: function (o) {
		this.Clear();
		var hMenu = api.CreatePopupMenu();
		RunEvent3("LocationPopup", hMenu);
		var pt = GetPos(o, true);
		window.g_menu_click = true;
		var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x + o.offsetWidth, pt.y + o.offsetHeight, te.hwnd, null, null);
		api.DestroyMenu(hMenu);
		var FolderItem = nVerb ? this.Items[nVerb - 1] : null;
		this.Clear();
		this.Invoke(FolderItem);
	}
};

RunEvent1 = function () {
	var args = Array.apply(null, arguments);
	var en = args.shift();
	var eo = eventTE[en.toLowerCase()];
	for (var i in eo) {
		try {
			eo[i].apply(eo[i], args);
		} catch (e) {
			ShowError(e, en, i);
		}
	}
}

RunEvent2 = function () {
	var args = Array.apply(null, arguments);
	var en = args.shift();
	var eo = eventTE[en.toLowerCase()];
	for (var i in eo) {
		try {
			var hr = eo[i].apply(null, args);
			if (isFinite(hr) && hr != S_OK) {
				return hr;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	return S_OK;
}

RunEvent3 = function () {
	var args = Array.apply(null, arguments);
	var en = args.shift();
	var eo = eventTE[en.toLowerCase()];
	for (var i in eo) {
		try {
			var hr = eo[i].apply(null, args);
			if (isFinite(hr)) {
				return hr;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
}

RunEvent4 = function () {
	var args = Array.apply(null, arguments);
	var en = args.shift();
	var eo = eventTE[en.toLowerCase()];
	for (var i in eo) {
		try {
			var r = eo[i].apply(null, args);
			if (r !== void 0) {
				return r;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
}

AddEvent = function (Name, fn, priority) {
	if (Name) {
		var en = Name.toLowerCase();
		var s = en.replace(/\d$/g, "");
		if (g_.event[s] && !te["On" + s]) {
			te["On" + s] = g_.event[s];
		}
		if (!eventTE[en]) {
			eventTE[en] = api.CreateObject("Array");
		}
		if (!eventTA[en]) {
			eventTA[en] = [];
		}
		if (priority) {
			eventTE[en].unshift(fn);
			eventTA[en].unshift(g_.Error_source);
		} else {
			eventTE[en].push(fn);
			eventTA[en].push(g_.Error_source);
		}
	}
}

AddEnv = function (Name, fn) {
	eventTE.Environment[Name.toLowerCase()] = fn;
}

AddEventId = function (Name, Id, fn) {
	var en = Name.toLowerCase();
	if (!eventTE[en]) {
		eventTE[en] = {};
	}
	eventTE[en][Id.toLowerCase()] = fn;
}

AddonDisabled = function (Id) {
	SaveConfig();
	RunEvent1("AddonDisabled", Id);
	if (eventTE.addondisabledex) {
		Id = Id.toLowerCase();
		if (Id == "tabgroups") {
			return;
		}
		var fn = eventTE.addondisabledex[Id];
		if (fn) {
			delete eventTE.addondisabledex[Id];
			AddEventEx(window, "beforeunload", fn);
		}
	}
	CollectGarbage();
}

function ApplyLangTag(o) {
	if (o) {
		for (i = o.length; i--;) {
			var s, s1;
			if (s = o[i].innerHTML) {
				if (s != (s1 = s.replace(/(\s*<[^>]*?>\s*)|([^<>]*)|/gm, function (strMatch, ref1, ref2) {
					var r = ref1 || ref2;
					return /^\s*</.test(r) ? r : amp2ul(GetTextR(r.replace(/&amp;/ig, "&")));
				}))) {
					o[i].innerHTML = s1;
				}
			}
			s = o[i].title;
			if (s) {
				o[i].title = GetTextR(s).replace(/\(&\w\)|&/, "");
			}
			s = o[i].alt;
			if (s) {
				o[i].alt = GetTextR(s).replace(/\(&\w\)|&/, "");
			}
		}
	}
}

function ApplyLang(doc) {
	var i, o, h = 0;
	var FaceName = MainWindow.DefaultFont.lfFaceName;
	if (doc.body) {
		doc.body.style.fontFamily = FaceName;
		doc.body.style.fontSize = Math.abs(MainWindow.DefaultFont.lfHeight) + "px";
		doc.body.style.fontWeight = MainWindow.DefaultFont.lfWeight;
	}
	ApplyLangTag(doc.getElementsByTagName("label"));
	ApplyLangTag(doc.getElementsByTagName("button"));
	ApplyLangTag(doc.getElementsByTagName("li"));
	o = doc.getElementsByTagName("a");
	if (o) {
		ApplyLangTag(o);
		for (i = o.length; i--;) {
			if (o[i].className == "treebutton" && o[i].innerHTML == "") {
				o[i].innerHTML = BUTTONS.opened;
			}
		}
	}
	o = doc.getElementsByTagName("input");
	if (o) {
		ApplyLangTag(o);
		for (i = o.length; i--;) {
			if (!h && o[i].type == "text") {
				h = o[i].offsetHeight;
			}
			o[i].placeholder = GetTextR(o[i].placeholder).replace(/\(&\w\)|&/, "");
			if (o[i].type == "button") {
				o[i].value = GetTextR(o[i].value).replace(/\(&\w\)|&/, "");
			}
			var s = ImgBase64(o[i], 0);
			if (s) {
				o[i].src = s;
				if (o[i].type == "text") {
					o[i].style.backgroundImage = "url('" + s + "')";
				}
			}
		}
	}
	o = doc.getElementsByTagName("img");
	if (o) {
		ApplyLangTag(o);
		for (i = o.length; i--;) {
			var s = ImgBase64(o[i], 0);
			if (s) {
				o[i].src = s;
			}
			if (!o[i].ondragstart) {
				o[i].draggable = false;
			}
		}
	}
	o = doc.getElementsByTagName("select");
	if (o) {
		for (i = o.length; i--;) {
			o[i].title = delamp(GetTextR(o[i].title));
			for (var j = 0; j < o[i].length; ++j) {
				o[i][j].text = GetTextR(o[i][j].text.replace(/^\n/, "").replace(/\n$/, ""));
			}
		}
	}
	o = doc.getElementsByTagName("textarea");
	if (o) {
		for (i = o.length; i--;) {
			o[i].onkeydown = InsertTab;
		}
	}
	o = doc.getElementsByTagName("form");
	if (o) {
		for (i = o.length; i--;) {
			o[i].onsubmit = function () { return false };
		}
	}
	doc.title = GetTextR(doc.title);
	setTimeout(function () {
		var hwnd = api.GetParent(api.GetWindow(doc));
		var s = api.GetWindowText(hwnd);
		if (/ \-+ .*$/.test(s)) {
			api.SetWindowText(hwnd, s.replace(/ \-+ .*$/, ""));
		}
	}, 500);
}

function amp2ul(s) {
	s = s.replace(/&amp;/ig, "&");
	if (/@.*\..*,\-?\d+/.test(s)) {
		try {
			var lk = wsh.CreateShortCut(".lnk");
			lk.Description = s;
			s = lk.Description;
		} catch (e) { }
	}
	return /;/.test(s) ? s : s.replace(/&(.)/ig, "<u>$1</u>");
}

function delamp(s) {
	s = s.replace(/&amp;/ig, "&");
	return /;/.test(s) ? s : s.replace(/&/ig, "");
}

function ImgBase64(o, index) {
	var src = ExtractMacro(te, o.src);
	var s = MakeImgSrc(src, index, false, o.height);
	if (!s && api.StrCmpI(o.src, src)) {
		return src.replace(location.href.replace(/[^\/]*$/, ""), "file:///");
	}
	return s;
}

function MakeImgDataEx(src, bSimple, h) {
	return bSimple || REGEXP_IMAGE.test(src) ? src : MakeImgSrc(src, 0, false, h);
}

function MakeImgSrc(src, index, bSrc, h) {
	var fn;
	src = api.PathUnquoteSpaces(ExtractMacro(te, src));
	if (!/^file:/i.test(src) && REGEXP_IMAGE.test(src)) {
		return src;
	}
	if (g_.IEVer < 8) {
		var res = /^bitmap:(.+)/i.exec(src);
		if (res) {
			fn = fso.BuildPath(te.Data.DataFolder, "cache\\bitmap\\" + res[1].replace(/[:\\\/]/g, "$") + ".png");
		} else {
			res = /^icon:(.+)/i.exec(src);
			if (res) {
				fn = fso.BuildPath(te.Data.DataFolder, "cache\\icon\\" + (res[1].replace(/[:\\\/]/g, "$")) + ".png");
			} else if (src && !REGEXP_IMAGE.test(src)) {
				fn = fso.BuildPath(te.Data.DataFolder, "cache\\file\\" + (api.PathCreateFromUrl(src).replace(/[:\\\/]/g, "$")) + ".png");
			}
		}
		if (fn && fso.FileExists(fn)) {
			return fn;
		}
	}
	var image = MakeImgData(src, index, h);
	if (image) {
		if (g_.IEVer >= 8) {
			return image.DataURI("image/png");
		}
		if (fn) {
			image.Save(fn);
			return fn;
		}
	}
	return bSrc ? src : "";
}

function MakeImgData(src, index, h) {
	var hIcon = MakeImgIcon(src, index, h);
	if (hIcon) {
		var image = api.CreateObject("WICBitmap").FromHICON(hIcon);
		api.DestroyIcon(hIcon);
		return image;
	}
}

function MakeImgIcon(src, index, h, bIcon) {
	var hIcon = null;
	var res = /^bitmap:(.+)/i.exec(src);
	if (res) {
		var icon = res[1].split(",");
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
	res = /^icon:(.+)/i.exec(src);
	if (res) {
		var icon = res[1].split(",");
		if (icon == "shell32.dll") {
			var dw = { 3: SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES, 4: SHGFI_SYSICONINDEX | SHGFI_OPENICON | SHGFI_USEFILEATTRIBUTES }[res[1]];
			if (dw) {
				var sfi = api.Memory("SHFILEINFO");
				api.SHGetFileInfo("*", FILE_ATTRIBUTE_DIRECTORY, sfi, sfi.Size, dw);
				return GetHICON(sfi.iIcon, h, ILD_NORMAL);
			}
		}
		var phIcon = api.Memory("HANDLE");
		if (icon[index * 4 + 2]) {
			h = icon[index * 4 + 2];
		} else if (!h) {
			h = api.GetSystemMetrics(SM_CYSMICON);
		}
		if (h > 16) {
			api.SHDefExtractIcon(icon[index * 4], icon[index * 4 + 1], 0, phIcon, null, h);
		} else {
			api.SHDefExtractIcon(icon[index * 4], icon[index * 4 + 1], 0, null, phIcon, h << 16);
		}
		if (phIcon[0]) {
			return phIcon[0];
		}
	}
	if (src && (bIcon || /\*/.test(src) || !REGEXP_IMAGE.test(src))) {
		var sfi = api.Memory("SHFILEINFO");
		if (/\*/.test(src)) {
			api.SHGetFileInfo(src, 0, sfi, sfi.Size, SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES);
		} else {
			if (/^file:/i.test(src)) {
				src = api.PathCreateFromUrl(src) || src;
			}
			var pidl = api.ILCreateFromPath(api.PathUnquoteSpaces(src));
			if (pidl) {
				api.SHGetFileInfo(pidl, 0, sfi, sfi.Size, SHGFI_SYSICONINDEX | SHGFI_PIDL);
			}
		}
		return GetHICON(sfi.iIcon, h, ILD_NORMAL);
	}
}

LoadImgDll = function (icon, index) {
	var hModule = api.LoadLibraryEx(fso.BuildPath(system32, icon[index * 4]), 0, LOAD_LIBRARY_AS_DATAFILE);
	if (!hModule && (icon[index * 4] || "").toLowerCase() == "ieframe.dll") {
		if (icon[index * 4 + 1] >= 500) {
			hModule = api.LoadLibraryEx(fso.BuildPath(system32, "browseui.dll"), 0, LOAD_LIBRARY_AS_DATAFILE);
		} else {
			hModule = api.LoadLibraryEx(fso.BuildPath(system32, "shell32.dll"), 0, LOAD_LIBRARY_AS_DATAFILE);
		}
	}
	return hModule;
}

GetText = function (id) {
	try {
		id = id.replace(/&amp;/g, "&");
		return MainWindow.Lang[id.toLowerCase()] || id;
	} catch (e) { }
	return id;
}

GetTextR = function (id) {
	var res = /^\@(.+),-(\d+)(\[[^\]]+\])?$/i.exec(id);
	if (res) {
		var hModule = api.LoadLibraryEx(res[1], 0, LOAD_LIBRARY_AS_DATAFILE);
		var s = api.LoadString(hModule, api.LowPart(res[2]));
		if (!s && res[3]) {
			var ar = res[3].substr(1, res[3].length - 2).split("|");
			for (var i = 0; i < ar.length && !s; ++i) {
				res = /^-(\d+)$/.exec(ar[i]);
				s = res ? api.LoadString(hModule, api.LowPart(res[1])) : GetTextR(ar[i]);
			}
		}
		if (hModule) {
			api.FreeLibrary(hModule);
		}
		if (s) {
			return s;
		}
	}
	res = /^({[0-9a-f\-]+} \d+)\|?(.*)$/i.exec(id);
	if (res) {
		var s = api.PSGetDisplayName(res[1]);
		return (s && s.indexOf("{") < 0) ? s : GetTextR(res[2]);
	}
	res = /^(\d+)\-bit$/i.exec(id);
	if (res) {
		return (api.LoadString(hShell32, 31092) || "%s (32-bit)").replace(/^%s\s\(|\)$/g, "").replace(/\d+/, res[1]);
	}
	return GetText(id) || "";
}

function LoadLang2(filename) {
	var xml = api.CreateObject("Msxml2.DOMDocument");
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
	for (var i = 0; i < items.length; ++i) {
		var item = items[i];
		SetLang2(item.getAttribute("s").replace("\\t", "\t").replace("\\n", "\n"), item.text.replace("\\t", "\t").replace("\\n", "\n"));
	}
}

SetLang2 = function (s, v) {
	var sl = s.toLowerCase();
	if (!MainWindow.Lang[sl] && !MainWindow.LangSrc[v]) {
		MainWindow.Lang[sl] = v;
		MainWindow.LangSrc[v] = s;
		if (/&/.test(s)) {
			SetLang2(s.replace(/\(&\w\)|&/, ""), v.replace(/\(&\w\)|&/, ""));
		}
		if (/\.\.\.$/.test(s)) {
			SetLang2(StripAmp(s), StripAmp(v));
		}
	}
}

LoadXml = function (filename, nGroup) {
	var items;
	var xml = filename;
	g_.fTCs = 0;
	if ("string" === typeof filename) {
		filename = api.PathUnquoteSpaces(ExtractMacro(te, filename));
		if (fso.FileExists(filename)) {
			xml = api.CreateObject("Msxml2.DOMDocument");
			xml.async = false;
			xml.load(filename);
		}
	}
	try {
		items = xml.getElementsByTagName('Ctrl');
	} catch (e) {
		return;
	}
	++g_.LockUpdate;
	te.LockUpdate();
	if (!nGroup) {
		var cTC = te.Ctrls(CTRL_TC);
		for (i in cTC) {
			cTC[i].Close();
		}
	}
	for (var i = 0; i < items.length; ++i) {
		var item = items[i];
		switch (item.getAttribute("Type") - 0) {
			case CTRL_TC:
				var TC = te.CreateCtrl(CTRL_TC, item.getAttribute("Left"), item.getAttribute("Top"), item.getAttribute("Width"), item.getAttribute("Height"), item.getAttribute("Style"), item.getAttribute("Align"), item.getAttribute("TabWidth"), item.getAttribute("TabHeight"));
				TC.Data.Group = nGroup || Number(item.getAttribute("Group")) || 0;
				var tabs = item.getElementsByTagName('Ctrl');
				for (var i2 = 0; i2 < tabs.length; ++i2) {
					var tab = tabs[i2];
					var Path = tab.getAttribute("Path");
					var logs = tab.getElementsByTagName('Log');
					var nLogCount = logs.length;
					if (nLogCount > 1) {
						Path = api.CreateObject("FolderItems");
						for (var i3 = 0; i3 < nLogCount; ++i3) {
							Path.AddItem(logs[i3].getAttribute("Path"));
						}
						Path.Index = tab.getAttribute("LogIndex");
					}
					var FV = TC.Selected.Navigate2(Path, SBSP_NEWBROWSER, tab.getAttribute("Type"), tab.getAttribute("ViewMode"), tab.getAttribute("FolderFlags"), tab.getAttribute("Options"), tab.getAttribute("ViewFlags"), tab.getAttribute("IconSize"), tab.getAttribute("Align"), tab.getAttribute("Width"), tab.getAttribute("Flags"), tab.getAttribute("EnumFlags"), tab.getAttribute("RootStyle"), tab.getAttribute("Root"));
					if (!FV.FilterView) {
						FV.FilterView = tab.getAttribute("FilterView");
						FV.Data.Lock = api.LowPart(tab.getAttribute("Lock")) != 0;
						Lock(TC, i2, false);
					}
					ChangeTabName(FV);
					MainWindow.RunEvent1("LoadFV", FV, tab);
				}
				TC.SelectedIndex = item.getAttribute("SelectedIndex");
				TC.Visible = api.LowPart(item.getAttribute("Visible"));
				if (TC.Visible) {
					g_.focused = TC.Selected;
					++g_.fTCs;
				}
				MainWindow.RunEvent1("LoadTC", TC, item);
				break;
		}
	}
	if (!nGroup) {
		MainWindow.RunEvent1("LoadWindow", xml);
	}
	te.UnlockUpdate();
	--g_.LockUpdate;
}

SaveXmlTC = function (Ctrl, xml, nGroup) {
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
							--nLogIndex;
						}
					}
					if (bLogSaved) {
						item2.setAttribute("LogIndex", nLogIndex);
					}
				}
			}
			MainWindow.RunEvent1("SaveFV", FV, item2);
			item.appendChild(item2);
			bEmpty = false;
		}
	}
	MainWindow.RunEvent1("SaveTC", Ctrl, item);
	xml.documentElement.appendChild(item);
}

SaveConfigXML = function (filename) {
	var xml = CreateXml(true);
	var root = xml.documentElement;

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

	MainWindow.RunEvent1("SaveConfig", xml);
	try {
		xml.save(api.PathUnquoteSpaces(filename));
	} catch (e) {
		if (e.number != E_ACCESSDENIED) {
			ShowError(e, [GetText("Save"), filename].join(": "));
		}
	}
}

SaveXml = function (filename) {
	var xml = CreateXml(true);
	var root = xml.documentElement;
	var item = xml.createElement("Window");
	if (!api.IsZoomed(te.hwnd) && !api.IsIconic(te.hwnd)) {
		api.GetWindowRect(te.hwnd, te.Data.rcWindow);
	}
	item.setAttribute("Left", te.Data.rcWindow.left);
	item.setAttribute("Top", te.Data.rcWindow.top);
	item.setAttribute("Width", te.Data.rcWindow.right - te.Data.rcWindow.left);
	item.setAttribute("Height", te.Data.rcWindow.bottom - te.Data.rcWindow.top);
	item.setAttribute("CmdShow", api.IsZoomed(te.hwnd) ? SW_SHOWMAXIMIZED : te.CmdShow);
	item.setAttribute("DPI", screen.deviceYDPI);
	root.appendChild(item);

	var TC = te.Ctrl(CTRL_TC);
	var cTC = te.Ctrls(CTRL_TC);
	for (var i in cTC) {
		if (cTC[i].Id != TC.Id) {
			SaveXmlTC(cTC[i], xml);
		}
	}
	SaveXmlTC(TC, xml);

	MainWindow.RunEvent1("SaveWindow", xml);
	try {
		xml.save(api.PathUnquoteSpaces(ExtractMacro(te, filename)));
	} catch (e) {
		if (e.number != E_ACCESSDENIED) {
			ShowError(e, [GetText("Save"), filename].join(": "));
		}
	}
}

GetKeyKey = function (strKey) {
	var nShift = api.sscanf(strKey, "$%x");
	if (nShift) {
		return nShift;
	}
	strKey = strKey.toUpperCase();
	for (var j = 0; j < MainWindow.g_.KeyState.length; ++j) {
		var s = MainWindow.g_.KeyState[j][0].toUpperCase() + "+";
		var i = strKey.indexOf(s);
		if (i >= 0) {
			strKey = strKey.substr(0, i) + strKey.substr(i + s.length);
			nShift |= MainWindow.g_.KeyState[j][1];
		}
	}
	return nShift | MainWindow.g_.KeyCode[strKey];
}

GetKeyName = function (strKey, bEn) {
	var nKey = api.sscanf(strKey, "$%x");
	if (nKey) {
		var s = api.GetKeyNameText((nKey & 0x17f) << 16);
		if (s) {
			var arKey = [];
			for (var i = 0, z = MainWindow.g_.KeyState.length; i < z; ++i) {
				var j = bEn ? (i + 3) % z : i;
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

GetKeyShift = function () {
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

function SetKeyData(mode, strKey, path, type, km, o) {
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

function SendShortcutKeyFV(Key) {
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

CreateTab = function (Ctrl, pt) {
	var FV = GetFolderView(Ctrl, pt);
	NavigateFV(FV, HOME_PATH || api.ILIsEqual(FV, "about:blank") ? HOME_PATH : FV, SBSP_NEWBROWSER);
}

Navigate = function (Path, wFlags) {
	NavigateFV(te.Ctrl(CTRL_FV), Path, wFlags);
}

NavigateFV = function (FV, Path, wFlags, bInputed) {
	if (!FV) {
		var TC = te.CreateCtrl(CTRL_TC, 0, 0, "100%", "100%", te.Data.Tab_Style, te.Data.Tab_Align, te.Data.Tab_TabWidth, te.Data.Tab_TabHeight);
		FV = TC.Selected;
	}
	if ("string" === typeof Path) {
		Path = ExtractMacro(FV, Path).replace(/^\s+|\s*$/g, "");
		if (/\?|\*/.test(Path)) {
			if (!/\\\\\?\\|:/.test(Path)) {
				FV.FilterView = Path;
				FV.Refresh();
				return;
			}
		}
		if (/^file:/.test(Path)) {
			Path = decodeURI(Path);
		}
	}
	if (GetLock(FV)) {
		wFlags |= SBSP_NEWBROWSER;
	}
	if (bInputed) {
		RunEvent1("LocationEntered", FV, Path, wFlags);
		var r = MainWindow.RunEvent4("LocationEntered2", FV, Path, wFlags);
		if (r !== void 0) {
			return r;
		}
	}
	FV.Navigate(Path, wFlags);
	FV.Focus();
}

GetOpenMode = function () {
	return window.g_menu_button == 3 ? SBSP_NEWBROWSER : GetNavigateFlags();
}

IsDrag = function (pt1, pt2) {
	if (pt1 && pt2) {
		try {
			return (Math.abs(pt1.x - pt2.x) > api.GetSystemMetrics(SM_CXDRAG) || Math.abs(pt1.y - pt2.y) > api.GetSystemMetrics(SM_CYDRAG));
		} catch (e) { }
	}
	return false;
}

ChangeTab = function (TC, nMove) {
	var nCount = TC.Count;
	TC.SelectedIndex = (TC.SelectedIndex + nCount + nMove) % nCount;
}

ShowOptions = function (s) {
	try {
		var dlg = g_.dlgs.Options;
		if (dlg) {
			dlg.Window.SetTab(s);
			dlg.Focus();
			return;
		}
	} catch (e) { }
	g_.dlgs.Options = ShowDialog("options.html",
		{
			width: te.Data.Conf_OptWidth, height: te.Data.Conf_OptHeight,
			Data: s, event:
			{
				onbeforeunload: function () {
					delete MainWindow.g_.dlgs.Options;
				}
			}
		})
}

ShowDialog = function (fn, opt) {
	opt.opener = window;
	if (!/:/.test(fn)) {
		fn = location.href.replace(/[^\/]*$/, fn);
	}
	var r = opt.r || Math.abs(MainWindow.DefaultFont.lfHeight) / 12;
	var h = api.GetWindowLongPtr(te.hwnd, GWL_STYLE) & WS_CAPTION ? 0 : api.GetSystemMetrics(SM_CYCAPTION);
	return te.CreateCtrl(CTRL_SW, fn, opt, document, (opt.width > 99 ? opt.width : 750) * r, (opt.height > 99 ? opt.height : 530) * r + h, opt.left, opt.top);
}

LoadLayout = function () {
	var commdlg = api.CreateObject("CommonDialog");
	commdlg.InitDir = fso.BuildPath(te.Data.DataFolder, "layout");
	commdlg.Filter = MakeCommDlgFilter("*.xml");
	commdlg.Flags = OFN_FILEMUSTEXIST;
	if (commdlg.ShowOpen()) {
		LoadXml(commdlg.FileName);
	}
	return S_OK;
}

SaveLayout = function () {
	var commdlg = api.CreateObject("CommonDialog");
	commdlg.InitDir = fso.BuildPath(te.Data.DataFolder, "layout");
	commdlg.Filter = MakeCommDlgFilter("*.xml");
	commdlg.DefExt = "xml";
	commdlg.Flags = OFN_OVERWRITEPROMPT;
	if (commdlg.ShowSave()) {
		SaveXml(commdlg.FileName);
	}
	return S_OK;
}

ReloadCustomize = function () {
	te.Data.bReload = false;
	CloseSubWindows();
	Finalize();
	te.Reload();
	return S_OK;
}

GetPos = function (o, bScreen, bAbs, bPanel, bBottom) {
	if ("number" === typeof bScreen) {
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
		if (bAbs || !bPanel || !/absolute/i.test(o.style.position)) {
			x += o.offsetLeft - (bAbs ? 0 : o.scrollLeft);
			y += o.offsetTop - (bAbs ? 0 : o.scrollTop);
			o = o.offsetParent;
		} else {
			break;
		}
	}
	var pt = api.Memory("POINT");
	pt.x = x;
	pt.y = y;
	return pt;
}

HitTest = function (o, pt) {
	if (o) {
		var p = GetPos(o, 1);
		if (pt.x >= p.x && pt.x < p.x + o.offsetWidth && pt.y >= p.y && pt.y < p.y + o.offsetHeight) {
			o = o.offsetParent;
			p = GetPos(o, 1);
			return pt.x >= p.x && pt.x < p.x + o.offsetWidth && pt.y >= p.y && pt.y < p.y + o.offsetHeight;
		}
	}
	return false;
}

PtInRect = function (rc, pt) {
	return pt.x >= rc.left && pt.x < rc.right && pt.y >= rc.top && pt.y < rc.bottom;
}

DeleteItem = function (path, fFlags) {
	if (IsExists(path)) {
		api.SHFileOperation(FO_DELETE, path, null, fFlags || FOF_SILENT | FOF_NOCONFIRMATION | FOF_NOERRORUI, false);
	}
}

IsExists = function (path) {
	var wfd = api.Memory("WIN32_FIND_DATA");
	var hFind = api.FindFirstFile(path.replace(/\\$/, ""), wfd);
	api.FindClose(hFind);
	return hFind != INVALID_HANDLE_VALUE;
}

CreateNew = function (path, fn) {
	if (fn && !IsExists(path)) {
		try {
			fn(path);
		} catch (e) {
			if (/^[A-Z]:\\|^\\\\\w/i.test(path)) {
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
					for (var i = 1; i < ar.length; ++i) {
						fso.CreateFolder(path4);
						path4 = fso.BuildPath(path4, ar[i]);
					}
					fn(path4);
					api.SHFileOperation(FO_MOVE, path3, fso.GetParentFolderName(path), FOF_SILENT | FOF_NOCONFIRMATION | FOF_NOERRORUI, false);
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

SetFileTime = function (path, ctime, atime, mtime) {
	var b = MainWindow.RunEvent3("SetFileTime", path, ctime, atime, mtime);
	if (isFinite(b)) {
		return b;
	}
	return api.SetFileTime(path, ctime, atime, mtime);
}

SetFileAttributes = function (path, attr) {
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

CreateFolder = function (path) {
	var r = MainWindow.RunEvent4("CreateFolder", path);
	if (r !== void 0) {
		return r;
	}
	CreateNew(path, function (strPath) {
		fso.CreateFolder(strPath.replace(/^\s*/, ""));
	});
}

CreateFile = function (path) {
	var r = MainWindow.RunEvent4("CreateFile", path);
	if (r !== void 0) {
		return r;
	}
	CreateNew(path, CreateFile2);
}

CreateFolder2 = function (path) {
	if (!fso.FolderExists(path)) {
		CreateFolder(path);
	}
}

CreateFile2 = function (path) {
	var ext = fso.GetExtensionName(path);
	if (ext) {
		var s, r = "HKCR\\." + ext + "\\";
		try {
			s = wsh.regRead(r);
			try {
				wsh.RegRead(r + "ShellNew\\");
			} catch (e) {
				r += s + "\\";
				wsh.RegRead(r + "\\ShellNew\\");
			}
			r += "ShellNew\\";
			var ar = ['Command', 'Data', 'FileName'];
			for (var i in ar) {
				try {
					s = wsh.RegRead(r + ar[i]);
				} catch (e) {
					continue;
				}
				if (s) {
					if (i == 2) {
						r = fso.BuildPath(wsh.SpecialFolders("Templates"), s);
						if (!fso.FileExists(r)) {
							r = wsh.ExpandEnvironmentStrings("%SystemRoot%\\ShellNew\\") + s;
						}
						fso.CopyFile(r, path);
						SetFileTime(path, null, null, new Date());
						return;
					}
					if (i == 1) {
						var a = fso.CreateTextFile(path);
						a.Write(s);
						a.Close();
						return;
					}
					ShellExecute(s.replace("%1", path), null, SW_SHOWNORMAL);
					return;
				}
			}
		} catch (e) { }
	}
	fso.CreateTextFile(path).Close();
}

FormatDateTime = function (s) {
	return new Date(s).getTime() > 0 ? (api.GetDateFormat(LOCALE_USER_DEFAULT, 0, s, api.GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE)) + " " + api.GetTimeFormat(LOCALE_USER_DEFAULT, 0, s, api.GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STIMEFORMAT))) : "";
};

GetConsts = function (s) {
	var Result = window[s.replace(/\s/, "")];
	if (Result !== void 0) {
		return Result;
	}
	return s;
}

Navigate2 = function (path, NewTab) {
	var a = path.toString().split("\n");
	for (var i in a) {
		var s = a[i].replace(/^\s+/, "");
		if (s != "") {
			Navigate(s, NewTab);
			NewTab |= SBSP_NEWBROWSER;
		}
	}
}

ExecOpen = function (Ctrl, s, type, hwnd, pt, NewTab) {
	var nLock = 0;
	var line = s.split("\n");
	var bLast = (Addons.TabPositon || {}).nNew;
	var bRev = (NewTab & SBSP_ACTIVATE_NOFOCUS) && !bLast;
	var bFirst = false;
	var FV = GetFolderView(Ctrl, pt);
	while (line.length) {
		if (line.length > 1 && api.ILIsEqual(FV, "about:blank")) {
			if (NewTab & SBSP_ACTIVATE_NOFOCUS) {
				NewTab &= ~SBSP_ACTIVATE_NOFOCUS;
				bFirst = true;
			}
			++g_.LockUpdate;
			++nLock;
			bRev = false;
		}
		var s = bRev ? line.pop() : line.shift();
		if (s) {
			NavigateFV(FV, ExtractPath(Ctrl, s, pt), NewTab);
			NewTab |= SBSP_NEWBROWSER;
		}
	}
	if (bFirst) {
		FV.Parent.SelectedIndex = 0;
	}
	g_.LockUpdate -= nLock;
	return S_OK;
}

DropOpen = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop) {
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

Exec = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop) {
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
	var hr = MainWindow.RunEvent3("Exec", Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop, window.FV);
	return isFinite(hr) ? hr : window.Handled;
}

ExecScriptEx = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop, FV) {
	var fn = null;
	try {
		if (/J.*Script/i.test(type)) {
			fn = { Handled: new Function(s) };
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

DropScript = function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop, FV) {
	if (!pdwEffect) {
		pdwEffect = api.Memory("DWORD");
	}
	if (/EnableDragDrop/.test(s)) {
		return ExecScriptEx(Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop, FV);
	}
	pdwEffect[0] = DROPEFFECT_NONE;
	return E_NOTIMPL;
}

ExtractPath = function (Ctrl, s, pt) {
	s = api.PathUnquoteSpaces(ExtractMacro(Ctrl, GetConsts(s)));
	if (/^\.|^\\$/.test(s)) {
		var FV = GetFolderView(Ctrl, pt);
		if (FV) {
			if (s == "\\") {
				return fso.GetDriveName(FV.FolderItem.Path) + s;
			}
			if (s == "..") {
				return api.GetDisplayNameOf(api.ILGetParent(FV), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL);
			}
			var res = /\.\.\\(.*)/.exec(s);
			if (res) {
				return fso.BuildPath(api.GetDisplayNameOf(api.ILGetParent(FV), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL), res[1]);
			}
			res = /\.\\(.*)/.exec(s);
			if (res) {
				return fso.BuildPath(api.GetDisplayNameOf(FV, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL), res[1]);
			}
		}
	}
	return s;
}

ExtractMacro = MainWindow.ExtractMacro || function (Ctrl, s) {
	if ("string" ===  typeof s) {
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

PathMatchEx = function (path, s) {
	var hr = MainWindow.RunEvent3("PathMatch", path, s);
	if (isFinite(hr)) {
		return hr;
	}
	var res = /^\/(.*)\/(.*)/.exec(s);
	return res ? new RegExp(res[1], res[2]).test(path) : api.PathMatchSpec(path, s);
}

IsFolderEx = function (Item) {
	if (Item) {
		if (Item.IsFolder && !api.ILIsParent(ssfBITBUCKET, Item, true)) {
			var wfd = api.Memory("WIN32_FIND_DATA");
			var hr = api.SHGetDataFromIDList(Item, SHGDFIL_FINDDATA, wfd, wfd.Size);
			return (hr < 0) || (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 || !/^[A-Z]:\\|^\\\\\w.*\\.*\\/i.test(Item.Path);
		}
		return !Item.IsFileSystem && Item.IsBrowsable;
	}
	return false;
}

OpenMenu = function (items, SelItem) {
	var arMenu;
	var path = "";
	if (SelItem) {
		if ("object" === typeof SelItem) {
			var link = SelItem.ExtendedProperty("linktarget");
			path = String(api.GetDisplayNameOf(SelItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX));
			if (link && /\.lnk$/i.test(path)) {
				for (var i = items.length; --i >= 0;) {
					var s = items[i].getAttribute("Filter");
					if (/\.lnk/i.test(s) && PathMatchEx(path, s)) {
						link = path;
						break;
					}
				}
				path = link;
			}
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
	for (var i = 0; i < items.length; ++i) {
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

ExecMenu3 = function (Ctrl, Name, x, y) {
	window.Ctrl = Ctrl;
	setTimeout(function () {
		ExecMenu2(Name, x, y);
	}, 99);;
}

ExecMenu2 = function (Name, x, y) {
	if (!pt) {
		pt = api.Memory("POINT");
	}
	pt.x = x;
	pt.y = y;
	ExecMenu(Ctrl, Name, pt, 0);
}

AdjustMenuBreak = function (hMenu) {
	var mii = api.Memory("MENUITEMINFO");
	mii.cbSize = mii.Size;
	var uFlags = 0;
	for (var i = api.GetMenuItemCount(hMenu); i-- > 0;) {
		mii.fMask = MIIM_FTYPE | MIIM_SUBMENU;
		api.GetMenuItemInfo(hMenu, i, true, mii);
		if (api.GetMenuString(hMenu, i, MF_BYPOSITION) != "") {
			mii.fType |= uFlags;
			api.SetMenuItemInfo(hMenu, i, true, mii);
			uFlags = 0;
			if (mii.hSubMenu) {
				AdjustMenuBreak(mii.hSubMenu);
			}
			continue;
		}
		if (mii.hSubMenu) {
			AdjustMenuBreak(mii.hSubMenu);
			continue;
		}
		var u = mii.fType & (MFT_MENUBREAK | MFT_MENUBARBREAK);
		if (u && api.DeleteMenu(hMenu, i, MF_BYPOSITION)) {
			++i;
			uFlags = u;
		} else {
			uFlags = 0;
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

teMenuGetElementsByTagName = function (Name) {
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

ExecMenu = function (Ctrl, Name, pt, Mode, bNoExec) {
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
	eventTE.menucommand = [];
	var arMenu, item;
	if (items) {
		arMenu = OpenMenu(items, SelItem);
		for (var i = 0; i < arMenu.length; ++i) {
			item = items[arMenu[i]];
			if (!/^menus$/i.test(item.getAttribute("Type"))) {
				break;
			}
			item = null;
		}
		var nBase = api.LowPart(menus[0].getAttribute("Base"));
		if (nBase == 1) {
			if (api.LowPart(menus[0].getAttribute("Pos")) < 0) {
				for (var i = arMenu.length; i-- > 0;) {
					item = items[arMenu[i]];
					if (!/^menus$/i.test(item.getAttribute("Type"))) {
						break;
					}
					item = null;
				}
				if (arMenu.length > 1) {
					for (var i = arMenu.length; i--;) {
						var nLevel = 0;
						if (/^menus$/i.test(items[arMenu[i]].getAttribute("Type"))) {
							var s = String(items[arMenu[i]].text).toLowerCase();
							if (s == "close") {
								++nLevel;
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
			g_nPos = MakeMenus(hMenu, menus, arMenu, items, Ctrl, pt, 0, null, true);
			var eo = eventTE[Name.toLowerCase()];
			for (var i in eo) {
				try {
					g_nPos = eo[i](Ctrl, hMenu, g_nPos, Selected, SelItem, ContextMenu, Name, pt);
				} catch (e) {
					ShowError(e, Name, i);
				}
			}
			for (var i in eventTE.menus) {
				try {
					g_nPos = eventTE.menus[i](Ctrl, hMenu, g_nPos, Selected, SelItem, ContextMenu, Name, pt);
				} catch (e) {
					ShowError(e, "Menus", i);
				}
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
							pt.x = rc.left;
							pt.y = rc.top;
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
			if (window.g_menu_button == 2 && api.PathMatchSpec(s, "Open;Open in new tab;Open in background")) {
				PopupContextMenu(item.text);
				return S_OK;
			}
			Exec(Ctrl, item.text, window.g_menu_button == 3 && s == "Open" ? "Open in new tab" : s, Ctrl.hwnd, pt);
			return S_OK;
		}
		if (Mode != 2) {
			return (nBase != 2 || ContextMenu) ? S_OK : S_FALSE;
		}
	}
	return S_FALSE;
}

ExecMenu4 = function (Ctrl, Name, pt, hMenu, arContextMenu, nVerb, FV) {
	if (ExtraMenuCommand[nVerb]) {
		if (ExtraMenuCommand[nVerb](Ctrl, pt, Name, nVerb) != S_FALSE) {
			api.DestroyMenu(hMenu);
			return S_OK;
		}
	}
	for (var i in eventTE.menucommand) {
		var hr = eventTE.menucommand[i](Ctrl, pt, Name, nVerb, hMenu);
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

CopyMenu = function (hSrc, hDest) {
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

GetViewMenu = function (arContextMenu, FV, hMenu, uCMF) {
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

GetBaseMenuEx = function (hMenu, nBase, FV, Selected, uCMF, Mode, SelItem, arContextMenu) {
	var ContextMenu;
	for (var i in eventTE.getbasemenuex) {
		ContextMenu = eventTE.getbasemenuex[i](hMenu, nBase, FV, Selected, uCMF, Mode, SelItem, arContextMenu);
		if (ContextMenu !== void 0) {
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
					ContextMenu.QueryContextMenu(hMenu, 0, 0x6001, 0x6fff, uCMF | CMF_ITEMMENU);
					if (!Items.Count) {
						SetRenameMenu(ContextMenu.idCmdFirst);
					}
				}
			} else if (FV) {
				ContextMenu = GetViewMenu(arContextMenu, FV, hMenu, uCMF);
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
				ContextMenu = GetViewMenu(arContextMenu, FV, hMenu, uCMF);
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
			var dir = [GetText("Check for updates"), GetText("Get Add-ons"), null, api.sprintf(99, GetText("&About %s"), "Tablacus Explorer")];
			for (var i = 0; i < dir.length; ++i) {
				var s = dir[i];
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | (s === null ? MF_SEPARATOR : MF_STRING), i + 0x4011, s);
			}
			AddEvent("MenuCommand", function (Ctrl, pt, Name, nVerb) {
				var s = [CheckUpdate, GetAddons, null, ShowAbout][nVerb - 0x4011];
				if (s) {
					s(Ctrl, pt, Name, nVerb);
					return S_OK;
				}
			});
			break;
		case 8:
			api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 0x4001, GetText("&Add to Favorites..."));
			ExtraMenuCommand[0x4001] = AddFavoriteEx;
			api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, 0x4002, GetText("&Edit"));
			ExtraMenuCommand[0x4002] = function () {
				ShowOptions("Tab=Menus&Menus=Favorites");
			};
			api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_SEPARATOR, 0, null);
			break;
		default:
			break;
	}
	return ContextMenu;
}

MenuDbInit = function (hMenu, oMenu, oMenu2) {
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

MenuDbReplace = function (hMenu, oMenu, hMenu2) {
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

GetAccelerator = function (s) {
	var res = /&(.)/.exec(s);
	return res ? res[1] : "";
}

GetNetworkIcon = function (path) {
	if (api.PathIsNetworkPath(path)) {
		if (/^\\\\[^\\]+$/.test(path)) {
			return "icon:shell32.dll,15";
		}
		if (fso.GetDriveName(path) == path.replace(/\\$/, "")) {
			if (/^\\\\/.test(path)) {
				return WINVER >= 0x600 ? "icon:shell32.dll,275" : "icon:shell32.dll,85";
			}
			return WINVER >= 0x600 ? "icon:shell32.dll,273" : "icon:shell32.dll,9";
		}
		return "folder:closed";
	}
}

RemoveSubMenu = function (hMenu, wID) {
	if (hMenu && wID) {
		var mii = api.Memory("MENUITEMINFO");
		mii.cbSize = mii.Size;
		mii.fMask = MIIM_SUBMENU | MIIM_FTYPE;
		if (api.GetMenuItemInfo(hMenu, wID, false, mii)) {
			api.DestroyMenu(mii.hSubMenu);
			mii.hSubMenu = 0;
			mii.fType &= ~MF_POPUP;
			api.SetMenuItemInfo(hMenu, wID, false, mii);
		}
	}
}

AddMenuIconFolderItem = function (mii, FolderItem, nHeight) {
	var path = MainWindow.GetIconImage(FolderItem, GetSysColor(COLOR_WINDOW), 2);
	MenusIcon(mii, path || api.GetDisplayNameOf(FolderItem, SHGDN_FORPARSING | SHGDN_ORIGINAL), nHeight, !path);
}

AddMenuImage = function (mii, image, id, nHeight) {
	mii.hbmpItem = image.GetHBITMAP(WINVER >= 0x600 ? -2 : GetSysColor(COLOR_MENU));
	if (mii.hbmpItem) {
		mii.fMask |= MIIM_BITMAP;
		if (id) {
			MainWindow.g_arBM[[id, nHeight].join("\t")] = mii.hbmpItem;
		} else {
			MainWindow.g_arBM.push(mii.hbmpItem);
		}
	}
}

MenusIcon = function (mii, src, nHeight, bIcon) {
	var image;
	mii.cbSize = mii.Size;
	if (src && src !== "-") {
		if ("string" === typeof src) {
			src = api.PathUnquoteSpaces(ExtractMacro(te, src));
			if (!/:|^\\\\/i.test(src)) {
				src = fso.BuildPath(te.Data.Installed, "script\\" + src);
			}
			mii.hbmpItem = MainWindow.g_arBM[[src, nHeight].join("\t")];
			if (mii.hbmpItem) {
				mii.fMask = mii.fMask | MIIM_BITMAP;
				return;
			}
		}
		var h16 = GetIconSize(0, 16);
		var h = nHeight < h16 ? GetIconSize(0, nHeight || 16) : nHeight || h16;
		if (/^object$/.test(src)) {
			image = src;
		} else {
			image = api.CreateObject("WICBitmap");
			if (bIcon || !image.FromFile(src)) {
				var hIcon = MakeImgIcon(src, 0, h, bIcon);
				image.FromHICON(hIcon);
				api.DestroyIcon(hIcon);
			}
		}
		if (h != image.GetHeight()) {
			image = image.GetThumbnailImage(h * image.GetWidth() / image.GetHeight(), h) || image;
		}
		AddMenuImage(mii, image, src);
	}
}

MakeMenus = function (hMenu, menus, arMenu, items, Ctrl, pt, nMin, arItem, bTrans) {
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
	for (var i = 0; i < nLen; ++i) {
		var item = items[arMenu[i]];
		var s = (item.getAttribute("Name") || item.getAttribute("Mouse") || GetKeyName(item.getAttribute("Key")) || "").replace(/\\t/i, "\t");
		var strType = String(item.getAttribute("Type")).toLowerCase();
		var path = ExtractMacro(te, item.text);
		var strFlag = strType == "menus" ? item.text.toLowerCase() : "";
		var icon = item.getAttribute("Icon");
		if (!icon && te.Data.Conf_MenuIcon) {
			if (api.PathMatchSpec(strType, "Open;Open in new tab;Open in background")) {
				var pidl = api.ILCreateFromPath(path);
				if (!api.PathIsNetworkPath(api.PathUnquoteSpaces(path))) {
					if (api.ILIsEmpty(pidl) || pidl.Unavailable) {
						var res = /(.*?)\n/.exec(path);
						if (res) {
							pidl = api.ILCreateFromPath(res[1]);
						}
					}
				}
				icon = MainWindow.GetIconImage(pidl, GetSysColor(COLOR_WINDOW), true);
			} else if (api.PathMatchSpec(strType, "Exec;Selected items")) {
				var arg = api.CommandLineToArgv(path);
				if (!api.PathIsNetworkPath(arg[0])) {
					var pidl = api.ILCreateFromPath(arg[0]);
					if (!api.ILIsEmpty(pidl) && !pidl.Unavailable) {
						icon = MainWindow.GetIconImage(pidl, GetSysColor(COLOR_WINDOW), true);
					}
				}
			}
		}
		if (strFlag == "close") {
			hMenus.pop();
			if (!hMenus.length) {
				break;
			}
		} else {
			var ar = s.split(/\t/);
			ar[0] = ExtractMacro(te, ar[0]);
			if (bTrans && !item.getAttribute("Org")) {
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
				MenusIcon(mii, icon, item.getAttribute("Height"));
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
					MenusIcon(mii, icon, item.getAttribute("Height"));
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

SaveXmlEx = function (filename, xml) {
	try {
		filename = fso.BuildPath(te.Data.DataFolder, "config\\" + filename);
		xml.save(filename);
	} catch (e) {
		if (e.number != E_ACCESSDENIED) {
			ShowError(e, [GetText("Save"), filename].join(": "));
		}
	}
}

BlurId = function (Id) {
	document.getElementById(Id).blur();
}

RunCommandLine = function (s) {
	var re = /\/select,([^,]+)/i.exec(s);
	if (re) {
		var arg = api.CommandLineToArgv(re[1]);
		Navigate(fso.GetParentFolderName(arg[0]), SBSP_NEWBROWSER);
		(function (Item) {
			setTimeout(function () {
				var FV = te.Ctrl(CTRL_FV);
				FV.SelectItem(Item, SVSI_SELECT | SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS);
			}, 99);
		})(arg[0]);
		return;
	}
	var arg = api.CommandLineToArgv(s.replace(/^\/e,|^\/n,|^\/root,/ig, ""));
	arg.shift();
	s = arg.join(" ");
	if (/^[A-Z]:\\|^\\\\/i.test(s) && IsExists(s)) {
		Navigate(s, SBSP_NEWBROWSER);
		return;
	}
	while (s = arg.shift()) {
		var ar = s.split(",");
		if (ar.length > 1) {
			Exec(te, GetSourceText(ar[1]), GetSourceText(ar[0]), te.hwnd, api.Memory("POINT"));
			continue;
		}
		Navigate(s, SBSP_NEWBROWSER);
	}
}

OpenNewProcess = function (fn, ex, mode, vOperation) {
	var uid;
	do {
		uid = String(Math.random()).replace(/^0?\./, "");
	} while (MainWindow.Exchange[uid]);
	MainWindow.Exchange[uid] = ex;
	return ShellExecute([api.PathQuoteSpaces(api.GetModuleFileName(null)), mode ? '/open' : '/run', fn, uid].join(" "), vOperation, mode ? SW_SHOWNORMAL : SW_SHOWNOACTIVATE);
}

GetAddonInfo = function (Id) {
	var info = [];

	var path = te.Data.Installed;
	var xml = api.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	var xmlfile = fso.BuildPath(path, "addons\\" + Id + "\\config.xml");
	if (fso.FileExists(xmlfile)) {
		xml.load(xmlfile);

		GetAddonInfo2(xml, info, "General", true);
		var lang = GetLangId();
		if (!/^en/.test(lang)) {
			GetAddonInfo2(xml, info, "en", true);
			info.Name = GetTextR(info.Name);
		}
		var res = /(\w+)_/.exec(lang);
		if (res && !/zh_cn/i.test(lang)) {
			GetAddonInfo2(xml, info, res[1]);
		}
		GetAddonInfo2(xml, info, lang);
		if (!info.Name) {
			info.Name = Id;
		}
	}
	return info;
}

GetAddonInfoName = function (Id) {
	var info = GetAddonInfo(Id);
	return info.Name;
}

GetAddonInfo2 = function (xml, info, Tag, bTrans) {
	var items = xml.getElementsByTagName(Tag);
	if (items.length) {
		var item = items[0].childNodes;
		for (var i = 0; i < item.length; ++i) {
			var n = item[i].tagName;
			var s = item[i].textContent || item[i].text;
			info[n] = (bTrans && /Name|Description/i.test(n) ? GetText(s) : s);
		}
	}
}

OpenXml = function (strFile, bAppData, bEmpty, strInit) {
	var xml = api.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	var path = fso.BuildPath(te.Data.DataFolder, "config\\" + strFile);
	if (fso.FileExists(path) && xml.load(path)) {
		return xml;
	}
	if (!bAppData) {
		path = fso.BuildPath(te.Data.Installed, "config\\" + strFile);
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
	path = fso.BuildPath(te.Data.Installed, "init\\" + strFile);
	if (fso.FileExists(path) && xml.load(path)) {
		return xml;
	}
	return bEmpty ? xml : null;
}

CreateXml = function (bRoot) {
	var xml = api.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	xml.appendChild(xml.createProcessingInstruction("xml", 'version="1.0" encoding="UTF-8"'));
	if (bRoot) {
		xml.appendChild(xml.createElement("TablacusExplorer"));
	}
	return xml;
}

DownloadFile = function (url, fn) {
	return api.URLDownloadToFile(null, url, fn);
}

Extract = function (Src, Dest, xhr) {
	var hr;
	if (xhr) {
		hr = DownloadFile(xhr, Src);
		if (hr) {
			return hr;
		}
	}
	for (var i in eventTE.extract) {
		hr = eventTE.extract[i](Src, Dest);
		if (isFinite(hr)) {
			return hr;
		}
	}
	return api.Extract(fso.BuildPath(system32, "zipfldr.dll"), "{E88DCCE0-B7B3-11d1-A9F0-00AA0060FA31}", Src, Dest);
}

OptionRef = function (Id, s, pt) {
	return RunEvent4("OptionRef", Id, s, pt);
}

OptionDecode = function (Id, p) {
	return RunEvent3("OptionDecode", Id, p);
}

OptionEncode = function (Id, p) {
	return RunEvent3("OptionEncode", Id, p);
}

function GetAddons() {
	ShowOptions("Tab=Get Addons");
}

function CheckUpdate(arg) {
	MainWindow.OpenHttpRequest("https://api.github.com/repos/tablacus/TablacusExplorer/releases/latest", "http://tablacus.github.io/TablacusExplorerAddons/te/releases.json", CheckUpdate2, arg);
}

function CheckUpdate2(xhr, url, arg1) {
	var arg = {};
	var Text = xhr.get_responseText ? xhr.get_responseText() : xhr.responseText;
	var json = window.JSON ? JSON.parse(Text) : api.GetScriptDispatch("function fn () { return " + Text + "}", "JScript", {}).fn();
	if (json.assets && json.assets[0]) {
		arg.size = json.assets[0].size / 1024;
		arg.url = json.assets[0].browser_download_url;
	}
	if (!arg.url) {
		return;
	}
	arg.file = fso.GetFileName(arg.url.replace(/\//g, "\\"));
	var ver = 0;
	res = /(\d+)/.exec(arg.file);
	if (res) {
		ver = api.Add(20000000, res[1]);
	}
	if (ver <= AboutTE(0)) {
		if ((arg1 && arg1.silent) || MessageBox(AboutTE(2) + "\n" + GetText("the latest version"), TITLE, MB_ICONINFORMATION)) {
			if (api.GetKeyState(VK_SHIFT) >= 0 || api.GetKeyState(VK_CONTROL) >= 0) {
				MainWindow.RunEvent1("CheckUpdate", arg1);
				return;
			}
		}
	}
	if (!(arg1 && arg1.noconfirm)) {
		var s = (api.LoadString(hShell32, 60) || "%").replace(/%.*/, api.sprintf(99, "%d.%d.%d (%.1lfKB)", ver / 10000 % 100, ver / 100 % 100, ver % 100, arg.size));
		if (!confirmOk([GetText("Update available"), s, GetText("Do you want to install it now?")].join("\n"))) {
			return;
		}
	}
	arg.temp = fso.BuildPath(fso.GetSpecialFolder(2).Path, "tablacus");
	CreateFolder2(arg.temp);
	wsh.CurrentDirectory = arg.temp;
	arg.InstalledFolder = te.Data.Installed;
	arg.zipfile = fso.BuildPath(arg.temp, arg.file);
	arg.temp += "\\explorer";
	DeleteItem(arg.temp);
	CreateFolder2(arg.temp);
	MainWindow.OpenHttpRequest(arg.url, "http://tablacus.github.io/TablacusExplorerAddons/te/" + (arg.url.replace(/^.*\//, "")), CheckUpdate3, arg);
}

function CheckUpdate3(xhr, url, arg) {
	var hr = Extract(arg.zipfile, arg.temp, xhr);
	if (hr) {
		MessageBox([api.LoadString(hShell32, 4228).replace(/^\t/, "").replace("%d", api.sprintf(99, "0x%08x", hr)), GetText("Extract"), fso.GetFileName(arg.zipfile)].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
		return;
	}
	var te_exe = arg.temp + "\\te64.exe";
	var nDog = 300;
	while (!fso.FileExists(te_exe)) {
		if (wsh.Popup(GetText("Please wait."), 1, TITLE, MB_OKCANCEL) == IDCANCEL || nDog-- == 0) {
			return;
		}
	}
	var arDel = [];
	var addons = arg.temp + "\\addons";
	if (fso.FolderExists(arg.temp + "\\config")) {
		arDel.push(arg.temp + "\\config");
	}
	for (var i = 32; i <= 64; i += 32) {
		te_exe = arg.temp + '\\te' + i + '.exe';
		var te_old = fso.BuildPath(te.Data.Installed, 'te' + i + '.exe');
		if (!fso.FileExists(te_old) || fso.GetFileVersion(te_exe) == fso.GetFileVersion(te_old)) {
			arDel.push(te_exe);
		}
	}
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
	var ppid = api.Memory("DWORD");
	api.GetWindowThreadProcessId(te.hwnd, ppid);
	arg.pid = ppid[0];
	MainWindow.CreateUpdater(arg);
}

function ShowAbout() {
	ShowDialog(fso.BuildPath(te.Data.Installed, "script\\dialog.html"), { MainWindow: MainWindow, Query: "about", Modal: false, width: 640, height: 360 });
}

function EscapeUpdateFile(s) {
	return s.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

confirmYN = function (s, title) {
	return MessageBox(s, title, MB_ICONQUESTION | MB_YESNO) == IDYES;
}

confirmOk = function (s, title) {
	return MessageBox(s || "Are you sure?", title, MB_ICONQUESTION | MB_OKCANCEL) == IDOK;
}

MessageBox = function (s, title, uType) {
	return api.MessageBox(api.GetForegroundWindow(), GetTextR(s), GetTextR(title) || TITLE, uType);
}

createHttpRequest = function () {
	try {
		return window.XMLHttpRequest && g_.IEVer >= 9 ? new XMLHttpRequest() : api.CreateObject("Msxml2.XMLHTTP");
	} catch (e) {
		return api.CreateObject("Microsoft.XMLHTTP");
	}
}

OpenHttpRequest = function (url, alt, fn, arg) {
	var xhr = createHttpRequest();
	xhr.onreadystatechange = function () {
		if (xhr.readyState == 4) {
			if (arg && arg.pcRef) {
				--arg.pcRef[0];
			}
			if (xhr.status == 200) {
				return fn(xhr, url, arg);
			}
			if (/^http/.test(alt)) {
				return OpenHttpRequest(/^https/.test(url) && alt == "http" ? url.replace(/^https/, alt) : alt, '', fn, arg);
			}
			MessageBox([api.sprintf(999, api.LoadString(hShell32, 4227).replace(/^\t/, ""), xhr.status), url].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
		}
	}
	if (/ml$/i.test(url)) {
		url += "?" + Math.floor(new Date().getTime() / 60000);
	}
	if (arg && arg.pcRef) {
		++arg.pcRef[0];
	}
	xhr.open("GET", url, false);
	try {
		xhr.send(null);
	} catch (e) { }
}

InputDialog = function (text, defaultText) {
	return prompt(GetTextR(text), defaultText);
}

AddonOptions = function (Id, fn, Data, bNew) {
	var sParent = te.Data.Installed;
	LoadLang2(fso.BuildPath(sParent, "addons\\" + Id + "\\lang\\" + GetLangId() + ".xml"));
	var items = te.Data.Addons.getElementsByTagName(Id);
	if (!items.length) {
		var root = te.Data.Addons.documentElement;
		if (root) {
			root.appendChild(te.Data.Addons.createElement(Id));
		}
	}
	var info = GetAddonInfo(Id);
	var sURL = "addons\\" + Id + "\\options.html";
	if (!Data) {
		Data = {};
	}
	Data.id = Id;
	var sFeatures = info.Options;
	if (/^Location$/i.test(sFeatures)) {
		sFeatures = "Common:6:6";
	}
	var res = /Common:([\d,]+):(\d)/i.exec(sFeatures);
	if (res) {
		sURL = "script\\location.html";
		Data.show = res[1];
		Data.index = res[2];
		sFeatures = 'Default';
	}
	sURL = fso.BuildPath(te.Data.Installed, sURL);
	var opt = { MainWindow: MainWindow, Data: Data, event: {} };
	if (fn) {
		opt.event.TEOk = fn;
	} else if (window.g_Chg) {
		opt.event.TEOk = function () {
			g_Chg.Addons = true;
		}
	}
	if (bNew || window.Addon == 1 || api.GetKeyState(VK_CONTROL) < 0) {
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
		var opt = { MainWindow: MainWindow, Data: Data, event: {} };
		if (fn) {
			opt.event.TEOk = fn;
		} else if (window.g_Chg) {
			opt.event.TEOk = function () {
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
		return;
	}
	if (!g_.elAddons[Id]) {
		opt.event.onload = function () {
			var cInput = el.contentWindow.document.getElementsByTagName('input');
			for (var i in cInput) {
				if (/^ok$|^cancel$/i.test(cInput[i].className)) {
					cInput[i].style.display = 'none';
				}
			}
			el.contentWindow.g_.Inline = true;
		}
		external.WB.Data = opt;
		var el = document.createElement('iframe');
		el.id = 'panel1_' + Id;
		el.src = sURL;
		el.style.cssText = 'width: 100%; border: 0; padding: 0; margin: 0';
		g_.elAddons[Id] = el;
		var o = document.getElementById('panel1_2');
		o.style.display = "block";
		o.appendChild(el);
		o = document.getElementById('tab1_');
		o.insertAdjacentHTML("BeforeEnd", '<label id="tab1_' + Id + '" class="button" style="width: 100%" onmousedown="ClickTree(this);">' + info.Name + '</label><br>');
	}
	ClickTree(document.getElementById('tab1_' + Id));
}

function CalcVersion(s) {
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

GethwndFromPid = function (ProcessId, bDT) {
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
			var ppid = api.Memory("DWORD");
			api.GetWindowThreadProcessId(hwnd, ppid);
			if (ProcessId == ppid[0]) {
				return hwnd;
			}
		}
	} while (hwnd = api.GetWindow(hwnd, GW_HWNDNEXT));
	return null;
}

PopupContextMenu = function (Item, FV) {
	if ("string" === typeof Item) {
		var arg = api.CommandLineToArgv(Item);
		Item = api.CreateObject("FolderItems");
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

GetAddonElement = function (id) {
	var items = te.Data.Addons.getElementsByTagName(id.toLowerCase());
	if (items.length) {
		return items[0];
	}
	return {
		getAttribute: function (s) {
			return "";
		},
		setAttribute: function () { }
	}
}

GetAddonOption = function (id, strTag) {
	return GetAddonElement(id).getAttribute(strTag);
}

GetAddonOptionEx = function (id, strTag) {
	return api.LowPart(GetAddonOption(id, strTag));
}

GetInnerFV = function (id) {
	var TC = te.Ctrl(CTRL_TC, id);
	if (TC && TC.SelectedIndex >= 0) {
		return TC.Selected;
	}
	return null;
}

OpenInExplorer = function (pid1) {
	if (pid1) {
		CancelWindowRegistered();
		var ar = [];
		pid1 = pid1.FolderItem || pid1;
		var pid = pid1;
		if (!api.ILIsParent(ssfNETWORK, pid, false)) {
			for (var n = 99; !api.ILIsEmpty(pid) && n--; pid = api.ILGetParent(pid)) {
				var path = api.GetDisplayNameOf(pid, SHGDN_FORPARSING | SHGDN_INFOLDER | SHGDN_ORIGINAL);
				if (!path || /\\/.test(path)) {
					ar = [];
					break;
				}
				ar.unshift(path);
			}
		}
		if (!ar.length) {
			ar = [api.GetDisplayNameOf(pid1, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL)];
		}
		api.CreateProcess([api.PathQuoteSpaces(wsh.ExpandEnvironmentStrings("%SystemRoot%\\explorer.exe")), api.PathQuoteSpaces(ar.join("\\"))].join(" "), null, 0, 1, 0);
	}
}

CancelWindowRegistered = function () {
	clearTimeout(g_.tidWindowRegistered);
	g_.bWindowRegistered = false;
	g_.tidWindowRegistered = setTimeout(function () {
		g_.bWindowRegistered = true;
	}, 9999);
}

ShowDialogEx = function (mode, w, h, ele) {
	ShowDialog(fso.BuildPath(te.Data.Installed, "script\\dialog.html"), { MainWindow: MainWindow, Query: mode, width: w, height: h, element: ele });
}

ShowNew = function (Ctrl, pt, Mode) {
	var FV = GetFolderView(Ctrl, pt);
	var path = api.GetDisplayNameOf(FV, SHGDN_FORPARSING | SHGDN_ORIGINAL);
	if (/^[A-Z]:\\|^\\\\/i.test(path)) {
		ShowDialog(fso.BuildPath(te.Data.Installed, "script\\dialog.html"), { MainWindow: MainWindow, Query: "new", Mode: Mode, path: path, FV: FV, Modal: false, width: 480, height: 120 });
	}
}

CreateNewFolder = function (Ctrl, pt) {
	ShowNew(Ctrl, pt, "folder");
	return S_OK;
}

CreateNewFile = function (Ctrl, pt) {
	ShowNew(Ctrl, pt, "file");
	return S_OK;
}

InputMouse = function (o) {
	ShowDialogEx("mouse", 500, 420, o || (document.E && document.E.MouseMouse) || document.F.MouseMouse || document.F.Mouse);
}

InputKey = function (o) {
	ShowDialogEx("key", 320, 120, o || (document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key);
}

ShowIconEx = function (o, mode) {
	ShowDialogEx("icon" + (mode || ""), 640, 480, o || document.F.Icon);
}

ShowLocationEx = function (s) {
	ShowDialog(fso.BuildPath(te.Data.Installed, "script\\location.html"), { MainWindow: MainWindow, Data: s });
}

function MakeKeySelect() {
	var oa = document.getElementById("_KeyState");
	if (oa) {
		var ar = [];
		for (var i = 0; i < 4; ++i) {
			var s = MainWindow.g_.KeyState[i][0];
			ar.push('<label><input type="checkbox" onclick="KeyShift(this)" id="_Key', s, '">', s, '&nbsp;</label>');
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
	s.sort(function (a, b) {
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
			o.text = j + "\x80";
		}
	}
}

function SetKeyShift() {
	var key = ((document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key).value;
	for (var i = 0; i < MainWindow.g_.KeyState.length; ++i) {
		var s = MainWindow.g_.KeyState[i][0];
		var o = document.getElementById("_Key" + s);
		if (o) {
			o.checked = key.indexOf(s + "+") >= 0;
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

function KeyShift(o) {
	var oKey = (document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key;
	var key = oKey.value;
	var shift = o.id.replace(/^_Key(.*)/, "$1+");
	key = key.replace(shift, "");
	if (o.checked) {
		key = shift + key;
	}
	oKey.value = key;
}

function KeySelect(o) {
	var oKey = (document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key;
	oKey.value = oKey.value.replace(/(\+)[^\+]*$|^[^\+]*$/, "$1") + o[o.selectedIndex].value;
}

GetLangId = function (nDefault) {
	if (!nDefault && te.Data.Conf_Lang) {
		return te.Data.Conf_Lang;
	}
	var lang = navigator.userLanguage.replace(/\-/, '_').toLowerCase();
	if (nDefault != 2) {
		if (!fso.FileExists(fso.BuildPath(te.Data.Installed, "lang\\" + lang + ".xml"))) {
			lang = lang.replace(/_.*/, "");
		}
	}
	if (!te.Data.Conf_Lang) {
		te.Data.Conf_Lang = lang;
	}
	return lang;
}

GetSourceText = function (s) {
	try {
		return (MainWindow.LangSrc || LangSrc)[s] || s;
	} catch (e) {
		return s;
	}
}

GetFolderView = function (Ctrl, pt, bStrict) {
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

GetSelectedArray = function (Ctrl, pt, bPlus) {
	var Selected, SelItem;
	var FV = null;
	var bSel = true;
	switch (Ctrl.Type) {
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
		Selected = api.CreateObject("FolderItems");
		if (bPlus) {
			Selected.AddItem(SelItem);
		}
	}
	return [Selected, SelItem, FV];
}

StripAmp = function (s) {
	return String(s).replace(/\(&\w\)|&/, "").replace(/\.\.\.$/, "");
}

EncodeSC = function (s) {
	return String(s).replace(/[&"<>]/g, function (strMatch) {
		return "&#" + strMatch.charCodeAt(0) + ";";
	});
}

DecodeSC = function (s) {
	return String(s).replace(/&([\w#]{1,5});/g, function (strMatch, ref) {
		var res = /^#x([\da-f]+)$/i.exec(ref)
		if (res) {
			return String.fromCharCode(parseInt(res[1], 16));
		}
		res = /^#(\d+)$/.exec(ref)
		if (res) {
			return String.fromCharCode(res[1]);
		}
		return { quot: '"', amp: '&', lt: '<', gt: '>' }[ref.toLowerCase()] || '&' + ref + ';';
	});
}

function GetGestureX(ar) {
	var o, s = "";
	while (o = ar.shift()) {
		if (api.GetKeyState(o[1]) < 0) {
			s += o[0];
		}
	}
	return s;
}

GetGestureKey = function () {
	return GetGestureX([
		["S", VK_SHIFT],
		["C", VK_CONTROL],
		["A", VK_MENU]
	]);
}

GetGestureButton = function () {
	return GetGestureX([
		["1", VK_LBUTTON],
		["2", VK_RBUTTON],
		["3", VK_MBUTTON],
		["4", VK_XBUTTON1],
		["5", VK_XBUTTON2]
	]);
}

GetWebColor = function (c) {
	var res = /(\d{1,3}) (\d{1,3}) (\d{1,3})/.exec(c);
	if (res) {
		c = res[3] * 65536 + res[2] * 256 + res[1] * 1;
	}
	return isNaN(c) && /^#[0-9a-f]{3,6}$/i ? c : api.sprintf(8, "#%06x", ((c & 0xff) << 16) | (c & 0xff00) | ((c & 0xff0000) >> 16));
}

GetWinColor = function (c) {
	var res = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(c);
	if (res) {
		return Number(["0x", res[3], res[2], res[1]].join(""));
	}
	res = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(c);
	if (res) {
		return Number(["0x", res[3], res[3], res[2], res[2], res[1], res[1]].join(""));
	}
	res = /(\d{1,3}) (\d{1,3}) (\d{1,3})/.exec(c);
	if (res) {
		return res[3] * 65536 + res[2] * 256 + res[1] * 1;
	}
	return c;
}

ChooseColor = function (c) {
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

ChooseWebColor = function (c) {
	c = ChooseColor(GetWinColor(c));
	if (isFinite(c)) {
		return GetWebColor(c);
	}
}

SetCursor = function (o, s) {
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

function MouseOver(o) {
	o = o.srcElement || o;
	if (/^button$|^menu$/i.test(o.className)) {
		if (g_.objHover && o != g_.objHover) {
			MouseOut();
		}
		var pt = api.Memory("POINT");
		api.GetCursorPos(pt, true);
		var ptc = pt.Clone();
		api.ScreenToClient(api.GetWindow(document), ptc);
		if (o == document.elementFromPoint(ptc.x, ptc.y) || HitTest(o, pt)) {
			g_.objHover = o;
			o.className = 'hover' + o.className;
		}
	}
}

function MouseOut(s) {
	if (g_.objHover) {
		if ("string" !== typeof s || g_.objHover.id.indexOf(s) >= 0) {
			if (g_.objHover.className == 'hoverbutton') {
				g_.objHover.className = 'button';
			} else if (g_.objHover.className == 'hovermenu') {
				g_.objHover.className = 'menu';
			}
			g_.objHover = null;
		}
	}
	return S_OK;
}

InsertTab = function (e) {
	if (!e) {
		e = event;
	}
	var ot = e.srcElement;
	if (e.keyCode == VK_TAB) {
		ot.focus();
		if (document.all && document.selection) {
			var selection = document.selection.createRange();
			if (selection) {
				selection.text += "\t";
				return false;
			}
		}
		var i = ot.selectionEnd;
		var s = ot.value;
		ot.value = s.substr(0, i) + "\t" + s.substr(i, s.length);
		ot.selectionStart = ++i;
		ot.selectionEnd = i;
		return false;
	}
	return true;
}

RegEnumKey = function (hKey, Name) {
	var server = te.GetObject("winmgmts:\\\\.\\root\\default:StdRegProv");
	var method = server.Methods_.Item("EnumKey");
	var iParams = method.InParameters.SpawnInstance_();
	iParams.hDefKey = hKey;
	iParams.sSubKeyName = Name;
	var r = server.ExecMethod_(method.Name, iParams).sNames;
	return r !== null ? r.toArray() : [];
}

FindText = function () {
	api.OleCmdExec(document, null, 32, 0, 0);
}

OpenDialogEx = function (path, filter, bFilesOnly) {
	var commdlg = te.CommonDialog;
	var te_path = te.Data.Installed;
	var res = /^\.\.(\/.*)/.exec(path);
	if (res) {
		path = te_path + (res[1].replace(/\//g, "\\"));
	}
	path = api.PathUnquoteSpaces(ExtractMacro(te, path));
	if (!api.PathIsDirectory(path)) {
		path = fso.GetParentFolderName(path);
		if (!api.PathIsDirectory(path)) {
			path = fso.GetDriveName(te_path);
		}
	}
	commdlg.InitDir = path;
	commdlg.Filter = MakeCommDlgFilter(filter);
	commdlg.Flags = OFN_FILEMUSTEXIST | OFN_EXPLORER | OFN_ENABLESIZING | (bFilesOnly ? 0 : OFN_ENABLEHOOK);
	if (commdlg.ShowOpen()) {
		return api.PathQuoteSpaces(commdlg.FileName);
	}
}

OpenDialog = function (path, bFilesOnly) {
	return OpenDialogEx(path, null, bFilesOnly);
}

ChooseFolder = function (path, pt) {
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

BrowseForFolder = function (path) {
	return OpenDialogEx(path, api.LoadString(hShell32, 4131) + "|<Folder>");
}

InvokeCommand = function (Items, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon, FV, uCMF) {
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

SetRenameMenu = function (n) {
	ExtraMenuCommand[CommandID_RENAME + n - 1] = function (Ctrl, pt, Name, nVerb) {
		if (Ctrl.Type == CTRL_TV) {
			setTimeout(function () {
				Ctrl.Focus();
				wsh.SendKeys("{F2}");
			}, 99);
		} else {
			return S_FALSE;
		}
	};
}

ShowError = function (e, s, i) {
	var sl = (s || "").toLowerCase();
	if (isFinite(i)) {
		if (eventTA[sl][i]) {
			s = eventTA[sl][i] + " : " + s;
		}
	}
	if (!g_.ShowError) {
		g_.ShowError = true;
		setTimeout(function () {
			g_.ShowError = MessageBox([e.stack || e.description || e.toString(), s, AboutTE(3)].join("\n\n"), TITLE, MB_OKCANCEL) != IDOK;
		}, 99);
	}
}

ApiStruct = function (oTypedef, nAli, oMemory) {
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
	this.Memory = "object" === typeof oMemory ? oMemory : api.Memory("BYTE", this.Size);
	this.Read = function (Id) {
		var ar = this.Typedef[Id];
		if (ar) {
			return this.Memory.Read(ar[3], ar[0]);
		}
	};
	this.Write = function (Id, Data) {
		var ar = this.Typedef[Id];
		if (ar) {
			this.Memory.Write(ar[3], ar[0], Data);
		}
	};
}

FindChildByClass = function (hwnd, s) {
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

GetNavigateFlags = function (FV, bParent) {
	if (!FV && OpenMode != SBSP_NEWBROWSER) {
		FV = te.Ctrl(CTRL_FV);
	}
	return (!bParent && api.GetKeyState(VK_MBUTTON) < 0) || api.GetKeyState(VK_CONTROL) < 0 || GetLock() ? SBSP_NEWBROWSER : OpenMode;
}

AddEvent("ConfigChanged", function (s) {
	te.Data["bSave" + s] = true;
});

GetSysColor = function (i) {
	var c = g_.Colors[i];
	return c !== void 0 ? c : api.GetSysColor(i);
}

SetSysColor = function (i, color) {
	g_.Colors[i] = color;
	if (i == COLOR_BTNFACE) {
		te.Background = isFinite(color) ? api.CreateSolidBrush(color) : null;
	}
}

ShellExecute = function (s, vOperation, nShow, vDir2, pt) {
	var cmd = ExtractMacro(te, s);
	var res = /^\s*"([^"]*)"\s*(.*)/.exec(cmd) || /^\s*([^\s]*)\s*(.*)/.exec(cmd);
	if (res) {
		var vDir = fso.GetParentFolderName(res[1]) || vDir2;
		if (pt && vDir.Type) {
			vDir = (GetFolderView(Ctrl, pt) || { FolderItem: {} }).FolderItem.Path;
		}
		return sha.ShellExecute(res[1], res[2], vDir, vOperation, nShow);
	}
}

CreateFont = function (LogFont) {
	var key = [LogFont.lfFaceName, LogFont.lfHeight, LogFont.lfCharSet, LogFont.lfWeight, LogFont.lfItalic, LogFont.lfUnderline].join("\t");
	var hFont = te.Data.Fonts[key];
	if (!hFont) {
		hFont = api.CreateFontIndirect(LogFont);
		te.Data.Fonts[key] = hFont;
	}
	return hFont;
}

Activate = function (o, id) {
	var TC = te.Ctrl(CTRL_TC);
	if (TC && TC.Id != id) {
		var FV = GetInnerFV(id);
		if (FV) {
			FV.Focus();
			if (o) {
				setTimeout(function () {
					o.focus();
				}, 99)
			}
		}
	}
}

function DetectProcessTag(e) {
	var el = (e || event).srcElement;
	return /input|textarea/i.test(el.tagName) || /selectable/i.test(el.className);
}

AddEventEx(window, "load", function () {
	document.body.onselectstart = DetectProcessTag;
	document.body.oncontextmenu = DetectProcessTag;
	document.body.onmousewheel = function () {
		return api.GetKeyState(VK_CONTROL) >= 0;
	};
});

Alt = function () {
	return S_OK;
}

GetSavePath = function (FolderItem) {
	var path = api.GetDisplayNameOf(FolderItem, SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
	if (!/^[A-Z]:\\|^\\\\\w/i.test(path)) {
		var res = IsSearchPath(FolderItem);
		if (res) {
			return res[0];
		}
	}
	if (/\?/.test(path)) {
		var nCount = api.ILGetCount(FolderItem);
		path = [];
		while (nCount-- > 0) {
			path.unshift(api.GetDisplayNameOf(FolderItem, (nCount > 0 ? SHGDN_FORADDRESSBAR : 0) | SHGDN_FORPARSING | SHGDN_INFOLDER | SHGDN_ORIGINAL));
			FolderItem = api.ILRemoveLastID(FolderItem);
		}
		return path.join("\\")
	}
	return path;
}

ExecAddonScript = function (type, s, fn, arError, o, arStack) {
	var sc = api.GetScriptDispatch(s, type, o,
		function (ei, SourceLineText, dwSourceContext, lLineNumber, CharacterPosition) {
			arError.push([api.SysAllocString(ei.bstrDescription), fn, api.sprintf(16, "Line: %d", lLineNumber)].join("\n"));
		}
	);
	if (sc && arStack) {
		arStack.push(sc);
	}
	return sc;
}

LoadAddon = function (ext, Id, arError, param) {
	var r;
	try {
		var sc;
		var ar = ext.split(".");
		if (ar.length == 1) {
			ar.unshift("script");
		}
		var fn = "addons" + "\\" + Id + "\\" + ar.join(".");
		var ado = OpenAdodbFromTextFile(fn, "utf-8");
		if (ado) {
			var s = ado.ReadText();
			ado.Close();
			if (s) {
				if (ar[1] == "js") {
					sc = new Function(s);
				} else if (ar[1] == "vbs") {
					sc = ExecAddonScript("VBScript", s, fn, arError, { "_Addon_Id": { "Addon_Id": Id }, window: window }, Addons["_stack"]);
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
		arError.push([e.stack || e.description || e.toString(), fn].join("\n"));
	}
	return r;
}

function CloseSubWindows() {
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
}

AddEventEx(window, "beforeunload", CloseSubWindows);

function FireEvent(o, event) {
	if (o) {
		if (o.fireEvent) {
			return o.fireEvent('on' + event);
		} else if (document.createEvent) {
			var evt = document.createEvent("HTMLEvents");
			evt.initEvent(event, true, true);
			return !o.dispatchEvent(evt);
		}
	}
}

RemoveCommand = function (hMenu, ContextMenu, strDelete) {
	if (ContextMenu) {
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
}

DeleteTempFolder = function () {
	DeleteItem(fso.BuildPath(fso.GetSpecialFolder(2).Path, "tablacus"));
}

PerformUpdate = function () {
	var oExec = wsh.Exec(g_.strUpdate);
	wsh.AppActivate(oExec.ProcessID);
}

OpenContains = function (Ctrl, pt) {
	var Items = GetSelectedItems(Ctrl, pt);
	for (var j in Items) {
		var Item = Items.Item(j);
		var path = Item.Path;
		Navigate(fso.GetParentFolderName(path), SBSP_NEWBROWSER);
		(function (Item) {
			setTimeout(function () {
				var FV = te.Ctrl(CTRL_FV);
				FV.SelectItem(Item, SVSI_SELECT | SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS);
			}, 99);
		})(Item);
	}
}

function OrganizePath(fn, base)
{
	if (/"/.test(fn)) {
		fn = api.PathUnquoteSpaces(fn);
	}
	fn = ExtractMacro(te, fn);
	if (base && !/^[A-Z]:\\|^\\\\\w/i.test(fn)) {
		fn = fso.BuildPath(base, fn);
	}
	return fn;
}

OpenAdodbFromTextFile = function (fn, charset) {
	fn = OrganizePath(fn, te.Data.Installed);
	var ado = api.CreateObject("ads");
	if (charset) {
		try {
			ado.CharSet = charset;
			ado.Open();
			ado.LoadFromFile(fn);
		} catch (e) {
			ado.Close();
			return;
		}
		return ado;
	}
	try {
		ado.Type = adTypeBinary;
		ado.Open();
		ado.LoadFromFile(fn);
		var s = ado.Read(8192);
	} catch (e) {
		ado.Close();
		return;
	}
	charset = MainWindow.RunEvent4("DetectCharSet", s, fn);
	if (!charset) {
		s = api.SysAllocString(s, 28591);
		charset = "_autodetect_all";
		if (/^\xEF\xBB\xBF|^([\x09\x0A\x0D\x20-\x7E]|[\xC2-\xDF][\x80-\xBF]|\xE0[\xA0-\xBF][\x80-\xBF]|[\xE1-\xEC\xEE\xEF][\x80-\xBF]{2}|\xED[\x80-\x9F][\x80-\xBF]|\xF0[\x90-\xBF][\x80-\xBF]{2}|[\xF1-\xF3][\x80-\xBF]{3}|\xF4[\x80-\x8F][\x80-\xBF]{2})*[\x80-\xBF]{0,3}$/.test(s)) {
			charset = 'utf-8';
		} else if (/^\xFF\xFE|^\xFE\xFF/.test(s)) {
			charset = 'unicode';
		}
	}
	ado.Position = 0;
	ado.Type = adTypeText;
	ado.CharSet = charset;
	return ado;
}

WmiProcess = function (arg, fn) {
	try {
		var server = te.GetObject("winmgmts:\\\\.\\root\\cimv2");
		if (server) {
			var cols = server.ExecQuery("SELECT * FROM Win32_Process " + arg);
			for (var list = new Enumerator(cols); !list.atEnd(); list.moveNext()) {
				fn(list.item());
			}
		}
	} catch (e) { }
}

function CalcElementHeight(o, em) {
	if (o) {
		if (g_.IEVer >= 9) {
			o.style.height = "calc(100vh - " + em + "em)";
		} else {
			var h = (document.documentElement || document.body).clientHeight;
			h += MainWindow.DefaultFont.lfHeight * em;
			if (h > 0) {
				o.style.height = h + 'px';
				o.style.height = 2 * h - o.offsetHeight + "px";
			}
		}
	}
}

AddFavoriteEx = function (Ctrl, pt) {
	GetFolderView(Ctrl, pt).Focus();
	AddFavorite();
	return S_OK
}

importScript = function (fn) {
	var hr = E_FAIL;
	var ado = OpenAdodbFromTextFile(fn, "utf-8");
	if (ado) {
		hr = ExecScriptEx(window.Ctrl, ado.ReadText(), /\.vbs/i.test(fn) ? "VBScript" : "JScript", window.pt, window.dataObj, window.grfKeyState, window.pdwEffect, window.bDrop);
		ado.Close();
	}
	return hr;
}

function SameFolderItems(Items1, Items2) {
	var i = Items1.Count;
	if (i != Items2.Count) {
		return false;
	}
	while (i-- > 0) {
		if (!api.ILIsEqual(Items1.Item(i), Items2.Item(i))) {
			return false;
		}
	}
	return true;
}

function GetSelectedItems(Ctrl, pt) {
	var FV = GetFolderView(Ctrl, pt);
	if (FV) {
		return FV.SelectedItems();
	}
}

function GetThumbnail(image, m, f) {
	var w = image.GetWidth(), h = image.GetHeight(), z = m / Math.max(w, h);
	if (z == 1 || (f && z > 1)) {
		return image;
	}
	return image.GetThumbnailImage(w * z, h * z);
}

function AddonBeforeRemove(Id) {
	CollectGarbage();
	var arError = [];
	var r = LoadAddon("remove.js", Id, arError);
	if (arError.length) {
		MessageBox(arError.join("\n\n"), TITLE, MB_OK);
	}
	return r;
}
function ColumnsReplace(Ctrl, pid, fmt, fn, priority) {
	if (!Ctrl.ColumnsReplace) {
		try {
			Ctrl.ColumnsReplace = api.CreateObject("Object");
		} catch (e) {
			return;
		}
		if (!Ctrl.ColumnsReplace) {
			return;
		}
	}
	var n = api.PSGetDisplayName(pid);
	fn.fmt = fmt;
	if (Ctrl.ColumnsReplace[n] && priority != 2) {
		if (Ctrl.ColumnsReplace[n] === fn) {
			return;
		}
		if (!Ctrl.ColumnsReplace[n].push) {
			Ctrl.ColumnsReplace[n] = [Ctrl.ColumnsReplace[n]];
			Ctrl.ColumnsReplace[n].fmt = Ctrl.ColumnsReplace[n][0].fmt;
		}
		for (var i = Ctrl.ColumnsReplace[n]; i--;) {
			if (Ctrl.ColumnsReplace[n][i] === fn) {
				return;
			}
		}
		if (priority) {
			Ctrl.ColumnsReplace[n].push(fn);
		} else {
			Ctrl.ColumnsReplace[n].unshift(fn);
		}
	} else {
		Ctrl.ColumnsReplace[n] = fn;
	}
}

function CustomSort(FV, id, r, fnAdd, fnComp) {
	var Progress = api.CreateObject("ProgressDialog");
	Progress.StartProgressDialog(te.hwnd, null, 2);
	var Name = api.PSGetDisplayName(id) || id;
	Progress.SetLine(1, api.LoadString(hShell32, 13585) || api.LoadString(hShell32, 6478), true);
	FV.Parent.LockUpdate();
	try {
		var Items = FV.Items();
		var List = [];
		for (var i = Items.Count; i--;) {
			List.push([i, fnAdd(Items.Item(i), FV)]);
		}
		List.sort(fnComp);
		if (r) {
			List = List.reverse();
		}
		var ViewMode = api.SendMessage(FV.hwndList, LVM_GETVIEW, 0, 0);
		if (ViewMode == 1 || ViewMode == 3) {
			api.SendMessage(FV.hwndList, LVM_SETVIEW, 4, 0);
		}
		var FolderFlags = FV.FolderFlags;
		FV.FolderFlags = FolderFlags | FWF_AUTOARRANGE;
		FV.GroupBy = "System.Null";
		var pt = api.Memory("POINT");
		FV.GetItemPosition(Items.Item(0), pt);
		var nMax = List.length;
		var p = nMax / 100;
		Progress.SetLine(1, api.LoadString(hShell32, 50690) + Name, true);
		for (var i = 0; i < nMax && !Progress.HasUserCancelled(); ++i) {
			Progress.SetTitle((i / p).toFixed(0) + "%");
			Progress.SetProgress(i, nMax);
			FV.SelectAndPositionItem(Items.Item(List[i][0]), 0, pt);
		}
		api.SendMessage(FV.hwndList, LVM_SETVIEW, ViewMode, 0);
		FV.FolderFlags = FolderFlags;
		var hHeader = api.SendMessage(FV.hwndList, LVM_GETHEADER, 0, 0);
		var item = api.Memory("HDITEM");
		item.mask = HDI_TEXT | HDI_FORMAT;
		item.pszText = api.Memory("WCHAR", 260);
		item.cchTextMax = 260;
		if (id) {
			for (var i = api.SendMessage(hHeader, HDM_GETITEMCOUNT, 0, 0); i-- > 0;) {
				api.SendMessage(hHeader, HDM_GETITEM, i, item);
				if (Name == api.SysAllocString(item.pszText)) {
					item.mask = HDI_FORMAT;
					item.fmt |= r ? HDF_SORTDOWN : HDF_SORTUP;
					api.SendMessage(hHeader, HDM_SETITEM, i, item);
					break;
				}
			}
		}
	} catch (e) { }
	FV.Parent.UnlockUpdate(true);
	Progress.StopProgressDialog();
}

function GetImgTag(o, h) {
	if (o.src) {
		var ar = ['<img'];
		for (var n in o) {
			if (o[n]) {
				ar.push(' ', n, '="', EncodeSC(api.PathUnquoteSpaces(o[n])), '"');
			}
		}
		if (h) {
			h = Number(h) ? h + 'px' : EncodeSC(h);
			ar.push(' width="', h, '" height="', h, '"');
		}
		ar.push('>');
		return ar.join("");
	}
	var ar = ['<span'];
	for (var n in o) {
		if (n != "title" && o[n]) {
			ar.push(' ', n, '="', EncodeSC(o[n]), '"');
		}
	}
	ar.push('>', EncodeSC(o.title), '</span>');
	return ar.join("");
}

function GetIconSize(h, a) {
	return h || a * screen.logicalYDPI / 96 || window.IconSize;
}

function MakeCommDlgFilter(arg) {
	var ar = arg ? arg.join ? arg : [arg] : [];
	var result = [];
	var bAll = true;
	for (var i = 0; i < ar.length; ++i) {
		var s = ar[i];
		bAll &= s.indexOf("*.*") < 0;
		if (/[\|#]/.test(s)) {
			result.push(s.replace(/[#@]/g, "|").replace(/[\0\|]$/, ""));
			continue;
		}
		var res = /\(([^\)]+)\)/.exec(s);
		if (res) {
			result.push(s, res[1]);
			continue;
		}
		var sfi = api.Memory("SHFILEINFO");
		api.SHGetFileInfo(s, 0, sfi, sfi.Size, SHGFI_TYPENAME);
		result.push(sfi.szTypeName + " (" + s + ")", s);
	}
	if (bAll) {
		result.push(api.LoadString(hShell32, 34193) || "All files", "*.*");
	}
	return result.join("|");
}

function GetHICON(iIcon, h, flags) {
	var ar = [SHIL_JUMBO, SHIL_EXTRALARGE, SHIL_LARGE, SHIL_SMALL]
	for (var i = ar.length; h > te.Data.SHILS[ar[--i]].cy && i;) { }
	return api.ImageList_GetIcon(te.Data.SHIL[ar[i]], iIcon, flags);
}

function GetBGRA (c, a) {
	return ((c & 0xff) << 16) | (c & 0xff00) | ((c & 0xff0000) >> 16) | a << 24;
}

function ExtractFilter(s) {
	return (ExtractMacro(te, s) || "").replace(/[\r\n;]+/g, ";").replace(/^;+|;+$|"/g, "");
}

function GetEnum(FolderItem, bShowHidden)
{
	if (FolderItem.Enum) {
		return FolderItem.Enum(FolderItem);
	}
	if (FolderItem.IsFolder) {
		var Folder = FolderItem.GetFolder;
		if (Folder) {
			var Items = Folder.Items();
			if (bShowHidden || api.GetKeyState(VK_SHIFT) < 0) {
				try {
					Items.Filter(SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN, "*");
				} catch (e) { }
			}
			return api.CreateObject("FolderItems", Items);
		}
	}
}

function LoadDBFromTSV(DB, fn)
{
	var ado = OpenAdodbFromTextFile(fn, "utf-8");
	if (ado) {
		while (!ado.EOS) {
			var res = /^([^\t]*)\t(.*)$/.exec(ado.ReadText(adReadLine));
			if (res) {
				DB.Set(res[1], res[2]);
			}
		}
		ado.Close();
	}
}

function SaveDBToTSV(DB, fn)
{
	try {
		var ado = api.CreateObject("ads");
		ado.CharSet = "utf-8";
		ado.Open();
		DB.ENumCB(function (n, s) {
			ado.WriteText([n, s].join("\t") + "\r\n");
		});
		ado.SaveToFile(fn, adSaveCreateOverWrite);
		ado.Close();
	} catch (e) { }
}

function BasicDB(name)
{
	this.DB = api.CreateObject("Object");

	this.Get = function (n) {
		return this.DB[n] || "";
	}

	this.Set = function (n, s) {
		var s0 = this.DB[n];
		if (s0 != s) {
			if (s) {
				this.DB[n] = s;
			} else {
				delete this.DB[n];
			}
			this.bChanged = true;
			var fn = api.ObjGetI(this, "OnChange");;
			fn && fn(n, s, s0);
		}
	}

	this.ENumCB = function (fncb) {
		for (var n in this.DB) {
			fncb(n, this.DB[n]);
		}
	}

	this.Clear = function () {
		for (var n in this.DB) {
			delete this.DB[n];
		}
		return this;
	}

	this.Load = function ()
	{
		LoadDBFromTSV(this, this.path);
		return this;
	}

	this.Save = function () {
		if (this.bChanged) {
			SaveDBToTSV(this, this.path);
			this.bChanged = false;
		}
	}

	this.Close = function () {}

	this.path = OrganizePath(name, fso.BuildPath(te.Data.DataFolder, "config")) + ".tsv";
}

SimpleDB = BasicDB;
