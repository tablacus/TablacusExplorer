//Tablacus Explorer

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
eventTE = api.CreateObject("Object");
eventTE.Environment = api.CreateObject("Object");
eventTA = api.CreateObject("Object");

g_ptDrag = api.Memory("POINT");
Addons = api.CreateObject("Object");
Addons["_stack"] = api.CreateObject("Array");

g_ = api.CreateObject("Object");
g_.Colors = api.CreateObject("Object")
g_.KeyCode = api.CreateObject("Object");
g_.KeyState = api.CreateObject("Array");
var ar = [
	[0x1d0000, 0x2000],
	[0x2a0000, 0x1000],
	[0x380000, 0x4000],
	["Win", 0x8000],
	["Ctrl", 0x2000],
	["Shift", 0x1000],
	["Alt", 0x4000]
];
for (var i in ar) {
	var a2 = api.CreateObject("Array");
	a2.push(ar[i][0], ar[i][1]);
	g_.KeyState.push(a2);
}
g_.stack_TC = api.CreateObject("Object");
g_.dlgs = api.CreateObject("Object");
g_.tidWindowRegistered = null;
g_.bWindowRegistered = true;
g_.xmlWindow = null;
g_.elAddons = api.CreateObject("Object");
g_.event = api.CreateObject("Object");
g_.tid_rf = api.CreateObject("Array");
g_.Autocomplete = api.CreateObject("Object");
g_.LockUpdate = 0;
g_.ptMenuDrag = api.Memory("POINT");
g_.Locations = api.CreateObject("Object");
g_.IEVer = window.chrome ? (/Edg\/(\d+)/.test(navigator.appVersion) ? RegExp.$1 : 12) : window.document && (document.documentMode || (/MSIE 6/.test(navigator.appVersion) ? 6 : 7));

if (g_.IEVer < 10) {
	(function (f) {
		if (f) {
			setTimeout = function (fn, tm) {
				var args = Array.prototype.slice.call(arguments, 2);
				try {
					return f(function () {
						try {
							if ("string" === typeof fn) {
								fn = new Function(fn);
							}
							fn.apply(fn, args);
						} catch (e) {
							ShowError(e, fn.toString());
						}
					}, tm);
				} catch (e) {
					ShowError(e);
				}
			}
		}
	})(setTimeout);
} else if (g_.IEVer > 11) {
	setTimeout = function (fn, tm) {
		api.OutputDebugString(fn + "\n");
	}

	clearTimeout = function (tid) {
		api.Invoke(UI.clearTimeout, tid);
	}
}

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
			eventTA[en] = api.CreateObject("Array");
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

ClearEvent = function (Name) {
	if (Name) {
		var en = Name.toLowerCase();
		delete eventTE[en];
		delete eventTA[en];
	}
}

AddEnv = function (Name, fn) {
	eventTE.Environment[Name.toLowerCase()] = fn;
}

AddEventId = function (Name, Id, fn) {
	var en = Name.toLowerCase();
	if (!eventTE[en]) {
		eventTE[en] = api.CreateObject("Object");
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
			api.Invoke(UI.AddEventEx, null, "beforeunload", fn);
		}
	}
	CollectGarbage();
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
					var FV = TC.Selected;
					FV.Navigate2(Path, SBSP_NEWBROWSER, tab.getAttribute("Type"), tab.getAttribute("ViewMode"), tab.getAttribute("FolderFlags"), tab.getAttribute("Options"), tab.getAttribute("ViewFlags"), tab.getAttribute("IconSize"), tab.getAttribute("Align"), tab.getAttribute("Width"), tab.getAttribute("Flags"), tab.getAttribute("EnumFlags"), tab.getAttribute("RootStyle"), tab.getAttribute("Root"));
					if (!FV.FilterView) {
						FV.FilterView = tab.getAttribute("FilterView");
					}
					FV.Data.Lock = api.LowPart(tab.getAttribute("Lock")) != 0;
					Lock(TC, i2, false);
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

GetAddons = function () {
	ShowOptions("Tab=Get Addons");
}

GetIconPacks = function () {
	ShowOptions("Tab=Get Icons");
}

CheckUpdate = function (arg) {
	MainWindow.OpenHttpRequest("https://api.github.com/repos/tablacus/TablacusExplorer/releases/latest", "http://tablacus.github.io/TablacusExplorerAddons/te/releases.json", CheckUpdate2, arg);
}

CheckUpdate2 = function (xhr, url, arg1) {
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

CheckUpdate3 = function (xhr, url, arg) {
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
	for (var list = api.CreateObject("Enum", fso.GetFolder(addons).SubFolders); !list.atEnd(); list.moveNext()) {
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

ShowAbout = function () {
	ShowDialog(fso.BuildPath(te.Data.Installed, "script\\dialog.html"), { MainWindow: MainWindow, Query: "about", Modal: false, width: 640, height: 360 });
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

AddEvent("ConfigChanged", function (s) {
	te.Data["bSave" + s] = true;
});

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


AddonBeforeRemove = function (Id) {
	CollectGarbage();
	var arError = [];
	var r = LoadAddon("remove.js", Id, arError);
	if (arError.length) {
		MessageBox(arError.join("\n\n"), TITLE, MB_OK);
	}
	return r;
}

CreateSyncFunction = function (s) {
	return new Function(s);
}

BlurId = function (Id) {
	api.Invoke(UI.Blur, id);
}

ShowError = function (e, s, i) {
	api.Invoke(UI.ShowError, e, s, i);
}
