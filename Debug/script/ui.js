// Tablacus Explorer

ui_ = {
	IEVer: window.chrome ? (/Edg\/(\d+)/.test(navigator.appVersion) ? RegExp.$1 : 12) : window.document && (document.documentMode || (/MSIE 6/.test(navigator.appVersion) ? 6 : 7)),
	bWindowRegistered: true,
	Panels: {}
};

UI = api.CreateObject("Object");

UI.OpenHttpRequest = OpenHttpRequest = function (url, alt, fn, arg) {
	var xhr = createHttpRequest();
	var fnLoaded = function () {
		if (arg && arg.pcRef) {
			arg.pcRef[0] = arg.pcRef[0] - 1;
		}
		if (xhr.status == 200) {
			return fn(xhr, url, arg);
		}
		if (/^http/.test(alt)) {
			return UI.OpenHttpRequest(/^https/.test(url) && alt == "http" ? url.replace(/^https/, alt) : alt, '', fn, arg);
		}
		ShowXHRError(url, xhr.status);
	}
	if (xhr.onload !== void 0) {
		xhr.onload = fnLoaded;
	} else {
		xhr.onreadystatechange = function () {
			if (xhr.readyState == 4) {
				fnLoaded();
			}
		}
	}
	if (/ml$/i.test(url)) {
		url += "?" + Math.floor(new Date().getTime() / 60000);
	}
	if (arg && arg.pcRef) {
		arg.pcRef[0] = arg.pcRef[0] + 1;
	}
	if (window.chrome && /\.zip$/i.test(url)) {
		xhr.responseType = "blob";
	}
	xhr.open("GET", url, true);
	xhr.send();
}

UI.setTimeoutAsync = function (fn, tm, a, b, c, d) {
	setTimeout(fn, tm, a, b, c, d);
}

UI.clearTimeout = clearTimeout;

UI.RunEvent1 = function () {
	var args = Array.apply(null, arguments);
	var en = args.shift();
	var eo = eventTE[en.toLowerCase()];
	var nLen = GetLength(eo);
	for (var i = 0; i < nLen; ++i) {
		var fn = eo[i];
		try {
			fn.apply(fn, args);
		} catch (e) {
			ShowError(e, en, i);
		}
	}
}

UI.ShowError = function (cb, e, s) {
	setTimeout(cb, 99, cb, e, s);
}

UI.SetDisplay = function (Id, s) {
	var o = document.getElementById(Id);
	if (o) {
		o.style.display = s;
	}
}

UI.BlurId = BlurId = function (Id) {
	document.getElementById(Id).blur();
}

UI.ShowOptions = function (opt) {
	opt.event.onbeforeunload = function () {
		MainWindow.g_.dlgs.Options = void 0;
	};
	g_.dlgs.Options = ShowDialog("options.html", opt);
}

function readAsDataURL(blob) {
	return new Promise(function (resolve, reject) {
		var reader = new FileReader();
		reader.onload = function () {
			resolve(reader.result);
		};
		reader.onerror = function () {
			reject(reader.error);
		};
		reader.readAsDataURL(blob);
	});
}

UI.Extract = Extract = function (Src, Dest, xhr) {
	var hr;
	if (xhr) {
		if (window.chrome && xhr.response) {
			xhr = readAsDataURL(xhr.response);
		}
		hr = DownloadFile(xhr, Src);
		if (hr) {
			return hr;
		}
	}
	var eo = MainWindow.eventTE.extract;
	var nLen = GetLength(eo);
	for (var i = 0; i < nLen; ++i) {
		var fn = eo[i];
		try {
			fn.apply(Src, Dest);
		} catch (e) {
			ShowError(e, "Extract", i);
		}
	}
	return api.Extract(BuildPath(system32, "zipfldr.dll"), "{E88DCCE0-B7B3-11d1-A9F0-00AA0060FA31}", Src, Dest);
}

UI.DownloadFile = DownloadFile = function (url, fn) {
	return api.URLDownloadToFile(null, url, fn);
}

UI.CheckUpdate2 = function (xhr, url, arg1) {
	var arg = api.CreateObject("Object");
	var Text = xhr.get_responseText ? xhr.get_responseText() : xhr.responseText;
	var json = window.JSON ? JSON.parse(Text) : api.GetScriptDispatch("function fn () { return " + Text + "}", "JScript", {}).fn();
	if (json.assets && json.assets[0]) {
		arg.size = json.assets[0].size / 1024;
		arg.url = json.assets[0].browser_download_url;
	}
	if (!arg.url) {
		return;
	}
	arg.file = fso.GetFileName((arg.url).replace(/\//g, "\\"));
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
	arg.temp = BuildPath(wsh.ExpandEnvironmentStrings("%TEMP%"), "tablacus");
	CreateFolder2(arg.temp);
	arg.InstalledFolder = te.Data.Installed;
	arg.zipfile = BuildPath(arg.temp, arg.file);
	arg.temp = arg.temp + "\\explorer";
	DeleteItem(arg.temp);
	CreateFolder2(arg.temp);
	OpenHttpRequest(arg.url, "http://tablacus.github.io/TablacusExplorerAddons/te/" + ((arg.url).replace(/^.*\//, "")), UI.CheckUpdate3, arg);
}

UI.CheckUpdate3 = function (xhr, url, arg) {
	var hr = Extract(arg.zipfile, arg.temp, xhr);
	if (hr) {
		MessageBox([(api.LoadString(hShell32, 4228)).replace(/^\t/, "").replace("%d", api.sprintf(99, "0x%08x", hr)), GetText("Extract"), fso.GetFileName(arg.zipfile)].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
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
		var te_old = BuildPath(te.Data.Installed, 'te' + i + '.exe');
		if (!fso.FileExists(te_old) || fso.GetFileVersion(te_exe) == fso.GetFileVersion(te_old)) {
			arDel.push(te_exe);
		}
	}
	for (var list = api.CreateObject("Enum", fso.GetFolder(addons).SubFolders); !list.atEnd(); list.moveNext()) {
		var n = list.item().Name;
		var items = te.Data.Addons.getElementsByTagName(n);
		if (!items || GetLength(items) == 0) {
			arDel.push(BuildPath(addons, n));
		}
	}
	if (arDel.length) {
		api.SHFileOperation(FO_DELETE, arDel, null, FOF_SILENT | FOF_NOCONFIRMATION, false);
	}
	var ppid = api.Memory("DWORD");
	api.GetWindowThreadProcessId(te.hwnd, ppid);
	arg.pid = ppid[0];
	MainWindow.CreateUpdater(arg);
}

UI.GetIconSize = GetIconSize = function (h, a) {
	return h || a * screen.logicalYDPI / 96 || window.IconSize;
}

InitUI = function () {
	ui_.DoubleClickTime = sha.GetSystemInformation("DoubleClickTime");
	var arg = te.Arguments;
	if (arg) {
		window.dialogArguments = arg;
		$.dialogArguments = arg;
	}
	if (window.dialogArguments) {
		for (var list = api.CreateObject("Enum", dialogArguments.event); !list.atEnd(); list.moveNext()) {
			var j = list.item();
			var res = /^on(.+)/i.exec(j);
			if (res) {
				AddEventEx(window, res[1], dialogArguments.event[j]);
			} else {
				window[j] = dialogArguments.event[j];
			}
		}
	}
	var x;
	if (!window.MainWindow) {
		window.MainWindow = $;
		while (x = MainWindow.dialogArguments || MainWindow.opener) {
			MainWindow = x;
			if (x = MainWindow.MainWindow) {
				MainWindow = x;
			}
		}
		if (x = MainWindow.$) {
			MainWindow = x;
		}
	}
	window.ParentWindow = (window.dialogArguments ? dialogArguments.opener : window.opener);
	if (ParentWindow || !window.te) {
		te = MainWindow.te;
	}
	if (x = ParentWindow && ParentWindow.$) {
		ParentWindow = x;
	}
	var uid = location.hash.replace(/\D/g, "");
	if (!window.dialogArguments && !window.opener) {
		var arg = api.CommandLineToArgv(api.GetCommandLine());
		if (arg.Count > 3 && SameText(arg[1], '/open')) {
			uid = arg[3];
		}
	}
	if (uid) {
		for (var esw = api.CreateObject("Enum", sha.Windows()); !esw.atEnd(); esw.moveNext()) {
			var x = esw.item();
			if (x && x.Document) {
				var w = x.Document.parentWindow;
				if (w && w.te && w.Exchange) {
					var a = w.Exchange[uid];
					if (a) {
						dialogArguments = a;
						MainWindow = w;
						te = w.te;
						AddEventEx(window, "beforeunload", function () {
							try {
								delete MainWindow.Exchange[uid];
							} catch (e) { }
						});
						break;
					}
				}
			}
		}
	}
};

LoadScripts = function (js1, js2, cb) {
	InitUI();
	if (window.chrome) {
		js1.unshift("consts.js", "common.js", "syncb.js");
		var s = [];
		var arFN = [];
		var line = 1;
		var strParent = (fso.GetParentFolderName(api.GetModuleFileName(null))).replace(/\\$/, "");
		for (var i = 0; i < js1.length; i++) {
			var fn = strParent + "\\script\\" + js1[i];
			var ado = api.CreateObject("ads");
			if (ado) {
				ado.CharSet = "utf-8";
				ado.Open();
				ado.LoadFromFile(fn);
				var src = RemoveAsync(ado.ReadText());
				s.push(src);
				ado.Close();
				if (src && /sync/.test(fn)) {
					arFN = arFN.concat(src.match(/\s\w+\s*=\s*function/g).map(function (s) {
						return s.replace(/^\s|\s*=.*$/g, "");
					}));
				}
				api.OutputDebugString([js1[i], "Start line:", line, "\n"].join(" "));
				line += src.split("\n").length;
			}
		}
		js1.length = 0;
		var CopyObj = function (to, o, ar) {
			if (!to) {
				to = api.CreateObject("Object");
			}
			ar.forEach(function (key) {
				var a = o[key];
				if (!to[key] && a !== void 0) {
					to[key] = a;
				}
			});
			return to;
		}
		document.documentMode = (/Edg\/(\d+)/.test(navigator.appVersion) ? RegExp.$1 : 12);
		CopyObj($, window, ["te", "api", "chrome", "document", "UI", "MainWindow"]);
		$.location = CopyObj(null, location, ["hash", "href"]);
		$.navigator = CopyObj(null, navigator, ["appVersion", "language"]);
		$.screen = CopyObj(null, screen, ["deviceXDPI", "deviceYDPI"]);
		var o = api.CreateObject("Object");
		o.window = $;
		$.$JS = api.GetScriptDispatch(s.join(""), "JScript", o);
		CopyObj(window, $, ["g_", "Addons", "eventTE", "eventTA", "Threads", "SimpleDB", "BasicDB;"]);
		CopyObj(window, $, arFN);
		CopyObj($, window, ["fso", "wsh"]);
		$.g_.hwndBrowser = GetBrowserWindow();
	} else {
		$ = window;
	}
	LoadScript(js1.concat(js2), cb);
};

ApplyLangTag = function (o) {
	if (o) {
		for (var i = o.length; i--;) {
			var s, s1;
			if (s = o[i].innerHTML) {
				var ar = s.split("<");
				var re1 = /^([^>]*>\s*)(.+)(\s*)$/;
				var re2 = /^(\s*)(.+)(\s*)$/;
				for (var j = ar.length; j-- > 0;) {
					var res = re1.exec(ar[j]) || re2.exec(ar[j])
					if (res) {
						ar[j] = res[1] + amp2ul(GetTextR(res[2].replace(/&amp;/ig, "&"))) + res[3];
					}
				}
				s1 = ar.join("<");
				if (s != s1) {
					o[i].innerHTML = s1;
				}
			}
			s = o[i].title;
			if (s) {
				o[i].title = (GetTextR(s)).replace(/\(&\w\)|&/, "");
			}
			s = o[i].alt;
			if (s) {
				o[i].alt = (GetTextR(s)).replace(/\(&\w\)|&/, "");
			}
		}
	}
}

ApplyLang = function (doc) {
	var i, o, s, h = 0;
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
			if (s = o[i].placeholder) {
				o[i].placeholder = (GetTextR(s)).replace(/\(&\w\)|&/, "");
			}
			if (o[i].type == "button") {
				if (s = o[i].value) {
					o[i].value = (GetTextR(s)).replace(/\(&\w\)|&/, "");
				}
			}
			if (s = ImgBase64(o[i], 0)) {
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
				var s = GetTextR(o[i][j].text);
				o[i][j].text = s.replace(/^\n/, "").replace(/\n$/, "");
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
}

ImgBase64 = function (el, index) {
	var src = ExtractMacro(te, el.src);
	var s = MakeImgSrc(src, index, false, el.height);
	if (!s && !SameText(el.src, src)) {
		return src.replace(location.href.replace(/[^\/]*$/, ""), "file:///");
	}
	return s;
}

FireEvent = function (o, event) {
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
	return { x: x, y: y };
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

GetBrowserWindow = function () {
	return (window.chrome ? chrome.webview.hostObjects.hwnd : api.GetWindow(document));
}

GetTopWindow = function (hwnd) {
	var hwnd1 = hwnd || GetBrowserWindow();
	while (hwnd1 = api.GetParent(hwnd1)) {
		hwnd = hwnd1;
	}
	return hwnd;
}

MouseOver = function (o) {
	o = o.srcElement || o;
	if (/^button$|^menu$/i.test(o.className)) {
		if (ui_.objHover && o != ui_.objHover) {
			MouseOut();
		}
		var pt = api.Memory("POINT");
		api.GetCursorPos(pt);
		var ptc = pt.Clone();
		api.ScreenToClient(GetBrowserWindow(), ptc);
		if (o == document.elementFromPoint(ptc.x, ptc.y) || HitTest(o, pt)) {
			ui_.objHover = o;
			o.className = 'hover' + o.className;
		}
	}
}

MouseOut = function (s) {
	if (ui_.objHover) {
		if ("string" !== typeof s || ui_.objHover.id.indexOf(s) >= 0) {
			if (ui_.objHover.className == 'hoverbutton') {
				ui_.objHover.className = 'button';
			} else if (ui_.objHover.className == 'hovermenu') {
				ui_.objHover.className = 'menu';
			}
			ui_.objHover = null;
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
		ot.value = s.slice(0, i) + "\t" + s.slice(i);
		ot.selectionStart = ++i;
		ot.selectionEnd = i;
		return false;
	}
	return true;
}

DetectProcessTag = function (e) {
	var el = (e || event).srcElement;
	return /input|textarea/i.test(el.tagName) || /selectable/i.test(el.className);
}

AddEventEx(window, "load", function () {
	document.body.onselectstart = DetectProcessTag;
	if (!window.chrome) {
		document.body.oncontextmenu = DetectProcessTag;
	}
	document.body.onmousewheel = function () {
		return api.GetKeyState(VK_CONTROL) >= 0;
	};
});

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

GetImgTag = function (o, h) {
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

LoadAddon = function (ext, Id, arError, param) {
	var r;
	try {
		var sc;
		var ar = ext.split(".");
		if (ar.length == 1) {
			ar.unshift("script");
		}
		var fn = BuildPath("addons", Id, ar.join("."));
		var ado = OpenAdodbFromTextFile(fn, "utf-8");
		if (ado) {
			var s = ado.ReadText();
			ado.Close();
			if (s) {
				if (ar[1] == "js") {
					sc = new Function(s);
				} else if (ar[1] == "vbs") {
					var o = api.CreateObject("Object");
					o["_Addon_Id"] = api.CreateObject("Object");
					o["_Addon_Id"].Addon_Id = Id;
					o.window = window;
					sc = ExecAddonScript("VBScript", s, fn, arError, o, Addons["_stack"]);
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

UI.OnLoad = function () {
	UI.Addons = api.CreateObject("Object");
	AddEventEx(window, "beforeunload", CloseSubWindows);
}

//Options
AddonOptions = function (Id, fn, Data, bNew) {
	var sParent = te.Data.Installed;
	LoadLang2(BuildPath(sParent, "addons\\" + Id + "\\lang\\" + GetLangId() + ".xml"));
	var items = te.Data.Addons.getElementsByTagName(Id);
	if (!GetLength(items)) {
		var root = te.Data.Addons.documentElement;
		if (root) {
			root.appendChild(te.Data.Addons.createElement(Id));
		}
	}
	var info = GetAddonInfo(Id);
	var sURL = "addons\\" + Id + "\\options.html";
	if (!Data) {
		Data = api.CreateObject("Object");;
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
	sURL = BuildPath(te.Data.Installed, sURL);
	var opt = api.CreateObject("Object");
	opt.MainWindow = MainWindow.$;
	opt.Data = Data;
	opt.event = api.CreateObject("Object");
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
		opt = api.CreateObject("Object");
		opt.MainWindow = MainWindow;
		opt.Data = Data;
		opt.event = api.CreateObject("Object");
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
			MainWindow.g_.dlgs[Id] = void 0;
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
		te.Arguments = opt;
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
	var opt = api.CreateObject("Object");
	opt.MainWindow = MainWindow;
	opt.Data = s;
	ShowDialog(BuildPath(te.Data.Installed, "script\\location.html"), opt);
}

MakeKeySelect = function () {
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

SetKeyShift = function () {
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

CalcElementHeight = function (o, em) {
	if (o) {
		if (ui_.IEVer >= 9) {
			o.style.height = "calc(100vh - " + em + "em)";
			return;
		}
		var h = document.documentElement.clientHeight || document.body.clientHeight;
		h += MainWindow.DefaultFont.lfHeight * em;
		if (h > 0) {
			o.style.height = h + 'px';
			o.style.height = 2 * h - o.offsetHeight + "px";
		}
	}
}

KeyShift = function (o) {
	var oKey = (document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key;
	var key = oKey.value;
	var shift = o.id.replace(/^_Key(.*)/, "$1+");
	key = key.replace(shift, "");
	if (o.checked) {
		key = shift + key;
	}
	oKey.value = key;
}

KeySelect = function (o) {
	var oKey = (document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key;
	oKey.value = oKey.value.replace(/(\+)[^\+]*$|^[^\+]*$/, "$1") + o[o.selectedIndex].value;
}

createHttpRequest = function () {
	var xhr = RunEvent4("createHttpRequest");
	if (xhr !== void 0) {
		return xhr;
	}
	try {
		return window.XMLHttpRequest && ui_.IEVer >= 9 ? new XMLHttpRequest() : api.CreateObject("Msxml2.XMLHTTP");
	} catch (e) {
		return api.CreateObject("Microsoft.XMLHTTP");
	}
}

ShowXHRError = function (url, status) {
	MessageBox([api.sprintf(999, (api.LoadString(hShell32, 4227)).replace(/^\t/, ""), status), url].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
}

if (!window.JSON) {
	JSON = {
		parse: function (s) {
			return new Function('return ' + (s || {}))();
		},
		stringify: function (o) {
			var ar = [];
			if (Array.isArray(o)) {
				for (var i = 0; i < o.length; ++i) {
					if ("object" === typeof o[i]) {
						ar.push('"' + this.stringify(o[i]) + '"');
					} else {
						ar.push('"' + o[i] + '"');
					}
				}
				return '[' + ar.join(",") + "]";
			}
			for (var n in o) {
				if ("object" === typeof o[n]) {
					ar.push('"' + n + '":"' + this.stringify(o[n]) + '"');
				} else {
					ar.push('"' + n + '":"' + o[n] + '"');
				}
			}
			return '{' + ar.join(",") + "}";
		}
	}
}

