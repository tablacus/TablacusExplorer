// Tablacus Explorer

ui_ = {
	IEVer: window.chrome ? (/Edg\/(\d+)/.test(navigator.appVersion) ? RegExp.$1 : 12) : window.document && (document.documentMode || (/MSIE 6/.test(navigator.appVersion) ? 6 : 7)),
	bWindowRegistered: true,
	Panels: {},
	eventTE: {}
};

InitUI = function () {
	if (window.chrome) {
		te = parent.chrome.webview.hostObjects.te;
		api = te.WindowsAPI0.CreateObject("api");
		fso = api.CreateObject("fso");
		sha = api.CreateObject("sha");
		wsh = api.CreateObject("wsh");
		$ = api.CreateObject("Object");
		WebBrowser = te.Ctrl(CTRL_WB);
	}
	system32 = api.GetDisplayNameOf(ssfSYSTEM, SHGDN_FORPARSING);
	hShell32 = api.GetModuleHandle(BuildPath(system32, "shell32.dll"));

	osInfo = api.Memory("OSVERSIONINFOEX");
	osInfo.dwOSVersionInfoSize = osInfo.Size;
	api.GetVersionEx(osInfo);
	WINVER = osInfo.dwMajorVersion * 0x100 + osInfo.dwMinorVersion;

	if (WINVER > 0x603) {
		BUTTONS = {
			opened: '<b style="font-family: Consolas; transform: scale(1.2,1) rotate(-90deg)">&lt;</b>',
			closed: '<b style="font-family: Consolas; transform: scale(1,1.2) translateX(1px); opacity: 0.6">&gt;</b>',
			parent: '&laquo;',
			next: '<b style="font-family: Consolas; opacity: 0.6; transform: scale(0.75,0.9); text-shadow: 1px 0">&gt;</b>',
			dropdown: '<b style="font-family: Consolas; transform: scale(1.2,1) rotate(-90deg) translateX(2px); opacity: 0.6; width: 1em; display: inline-block">&lt;</b>'
		};
	} else {
		try {
			var s = wsh.regRead("HKCU\\Software\\Microsoft\\Internet Explorer\\Settings\\Always Use My Font Face");
		} catch (e) {
			s = 0;
		}
		BUTTONS = {
			opened: '<span style="font-size: 10pt; transform: translateY(-2pt)">&#x25e2;</span>',
			closed: '<span style="font-size: 10pt; transform: scale(1,1.4)">&#x25b7;</span>',
			parent: '&laquo;',
			next: s ? '&#x25ba;' : '<span style="font-family: Marlett">4</span>',
			dropdown: s ? '&#x25bc;' : '<span style="font-family: Marlett">6</span>'
		};
	}

	if (api.SHTestTokenMembership(null, 0x220) && WINVER >= 0x600) {
		TITLE += ' [' + (api.LoadString(hShell32, 25167) || "Admin").replace(/;.*$/, "") + ']';
	}

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
	if (x = ParentWindow) {
		if (x && x.$) {
			ParentWindow = x;
		}
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

	UI = api.CreateObject("Object");
	UI.Addons = api.CreateObject("Object");

	UI.OpenHttpRequest = OpenHttpRequest = function (url, alt, fn, arg) {
		var xhr = createHttpRequest();
		var fnLoaded = function () {
			if (fn) {
				if (arg && arg.pcRef) {
					arg.pcRef[0] = arg.pcRef[0] - 1;
				}
				if (xhr.status == 200) {
					fn(xhr, url, arg);
					fn = void 0;
					return;
				}
				if (/^http/.test(alt)) {
					UI.OpenHttpRequest(/^https/.test(url) && alt == "http" ? url.replace(/^https/, alt) : alt, '', fn, arg);
					return;
				}
				ShowXHRError(url, xhr.status);
			}
		}
		xhr.onload = fnLoaded;
		xhr.onreadystatechange = function () {
			if (xhr.readyState == 4) {
				fnLoaded();
			}
		}
		if (/ml$/i.test(url)) {
			url += "?" + Math.floor(new Date().getTime() / 60000);
		}
		if (arg && arg.pcRef) {
			arg.pcRef[0] = arg.pcRef[0] + 1;
		}
		if (window.chrome && /\.zip$|\.nupkg$/i.test(url)) {
			xhr.responseType = "blob";
		}
		xhr.open("GET", url, true);
		xhr.send();
	}

	UI.ReloadCustomize = ReloadCustomize = function () {
		te.Data.bReload = false;
		CloseSubWindows();
		Finalize();
		te.Reload();
		return S_OK;
	}

	UI.RunEvent = function () {
		var args = Array.apply(null, arguments);
		var s = args.shift();
		if (window.chrome) {
			s = "(() => {" + s + "\n})();";
		} else {
			s = RemoveAsync(s);
		}
		var fn = new Function(s);
		fn.apply(fn, args);
	}

	UI.setTimeoutAsync = function (fn, tm, a, b, c, d) {
		setTimeout(fn, tm, a, b, c, d);
	}

	UI.clearTimeout = clearTimeout;

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
		g_.dlgs.Options = ShowDialog("options.html", opt);
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
		var hr = MainWindow.RunEvent4("DownloadFile", url, fn);
		return hr != null ? hr : api.URLDownloadToFile(null, url, fn);
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
		arg.file = GetFileName((arg.url).replace(/\//g, "\\"));
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
			MessageBox([(api.LoadString(hShell32, 4228)).replace(/^\t/, "").replace("%d", api.sprintf(99, "0x%08x", hr)), GetText("Extract"), GetFileName(arg.zipfile)].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
			return;
		}
		var te_exe = arg.temp + "\\te64.exe";
		var nDog = 300;
		while (!$.fso.FileExists(te_exe)) {
			if (wsh.Popup(GetText("Please wait."), 1, TITLE, MB_OKCANCEL) == IDCANCEL || nDog-- == 0) {
				return;
			}
		}
		var arDel = [];
		var addons = arg.temp + "\\addons";
		if ($.fso.FolderExists(arg.temp + "\\config")) {
			arDel.push(arg.temp + "\\config");
		}
		for (var i = 32; i <= 64; i += 32) {
			te_exe = arg.temp + '\\te' + i + '.exe';
			var te_old = BuildPath(te.Data.Installed, 'te' + i + '.exe');
			if (!$.fso.FileExists(te_old) || $.fso.GetFileVersion(te_exe) == $.fso.GetFileVersion(te_old)) {
				arDel.push(te_exe);
			}
		}
		for (var list = api.CreateObject("Enum", $.fso.GetFolder(addons).SubFolders); !list.atEnd(); list.moveNext()) {
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

	UI.OnLoad = function () {
		AddEventEx(window, "beforeunload", CloseSubWindows);
	}

	UI.DialogResult = function (opt, returnValue) {
		var e = GetElementEx(opt.Id)
		if (e) {
			e.value = returnValue;
		}
	}

	UI.CloseWindow = CloseWindow = function () {
		if (window.chrome) {
			api.PostMessage(GetTopWindow(), WM_CLOSE, 0, 0);
			return;
		}
		window.close();
	}

	UI.Autocomplete = function (s, path) {
		var dl = document.getElementById("AddressList");
		while (dl.lastChild) {
			dl.removeChild(dl.lastChild);
		}
		if (g_.Autocomplete.Path == path) {
			var ar = s.split("\t");
			for (var i = 0; i < ar.length; ++i) {
				var el = document.createElement("option");
				el.value = ar[i];
				dl.appendChild(el);
			}
		}
	}
};

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

LoadScripts = function (js1, js2, cb) {
	InitUI();
	if (window.chrome) {
		js1.unshift("script/consts.js", "script/common.js", "script/syncb.js");
		var s = [];
		var arFN = ["fso", "sha", "wsh", "wnw"];
		var line = 1;
		var strParent = GetParentFolderName(api.GetModuleFileName(null));
		for (var i = 0; i < js1.length; i++) {
			var fn = BuildPath(strParent, js1[i]);
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
		screen.deviceYDPI = parent.deviceYDPI;
		ui_.Zoom = parent.deviceYDPI / 96;
		CopyObj($, window, ["te", "api", "chrome", "document", "UI", "MainWindow"]);
		$.location = CopyObj(null, location, ["hash", "href"]);
		$.navigator = CopyObj(null, navigator, ["appVersion", "language"]);
		$.screen = CopyObj(null, screen, ["deviceXDPI", "deviceYDPI"]);
		var o = api.CreateObject("Object");
		o.window = $;
		$.$JS = api.GetScriptDispatch(s.join(""), "JScript", o);
		CopyObj(window, $, ["g_", "Common", "Threads"]);
		CopyObj(window, $, arFN);
		var doc = api.CreateObject("Object");
		doc.parentWindow = $;
		WebBrowser.Document = doc;
	} else {
		$ = window;
		ui_.Zoom = 1;
	}
	Addons = {
		"_stack": api.CreateObject("Array")
	};
	LoadScript(js1.concat(js2), cb);
};

function _InvokeMethod() {
	var args;
	var ar = api.ObjGetI(te, "fn");
	while (args = ar.shift()) {
		api.OutputDebugString(["InvokeMethod:", api.ObjGetI(ar, "length"), "\n"].join(","));
		var fn = args.shift();
		api.OutputDebugString(["InvokeMethod:", fn.toString().split("\n")[0]].join(","));
		var hr = fn.apply(fn, args);
		api.OutputDebugString("InvokeMethod: " + hr + "," + api.ObjGetI(ar, "length") + "\n");
	}
}

ApplyLangTag = function (o) {
	if (o) {
		for (var i = 0; i < o.length; ++i) {
			(function (el) {
				var s;
				if (s = el.childNodes) {
					for (var j = s.length; j-- > 0;) {
						if (!s[j].tagName) {
							s[j].data = amp2ul(GetTextR(s[j].data.replace(/&amp;/ig, "&")));
						}
					}
				}
				if (s = el.title) {
					el.title = (GetTextR(s)).replace(/\(&\w\)|&/, "");
				}
				if (s = el.alt) {
					el.alt = (GetTextR(s)).replace(/\(&\w\)|&/, "");
				}
			})(o[i]);
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
			(function (el) {
				if (!h && SameText(el.type, "text")) {
					h = el.offsetHeight;
				}
				if (s = el.placeholder) {
					el.placeholder = (GetTextR(s)).replace(/\(&\w\)|&/, "");
				}
				if (SameText(el.type, "button")) {
					if (s = el.value) {
						el.value = (GetTextR(s)).replace(/\(&\w\)|&/, "");
					}
				}
				if (s = ImgBase64(el, 0)) {
					el.src = s;
					if (SameText(el.type, "text")) {
						el.style.backgroundImage = "url('" + s + "')";
					}
				}
			})(o[i]);
		}
	}
	o = doc.getElementsByTagName("img");
	if (o) {
		ApplyLangTag(o);
		for (i = o.length; i--;) {
			(function (el) {
				var s = ImgBase64(el, 0);
				if (s) {
					el.src = s;
				}
				if (!el.ondragstart) {
					el.draggable = false;
				}
			})(o[i]);
		}
	}
	o = doc.getElementsByTagName("select");
	if (o) {
		for (i = o.length; i--;) {
			(function (el) {
				el.title = delamp(GetTextR(el.title));
				for (var j = 0; j < el.length; ++j) {
					el[j].text = (GetTextR(el[j].text)).replace(/^\n/, "").replace(/\n$/, "");
				}
			})(o[i]);
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

ImgBase64 = function (el, index, h) {
	var org = el.src || el.getAttribute("data-src");
	var src = ExtractMacro(te, org);
	var s = MakeImgSrc(src, index, false, h || el.height);
	if ("string" === typeof s && org != src) {
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

GetRect = function (o, f) {
	var rc = api.Memory("RECT");
	var pt = GetPos(o, f);
	rc.left = pt.x;
	rc.top = pt.y;
	rc.right = pt.x + o.offsetWidth;
	rc.bottom = pt.y + o.offsetHeight;
	return rc;
}

GetPos = function (o, bScreen, bAbs, bPanel, bBottom) {
	if ("number" === typeof bScreen) {
		bAbs = bScreen & 2;
		bPanel = bScreen & 4;
		bBottom = bScreen & 8;
		bScreen &= 1;
	}
	var x = bScreen ? screenLeft * ui_.Zoom : 0;
	var y = bScreen ? screenTop * ui_.Zoom : 0;
	if (bBottom) {
		y += o.offsetHeight;
	}
	var rc = o.getBoundingClientRect();
	return { x: x + rc.left, y: y + rc.top };
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

GetTopWindow = function (hwnd) {
	var hwnd1 = hwnd || WebBrowser.hwnd;
	while (hwnd1 = api.GetParent(hwnd1)) {
		hwnd = hwnd1;
	}
	return hwnd;
}

CloseSubWindows = function () {
	var hwnd = GetTopWindow();;
	var hwnd1 = hwnd;
	while (hwnd1 = api.FindWindowEx(null, hwnd1, null, null)) {
		if (hwnd == api.GetWindowLongPtr(hwnd1, GWLP_HWNDPARENT)) {
			api.PostMessage(hwnd1, WM_CLOSE, 0, 0);
		}
	}
}

MouseOver = function (o) {
	o = o.srcElement || o;
	if (/^button$|^menu$/i.test(o.className)) {
		if (ui_.objHover && o != ui_.objHover) {
			MouseOut();
		}
		var bHover = window.chrome;
		if (!bHover) {
			var pt = api.Memory("POINT");
			api.GetCursorPos(pt);
			var ptc = pt.Clone();
			api.ScreenToClient(WebBrowser.hwnd, ptc);
			bHover = (o == document.elementFromPoint(ptc.x, ptc.y) || HitTest(o, pt));
		}
		if (bHover) {
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
	var el = e.srcElement;
	return /input|textarea/i.test(el.tagName) || /selectable/i.test(el.className);
}

GetFolderViewEx = function (Ctrl, pt, bStrict) {
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
			o = o.offsetParent;
		}
		return te.Ctrl(CTRL_FV);
	}
	return GetFolderView(Ctrl, pt, bStrict);
}

AddEventEx(window, "load", function () {
	document.body.onselectstart = DetectProcessTag;
	if (!window.chrome) {
		document.body.oncontextmenu = DetectProcessTag;
	}
	if (window.chrome) {
		document.body.addEventListener('mousewheel', function (e) {
			if (e.ctrlKey) {
				e.preventDefault();
			}
		}, { passive: false });
	} else {
		document.body.onmousewheel = function (e) {
			return !e.ctrlKey;
		};
	}
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
		o.src = ImgBase64(o, 0, Number(h))
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
		var s = ReadTextFile(fn);
		if (s) {
			if (ar[1] == "js") {
				if (window.chrome) {
					s = "(() => {" + s + "\n})();";
				} else {
					s = RemoveAsync(s);
				}
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
	} catch (e) {
		arError.push([e.stack || e.description || e.toString(), fn].join("\n"));
	}
	return r;
}

//Options
AddonOptions = function (Id, fn, Data, bNew) {
	LoadLang2(BuildPath("addons", Id, "lang", GetLangId() + ".xml"));
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
			MainWindow.g_.dlgs[Id] = void 0;
		}
		opt = api.CreateObject("Object");
		opt.MainWindow = $;
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
		opt.event.onbeforeunload = "MainWindow.g_.dlgs['" +Id +"'] = void 0;";
		MainWindow.g_.dlgs[Id] = ShowDialog(sURL, opt);
		return;
	}
	if (!ui_.elAddons[Id]) {
		if (!/location\.html$/.test(sURL)) {
			opt.event.onload = function () {
				var cInput = el.contentWindow.document.getElementsByTagName('input');
				for (var i in cInput) {
					if (/^ok$|^cancel$/i.test(cInput[i].className)) {
						cInput[i].style.display = 'none';
					}
				}
				el.contentWindow.g_Inline = true;
			}
		}
		te.Arguments = opt;
		var el = document.createElement('iframe');
		el.id = 'panel1_' + Id;
		if (window.chrome) {
			el.srcdoc = (ReadTextFile(sURL)).replace(/(<head[^>]*>)/i, '$1<base href="' + sURL.replace(/[^\/\\]*$/, "") + '">');
		} else {
			el.src = sURL;
		}
		el.style.cssText = 'width: 100%; border: 0; padding: 0; margin: 0';
		ui_.elAddons[Id] = el;
		var o = document.getElementById('panel1_2');
		o.style.display = "block";
		o.appendChild(el);
		o = document.getElementById('tab1_');
		o.insertAdjacentHTML("BeforeEnd", '<label id="tab1_' + Id + '" class="button" style="width: 100%" onmousedown="ClickTree(this);">' + info.Name + '</label><br>');
	}
	ClickTree(document.getElementById('tab1_' + Id));
}

GetElementEx = function (Id) {
	var e = document.getElementById(Id);
	if (e) {
		return e;
	}
	var ar = Id.split("::");
	return document.forms[ar[0]].elements[ar[1]];
}

GetElementIdEx = function (e) {
	if ("string" === typeof e) {
		return e;
	}
	if (e.id) {
		return e.id;
	}
	return e.form.name + "::" + e.name;
}

InputMouse = function (o) {
	ShowDialogEx("mouse", 500, 420, GetElementIdEx(o || (document.E && document.E.MouseMouse) || document.F.MouseMouse || document.F.Mouse));
}

InputKey = function (o) {
	ShowDialogEx("key", 320, 120, GetElementIdEx(o || (document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key));
}

ShowIconEx = function (o, mode) {
	ShowDialogEx("icon" + (mode || ""), 640, 480, GetElementIdEx(o || document.F.Icon));
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
		return a > b ? 1 : a < b ? -1 : 0;
	});
	var j = "";
	for (i in s) {
		if (j != s[i]) {
			j = s[i];
			var o = oa[++oa.length - 1];
			o.value = j;
			o.text = j + "\x20";
		}
	}
}

SetKeyShift = function () {
	var key = ((document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key).value;
	key = key.replace(/^(.+),.+/, "$1");
	var nLen = GetLength(MainWindow.g_.KeyState);
	for (var i = 0; i < nLen; ++i) {
		var s = MainWindow.g_.KeyState[i][0];
		var o = document.getElementById("_Key" + s);
		if (o) {
			o.checked = key.indexOf(s + "+") >= 0;
		}
		key = key.replace(s + "+", "");
	}
	o = document.getElementById("_KeySelect");
	for (var i = o.length; i--;) {
		if (SameText(key, o[i].value)) {
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
	var ar = oKey.value == "," ? [","] : oKey.value.split(",");
	ar[0] = ar[0].replace(/(\+)[^\+]*$|^[^\+]*$/, "$1") + o[o.selectedIndex].value;
	oKey.value = ar.join(",");
}

createHttpRequest = function () {
	var xhr = MainWindow.RunEvent4("createHttpRequest");
	if (xhr != null) {
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

ClearAutocomplete = function () {
	for (var dl = document.getElementById("AddressList"); dl.lastChild;) {
		dl.removeChild(dl.lastChild);
	}
	g_.Autocomplete.Path = "";
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

