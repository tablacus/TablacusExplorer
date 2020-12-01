// Tablacus Explorer

ui_ = {
	IEVer: window.chrome ? (/Edg\/(\d+)/.test(navigator.appVersion) ? RegExp.$1 : 12) : window.document && (document.documentMode || (/MSIE 6/.test(navigator.appVersion) ? 6 : 7)),
	bWindowRegistered: true,
	Zoom: 1,
	eventTE: {}
};

InitUI = async function () {
	if (window.chrome) {
		te = await parent.chrome.webview.hostObjects.te;
		api = await te.WindowsAPI0.CreateObject("api");
		fso = await api.CreateObject("fso");
		sha = await api.CreateObject("sha");
		wsh = await api.CreateObject("wsh");
		$ = await api.CreateObject("Object");
		WebBrowser = await te.Ctrl(CTRL_WB);
	}
	system32 = await api.GetDisplayNameOf(ssfSYSTEM, SHGDN_FORPARSING);
	hShell32 = await api.GetModuleHandle(BuildPath(system32, "shell32.dll"));

	osInfo = await api.Memory("OSVERSIONINFOEX");
	osInfo.dwOSVersionInfoSize = await osInfo.Size;
	await api.GetVersionEx(osInfo);
	WINVER = await osInfo.dwMajorVersion * 0x100 + await osInfo.dwMinorVersion;

	if (WINVER > 0x603) {
		BUTTONS = {
			opened: '<b style="font-family: Consolas; transform: scale(1.2,1) rotate(-90deg); display: inline-block">&lt;</b>',
			closed: '<b style="font-family: Consolas; transform: scale(1,1.2) translateX(1px); opacity: 0.6">&gt;</b>',
			parent: '&laquo;',
			next: '<b style="font-family: Consolas; opacity: 0.6; transform: scale(0.75,0.9); text-shadow: 1px 0">&gt;</b>',
			dropdown: '<b style="font-family: Consolas; transform: scale(1.2,1) rotate(-90deg) translateX(2px); opacity: 0.6; width: 1em; display: inline-block">&lt;</b>'
		};
	} else {
		let s;
		try {
			s = await wsh.regRead("HKCU\\Software\\Microsoft\\Internet Explorer\\Settings\\Always Use My Font Face");
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

	if (await api.SHTestTokenMembership(null, 0x220) && WINVER >= 0x600) {
		TITLE += ' [' + (await api.LoadString(hShell32, 25167) || "Admin").replace(/;.*$/, "") + ']';
	}
	ui_.TEPath = await api.GetModuleFileName(null);
	ui_.Installed = GetParentFolderName(ui_.TEPath);
	ui_.DoubleClickTime = await sha.GetSystemInformation("DoubleClickTime");
	ui_.hwnd = await te.hwnd;
	let arg = await te.Arguments;
	if (arg) {
		window.dialogArguments = arg;
		$.dialogArguments = arg;
	}
	if (await window.dialogArguments) {
		for (let list = await api.CreateObject("Enum", await dialogArguments.event); !await list.atEnd(); await list.moveNext()) {
			const j = await list.item();
			const res = /^on(.+)/i.exec(j);
			if (res) {
				AddEventEx(window, res[1], await dialogArguments.event[j]);
			} else {
				window[j] = await dialogArguments.event[j];
			}
		}
	}
	let x;
	if (!window.MainWindow) {
		window.MainWindow = await $;
		while (x = await MainWindow.dialogArguments || await MainWindow.opener) {
			MainWindow = x;
			if (x = await MainWindow.MainWindow) {
				MainWindow = x;
			}
		}
		if (x = await MainWindow.$) {
			MainWindow = x;
		}
	}
	window.ParentWindow = (await window.dialogArguments ? await dialogArguments.opener : window.opener);
	if (ParentWindow || !window.te) {
		te = MainWindow.te;
	}
	if (x = ParentWindow) {
		if (x && await x.$) {
			ParentWindow = x;
		}
	}
	let uid = location.hash.replace(/\D/g, "");
	if (!await window.dialogArguments && !window.opener) {
		arg = await api.CommandLineToArgv(await api.GetCommandLine());
		if (await arg.Count > 3 && SameText(await arg[1], '/open')) {
			uid = await arg[3];
		}
	}
	if (uid) {
		for (let esw = await api.CreateObject("Enum", await sha.Windows()); !await esw.atEnd(); await esw.moveNext()) {
			x = await esw.item();
			if (x && await x.Document) {
				const w = await x.Document.parentWindow;
				if (w && await w.te && await w.Exchange) {
					const a = await w.Exchange[uid];
					if (a) {
						dialogArguments = a;
						MainWindow = w;
						te = await w.te;
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

	UI = await api.CreateObject("Object");
	UI.Addons = await api.CreateObject("Object");

	UI.Invoke = function () {
		Invoke(Array.apply(null, arguments));
	}

	UI.Autocomplete = async function (s, path) {
		const dl = document.getElementById("AddressList");
		while (dl.lastChild) {
			dl.removeChild(dl.lastChild);
		}
		if (await g_.Autocomplete.Path == path) {
			const ar = s.split("\t");
			for (let i = 0; i < ar.length; ++i) {
				const el = document.createElement("option");
				el.value = ar[i];
				dl.appendChild(el);
			}
		}
	}
};

InvokeUI = function () {
	if (arguments.length == 2 && arguments[1].unshift) {
		const args = Array.apply(null, arguments[1]);
		args.unshift(arguments[0]);
		Invoke(args);
		return S_OK;
	}
	Invoke(Array.apply(null, arguments));
	return S_OK;
}

ExecJavaScript = async function () {
	const args = Array.apply(null, arguments);
	const s = FixScript(args.shift(), window.chrome);
	new Function(s).apply(null, args);
}

OpenHttpRequest = async function (url, alt, fn, arg) {
	const xhr = await createHttpRequest();
	const fnLoaded = async function () {
		if (fn) {
			if (arg && await arg.pcRef) {
				arg.pcRef[0] = await arg.pcRef[0] - 1;
			}
			if (await xhr.status == 200) {
				if (fn) {
					if ("string" === typeof fn) {
						InvokeUI(fn, [xhr, url, arg]);
					} else {
						fn(xhr, url, arg);
					}
					fn = void 0;
				}
				return;
			}
			if (/^http/.test(alt)) {
				OpenHttpRequest(/^https/.test(url) && alt == "http" ? url.replace(/^https/, alt) : alt, '', fn, arg);
				return;
			}
			ShowXHRError(url, await xhr.status);
		}
	}
	xhr.onload = fnLoaded;
	xhr.onreadystatechange = async function () {
		if (await xhr.readyState == 4) {
			fnLoaded();
		}
	}
	if (/ml$/i.test(url)) {
		url += "?" + Math.floor(new Date().getTime() / 60000);
	}
	if (arg && await arg.pcRef) {
		arg.pcRef[0] = await arg.pcRef[0] + 1;
	}
	if (window.chrome && /\.zip$|\.nupkg$/i.test(url)) {
		xhr.responseType = "blob";
	}
	xhr.open("GET", url, true);
	xhr.send();
}

DownloadFile = async function (url, fn) {
	const hr = await MainWindow.RunEvent4("DownloadFile", url, fn);
	return hr != null ? hr : await api.URLDownloadToFile(null, url, fn);
}

ReadAsDataURL = function (blob) {
	return new Promise(function (resolve, reject) {
		const reader = new FileReader();
		reader.onload = function () {
			resolve(reader.result);
		};
		reader.onerror = function () {
			reject(reader.error);
		};
		reader.readAsDataURL(blob);
	});
}

Extract = async function (Src, Dest, xhr) {
	let hr;
	if (xhr) {
		if (window.chrome && await xhr.response) {
			xhr = await ReadAsDataURL(xhr.response);
		}
		hr = await DownloadFile(xhr, Src);
		if (hr) {
			return hr;
		}
	}
	hr = await MainWindow.RunEvent4("Extract", Src, Dest);
	return hr != null ? hr : api.Extract(BuildPath(system32, "zipfldr.dll"), "{E88DCCE0-B7B3-11d1-A9F0-00AA0060FA31}", Src, Dest);
}

CheckUpdate2 = async function (xhr, url, arg1) {
	const arg = await api.CreateObject("Object");
	const Text = await xhr.get_responseText ? await xhr.get_responseText() : xhr.responseText;
	const json = JSON.parse(Text);
	if (json.assets && json.assets[0]) {
		arg.size = json.assets[0].size / 1024;
		arg.url = json.assets[0].browser_download_url;
	}
	if (!await arg.url) {
		return;
	}
	arg.file = GetFileName((await arg.url).replace(/\//g, "\\"));
	let ver = 0;
	res = /(\d+)/.exec(await arg.file);
	if (res) {
		ver = await api.Add(20000000, res[1]);
	}
	if (ver <= await AboutTE(0)) {
		if ((arg1 && GetNum(await arg1.silent)) || await MessageBox(await AboutTE(2) + "\n" + await GetText("the latest version"), TITLE, MB_ICONINFORMATION)) {
			if (await api.GetKeyState(VK_SHIFT) >= 0 || await api.GetKeyState(VK_CONTROL) >= 0) {
				MainWindow.RunEvent1("CheckUpdate", arg1);
				return;
			}
		}
	}
	if (!(arg1 && GetNum(await arg1.noconfirm))) {
		const s = (await api.LoadString(hShell32, 60) || "%").replace(/%.*/, await api.sprintf(99, "%d.%d.%d (%.1lfKB)", ver / 10000 % 100, ver / 100 % 100, ver % 100, await arg.size));
		if (!await confirmOk([await GetText("Update available"), s, await GetText("Do you want to install it now?")].join("\n"))) {
			return;
		}
	}
	arg.temp = BuildPath(await wsh.ExpandEnvironmentStrings("%TEMP%"), "tablacus");
	await CreateFolder2(await arg.temp);
	arg.InstalledFolder = ui_.Installed;
	arg.zipfile = BuildPath(await arg.temp, await arg.file);
	arg.temp = await arg.temp + "\\explorer";
	await DeleteItem(await arg.temp);
	await CreateFolder2(await arg.temp);
	OpenHttpRequest(await arg.url, "http://tablacus.github.io/TablacusExplorerAddons/te/" + ((await arg.url).replace(/^.*\//, "")), "CheckUpdate3", arg);
}

CheckUpdate3 = async function (xhr, url, arg) {
	let hr = await Extract(await arg.zipfile, await arg.temp, xhr);
	if (hr) {
		await MessageBox([(await api.LoadString(hShell32, 4228)).replace(/^\t/, "").replace("%d", await api.sprintf(99, "0x%08x", hr)), await GetText("Extract"), GetFileName(arg.zipfile)].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
		return;
	}
	let te_exe = await arg.temp + "\\te64.exe";
	let nDog = 300;
	while (!await $.fso.FileExists(te_exe)) {
		if (await wsh.Popup(await GetText("Please wait."), 1, TITLE, MB_OKCANCEL) == IDCANCEL || nDog-- == 0) {
			return;
		}
	}
	const arDel = [];
	const addons = await arg.temp + "\\addons";
	if (await $.fso.FolderExists(await arg.temp + "\\config")) {
		arDel.push(await arg.temp + "\\config");
	}
	for (let i = 32; i <= 64; i += 32) {
		te_exe = await arg.temp + '\\te' + i + '.exe';
		const te_old = BuildPath(ui_.Installed, 'te' + i + '.exe');
		if (!await $.fso.FileExists(te_old) || await $.fso.GetFileVersion(te_exe) == await $.fso.GetFileVersion(te_old)) {
			arDel.push(te_exe);
		}
	}
	for (let list = await api.CreateObject("Enum", await $.fso.GetFolder(addons).SubFolders); !await list.atEnd(); await list.moveNext()) {
		const n = await list.item().Name;
		const items = await te.Data.Addons.getElementsByTagName(n);
		if (!items || GetLength(items) == 0) {
			arDel.push(BuildPath(addons, n));
		}
	}
	if (arDel.length) {
		await api.SHFileOperation(FO_DELETE, arDel, null, FOF_SILENT | FOF_NOCONFIRMATION, false);
	}
	const ppid = await api.Memory("DWORD");
	await api.GetWindowThreadProcessId(ui_.hwnd, ppid);
	arg.pid = await ppid[0];
	MainWindow.CreateUpdater(arg);
}

LoadScripts = async function (js1, js2, cb) {
	await InitUI();
	if (window.chrome) {
		js1.unshift("script/consts.js", "script/common.js", "script/syncb.js");
		const s = [];
		let arFN = ["fso", "sha", "wsh", "wnw"];
		let line = 1;
		for (let i = 0; i < js1.length; i++) {
			const fn = BuildPath(ui_.Installed, js1[i]);
			const ado = await api.CreateObject("ads");
			if (ado) {
				ado.CharSet = "utf-8";
				await ado.Open();
				await ado.LoadFromFile(fn);
				const src = FixScript(await ado.ReadText());
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
		const CopyObj = async function (to, o, ar) {
			if (!to) {
				to = await api.CreateObject("Object");
			}
			ar.forEach(async function (key) {
				const a = await o[key];
				if (!await to[key] && a != null) {
					to[key] = a;
				}
			});
			return to;
		}
		document.documentMode = (/Edg\/(\d+)/.test(navigator.appVersion) ? RegExp.$1 : 12);
		screen.deviceYDPI = parent.deviceYDPI;
		ui_.Zoom = parent.deviceYDPI / 96;
		await CopyObj($, window, ["te", "api", "chrome", "document", "UI", "MainWindow"]);
		$.location = await CopyObj(null, location, ["hash", "href"]);
		$.navigator = await CopyObj(null, navigator, ["appVersion", "language"]);
		$.screen = await CopyObj(null, screen, ["deviceYDPI"]);
		const o = await api.CreateObject("Object");
		o.window = $;
		$.$JS = await api.GetScriptDispatch(s.join(""), "JScript", o);
		await CopyObj(window, $, ["g_", "Common", "Sync", "Threads"]);
		await CopyObj(window, $, arFN);
		const doc = await api.CreateObject("Object");
		doc.parentWindow = $;
		WebBrowser.Document = doc;
	} else {
		$ = window;
	}
	window.Addons = {
		"_stack": await api.CreateObject("Array")
	};
	LoadScript(js1.concat(js2), cb);
};

async function _InvokeMethod() {
	let args;
	const ar = await api.ObjGetI(te, "fn");
	while (args = await ar.shift()) {
		const fn = args.shift();
		await fn.apply(fn, args);
	}
}

ApplyLangTag = async function (o) {
	if (o) {
		for (let i = 0; i < o.length; ++i) {
			(async function (el) {
				let s;
				if (s = el.childNodes) {
					for (let j = s.length; j-- > 0;) {
						if (!s[j].tagName) {
							s[j].data = amp2ul(await GetTextR(s[j].data.replace(/&amp;/ig, "&")));
						}
					}
				}
				if (s = el.title) {
					el.title = (await GetTextR(s)).replace(/\(&\w\)|&/, "");
				}
				if (s = el.alt) {
					el.alt = (await GetTextR(s)).replace(/\(&\w\)|&/, "");
				}
			})(o[i]);
		}
	}
}

ApplyLang = async function (doc) {
	if (!doc) {
		doc = document;
	}
	let i, o, s, h = 0;
	const FaceName = await MainWindow.DefaultFont.lfFaceName;
	if (doc.body) {
		doc.body.style.fontFamily = FaceName;
		doc.body.style.fontSize = Math.abs(await MainWindow.DefaultFont.lfHeight) + "px";
		doc.body.style.fontWeight = await MainWindow.DefaultFont.lfWeight;
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
			(async function (el) {
				if (!h && SameText(el.type, "text")) {
					h = el.offsetHeight;
				}
				if (s = el.placeholder) {
					el.placeholder = (await GetTextR(s)).replace(/\(&\w\)|&/, "");
				}
				if (SameText(el.type, "button")) {
					if (s = el.value) {
						el.value = (await GetTextR(s)).replace(/\(&\w\)|&/, "");
					}
				}
				if (s = await ImgBase64(el, 0)) {
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
			(async function (el) {
				const s = await ImgBase64(el, 0);
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
			(async function (el) {
				el.title = delamp(await GetTextR(el.title));
				for (let j = 0; j < el.length; ++j) {
					el[j].text = (await GetTextR(el[j].text)).replace(/^\n/, "").replace(/\n$/, "");
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
	doc.title = await GetTextR(doc.title);
}

ImgBase64 = async function (el, index, h) {
	const org = el.src || el.getAttribute("data-src");
	const src = await ExtractMacro(te, org);
	const s = await MakeImgSrc(src, index, false, h || el.height);
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
			const evt = document.createEvent("HTMLEvents");
			evt.initEvent(event, true, true);
			return !o.dispatchEvent(evt);
		}
	}
}

GetRect = async function (o, f) {
	const rc = await api.Memory("RECT");
	const pt = GetPos(o, f);
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
	const x = bScreen ? screenLeft * ui_.Zoom : 0;
	let y = bScreen ? screenTop * ui_.Zoom : 0;
	if (bBottom) {
		y += o.offsetHeight;
	}
	const rc = o.getBoundingClientRect();
	return { x: x + rc.left, y: y + rc.top };
}

GetPosEx = async function (el, n) {
	const p = GetPos(el, n);
	const pt = await api.Memory("POINT");
	pt.x = p.x;
	pt.y = p.y;
	return pt;
}

HitTest = async function (o, pt) {
	if (o) {
		let p = GetPos(o, 1);
		if (await pt.x >= p.x && await pt.x < p.x + o.offsetWidth && await pt.y >= p.y && await pt.y < p.y + o.offsetHeight) {
			o = o.offsetParent;
			p = GetPos(o, 1);
			return await pt.x >= p.x && await pt.x < p.x + o.offsetWidth && await pt.y >= p.y && await pt.y < p.y + o.offsetHeight;
		}
	}
	return false;
}

GetTopWindow = async function (hwnd) {
	let hwnd1 = hwnd || await WebBrowser.hwnd;
	while (hwnd1 = await api.GetParent(hwnd1)) {
		hwnd = hwnd1;
	}
	return hwnd;
}

CloseWindow = async function () {
	if (window.chrome) {
		api.PostMessage(await GetTopWindow(), WM_CLOSE, 0, 0);
		return;
	}
	window.close();
}

CloseSubWindows = async function () {
	const hwnd = await GetTopWindow();;
	let hwnd1 = hwnd;
	while (hwnd1 = await api.FindWindowEx(null, hwnd1, null, null)) {
		if (hwnd == await api.GetWindowLongPtr(hwnd1, GWLP_HWNDPARENT)) {
			api.PostMessage(hwnd1, WM_CLOSE, 0, 0);
		}
	}
}

MouseOver = async function (o) {
	if ("string" === typeof o) {
		o = document.getElementById(o);
	}
	o = o.srcElement || o;
	if (/^button$|^menu$/i.test(o.className)) {
		if (ui_.objHover && o != ui_.objHover) {
			MouseOut();
		}
		let bHover = window.chrome;
		if (!bHover) {
			const pt = await api.Memory("POINT");
			api.GetCursorPos(pt);
			const ptc = await pt.Clone();
			await api.ScreenToClient(WebBrowser.hwnd, ptc);
			bHover = (o == document.elementFromPoint(ptc.x, ptc.y) || await HitTest(o, pt));
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
	const ot = e.srcElement;
	if (e.keyCode == VK_TAB) {
		ot.focus();
		if (document.all && document.selection) {
			const selection = document.selection.createRange();
			if (selection) {
				selection.text += "\t";
				return false;
			}
		}
		let i = ot.selectionEnd;
		const s = ot.value;
		ot.value = s.slice(0, i) + "\t" + s.slice(i);
		ot.selectionStart = ++i;
		ot.selectionEnd = i;
		return false;
	}
	return true;
}

DetectProcessTag = function (ev) {
	const el = (ev || event).srcElement;
	return el && (/input|textarea/i.test(el.tagName) || /selectable/i.test(el.className));
}

GetFolderView = GetFolderViewEx = async function (Ctrl, pt, bStrict) {
	if (!Ctrl) {
		return await te.Ctrl(CTRL_FV);
	}
	const nType = await Ctrl.Type;
	if (nType == null) {
		let o = Ctrl.offsetParent;
		while (o) {
			const res = /^Panel_(\d+)$/.exec(o.id);
			if (res) {
				return await te.Ctrl(CTRL_TC, res[1]).Selected;
			}
			o = o.offsetParent;
		}
		return await te.Ctrl(CTRL_FV);
	}
	if (nType <= CTRL_EB) {
		return Ctrl;
	}
	if (nType == CTRL_TV) {
		return await Ctrl.FolderView;
	}
	if (nType != CTRL_TC) {
		return await te.Ctrl(CTRL_FV);
	}
	if (pt) {
		const FV = await Ctrl.HitTest(pt);
		if (FV) {
			return FV;
		}
	}
	if (!bStrict || !pt) {
		return await Ctrl.Selected;
	}
}

AddFavoriteEx = async function (Ctrl, pt) {
	const FV = await GetFolderViewEx(Ctrl, pt);
	await FV.Focus();
	AddFavorite();
	return S_OK
}

SelectItem = function (FV, path, wFlags, tm) {
	setTimeout(async function () {
		if (FV) {
			if (SameText(await FV.FolderItem.Path, GetParentFolderName(path))) {
				FV.SelectItem(path, wFlags);
			}
		}
	}, tm);
}

AddEventEx(window, "load", function () {
	document.body.onselectstart = DetectProcessTag;
	if (window.chrome) {
		document.body.addEventListener('mousewheel', function (ev) {
			if (ev.ctrlKey) {
				ev.preventDefault();
			}
		}, { passive: false });
		document.body.addEventListener('keydown', function (ev) {
			if ("F5" === ev.key) {
				ev.preventDefault();
			}
		}, false);
	} else {
		document.body.oncontextmenu = DetectProcessTag;
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
			const el = o.getElementsByTagName("*");
			for (let i in el) {
				if (el[i].style) {
					el[i].style.cursor = s;
				}
			}
		}
	}
}

GetImgTag = async function (o, h) {
	if (o.src) {
		o.src = await ImgBase64(o, 0, Number(h))
		const ar = ['<img'];
		for (let n in o) {
			if (o[n]) {
				ar.push(' ', n, '="', EncodeSC(StripAmp(await GetText(await api.PathUnquoteSpaces(o[n])))), '"');
			}
		}
		if (h) {
			h = Number(h) ? h + 'px' : EncodeSC(h);
			ar.push(' width="', h, '" height="', h, '"');
		}
		ar.push('>');
		return ar.join("");
	}
	const ar = ['<span'];
	for (let n in o) {
		if (n != "title" && o[n]) {
			ar.push(' ', n, '="', EncodeSC(o[n]), '"');
		}
	}
	ar.push('>', EncodeSC(o.title), '</span>');
	return ar.join("");
}

LoadAddon = async function (ext, Id, arError, param, bDisabled) {
	let r, fn;
	try {
		let sc;
		const ar = ext.split(".");
		if (ar.length == 1) {
			ar.unshift("script");
		}
		fn = BuildPath("addons", Id, ar.join("."));
		const s = await ReadTextFile(fn);
		if (s && (!bDisabled || /await|\$\./.test(s))) {
			if (ar[1] == "js") {
				sc = new Function(FixScript(s, window.chrome));
			} else if (ar[1] == "vbs") {
				const o = await api.CreateObject("Object");
				o["_Addon_Id"] = await api.CreateObject("Object");
				o["_Addon_Id"].Addon_Id = Id;
				o.window = window;
				sc = await ExecAddonScript("VBScript", s, fn, arError, o, Addons["_stack"]);
			}
			if (sc) {
				r = await sc(Id);
				if (param) {
					const res = /[\r\n\s]Default\s*=\s*["'](.*)["'];/.exec(s);
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

ReloadCustomize = async function () {
	te.Data.bReload = false;
	await CloseSubWindows();
	await Finalize();
	te.Reload();
	return S_OK;
}

BlurId = function (Id) {
	document.getElementById(Id).blur();
}

SetDisplay = function (Id, s) {
	const o = document.getElementById(Id);
	if (o) {
		o.style.display = s;
	}
}

//Options
AddonOptions = async function (Id, fn, Data, bNew) {
	await LoadLang2(BuildPath("addons", Id, "lang", await GetLangId() + ".xml"));
	const items = await te.Data.Addons.getElementsByTagName(Id);
	if (!GetLength(items)) {
		const root = await te.Data.Addons.documentElement;
		if (root) {
			root.appendChild(await te.Data.Addons.createElement(Id));
		}
	}
	const info = await GetAddonInfo(Id);
	let sURL = "addons\\" + Id + "\\options.html";
	if (!Data) {
		Data = await api.CreateObject("Object");;
	}
	Data.id = Id;
	let sFeatures = await info.Options;
	if (/^Location$/i.test(sFeatures)) {
		sFeatures = "Common:6:6";
	}
	let res = /Common:([\d,]+):(\d)/i.exec(sFeatures);
	if (res) {
		sURL = "script\\location.html";
		Data.show = res[1];
		Data.index = res[2];
		sFeatures = 'Default';
	}
	sURL = BuildPath(ui_.Installed, sURL);
	let opt = await api.CreateObject("Object");
	opt.MainWindow = MainWindow.$;
	opt.Data = Data;
	opt.event = await api.CreateObject("Object");
	if (fn) {
		opt.event.TEOk = fn;
	} else if (window.g_Chg) {
		opt.event.TEOk = function () {
			g_Chg.Addons = true;
		}
	}
	if (bNew || window.Addon == 1 || await api.GetKeyState(VK_CONTROL) < 0) {
		if (/^Default$/i.test(sFeatures)) {
			sFeatures = 'Width: 640; Height: 480';
		}
		try {
			const dlg = await MainWindow.g_.dlgs[Id];
			if (dlg) {
				if (await api.IsWindowVisible(await dlg.Document.parentWindow.WebBrowser.hwnd)) {
					dlg.Focus();
					return;
				}
			}
		} catch (e) {
			MainWindow.g_.dlgs[Id] = void 0;
		}
		opt = await api.CreateObject("Object");
		opt.MainWindow = await $;
		opt.Data = Data;
		opt.event = await api.CreateObject("Object");
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
		MainWindow.g_.dlgs[Id] = await ShowDialog(sURL, opt);
		return;
	}
	if (!ui_.elAddons[Id]) {
		if (!/location\.html$/.test(sURL)) {
			opt.event.onload = function () {
				const cInput = el.contentWindow.document.getElementsByTagName('input');
				for (let i in cInput) {
					if (/^ok$|^cancel$/i.test(cInput[i].className)) {
						cInput[i].style.display = 'none';
					}
				}
				el.contentWindow.g_Inline = true;
			}
		}
		te.Arguments = opt;
		const el = document.createElement('iframe');
		el.id = 'panel1_' + Id;
		el.src = sURL;
		el.style.cssText = 'width: 100%; border: 0; padding: 0; margin: 0';
		ui_.elAddons[Id] = el;
		let o = document.getElementById('panel1_2');
		o.style.display = "block";
		o.appendChild(el);
		o = document.getElementById('tab1_');
		o.insertAdjacentHTML("BeforeEnd", '<label id="tab1_' + Id + '" class="button" style="width: 100%" onmousedown="ClickTree(this);">' + await info.Name + '</label><br>');
	}
	ClickTree(document.getElementById('tab1_' + Id));
}

GetElementEx = function (Id) {
	const el = document.getElementById(Id);
	if (el) {
		return el;
	}
	const ar = Id.split("::");
	return document.forms[ar[0]].elements[ar[1]];
}

GetElementIdEx = function (el) {
	if ("string" === typeof el) {
		return el;
	}
	if (el.id) {
		return el.id;
	}
	return el.form.name + "::" + el.name;
}

SetElement = async function (Id, v) {
	(GetElementEx(Id) || {}).value = v;
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

ShowLocationEx = async function (s) {
	const opt = await api.CreateObject("Object");
	opt.MainWindow = MainWindow;
	opt.Data = s;
	ShowDialog(BuildPath(ui_.Installed, "script\\location.html"), opt);
}

MakeKeySelect = async function () {
	let s;
	let oa = document.getElementById("_KeyState");
	if (oa) {
		const ar = [];
		for (let i = 0; i < 4; ++i) {
			s = await MainWindow.g_.KeyState[i][0];
			ar.push('<label><input type="checkbox" onclick="KeyShift(this)" id="_Key', s, '">', s, '&nbsp;</label>');
		}
		oa.insertAdjacentHTML("AfterBegin", ar.join(""));
	}
	oa = document.getElementById("_KeySelect");
	oa.length = 0;
	oa[++oa.length - 1].value = "";
	oa[oa.length - 1].text = await GetText("Select");
	s = [];
	for (let j = 256; j >= 0; j -= 256) {
		for (let i = 128; i > 0; i--) {
			const v = await api.GetKeyNameText((i + j) * 0x10000);
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
	let j = "";
	for (i in s) {
		if (j != s[i]) {
			j = s[i];
			const o = oa[++oa.length - 1];
			o.value = j;
			o.text = j + "\x20";
		}
	}
}

SetKeyShift = async function () {
	let key = ((document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key).value;
	key = key.replace(/^(.+),.+/, "$1");
	const nLen = await GetLength(await MainWindow.g_.KeyState);
	let o;
	for (let i = 0; i < nLen; ++i) {
		const s = await MainWindow.g_.KeyState[i][0];
		if (o = document.getElementById("_Key" + s)) {
			o.checked = key.indexOf(s + "+") >= 0;
		}
		key = key.replace(s + "+", "");
	}
	o = document.getElementById("_KeySelect");
	for (let i = o.length; i--;) {
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
		let h = document.documentElement.clientHeight || document.body.clientHeight;
		h += MainWindow.DefaultFont.lfHeight * em;
		if (h > 0) {
			o.style.height = h + 'px';
			o.style.height = 2 * h - o.offsetHeight + "px";
		}
	}
}

KeyShift = function (o) {
	const oKey = (document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key;
	let key = oKey.value;
	const shift = o.id.replace(/^_Key(.*)/, "$1+");
	key = key.replace(shift, "");
	if (o.checked) {
		key = shift + key;
	}
	oKey.value = key;
}

KeySelect = function (o) {
	const oKey = (document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key;
	const ar = oKey.value == "," ? [","] : oKey.value.split(",");
	ar[0] = ar[0].replace(/(\+)[^\+]*$|^[^\+]*$/, "$1") + o[o.selectedIndex].value;
	oKey.value = ar.join(",");
}

createHttpRequest = async function () {
	const xhr = await MainWindow.RunEvent4("createHttpRequest");
	if (xhr != null) {
		return xhr;
	}
	try {
		return window.XMLHttpRequest && ui_.IEVer >= 9 ? new XMLHttpRequest() : await api.CreateObject("Msxml2.XMLHTTP");
	} catch (e) {
		return await api.CreateObject("Microsoft.XMLHTTP");
	}
}

ShowXHRError = async function (url, status) {
	MessageBox([await api.sprintf(999, (await api.LoadString(hShell32, 4227)).replace(/^\t/, ""), status), url].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
}

ClearAutocomplete = function () {
	for (let dl = document.getElementById("AddressList"); dl.lastChild;) {
		dl.removeChild(dl.lastChild);
	}
	g_.Autocomplete.Path = "";
}

GetXmlItems = window.chrome ? async function (items) {
	return JSON.parse(await XmlItems2Json(items));
} : function (items) {
	const ar = [];
	if (items) {
		for (let i = 0; i < items.length; ++i) {
			const item = items[i];
			if (item) {
				const o = {};
				if (item.text) {
					o.text = item.text;
				}
				const attrs = item.attributes;
				if (attrs) {
					for (let j = 0; j < attrs.length; ++j) {
						o[attrs[j].name] = attrs[j].value;
					}
				}
				ar.push(o);
			}
		}
	}
	return ar;
}

SyncExec = async function (cb, el) {
	cb(await GetFolderViewEx(el));
}

if (window.chrome) {
	GetAddonElement = async function (id) {
		const item = await $.GetAddonElement(id);
		const o = {
			item: item,
			db: JSON.parse(await XmlItem2Json(item)),
			getAttribute: function (s) {
				return this.db[s];
			},
			setAttribute: function (s, v) {
				this.db[s] = v;
				this.item.setAttribute(s, v);
			},
			removeAttribute: function (s) {
				delete this.db[s];
				this.item.removeAttribute(s, v);
			}
		};
		o.__defineGetter__("attributes", function () {
			const ar = [];
			for (let n in this.db) {
				ar.push({
					name: n, value: this.db[n]
				});
			}
			return ar;
		});
		return o;
	}
}
