// Tablacus Explorer

if (!window.addEventListener && window.attachEvent) {
	window.addEventListener = function (n, fn) {
		window.attachEvent("on" + n, fn);
	}
	document.addEventListener = function (n, fn) {
		document.attachEvent("on" + n, fn);
	}
	document.body.addEventListener = function (n, fn) {
		document.body.attachEvent("on" + n, fn);
	}
}

if (!window.devicePixelRatio) {
	window.devicePixelRatio = 1;
}

ui_ = {
	IEVer: window.chrome ? 12 : ScriptEngineMajorVersion() > 8 ? ScriptEngineMajorVersion() : ScriptEngineMinorVersion(),
	Zoom: 1,
	tmDown: 0,
	eventTE: {},
	MiscIcon: {}
};

InitUI = async function () {
	if (window.chrome) {
		te = await parent.chrome.webview.hostObjects.te;
		parent.chrome.webview.hostObjects.options.shouldSerializeDates = true;
		api = await te.WindowsAPI0.CreateObject("api");
		fso = api.CreateObject("fso");
		sha = api.CreateObject("sha");
		wsh = api.CreateObject("wsh");
		$ = await api.CreateObject("Object");
		$.Addon = window.Addon;
		WebBrowser = await te.Ctrl(CTRL_WB);
		const fn = function R(api) {
			var r = ["undefined" != typeof ScriptEngineMajorVersion && ScriptEngineMajorVersion(), api.Memory("OSVERSIONINFOEX")];
			api.GetVersionEx(r[1]);
			r.push(r[1].dwMajorVersion * 0x100 + r[1].dwMinorVersion);
			return api.CreateObject("SafeArray", r);
		}
		const r = await api.GetScriptDispatch(fn.toString(), "JScript").R(api);
		if (r[0]) {
			ui_.ScriptEngineMajorVersion = r[0];
			ScriptEngineMajorVersion = function () {
				return ui_.ScriptEngineMajorVersion;
			};
		}
		osInfo = r[1];
		WINVER = r[2];
	} else {
		try {
			ui_.NoCssFont = wsh.regRead("HKCU\\Software\\Microsoft\\Internet Explorer\\Settings\\Always Use My Font Face");
		} catch (e) { }
	}
	if (WINVER > 0x603) {
		BUTTONS = {
			opened: '<b style="font-family: Consolas; transform: scale(1.2,1) rotate(-90deg); display: inline-block">&lt;</b>',
			closed: '<b style="font-family: Consolas; transform: scale(1,1.2) translateX(1px); opacity: 0.6">&gt;</b>',
			parent: '&laquo;',
			next: '<b style="font-family: Consolas; opacity: 0.6; transform: scale(0.75,0.9); text-shadow: 1px 0">&gt;</b>',
			dropdown: '<b style="font-family: Consolas; transform: scale(1.2,1) rotate(-90deg) translateX(2px); opacity: 0.6; width: 1em; display: inline-block">&lt;</b>'
		};
	} else {
		BUTTONS = {
			opened: '<span style="font-size: 10pt; transform: translateY(-2pt)">&#x25e2;</span>',
			closed: '<span style="font-size: 10pt; transform: scale(1,1.4)">&#x25b7;</span>',
			parent: '&laquo;',
			next: ui_.NoCssFont ? '&#x25ba;' : '<span style="font-family: Marlett">4</span>',
			dropdown: ui_.NoCssFont ? '&#x25bc;' : '<span style="font-family: Marlett">6</span>'
		};
	}
	const r = await Promise.all([api.GetDisplayNameOf(ssfSYSTEM, SHGDN_FORPARSING), api.GetModuleFileName(null), sha.GetSystemInformation("DoubleClickTime"), te.hwnd, api.SHTestTokenMembership(null, 0x220), te.Arguments, api.CreateObject("Object"), api.CreateObject("Object"), api.sizeof("HANDLE")]);
	system32 = r[0];
	ui_.TEPath = r[1];
	ui_.Installed = GetParentFolderName(ui_.TEPath);
	ui_.DoubleClickTime = r[2];
	ui_.hwnd = r[3];
	ui_.bit = r[8] * 8;
	hShell32 = await api.GetModuleHandle(BuildPath(system32, "shell32.dll"));
	if (r[4] && WINVER >= 0x600) {
		TITLE += ' [' + (await api.LoadString(hShell32, 25167) || "Admin").replace(/;.*$/, "") + ']';
	}
	document.title = TITLE;
	let arg = r[5];
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
		window.MainWindow = $;
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
		const sw = await sha.Windows();
		const nCount = await sw.Count;
		for (let i = 0; i < nCount; ++i) {
			x = await sw.item(i);
			if (x && await x.Document) {
				const w = await x.Document.parentWindow;
				if (w && await w.te && await w.Exchange) {
					const a = await w.Exchange[uid];
					if (a) {
						dialogArguments = a;
						MainWindow = w;
						te = await w.te;
						window.addEventListener("unload", function () {
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

	UI = r[6];
	UI.Addons = r[7];

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
	const s = '{' + FixScript(args.shift(), window.chrome) + '}';
	try {
		new Function("Ctrl", "FV", "type", "hwnd", "pt", s).apply(null, args);
	} catch (e) {
		ShowError(e, s);
	}
}

AddEventUI = function (Name, fn, priority) {
	if (Name) {
		const en = Name.toLowerCase();
		if (!ui_.eventTE[en]) {
			ui_.eventTE[en] = [];
		}
		if (priority) {
			ui_.eventTE[en].unshift(fn);
		} else {
			ui_.eventTE[en].push(fn);
		}
	}
}

ClearEventUI = function (Name) {
	delete ui_.eventTE[Name.toLowerCase()];
}

RunEventUI1 = async function () {
	const args = Array.apply(null, arguments);
	const en = args.shift();
	const eo = ui_.eventTE[en.toLowerCase()];
	for (let i in eo) {
		try {
			await eo[i].apply(eo[i], args);
		} catch (e) {
			await api.Invoke(eo[i], args);
		}
	}
}

OpenHttpRequest = async function (url, alt, fn, arg) {
	const xhr = await createHttpRequest();
	const fnLoaded = async function () {
		if (fn) {
			if (await xhr.status == 200) {
				if (fn) {
					CalcRef(arg && await arg.pcRef, 0, -1);
					if ("string" === typeof fn) {
						InvokeUI(fn, [xhr, url, arg]);
					} else {
						fn(xhr, url, arg);
					}
					fn = void 0;
				}
				return;
			}
			fnError();
		}
	}
	const fnError = async function () {
		if (/^http/.test(alt)) {
			CalcRef(arg && await arg.pcRef, 0, -1);
			OpenHttpRequest(/^https/.test(url) && alt == "http" ? url.replace(/^https/, alt) : alt, '', fn, arg);
			return;
		}
		ShowXHRError(url, await xhr.status);
	}
	xhr.onload = fnLoaded;
	xhr.onerror = fnError;
	if (!window.chrome) {
		xhr.onreadystatechange = async function () {
			if (await xhr.readyState == 4) {
				fnLoaded();
			}
		}
	}
	if (/ml$|\.json$/i.test(url)) {
		url += "?" + Math.floor(new Date().getTime() / 60000);
	}
	CalcRef(arg && await arg.pcRef, 0, 1);
	if (window.chrome) {
		const rt = arg && await arg.responseType;
		if (rt) {
			xhr.responseType = rt;
		} else if (/\.zip$|\.exe$/i.test(url)) {
			xhr.responseType = "blob";
		}
	}
	xhr.open("GET", url, true);
	xhr.send();
}

DownloadFile = async function (url, fn) {
	let hr;
	if (window.chrome && await url.response) {
		url = await ReadAsDataURL(url.response);
	} else {
		hr = await MainWindow.RunEvent4("DownloadFile", url, fn);
	}
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
		hr = await DownloadFile(xhr, Src);
		if (hr) {
			return hr;
		}
	}
	hr = await MainWindow.RunEvent4("Extract", Src, Dest);
	return hr != null ? hr : await api.Extract(BuildPath(system32, "zipfldr.dll"), "{E88DCCE0-B7B3-11d1-A9F0-00AA0060FA31}", Src, Dest);
}

CheckUpdate2 = async function (xhr, url, arg1) {
	const arg = await api.CreateObject("Object");
	const Text = await xhr.get_responseText ? await xhr.get_responseText() : xhr.responseText;
	let json = JSON.parse(Text);
	if (json.length) {
		json = json[0];
	}
	if (json.assets && json.assets[0]) {
		arg.size = json.assets[0].size / 1024;
		arg.url = json.assets[0].browser_download_url;
	}
	if (!await arg.url) {
		return;
	}
	arg.file = GetFileName((await arg.url).replace(/\//g, "\\"));
	const ver = (res = /(\d+)/.exec(await arg.file)) ? GetNum(res[1]) + 2e7 : 0;
	if (ver <= await AboutTE(0)) {
		let bUpdate = arg1 && GetNum(await arg1.silent);
		if (!bUpdate) {
			const res = /^(.*) (\d.*)/i.exec(await AboutTE(2));
			bUpdate = await MessageBox(amp2ul(await GetText("%s is up to date.")).replace("%s", res[1]) + "\n" + res[2], TITLE, MB_ICONINFORMATION);
		}
		if (bUpdate) {
			if (await api.GetKeyState(VK_SHIFT) >= 0 || await api.GetKeyState(VK_CONTROL) >= 0) {
				MainWindow.RunEvent1("CheckUpdate", arg1);
				return;
			}
		}
	}
	if (!(arg1 && GetNum(await arg1.noconfirm))) {
		const s = [await GetTextR("@mstask.dll,-319"), " ", Math.floor(ver / 10000) % 100, ".", Math.floor(ver / 100) % 100, ".", ver % 100, " (", (await arg.size).toFixed(1), "KB)"];
		const lang = (await GetLangId()).replace(/\W.*$/, "");
		if (res = new RegExp("__" + lang + ":([^_]+)__", "i").exec(json.body) || /__([^_]+)__/.exec(json.body)) {
			s.push("\n", res[1]);
		}
		if (!await confirmOk([await GetText("Update available"), s.join(""), "", await GetText("Do you want to install it now?")].join("\n"))) {
			return;
		}
	}
	const temp = await GetTempPath(3);
	await CreateFolder(temp);
	arg.zipfile = BuildPath(temp, await arg.file);
	arg.temp = BuildPath(temp, "explorer");
	await CreateFolder(await arg.temp);
	OpenHttpRequest(await arg.url, "http://tablacus.github.io/TablacusExplorerAddons/te/" + ((await arg.url).replace(/^.*\//, "")), "CheckUpdate3", arg);
}

CheckUpdate3 = async function (xhr, url, arg) {
	const hr = await Extract(await arg.zipfile, await arg.temp, xhr);
	if (hr) {
		ShowExtractError(hr, GetFileName(arg.zipfile));
		return;
	}
	MainWindow.CreateUpdater(arg);
}

LoadScripts = async function (js1, js2, cb) {
	await InitUI();
	if (window.chrome) {
		js1.unshift("script/consts.js", "script/common.js", "script/syncb.js");
		const CopyObj = async function (to, o, ar) {
			if (!to) {
				to = await api.CreateObject("Object");
			}
			let v = [];
			const n = ar.length;
			for (let i = 0; i < n; ++i) {
				v[i] = o[ar[i]];
				v[i + n] = to[ar[i]];
			}
			v = await Promise.all(v);
			for (let i = 0; i < n; ++i) {
				if (v[i] != null && !v[i + n]) {
					to[ar[i]] = v[i];
				}
			}
			return to;
		}
		document.documentMode = 12;
		const hdc = await api.GetDC(ui_.hwnd);
		screen.deviceYDPI = await api.GetDeviceCaps(hdc, 90);
		api.ReleaseDC(ui_.hwnd, hdc);
		await CopyObj($, window, ["te", "api", "chrome", "document", "UI", "MainWindow"]);
		let po = [];
		po.push(CopyObj(null, location, ["hash", "href"]));
		po.push(CopyObj(null, navigator, ["language"]));
		po.push(CopyObj(null, screen, ["deviceYDPI"]));
		po.push(api.CreateObject("Object"));
		po.push(api.CreateObject("Object"));
		po = await Promise.all(po);
		$.location = po.shift();
		$.navigator = po.shift();
		$.screen = po.shift();
		const o = po.shift();
		const doc = po.shift();
		o.window = $;
		const s = [];
		let arFN = [];
		let line = 1;
		for (let i = 0; i < js1.length; i++) {
			const fn = BuildPath(ui_.Installed, js1[i]);
			const ado = await api.CreateObject("ads");
			if (ado) {
				ado.CharSet = "utf-8";
				await ado.Open();
				await ado.LoadFromFile(fn);
				let src = await ado.ReadText();
				if (!/consts\.js$/.test(fn)) {
					src = FixScript(src);
				}
				s.push(src);
				ado.Close();
				if (src && /sync/.test(fn)) {
					arFN = arFN.concat(src.match(/\s\w+\s*=\s*function/g).map(function (s) {
						return s.replace(/^\s|\s*=.*$/g, "");
					}));
				}
				api.OutputDebugString([js1[i], "Start line:", line, "\n"].join(" "));
				line += src.replace(/[^\n]/g, "").length;
			}
		}
		js1.length = 0;
		$.$JS = await api.GetScriptDispatch(s.join(""), "JScript", o);
		await CopyObj(window, $, arFN);
		doc.parentWindow = $;
		WebBrowser.Document = doc;
		delete fso;
		Object.defineProperty(window, "fso", {
			get: function () {
				return $.fso;
			}
		});
		delete wsh;
		Object.defineProperty(window, "wsh", {
			get: function () {
				return $.wsh;
			}
		});
		delete sha;
		Object.defineProperty(window, "sha", {
			get: function () {
				return $.sha;
			}
		});
		Object.defineProperty(window, "wnw", {
			get: function () {
				return $.wnw;
			}
		});
		Object.defineProperty(window, "Common", {
			get: function () {
				return $.Common;
			}
		});
		Object.defineProperty(window, "Sync", {
			get: function () {
				return $.Sync;
			}
		});
		Object.defineProperty(window, "Threads", {
			get: function () {
				return $.Threads;
			}
		});
		Object.defineProperty(window, "FolderMenu", {
			get: function () {
				return $.FolderMenu;
			}
		});
		Object.defineProperty(window, "clipboardData", {
			get: function () {
				return $.clipboardData;
			}
		});
		Object.defineProperty(window, "g_", {
			get: function () {
				return $.g_;
			}
		});
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

AddRule = function (s) {
	const css = document.styleSheets.item(0);
	if (css.insertRule) {
		css.insertRule(s, css.cssRules.length);
	} else if (css.addRule) {
		const s2 = s.split(/\s*{\s*/);
		css.addRule(s2[0], s2[1].replace(/}.*/m, ""));
	}
}

ReadCss = async function (s) {
	const ar = s.split(/}\s*/m);
	while (s = ar.shift()) {
		try {
			const res = /\@import url\("([^"]+)"\);\s*(.*)/m.exec(s);
			if (res) {
				LoadCss(res[1]);
			}
			AddRule(s.replace(/^@[^;]*;\s*/gm, "") + "}");
		} catch (e) { }
	}
}

LoadCss = async function (fn) {
	ReadCss(await ReadTextFile(fn));
}

ApplyLangTag = async function (o) {
	if (o) {
		for (let i = 0; i < o.length; ++i) {
			(async function (el) {
				let s, cn;
				if (s = el.title) {
					el.title = (await GetTextR(s)).replace(/\(&\w\)|&/, "");
				}
				if (s = el.getAttribute("alt")) {
					if (SameText(el.tagName, "label")) {
						if (s = await GetAltText(s)) {
							if (cn = el.childNodes) {
								for (let j = cn.length; j-- > 0;) {
									const tn = cn[j].tagName;
									if (tn) {
										if (SameText(tn, "input")) {
											continue;
										}
									} else if (s) {
										cn[j].data = s.replace(/\(&\w\)|&/, "");
										s = "";
										continue;
									}
									el.removeChild(cn[j]);
								}
							}
							el.removeAttribute("alt");
							return;
						}
					}
				}
				if (cn = el.childNodes) {
					for (let j = cn.length; j-- > 0;) {
						if (!cn[j].tagName) {
							cn[j].data = amp2ul(await GetTextR(cn[j].data.replace(/&amp;/ig, "&")));
						}
					}
				}
			})(o[i]);
		}
	}
}

ApplyLang = async function (doc) {
	if (!doc) {
		doc = document;
	}
	const ButtonIcon = {
		"Up": 0xe74a,
		"Down": 0xe74b,
		"Remove": 0xe74d,
		"Delete": 0xe74d,
		"Portable": 0xe821,
		"Select": 0xea37,
		"Add": 0xe710,
		"Replace": 0xe74e,
		"Browse": 0xe712,
		"Search": 0xe721,
		"Default": 0xe777,
		"Input": 0xe961,
		"Details": 0xe946,
		"Options": 0xe713,
		"Refresh": 0xe72c,
		"Open": 0xe8e5,
		"Menus": 0xed0e,
		"Test": 0xe768,
		"None": 0xe75c,
		"Swap": 0xe8ab
	}
	let s, h = 0;
	if (doc.body) {
		const r = await Promise.all([MainWindow.DefaultFont.lfFaceName, MainWindow.DefaultFont.lfHeight, MainWindow.DefaultFont.lfWeight, MainWindow.g_.IconFont]);
		doc.body.style.fontFamily = r[0];
		doc.body.style.fontSize = Math.abs(r[1]) + "px";
		doc.body.style.fontWeight = r[2];
		if (!ui_.NoCssFont) {
			ui_.IconFont = r[3];
		}
	}
	ApplyLangTag(doc.getElementsByTagName("label"));
	ApplyLangTag(doc.getElementsByTagName("li"));
	let o = doc.getElementsByTagName("a");
	if (o) {
		ApplyLangTag(o);
		for (let i = o.length; i--;) {
			if (o[i].className == "treebutton" && o[i].innerHTML == "") {
				o[i].innerHTML = BUTTONS.opened;
			}
		}
	}
	o = doc.getElementsByTagName("button");
	if (o) {
		let r = [];
		for (let i = o.length; i--;) {
			let s = o[i].innerHTML;
			if (s) {
				r[i * 2] = GetTextR(s);
				const n = s.replace(/\.+$/, "").toLowerCase();
				r[i * 2 + 1] = "string" === ui_.MiscIcon[n] ? ui_.MiscIcon[n] : GetMiscIcon(n);
			}
		}
		(async function (o, r) {
			if (window.chrome) {
				r = await Promise.all(r);
			}
			for (let i = o.length; i--;) {
				const el = o[i];
				let s = el.innerHTML;
				if (s) {
					const v = r[i * 2].replace(/\(&\w\)|&/, "");
					let icon = r[i * 2 + 1];
					if (icon) {
						el.innerHTML = await GetImgTag({
							title: v,
							src: icon
						}, GetNum(window.getComputedStyle ? getComputedStyle(document.body).getPropertyValue('font-size') : document.body.currentStyle.fontSize) * 4 / 3);
						el.className += " svgiconbutton";
					} else if (icon = ui_.IconFont && ButtonIcon[s.replace(/\.+$/, "")]) {
						el.innerHTML = String.fromCodePoint(icon);
						el.style.fontFamily = ui_.IconFont;
						if (!el.title) {
							el.title = v;
						}
						el.className += " fonticonbutton";
					} else if (s != v) {
						el.innerHTML = v;
					}
				}
			}
		})(o, r);
	}
	o = doc.getElementsByTagName("input");
	if (o) {
		ApplyLangTag(o);
		for (let i = o.length; i--;) {
			(async function (el) {
				if (!h && SameText(el.type, "text")) {
					h = el.offsetHeight;
				}
				if (s = el.placeholder) {
					el.placeholder = (await GetTextR(s)).replace(/\(&\w\)|&/, "");
				}
				if (SameText(el.type, "button")) {
					if (s = el.value) {
						const icon = ui_.IconFont && ButtonIcon[s.replace(/\.+$/, "")];
						const v = (await GetTextR(s)).replace(/\(&\w\)|&/, "");
						if (icon) {
							el.value = String.fromCodePoint(icon);
							el.style.fontFamily = ui_.IconFont;
							if (!el.title) {
								el.title = v;
							}
							el.className += " fonticonbutton";
						} else {
							el.value = v;
						}
					}
				}
				if (s = await ImgBase64(el, 0)) {
					el.src = s;
					if (SameText(el.type, "text")) {
						el.style.backgroundImage = "url('" + s + "')";
					}
				}
				el.spellcheck = false;
			})(o[i]);
		}
	}
	o = doc.getElementsByTagName("img");
	if (o) {
		ApplyLangTag(o);
		for (let i = o.length; i--;) {
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
		for (let i = o.length; i--;) {
			if (/translate/i.test(o[i].className)) {
				(async function (el) {
					el.title = delamp(await GetTextR(el.title));
					for (let j = 0; j < el.length; ++j) {
						el[j].text = (await GetTextR(el[j].text)).replace(/^\n/, "").replace(/\n$/, "");
					}
				})(o[i]);
			}
		}
	}
	o = doc.getElementsByTagName("textarea");
	if (o) {
		for (let i = o.length; i--;) {
			o[i].onkeydown = InsertTab;
			o[i].spellcheck = false;

		}
	}
	o = doc.getElementsByTagName("form");
	if (o) {
		for (let i = o.length; i--;) {
			o[i].onsubmit = function () { return false };
		}
	}
	doc.title = await GetTextR(doc.title);
}

ImgBase64 = async function (el, index, h) {
	return await MakeImgSrc(el.src || el.getAttribute("data-src"), index, false, h || el.height || window.IconSize);
}

FireEvent = function (o, event) {
	if (o) {
		if (window.chrome) {
			const evt = new Event(event);
			return !o.dispatchEvent(evt);
		} else if (document.createEvent) {
			const evt = document.createEvent("HTMLEvents");
			evt.initEvent(event, true, true);
			return !o.dispatchEvent(evt);
		} else if (o.fireEvent) {
			return o.fireEvent('on' + event);
		}
	}
}

GetRect = async function (el, f) {
	const rc = await api.Memory("RECT");
	if (el) {
		const pt = GetPos(el, f);
		await api.SetRect(rc, pt.x, pt.y, pt.x + el.offsetWidth, pt.y + el.offsetHeight);
	}
	return rc;
}

GetPos = function (el, bScreen, bAbs, bPanel, bBottom) {
	if (!el) {
		return { x: -32768, y: -32768 };
	}
	if ("number" === typeof bScreen) {
		bAbs = bScreen & 2;
		bPanel = bScreen & 4;
		bBottom = bScreen & 8;
		bScreen &= 1;
	}
	let rc;
	const pt = {};
	try {
		rc = el.getBoundingClientRect();
		pt.x = rc.left;
		pt.y = rc.top;
	} catch (e) {
		pt.x = el.offsetLeft;
		pt.y = el.offsetTop;
	}
	if (window.chrome && window.frameElement) {
		rc = frameElement.getBoundingClientRect();
		pt.x += rc.left;
		pt.y += rc.top;
	}
	if (bScreen) {
		pt.x += screenLeft;
		pt.y += screenTop;
	}
	if (bBottom) {
		pt.y += el.offsetHeight;
	}
	return pt;
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
		const r = await Promise.all([pt.x, pt.y]);
		if (r[0] >= p.x && r[0] < p.x + o.offsetWidth && r[1] >= p.y && r[1] < p.y + o.offsetHeight) {
			o = o.offsetParent;
			p = GetPos(o, 1);
			return r[0] >= p.x && r[0] < p.x + o.offsetWidth && r[1] >= p.y && r[1] < p.y + o.offsetHeight;
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
	CloseWindows(await GetTopWindow(), "TablacusExplorer*");
}

MouseOver = async function (o) {
	if ("string" === typeof o) {
		o = document.getElementById(o);
		if (!o) {
			return;
		}
	}
	o = o.target || o.srcElement || o;
	if (/button|menu/.test(o.className)) {
		if (ui_.objHover && o != ui_.objHover) {
			MouseOut();
		}
		let bHover = window.chrome;
		if (!bHover) {
			const pt = await api.GetCursorPos();
			const ptc = await pt.Clone();
			await api.ScreenToClient(await WebBrowser.hwnd, ptc);
			bHover = (o == document.elementFromPoint(await ptc.x, await ptc.y) || await HitTest(o, pt));
		}
		if (bHover) {
			ui_.objHover = o;
			const ar = o.className.split(/\s+/);
			for (let i = ar.length; --i >= 0;) {
				if (/^button$|^menu$/.test(ar[i])) {
					ar[i] = 'hover' + ar[i];
				}
			}
			o.className = ar.join(" ");
		}
	}
}

MouseOut = function (s) {
	if (ui_.objHover) {
		if ("string" !== typeof s || ui_.objHover.id.indexOf(s) >= 0) {
			const ar = ui_.objHover.className.split(/\s+/);
			for (let i = ar.length; --i >= 0;) {
				if (ar[i] == 'hoverbutton') {
					ar[i] = 'button';
				} else if (ar[i] == 'hovermenu') {
					ar[i] = 'menu';
				}
			}
			ui_.objHover.className = ar.join(" ");
			ui_.objHover = null;
		}
	}
	return S_OK;
}

InsertTab = function (ev) {
	ev = ev || event;
	const ot = ev.target || ev.srcElement;
	if (ev.keyCode ? ev.keyCode == VK_TAB : "Tab" === ev.key) {
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

PreventDefault = function (ev) {
	ev = ev || event;
	if (ev.preventDefault) {
		ev.preventDefault();
	} else {
		ev.returnValue = false;
	}
}

DetectProcessTag = function (ev) {
	ev = ev || event;
	const el = ev.target || ev.srcElement;
	return el && (/input|textarea/i.test(el.tagName) || /selectable/i.test(el.className));
}

GetFolderView = GetFolderViewEx = async function (Ctrl, pt, bStrict) {
	if (!Ctrl) {
		return await te.Ctrl(CTRL_FV);
	}
	const nType = await Ctrl.Type;
	if (nType == null) {
		let o = (Ctrl.target || Ctrl.srcElement || Ctrl).offsetParent;
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
		const FV = await pt.Target || await Ctrl.HitTest(pt);
		if (FV) {
			return FV;
		}
	}
	if (!bStrict || !pt) {
		return await Ctrl.Selected;
	}
}

AddFavoriteEx = async function (Ctrl, pt) {
	const FV = await GetFolderView(Ctrl, pt);
	await FV.Focus();
	AddFavorite();
	return S_OK
}

SelectItem = function (FV, path, wFlags, tm, bCheck) {
	setTimeout(async function () {
		if (FV) {
			const FolderItem = await FV.FolderItem;
			if (FolderItem && SameText(await FolderItem.Path, GetParentFolderName(path))) {
				await FV.SelectItem(path, wFlags);
				if (bCheck && (wFlags & SVSI_FOCUSED)) {
					setTimeout(async function () {
						if (!await api.ILIsEqual(path, await FV.FocusedItem)) {
							await FV.Refresh();
							SelectItem(FV, path, wFlags, tm);
						}
					}, 500);
				}
			}
		}
	}, tm);
}

window.addEventListener("load", function () {
	document.body.onselectstart = DetectProcessTag;
	document.body.oncontextmenu = DetectProcessTag;
	document.body.addEventListener('keydown', function (ev) {
		ev = ev || event;
		if ((ev.keyCode ? ev.keyCode == VK_F5 : "F5" === ev.key) || (ev.ctrlKey && "r" === ev.key)) {
			PreventDefault(ev);
		}
	});

	if (window.chrome) {
		document.body.addEventListener('mousewheel', function (ev) {
			if (ev.ctrlKey) {
				ev.preventDefault();
			}
		}, { passive: false });
		document.body.addEventListener('mousedown', function (ev) {
			if (ev.buttons & 4) {
				ev.preventDefault();
			}
		}, { passive: false });
		if (window.Addon != 1) {
			window.addEventListener('dragover', PreventDefault);
			window.addEventListener('drop', PreventDefault);
		}
	} else {
		document.body.onmousewheel = function (ev) {
			return ev ? !ev.ctrlKey : api.GetKeyState(VK_CONTROL) >= 0;
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
				sc = new AsyncFunction(s);
			} else if (ar[1] == "vbs") {
				const o = await api.CreateObject("Object");
				o["_Addon_Id"] = await api.CreateObject("Object");
				o["_Addon_Id"].Addon_Id = Id;
				o.window = $;
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
		arError.push([e.stack || e.message, fn].join("\n"));
	}
	return r;
}

AddonBeforeRemove = async function (Id) {
	let arError = await api.CreateObject("Array");
	const r = LoadAddon("remove.js", Id, arError);
	arError = await api.CreateObject("SafeArray", arError);
	if (arError.length) {
		setTimeout(async function (arError) {
			MessageBox(arError.join("\n\n"), TITLE, MB_ICONSTOP | MB_OK);
		}, 500, arError);
	}
	return r;
}

FinalizeUI = async function () {
	await CloseSubWindows();
	if (await g_.bFinalized) {
		return;
	}
	g_.bFinalized = true;
	await RunEventUI1("Finalize");
	await FinalizeEx();
}

ReloadCustomize = async function () {
	await FinalizeUI();
	te.Data.bReload = false;
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

FocusElement = function (el) {
	if (ui_.tmActivate) {
		clearTimeout(ui_.tmActivate);
		delete ui_.tmActivate;
	}
	if (el) {
		WebBrowser.Focus();
		el.focus();
	} else if (document.activeElement) {
		document.activeElement.blur();
	}
}

if (window.chrome) {
	GetAddonElement = function (id) {
		const items = ui_.Addons.getElementsByTagName(id.toLowerCase());
		if (items.length) {
			return items[0];
		}
		const item = ui_.Addons.createElement(id.toLowerCase());
		ui_.Addons.documentElement.appendChild(item);
		return item;
	}
}

//Options
AddonOptions = async function (Id, fn, Data, bNew) {
	CloseFindDialog();
	await LoadLang2(BuildPath("addons", Id, "lang", await GetLangId() + ".xml"));
	const items = await te.Data.Addons.getElementsByTagName(Id);
	if (!GetLength(items)) {
		const root = await te.Data.Addons.documentElement;
		if (root) {
			root.appendChild(await te.Data.Addons.createElement(Id));
		}
	}
	const info = GetAddonInfo(Id);
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
			if (dlg && await api.IsWindowVisible(await dlg.hwnd)) {
				dlg.Focus();
				return;
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
		opt.event.onunload = "MainWindow.g_.dlgs['" + Id + "'] = void 0;";
		MainWindow.g_.dlgs[Id] = await ShowDialog(sURL, opt);
		return;
	}
	if (!ui_.elAddons[Id]) {
		te.Arguments = opt;
		const el = document.createElement('iframe');
		el.id = 'panel1_' + Id;
		el.style.cssText = 'width: 100%; border: 0; padding: 0; margin: 0';
		el.onload = function () {
			el.contentWindow.document.body.style.backgroundColor = window.getComputedStyle ? getComputedStyle(document.body).getPropertyValue('background-color') : document.body.currentStyle.backgroundColor;
		}
		el.src = window.chrome ? sURL.replace(/%/g, '%25').replace(/#/g, '%23') : sURL;
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
		let ar = [];
		for (let i = 0; i < 4; ++i) {
			ar[i] = MainWindow.g_.KeyState[i][0];
		}
		ar = await Promise.all(ar);
		for (let i = 0; i < 4; ++i) {
			ar[i] = '<label><input type="checkbox" onclick="KeyShift(this)" id="_Key' + ar[i] + '">' + ar[i] + '&nbsp;</label>';
		}
		oa.insertAdjacentHTML("AfterBegin", ar.join(""));
	}
	oa = document.getElementById("_KeySelect");
	oa.length = 0;
	oa[++oa.length - 1].value = "";
	oa[oa.length - 1].text = await GetText("Select");
	s = [];
	for (let i = 384; i > 0; i--) {
		if (!(i & 128)) {
			s.push(api.GetKeyNameText(i * 0x10000));
		}
	}
	if (window.chrome) {
		s = await Promise.all(s);
	}
	s.sort(function (a, b) {
		if (a.length == 1 || b.length == 1) {
			return a.length - b.length || a.localeCompare(b);
		}
		return ui_.IEVer > 10 ? a.localeCompare(b, {}, { numeric: true }) : api.StrCmpLogical(a, b);
	});
	let v = "";
	for (i in s) {
		if (v != s[i]) {
			if (v = s[i]) {
				if (v.charCodeAt(0) > 32) {
					const o = oa[++oa.length - 1];
					o.text = o.value = v;
				}
			}
		}
	}
}

SetKeyShift = async function () {
	let key = ((document.E && document.E.KeyKey) || document.F.KeyKey || document.F.Key).value;
	key = key.replace(/^(.+),.+/, "$1");
	const nLen = await GetLength(await MainWindow.g_.KeyState);
	let o, r = [];
	for (let i = 0; i < nLen; ++i) {
		r.push(MainWindow.g_.KeyState[i][0]);
	}
	r = await Promise.all(r);
	for (let i = 0; i < nLen; ++i) {
		const s = r[i];
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

CalcElementHeight = function (o, em, r) {
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
		if (!r) {
			setTimeout(CalcElementHeight, 999, o, em, 1);
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
	MessageBox([(await GetTextR("@shell32,-4227")).replace(/^\t/, "").replace("%d", status), url].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
}

ClearAutocomplete = function () {
	for (let dl = document.getElementById("AddressList"); dl.lastChild;) {
		dl.removeChild(dl.lastChild);
	}
	g_.Autocomplete.Path = "";
}

OpenXmlUI = async function (strFile, bAppData, bEmpty, strInit) {
	const s = await ReadXmlFile(strFile, bAppData, bEmpty, strInit);
	if (s) {
		if (window.DOMParser) {
			return new DOMParser().parseFromString(s, "application/xml");
		}
		const xml = api.CreateObject("Msxml2.DOMDocument");
		xml.async = false;
		xml.loadXML(s);
		return xml;
	}
}

CreateXmlUI = function (bRoot) {
	if (window.DOMParser) {
		const xml = new DOMParser().parseFromString('<TablacusExplorer/>', "application/xml");
		if (!bRoot) {
			xml.removeChild(xml.documentElement);
		}
		return xml;
	}
	return CreateXml(bRoot);
}

SaveXmlExUI = async function (fn, xml) {
	fn = BuildPath(await te.Data.DataFolder, "config\\" + fn);
	const r = (await WriteTextFile(fn, window.XMLSerializer ? new XMLSerializer().serializeToString(xml) : xml.xml) || "").split("\t");
	if (r[0] && r[0] != E_ACCESSDENIED) {
		const e = await api.CreateObject("Object");
		e.message = r[1];
		ShowError(e, [await GetText("Save"), fn].join(": "));
	}
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

MenuGetElementsByTagNameUI = async function (Name) {
	if (window.chrome) {
		const xml = new DOMParser().parseFromString(await te.Data.xmlMenus.xml, "application/xml");
		let menus = xml.getElementsByTagName(Name);
		if (!menus || !menus.length) {
			const altMenu = {
				"ViewContext": "Background",
				"Background": "ViewContext",
				"TaskTray": "Systray",
				"Systray": "TaskTray"
			}
			menus = xml.getElementsByTagName(altMenu[Name]);
		}
		return menus;
	}
	return teMenuGetElementsByTagName(Name);
}

SyncExec = async function (cb, o, n) {
	let pt;
	if (n) {
		pt = await GetPosEx(o, n);
	} else if (o && (o.target || o.srcElement)) {
		pt = await api.Memory("POINT");
		pt.x = o.screenX;
		pt.y = o.screenY;
	}
	const FV = await GetFolderView(o);
	await FV.Focus();
	cb(FV, pt);
}

GetWidth = function (s) {
	return (s && GetNum(s) == s) ? (s + "px") : s;
}

KeyDownEvent = function (ev, vEnter, vCancel) {
	if (vEnter != null && ev.keyCode ? ev.keyCode == VK_RETURN : /^Enter/i.test(ev.key)) {
		if (/object|function/.test(typeof vEnter)) {
			vEnter();
			return false;
		} else {
			return vEnter;
		}
	}
	if (vCancel && ev.keyCode ? ev.keyCode == VK_ESCAPE : /^Esc/i.test(ev.key)) {
		if (/object|function/.test(typeof vCancel)) {
			vCancel();
			return false;
		} else {
			return vCancel;
		}
	}
	return true;
}

GetIconSizeEx = function (item) {
	return GetIconSize(item.getAttribute("IconSize"), /Inner/i.test(item.getAttribute("Location")) && ui_.InnerIconSize);
}

ConfirmThenExec = async function (msg, fn, arg) {
	const el = document.activeElement;
	if (!el || el == ui_.elConfirm && ui_.ConfirmMenu && (new Date().getTime() - ui_.ConfirmMenu) < 500) {
		return;
	}
	const hMenu = await api.CreatePopupMenu();
	await api.InsertMenu(hMenu, 0, MF_BYPOSITION | MF_STRING, 1, await msg);
	await api.InsertMenu(hMenu, 1, MF_BYPOSITION | MF_STRING, 0, await GetText("Cancel"));
	const pt = GetPos(el || o, 9);
	const n = await api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, ui_.hwnd, null, null);
	api.DestroyMenu(hMenu);
	if (n) {
		return fn.apply(fn, arg || []);
	}
	ui_.ConfirmMenu = new Date().getTime();
	ui_.elConfirm = el;
}

SetMiscIcon = function (n, s) {
	ui_.MiscIcon[n] = s;
}
