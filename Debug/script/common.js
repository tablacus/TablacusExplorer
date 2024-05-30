// Tablacus Explorer

g_sep = "` ~";

if ("async ") {
	AsyncFunction = Object.getPrototypeOf(async function () { }).constructor;
} else {
	AsyncFunction = function (s) {
		return Function(FixScript(s));
	};
}

importScript = async function (fn) {
	let hr = E_FAIL, s;
	if (window.ReadTextFile) {
		s = await ReadTextFile(fn);
	} else {
		if (!/^[A-Z]:\\|^\\\\\w/i.test(fn)) {
			fn = BuildPath(GetParentFolderName(await api.GetModuleFileName(null)), fn);
		}
		let ado = await api.CreateObject("ads");
		ado.CharSet = "utf-8";
		await ado.Open();
		await ado.LoadFromFile(fn);
		s = await ado.ReadText();
		ado.Close();
	}
	if (s) {
		if (/\.vbs$/i.test(fn)) {
			hr = ExecScriptEx(window.Ctrl, s, "VBScript", await $.pt, await $.dataObj, await $.grfKeyState, await $.pdwEffect, await $.bDrop);
		} else {
			await new AsyncFunction(s)();
			hr = S_OK;
		}
	}
	return hr;
}

if (!window.InitUI && !window.chrome) {
	if (window.alert) {
		importScript("script\\ui.js");
		InitUI();
	}
	if (!window.g_) {
		importScript("script\\sync.js");
	}
}

Invoke = async function (args, cb) {
	const ar = args.shift().split(".");
	let fn = window, s, parent;
	while (s = ar.shift()) {
		parent = fn;
		fn = fn[s];
		if (!fn) {
			return;
		}
	}
	if (cb) {
		InvokeFunc(cb, [await fn.apply(parent, args)]);
		return;
	}
	fn.apply(parent, args);
}

AddEvent = function (Name, fn, priority) {
	if (/^arrange$|^finalize$|^layout$|^load$|^panelcreated$|^resize$/i.test(Name)) {
		InvokeUI("AddEventUI", Array.apply(null, arguments));
		if (/^finalize$/.test(Name)) {
			return;
		}
	}
	AddEvent2(Name, fn, priority);
}

ClearEvent = function (Name) {
	if (/^arrange$|^finalize$|^layout$|^load$|^panelcreated$|^resize$/i.test(Name)) {
		InvokeUI("ClearEventUI", [Name]);
		if (/^finalize$/.test(Name)) {
			return;
		}
	}
	ClearEvent2(Name);
}

SameText = function (s1, s2) {
	return String(s1).toLowerCase() == String(s2).toLowerCase();
}

StartsText = function (s1, s2) {
	return String(s1).toLowerCase() == String(s2).slice(0, s1.length).toLowerCase();
}

GetLength = async function (o) {
	return o ? (o.length || await api.ObjGetI(o, "length")) : 0;
}

GetFileName = function (s) {
	const res = /([^\\\/]*)$/.exec(s);
	return res ? res[1] : "";
}

PathQuoteSpaces = function (s) {
	return /\s/.test(s) ? '"' + s + '"' : s;
}

PathUnquoteSpaces = function (s) {
	const res = /^"(.*)"\0?$/.exec(s);
	return res ? res[1] : s;
}

CreateTab = async function (Ctrl, pt) {
	const FV = await GetFolderView(Ctrl, pt);
	HOME_PATH = await $.HOME_PATH;
	NavigateFV(FV, HOME_PATH || await api.ILIsEqual(FV, "about:blank") ? HOME_PATH : FV, SBSP_NEWBROWSER);
	return S_OK;
}

StripAmp = function (s) {
	return String(s).replace(/\(&\w\)|&/, "").replace(/ ?\.\.\.$/, "");
}

EncodeSC = function (s) {
	return String(s).replace(/[&"<>]/g, function (strMatch) {
		return "&#" + strMatch.charCodeAt(0) + ";";
	});
}

DecodeSC = function (s) {
	return String(s).replace(/&([\w#]{1,5});/g, function (strMatch, ref) {
		let res = /^#x([\da-f]+)$/i.exec(ref)
		if (res) {
			return String.fromCodePoint(parseInt(res[1], 16));
		}
		res = /^#(\d+)$/.exec(ref)
		if (res) {
			return String.fromCodePoint(res[1]);
		}
		return { quot: '"', amp: '&', lt: '<', gt: '>' }[ref.toLowerCase()] || '&' + ref + ';';
	});
}

ExtractAttr = function (o, ar, re) {
	for (let n in o) {
		const s = o[n];
		if (s != null && (!re || !re.test(s))) {
			ar.push(' ', n, '="', EncodeSC(StripAmp(s)).replace(/"/g, ""), '"');
		}
	}
}

AddClass = function (o, s) {
	const ar = [];
	if (o.className || o["class"]) {
		ar.push(o.className || o["class"]);
	}
	ar.push(s);
	o["class"] = ar.join(" ");
}

GetImgTag = async function (o, h) {
	if (o.onclick) {
		if (!o.onmouseover) {
			o.onmouseover = "MouseOver(this)";
		}
		if (!o.nmouseout) {
			o.onmouseout = "MouseOut()";
		}
	}
	if (o.src = PathUnquoteSpaces(o.src)) {
		let res = !(window.ui_ || te.Data).NoCssFont && /^font:([^,]*),(.+)/i.exec(await MainWindow.RunEvent4("ReplaceIcon", o.src) || o.src);
		if (res) {
			const FontFace = res[1].replace(/\"/g, '\\"');
			let c = res[2];
			if (/[\da-fx,]+/.test(c)) {
				c = c.split(",");
				c = String.fromCodePoint(c.length > 1 ? parseInt(c[0]) * 256 + parseInt(c[1]) : parseInt(c[0]));
			}
			h = h || window.IconSize;
			let h2 = Number(h) ? await CalcFontSize(FontFace, h, c) + "px" : EncodeSC(h);
			if (h2 != "0px") {
				h = Number(h) ? h + "px" : EncodeSC(h);
				const ar = ['font-family:', FontFace, '; font-size:', h2, '; line-height:', h, ';', (o.style || "") ];
				o.style = ar.join("");
				AddClass(o, "fonticon");
			} else if (o.src == o.title) {
				o.style = "display: none";
			} else {
				c =  o.title;
			}
			const ar = ['<span'];
			ExtractAttr(o, ar, /src/i);
			ar.push('>', c, '</span>');
			return ar.join("");
		}
		o.org = o.src;
		res = (window.chrome || g_.IEVer > 8) && /\.svg$/i.test(o.src);
		if (!res) {
			o.src = await ImgBase64(o, 0, Number(h));
			res = (window.chrome || g_.IEVer > 8) && /\.svg$/i.test(o.src);
		}
		if (res) {
			AddClass(o, "svgicon");
			o.org = void 0;
			const ar = ['<span'];
			h = h || window.IconSize;
			h = Number(h) ? h + "px" : EncodeSC(h);
			ExtractAttr(o, ar, /src/i);
			ar.push('>');
			if (res = /(<svg)([\w\W]*?>)([\w\W]*?<\/svg[^>]*>)/i.exec(await ReadTextFile(o.src))) {
				ar.push(res[1], ' style="max-width:' + h + ';height:' + h + '" ', res[2].replace(/\s+width="[^"]*"|\s+height="[^"]*"/ig, ""), res[3]);
			}
			ar.push("</span>");
			return ar.join("");
		}
		if (!o.draggable) {
			o.draggable = o.ondragstart != null;
		}
		const ar = ['<img'];
		AddClass(o, "imgicon");
		ExtractAttr(o, ar);
		if (h) {
			h = Number(h) ? h + 'px' : EncodeSC(h);
			ar.push(' width="', h, '" height="', h, '"');
		}
		ar.push('>');
		return ar.join("");
	}
	const ar = ['<span'];
	await ExtractAttr(o, ar, /title/i);
	ar.push('>', EncodeSC(o.title || ""), '</span>');
	return ar.join("");
}

GetIconSize = function (h, a) {
	return h || a * screen.deviceYDPI / 96 || window.IconSize;
}

GetGestureX = async function (ar) {
	let o, s = "";
	while (o = await ar.shift()) {
		if (await api.GetKeyState(o[1]) < 0) {
			s += o[0];
		}
	}
	return s;
}

GetGestureKey = async function () {
	return await GetGestureX([
		["S", VK_SHIFT],
		["C", VK_CONTROL],
		["A", VK_MENU]
	]);
}

GetGestureButton = async function () {
	return await GetGestureX([
		["1", VK_LBUTTON],
		["2", VK_RBUTTON],
		["3", VK_MBUTTON],
		["4", VK_XBUTTON1],
		["5", VK_XBUTTON2]
	]);
}

GetWebColor = function (c) {
	const res = /(\d{1,3}) (\d{1,3}) (\d{1,3})/.exec(c);
	if (res) {
		c = res[3] * 65536 + res[2] * 256 + res[1] * 1;
	}
	return isNaN(c) && /^#[0-9a-f]{3,6}$/i ? c : "#" + (('00000' + (((c & 0xff) << 16) | (c & 0xff00) | ((c & 0xff0000) >> 16)).toString(16)).slice(-6));
}

GetWinColor = async function (c) {
	let res = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(c);
	if (res) {
		return Number(["0x", res[3], res[2], res[1]].join(""));
	}
	res = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(c);
	if (res) {
		return Number(["0x", res[3], res[3], res[2], res[2], res[1], res[1]].join(""));
	}
	res = /(\d{1,3})[ ,]+(\d{1,3})[ ,]+(\d{1,3})/.exec(c);
	if (res) {
		return res[3] * 65536 + res[2] * 256 + res[1] * 1;
	}
	try {
		res = /(\d{1,3}) (\d{1,3}) (\d{1,3})/.exec(await wsh.regRead(["HKCU", "Control Panel", "Colors", c].join("\\")));
		if (res) {
			return res[3] * 65536 + res[2] * 256 + res[1] * 1;
		}
	} catch (e) { }
	return c;
}

FindText = async function () {
	if (window.chrome || g_.IEVer < 8) {
		SetModifierKeys(0);
		wsh.SendKeys("^f");
	} else {
		api.OleCmdExec(document, null, 32, 0, 0);
	}
}

SetRenameMenu = function () { }

Alt = function () {
	return S_OK;
}

GetBGRA = function (c, a) {
	return ((c & 0xff) << 16) | (c & 0xff00) | ((c & 0xff0000) >> 16) | a << 24;
}

delamp = function (s) {
	s = s.replace(/&amp;/ig, "&");
	return /;/.test(s) ? s : s.replace(/&/ig, "");
}

GetConsts = function (s) {
	const r = ("string" === typeof s) && window[s.replace(/\s/, "")];
	return /^string$|^number$/.test(typeof r) ? r : s;
}

CalcVersion = function (s) {
	let r = 0;
	let res = /(\d+)\.(\d+)\.(\d+)\.(\d+)/.exec(s);
	if (res) {
		r = "";
		for (let i = 1; i < 5; ++i) {
			r += ('00000' + res[i]).slice(-6);
		}
		return r;
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

LoadImgDll = async function (icon, index) {
	const i4 = (index || 0) * 4;
	let hModule = await api.LoadLibraryEx(BuildPath(system32, icon[i4]), 0, LOAD_LIBRARY_AS_DATAFILE);
	if (!hModule && SameText(icon[i4], "ieframe.dll")) {
		if (icon[i4 + 1] >= 500) {
			hModule = await api.LoadLibraryEx(BuildPath(system32, "browseui.dll"), 0, LOAD_LIBRARY_AS_DATAFILE);
		} else {
			hModule = await api.LoadLibraryEx(BuildPath(system32, "shell32.dll"), 0, LOAD_LIBRARY_AS_DATAFILE);
		}
	}
	return hModule;
}

DeleteItem = async function (path, fFlags) {
	if (/\0/.test(path)) {
		path = path.split("\0");
	}
	if ("string" !== typeof path || await IsExists(path)) {
		return await api.SHFileOperation(FO_DELETE, path, null, fFlags || FOF_SILENT | FOF_NOCONFIRMATION | FOF_NOERRORUI, false);
	}
}

amp2ul = function (s) {
	return s.replace(/&amp;|&/ig, "");
}

GetAddonInfo2 = async function (xml, info, Tag) {
	const items = xml.getElementsByTagName(Tag);
	if (items.length) {
		const item = items[0].childNodes;
		for (let i = 0; i < item.length; ++i) {
			const item1 = item[i];
			const n = item1.tagName;
			const s = item1.text || item1.textContent || "";
			if (/^en/i.test(Tag) && /Name|Description/i.test(n)) {
				const st = await GetAltText(s);
				info[n] = st || s;
				if (st) {
					info['$' + n] = s;
				}
			} else {
				info[n] = s;
			}
		}
	}
}

SetXmlText = function (item, s) {
	if ("string" === typeof item.text) {
		item.text = s;
	} else {
		item.textContent = s;
	}
}

GetWinIcon = function () {
	for (let i = 0; i < arguments.length; i += 2) {
		if (WINVER >= arguments[i]) {
			return arguments[i + 1];
		}
	}
}

GetEncodeType = function (fn) {
	if (/\.jpe?g?$|\.jfif$/i.test(fn)) {
		return "image/jpeg";
	}
	if (/\.gif$/i.test(fn)) {
		return "image/gif";
	}
	return "image/png";
}

EscapeJsonObj = {
	"\\": "\\\\",
	'"': '\\"',
	"/": "\\/",
	"\b": "\\b",
	"\f": "\\f",
	"\n": "\\n",
	"\r": "\\r",
	"\t": "\\t"
}

EscapeJson = function (s) {
	return String(s).replace(/([\\"\/\b\f\n\r\t])/g, function (all, s) {
		return EscapeJsonObj[s];
	});
};

if ("undefined" === typeof JSON) {
	JSON = {
		parse: function (s) {
			return new Function('return ' + (s || {}))();
		},
		stringify: function (o) {
			const ar = [];
			if (Array.isArray(o)) {
				for (let i = 0; i < o.length; ++i) {
					if ("object" === typeof o[i]) {
						ar.push(this.stringify(o[i]));
					} else {
						ar.push('"' + EscapeJson(o[i]) + '"');
					}
				}
				return '[' + ar.join(",") + "]";
			}
			for (let n in o) {
				if ("object" === typeof o[n]) {
					ar.push('"' + EscapeJson(n) + '":' + this.stringify(o[n]));
				} else {
					ar.push('"' + EscapeJson(n) + '":"' + EscapeJson(o[n]) + '"');
				}
			}
			return '{' + ar.join(",") + "}";
		}
	}
}

if ("undefined" === typeof ActiveXObject) {
	ActiveXObject = function (s) {
		return api.CreateObject(s);
	}
}

if (!String.fromCodePoint) {
	String.fromCodePoint = function (n) {
		const cp = n - 0x10000;
		return cp < 0 ? String.fromCharCode(n) : String.fromCharCode(0xd800 | (cp >> 10), 0xDC00 | (cp & 0x3ff));
	}
}

if (!String.prototype.trim) {
	String.prototype.trim = function () {
		return this.replace(/^\s+|\s+$/g, "");
	}
}

if (!Array.isArray) {
	Array.isArray = function (arg) {
		return Object.prototype.toString.call(arg) === '[object Array]';
	};
}
