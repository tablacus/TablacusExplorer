// Tablacus Explorer

g_sep = "` ~";

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
			hr = ExecScriptEx(await window.Ctrl, s, "VBScript", await $.pt, await $.dataObj, await $.grfKeyState, await $.pdwEffect, await $.bDrop);
		} else {
			new Function(FixScript(s, window.chrome && window.alert))();
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

Invoke = function (args) {
	const ar = args.shift().split(".");
	let fn = window, s, parent;
	while (s = ar.shift()) {
		parent = fn;
		fn = fn[s];
		if (!fn) {
			return;
		}
	}
	fn.apply(parent, args);
}

GetNum = function (s) {
	return "number" === typeof s ? s : Number("string" === typeof s ? s.replace(/[^\d\-\.].*/, "") : s) || 0;
}

SameText = function (s1, s2) {
	return String(s1).toLowerCase() == String(s2).toLowerCase();
}

GetLength = async function (o) {
	return o ? (o.length || await api.ObjGetI(o, "length")) : 0;
}

GetFileName = function (s) {
	const res = /([^\\\/]*)$/.exec(s);
	return res ? res[1] : "";
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
		let res = /^#x([\da-f]+)$/i.exec(ref)
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
	return isNaN(c) && /^#[0-9a-f]{3,6}$/i ? c : "#" + (('00000' + (((c & 0xff) << 16) | (c & 0xff00) | ((c & 0xff0000) >> 16)).toString(16)).substr(-6));
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
	if (window.chrome || ui_.IEVer < 8) {
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
	const Result = window[s.replace(/\s/, "")];
	if (Result != null) {
		return Result;
	}
	return s;
}

CalcVersion = function (s) {
	let r = 0;
	let res = /(\d+)\.(\d+)\.(\d+)\.(\d+)/.exec(s);
	if (res) {
		r = "";
		for (let i = 1; i < 5; ++i) {
			r += ('00000' + res[i]).substr(-6);
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
		return api.SHFileOperation(FO_DELETE, path, null, fFlags || FOF_SILENT | FOF_NOCONFIRMATION | FOF_NOERRORUI, false);
	}
}

amp2ul = function (s) {
	return s.replace(/&amp;|&/ig, "");
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
	return s.replace(/([\\"\/\b\f\n\r\t])/g, function (all, s) {
		return EscapeJsonObj[s];
	});
};

GetAddonInfo2 = async function (xml, info, Tag, bTrans) {
	const items = await xml.getElementsByTagName(Tag);
	if (await GetLength(items)) {
		const item = await items[0].childNodes;
		const nLen = await GetLength(item);
		for (let i = 0; i < nLen; ++i) {
			const item1 = await item[i];
			const n = await item1.tagName;
			const s = await item1.text || await item1.textContent;
			info[n] = (bTrans && /Name|Description/i.test(n) ? await GetText(s) : s);
		}
	}
}

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
