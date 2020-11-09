// Tablacus Explorer

g_sep = "` ~";

importScript = function (fn) {
	var hr = E_FAIL;
	var s;
	if (window.ReadTextFile) {
		s = ReadTextFile(fn);
	} else {
		if (!/^[A-Z]:\\|^\\\\\w/i.test(fn)) {
			fn = BuildPath(GetParentFolderName(api.GetModuleFileName(null)), fn);
		}
		var ado = api.CreateObject("ads");
		ado.CharSet = "utf-8";
		ado.Open();
		ado.LoadFromFile(fn);
		s = ado.ReadText();
		ado.Close();
	}
	if (s) {
		if (/\.vbs$/i.test(fn)) {
			hr = ExecScriptEx(window.Ctrl, s, "VBScript", $.pt, $.dataObj, $.grfKeyState, $.pdwEffect, $.bDrop);
		} else {
			new Function(window.chrome && window.alert ? "(() => {" + s + "\n})();" : RemoveAsync(s))();
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

GetNum = function (s) {
	return "number" === typeof s ? s : Number("string" === typeof s ? s.replace(/[^\d\-\.].*/, "") : s) || 0;
}

SameText = function (s1, s2) {
	return String(s1).toLowerCase() == String(s2).toLowerCase();
}

GetLength = function (o) {
	return o ? (o.length || api.ObjGetI(o, "length")) : 0;
}

GetFileName = function (s) {
	var res = /([^\\\/]*)$/.exec(s);
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

GetIconSize = function (h, a) {
	return h || a * screen.deviceYDPI / 96 || window.IconSize;
}

GetGestureX = function (ar) {
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
	return isNaN(c) && /^#[0-9a-f]{3,6}$/i ? c : "#" + (('00000' + (((c & 0xff) << 16) | (c & 0xff00) | ((c & 0xff0000) >> 16)).toString(16)).substr(-6));
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
	res = /(\d{1,3})[ ,]+(\d{1,3})[ ,]+(\d{1,3})/.exec(c);
	if (res) {
		return res[3] * 65536 + res[2] * 256 + res[1] * 1;
	}
	try {
		res = /(\d{1,3}) (\d{1,3}) (\d{1,3})/.exec(wsh.regRead(["HKCU", "Control Panel", "Colors", c].join("\\")));
		if (res) {
			return res[3] * 65536 + res[2] * 256 + res[1] * 1;
		}
	} catch (e) { }
	return c;
}

FindText = function () {
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
	var Result = window[s.replace(/\s/, "")];
	if (Result !== void 0) {
		return Result;
	}
	return s;
}

CalcVersion = function (s) {
	var r = 0;
	var res = /(\d+)\.(\d+)\.(\d+)\.(\d+)/.exec(s);
	if (res) {
		var r = "";
		for (var i = 1; i < 5; ++i) {
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

LoadImgDll = function (icon, index) {
	var i4 = (index || 0) * 4;
	api.OutputDebugString([i4, system32, icon[i4]].join(",") + "\n");
	var hModule = api.LoadLibraryEx(BuildPath(system32, icon[i4]), 0, LOAD_LIBRARY_AS_DATAFILE);
	if (!hModule && SameText(icon[i4], "ieframe.dll")) {
		if (icon[i4 + 1] >= 500) {
			hModule = api.LoadLibraryEx(BuildPath(system32, "browseui.dll"), 0, LOAD_LIBRARY_AS_DATAFILE);
		} else {
			hModule = api.LoadLibraryEx(BuildPath(system32, "shell32.dll"), 0, LOAD_LIBRARY_AS_DATAFILE);
		}
	}
	return hModule;
}

amp2ul = function (s) {
	return s.replace(/&amp;|&/ig, "");
}

EscapeJson = function (s) {
	return s.replace(/([\\|"|\/])/g, '\\$1').replace(/[\b]/g, '\\b').replace(/[\f]/g, '\\f').replace(/[\n]/g, '\\n').replace(/[\r]/g, '\\r').replace(/[\t]/g, '\\t');
};

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
						ar.push(this.stringify(o[i]));
					} else {
						ar.push('"' + EscapeJson(o[i]) + '"');
					}
				}
				return '[' + ar.join(",") + "]";
			}
			for (var n in o) {
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
