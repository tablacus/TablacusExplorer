// Tablacus Explorer

(function () {
	system32 = api.GetDisplayNameOf(ssfSYSTEM, SHGDN_FORPARSING);
	hShell32 = api.GetModuleHandle(fso.BuildPath(system32, "shell32.dll"));

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
		delete s;
	}

	if (api.SHTestTokenMembership(null, 0x220) && WINVER >= 0x600) {
		TITLE += ' [' + (api.LoadString(hShell32, 25167) || "Admin").replace(/;.*$/, "") + ']';
	}
})();

importScript = function (fn) {
	var hr = E_FAIL;
	var ado;
	if (window.OpenAdodbFromTextFile) {
		ado = OpenAdodbFromTextFile(fn, "utf-8");
	} else {
		if (!/^[A-Z]:\\|^\\\\\w/i.test(fn)) {
			fn = fso.BuildPath(te.Data.Installed, fn);
		}
		var ado = api.CreateObject("ads");
		ado.CharSet = "utf-8";
		ado.Open();
		ado.LoadFromFile(fn);
	}
	if (ado) {
		if (/\.vbs/i.test(fn)) {
			hr = ExecScriptEx(window.Ctrl, ado.ReadText(), "VBScript", $.pt, $.dataObj, $.grfKeyState, $.pdwEffect, $.bDrop);
		} else {
			new Function(ado.ReadText())();
			hr = S_OK;
		}
		ado.Close();
	}
	return hr;
}

if (!window.UI) {
	importScript("script\\sync.js");
	importScript("script\\ui.js");
}

GetInt = function (s) {
	return "number" === typeof s ? s : Number("string" === typeof s ? s.replace(/\D.*/, "") : s) || 0;
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
	api.OleCmdExec(document, null, 32, 0, 0);
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

createHttpRequest = function () {
	try {
		return window.XMLHttpRequest && ui_.IEVer >= 9 ? new XMLHttpRequest() : api.CreateObject("Msxml2.XMLHTTP");
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
