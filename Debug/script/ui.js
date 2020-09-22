// Tablacus Explorer

ui_ = {
	Panels: {}
};
UI = api.CreateObject("Object");
UI.Addons = api.CreateObject("Object");

UI.clearTimeout = clearTimeout;

UI.RunEvent1 = function () {
	var args = Array.apply(null, arguments);
	var en = args.shift();
	var eo = eventTE[en.toLowerCase()];
	for (var list = api.CreateObject("Enum", eo); !list.atEnd(); list.moveNext()) {
		var fn = eo[list.item()];
		try {
			fn.apply(fn, args);
		} catch (e) {
			ShowError(e, en, list.item());
		}
	}
}

UI.ShowError = function (e, s, i) {
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

UI.SetDisplay = function (Id, s) {
	var o = document.getElementById(Id);
	if (o) {
		o.style.display = s;
	}
}

UI.ApplyLangTag = function (o) {
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

UI.Blur = function (Id) {
	document.getElementById(Id).blur();
}

function _InvokeMethod() {
	var args;
	var ar = api.ObjGetI(te, "fn");
	while (args = ar.shift()) {
		api.OutputDebugString(["InvokeMethod:", api.ObjGetI(ar, "length"), "\n"].join(","));
		var fn = args.shift();
		api.OutputDebugString(["InvokeMethod:", fn.toString().split("\n")[0]].join(","));
		fn.apply(fn, args);
		api.OutputDebugString("InvokeMethod: OK:" + api.ObjGetI(ar, "length") + "\n");
	}
}

ApplyLang = function (doc) {
	var i, o, h = 0;
	var FaceName = MainWindow.DefaultFont.lfFaceName;
	if (doc.body) {
		doc.body.style.fontFamily = FaceName;
		doc.body.style.fontSize = Math.abs(MainWindow.DefaultFont.lfHeight) + "px";
		doc.body.style.fontWeight = MainWindow.DefaultFont.lfWeight;
	}
	UI.ApplyLangTag(doc.getElementsByTagName("label"));
	UI.ApplyLangTag(doc.getElementsByTagName("button"));
	UI.ApplyLangTag(doc.getElementsByTagName("li"));
	o = doc.getElementsByTagName("a");
	if (o) {
		UI.ApplyLangTag(o);
		for (i = o.length; i--;) {
			if (o[i].className == "treebutton" && o[i].innerHTML == "") {
				o[i].innerHTML = BUTTONS.opened;
			}
		}
	}
	o = doc.getElementsByTagName("input");
	if (o) {
		UI.ApplyLangTag(o);
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
		UI.ApplyLangTag(o);
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

MouseOver = function (o) {
	o = o.srcElement || o;
	if (/^button$|^menu$/i.test(o.className)) {
		if (ui_.objHover && o != ui_.objHover) {
			MouseOut();
		}
		var pt = api.Memory("POINT");
		api.GetCursorPos(pt);
		var ptc = pt.Clone();
		api.ScreenToClient(api.GetWindow(document), ptc);
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

CloseSubWindows = function () {
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

AddEventEx(window, "beforeunload", CloseSubWindows);

UI.OnLoad = function () {
}

//Options
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
		if (g_.IEVer >= 9) {
			o.style.height = "calc(100vh - " + em + "em)";
		} else {
			var h = document.documentElement.clientHeight || document.body.clientHeight;
			h += MainWindow.DefaultFont.lfHeight * em;
			if (h > 0) {
				o.style.height = h + 'px';
				o.style.height = 2 * h - o.offsetHeight + "px";
			}
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

