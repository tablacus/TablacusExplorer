//Tablacus Explorer

Addon = 1;

function Resize2() {
	if (ui_.tidResize) {
		clearTimeout(ui_.tidResize);
		ui_.tidResize = void 0;
	}
	ResetScroll();
	var o = document.getElementById("toolbar");
	var offsetTop = o ? o.offsetHeight : 0;

	var h = 0;
	o = document.getElementById("bottombar");
	var offsetBottom = o.offsetHeight;
	o = document.getElementById("client");
	var ode = document.documentElement || document.body;
	if (o) {
		h = ode.offsetHeight - offsetBottom - offsetTop;
		if (h < 0) {
			h = 0;
		}
		o.style.height = h + "px";
	}
	ResizeSizeBar("Left", h);
	ResizeSizeBar("Right", h);
	o = document.getElementById("Background");
	pt = GetPos(o);
	te.offsetLeft = pt.x;
	te.offsetRight = ode.offsetWidth - o.offsetWidth - te.offsetLeft;
	te.offsetTop = pt.y;
	pt = GetPos(document.getElementById("bottombar"));
	te.offsetBottom = ode.offsetHeight - pt.y;
	RunEvent1("Resize");
	api.PostMessage(te.hwnd, WM_SIZE, 0, 0);
}

function ResizeSizeBar (z, h) {
	var o = ui_.Locations;
	var w = (o[z + "Bar1"] || o[z + "Bar2"] || o[z + "Bar3"]) ? te.Data["Conf_" + z + "BarWidth"] : 0;
	o = document.getElementById(z.toLowerCase() + "bar");
	if (w > 0) {
		o.style.display = "";
		if (w != o.offsetWidth) {
			o.style.width = w + "px";
			for (var i = 1; i <= 3; ++i) {
				document.getElementById(z + "Bar" + i).style.width = w + "px";
			}
			document.getElementById(z.toLowerCase() + "barT").style.width = w + "px";
		}
	} else {
		o.style.display = "none";
	}
	document.getElementById(z.toLowerCase() + "splitter").style.display = w ? "" : "none";

	o = document.getElementById(z.toLowerCase() + "barT");
	var i = h;
	o.style.height = Math.max(i, 0) + "px";
	i = o.clientHeight - o.style.height.replace(/\D/g, "");

	var h2 = o.clientHeight - document.getElementById(z + "Bar1").offsetHeight - document.getElementById(z + "Bar3").offsetHeight;
	document.getElementById(z + "Bar2").style.height = Math.abs(h2 - i) + "px";
}

function ResetScroll() {
	if (document.documentElement && document.documentElement.scrollLeft) {
		document.documentElement.scrollLeft = 0;
	}
}

function PanelCreated (Ctrl) {
	UI.RunEvent1("PanelCreated", Ctrl);
}

GetAddonLocation = function (strName) {
	var items = te.Data.Addons.getElementsByTagName(strName);
	return (api.ObjGetI(items, "length") ? items[0].getAttribute("Location") : null);
}

SetAddon = function (strName, Location, Tag, strVAlign) {
	if (strName) {
		var s = GetAddonLocation(strName);
		if (s) {
			Location = s;
		}
	}
	if (Tag) {
		if (Tag.join) {
			Tag = Tag.join("");
		}
		var o = document.getElementById(Location);
		if (o) {
			if ("string" === typeof Tag) {
				o.insertAdjacentHTML("BeforeEnd", Tag);
			} else {
				o.appendChild(Tag);
			}
			o.style.display = (ui_.IEVer >= 8 && o.tagName.toLowerCase() == "td") ? "table-cell" : "block";
			if (strVAlign && !o.style.verticalAlign) {
				o.style.verticalAlign = strVAlign;
			}
		} else if (Location == "Inner") {
			AddEvent("PanelCreated", function (Ctrl) {
				SetAddon(null, "Inner1Left_" + Ctrl.Id, Tag.replace(/\$/g, Ctrl.Id));
			});
		}
		if (strName) {
			if (!ui_.Locations[Location]) {
				ui_.Locations[Location] = [];
			}
			var res = /<img.*?src=["'](.*?)["']/i.exec(String(Tag));
			if (res) {
				strName += "\0" + res[1];
			}
			ui_.Locations[Location].push(strName);
		}
	}
	return Location;
}

UI.Resize = function () {
	if (!ui_.tidResize) {
		clearTimeout(ui_.tidResize);
	}
	ui_.tidResize = setTimeout(Resize2, 500);
}

Resize = UI.Resize;

UI.OpenInExplorer = function (Path) {
	setTimeout(function (Path) {
		$.OpenInExplorer(Path);
	}, 99, Path);
}

UI.StartGestureTimer = function () {
	var i = te.Data.Conf_GestureTimeout;
	if (i) {
		clearTimeout(g_.mouse.tidGesture);
		g_.mouse.tidGesture = setTimeout(function () {
			g_.mouse.EndGesture(true);
		}, i);
	}
}

UI.Focus = function (o, tm) {
	if (o) {
		setTimeout(function () {
			o.focus();
		}, tm)
	}
}

UI.FocusFV = function (tm) {
	setTimeout(function () {
		var hFocus = api.GetFocus();
		if (!hFocus || hFocus == te.hwnd) {
			var FV = te.Ctrl(CTRL_FV);
			if (FV) {
				FV.Focus();
			}
		}
	}, tm || ui_.DoubleClickTime);
}

UI.FocusWB = function () {
	var o = document.activeElement;
	if (o) {
		if (!/input|textarea/i.test(o.tagName)) {
			setTimeout(function () {
				if (o === document.activeElement) {
					GetFolderView().Focus();
				}
			}, ui_.DoubleClickTime, o);
		}
	}
}

UI.SelectItem = function (FV, path, wFlags, tm) {
	setTimeout(function () {
		FV.SelectItem(path, wFlags);
	}, tm);
}

UI.SelectNewItem = function () {
	setTimeout(function () {
		var FV = te.Ctrl(CTRL_FV);
		if (FV) {
			if (!api.StrCmpI(FV.FolderItem.Path, fso.GetParentFolderName(g_.NewItemPath))) {
				FV.SelectItem(g_.NewItemPath, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_DESELECTOTHERS | SVSI_SELECTIONMARK | SVSI_SELECT);
			}
		}
	}, 800);
}

UI.SelectNext = function (FV) {
	setTimeout(function () {
		if (!api.SendMessage(FV.hwndList, LVM_GETEDITCONTROL, 0, 0) || WINVER < 0x600) {
			FV.SelectItem(FV.Item(FV.GetFocusedItem + (api.GetKeyState(VK_SHIFT) < 0 ? -1 : 1)) || api.GetKeyState(VK_SHIFT) < 0 ? FV.ItemCount(SVGIO_ALLVIEW) - 1 : 0, SVSI_EDIT | SVSI_FOCUSED | SVSI_SELECT | SVSI_DESELECTOTHERS);
		}
	}, 99);
}

UI.EndMenu = function () {
	setTimeout(function () {
		api.EndMenu();
	}, 200);
}

UI.ShowStatusText = function (Ctrl, Text, iPart, tm) {
	if (ui_.Status && ui_.Status[5]) {
		clearTimeout(ui_.Status[5]);
		delete ui_.Status;
	}
	ui_.Status = [Ctrl, Text, iPart, tm, new Date().getTime(), setTimeout(function () {
		if (ui_.Status) {
			if (new Date().getTime() - ui_.Status[4] > ui_.Status[3] / 2) {
				$.ShowStatusText(ui_.Status[0], ui_.Status[1], ui_.Status[2]);
				delete ui_.Status;
			}
		}
	}, tm)];
}

UI.CancelWindowRegistered = function () {
	clearTimeout(ui_.tidWindowRegistered);
	ui_.bWindowRegistered = false;
	ui_.tidWindowRegistered = setTimeout(function () {
		ui_.bWindowRegistered = true;
	}, 9999);
}

UI.ExecGesture = function (Ctrl, hwnd, pt, str) {
	setTimeout(function () {
		g_.mouse.Exec(Ctrl, hwnd, pt, str);
	}, 99);
}
;
(function () {
	var args = ["sync.js", "sync1.js"];
	if (window.chrome) {
		args.unshift("consts.js", "common.js");
		var s = ["MainWindow = window;"];
		var arFN = [];
		var line = 1;
		for (var i = 0; i < args.length; i++) {
			var fn = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "script\\" + args[i]);
			var ado = api.CreateObject("ads");
			if (ado) {
				ado.CharSet = "utf-8";
				ado.Open();
				ado.LoadFromFile(fn);
				var src = ado.ReadText();
				s.push(src);
				ado.Close();
				if (window.chrome && src && /sync/.test(fn)) {
					arFN = arFN.concat(src.match(/\s\w+\s*=\s*function/g).map(function (s) {
						return s.replace(/^\s|\s*=.*$/g, "");
					}));
				}
				api.OutputDebugString([fso.GetFileName(args[i]), "Start line:", line, "\n"].join(""));
				line += src.split("\n").length;
			}
		}
		var CopyObj = function (to, o, ar) {
			if (!to) {
				to = api.CreateObject("Object");
			}
			ar.forEach(function (key) {
				to[key] = o[key];
			});
			return to;
		}
		document.documentMode = (/Edg\/(\d+)/.test(navigator.appVersion) ? RegExp.$1 : 12);
		CopyObj($, window, ["te", "api", "chrome", "document", "UI"]);
		$.location = CopyObj(null, location, ["hash", "href"]);
		$.navigator = CopyObj(null, navigator, ["appVersion", "language"]);
		$.screen = CopyObj(null, screen, ["deviceXDPI", "deviceYDPI"]);

		var o = api.CreateObject("Object");
		o.window = $;
		$.$JS = api.GetScriptDispatch(s.join(""), "JScript", o);
		CopyObj(window, $, ["g_", "Addons", "eventTE", "eventTA", "Threads", "SimpleDB", "BasicDB;"]);
		CopyObj(window, $, arFN);
	} else {
		$ = window;
		for (var i = 0; i < args.length; ++i) {
			var el = document.createElement("script");
			el.src = args[i];
			document.body.appendChild(el);
		}
	}
	MainWindow = $;
	var el = document.createElement("script");
	el.src = "ui1.js";
	document.body.appendChild(el);
})();

