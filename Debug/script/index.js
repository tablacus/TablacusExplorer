// Tablacus Explorer

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
	ResizeSideBar("Left", h);
	ResizeSideBar("Right", h);
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

function ResizeSideBar(z, h) {
	var o = g_.Locations;
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
	var th = Math.round(Math.max(h, 0));
	o.style.height = th + "px";

	var h2 = Math.max(o.clientHeight - document.getElementById(z + "Bar1").offsetHeight - document.getElementById(z + "Bar3").offsetHeight, 0);
	document.getElementById(z + "Bar2").style.height = h2 + "px";
}

function ResetScroll () {
	if (document.documentElement && document.documentElement.scrollLeft) {
		document.documentElement.scrollLeft = 0;
	}
}


function PanelCreated(Ctrl) {
	RunEvent1("PanelCreated", Ctrl);
}

GetAddonLocation = function (strName) {
	var items = te.Data.Addons.getElementsByTagName(strName);
	return (GetLength(items) ? items[0].getAttribute("Location") : null);
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
			o.style.display = (ui_.IEVer >= 8 && SameText(o.tagName, "td")) ? "table-cell" : "block";
			if (strVAlign && !o.style.verticalAlign) {
				o.style.verticalAlign = strVAlign;
			}
		} else if (Location == "Inner") {
			AddEvent("PanelCreated", function (Ctrl) {
				SetAddon(null, "Inner1Left_" + Ctrl.Id, Tag.replace(/\$/g, Ctrl.Id));
			});
		}
		if (strName) {
			if (!g_.Locations[Location]) {
				g_.Locations[Location] = api.CreateObject("Array");
			}
			var res = /<img.*?src=["'](.*?)["']/i.exec(String(Tag));
			if (res) {
				strName += "\t" + res[1];
			}
			g_.Locations[Location].push(strName);
		}
	}
	return Location;
}

Resize = UI.Resize = function () {
	if (!ui_.tidResize) {
		clearTimeout(ui_.tidResize);
	}
	ui_.tidResize = setTimeout(Resize2, 500);
}

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
			if (SameText(FV.FolderItem.Path, GetParentFolderName(g_.NewItemPath))) {
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

UI.InitWindow = function (cb, cb2) {
	setTimeout(function () {
		Resize();
		(cb)();
		setTimeout(function () {
			ApplyLang(document);
			Resize();
			(cb2)();
		}, 500);
	}, 99);
}

UI.ExitFullscreen = function () {
	if (document.msExitFullscreen) {
		document.msExitFullscreen();
	}
}

importJScript = $.importScript;

te.OnArrange = function (Ctrl, rc, cb) {
	var Type = Ctrl.Type;
	if (Type == CTRL_TE) {
		ui_.TCPos = {};
	}
	RunEvent1("Arrange", Ctrl, rc, cb);
	if (Type == CTRL_TC) {
		var Id = Ctrl.Id;
		var o = ui_.Panels[Id];
		if (!o) {
			var s = ['<table id="Panel_$" class="layout" style="position: absolute; z-index: 1;">'];
			s.push('<tr><td id="InnerLeft_$" class="sidebar" style="width: 0; display: none; overflow: auto"></td><td style="width: 100%"><div id="InnerTop_$" style="display: none"></div>');
			s.push('<table id="InnerTop2_$" class="layout">');
			s.push('<tr><td id="Inner1Left_$" class="toolbar1"></td><td id="Inner1Center_$" class="toolbar2" style="white-space: nowrap"></td><td id="Inner1Right_$" class="toolbar3"></td></tr></table>');
			s.push('<table id="InnerView_$" class="layout" style="width: 100%"><tr><td id="Inner2Left_$" style="width: 0"></td><td id="Inner2Center_$"></td><td id="Inner2Right_$" style="width: 0; overflow: auto"></td></tr></table>');
			s.push('<div id="InnerBottom_$"></div></td><td id="InnerRight_$" class="sidebar" style="width: 0; display: none"></td></tr></table>');
			document.getElementById("Panel").insertAdjacentHTML("BeforeEnd", s.join("").replace(/\$/g, Id));
			PanelCreated(Ctrl);
			o = document.getElementById("Panel_" + Id);
			ui_.Panels[Id] = o;
			ApplyLang(o);
			ChangeView(Ctrl.Selected);
		}
		o.style.left = rc.left + "px";
		o.style.top = rc.top + "px";
		if (Ctrl.Visible) {
			var s = [Ctrl.Left, Ctrl.Top, Ctrl.Width, Ctrl.Height].join(",");
			if (ui_.TCPos[s] && ui_.TCPos[s] != Id) {
				Ctrl.Close();
				return;
			} else {
				ui_.TCPos[s] = Id;
			}
			o.style.display = (ui_.IEVer >= 8 && SameText(o.tagName, "td")) ? "table-cell" : "block";
		} else {
			o.style.display = "none";
		}
		o.style.width = Math.max(rc.right - rc.left, 0) + "px";
		o.style.height = Math.max(rc.bottom - rc.top, 0) + "px";
		rc.top = rc.top + document.getElementById("InnerTop_" + Id).offsetHeight + document.getElementById("InnerTop2_" + Id).offsetHeight;
		var w1 = 0, w2 = 0, x = '';
		for (var i = 0; i <= 1; ++i) {
			w1 += Number(document.getElementById("Inner" + x + "Left_" + Id).style.width.replace(/\D/g, "")) || 0;
			w2 += Number(document.getElementById("Inner" + x + "Right_" + Id).style.width.replace(/\D/g, "")) || 0;
			x = '2';
		}
		rc.left = rc.left + w1;
		rc.right = rc.right - w2;
		rc.bottom = rc.bottom - document.getElementById("InnerBottom_" + Id).offsetHeight;
		o = document.getElementById("Inner2Center_" + Id).style;
		o.width = Math.max(rc.right - rc.left, 0) + "px";
		o.height = Math.max(rc.bottom - rc.top, 0) + "px";
		(cb)(Ctrl, rc);
	}
}

g_.event.windowregistered = function (Ctrl) {
	if (ui_.bWindowRegistered) {
		RunEvent1("WindowRegistered", Ctrl);
	}
	ui_.bWindowRegistered = true;
}


ArrangeAddons = function () {
	g_.Locations = api.CreateObject("Object");
	$.IconSize = te.Data.Conf_IconSize || screen.deviceYDPI / 4;
	var xml = OpenXml("addons.xml", false, true);
	te.Data.Addons = xml;
	if (api.GetKeyState(VK_SHIFT) < 0 && api.GetKeyState(VK_CONTROL) < 0) {
		IsSavePath = function (path) {
			return false;
		}
		return;
	}
	var AddonId = [];
	var root = xml.documentElement;
	if (root) {
		var items = root.childNodes;
		if (items) {
			var arError = api.CreateObject("Array");
			for (var i = 0; i < GetLength(items); ++i) {
				var item = items[i];
				var Id = item.nodeName;
				g_.Error_source = Id;
				if (!AddonId[Id]) {
					var Enabled = GetNum(item.getAttribute("Enabled"));
					if (Enabled & 6) {
						LoadLang2(BuildPath(te.Data.Installed, "addons", Id, "lang", GetLangId() + ".xml"));
					}
					if (Enabled & 8) {
						LoadAddon("vbs", Id, arError);
					}
					if (Enabled & 1) {
						LoadAddon("js", Id, arError);
					}
					AddonId[Id] = true;
				}
				g_.Error_source = "";
			}
			if (arError.length || arError.Count) {
				setTimeout(function (arError) {
					if (MessageBox(arError.join("\n\n"), TITLE, MB_OKCANCEL) != IDCANCEL) {
						te.Data.bErrorAddons = true;
						ShowOptions("Tab=Add-ons");
					}
				}, 500, arError);
			}
		}
	}
	RunEventUI("BrowserCreatedEx");
	RunEvent1("BrowserCreated", document);
	var cl = GetWinColor(window.getComputedStyle ? getComputedStyle(document.body).getPropertyValue('background-color') : document.body.currentStyle.backgroundColor);
	ArrangeAddons1(cl);
}

// Events

AddEvent("VisibleChanged", function (Ctrl) {
	if (Ctrl.Type == CTRL_TC) {
		var o = ui_.Panels[Ctrl.Id];
		if (o) {
			if (Ctrl.Visible) {
				o.style.display = (ui_.IEVer >= 8 && SameText(o.tagName, "td")) ? "table-cell" : "block";
				ChangeView(Ctrl.Selected);
			} else {
				o.style.display = "none";
			}
		}
	}
});

AddEvent("SystemMessage", function (Ctrl, hwnd, msg, wParam, lParam) {
	if (Ctrl.Type == CTRL_WB) {
		if (msg == WM_KILLFOCUS) {
			var o = document.activeElement;
			if (o) {
				var s = o.style.visibility;
				o.style.visibility = "hidden";
				o.style.visibility = s;
				FireEvent(o, "blur");
			}
		}
	}
});

// Browser Events

AddEventEx(window, "load", function () {
	ApplyLang(document);
	if (api.GetKeyState(VK_SHIFT) < 0 && api.GetKeyState(VK_CONTROL) < 0) {
		ShowOptions("Tab=Add-ons");
	}
});

AddEventEx(window, "resize", Resize);

AddEventEx(window, "beforeunload", Finalize);

AddEventEx(window, "blur", ResetScroll);

AddEventEx(document, "MSFullscreenChange", function () {
	FullscreenChanged(document.msFullscreenElement != void 0);
});

(function () {
	UI.OnLoad();
	InitCode();
	DefaultFont = $.DefaultFont;
	HOME_PATH = $.HOME_PATH;
	InitMouse();
	OpenMode = $.OpenMode;
	InitMenus();
	LoadLang();
	ArrangeAddons();
	setTimeout(InitWindow, 9);
})();
