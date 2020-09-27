// Tablacus Explorer

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
			o = document.createElement("table");
			o.id = "Panel_" + Id;
			o.className = "layout";
			o.style.position = "absolute";
			o.style.zIndex = 1;
			var s = '<tr><td id="InnerLeft_$" class="sidebar" style="width: 0; display: none; overflow: auto"></td><td style="width: 100%"><div id="InnerTop_$" style="display: none"></div>';
			s += '<table id="InnerTop2_$" class="layout"><tr><td id="Inner1Left_$" class="toolbar1"></td><td id="Inner1Center_$" class="toolbar2" style="white-space: nowrap"></td><td id="Inner1Right_$" class="toolbar3"></td></tr></table>';
			s += '<table id="InnerView_$" class="layout" style="width: 100%"><tr><td id="Inner2Left_$" style="width: 0"></td><td id="Inner2Center_$"></td><td id="Inner2Right_$" style="width: 0; overflow: auto"></td></tr></table>';
			s += '<div id="InnerBottom_$"></div></td><td id="InnerRight_$" class="sidebar" style="width: 0; display: none"></td></tr>';
			o.innerHTML = s.replace(/\$/g, Id);
			document.getElementById("Panel").appendChild(o);
			PanelCreated(Ctrl);
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
			o.style.display = (ui_.IEVer >= 8 && o.tagName.toLowerCase() == "td") ? "table-cell" : "block";
		} else {
			o.style.display = "none";
		}
		o.style.width = Math.max(rc.right - rc.left, 0) + "px";
		o.style.height = Math.max(rc.bottom - rc.top, 0) + "px";
		rc.top = rc.top + document.getElementById("InnerTop_" + Id).offsetHeight + document.getElementById("InnerTop2_" + Id).offsetHeight;
		var w1 = 0, w2 = 0, x = '';
		for (var i = 0; i <= 1; ++i) {
			w1 += api.LowPart(document.getElementById("Inner" + x + "Left_" + Id).style.width.replace(/\D/g, ""));
			w2 += api.LowPart(document.getElementById("Inner" + x + "Right_" + Id).style.width.replace(/\D/g, ""));
			x = '2';
		}
		rc.left = rc.left + w1;
		rc.right = rc.right - w2;
		rc.bottom = rc.bottom - document.getElementById("InnerBottom_" + Ctrl.Id).offsetHeight;
		o = document.getElementById("Inner2Center_" + Ctrl.Id).style;
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

UI.OnLoad();

ArrangeAddons = function () {
	ui_.Locations = {};
	$.IconSize = te.Data.Conf_IconSize || screen.logicalYDPI / 4;
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
			for (var i = 0; i < api.ObjGetI(items, "length"); ++i) {
				var item = items[i];
				var Id = item.nodeName;
				g_.Error_source = Id;
				if (!AddonId[Id]) {
					var Enabled = api.LowPart(item.getAttribute("Enabled"));
					if (Enabled & 6) {
						LoadLang2(fso.BuildPath(te.Data.Installed, "addons\\" + Id + "\\lang\\" + GetLangId() + ".xml"));
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
	UI.RunEvent1("BrowserCreated", document);
	$.ArrangeAddons1();
}

// Events

AddEvent("VisibleChanged", function (Ctrl) {
	if (Ctrl.Type == CTRL_TC) {
		var o = ui_.Panels[Ctrl.Id];
		if (o) {
			if (Ctrl.Visible) {
				o.style.display = (ui_.IEVer >= 8 && o.tagName.toLowerCase() == "td") ? "table-cell" : "block";
				ChangeView(Ctrl.Selected);
			} else {
				o.style.display = "none";
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

//
InitCode();
DefaultFont = $.DefaultFont;
HOME_PATH = $.HOME_PATH;
InitMouse();
OpenMode = $.OpenMode;
InitMenus();
LoadLang();
ArrangeAddons();
setTimeout(InitWindow);
