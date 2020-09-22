//Tablacus Explorer

Addon = 1;

function Resize2() {
	if (ui_.tidResize) {
		clearTimeout(ui_.tidResize);
		ui_.tidResize = void 0;
	}
	if (!te.Ctrl(CTRL_FV)) {
		Resize();
		return;
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
	UI.RunEvent1("Resize");
	api.PostMessage(te.hwnd, WM_SIZE, 0, 0);
}

function ResizeSizeBar (z, h) {
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
	return (items.length ? items[0].getAttribute("Location") : null);
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
			o.style.display = (g_.IEVer >= 8 && o.tagName.toLowerCase() == "td") ? "table-cell" : "block";
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
				g_.Locations[Location] = [];
			}
			var res = /<img.*?src=["'](.*?)["']/i.exec(String(Tag));
			if (res) {
				strName += "\0" + res[1];
			}
			g_.Locations[Location].push(strName);
		}
	}
	return Location;
}

UI.Resize = function () {
	if (!ui_.tidResize) {
		clearTimeout(ui_.tidResize);
		ui_.tidResize = setTimeout(Resize2, 500);
	}
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
	}, tm);
}

UI.FocusWB = function () {
	var o = document.activeElement;
	if (o) {
		if (!/input|textarea/i.test(o.tagName)) {
			setTimeout(function () {
				if (o === document.activeElement) {
					GetFolderView().Focus();
				}
			}, 99, o);
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
	g_.Status = [Ctrl, Text, iPart, tm, new Date().getTime(), setTimeout(function () {
		if (g_.Status) {
			if (new Date().getTime() - g_.Status[4] > g_.Status[3] / 2) {
				$.ShowStatusText(g_.Status[0], g_.Status[1], g_.Status[2]);
				delete g_.Status;
			}
		}
	}, tm)];
}

UI.CancelWindowRegistered = function () {
	clearTimeout(g_.tidWindowRegistered);
	g_.bWindowRegistered = false;
	g_.tidWindowRegistered = setTimeout(function () {
		g_.bWindowRegistered = true;
	}, 9999);
}

UI.ExecGesture = function (Ctrl, hwnd, pt, str) {
	setTimeout(function () {
		g_.mouse.Exec(Ctrl, hwnd, pt, str);
	}, 99);
}

var args = ["script\\sync.js", "script\\sync1.js"];
if (window.chrome) {
	te.fn = api.CreateObject("Array");
	args.unshift("script\\consts.js", "script\\common.js");
}

var s = ["MainWindow = window;"];
for (var i = 0; i < args.length; i++) {
	var fn = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), args[i]);
	var ado = api.CreateObject("ads");
	if (ado) {
		ado.CharSet = "utf-8";
		ado.Open();
		ado.LoadFromFile(fn);
		var line = s.join("").split("\n").length;
		s.push(ado.ReadText());
		ado.Close();
		api.OutputDebugString([fso.GetFileName(args[i]), "Start line:", line, "\n"].join(""));
	}
}
if (window.chrome) {
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
	$JS = api.GetScriptDispatch(s.join(""), "JScript", o);
	CopyObj(window, $, ["g_", "Addons", "AboutTE", "AddEvent", "ClearEvent", "AddEnv", "AddEventId", "eventTE", "eventTA", "Threads", "BlurId", "ChangeView", "ShowError", "ShowStatusText"]);
	$.g_.$JS = $JS;
} else {
	$ = window;
	$JS = new Function(s.join(""));
	$JS();
}
MainWindow = $;

te.OnArrange = function (Ctrl, rc, cb) {
	UI.RunEvent1("Arrange", Ctrl, rc);
	if (Ctrl.Type == CTRL_TC) {
		var o = ui_.Panels[Ctrl.Id];
		if (!o) {
			o = document.createElement("table");
			o.id = "Panel_" + Ctrl.Id;
			o.className = "layout";
			o.style.position = "absolute";
			o.style.zIndex =  1;
			var s = '<tr><td id="InnerLeft_$" class="sidebar" style="width: 0; display: none; overflow: auto"></td><td style="width: 100%"><div id="InnerTop_$" style="display: none"></div>';
			s += '<table id="InnerTop2_$" class="layout"><tr><td id="Inner1Left_$" class="toolbar1"></td><td id="Inner1Center_$" class="toolbar2" style="white-space: nowrap"></td><td id="Inner1Right_$" class="toolbar3"></td></tr></table>';
			s += '<table id="InnerView_$" class="layout" style="width: 100%"><tr><td id="Inner2Left_$" style="width: 0"></td><td id="Inner2Center_$"></td><td id="Inner2Right_$" style="width: 0; overflow: auto"></td></tr></table>';
			s += '<div id="InnerBottom_$"></div></td><td id="InnerRight_$" class="sidebar" style="width: 0; display: none"></td></tr>';
			o.innerHTML = s.replace(/\$/g, Ctrl.Id);
			document.getElementById("Panel").appendChild(o);
			PanelCreated(Ctrl);
			ui_.Panels[Ctrl.Id] = o;
			ApplyLang(o);
			ChangeView(Ctrl.Selected);
		}
		o.style.left = rc.left + "px";
		o.style.top = rc.top + "px";
		if (Ctrl.Visible) {
			var s = [Ctrl.Left, Ctrl.Top, Ctrl.Width, Ctrl.Height].join(",");
			if (g_.TCPos[s] && g_.TCPos[s] != Ctrl.Id) {
				Ctrl.Close();
				return;
			} else {
				g_.TCPos[s] = Ctrl.Id;
			}
			o.style.display = (g_.IEVer >= 8 && o.tagName.toLowerCase() == "td") ? "table-cell" : "block";
		} else {
			o.style.display = "none";
		}
		o.style.width = Math.max(rc.right - rc.left, 0) + "px";
		o.style.height = Math.max(rc.bottom - rc.top, 0) + "px";
		rc.top += document.getElementById("InnerTop_" + Ctrl.Id).offsetHeight + document.getElementById("InnerTop2_" + Ctrl.Id).offsetHeight;
		var w1 = 0, w2 = 0, x = '';
		for (var i = 0; i <= 1; ++i) {
			w1 += api.LowPart(document.getElementById("Inner" + x + "Left_" + Ctrl.Id).style.width.replace(/\D/g, ""));
			w2 += api.LowPart(document.getElementById("Inner" + x + "Right_" + Ctrl.Id).style.width.replace(/\D/g, ""));
			x = '2';
		}
		rc.left += w1;
		rc.right -= w2;
		rc.bottom -= document.getElementById("InnerBottom_" + Ctrl.Id).offsetHeight;
		o = document.getElementById("Inner2Center_" + Ctrl.Id).style;
		o.width = Math.max(rc.right - rc.left, 0) + "px";
		o.height = Math.max(rc.bottom - rc.top, 0) + "px";
		api.OutputDebugString(["Arrange OK",rc.left, rc.top, rc.right, rc.bottom, "\n"].join(","));
		cb(Ctrl, rc);
	} else if (Ctrl.Type == CTRL_TE) {
		g_.TCPos = api.CreateObject("Object");
		var rcClient = api.Memory("RECT");
		api.GetClientRect(te.hwnd, rcClient);
		rcClient.left += te.offsetLeft;
		rcClient.top += te.offsetTop;
		rcClient.right -= te.offsetRight;
		rcClient.bottom -= te.offsetBottom;
		te.LockUpdate(1);
		try {
			var cTC = te.Ctrls(CTRL_TC, true);
			for (var i = 0; i < cTC.Count; ++i) {
				var TC = cTC[i];
				var rc = api.Memory("RECT");
				var res = /(\d+)%/.exec(TC.Left);
				if (res) {
					rc.left = res[1] * (rcClient.right - rcClient.left) / 100 + rcClient.left;
				} else {
					rc.left = TC.Left + (TC.Left >= 0 ? rcClient.left : rcClient.right);
				}
				if (TC.Left) {
					++rc.left;
				}
				res = /(\d+)%/.exec(TC.Top);
				if (res) {
					rc.top = res[1] * (rcClient.bottom - rcClient.top) / 100 + rcClient.top;
				} else {
					rc.top = TC.Top + (TC.Top >= 0 ? rcClient.top : rcClient.bottom);
				}
				if (TC.Top) {
					++rc.top;
				}
				res = /(\d+)%/.exec(TC.Width);
				if (res) {
					rc.right = res[1] * (rcClient.right - rcClient.left) / 100 + rc.left;
				} else {
					rc.right = TC.Width + (TC.Width >= 0 ? rc.left : rcClient.right);
				}
				if (rc.right >= rcClient.right) {
					rc.right = rcClient.right;
				} else {
					--rc.right;
				}
				res = /(\d+)%/.exec(TC.Height);
				if (res) {
					rc.bottom = res[1] * (rcClient.bottom - rcClient.top) / 100 + rc.top;
				} else {
					rc.bottom = TC.Height + (TC.Height >= 0 ? rc.top : rcClient.bottom);
				}
				if (rc.bottom >= rcClient.bottom) {
					rc.bottom = rcClient.bottom;
				} else {
					--rc.bottom;
				}
				te.OnArrange(TC, rc, cb);
			}
		} catch (e) { }
		te.UnlockUpdate(1);
	}
}

te.OnVisibleChanged = function (Ctrl) {
	api.OutputDebugString(["OnVisibleChanged", Ctrl.id,"\n"].join(","));
	UI.RunEvent1("VisibleChanged", Ctrl);
}

UI.OnLoad();

ArrangeAddons = function () {
	g_.Locations = {};
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
			var arError = [];
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
			if (arError.length) {
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
	UI.RunEvent1("Load");
	ClearEvent("Load");
	var cHwnd = [te.Ctrl(CTRL_WB).hwnd, te.hwnd];
	var cl = $.GetWinColor(window.getComputedStyle ? getComputedStyle(document.body).getPropertyValue('background-color') : document.body.currentStyle.backgroundColor);
	for (var i = cHwnd.length; i--;) {
		var hOld = api.SetClassLongPtr(cHwnd[i], GCLP_HBRBACKGROUND, api.CreateSolidBrush(cl));
		if (hOld) {
			api.DeleteObject(hOld);
		}
	}
	var hwnd, p = api.Memory("WCHAR", 11);
	p.Write(0, VT_LPWSTR, "ShellState");
	var cFV = te.Ctrls(CTRL_FV);
	for (var i in cFV) {
		if (hwnd = cFV[i].hwndView) {
			api.SendMessage(hwnd, WM_SETTINGCHANGE, 0, p);
		}
	}
}

InitCode = function () {
	var types =
	{
		Key: ["All", "List", "Tree", "Browser", "Edit", "Menus"],
		Mouse: ["All", "List", "List_Background", "Tree", "Tabs", "Tabs_Background", "Browser"]
	};
	var i;
	for (i = 0; i < 3; ++i) {
		g_.KeyState[i][0] = api.GetKeyNameText(g_.KeyState[i][0]);
	}
	i = g_.KeyState.length;
	while (i-- > 4 && g_.KeyState[i][0] == g_.KeyState[i - 4][0]) {
		g_.KeyState.pop();
	}
	for (var j = 256; j >= 0; j -= 256) {
		for (var i = 128; i > 0; i--) {
			var v = api.GetKeyNameText((i + j) * 0x10000);
			if (v && v.charCodeAt(0) > 32) {
				g_.KeyCode[v.toUpperCase()] = i + j;
			}
		}
	}
	for (var i in types) {
		eventTE[i] = api.CreateObject("Object");
		for (var j in types[i]) {
			eventTE[i][types[i][j]] = api.CreateObject("Object");
		}
	}
	DefaultFont = api.Memory("LOGFONT");
	$.DefaultFont = DefaultFont;
	api.SystemParametersInfo(SPI_GETICONTITLELOGFONT, DefaultFont.Size, DefaultFont, 0);
	HOME_PATH = te.Data.Conf_NewTab || HOME_PATH;
	$.HOME_PATH = HOME_PATH;
}

InitMouse = function () {
	te.Data.Conf_Gestures = isFinite(te.Data.Conf_Gestures) ? Number(te.Data.Conf_Gestures) : 2;
	if ("string" === typeof te.Data.Conf_TrailColor) {
		te.Data.Conf_TrailColor = GetWinColor(te.Data.Conf_TrailColor);
	}
	if (!isFinite(te.Data.Conf_TrailColor) || te.Data.Conf_TrailColor === "") {
		te.Data.Conf_TrailColor = 0xff00;
	}
	te.Data.Conf_TrailSize = isFinite(te.Data.Conf_TrailSize) ? Number(te.Data.Conf_TrailSize) : 2;
	te.Data.Conf_GestureTimeout = isFinite(te.Data.Conf_GestureTimeout) ? Number(te.Data.Conf_GestureTimeout) : 3000;
	te.Data.Conf_Layout = isFinite(te.Data.Conf_Layout) ? Number(te.Data.Conf_Layout) : 0x80;
	te.Data.Conf_NetworkTimeout = isFinite(te.Data.Conf_NetworkTimeout) ? Number(te.Data.Conf_NetworkTimeout) : 2000;
	te.Data.Conf_WheelSelect = isFinite(te.Data.Conf_WheelSelect) ? Number(te.Data.Conf_WheelSelect) : 1;
	te.SizeFormat = (te.Data.Conf_SizeFormat || "").replace(/^0x/i, "");
	te.HiddenFilter = ExtractFilter(te.Data.Conf_HiddenFilter);
	te.DragIcon = !api.LowPart(te.Data.Conf_NoDragIcon);
	var ar = ['AutoArrange', 'ColumnEmphasis', 'DateTimeFormat', 'Layout', 'LibraryFilter', 'NetworkTimeout', 'ShowInternet', 'ViewOrder'];
	for (var i = ar.length; i--;) {
		te[ar[i]] = te.Data['Conf_' + ar[i]];
	}
	OpenMode = te.Data.Conf_OpenMode ? SBSP_NEWBROWSER : SBSP_SAMEBROWSER;
}

InitMenus = function () {
	te.Data.xmlMenus = OpenXml("menus.xml", false, true);
	var root = te.Data.xmlMenus.documentElement;
	if (root) {
		menus = root.childNodes;
		for (var i = api.ObjGetI(menus, "length"); i--;) {
			items = menus[i].getElementsByTagName("Item");
			for (var j = api.ObjGetI(items, "length"); j--;) {
				a = items[j].getAttribute("Name").split(/\\t/);
				if (a.length > 1) {
					$.SetKeyExec("List", a[1], items[j].text, items[j].getAttribute("Type"), true);
				}
			}
		}
	}
}

InitWindow = function () {
	if (g_.xmlWindow && "string" !== typeof g_.xmlWindow) {
		$.LoadXml(g_.xmlWindow);
	}
	if (te.Ctrls(CTRL_TC).Count == 0) {
		var TC = te.CreateCtrl(CTRL_TC, 0, 0, "100%", "100%", te.Data.Tab_Style, te.Data.Tab_Align, te.Data.Tab_TabWidth, te.Data.Tab_TabHeight);
		TC.Selected.Navigate2(HOME_PATH, SBSP_NEWBROWSER, te.Data.View_Type, te.Data.View_ViewMode, te.Data.View_fFlags, te.Data.View_Options, te.Data.View_ViewFlags, te.Data.View_IconSize, te.Data.Tree_Align, te.Data.Tree_Width, te.Data.Tree_Style, te.Data.Tree_EnumFlags, te.Data.Tree_RootStyle, te.Data.Tree_Root);
	}
	g_.xmlWindow = void 0;
	setTimeout(function () {
		var a, i, j, menus, items;
		Resize();
		var cTC = te.Ctrls(CTRL_TC);
		for (var i in cTC) {
			if (cTC[i].SelectedIndex >= 0) {
				ChangeView(cTC[i].Selected);
			}
		}
		api.ShowWindow(te.hwnd, te.CmdShow);
		if (te.CmdShow == SW_SHOWNOACTIVATE) {
			api.SetWindowPos(te.hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
		}
		te.CmdShow = SW_SHOWNORMAL;
		setTimeout(function () {
			$.RunCommandLine(api.GetCommandLine());
			api.PostMessage(te.hwnd, WM_SIZE, 0, 0);
		}, 500);
	}, 99);
}

// Events

AddEvent("VisibleChanged", function (Ctrl) {
	if (Ctrl.Type == CTRL_TC) {
		api.OutputDebugString([Ctrl.Id, Ctrl.Visible, "\n"].join(","))
		var o = ui_.Panels[Ctrl.Id];
		if (o) {
			api.OutputDebugString([Ctrl.Visible, "\n"].join(","))
			if (Ctrl.Visible) {
				o.style.display = (g_.IEVer >= 8 && o.tagName.toLowerCase() == "td") ? "table-cell" : "block";
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
		$.ShowOptions("Tab=Add-ons");
	}
});

AddEventEx(window, "resize", Resize);

AddEventEx(window, "beforeunload", $.Finalize);

AddEventEx(window, "blur", ResetScroll);

AddEventEx(document, "MSFullscreenChange", function () {
	if (document.msFullscreenElement) {
		var cTC = te.Ctrls(CTRL_TC, true);
		for (var i in cTC) {
			var TC = cTC[i];
			g_.stack_TC.push(TC);
			TC.Visible = false;
		}
	} else {
		while (api.ObjGetI(g_.stack_TC, "length")) {
			g_.stack_TC.pop().Visible = true;
		}
	}
});

//
InitCode();
InitMouse();
InitMenus();
LoadLang();
ArrangeAddons();
setTimeout(InitWindow);
