// Tablacus Explorer

Resize = async function () {
	ResetScroll();
	let o = document.getElementById("toolbar");
	const offsetTop = o ? o.offsetHeight : 0;
	let h = 0;
	o = document.getElementById("bottombar");
	const offsetBottom = o.offsetHeight;
	o = document.getElementById("client");
	const ode = document.documentElement || document.body;
	if (o) {
		h = Math.max(ode.offsetHeight - offsetBottom - offsetTop, 0);
		o.style.height = h + "px";
	}
	await Promise.all([ResizeSideBar("Left", h), ResizeSideBar("Right", h)]);
	o = document.getElementById("Background");
	pt = GetPos(o);
	te.offsetLeft = pt.x;
	te.offsetRight = ode.offsetWidth - o.offsetWidth - pt.x;
	te.offsetTop = pt.y;
	pt = GetPos(document.getElementById("bottombar"));
	te.offsetBottom = ode.offsetHeight - pt.y;
	if (ui_.Show) {
		await RunEventUI1("Resize");
	}
	api.PostMessage(ui_.hwnd, WM_SIZE, 0, 0);
}

ResizeSideBar = async function (z, h) {
	let o = g_.Locations;
	const r = await Promise.all([te.Data["Conf_" + z + "BarWidth"], o[z + "Bar1"], o[z + "Bar2"], o[z + "Bar3"]]);
	const w = (r[1] || r[2] || r[3]) ? r[0] : 0;
	o = document.getElementById(z.toLowerCase() + "bar");
	if (w > 0) {
		o.style.display = "";
		if (w != o.offsetWidth) {
			o.style.width = w + "px";
			for (let i = 1; i <= 3; ++i) {
				document.getElementById(z + "Bar" + i).style.width = w + "px";
			}
			document.getElementById(z.toLowerCase() + "barT").style.width = w + "px";
		}
	} else {
		o.style.display = "none";
	}
	document.getElementById(z.toLowerCase() + "splitter").style.display = w ? "" : "none";

	o = document.getElementById(z.toLowerCase() + "barT");
	const th = Math.round(Math.max(h, 0));
	o.style.height = th + "px";

	const h2 = Math.max(o.clientHeight - document.getElementById(z + "Bar1").offsetHeight - document.getElementById(z + "Bar3").offsetHeight, 0);
	document.getElementById(z + "Bar2").style.height = h2 + "px";
}

ResetScroll = function () {
	if (document.documentElement && document.documentElement.scrollLeft) {
		document.documentElement.scrollLeft = 0;
	}
}

PanelCreated = async function (Ctrl, Id) {
	await RunEventUI1("PanelCreated", Ctrl, Id);
	await Resize();
	setTimeout(ChangeView, 9, await Ctrl.Selected);
}

Activate = async function (el, id) {
	const TC = await te.Ctrl(CTRL_TC);
	if (TC && await TC.Id != id) {
		const FV = await GetInnerFV(id);
		if (FV) {
			FV.Focus();
			setTimeout(FocusElement, 9, el);
		}
	}
}

SetAddon = async function (strName, Location, Tag, strVAlign) {
	if (strName) {
		const s = await GetAddonLocation(strName);
		if (s) {
			Location = s;
		}
	}
	if (Tag) {
		if (Tag.join) {
			Tag = Tag.join("");
		}
		const re = /(<[^<>]*placeholder=")([^"<>]+)/;
		const res = re.exec(Tag);
		if (res) {
			Tag = Tag.replace(re, res[1] + await GetTextR(res[2]));
		}
		const o = document.getElementById(Location);
		if (o) {
			if ("string" === typeof Tag) {
				o.insertAdjacentHTML("beforeend", Tag);
			} else {
				o.appendChild(Tag);
			}
			o.style.display = (ui_.IEVer >= 8 && SameText(o.tagName, "td")) ? "table-cell" : "block";
			if (strVAlign && !o.style.verticalAlign) {
				o.style.verticalAlign = strVAlign;
			}
			const res = /^(LeftBar|RightBar)/.exec(Location);
			if (res) {
				const s = "Conf_" + res[1] + "Width";
				if (!await te.Data[s]) {
					te.Data[s] = 178;
				}
			}
		} else if (Location == "Inner") {
			AddEventUI("PanelCreated", function (Ctrl, Id) {
				return SetAddon(null, "Inner1Left_" + Id, Tag.replace(/\$/g, Id));
			});
		} else if (Location == "InnerRight") {
			AddEventUI("PanelCreated", function (Ctrl, Id) {
				return SetAddon(null, "Inner1Right_" + Id, Tag.replace(/\$/g, Id));
			});
		} else if (Location == "InnerBottom") {
			AddEventUI("PanelCreated", function (Ctrl, Id) {
				return SetAddon(null, "InnerBottom2Left_" + Id, Tag.replace(/\$/g, Id));
			});
		} else if (Location == "InnerBottomRight") {
			AddEventUI("PanelCreated", function (Ctrl, Id) {
				return SetAddon(null, "InnerBottom2Right_" + Id, Tag.replace(/\$/g, Id));
			});
		}
		if (strName) {
			if ("string" === typeof Tag) {
				const res = /<img.*?org=["'](.*?)["']/i.exec(Tag) || /<span.*?src=["'](.*?)["']/i.exec(Tag);
				if (res) {
					strName += "\t" + res[1];
				}
			}
			await SetAddonLocation(Location, strName);
		}
	}
	return Location;
}

RunSplitter = async function (ev, n) {
	if ((ev.buttons != null ? ev.buttons : ev.button) == 1) {
		api.ObjPutI(await g_.mouse, "Capture", n);
		api.SetCapture(ui_.hwnd);
	}
}

DisableImage = function (img, bDisable) {
	if (img) {
		if (window.chrome) {
			img.style.filter = bDisable ? "grayscale(1) opacity(.5)" : "";
		} else if (ui_.IEVer >= 10) {
			let s = decodeURIComponent(img.src);
			const res = /^data:image\/svg.*?href="([^"]*)/i.exec(s);
			if (bDisable) {
				if (!res && /^IMG$/i.test(img.tagName)) {
					if (/^file:/i.test(s)) {
						let image;
						if (image = api.CreateObject("WICBitmap").FromFile(api.PathCreateFromUrl(s))) {
							s = image.DataURI("image/png");
						}
					}
					img.src = "data:image/svg+xml," + encodeURIComponent(['<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ', img.offsetWidth, ' ', img.offsetHeight, '"><filter id="G"><feColorMatrix type="saturate" values="0.1" /></filter><image width="', img.width, '" height="', img.height, '" xlink:href="', s, '" filter="url(#G)"></image></svg>'].join(""));
				}
				img.style.opacity = .5;
			} else {
				if (res) {
					img.src = res[1];
				}
				img.style.opacity = 1;
			}
		} else {
			img.style.filter = bDisable ? "gray(), alpha(style=0,opacity=50);" : "";
		}
	}
}

StartGestureTimer = async function () {
	const i = await te.Data.Conf_GestureTimeout;
	if (i) {
		clearTimeout(await g_.mouse.tidGesture);
		api.ObjPutI(await g_.mouse, "tidGesture", setTimeout(function () {
			g_.mouse.EndGesture(true);
		}, i));
	}
}

FocusFV1 = function (Id) {
	let el;
	if (document.activeElement) {
		if (/input|textarea/i.test(document.activeElement.tagName)) {
			return;
		}
		const rc = document.activeElement.getBoundingClientRect();
		el = document.elementFromPoint((rc.left + rc.right) / 2, (rc.top + rc.bottom) / 2);
	}
	if (!el || !/input|textarea/i.test(el.tagName)) {
		FocusFV2("number" === typeof Id ? Id : null);
	}
}

FocusFV = function (Id) {
	const tm = new Date().getTime() - ui_.tmDown;
	setTimeout(FocusFV1, tm < ui_.DoubleClickTime ? tm : 9, Id);
}

ToggleFullscreen = async function () {
	if (await ExitFullscreen()) {
		return;
	}
	FullscreenChanged(true, true);
}

ExitFullscreen = async function () {
	if (document.msFullscreenElement != null) {
		document.msExitFullscreen();
		return true;
	}
	if (document.fullscreenElement != null) {
		document.webkitCancelFullScreen();
		return true;
	}
	if (await g_.Fullscreen) {
		FullscreenChanged(false, true);
		return true;
	}
}

MoveSplitter = function (x, n) {
	const w = document.documentElement.offsetWidth || document.body.offsetWidth;
	if (x >= w) {
		x = w - 1;
	}
	if (n == 1) {
		te.Data["Conf_LeftBarWidth"] = x;
	} else if (n == 2) {
		te.Data["Conf_RightBarWidth"] = w - x;
	}
	Resize();
}

ShowStatusTextEx = function (Ctrl, Text, iPart, tm) {
	if (ui_.Status && ui_.Status[5]) {
		clearTimeout(ui_.Status[5]);
		delete ui_.Status;
	}
	ui_.Status = [Ctrl, Text, iPart, tm, new Date().getTime(), setTimeout(function () {
		if (ui_.Status) {
			if (new Date().getTime() - ui_.Status[4] > ui_.Status[3] / 2) {
				ShowStatusText(ui_.Status[0], ui_.Status[1], ui_.Status[2]);
				delete ui_.Status;
			}
		}
	}, tm)];
}

importJScript = $.importScript;

OnArrange = async function (Ctrl, rc, Type, Id, FV) {
	if (Type == CTRL_TE) {
		ui_.TCPos = {};
	}
	await RunEventUI1("Arrange", Ctrl, rc, Type, Id, FV);
	if (Type == CTRL_TC) {
		const p = [rc.left, rc.top, rc.right, rc.bottom, Ctrl.Visible, Ctrl.Left, Ctrl.Top, Ctrl.Width, Ctrl.Height];
		if (!document.getElementById("Panel_" + Id)) {
			const s = ['<table id="Panel_', Id, '" class="layout fixed" style="position: absolute; z-index: 1; color: inherit; visibility: hidden" onclick="FocusFV1(', Id, ')">'];
			s.push('<tr><td id="InnerLeft_', Id, '" class="sidebar" style="width: 0; display: none; overflow: auto"></td><td class="full"><div id="InnerTop_', Id, '" style="display: none"></div>');
			s.push('<table id="InnerTop2_', Id, '" class="layout">');
			s.push('<tr><td id="Inner1Left_', Id, '" class="toolbar1"></td><td id="Inner1Center_', Id, '" class="toolbar2 nowrap"></td><td id="Inner1Right_', Id, '" class="toolbar3"></td></tr></table>');
			s.push('<table id="InnerView_', Id, '" class="layout full"><tr><td id="Inner2Left_', Id, '" style="width: 0"></td><td id="Inner2Center_', Id, '" class="full"></td><td id="Inner2Right_', Id, '" style="width: 0; overflow: auto"></td></tr></table>');
			s.push('<table id="InnerBottomT_', Id, '" class="layout full"><tr><td id="InnerBottom_', Id, '"></td></tr><tr><td id="InnerBottom2Left_', Id, '" class="toolbar1 full"></td><td id="InnerBottom2Right_', Id, '" class="toolbar3"></td></tr></table></td><td id="InnerRight_', Id, '" class="sidebar" style="width: 0; display: none"></td></tr></table>');
			document.getElementById("Panel").insertAdjacentHTML("beforeend", s.join(""));
			p.push(PanelCreated(Ctrl, Id));
		}
		Promise.all(p).then(function (r) {
			const o = document.getElementById("Panel_" + Id);
			o.style.left = r[0] + "px";
			o.style.top = r[1] + "px";
			if (r[4]) {
				const s = r.slice(5, 8).join(",");
				if (ui_.TCPos[s] && ui_.TCPos[s] != Id) {
					Ctrl.Close();
					return;
				}
				ui_.TCPos[s] = Id;
				o.style.display = "";
			} else {
				o.style.display = "none";
				return;
			}
			let w = Math.max(r[2] - r[0], 0) + "px", h = Math.max(r[3] - r[1], 0) + "px";
			if (o.style.width != w) {
				o.style.width = w;
				ui_.bResize = true;
			}
			if (o.style.height != h) {
				o.style.height = h;
				ui_.bResize = true;
			}
			let el = document.getElementById("InnerLeft_" + Id);
			if (!/none/i.test(el.style.display)) {
				r[0] += el.offsetWidth;
			}
			el = document.getElementById("Inner2Left_" + Id);
			if (!/none/i.test(el.style.display)) {
				r[0] += el.offsetWidth;
			}
			r[1] += document.getElementById("InnerTop_" + Id).offsetHeight + document.getElementById("InnerTop2_" + Id).offsetHeight;
			el = document.getElementById("InnerRight_" + Id);
			if (!/none/i.test(el.style.display)) {
				r[2] -= el.offsetWidth;
			}
			el = document.getElementById("Inner2Right_" + Id);
			if (!/none/i.test(el.style.display)) {
				r[2] -= el.offsetWidth;
			}
			r[3] -= document.getElementById("InnerBottomT_" + Id).offsetHeight;
			api.SetRect(rc, r[0], r[1], r[2], r[3]);
			document.getElementById("Inner2Center_" + Id).style.height = Math.max(r[3] - r[1], 0) + "px";
			Promise.all([te.ArrangeCB(Ctrl, rc)]).then(function () {
				o.style.visibility = "visible";
				if (ui_.Show == 1) {
					ui_.Show = 2;
					SetWindowAlpha(ui_.hwnd, 255);
					RunEvent1("VisibleChanged", te, true, CTRL_TE);
				}
				if (ui_.bResize) {
					RunEventUI1("Resize");
					ui_.bResize = false;
				}
			});
		});
	}
}

ArrangeAddons = async function () {
	let r = await Promise.all([InitAddonsXML(), api.CreateObject("Object"), api.GetKeyState(VK_SHIFT), api.GetKeyState(VK_CONTROL), GetLangId(), api.CreateObject("Array")]);
	const xml = window.chrome ? new DOMParser().parseFromString(r[0], "application/xml") : te.Data.Addons;
	ui_.Addons = xml;
	g_.Locations = r[1];
	if (r[2] < 0 && r[3] < 0) {
		IsSavePath = function (path) {
			return false;
		}
		return;
	}
	const LangId = r[4];
	let arError = r[5];
	delete r;
	const AddonId = {};
	const root = xml.documentElement;
	if (root) {
		const items = root.childNodes;
		if (items) {
			document.F.style.visibility = "hidden";
			for (let i = 0; i < items.length; ++i) {
				const item = items[i];
				const Id = item.nodeName;
				g_.Error_source = Id;
				if (!AddonId[Id]) {
					const Enabled = GetNum(item.getAttribute("Enabled"));
					if (Enabled) {
						if (Enabled & 6) {
							LoadLang2(BuildPath(ui_.Installed, "addons", Id, "lang", LangId + ".xml"));
						}
						if (Enabled & 8) {
							await LoadAddon("vbs", Id, arError, null, window.chrome && GetNum(item.getAttribute("Level")) < 2);
						}
						if (Enabled & 1) {
							await LoadAddon("js", Id, arError, null, window.chrome && GetNum(item.getAttribute("Level")) < 2);
						}
					}
					AddonId[Id] = true;
				}
				g_.Error_source = "";
			}
			if (window.chrome) {
				arError = await api.CreateObject("SafeArray", arError);
			}
			if (arError.length) {
				setTimeout(async function (arError) {
					if (await MessageBox(arError.join("\n\n"), TITLE, MB_OKCANCEL) != IDCANCEL) {
						te.Data.bErrorAddons = true;
						ShowOptions("Tab=Add-ons");
					}
				}, 500, arError);
			}
		}
	}
}

// Events

AddEvent("VisibleChanged", async function (Ctrl, Visible, Type, Id) {
	if ((Type != null ? Type : await Ctrl.Type) == CTRL_TC) {
		const o = document.getElementById("Panel_" + (Id != null ? Id : await Ctrl.Id));
		if (o) {
			if (Visible != null ? Visible : await Ctrl.Visible) {
				o.style.display = "";
				ChangeView(await Ctrl.Selected);
			} else {
				o.style.display = "none";
			}
		}
	} else if (Type == CTRL_TV) {
		ui_.bResize = true;
	}
});

AddEvent(WM_KILLFOCUS + "!", function (Ctrl, Type, hwnd, msg, wParam, lParam) {
	if (Type == CTRL_WB) {
		const o = document.activeElement;
		if (o) {
			const s = o.style.visibility;
			o.style.visibility = "hidden";
			o.style.visibility = s;
			FireEvent(o, "blur");
		}
	}
});

AddEvent(WM_SIZE + "!", function (Ctrl, Type, hwnd, msg, wParam, lParam) {
	if (Type == CTRL_TV) {
		Resize();
	}
});

// Browser Events

window.addEventListener("load", async function () {
	if (await api.GetKeyState(VK_SHIFT) < 0 && await api.GetKeyState(VK_CONTROL) < 0) {
		ShowOptions("Tab=Add-ons");
	}
});

window.addEventListener("resize", Resize);

window.addEventListener("unload", FinalizeUI);

window.addEventListener("blur", ResetScroll);

window.addEventListener("mouseup", FocusFV);

window.addEventListener("mousedown", function (ev) {
	ui_.tmDown = new Date().getTime();
	if (window.chrome) {
		g_.mouse.ptDown.x = ev.screenX * ui_.Zoom;
		g_.mouse.ptDown.y = ev.screenY * ui_.Zoom;
		g_.mouse.CancelContextMenu = false;
	}
});

document.addEventListener("MSFullscreenChange", function () {
	FullscreenChanged(document.msFullscreenElement != null, document.msFullscreenElement == document.body);
});

document.addEventListener("webkitfullscreenchange", function () {
	FullscreenChanged(document.fullscreenElement != null, document.fullscreenElement == document.body);
});

Init = async function () {
	te.Data.MainWindow = $;
	te.Data.NoCssFont = ui_.NoCssFont;
	await InitCode();
	InitCloud();
	const r = await Promise.all([te.Data.DataFolder, $.DefaultFont, $.HOME_PATH, $.OpenMode, $.DefaultFont.lfFaceName, $.DefaultFont.lfHeight, $.DefaultFont.lfWeight, te.Data.Conf_IconSize, te.Data.Conf_InnerIconSize, InitMenus(), LoadLang()]);
	ui_.DataFolder = r[0];
	DefaultFont = r[1];
	HOME_PATH = r[2];
	OpenMode = r[3];
	document.body.style.fontFamily = r[4];
	document.body.style.fontSize = Math.abs(r[5]) + "px";
	document.body.style.fontWeight = r[6];
	$.IconSize = IconSize = r[7] || screen.deviceYDPI / 4;
	ui_.InnerIconSize = r[8] ? r[8] * 96 / screen.deviceYDPI : 16;
	await ArrangeAddons();
	RunEventUI("BrowserCreatedEx");
	await RunEventUI1("Layout");
	document.F.style.visibility = "";
	te.OnArrange = OnArrange;
	await InitBG(window.getComputedStyle ? getComputedStyle(document.body).getPropertyValue('background-color') : document.body.currentStyle.backgroundColor, true);
	await InitWindow();
	await RunEventUI1("Load");
	ui_.Show = 1;
	Resize();
	AddEvent("BrowserCreatedEx", "setTimeout(async function () { SetWindowAlpha(await GetTopWindow(), 255); }, 99);");
	ClearEvent("Layout");
	ClearEvent("Load");
	g_.ShowError = true;
}
