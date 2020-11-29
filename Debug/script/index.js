// Tablacus Explorer

async function Resize2() {
	if (ui_.tidResize) {
		clearTimeout(ui_.tidResize);
		ui_.tidResize = void 0;
	}
	ResetScroll();
	let o = document.getElementById("toolbar");
	const offsetTop = o ? o.offsetHeight : 0;

	let h = 0;
	o = document.getElementById("bottombar");
	const offsetBottom = o.offsetHeight;
	o = document.getElementById("client");
	const ode = document.documentElement || document.body;
	if (o) {
		h = ode.offsetHeight - offsetBottom - offsetTop;
		if (h < 0) {
			h = 0;
		}
		o.style.height = h + "px";
	}
	await ResizeSideBar("Left", h);
	await ResizeSideBar("Right", h);
	o = document.getElementById("Background");
	pt = GetPos(o);
	te.offsetLeft = pt.x;
	te.offsetRight = ode.offsetWidth - o.offsetWidth - pt.x;
	te.offsetTop = pt.y;
	pt = GetPos(document.getElementById("bottombar"));
	te.offsetBottom = ode.offsetHeight - pt.y;
	RunEvent1("Resize");
	api.PostMessage(ui_.hwnd, WM_SIZE, 0, 0);
}

async function ResizeSideBar(z, h) {
	let o = await g_.Locations;
	const w = (await o[z + "Bar1"] || await o[z + "Bar2"] || await o[z + "Bar3"]) ? await te.Data["Conf_" + z + "BarWidth"] : 0;
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

function ResetScroll () {
	if (document.documentElement && document.documentElement.scrollLeft) {
		document.documentElement.scrollLeft = 0;
	}
}

function PanelCreated(Ctrl, Id) {
	if (/none/i.test(document.F.style.display)) {
		setTimeout(PanelCreated, 500, Ctrl, Id);
		return;
	}
	RunEvent1("PanelCreated", Ctrl, Id);
	ApplyLang(document.getElementById("Panel_" + Id));
	Resize();
}

Activate = async function (o, id) {
	const TC = await te.Ctrl(CTRL_TC);
	if (TC && await TC.Id != id) {
		const FV = await GetInnerFV(id);
		if (FV) {
			FV.Focus();
			setTimeout(function () {
				o.focus();
			}, 99);
		}
	}
}

GetAddonLocation = async function (strName) {
	const items = await te.Data.Addons.getElementsByTagName(strName);
	return (await GetLength(items) ? await items[0].getAttribute("Location") : null);
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
		} else if (Location == "Inner") {
			AddEvent("PanelCreated", function (Ctrl, Id) {
				SetAddon(null, "Inner1Left_" + Id, Tag.replace(/\$/g, Id));
			});
		}
		if (strName) {
			if (!await g_.Locations[Location]) {
				g_.Locations[Location] = await api.CreateObject("Array");
			}
			const res = /<img.*?src=["'](.*?)["']/i.exec(String(Tag));
			if (res) {
				strName += "\t" + res[1];
			}
			await g_.Locations[Location].push(strName);
		}
	}
	return Location;
}

RunSplitter = async function (ev, n) {
	if (ev.button == 0) {
		api.ObjPutI(await g_.mouse, "Capture", n);
		api.SetCapture(ui_.hwnd);
	}
}

Resize = function () {
	if (ui_.tidResize) {
		clearTimeout(ui_.tidResize);
	}
	ui_.tidResize = setTimeout(Resize2, 500);
}

DisableImage = function (img, bDisable) {
	if (img) {
		if (window.chrome) {
			if (bDisable) {
				const ar = [];
				for (let i = 0; i < img.style.length; ++i) {
					if (img.style[i] != "filter") {
						ar.push(img.style[i] + ":" + img.style[img.style[i]]);
					}
				}
				ar.push("filter:grayscale(1);");
				img.style = ar.join(";");
				img.style.opacity = .5;
			} else {
				img.style.filter  = "";
				img.style.opacity = 1;
			}
		} else if (ui_.IEVer >= 10) {
			let s = decodeURIComponent(img.src);
			const res = /^data:image\/svg.*?href="([^"]*)/i.exec(s);
			if (bDisable) {
				if (!res) {
					if (/^file:/i.test(s)) {
						let image;
						if (image = api.CreateObject("WICBitmap").FromFile(api.PathCreateFromUrl(s))) {
							s = image.DataURI("image/png");
						}
					}
					img.src = "data:image/svg+xml," + encodeURIComponent(['<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ', img.offsetWidth, ' ', img.offsetHeight, '"><filter id="G"><feColorMatrix type="saturate" values="0.1" /></filter><image width="', img.width, '" height="', img.height, '" opacity=".5" xlink:href="', s, '" filter="url(#G)"></image></svg>'].join(""));
				}
			} else if (res) {
				img.src = res[1];
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

FocusFV = function () {
	setTimeout(async function () {
		let el;
		if (document.activeElement) {
			const rc = document.activeElement.getBoundingClientRect();
			el = document.elementFromPoint(rc.left + 2, rc.top + 2);
		}
		if (!el || !/input|textarea/i.test(el.tagName)) {
			const FV = await GetFolderView();
			if (FV) {
				FV.Focus();
			}
		}
	}, ui_.DoubleClickTime);
}

SelectNext = function (FV) {
	setTimeout(async function () {
		if (!await api.SendMessage(await FV.hwndList, LVM_GETEDITCONTROL, 0, 0) || WINVER < 0x600) {
			FV.SelectItem(FV.Item(await FV.GetFocusedItem + (await api.GetKeyState(VK_SHIFT) < 0 ? -1 : 1)) || await api.GetKeyState(VK_SHIFT) < 0 ? FV.ItemCount(SVGIO_ALLVIEW) - 1 : 0, SVSI_EDIT | SVSI_FOCUSED | SVSI_SELECT | SVSI_DESELECTOTHERS);
		}
	}, 99);
}

CancelWindowRegistered = function () {
	clearTimeout(ui_.tidWindowRegistered);
	ui_.bWindowRegistered = false;
	ui_.tidWindowRegistered = setTimeout(function () {
		ui_.bWindowRegistered = true;
	}, 9999);
}

ExitFullscreen = function () {
	if (document.msExitFullscreen) {
		document.msExitFullscreen();
	} else if (document.exitFullscreen) {
		document.exitFullscreen();
	}
}

MoveSplitter = async function (x, n) {
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

ShowStatusTextEx = async function (Ctrl, Text, iPart, tm) {
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

te.OnArrange = async function (Ctrl, rc) {
	const Type = await Ctrl.Type;
	if (Type == CTRL_TE) {
		ui_.TCPos = {};
	}
	await RunEvent1("Arrange", Ctrl, rc);
	if (Type == CTRL_TC) {
		const Id = await Ctrl.Id;
		let o = document.getElementById("Panel_" + Id);
		if (!o) {
			const s = ['<table id="Panel_', Id, '" class="layout" style="position: absolute; z-index: 1;">'];
			s.push('<tr><td id="InnerLeft_', Id, '" class="sidebar" style="width: 0; display: none; overflow: auto"></td><td style="width: 100%"><div id="InnerTop_', Id, '" style="display: none"></div>');
			s.push('<table id="InnerTop2_', Id, '" class="layout">');
			s.push('<tr><td id="Inner1Left_', Id, '" class="toolbar1"></td><td id="Inner1Center_', Id, '" class="toolbar2" style="white-space: nowrap"></td><td id="Inner1Right_', Id, '" class="toolbar3"></td></tr></table>');
			s.push('<table id="InnerView_', Id, '" class="layout" style="width: 100%"><tr><td id="Inner2Left_', Id, '" style="width: 0"></td><td id="Inner2Center_', Id, '" style="width: 100%"></td><td id="Inner2Right_', Id, '" style="width: 0; overflow: auto"></td></tr></table>');
			s.push('<div id="InnerBottom_', Id, '"></div></td><td id="InnerRight_', Id, '" class="sidebar" style="width: 0; display: none"></td></tr></table>');
			document.getElementById("Panel").insertAdjacentHTML("beforeend", s.join(""));
			o = document.getElementById("Panel_" + Id);
			PanelCreated(Ctrl, Id);
		}
		Promise.all([rc.left, rc.top, rc.right, rc.bottom, Ctrl.Visible, Ctrl.Left, Ctrl.Top, Ctrl.Width, Ctrl.Height]).then(function (r) {
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
			}
			o.style.width = Math.max(r[2] - r[0], 0) + "px";
			o.style.height = Math.max(r[3] - r[1], 0) + "px";
			rc.left = r[0] + document.getElementById("InnerLeft_" + Id).offsetWidth + document.getElementById("Inner2Left_" + Id).offsetWidth;
			rc.top = r[1] + document.getElementById("InnerTop_" + Id).offsetHeight + document.getElementById("InnerTop2_" + Id).offsetHeight;
			rc.right = r[2] - document.getElementById("InnerRight_" + Id).offsetWidth - document.getElementById("Inner2Right_" + Id).offsetWidth;
			rc.bottom = r[3] - document.getElementById("InnerBottom_" + Id).offsetHeight;
			te.ArrangeCB(Ctrl, rc);
		});
	}
}

g_.event.windowregistered = function (Ctrl) {
	if (ui_.bWindowRegistered) {
		RunEvent1("WindowRegistered", Ctrl);
	}
	ui_.bWindowRegistered = true;
}

ArrangeAddons = async function () {
	g_.Locations = await api.CreateObject("Object");
	$.IconSize = IconSize = await te.Data.Conf_IconSize || screen.deviceYDPI / 4;
	const xml = await OpenXml("addons.xml", false, true);
	te.Data.Addons = xml;
	if (await api.GetKeyState(VK_SHIFT) < 0 && await api.GetKeyState(VK_CONTROL) < 0) {
		IsSavePath = function (path) {
			return false;
		}
		return;
	}
	const AddonId = [];
	const root = await xml.documentElement;
	if (root) {
		const items = await root.childNodes;
		if (items) {
			let arError = await api.CreateObject("Array");
			const LangId = await GetLangId();
			const nLen = await GetLength(items);
			for (let i = 0; i < nLen; ++i) {
				const item = await items[i];
				const Id = await item.nodeName;
				g_.Error_source = Id;
				if (!AddonId[Id]) {
					const Enabled = GetNum(await item.getAttribute("Enabled"));
					if (Enabled) {
						if (Enabled & 6) {
							LoadLang2(BuildPath(ui_.Installed, "addons", Id, "lang", LangId + ".xml"));
						}
						if (Enabled & 8) {
							LoadAddon("vbs", Id, arError, null, window.chrome && GetNum(await item.getAttribute("Level")) < 2);
						}
						if (Enabled & 1) {
							LoadAddon("js", Id, arError, null, window.chrome && GetNum(await item.getAttribute("Level")) < 2);
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
	RunEventUI("BrowserCreatedEx");
	const cl = await GetWinColor(window.getComputedStyle ? getComputedStyle(document.body).getPropertyValue('background-color') : document.body.currentStyle.backgroundColor);
	ArrangeAddons1(cl);
}

// Events

AddEvent("VisibleChanged", async function (Ctrl) {
	if (await Ctrl.Type == CTRL_TC) {
		const o = document.getElementById("Panel_" + await Ctrl.Id);
		if (o) {
			if (await Ctrl.Visible) {
				o.style.display = "";
				ChangeView(await Ctrl.Selected);
			} else {
				o.style.display = "none";
			}
		}
	}
});

AddEvent("SystemMessage", async function (Ctrl, hwnd, msg, wParam, lParam) {
	if (await Ctrl.Type == CTRL_WB) {
		if (msg == WM_KILLFOCUS) {
			const o = document.activeElement;
			if (o) {
				const s = o.style.visibility;
				o.style.visibility = "hidden";
				o.style.visibility = s;
				FireEvent(o, "blur");
			}
		}
	}
});

// Browser Events

AddEventEx(window, "load", async function () {
	if (await api.GetKeyState(VK_SHIFT) < 0 && await api.GetKeyState(VK_CONTROL) < 0) {
		ShowOptions("Tab=Add-ons");
	}
});

AddEventEx(window, "resize", Resize);

AddEventEx(window, "beforeunload", Finalize);

AddEventEx(window, "blur", ResetScroll);

AddEventEx(document, "MSFullscreenChange", function () {
	FullscreenChanged(document.msFullscreenElement != null);
});

AddEventEx(document, "FullscreenChange", function () {
	FullscreenChanged(document.fullscreenElement != null);
});

document.F.style.display = "none";
Init = async function () {
	te.Data.MainWindow = $;
	AddEventEx(window, "beforeunload", CloseSubWindows);
	await InitCode();
	DefaultFont = await $.DefaultFont;
	HOME_PATH = await $.HOME_PATH;
	await InitMouse();
	OpenMode = await $.OpenMode;
	await InitMenus();
	await LoadLang();
	await ApplyLang();
	await ArrangeAddons();
	ApplyLang();
	await InitWindow();
	document.F.style.display = "";
	WebBrowser.DropMode = 1;
}
