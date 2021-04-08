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
		h = ode.offsetHeight - offsetBottom - offsetTop;
		if (h < 0) {
			h = 0;
		}
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
	RunEvent1("Resize");
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
	Resize();
	setTimeout(async function () {
		ChangeView(await Ctrl.Selected);
	}, 99);
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
		} else if (Location == "Inner") {
			AddEventUI("PanelCreated", function (Ctrl, Id) {
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

FocusFV = function () {
	setTimeout(async function () {
		let el;
		if (document.activeElement) {
			const rc = document.activeElement.getBoundingClientRect();
			el = document.elementFromPoint(rc.left + 2, rc.top + 2);
		}
		if (!el || !/input|textarea/i.test(el.tagName)) {
			const hFocus = await api.GetFocus();
			if (hFocus == ui_.hwnd || await api.IsChild(ui_.hwnd, hFocus)) {
				const FV = await GetFolderView();
				if (FV) {
					if (!await api.PathMatchSpec(await api.GetClassName(hFocus), WC_EDIT + ";" + WC_TREEVIEW)) {
						FV.Focus();
					}
				}
			}
		}
	}, ui_.DoubleClickTime);
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

OnArrange = async function (Ctrl, rc) {
	const Type = await Ctrl.Type;
	if (Type == CTRL_TE) {
		ui_.TCPos = {};
	}
	await RunEvent1("Arrange", Ctrl, rc);
	if (Type == CTRL_TC) {
		const Id = await Ctrl.Id;
		let o = document.getElementById("Panel_" + Id);
		if (!o) {
			const s = ['<table id="Panel_', Id, '" class="layout" style="position: absolute; z-index: 1; color: inherit">'];
			s.push('<tr><td id="InnerLeft_', Id, '" class="sidebar" style="width: 0; display: none; overflow: auto"></td><td style="width: 100%"><div id="InnerTop_', Id, '" style="display: none"></div>');
			s.push('<table id="InnerTop2_', Id, '" class="layout">');
			s.push('<tr><td id="Inner1Left_', Id, '" class="toolbar1"></td><td id="Inner1Center_', Id, '" class="toolbar2" style="white-space: nowrap"></td><td id="Inner1Right_', Id, '" class="toolbar3"></td></tr></table>');
			s.push('<table id="InnerView_', Id, '" class="layout" style="width: 100%"><tr><td id="Inner2Left_', Id, '" style="width: 0"></td><td id="Inner2Center_', Id, '" style="width: 100%"></td><td id="Inner2Right_', Id, '" style="width: 0; overflow: auto"></td></tr></table>');
			s.push('<div id="InnerBottom_', Id, '"></div></td><td id="InnerRight_', Id, '" class="sidebar" style="width: 0; display: none"></td></tr></table>');
			document.getElementById("Panel").insertAdjacentHTML("beforeend", s.join(""));
			o = document.getElementById("Panel_" + Id);
			await PanelCreated(Ctrl, Id);
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
				return;
			}
			o.style.width = Math.max(r[2] - r[0], 0) + "px";
			o.style.height = Math.max(r[3] - r[1], 0) + "px";
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
			r[3] -= document.getElementById("InnerBottom_" + Id).offsetHeight;
			api.SetRect(rc, r[0], r[1], r[2], r[3]);
			document.getElementById("Inner2Center_" + Id).style.height = Math.max(r[3] - r[1], 0) + "px";
			te.ArrangeCB(Ctrl, rc);
		});
	}
}

ArrangeAddons = async function () {
	const r = await Promise.all([api.CreateObject("Object"), te.Data.Conf_IconSize, OpenXml("addons.xml", false, true), api.GetKeyState(VK_SHIFT), api.GetKeyState(VK_CONTROL), api.CreateObject("Array"), GetLangId()]);
	g_.Locations = r[0];
	$.IconSize = IconSize = r[1] || screen.deviceYDPI / 4;
	const xml = r[2];
	te.Data.Addons = xml;
	if (r[3] < 0 && r[4] < 0) {
		IsSavePath = function (path) {
			return false;
		}
		return;
	}
	const AddonId = {};
	const root = await xml.documentElement;
	if (root) {
		const items = await root.childNodes;
		if (items) {
			let arError = r[5];
			const LangId = r[6];
			const nLen = await GetLength(items);
			document.F.style.visibility = "hidden";
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
							await LoadAddon("vbs", Id, arError, null, window.chrome && GetNum(await item.getAttribute("Level")) < 2);
						}
						if (Enabled & 1) {
							await LoadAddon("js", Id, arError, null, window.chrome && GetNum(await item.getAttribute("Level")) < 2);
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
	if (window.chrome) {
		AddEvent("BrowserCreatedEx", "api.ShowWindow(await GetTopWindow(), SW_SHOWNORMAL);");
	}
	await RunEventUI1("Layout");
	setTimeout(async function () {
		AddEvent("Load", function () {
			te.OnArrange = OnArrange;
			document.F.style.visibility = "";
		});
		ArrangeAddons1(await GetWinColor(window.getComputedStyle ? getComputedStyle(document.body).getPropertyValue('background-color') : document.body.currentStyle.backgroundColor));
	}, 99);
}

GetMiscIcon = async function (n) {
	const s = BuildPath(ui_.DataFolder, "icons\\misc\\" + n + ".png");
	return await fso.FileExists(s) ? s : "";
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

window.addEventListener("load", async function () {
	if (await api.GetKeyState(VK_SHIFT) < 0 && await api.GetKeyState(VK_CONTROL) < 0) {
		ShowOptions("Tab=Add-ons");
	}
});

window.addEventListener("resize", Resize);

window.addEventListener("unload", FinalizeUI);

window.addEventListener("blur", ResetScroll);

window.addEventListener("mouseup", FocusFV);

document.addEventListener("MSFullscreenChange", function () {
	FullscreenChanged(document.msFullscreenElement != null);
});

document.addEventListener("FullscreenChange", function () {
	FullscreenChanged(document.fullscreenElement != null);
});

Init = async function () {
	te.Data.MainWindow = $;
	te.Data.NoCssFont = ui_.NoCssFont;
	await InitCode();
	const r = await Promise.all([te.Data.DataFolder, $.DefaultFont, $.HOME_PATH, $.OpenMode, $.DefaultFont.lfFaceName, $.DefaultFont.lfHeight, $.DefaultFont.lfWeight, InitMenus(), LoadLang()]);
	ui_.DataFolder = r[0];
	DefaultFont = r[1];
	HOME_PATH = r[2];
	OpenMode = r[3];
	document.body.style.fontFamily = r[4];
	document.body.style.fontSize = Math.abs(r[5]) + "px";
	document.body.style.fontWeight = r[6];
	await ArrangeAddons();
	await InitWindow();
	WebBrowser.DropMode = 1;
}
