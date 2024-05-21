const Addon_Id = "tabplus";
let item = GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("Icon", 1);
	item.setAttribute("Drive", 1);
	item.setAttribute("New", 1);
}
if (window.Addon == 1) {
	te.Tab = false;

	Addons.TabPlus = {
		Click: [],
		Button: [],
		Drag: [],
		Drop: [],
		pt: await api.Memory("POINT"),
		nCount: [],
		nIndex: [],
		bFlag: [],
		nFocused: -1,
		opt: {},
		str: {},
		tids: [],
		nSelected: [],

		Arrange: async function (Id, bWait) {
			delete Addons.TabPlus.tids[Id];
			const o = document.getElementById("tabplus_" + Id);
			if (o) {
				const wait = [];
				const TC = await te.Ctrl(CTRL_TC, Id);
				if (TC && await TC.Visible) {
					const nCount = await TC.Count;
					Addons.TabPlus.nIndex[Id] = await TC.SelectedIndex;
					Addons.TabPlus.nCount[Id] = nCount;
					if (o.lastChild && Addons.TabPlus.opt.New) {
						o.removeChild(o.lastChild);
					}
					let nDisp = o.getElementsByTagName("li").length;
					while (nDisp > nCount) {
						o.removeChild(o.lastChild);
						--nDisp;
					}
					while (nDisp < nCount) {
						let s = ['<li id="tabplus_', Id, '_', nDisp, '"'];
						if (Addons.TabPlus.opt.DragFolder || ui_.IEVer < 10) {
							s.push(' onmousedown="Addons.TabPlus.Down1(this)" onmousemove="Addons.TabPlus.Move(event, this)"');
						} else {
							s.push(' draggable = "true" ondragstart = "return Addons.TabPlus.Start5(event, this)" ondragend = "Addons.TabPlus.End5(event)"');
						}
						if (Addons.TabPlus.opt.Align > 1 && Addons.TabPlus.opt.Width) {
							s.push(' style="width: 100%"');
						}
						s.push('></li>');
						o.insertAdjacentHTML("beforeend", s.join(""));
						++nDisp;
					}
					if (Addons.TabPlus.opt.New) {
						let s = ['<li class="tab3" onclick="Addons.TabPlus.New(', Id, ');return false" title="', Addons.TabPlus.str.NewTab, '"'];
						if (Addons.TabPlus.opt.Align > 1 && Addons.TabPlus.opt.Width) {
							s.push(' style="text-align: center; width: 100%"');
						}
						s.push('>+</li>');
						o.insertAdjacentHTML("beforeend", s.join(""));
					}
					Addons.TabPlus.SetActiveColor(Id);
					const bRedraw = bWait || await api.GetKeyState(VK_LBUTTON) >= 0;
					for (let i = 0; i < nCount; ++i) {
						wait.push(Addons.TabPlus.Style(TC, i, bRedraw, wait));
					}
					Common.TabPlus.rc[Id] = await GetRect(o);
				}
				if (bWait) {
					await Promise.all(wait);
				}
			}
		},

		SelectionChanged: function (TC) {
			Promise.all([TC.Type, TC.Id, TC.Visible]).then(function (r) {
				if (r[0] == CTRL_TC && r[2] && !Addons.TabPlus.tids[r[1]]) {
					Addons.TabPlus.tids[r[1]] = setTimeout(function () {
						Addons.TabPlus.Arrange(r[1]);
					}, 99);
				}
			});
		},

		SetActiveColor: async function (Id) {
			const TC = await te.Ctrl(CTRL_TC);
			if (!TC || Id != await TC.Id) {
				return;
			}
			Addons.TabPlus.SetActiveColor2(Addons.TabPlus.nFocused, "");
			if (Addons.TabPlus.opt.Active) {
				Addons.TabPlus.SetActiveColor2(Id, "activecaption");
				Addons.TabPlus.nFocused = Id;
			}
		},

		SetActiveColor2: function (Id, s) {
			(document.getElementById("Panel_" + Id) || {}).className = s;
		},

		New: async function (Id) {
			const TC = await te.Ctrl(CTRL_TC, Id);
			if (TC) {
				CreateTab(await TC.Selected);
			}
		},

		Style: async function (TC, i, bRedraw, wait) {
			let r = await Promise.all([TC[i], TC.Id]);
			const FV = r[0];
			const Id = r[1];
			const o = document.getElementById("tabplus_" + Id + "_" + i);
			if (FV && o && await FV.Data) {
				const promise = [api.GetDisplayNameOf(FV, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING), RunEvent4("GetTabColor", FV), FV.Data.Lock, FV.Data.Protect, CanClose(FV), GetTabName(FV)]
				if (Addons.TabPlus.opt.Icon) {
					promise.push(GetIconImage(FV, CLR_DEFAULT | COLOR_BTNFACE));
				}
				r = await Promise.all(promise);
				const evTop = document.getElementById("tabplus_" + Id);
				const nOldHeight = evTop.offsetHeight;
				const path = r[0];
				if (Addons.TabPlus.opt.Tooltips) {
					o.title = path;
				}
				const cl = (r[1] || "").split(/\s/);
				const s = ['<table><tr class="full'];
				if (/^#/.test(cl[0]) && !cl[1]) {
					let c = Number(cl[0].replace(/^#/, "0x"));
					c = (c & 0xff0000) * .0045623779296875 + (c & 0xff00) * 2.29296875 + (c & 0xff) * 114;
					s.push(c > 127500 ? " lightbg" : " darkbg");
				}
				s.push('">');
				const r0 = Addons.TabPlus.opt.IconSize;
				let w = (Addons.TabPlus.opt.Close || r[2] || r[3]) ? -r0 : 0;
				if (!Addons.TabPlus.opt.NoLock && r[2]) {
					s.push('<td class="lockcell" style="padding-right: 2px; vertical-align: middle; width: ', r0, 'px">', Addons.TabPlus.ImgLock2, '</td>');
					w -= 2;
				} else if (Addons.TabPlus.opt.Protected && r[3]) {
					s.push('<td class="protectcell" style="padding-right: 2px; vertical-align: middle; width: ', r0, 'px">', Addons.TabPlus.ImgProtect, '</td>');
					w -= 2;
				}
				if (r[6]) {
					s.push('<td class="iconcell" style="padding-right: 3px; vertical-align: middle; width:', 20 * screen.deviceYDPI / 96, 'px">');
					if (Addons.TabPlus.opt.Drive) {
						const res = /^([A-Z]):/i.exec(path);
						if (res) {
							s.push('<span class="drive">', res[1], '</span>');
						}
					}
					const h = GetIconSize(0, 16);
					s.push('<img draggable="false" src="', r[6], '" style="width:', h, 'px; height:', h, 'px"></td>');
					w -= 20;
				} else if (Addons.TabPlus.opt.Drive) {
					s.push('<td class="drivecell" style="padding-right: 3px; vertical-align: middle; width: 12px">');
					const res = /^([A-Z]):/i.exec(path);
					if (res) {
						s.push('<span class="drive">', res[1], '</span>');
					}
					s.push('&nbsp;</td>');
					w -= 12;
				}
				s.push('<td class="namecell" style="vertical-align: middle;"><div style="overflow: hidden; white-space: nowrap;');
				const bUseClose = Addons.TabPlus.opt.Close && r[4] == S_OK;
				if (bUseClose && Addons.TabPlus.opt.Align > 1 && Addons.TabPlus.opt.Width) {
					w -= r0 * ui_.Zoom;
				}
				w += Addons.TabPlus.opt.Width;
				if (w > 0) {
					s.push((Addons.TabPlus.opt.Fix ? 'width: ' : 'max-width:'), w, 'px');
				}
				if (Addons.TabPlus.opt.Align > 1 && Addons.TabPlus.opt.Width) {
					s.push('; text-align: left; max-width: 100%');
				}
				const n = EncodeSC(r[5]);
				if (Addons.TabPlus.opt.Tooltips) {
					s.push('" title="', EncodeSC(path));
				}
				s.push('" >', n, '</div></td>');
				if (bUseClose) {
					s.push('<td class="closecell" style="vertical-align: middle; width: ', r0, 'px" align="right">', Addons.TabPlus.ImgClose.replace("`~`Arguments`~`", Id + "," + i), '</td>');
				}
				s.push('</tr></table>');
				if (!o.innerHTML || bRedraw) {
					o.innerHTML = s.join("");
				}
				const style = o.style;
				if (cl[0]) {
					if (ui_.IEVer > 9) {
						style.background = "none";
					} else {
						style.filter = "none";
					}
					style.backgroundColor = cl[0];
				} else {
					if (ui_.IEVer > 9) {
						style.background = "";
					} else if (style.filter) {
						style.filter = "";
					}
					style.backgroundColor = "";
				}
				style.color = cl[1] || "";
				const p = Addons.TabPlus.Class(TC, i, FV);
				if (wait) {
					wait.push(p);
				}
				if (nOldHeight != evTop.offsetHeight) {
					Addons.TabPlus.SetHeight(Id);
					Resize();
				}
			}
		},

		Class: async function (TC, i, FV) {
			if (!FV) {
				FV = await TC[i];
				if (!FV || !await FV.Data) {
					return;
				}
			}
			const r = await Promise.all([FV.Data.Lock, FV.Data.Protect, TC.SelectedIndex, TC.Count, TC.Id]);
			const o = document.getElementById("tabplus_" + r[4] + "_" + i);
			if (o) {
				const arClass = [];
				if (r[0]) {
					arClass.push("locked");
				}
				if (r[1]) {
					arClass.push("protected");
				}
				if (i == r[2]) {
					arClass.unshift('activetab');
					o.style.zIndex = r[3] + 1;
				} else {
					arClass.push(i < r[2] ? 'tab' : 'tab2');
					o.style.zIndex = r[3] - i;
				}
				o.className = arClass.join(" ");
				Addons.TabPlus.SetRect(r[4], i, o);
			};
		},

		SetRect: async function (Id, i, o) {
			let rcItem = await Common.TabPlus.rcItem[Id];
			if (!rcItem) {
				rcItem = await api.CreateObject("Array");
				Common.TabPlus.rcItem[Id] = rcItem;
			}
			rcItem[i] = await GetRect(o);
		},

		Down: async function (ev, Id) {
			Addons.TabPlus.buttons = (ev.buttons != null ? ev.buttons : ev.button);
			Addons.TabPlus.pt.x = ev.screenX * ui_.Zoom;
			Addons.TabPlus.pt.y = ev.screenY * ui_.Zoom;
			const TC = await te.Ctrl(CTRL_TC, Id);
			if (TC) {
				const n = await Addons.TabPlus.FromPt(Id, Addons.TabPlus.pt, true);
				Addons.TabPlus.Click = [Id, n];
				Addons.TabPlus.Button[Id] = await GetGestureButton();
				if (n >= 0) {
					if ((ev.buttons != null ? ev.buttons : ev.button) == 1) {
						TC.SelectedIndex = n;
						return false;
					}
				}
			}
			return true;
		},

		Up: async function (ev, Id) {
			const TC = await te.Ctrl(CTRL_TC, Id);
			if (TC) {
				const btn = Addons.TabPlus.Button[Id];
				if (btn == 1) {
					const n = await Addons.TabPlus.FromPt(Id, Addons.TabPlus.pt, true);
					if (n >= 0) {
						if (await Addons.TabPlus.IsClick(ev)) {
							(await TC[n]).Focus();
						}
					} else if ((ev.target || ev.srcElement).className === "tab0") {
						if (await Addons.TabPlus.IsClick(ev)) {
							Addons.TabPlus.GestureExec(TC, ev, btn);
						}
					}
				} else if (btn == 3) {
					if (await Addons.TabPlus.IsClick(ev)) {
						Addons.TabPlus.GestureExec(TC, ev, btn);
					}
				}
			}
			Addons.TabPlus.Click = [];
		},

		Down1: async function (el) {
			Addons.TabPlus.idDown = el.id;
		},

		Move: async function (ev, el) {
			const res = /^tabplus_(\d+)_(\d+)/.exec(Addons.TabPlus.idDown);
			if (res) {
				const buttons = ev.buttons != null ? ev.buttons : ev.button;
				if (buttons & 3) {
					if (!await Common.TabPlus.Drag5) {
						const pt = await api.Memory("POINT");
						pt.x = ev.screenX * ui_.Zoom;
						pt.y = ev.screenY * ui_.Zoom;
						if (await IsDrag(pt, Addons.TabPlus.pt) && !await IsDrag(await g_.mouse.ptDown, Addons.TabPlus.pt)) {
							Common.TabPlus.Drag5 = Addons.TabPlus.idDown;
							const DataObj = await api.CreateObject("FolderItems");
							DataObj.AddItem(await te.Ctrl(CTRL_TC, res[1])[res[2]].FolderItem);
							DataObj.dwEffect = DROPEFFECT_LINK;
							DoDragDrop(DataObj, DROPEFFECT_LINK | DROPEFFECT_COPY | DROPEFFECT_MOVE, false, function () {
								Common.TabPlus.Drag5 = void 0;
								Addons.TabPlus.pt = pt;
							});
						}
					}
				} else {
					Common.TabPlus.Drag5 = void 0;
				}
			}
		},

		Popup: async function (ev, Id) {
			const pt = await api.Memory("POINT");
			pt.x = ev.screenX * ui_.Zoom;
			pt.y = ev.screenY * ui_.Zoom;
			const TC = await te.Ctrl(CTRL_TC, Id);
			if (TC) {
				te.OnShowContextMenu(TC, await TC.hwnd, WM_CONTEXTMENU, 0, pt);
			}
		},

		GestureExec: async function (TC, ev, s) {
			if (TC) {
				const pt = await api.Memory("POINT");
				pt.x = ev.screenX * ui_.Zoom;
				pt.y = ev.screenY * ui_.Zoom;
				s = await GetGestureKey() + s;
				if (await TC.HitTest(pt, TCHT_ONITEM) < 0) {
					if (await GestureExec(TC, "Tabs_Background", s, pt, await TC.hwnd) === S_OK) {
						return;
					}
				}
				GestureExec(TC, "Tabs", s, pt, await TC.hwnd);
			}
		},

		DblClick: async function (ev, Id) {
			if (await Addons.TabPlus.IsClick(ev)) {
				const TC = await te.Ctrl(CTRL_TC, Id);
				Addons.TabPlus.GestureExec(TC, ev, Addons.TabPlus.Button[Id] + Addons.TabPlus.Button[Id]);
			}
		},

		IsClick: async function (ev) {
			const pt = await api.Memory("POINT");
			pt.x = ev.screenX * ui_.Zoom;
			pt.y = ev.screenY * ui_.Zoom;
			return !await IsDrag(pt, Addons.TabPlus.pt);
		},

		Select: function (id, nIndex) {
			te.Ctrl(CTRL_TC, id).SelectedIndex = nIndex;
		},

		Start5: function (ev, o) {
			if (Addons.TabPlus.buttons == 1) {
				ev.dataTransfer.effectAllowed = 'move';
				if (window.chrome) {
					ev.dataTransfer.setData("text/plain", o.title + "\n" + o.id);
				}
				Common.TabPlus.Drag5 = o.id;
				return true;
			}
			return false;
		},

		End5: async function (ev) {
			if (await api.GetKeyState(VK_RBUTTON) >= 0 && await api.GetKeyState(VK_ESCAPE) >= 0) {
				const res = /^tabplus_(\d+)_(\d+)/.exec(await Common.TabPlus.Drag5);
				if (res) {
					const pt = await api.Memory("POINT");
					pt.x = ev.screenX * ui_.Zoom;
					pt.y = ev.screenY * ui_.Zoom;
					const hwnd = await api.WindowFromPoint(pt);
					if (ui_.hwnd != hwnd && !await api.IsChild(ui_.hwnd, hwnd)) {
						Sync.TabPlus.DropTab(await te.Ctrl(CTRL_TC, res[1])[res[2]], hwnd, pt);
					}
				}
			}
			Common.TabPlus.Drag5 = void 0;
		},

		Close: async function (Id, i) {
			const FV = await te.Ctrl(CTRL_TC, Id)[i];
			FV.Close();
		},

		Wheel: async function (ev, Id) {
			const TC = await te.Ctrl(CTRL_TC, Id);
			if (TC) {
				const o = document.getElementById("tabplus_" + Id);
				if (o.clientWidth == o.offsetWidth) {
					Addons.TabPlus.GestureExec(TC, ev, ev.wheelDelta > 0 ? "8" : "9");
				}
			}
		},

		FromPt: async function (Id, pt, bNoClose) {
			const x = await pt.x - screenLeft * ui_.Zoom;
			const y = await pt.y - screenTop * ui_.Zoom;
			const re = new RegExp("tabplus_" + Id + "_(\\d+)", "");
			for (let el = document.elementFromPoint(x, y); el && !/UL/i.test(el.tagName); el = el.parentElement) {
				if (bNoClose && /closecell/.test(el.className)) {
					return -1;
				}
				const res = re.exec(el.id);
				if (res) {
					return res[1];
				}
			}
			return -1;
		},

		Resize: function () {
			if (Addons.TabPlus.tidResize) {
				return;
			}
			Addons.TabPlus.tidResize = setTimeout(async function () {
				delete Addons.TabPlus.tidResize;
				const cTC = await te.Ctrls(CTRL_TC, true, window.chrome);
				for (let j = cTC.length; j-- > 0;) {
					const Id = await cTC[j].Id;
					Addons.TabPlus.Arrange(Id);
					Addons.TabPlus.SetHeight(Id);
				}
			}, 500);
		},


		SetHeight: function (Id) {
			if (Addons.TabPlus.opt.Align > 1) {
				const elTab = document.getElementById("tabplus_" + Id);
				const elPanel = document.getElementById("Panel_" + Id);
				if (elTab && elPanel) {
					elTab.style.height = elPanel.style.height;
				}
			}
		},

		Over: async function (Id) {
			if (await Common.TabPlus.bDropping) {
				return;
			}
			const pt = await api.GetCursorPos();
			if (!await IsDrag(pt, await g_.ptDrag)) {
				const nIndex = await Addons.TabPlus.FromPt(Id, pt);
				if (nIndex >= 0) {
					const TC = await te.Ctrl(CTRL_TC, Id);
					if (Addons.TabPlus.opt.NoDragOpen) {
						setTimeout(Addons.TabPlus.Select, 500, Id, await TC.SelectedIndex);
					}
					TC.SelectedIndex = nIndex;
				}
			}
		},

		DragOver: function (Id) {
			if (Addons.TabPlus.tid) {
				clearTimeout(Addons.TabPlus.tid);
			}
			Addons.TabPlus.tid = setTimeout(Addons.TabPlus.Over, 300, Id);
		},

		DragLeave: function () {
			if (Addons.TabPlus.tid) {
				clearTimeout(Addons.TabPlus.tid);
				delete Addons.TabPlus.tid;
			}
		},

		SetRects: async function () {
			const cTC = await te.Ctrls(CTRL_TC, true, window.chrome);
			for (let i = 0; i < cTC.length; ++i) {
				const TC = cTC[i];
				const Id = await TC.Id;
				let o = document.getElementById("tabplus_" + Id);
				if (o) {
					Common.TabPlus.rc[Id] = await GetRect(o);
					const nCount1 = await TC.Count;
					Common.TabPlus.rcItem[Id] = await api.CreateObject("Array");
					for (let j = 0; j < nCount1; ++j) {
						o = document.getElementById("tabplus_" + Id + "_" + j);
						if (o) {
							Addons.TabPlus.SetRect(Id, j, o);
						}
					}
				}
			}
		}
	};

	Common.TabPlus = await api.CreateObject("Object");
	Common.TabPlus.opt = await api.CreateObject("Object");

	$.importScript("addons\\" + Addon_Id + "\\sync.js");

	AddEvent("PanelCreated", async function (Ctrl, Id) {
		const s = ['<ul id="tabplus_', Id, '" class="tab0" oncontextmenu="Addons.TabPlus.Popup(event,', Id, ');return false"'];
		s.push(' ondblclick="Addons.TabPlus.DblClick(event,', Id, ');return false" onmousewheel="Addons.TabPlus.Wheel(event,', Id, ')" onresize="Resize();"');
		s.push(' onmousedown="Addons.TabPlus.Down(event,', Id, ')" onmouseup="return Addons.TabPlus.Up(event,', Id, ')" onclick="return false;" style="width: 100%"></ul>');
		const n = Addons.TabPlus.opt.Align || 0;
		const arAlign = ["InnerTop_", "InnerBottom_", "InnerLeft_", "InnerRight_"];
		let o = document.getElementById(await SetAddon(null, arAlign[n] + Id, s.join("")));
		if (n > 1) {
			const w = (Number(Addons.TabPlus.opt.Width || 84) + 17 * screen.deviceYDPI / 96) + "px";
			o.style.width = w;
			o = document.getElementById("tabplus_" + Id);
			o.style.width = w;
			o.style.height = "0";
			o.style.overflow = "auto";
		} else {
			o.style.overflow = "hidden";
		}
		o.style.overflowX = "hidden";
		await Addons.TabPlus.Arrange(Id, true);
	});

	AddEvent("Lock", function (Ctrl, i, bLock) {
		Addons.TabPlus.Style(Ctrl, i, true)
	});

	AddEvent("VisibleChanged", async function (Ctrl, Visible) {
		Promise.all([Ctrl.Type, Ctrl.Visible, Ctrl.Id]).then(function (r) {
			if (r[0] == CTRL_TC) {
				if (r[1]) {
					Addons.TabPlus.SetActiveColor(r[2]);
				}
				Addons.TabPlus.Resize();
			}
		});
	});

	AddEvent("SelectionChanging", function (Ctrl) {
		Promise.all([Ctrl.Type, Ctrl.Id, Ctrl.SelectedIndex]).then(function (r) {
			if (r[0] == CTRL_TC) {
				Addons.TabPlus.nSelected[r[1]] = r[2];
			}
		});
	});

	AddEvent("SelectionChanged", function (Ctrl) {
		Promise.all([Ctrl.Type, Ctrl.Id, Ctrl.SelectedIndex]).then(function (r) {
			if (r[0] == CTRL_TC) {
				Addons.TabPlus.Class(Ctrl, Addons.TabPlus.nSelected[r[1]]);
				Addons.TabPlus.Class(Ctrl, r[2]);
				Addons.TabPlus.Arrange(r[1]);
			}
		});
	});

	AddEvent("Arrange", async function (Ctrl) {
		if (await Ctrl.Type == CTRL_TC) {
			Addons.TabPlus.Arrange(await Ctrl.Id);
		}
	});

	AddEvent("ChangeView", async function (Ctrl) {
		const TC = await Ctrl.Parent;
		if (TC) {
			const Id = await TC.Id;
			const i = Addons.TabPlus.nIndex[Id];
			let o = document.getElementById("tabplus_" + Id + "_" + i);
			if (o) {
				await Addons.TabPlus.Style(TC, i, true)
				o = document.getElementById("tabplus_" + Id);
				if (o) {
					if (await TC.Count + (Addons.TabPlus.opt.New ? 1 : 0) != o.getElementsByTagName("li").length) {
						o = null;
					}
				}
			}
			if (!o) {
				Addons.TabPlus.SelectionChanged(TC);
			}
		}
	});

	AddEvent("Create", async function (Ctrl) {
		Addons.TabPlus.SelectionChanged(Ctrl);
	});

	AddEvent("CloseView", async function (Ctrl) {
		const TC = await Ctrl.Parent;
		if (TC) {
			Addons.TabPlus.SelectionChanged(TC);
		}
	});

	AddEvent("Resize", Addons.TabPlus.Resize);

	const attrs = item.attributes;
	for (let i = attrs.length; i-- > 0;) {
		Common.TabPlus.opt[attrs[i].name] = Addons.TabPlus.opt[attrs[i].name] = attrs[i].value;
	}
	Addons.TabPlus.opt.Width = GetNum(Addons.TabPlus.opt.Width);
	Addons.TabPlus.opt.IconSize = GetNum(Addons.TabPlus.opt.IconsSize) || 13;
	Addons.TabPlus.ImgLock = Addons.TabPlus.opt.IconLock || ("string" === ui_.MiscIcon["lock"] ? ui_.MiscIcon["lock"] : await GetMiscIcon("lock")) || GetWinIcon(0x602, "font:Segoe UI Emoji,0x1f4cc", 0, "bitmap:ieframe.dll,545,13,2");
	const r0 = Math.round(Addons.TabPlus.opt.IconSize * screen.deviceYDPI / 96);
	const r = [GetImgTag({
		draggable: "false",
		src: Addons.TabPlus.ImgLock
	}, r0),
	GetImgTag({
		draggable: "false",
		onclick: "Addons.TabPlus.Close(`~`Arguments`~`)",
		"class": "button",
		title: await GetText("Close tab"),
		onmouseover: "MouseOver(this)",
		onmouseout: "MouseOut()",
		src: Addons.TabPlus.opt.IconClose || "font:Marlett,0x72"
	}, r0),
	GetImgTag({
		draggable: "false",
		src: Addons.TabPlus.opt.IconProtect || GetWinIcon(0xa00, "font:Segoe MDL2 Assets,0xea18", 0x602, "font:Segoe UI Emoji,0x26c9", 0, "font:Webdings,0x64")
	}, r0)];
	if (Addons.TabPlus.opt.New) {
		r.push(GetText("New tab"));
	}
	Promise.all(r).then(function (r) {
		Addons.TabPlus.ImgLock2 = r.shift();
		Addons.TabPlus.ImgClose = r.shift();
		Addons.TabPlus.ImgProtect = r.shift();
		Addons.TabPlus.str.NewTab = r.shift();
	});
	delete item;
} else {
	importScript("addons\\" + Addon_Id + "\\options.js");
}
