var Addon_Id = "tabplus";

var item = GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("Icon", 1);
	item.setAttribute("Drive", 1);
	item.setAttribute("New", 1);
}
if (window.Addon == 1) {
	Addons.TabPlus =
	{
		Click: [],
		Button: [],
		Drag: [],
		Drop: [],
		pt: api.Memory("POINT"),
		nCount: [],
		nIndex: [],
		bFlag: [],
		nFocused: -1,
		opt: [],
		tids: [],
		tidResize: null,
		tidCursor: null,

		Arrange: function (Id) {
			delete Addons.TabPlus.tids[Id];
			var o = document.getElementById("tabplus_" + Id);
			if (o) {
				var TC = te.Ctrl(CTRL_TC, Id);
				if (TC && TC.Visible) {
					Addons.TabPlus.tids[Id] = null;
					this.nIndex[Id] = TC.SelectedIndex;
					this.nCount[Id] = TC.Count;
					var tabs = o.getElementsByTagName("li");
					if (TC.Count + (this.opt.New ? 1 : 0) != tabs.length) {
						var s = [];
						for (var i = 0; i < this.nCount[Id]; i++) {
							this.Tab(s, TC, i);
						}
						if (this.opt.New) {
							s.push('<li class="tab3" onclick="Addons.TabPlus.New(', Id, ');return false" title="', this.opt.Tooltips ? GetText("New Tab") : "", '"');
							if (this.opt.Align > 1 && this.opt.Width) {
								s.push(' style="text-align: center; width: 100%"');
							}
							s.push(' />+</li>');
						}
						o.innerHTML = s.join("").replace(/\$/g, Id);
					}
					try {
						var FocusedId = te.Ctrl(CTRL_TC).Id;
						if (Id == FocusedId) {
							this.SetActiveColor(Id);
						}
					} catch (e) { }
					if (g_.IEVer < 10) {
						var n = (this.dtDown || 0) + 90000 - new Date().getTime();
						if (n >= 0) {
							var pt = api.Memory("POINT");
							api.GetCursorPos(pt);
							n = this.FromPt(Id, this.pt);
						}
						if (n >= 0) {
							if (this.Drag.length) {
								this.Cursor2("move");
							}
						} else {
							this.Drag = [];
							this.Cursor("default");
						}
					}
					for (var i = this.nCount[Id]; i--;) {
						this.Style(TC, i);
					}
				}
			}
		},

		SelectionChanged: function (TC, Id) {
			if (TC.Type == CTRL_TC && TC.Visible && !Addons.TabPlus.tids[TC.Id]) {
				Addons.TabPlus.tids[TC.Id] = setTimeout(function () {
					Addons.TabPlus.Arrange(TC.Id);
				}, 99);
			}
		},

		SetActiveColor: function (Id) {
			this.SetActiveColor2(this.nFocused, "");
			if (this.opt.Active) {
				this.SetActiveColor2(Id, "ActiveCaption");
				this.nFocused = Id;
			}
		},

		SetActiveColor2: function (Id, s) {
			var o = document.getElementById("Panel_" + Id);
			if (o) {
				o.style.backgroundColor = s;
			}
		},

		New: function (Id) {
			var TC = te.Ctrl(CTRL_TC, Id);
			if (TC) {
				var FV = TC.Selected;
				CreateTab(FV);
				TC.Move(TC.SelectedIndex, TC.Count - 1);
			}
		},

		Tab: function (s, TC, i) {
			var FV = TC[i];
			if (FV) {
				s.push('<li id="tabplus_$_', i, '" draggable="true" ondragstart="return Addons.TabPlus.Start5(this)" ondragend="Addons.TabPlus.End5(this)" onmousemove="Addons.TabPlus.Move(this, $)"');
				if (this.opt.Align > 1 && this.opt.Width) {
					s.push(' style="width: 100%"');
				}
				s.push('></li>');
			}
		},

		Style: function (TC, i) {
			var img;
			var FV = TC[i];
			var o = document.getElementById("tabplus_" + TC.Id + "_" + i);
			if (FV && o) {
				if (g_.IEVer >= 10) {
					if ((this.dtDown || 0) + 90000 - new Date().getTime() > 0) {
						var pt = api.Memory("POINT");
						api.GetCursorPos(pt);
						if (i == this.FromPt(TC.Id, this.pt)) {
							return;
						}
					}
				}
				var path = api.GetDisplayNameOf(FV.FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
				if (this.opt.Tooltips) {
					o.title = path;
				}
				var s = ['<table style="width: 100%"><tr style="width: 100%">'];
				try {
					var r0 = Math.round(13 * screen.deviceYDPI / 96);
					var w = (FV.Data.Lock || FV.Data.Protect || this.opt.Close) ? -r0 : 0;
					if (FV.Data.Lock && !this.opt.NoLock) {
						s.push('<td style="padding-right: 2px; vertical-align: middle; width: ', r0, 'px"><img draggable="false" src="', this.ImgLock, '" style="width: ', r0, 'px"></td>');
						w -= 2;
					} else if (FV.Data.Protect && this.opt.Protected) {
						s.push('<td style="padding-right: 2px; vertical-align: middle; width: ', r0, 'px"><img draggable="false" src="', this.ImgProtect, '" style="width: ', r0, 'px"></td>');
						w -= 2;
					}
					if (this.opt.Icon && (img = GetIconImage(FV, GetSysColor(COLOR_BTNFACE)))) {
						s.push('<td style="padding-right: 3px; vertical-align: middle; width: 20px">');
						if (this.opt.Drive) {
							var res = /^([A-Z]):/i.exec(path);
							if (res) {
								s.push('<span class="drive">', res[1], '</span>');
							}
						}
						s.push('<img draggable="false" src="', img, '" style="width: 1pc; height: 1pc"></td>');
						w -= 20;
					} else if (this.opt.Drive) {
						s.push('<td style="padding-right: 3px; vertical-align: middle; width: 12px">');
						var res = /^([A-Z]):/i.exec(path);
						if (res) {
							s.push('<span class="drive">', res[1], '</span>');
						}
						s.push('&nbsp;</td>');
						w -= 12;
					}
					s.push('<td style="vertical-align: middle;"><div style="overflow: hidden; white-space: nowrap;');
					if (this.opt.Close && CanClose(FV) == S_OK && this.opt.Align > 1 && this.opt.Width) {
						w -= r0;
					}
					w += Number(this.opt.Width) || 0;
					if (w > 0) {
						s.push((this.opt.Fix ? 'width: ' : 'max-width:'), w, 'px');
					}
					if (this.opt.Align > 1 && this.opt.Width) {
						s.push('; text-align: left; max-width: 100%');
					}
					var n = "";
					if (FV.FolderItem) {
						n = EncodeSC(GetTabName(FV));
						if (this.opt.Tooltips) {
							s.push('" title="', EncodeSC(FV.FolderItem.Path));
						}
					}
					s.push('" >', n, '</div></td>');
					if (this.opt.Close && CanClose(FV) == S_OK) {
						s.push('<td style="vertical-align: middle; width: ', r0, 'px" align="right">', this.ImgClose, r0, 'px" id="tabplus_', FV.Parent.Id, '_', i, 'x" class="button" title="', this.opt.Tooltips ? GetText("Close Tab") : "", '" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', this.ImgClose2, '</td>');
					}
				} catch (e) { }
				s.push('</tr></table>');
				o.innerHTML = s.join("");
				var style = o.style;
				var cl = RunEvent4("GetTabColor", FV);
				if (cl) {
					if (g_.IEVer >= 10) {
						style.background = "none";
					} else {
						style.filter = "none";
					}
					style.backgroundColor = cl;
					cl = api.sscanf(cl, "#%06x")
					cl = (cl & 0xff0000) * .0045623779296875 + (cl & 0xff00) * 2.29296875 + (cl & 0xff) * 114;
					style.color = cl > 127000 ? "black" : "white";
				} else {
					if (document.documentMode >= 10) {
						style.background = "";
					} else if (style.filter) {
						style.filter = "";
					}
					style.color = "";
					style.backgroundColor = "";
				}
				if (i == TC.SelectedIndex) {
					o.className = 'activetab';
					style.zIndex = TC.Count + 1;
				} else {
					o.className = i < TC.SelectedIndex ? 'tab' : 'tab2';
					style.zIndex = TC.Count - i;
				}
			}
		},

		Down: function (Id) {
			if (this.tidDrag) {
				clearTimeout(this.tidDrag);
				delete this.tidDrag;
			}
			this.dtDown = new Date().getTime();
			var TC = te.Ctrl(CTRL_TC, Id);
			if (TC) {
				api.GetCursorPos(this.pt);
				var n = this.FromPt(Id, this.pt);
				this.Click = [Id, n];
				this.Drag = [];
				this.Button[Id] = GetGestureButton();
				if (n >= 0) {
					if (api.GetKeyState(VK_LBUTTON) < 0) {
						var o = document.getElementById('tabplus_' + Id + '_' + n + 'x');
						if (o && HitTest(o, this.pt)) {
						} else {
							TC.SelectedIndex = n;
							var o = document.getElementById("tabplus_" + TC.Id + "_" + n);
							o.className = "activetab";
						}
						if (g_.IEVer < 10) {
							(function (Id, n) {
								this.tidDrag = setTimeout(function () {
									delete Addons.TabPlus.tidDrag;
									if (api.GetKeyState(VK_LBUTTON) < 0) {
										var pt = api.Memory("POINT");
										api.GetCursorPos(pt);
										if (n == Addons.TabPlus.FromPt(Id, pt)) {
											Addons.TabPlus.Cursor2("move");
											Addons.TabPlus.Drag = Addons.TabPlus.Click.slice(0);
										}
									}
								}, 99);
							})(Id, n);
						}
						return false;
					}
				}
			}
			return true;
		},

		Up: function (Id) {
			delete this.dtDown;
			var TC = te.Ctrl(CTRL_TC, Id);
			if (TC) {
				var pt = api.Memory("POINT");
				api.GetCursorPos(pt);
				if (/3/.test(this.Button[Id])) {
					Addons.TabPlus.GestureExec(TC, Id, this.Button[Id], pt);
					return false;
				}
				var nDrop = this.FromPt(Id, pt);
				if (nDrop < 0) {
					nDrop = TC.Count;
				} else if (/1/.test(this.Button[Id]) && !IsDrag(pt, this.pt)) {
					var o = document.getElementById('tabplus_' + Id + '_' + nDrop + 'x');
					if (o && HitTest(o, this.pt)) {
						TC[nDrop].Close();
					} else {
						setTimeout(function () {
							Addons.TabPlus.Arrange(Id);
						}, wsh.RegRead("HKCU\\Control Panel\\Mouse\\DoubleClickSpeed"));
					}
				}
				if (this.Drag.length && (this.Drag[0] != Id || this.Drag[1] != nDrop)) {
					if (g_.IEVer < 10) {
						te.Ctrl(CTRL_TC, this.Drag[0]).Move(this.Drag[1], nDrop, TC);
					}
					this.Drop = [];
				} else {
					setTimeout(function () {
						TC.Selected.Focus();
					}, 99);
				}
				if (this.Drag.length) {
					this.Cursor("default");
					this.Drag = [];
				}
			}
			this.Click = [];
			return true;
		},

		Move: function (o, Id) {
			if (this.Drag.length && g_.IEVer < 10) {
				if (api.GetKeyState(VK_LBUTTON) < 0) {
					this.Drop = [Id, o.id.replace(/^.*_\d+_/, '') - 0];
				} else {
					this.Cursor("default");
					this.Drag = [];
				}
			}
		},

		Cursor: function (s) {
			if (g_.IEVer < 10) {
				clearTimeout(Addons.TabPlus.tidCursor);
				var cTC = te.Ctrls(CTRL_TC);
				for (var j in cTC) {
					if (cTC[j].Visible) {
						SetCursor(document.getElementById('tabplus_' + cTC[j].Id), s);
					}
				}
			}
		},

		Cursor2: function (s) {
			this.tidCursor = setTimeout(function () {
				Addons.TabPlus.Cursor(s)
			}, 500);
		},

		Popup: function (Id) {
			var pt = api.Memory("POINT");
			api.GetCursorPos(pt);
			var TC = te.Ctrl(CTRL_TC, Id);
			if (TC) {
				var ShowContextMenu = te.OnShowContextMenu;
				ShowContextMenu(TC, TC.hwnd, WM_CONTEXTMENU, 0, pt);
			}
		},

		GestureExec: function (TC, Id, s, pt) {
			if (TC) {
				if (TC.HitTest(pt, TCHT_ONITEM) < 0) {
					if (GestureExec(TC, "Tabs_Background", GetGestureKey() + s, pt, TC.hwnd) === S_OK) {
						;
						return;
					}
				}
				GestureExec(TC, "Tabs", GetGestureKey() + s, pt, TC.hwnd);
			}
		},

		DblClick: function (Id) {
			api.GetCursorPos(pt);
			var TC = te.Ctrl(CTRL_TC, Id);
			Addons.TabPlus.GestureExec(TC, Id, this.Button[Id] + this.Button[Id], pt);
		},

		Over: function (Id) {
			var pt = api.Memory("POINT");
			api.GetCursorPos(pt);
			if (!IsDrag(pt, g_ptDrag) && !this.Drag5) {
				var nIndex = this.FromPt(Id, pt);
				if (nIndex >= 0) {
					te.Ctrl(CTRL_TC, Id).SelectedIndex = nIndex;
				}
			}
		},

		Start5: function (o) {
			if (api.GetKeyState(VK_LBUTTON) < 0) {
				clearTimeout(Addons.TabPlus.tidCursor);
				event.dataTransfer.effectAllowed = 'move';
				event.dataTransfer.setData("text", o.title);
				this.Drag5 = o.id;
				delete this.dtDown;
				return true;
			}
			return false;
		},

		End5: function (o) {
			var pt = api.Memory("POINT");
			api.GetCursorPos(pt);
			var hwnd = api.WindowFromPoint(pt);
			if (te.hwnd != hwnd && !api.IsChild(te.hwnd, hwnd) && api.GetKeyState(VK_RBUTTON) >= 0 && api.GetKeyState(VK_ESCAPE) >= 0) {
				var res = /^tabplus_(\d+)_(\d+)/.exec(this.Drag5);
				if (res) {
					var FV = te.Ctrl(CTRL_TC, res[1])[res[2]];
					OpenInExplorer(FV);
					FV.Close();
				}
			}
			this.Drag5 = null;
			this.Drag = [];
		},

		Over5: function (o) {
			if (this.Drag5) {
				if (event.preventDefault) {
					event.preventDefault();
				} else {
					event.returnValue = false;
				}
			}
		},

		Drop5: function (o) {
			var res = /^tabplus_(\d+)/.exec(o.id);
			if (res) {
				var Id = res[1];
				var TC = te.Ctrl(CTRL_TC, Id);
				if (TC) {
					var pt = api.Memory("POINT");
					api.GetCursorPos(pt);
					var nDrop = this.FromPt(Id, pt);
					if (nDrop < 0) {
						nDrop = TC.Count;
					}
					res = /^tabplus_(\d+)_(\d+)/.exec(Addons.TabPlus.Drag5);
					if (res) {
						if (res[1] != Id || res[2] != nDrop) {
							var TC1 = te.Ctrl(CTRL_TC, res[1]);
							TC1.Move(res[2], nDrop, TC);
							TC1.SelectedIndex = nDrop;
							this.Drop = [];
						}
					}
				}
			}
		},

		Wheel: function (Id) {
			var TC = te.Ctrl(CTRL_TC, Id);
			if (TC) {
				var o = document.getElementById("tabplus_" + Id);
				if (o.clientWidth == o.offsetWidth) {
					var pt = api.Memory("POINT");
					api.GetCursorPos(pt);
					Addons.TabPlus.GestureExec(TC, Id, event.wheelDelta > 0 ? "8" : "9", pt)
				}
			}
		},

		FromPt: function (Id, pt) {
			var ptc = pt.Clone();
			api.ScreenToClient(api.GetWindow(document), ptc);
			var re = new RegExp("tabplus_" + Id + "_(\\d+)", "");
			for (var el = document.elementFromPoint(ptc.x, ptc.y); el && !/UL/i.test(el.tagName); el = el.parentElement) {
				var res = re.exec(el.id);
				if (res) {
					return res[1];
				}
			}
			return -1;
		},

		TCFromPt: function (pt) {
			var cTC = te.Ctrls(CTRL_TC);
			for (var n in cTC) {
				var TC = cTC[n];
				if (TC.Visible) {
					if (HitTest(document.getElementById("tabplus_" + TC.Id), pt)) {
						return TC;
					}
				}
			}
		},

		Resize: function () {
			if (this.opt.Align > 1 && !this.tidResize) {
				this.tidResize = setTimeout(function () {
					Addons.TabPlus.tidResize = null;
					var cTC = te.Ctrls(CTRL_TC);
					for (var j in cTC) {
						var TC = cTC[j];
						if (TC.Visible) {
							var id = TC.Id;
							var o = document.getElementById("Panel_" + id);
							if (o) {
								document.getElementById("tabplus_" + id).style.height = o.clientHeight + "px";
							}
						}
					}
				}, 500);
			}
		}
	};

	AddEvent("PanelCreated", function (Ctrl) {
		var s = ['<ul id="tabplus_$" class="tab0" oncontextmenu="Addons.TabPlus.Popup($);return false"'];
		s.push(' ondblclick="Addons.TabPlus.DblClick($);return false" onmousewheel="Addons.TabPlus.Wheel($)" onresize="Resize();"');
		s.push(' onmousedown="Addons.TabPlus.Down($)" onmouseup="return Addons.TabPlus.Up($)" onclick="return false;" ondragover="Addons.TabPlus.Over5(this)" ondrop="Addons.TabPlus.Drop5(this)" style="width: 100%"></ul>');
		var n = Addons.TabPlus.opt.Align || 0;
		var arAlign = ["InnerTop_", "InnerBottom_", "InnerLeft_", "InnerRight_"];
		var o = document.getElementById(SetAddon(null, arAlign[n] + Ctrl.Id, s.join("").replace(/\$/g, Ctrl.Id)));
		if (n > 1) {
			var h = o.innerHeight;
			var w = (Number(Addons.TabPlus.opt.Width || 84) + 17) + "px";
			o.style.width = w;
			o = document.getElementById("tabplus_" + Ctrl.Id);
			o.style.width = w;
			o.style.height = "0";
			o.style.overflow = "auto";
		} else {
			o.style.overflow = "hidden";
		}
		o.style.overflowX = "hidden";
	});

	AddEvent("HitTest", function (Ctrl, pt, flags) {
		if (Ctrl.Type == CTRL_TC) {
			return Addons.TabPlus.FromPt(Ctrl.Id, pt);
		}
	});

	AddEvent("Lock", function (Ctrl, i, bLock) {
		Addons.TabPlus.Style(Ctrl, i)
	});

	AddEvent("DragEnter", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
		if (Ctrl.Type == CTRL_WB) {
			return S_OK;
		}
	});

	AddEvent("DragOver", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
		if (Ctrl.Type == CTRL_WB) {
			var TC = Addons.TabPlus.TCFromPt(pt);
			if (TC) {
				var nIndex = Addons.TabPlus.FromPt(TC.Id, pt);
				if (nIndex >= 0) {
					if (IsDrag(pt, g_ptDrag)) {
						clearTimeout(Addons.TabPlus.tid);
						g_ptDrag = pt.Clone();
						Addons.TabPlus.tid = setTimeout("Addons.TabPlus.Over(" + TC.Id + ");", 300);
					}
					if (dataObj.Count) {
						var Target = TC[nIndex].FolderItem;
						if (!api.ILIsEqual(dataObj.Item(-1), Target)) {
							var DropTarget = api.DropTarget(Target);
							if (DropTarget) {
								return DropTarget.DragOver(dataObj, grfKeyState, pt, pdwEffect);
							}
						}
					}
					pdwEffect[0] = DROPEFFECT_NONE;
					return S_OK;
				} else if (dataObj.Count && dataObj.Item(0).IsFolder) {
					pdwEffect[0] = DROPEFFECT_LINK;
					return S_OK;
				}
			}
		}
	});

	AddEvent("Drop", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
		if (Ctrl.Type == CTRL_WB && !Addons.TabPlus.Drag5) {
			var TC = Addons.TabPlus.TCFromPt(pt);
			if (TC) {
				var nIndex = Addons.TabPlus.FromPt(TC.Id, pt);
				if (nIndex >= 0) {
					var hr = S_FALSE;
					var DropTarget = TC[nIndex].DropTarget;
					if (DropTarget) {
						clearTimeout(Addons.TabPlus.tid);
						hr = DropTarget.Drop(dataObj, grfKeyState, pt, pdwEffect);
					}
					return hr;
				} else if (dataObj.Count) {
					for (var i = 0; i < dataObj.Count; i++) {
						var FV = TC.Selected.Navigate(dataObj.Item(i), SBSP_NEWBROWSER);
						TC.Move(FV.Index, TC.Count - 1);
					}
				}
			}
		}
	});

	AddEvent("DragLeave", function (Ctrl) {
		clearTimeout(Addons.TabPlus.tid);
		Addons.TabPlus.tid = null;
		return S_OK;
	});

	AddEvent("VisibleChanged", function (Ctrl) {
		if (Ctrl.Type == CTRL_TC) {
			if (Ctrl.Visible) {
				Addons.TabPlus.SetActiveColor(Ctrl.Id);
			}
			Addons.TabPlus.Resize();
		}
	});

	AddEvent("SelectionChanged", function (Ctrl) {
		if (Ctrl.Type == CTRL_TC) {
			Addons.TabPlus.Arrange(Ctrl.Id);
		}
	});

	AddEvent("Arrange", function (Ctrl) {
		if (Ctrl.Type == CTRL_TC) {
			Addons.TabPlus.Arrange(Ctrl.Id);
		}
	});

	AddEvent("ChangeView", function (Ctrl) {
		var TC = Ctrl.Parent;
		if (TC) {
			var i = Addons.TabPlus.nIndex[TC.Id];
			var o = document.getElementById("tabplus_" + TC.Id + "_" + i);
			if (o) {
				Addons.TabPlus.Style(TC, i)
				var o = document.getElementById("tabplus_" + TC.Id);
				if (o) {
					var tabs = o.getElementsByTagName("li");
					if (TC.Count + (Addons.TabPlus.opt.New ? 1 : 0) != tabs.length) {
						o = null;
					}
				}
			}
			if (!o) {
				Addons.TabPlus.SelectionChanged(TC, TC.Id);
			}
		}
	});

	AddEvent("Create", function (Ctrl) {
		Addons.TabPlus.SelectionChanged(Ctrl, Ctrl.Id);
	});

	AddEvent("CloseView", function (Ctrl) {
		var TC = Ctrl.Parent;
		if (TC) {
			Addons.TabPlus.SelectionChanged(TC, TC.Id);
		}
	});

	AddEvent("Resize", function () {
		Addons.TabPlus.Resize();
	});

	AddEvent("MouseMessage", function (Ctrl, hwnd, msg, mouseData, pt, wHitTestCode, dwExtraInfo) {
		if (msg == WM_MBUTTONDOWN || msg == WM_MBUTTONUP) {
			if (Ctrl.Type == CTRL_WB) {
				var TC = Addons.TabPlus.TCFromPt(pt);
				if (TC) {
					if (msg == WM_MBUTTONDOWN) {
						Addons.TabPlus.Down(TC.Id);
						Addons.TabPlus.Button[TC.Id] = GetGestureKey().replace(/3/, "") + "3";
					} else {
						Addons.TabPlus.Up(TC.Id);
					}
					return S_OK;
				}
			}
		}
	});

	//Init
	te.Tab = false;
	var attrs = item.attributes;
	for (var i = attrs.length; i-- > 0;) {
		Addons.TabPlus.opt[attrs[i].name] = attrs[i].value;
	}
	var r0 = Math.round(13 * screen.deviceYDPI / 96);
	Addons.TabPlus.ImgLock = MakeImgSrc(item.getAttribute("IconLock") || "bitmap:ieframe.dll,545,13,2", 0, true, r0);
	var s = item.getAttribute("IconClose");
	if (s) {
		Addons.TabPlus.ImgClose = '<img draggable="false" src="' + MakeImgSrc(s, 0, true, r0) + '" style="width: ';
		Addons.TabPlus.ImgClose2 = '';
	} else {
		Addons.TabPlus.ImgClose = '<span style="font-family: marlett; font-size: ';
		Addons.TabPlus.ImgClose2 = '&#x72;</span>';
	}
	Addons.TabPlus.ImgProtect = MakeImgSrc(item.getAttribute("IconProtect") || "icon:imageres.dll,1", 0, true, r0);
} else {
	var ado = OpenAdodbFromTextFile("addons\\" + Addon_Id + "\\options.html");
	if (ado) {
		var Icon = document.F.Icon;
		if (Icon) {
			Icon.name = "Icon_0";
		}
		SetTabContents(0, "General", ado.ReadText(adReadAll));
		ado.Close();
		document.getElementById("_Drive").innerHTML = api.LoadString(hShell32, 4122).replace(/ %c:?/, "");
	}
}
