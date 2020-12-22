//Tablacus Explorer

te.ClearEvents();
te.About = AboutTE(2);
Init = false;
g_arBM = [];

GetAddress = null;
ShowContextMenu = null;

Addon_Id = "";
g_pidlCP = api.ILRemoveLastID(ssfCONTROLS);
if (api.ILIsEmpty(g_pidlCP) || api.ILIsEqual(g_pidlCP, ssfDRIVES)) {
	g_pidlCP = ssfCONTROLS;
}

Refresh = function (Ctrl, pt) {
	return RunEvent4("Refresh", Ctrl, pt);
}

InputDialog = function (text, defaultText, cb, data) {
	const r = RunEvent4("InputDialog", text, defaultText);
	if (r !== void 0) {
		if (cb) {
			cb(r, data);
			return;
		}
		return r;
	}
	if (cb) {
		const opt = api.CreateObject("Object");
		opt.text = text;
		opt.defaultText = defaultText;
		opt.callback = function (text) {
			setTimeout(cb, 9, text, data);
		};
		ShowDialogEx("input", 480, 140, null, opt);
		return;
	}
	if (window.prompt) {
		return prompt(GetTextR(text), defaultText);
	}
	const rc = api.Memory("RECT");
	api.GetWindowRect(te.hwnd, rc);
	const t = 1440 / screen.deviceYDPI;
	const x = Math.min((rc.left + (rc.right - rc.left) / 2 - 186) * t, 32767);
	const y = Math.min((rc.top + (rc.bottom - rc.top) / 2 - 74) * t, 32767);
	return api.GetScriptDispatch('Function InputDialog(text, TITLE, defaultText, x, y)\nInputDialog = InputBox(text, TITLE, defaultText, x, y)\nEnd Function', "VBScript").InputDialog(GetTextR(text), TITLE, defaultText, x, y);
}

g_.mouse = {
	str: "",
	CancelContextMenu: false,
	ptGesture: api.Memory("POINT"),
	hwndGesture: null,
	tidGesture: null,
	bCapture: false,
	RButton: -1,
	bTrail: false,
	bDblClk: false,

	StartGestureTimer: function () {
		InvokeUI("StartGestureTimer");
	},

	EndGesture: function (button) {
		clearTimeout(this.tidGesture);
		if (this.bCapture) {
			api.ReleaseCapture();
			this.bCapture = false;
		}
		if (this.RButton >= 0) {
			this.RButtonDown(false)
		}
		this.str = "";
		g_bRButton = false;
		SetGestureText(Ctrl, "");
		if (this.bTrail) {
			api.RedrawWindow(te.hwnd, null, 0, RDW_INVALIDATE | RDW_FRAME | RDW_ALLCHILDREN);
			this.bTrail = false;
			this.ptText = null;
		}
	},

	RButtonDown: function (mode) {
		if (this.str == "2") {
			var item = api.Memory("LVITEM");
			item.iItem = this.RButton;
			item.mask = LVIF_STATE;
			item.stateMask = LVIS_SELECTED;
			api.SendMessage(this.hwndGesture, LVM_GETITEM, 0, item);
			if (!(item.state & LVIS_SELECTED)) {
				if (mode) {
					var Ctrl = te.CtrlFromWindow(this.hwndGesture);
					Ctrl.SelectItem(this.RButton, SVSI_SELECT | SVSI_FOCUSED | SVSI_DESELECTOTHERS);
				} else {
					var ptc = api.Memory("POINT");
					ptc = te.Data.pt.Clone();
					api.ScreenToClient(this.hwndGesture, ptc);
					api.SendMessage(this.hwndGesture, WM_RBUTTONDOWN, 0, ptc.x + (ptc.y << 16));
				}
			}
		}
		this.RButton = -1;
	},

	GetButton: function (msg, wParam) {
		var s = "";
		if (msg >= WM_LBUTTONDOWN && msg <= WM_LBUTTONDBLCLK) {
			if (api.GetKeyState(VK_RBUTTON) < 0) {
				g_.mouse.CancelContextMenu = true;
			}
			s = "1";
		}
		if (msg >= WM_RBUTTONDOWN && msg <= WM_RBUTTONDBLCLK) {
			s = "2";
		}
		if (msg >= WM_MBUTTONDOWN && msg <= WM_MBUTTONDBLCLK) {
			if (api.GetKeyState(VK_RBUTTON) < 0) {
				g_.mouse.CancelContextMenu = true;
			}
			s = "3";
		}
		if (msg >= WM_XBUTTONDOWN && msg <= WM_XBUTTONDBLCLK) {
			if (api.GetKeyState(VK_RBUTTON) < 0) {
				g_.mouse.CancelContextMenu = true;
			}
			switch (wParam >> 16) {
				case XBUTTON1:
					s = "4";
					break;
				case XBUTTON2:
					s = "5";
					break;
			}
		}
		return this.str.length ? s : GetGestureButton().replace(s, "") + s;
	},

	GetMode: function (Ctrl, pt) {
		switch (Ctrl ? Ctrl.Type : 0) {
			case CTRL_SB:
			case CTRL_EB:
				return Ctrl.ItemCount && Ctrl.ItemCount(SVGIO_SELECTION) ? "List" : "List_Background";
			case CTRL_TV:
				return "Tree";
			case CTRL_TC:
				return Ctrl.HitTest(pt, TCHT_ONITEM) >= 0 ? "Tabs" : "Tabs_Background";
			case CTRL_WB:
				return "Browser";
		}
	},

	Exec: function (Ctrl, hwnd, pt, str) {
		var str = GetGestureKey() + (str || this.str);
		this.EndGesture(false);
		te.Data.cmdMouse = str;
		if (!Ctrl) {
			return S_FALSE;
		}
		var s = this.GetMode(Ctrl, pt);
		if (s) {
			if (GestureExec(Ctrl, s, str, pt, hwnd) === S_OK) {
				return S_OK;
			}
			var res = /(.*)_Background/i.exec(s);
			if (res) {
				if (GestureExec(Ctrl, res[1], str, pt, hwnd) === S_OK) {
					return S_OK;
				}
			}
		} else {
			if (str.length) {
				window.g_menu_button = str;
				if (window.g_menu_click) {
					if (window.g_menu_click === true) {
						var hSubMenu = api.GetSubMenu(g_.menu_handle, g_.menu_pos);
						if (hSubMenu) {
							var mii = api.Memory("MENUITEMINFO");
							mii.cbSize = mii.Size;
							mii.fMask = MIIM_SUBMENU;
							if (api.SetMenuItemInfo(g_.menu_handle, g_.menu_pos, true, mii)) {
								api.DestroyMenu(hSubMenu);
							}
						}
					}
					if (str > 2) {
						window.g_menu_click = 4;
						var lParam = pt.x + (pt.y << 16);
						api.PostMessage(hwnd, WM_LBUTTONDOWN, 0, lParam);
						api.PostMessage(hwnd, WM_LBUTTONUP, 0, lParam);
					}
				}
			}
		}
		if (Ctrl.Type != CTRL_TE) {
			if (GestureExec(Ctrl, "All", str, pt, hwnd) === S_OK) {
				return S_OK;
			}
		}
		return S_FALSE;
	}
};

g_basic = {
	FuncI: function (s) {
		s = s.replace(/&|\.\.\.$/g, "");
		return this.Func[s] || api.ObjGetI(this.Func, s);
	},

	CmdI: function (s, s2) {
		var type = this.FuncI(s);
		if (type) {
			return type.Cmd[s2] || api.ObjGetI(type.Cmd, s2);
		}
	},

	Func: {
		"": {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				var lines = s.split(/\r?\n/);
				for (var i in lines) {
					var cmd = lines[i].split(",");
					var Id = cmd.shift();
					var hr = Exec(Ctrl, cmd.join(","), Id, hwnd, pt);
					if (hr != S_OK) {
						break;
					}
				}
				return S_OK;
			},

			Ref: function (s, pt) {
				var lines = s.split(/\r?\n/);
				var last = lines.length ? lines[lines.length - 1] : "";
				var res = /^([^,]+),$/.exec(last);
				if (res) {
					var Id = GetSourceText(res[1]);
					var r = OptionRef(Id, "", pt);
					if ("string" === typeof r) {
						return s + r + "\n";
					}
					return r;
				} else {
					var arFunc = [];
					RunEvent1("AddType", arFunc);
					var r = g_basic.Popup(arFunc, s, pt);
					return r == 1 ? 1 : s + (s.length && !/\n$/.test(s) ? "\n" : "") + r + ",";
				}
			}
		},

		Open: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				return ExecOpen(Ctrl, s, type, hwnd, pt, GetNavigateFlags(GetFolderView(Ctrl, pt)));
			},

			Drop: DropOpen,
			Ref: BrowseForFolder
		},

		"Open in new tab": {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				return ExecOpen(Ctrl, s, type, hwnd, pt, SBSP_NEWBROWSER);
			},

			Drop: DropOpen,
			Ref: BrowseForFolder
		},

		"Open in background": {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				return ExecOpen(Ctrl, s, type, hwnd, pt, SBSP_NEWBROWSER | SBSP_ACTIVATE_NOFOCUS);
			},

			Drop: DropOpen,
			Ref: BrowseForFolder
		},

		Filter: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				var FV = GetFolderView(Ctrl, pt);
				if (FV) {
					var s = ExtractMacro(Ctrl, s);
					FV.FilterView = s != "*" ? s : null;
					FV.Refresh();
				}
				return S_OK;
			}
		},

		Exec: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				var lines = ExtractMacro(Ctrl, s).split(/\r?\n/);
				for (var i in lines) {
					try {
						ShellExecute(lines[i], null, SW_SHOWNORMAL, Ctrl, pt);
					} catch (e) {
						ShowError(e, lines[i]);
					}
				}
				return S_OK;
			},

			Drop: function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop) {
				if (!pdwEffect) {
					pdwEffect = api.Memory("DWORD");
				}
				var re = /%Selected%/i;
				if (re.test(s)) {
					pdwEffect[0] = DROPEFFECT_LINK;
					if (bDrop) {
						var ar = [];
						for (var i = dataObj.Count; i > 0; ar.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(dataObj.Item(--i), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL)))) {
						}
						s = s.replace(re, ar.join(" "));
						ShellExecute(s, null, SW_SHOWNORMAL, Ctrl, pt);
					}
					return S_OK;
				} else {
					pdwEffect[0] = DROPEFFECT_NONE;
					return S_OK;
				}
			},

			Ref: OpenDialog
		},

		RunAs: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				s = ExtractMacro(Ctrl, s);
				try {
					ShellExecute(s, "RunAs", SW_SHOWNORMAL, Ctrl, pt);
				} catch (e) {
					ShowError(e, s);
				}
				return S_OK;
			},

			Ref: OpenDialog
		},

		JScript: {
			Exec: ExecScriptEx,
			Drop: DropScript,
			Ref: OpenDialog
		},

		JavaScript: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				if (/^"?[A-Z]:\\.*\.js"?\s*$|^"?\\\\\w.*\.js"?s*$/im.test(s)) {
					s = ReadTextFile(s);
				}
				InvokeUI("ExecJavaScript", [s, Ctrl, GetFolderView(Ctrl, pt), type, hwnd, pt]);
				return S_OK;
			},
			Ref: OpenDialog
		},

		VBScript: {
			Exec: ExecScriptEx,
			Drop: DropScript,
			Ref: OpenDialog
		},

		"Selected items": {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				var fn = g_basic.CmdI(type, s);
				if (fn) {
					return fn(Ctrl, pt);
				} else {
					return g_basic.Func.Exec.Exec(Ctrl, s + " %Selected%", "Exec", hwnd, pt);
				}
			},

			Drop: function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop) {
				var fn = g_basic.CmdI(type, s);
				if (fn) {
					pdwEffect[0] = DROPEFFECT_NONE;
					return S_OK;
				}
				return g_basic.Func.Exec.Drop(Ctrl, s + " %Selected%", "Exec", hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
			},

			Ref: function (s, pt) {
				var r = g_basic.Popup(g_basic.Func["Selected items"].Cmd, s, pt);
				if (SameText(r, GetText("Send to..."))) {
					var Folder = sha.NameSpace(ssfSENDTO);
					if (Folder) {
						var Items = Folder.Items();
						var hMenu = api.CreatePopupMenu();
						for (i = 0; i < Items.Count; ++i) {
							api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, i + 1, Items.Item(i).Name);
						}
						var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD | (pt.width ? TPM_RIGHTALIGN : 0), pt.x + pt.width, pt.y, te.hwnd, null, null);
						api.DestroyMenu(hMenu);
						if (nVerb) {
							return api.PathQuoteSpaces(Items.Item(nVerb - 1).Path);
						}
					}
					return 1;
				}
				if (SameText(r, GetText("Open with..."))) {
					r = OpenDialog(s);
					if (!r) {
						r = 1;
					}
				}
				return r;
			},

			Cmd: {
				Open: function (Ctrl, pt) {
					return OpenSelected(Ctrl, GetNavigateFlags(GetFolderView(Ctrl, pt)), pt);
				},
				"Open in new tab": function (Ctrl, pt) {
					return OpenSelected(Ctrl, SBSP_NEWBROWSER, pt);
				},
				"Open in background": function (Ctrl, pt) {
					return OpenSelected(Ctrl, SBSP_NEWBROWSER | SBSP_ACTIVATE_NOFOCUS, pt);
				},
				Exec: function (Ctrl, pt) {
					var Selected = GetSelectedArray(Ctrl, pt, true).shift();
					if (Selected) {
						InvokeCommand(Selected, 0, te.hwnd, null, null, null, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
					}
					return S_OK;
				},
				"Open with...": void 0,
				"Send to...": void 0
			},
			Enc: true
		},

		Tabs: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				var fn = g_basic.CmdI(type, s);
				if (fn) {
					fn(Ctrl, pt);
					return S_OK;
				}
				var FV = GetFolderView(Ctrl, pt, true);
				if (FV) {
					var res = /^(\-?)(\d+)/.exec(s);
					if (res) {
						var TC = FV.Parent;
						TC.SelectedIndex = res[1] ? TC.Count - res[2] : res[2];
					}
				}
				return S_OK;
			},

			Cmd: {
				"Close tab": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt, true);
					FV && FV.Close();
				},
				"Close other tabs": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						var TC = FV.Parent;
						var nIndex = GetFolderView(Ctrl, pt, true) ? FV.Index : -1;
						TC.LockUpdate();
						try {
							for (var i = TC.Count; i--;) {
								if (i != nIndex) {
									TC[i].Close();
								}
							}
						} catch (e) {
							ShowError(e);
						}
						TC.UnlockUpdate();
					}
				},
				"Close tabs on left": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt, true);
					if (FV) {
						var TC = FV.Parent;
						TC.LockUpdate();
						try {
							for (var i = FV.Index; i--;) {
								TC[i].Close();
							}
						} catch (e) {
							ShowError(e);
						}
						TC.UnlockUpdate();
					}
				},
				"Close tabs on right": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt, true);
					if (FV) {
						var TC = FV.Parent;
						var nIndex = FV.Index;
						TC.LockUpdate();
						try {
							for (var i = TC.Count; --i > nIndex;) {
								TC[i].Close();
							}
						} catch (e) {
							ShowError(e);
						}
						TC.UnlockUpdate();
					}
				},
				"Close all tabs": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						var TC = FV.Parent;
						TC.LockUpdate();
						try {
							for (var i = TC.Count; i--;) {
								TC[i].Close();
							}
						} catch (e) {
							ShowError(e);
						}
						TC.UnlockUpdate();
					}
				},
				"New tab": CreateTab,
				Lock: function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					FV && Lock(FV.Parent, FV.Index, true);
				},
				"Previous tab": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					FV && ChangeTab(FV.Parent, -1);
				},
				"Next tab": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					FV && ChangeTab(FV.Parent, 1);
				},
				Up: function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					FV && FV.Navigate(null, SBSP_PARENT | GetNavigateFlags(FV, true));
				},
				Back: function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						var wFlags = GetNavigateFlags(FV);
						if (wFlags & SBSP_NEWBROWSER) {
							var Log = FV.History;
							if (Log && Log.Index < Log.Count - 1) {
								FV.Navigate(Log[Log.Index + 1], wFlags);
							}
						} else {
							FV.Navigate(null, SBSP_NAVIGATEBACK);
						}
					}
				},
				Forward: function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						var wFlags = GetNavigateFlags(FV);
						if (wFlags & SBSP_NEWBROWSER) {
							var Log = FV.History;
							if (Log && Log.Index > 0) {
								FV.Navigate(Log[Log.Index - 1], wFlags);
							}
						} else {
							FV && FV.Navigate(null, SBSP_NAVIGATEFORWARD);
						}
					}
				},
				Refresh: Refresh,

				"Show frames": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						FV.Type = (FV.Type == CTRL_SB) ? CTRL_EB : CTRL_SB;
					}
				},
				"Switch Explorer engine": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						FV.Type = (FV.Type == CTRL_SB) ? CTRL_EB : CTRL_SB;
					}
				},
				"Open in Explorer": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					FV && OpenInExplorer(FV);
				}
			},

			Enc: true
		},

		Edit: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				var hMenu = te.MainMenu(FCIDM_MENU_EDIT);
				var nVerb = GetCommandId(hMenu, s);
				SendCommand(Ctrl, nVerb);
				api.DestroyMenu(hMenu);
				return S_OK;
			},

			Ref: function (s, pt) {
				return g_basic.PopupMenu(te.MainMenu(FCIDM_MENU_EDIT), null, pt);
			}
		},

		View: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				var hMenu = te.MainMenu(FCIDM_MENU_VIEW);
				var nVerb = GetCommandId(hMenu, s);
				SendCommand(Ctrl, nVerb);
				api.DestroyMenu(hMenu);
				return S_OK;
			},

			Ref: function (s, pt) {
				return g_basic.PopupMenu(te.MainMenu(FCIDM_MENU_VIEW), null, pt);
			}
		},

		Context: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				var Selected;
				var FV = Ctrl;
				if (Ctrl.Type <= CTRL_EB || Ctrl.Type == CTRL_TV) {
					Selected = Ctrl.SelectedItems();
				} else {
					FV = te.Ctrl(CTRL_FV);
					Selected = FV.SelectedItems();
				}
				if (Selected && Selected.Count) {
					var ContextMenu = api.ContextMenu(Selected, FV);
					if (ContextMenu) {
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS | CMF_CANRENAME);
						var nVerb = GetCommandId(hMenu, s, ContextMenu);
						FV.Focus();
						ContextMenu.InvokeCommand(0, te.hwnd, nVerb ? nVerb - 1 : s, null, null, SW_SHOWNORMAL, 0, 0);
						api.DestroyMenu(hMenu);
					}
				}
				return S_OK;
			},

			Ref: function (s, pt) {
				var FV = te.Ctrl(CTRL_FV);
				if (FV) {
					var Selected = FV.SelectedItems();
					if (!Selected.Count) {
						Selected = api.GetModuleFileName(null);
					}
					var ContextMenu = api.ContextMenu(Selected, FV);
					if (ContextMenu) {
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS | CMF_CANRENAME);
						return g_basic.PopupMenu(hMenu, ContextMenu, pt);
					}
				}
			}
		},

		Background: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				var FV = te.Ctrl(CTRL_FV);
				if (FV) {
					var ContextMenu = FV.ViewMenu();
					if (ContextMenu) {
						FV.Focus();
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS);
						var nVerb = GetCommandId(hMenu, s, ContextMenu);
						ContextMenu.InvokeCommand(0, te.hwnd, nVerb ? nVerb - 1 : s, null, null, SW_SHOWNORMAL, 0, 0);
						api.DestroyMenu(hMenu);
					}
				}
				return S_OK;
			},

			Ref: function (s, pt) {
				var FV = te.Ctrl(CTRL_FV);
				if (FV) {
					var ContextMenu = FV.ViewMenu();
					if (ContextMenu) {
						var hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS);
						return g_basic.PopupMenu(hMenu, ContextMenu, pt);
					}
				}
			}
		},

		Tools: {
			Cmd: {
				"New folder": CreateNewFolder,
				"New file": CreateNewFile,
				"Copy full path": function (Ctrl, pt) {
					var Selected = GetSelectedItems(Ctrl, pt);
					var s = "";
					var nCount = Selected.Count;
					if (nCount) {
						s = te.OnClipboardText(Selected);
					} else {
						s = api.PathQuoteSpaces(api.GetDisplayNameOf(FV, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL));
					}
					api.SetClipboardData(s);
					return S_OK;
				},
				"Run dialog": function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					api.ShRunDialog(te.hwnd, 0, FV ? FV.FolderItem.Path : null, null, null, 0);
					return S_OK;
				},
				Search: function (Ctrl, pt) {
					var FV = GetFolderView(Ctrl, pt);
					if (FV) {
						var s = InputDialog("Search", IsSearchPath(FV) ? GetFolderItemName(FV.FolderItem) : "");
						if (s) {
							FV.Search(s);
						} else if (s === "") {
							CancelFilterView(FV);
						}
					}
					return S_OK;
				},
				"Add to favorites": AddFavoriteEx,
				"Reload customize": ReloadCustomize,
				"Load layout": LoadLayout,
				"Save layout": SaveLayout,
				"Close application": function () {
					api.PostMessage(te.hwnd, WM_CLOSE, 0, 0);
					return S_OK;
				}
			},

			Enc: true
		},

		Options: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				ShowOptions("Tab=" + s);
				return S_OK;
			},

			Ref: function (s, pt) {
				return g_basic.Popup(g_basic.Func.Options.List, s, pt);
			},

			List: ["General", "Add-ons", "Menus", "Tabs", "Tree", "List"],

			Enc: true
		},

		Key: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				api.SetFocus(Ctrl.hwnd);
				wsh.SendKeys(s);
				return S_OK;
			}
		},

		"Load layout": {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				if (s) {
					LoadXml(s);
				} else {
					LoadLayout();
				}
				return S_OK;
			},

			Ref: OpenDialog
		},

		"Save layout": {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				if (s) {
					SaveXml(s);
				} else {
					SaveLayout();
				}
				return S_OK;
			},

			Ref: OpenDialog
		},

		"Add-ons": {
			Cmd: {},
			Enc: true
		},

		Menus: {
			Ref: function (s, pt) {
				return g_basic.Popup(g_basic.Func.Menus.List, s, pt);
			},

			List: ["Open", "Close", "Separator", "Break", "BarBreak"],

			Enc: true
		}
	},

	Exec: function (Ctrl, s, type, hwnd, pt) {
		var fn = g_basic.CmdI(type, s);
		if (!pt) {
			pt = api.Memory("POINT");
			api.GetCursorPos(pt);
		}
		fn && fn(Ctrl, pt);
		return S_OK;
	},

	Popup: function (Cmd, strDefault, pt) {
		var i, j, s;
		var ar = [];
		if (Cmd.length) {
			ar = Cmd;
		} else {
			for (i in Cmd) {
				ar.push(i);
			}
		}
		var hMenu = api.CreatePopupMenu();
		for (i = 0; i < ar.length; ++i) {
			if (ar[i]) {
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, i + 1, GetTextR(ar[i]));
			}
		}
		var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD | (pt.width ? TPM_RIGHTALIGN : 0), pt.x + pt.width, pt.y, te.hwnd, null, null);
		s = api.GetMenuString(hMenu, nVerb, MF_BYCOMMAND);
		api.DestroyMenu(hMenu);
		if (nVerb == 0) {
			return 1;
		}
		return s;
	},

	PopupMenu: function (hMenu, ContextMenu, pt) {
		var Verb;
		for (var i = api.GetMenuItemCount(hMenu); i--;) {
			if (api.GetMenuString(hMenu, i, MF_BYPOSITION)) {
				api.EnableMenuItem(hMenu, i, MF_ENABLED | MF_BYPOSITION);
			} else {
				api.DeleteMenu(hMenu, i, MF_BYPOSITION);
			}
		}
		window.g_menu_click = true;
		var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD | (pt.width ? TPM_RIGHTALIGN : 0), pt.x + pt.width, pt.y, te.hwnd, null, ContextMenu);
		if (nVerb == 0) {
			api.DestroyMenu(hMenu);
			return 1;
		}
		if (ContextMenu) {
			Verb = ContextMenu.GetCommandString(nVerb - 1, GCS_VERB);
		}
		if (!Verb) {
			Verb = window.g_menu_string;
			var res = /\t(.*)$/.exec(Verb);
			if (res) {
				Verb = res[1];
			}
		}
		api.DestroyMenu(hMenu);
		return Verb;
	}
};

RefreshEx = function (FV, tm, df) {
	if (FV.Data && FV.FolderItem && /^[A-Z]:\\|^\\\\\w/i.test(FV.FolderItem.Path)) {
		if (RunEvent4("RefreshEx", FV, tm, df) === void 0) {
			if (new Date().getTime() - (FV.Data.AccessTime || 0) > (df || 5000) || FV.Data.pathChk != FV.FolderItem.Path) {
				if (!FV.hwndView || FV.FolderItem.Unavailable || api.ILIsEqual(FV.FolderItem, FV.FolderItem.Alt)) {
					FV.Data.AccessTime = "!";
					FV.Data.pathChk = FV.FolderItem.Path;
					api.PathIsDirectory(function (hr, FV, Path) {
						if (FV.Data) {
							FV.Data.AccessTime = new Date().getTime();
							FV.Data.pathChk = void 0;
							if (Path == FV.FolderItem.Path) {
								if (hr < 0) {
									if (RunEvent4("Error", FV) === void 0) {
										if (FV.Unavailable > 3000) {
											FV.Suspend(2);
										}
									}
								} else if (hr == S_OK && FV.FolderItem.Unavailable) {
									FV.Refresh();
								}
							}
						}
					}, tm || -1, FV.FolderItem.Path, FV, FV.FolderItem.Path);
				}
			}
		}
	}
}

ChangeView = function (Ctrl) {
	const TC = te.Ctrl(CTRL_TC);
	if (TC && !g_.LockUpdate && Ctrl && Ctrl.FolderItem && Ctrl.FolderItem.Path != null) {
		ChangeTabName(Ctrl);
		if (Ctrl.hwndView) {
			RefreshEx(Ctrl, 5000, 5000);
		}
		RunEvent1("ChangeView", Ctrl);
		if (Ctrl.Id == Ctrl.Parent.Selected.Id) {
			if (Ctrl.Parent.Id == TC.Id) {
				RunEvent1("ChangeView1", Ctrl);
			}
			RunEvent1("ChangeView2", Ctrl);
		}
		RunEvent1("ConfigChanged", "Window");
	}
}

SetAddress = function (s) {
	RunEvent1("SetAddress", s);
}

ChangeTabName = function (Ctrl) {
	if (Ctrl.FolderItem) {
		Ctrl.Title = GetTabName(Ctrl);
	}
}

GetTabName = function (Ctrl) {
	if (Ctrl.FolderItem) {
		var res = /search\-ms:.*?crumb=([^&]+)/.exec(Ctrl.FilterView);
		if (res) {
			return decodeURIComponent(res[1]).replace(/~<(\*?)/, "$1");
		}
		return RunEvent4("GetTabName", Ctrl) || GetFolderItemName(Ctrl.FolderItem);
	}
}

GetFolderItemName = function (pid) {
	return pid ? RunEvent4("GetFolderItemName", pid) || api.GetDisplayNameOf(pid, SHGDN_INFOLDER | SHGDN_ORIGINAL) : "";
}

AddEvent("GetFolderItemName", function (pid) {
	var res = /search\-ms:.*?crumb=([^&]+)/.exec(pid.Path);
	if (res) {
		return decodeURIComponent(res[1]).replace(/~<(\*?)/, "$1");
	}
});

IsUseExplorer = function (pid) {
	return RunEvent3("UseExplorer", pid);
}

CloseView = function (Ctrl) {
	return RunEvent2("CloseView", Ctrl);
}

DeviceChanged = function (Ctrl) {
	RunEvent1("DeviceChanged", Ctrl);
}

ListViewCreated = function (Ctrl) {
	ChangeTabName(Ctrl);
	Ctrl.Data.AccessTime = "#";
	const res = /search\-ms:.*?crumb=([^&]+)/.exec(Ctrl.FilterView);
	if (res) {
		Ctrl.FilterView = null;
		Ctrl.Search(decodeURIComponent(res[1]).replace(/~<(\*?)/, "$1"));
		return S_OK;
	}
	RunEvent1("ListViewCreated", Ctrl);
}

TabViewCreated = function (Ctrl) {
	RunEvent1("TabViewCreated", Ctrl);
}

TreeViewCreated = function (Ctrl) {
	RunEvent1("TreeViewCreated", Ctrl);
}

SetAddrss = function (s) {
	RunEvent1("SetAddrss", s);
}

RestoreFromTray = function () {
	api.ShowWindow(te.hwnd, api.IsIconic(te.hwnd) ? SW_RESTORE : SW_SHOW);
	api.SetForegroundWindow(te.hwnd);
	setTimeout(function () {
		api.RedrawWindow(te.hwnd, null, 0, RDW_INVALIDATE | RDW_FRAME | RDW_ALLCHILDREN);
	}, 99);
	RunEvent1("RestoreFromTray");
}

Finalize = function () {
	Finalize = function () {};
	RunEvent1("Finalize");
	SaveConfig();
	Threads.Finalize();

	for (var i in g_.dlgs) {
		var dlg = g_.dlgs[i];
		try {
			if (dlg) {
				if (dlg.oExec) {
					if (dlg.oExec.Status == 0) {
						dlg.oExec.Terminate();
					}
				} else if (dlg.Document) {
					dlg.close();
				}
			}
		} catch (e) { }
	}
}

SetGestureText = function (Ctrl, Text) {
	var mode = g_.mouse.GetMode(Ctrl, g_.mouse.ptGesture);
	if (mode) {
		var s = eventTE.Mouse[mode][Text];
		if (!s) {
			var res = /(.*)_Background/i.exec(mode);
			if (res) {
				s = eventTE.Mouse[res[1]][Text];
			}
			if (!s) {
				s = eventTE.Mouse.All[Text];
			}
		}
		if (s && s[0][2]) {
			Text += "\n" + GetText(s[0][2]);
		}
	}
	RunEvent3("SetGestureText", Ctrl, Text);
	if (!te.Data.Conf_NoInfotip && Text.length > 1 && !/^[A-Z]+\d$/i.test(Text)) {
		g_.mouse.bTrail = true;
		var hdc = api.GetWindowDC(te.hwnd);
		if (hdc) {
			var rc = api.Memory("RECT");
			if (!g_.mouse.ptText) {
				g_.mouse.ptText = g_.mouse.ptGesture.Clone();
				api.ScreenToClient(te.hwnd, g_.mouse.ptText);
				g_.mouse.right = g_.mouse.bottom = -32767;
			}
			rc.left = g_.mouse.ptText.x;
			rc.top = g_.mouse.ptText.y;
			var hOld = api.SelectObject(hdc, CreateFont(DefaultFont));
			api.DrawText(hdc, Text, -1, rc, DT_CALCRECT | DT_NOPREFIX);
			if (g_.mouse.right < rc.right) {
				g_.mouse.right = rc.right;
			}
			if (g_.mouse.bottom < rc.bottom) {
				g_.mouse.bottom = rc.bottom;
			}
			api.Rectangle(hdc, rc.left - 2, rc.top - 1, g_.mouse.right + 2, g_.mouse.bottom + 1);
			api.DrawText(hdc, Text, -1, rc, DT_NOPREFIX);
			api.SelectObject(hdc, hOld);
			api.ReleaseDC(te.hwnd, hdc);
		}
	}
}

IsSavePath = function (path) {
	const en = "IsSavePath";
	const eo = eventTE[en.toLowerCase()];
	const args = [path];
	for (let i in eo) {
		try {
			if (api.Invoke(eo[i], args) === false) {
				return false;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	return true;
}

Lock = function (Ctrl, nIndex, turn) {
	const FV = Ctrl[nIndex];
	if (FV) {
		if (turn) {
			FV.Data.Lock = !FV.Data.Lock;
		}
		RunEvent1("Lock", Ctrl, nIndex, FV.Data.Lock);
	}
}

GetLock = function (FV) {
	return FV && FV.Data && FV.Data.Lock;
}

CanClose = function (FV) {
	if (FV && FV.Data) {
		if (GetLock(FV)) {
			return S_FALSE;
		}
		return RunEvent2("CanClose", FV);
	}
	return S_OK;
}

FontChanged = function () {
	RunEvent1("FontChanged");
}

FavoriteChanged = function () {
	InvokeUI("ExecJavaScript", ["delete ui_.MenuFavorites;"]);
	RunEvent1("FavoriteChanged");
}

LoadConfig = function (bDog) {
	var xml = OpenXml("window.xml", true, false);
	if (xml) {
		var items = xml.getElementsByTagName('Data');
		if (items.length && !bDog) {
			var path = items[0].getAttribute("Path");
			if (path) {
				path = ExtractMacro(te, path);
				if (fso.FolderExists(BuildPath(path, "config"))) {
					te.Data.DataFolder = path;
					LoadConfig(true);
					return;
				}
			}
		}
		var arKey = ["Conf", "Tab", "Tree", "View"];
		for (var j in arKey) {
			var key = arKey[j];
			var items = xml.getElementsByTagName(key);
			if (items.length == 0 && j == 0) {
				items = xml.getElementsByTagName('Config');
			}
			for (var i = 0; i < items.length; ++i) {
				var item = items[i];
				var s = item.text;
				if (s == "0") {
					s = 0;
				}
				te.Data[key + "_" + item.getAttribute("Id")] = s;
			}
		}
		te.Data.View_SizeFormat = void 0;
		g_.xmlWindow = xml;
		LoadWindowSettings(xml);
		return;
	}
	g_.xmlWindow = "Init";
	LoadWindowSettings();
}

LoadWindowSettings = function (xml) {
	if (fso.FileExists(te.Data.WindowSetting)) {
		xml = api.CreateObject("Msxml2.DOMDocument");
		xml.async = false;
		xml.load(te.Data.WindowSetting);
		if (xml) {
			g_.xmlWindow = xml;
		}
	} else {
		RunEvent1("ConfigChanged", "Config");
	}
	if (xml) {
		var items = xml.getElementsByTagName('Window');
		if (items.length) {
			var item = items[0];
			te.CmdShow = item.getAttribute("CmdShow");
			var x = GetNum(item.getAttribute("Left"));
			var y = GetNum(item.getAttribute("Top"));
			if (x > -30000 && y > -30000) {
				var w = GetNum(item.getAttribute("Width"));
				var h = GetNum(item.getAttribute("Height"));
				var z = GetNum(item.getAttribute("DPI")) / screen.deviceYDPI;
				if (z && z != 1) {
					x /= z;
					y /= z;
					w /= z;
					h /= z;
				}
				var pt = { x: x + w / 2, y: y };
				if (!api.MonitorFromPoint(pt, MONITOR_DEFAULTTONULL)) {
					var hMonitor = api.MonitorFromPoint(pt, MONITOR_DEFAULTTOPRIMARY);
					var mi = api.Memory("MONITORINFOEX");
					mi.cbSize = mi.Size;
					api.GetMonitorInfo(hMonitor, mi);
					x = mi.rcWork.left + (mi.rcWork.right - mi.rcWork.left - w) / 2;
					y = mi.rcWork.top + (mi.rcWork.bottom - mi.rcWork.top - h) / 2;
					if (y < mi.rcWork.top) {
						y = mi.rcWork.top;
					}
				}
				if (api.IsZoomed(te.hwnd)) {
					te.CmdShow = SW_SHOWMAXIMIZED;
				} else if (api.IsIconic(te.hwnd)) {
					te.CmdShow = SW_SHOWMINIMIZED;
				} else if (w && h) {
					api.MoveWindow(te.hwnd, x, y, w, h, true);
				}
				api.GetWindowRect(te.hwnd, te.Data.rcWindow);
			}
		}
	}
}

SaveConfig = function () {
	RunEvent1("SaveConfig");
	if (te.Data.bSaveMenus) {
		te.Data.bSaveMenus = false;
		SaveXmlEx("menus.xml", te.Data.xmlMenus);
	}
	if (te.Data.bSaveAddons) {
		te.Data.bSaveAddons = false;
		SaveXmlEx("addons.xml", te.Data.Addons);
	}
	if (te.Data.bSaveConfig) {
		te.Data.bSaveConfig = false;
		SaveConfigXML(BuildPath(te.Data.DataFolder, "config\\window.xml"));
	}
	if (te.Data.bSaveWindow) {
		te.Data.bSaveWindow = false;
		if (g_.Fullscreen) {
			while (g_.stack_TC.length) {
				g_.stack_TC.pop().Visible = true;
			}
			ExitFullscreen();
		}
		SaveXml(te.Data.WindowSetting);
	}
}

SaveAddons = function (Addons, bLoading) {
	if (bLoading) {
		var root = te.Data.Addons.documentElement;
		if (root) {
			var items = root.childNodes;
			if (items) {
				var nLen = items.length;
				for (var i = 0; i < nLen; ++i) {
					var item = items[i];
					var Id = item.nodeName;
					if (Addons[Id] == null) {
						Addons[Id] = item.getAttribute("Enabled");
					}
				}
			}
		}
	}
	te.Data.bErrorAddons = false;
	var xml = CreateXml();
	var root = xml.createElement("TablacusExplorer");
	for (var Id in Addons) {
		var item = null;
		var items = te.Data.Addons.getElementsByTagName(Id);
		if (items.length) {
			item = items[0].cloneNode(true);
		}
		if (!item) {
			item = xml.createElement(Id);
		}
		var Enabled = Addons[Id];
		if (Enabled) {
			var AddonFolder = BuildPath(te.Data.Installed, "addons", Id);
			Enabled = fso.FolderExists(AddonFolder + "\\lang") ? 2 : 0;
			if (fso.FileExists(AddonFolder + "\\script.vbs")) {
				Enabled |= 8;
			}
			if (fso.FileExists(AddonFolder + "\\script.js")) {
				Enabled |= 1;
			}
			Enabled = (Enabled & 9) ? Enabled : 4;
		} else {
			AddonDisabled(Id);
		}
		item.setAttribute("Enabled", Enabled);
		var info = await GetAddonInfo(Id);
		if (info.Level > 0) {
			item.setAttribute("Level", info.Level);
		}
		root.appendChild(item);
	}
	xml.appendChild(root);
	te.Data.Addons = xml;
	RunEvent1("ConfigChanged", "Addons");
}

AddEvent("Refresh", function (Ctrl, pt) {
	var FV = GetFolderView(Ctrl, pt);
	if (FV) {
		FV.Focus();
		FV.Refresh();
	}
});

AddFavorite = function (FolderItem) {
	var xml = te.Data.xmlMenus;
	var menus = xml.getElementsByTagName("Favorites");
	if (menus && menus.length) {
		var item = xml.createElement("Item");
		if (!FolderItem) {
			var FV = te.Ctrl(CTRL_FV);
			if (FV) {
				FolderItem = FV.FolderItem;
			}
		}
		if (!FolderItem) {
			return false;
		}
		var s = InputDialog("Add Favorite", GetFolderItemName(FolderItem));
		if (s) {
			item.setAttribute("Name", s.replace(/\\/g, "/"));
			item.setAttribute("Filter", "");
			var path = api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
			if ("string" === typeof path) {
				path = FolderItem.Path;
			}
			if (!FolderItem.Enum && fso.FileExists(path)) {
				path = api.PathQuoteSpaces(path);
				item.setAttribute("Type", "Exec");
			} else {
				item.setAttribute("Type", "Open");
			}
			item.text = path;
			menus[0].appendChild(item);
			SaveXmlEx("menus.xml", xml);
			FavoriteChanged();
			return true;
		}
	}
	return false;
}

CancelFilterView = function (FV) {
	if (IsSearchPath(FV)) {
		FV.Navigate(null, SBSP_PARENT);
		return S_OK;
	}
	if (FV.FilterView) {
		FV.FilterView = null;
		FV.Refresh();
		return S_OK;
	}
	return S_FALSE;
}

IsSearchPath = function (pid) {
	return /^search\-ms:.*?&crumb=location:([^&]*)/.exec("string" === typeof pid ? pid : api.GetDisplayNameOf(pid, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL));
}

GetCommandId = function (hMenu, s, ContextMenu) {
	var arMenu = [hMenu];
	var wId = 0;
	if (s) {
		var sl = GetTextR(s);
		var mii = api.Memory("MENUITEMINFO");
		mii.cbSize = mii.Size;
		mii.fMask = MIIM_SUBMENU | MIIM_ID | MIIM_FTYPE | MIIM_STATE;
		while (arMenu.length) {
			hMenu = arMenu.pop();
			if (ContextMenu) {
				ContextMenu.HandleMenuMsg(WM_INITMENUPOPUP, hMenu, 0);
			}
			var i = api.GetMenuItemCount(hMenu);
			if (i == 1 && ContextMenu && WINVER >= 0x600) {
				setTimeout(function () {
					api.EndMenu();
				}, 200);
				api.TrackPopupMenuEx(hMenu, TPM_RETURNCMD, 0, 0, te.hwnd, null, ContextMenu);
				i = api.GetMenuItemCount(hMenu);
			}
			while (i-- > 0) {
				if (api.GetMenuItemInfo(hMenu, i, true, mii) && !(mii.fType & MFT_SEPARATOR) && !(mii.fState & MFS_DISABLED)) {
					var title = api.GetMenuString(hMenu, i, MF_BYPOSITION);
					if (title) {
						var a = title.split(/\t/);
						if (api.PathMatchSpec(a[0], s) || api.PathMatchSpec(a[0], sl) || api.PathMatchSpec(a[1], s) || (ContextMenu && api.PathMatchSpec(ContextMenu.GetCommandString(mii.wID - 1, GCS_VERB), s))) {
							wId = mii.hSubMenu ? api.TrackPopupMenuEx(mii.hSubMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, te.hwnd, null, ContextMenu) : mii.wID;
							arMenu.length = 0;
							break;
						}
						if (mii.hSubMenu) {
							arMenu.push(mii.hSubMenu);
						}
					}
				}
			}
			if (ContextMenu) {
				ContextMenu.HandleMenuMsg(WM_UNINITMENUPOPUP, hMenu, 0);
			}
		}
	}
	return wId;
};

OpenSelected = function (Ctrl, NewTab, pt) {
	var ar = GetSelectedArray(Ctrl, pt, true);
	var Selected = ar.shift();
	var SelItem = ar.shift();
	var FV = ar.shift();
	if (Selected) {
		var Exec = [];
		for (var i in Selected) {
			var Item = Selected.Item(i);
			var bFolder = Item.IsFolder || (!Item.IsFileSystem && Item.IsBrowsable);
			if (!bFolder) {
				var path = Item.ExtendedProperty("linktarget");
				if (path) {
					bFolder = api.PathIsDirectory(path) === true;
				}
			}
			if (bFolder) {
				FV.Navigate(Item, NewTab);
				NewTab |= SBSP_NEWBROWSER;
			} else {
				Exec.push(Item);
			}
		}
		if (Exec.length) {
			if (Selected.Count != Exec.length) {
				Selected = api.CreateObject("FolderItems");
				for (var i = 0; i < Exec.length; ++i) {
					Selected.AddItem(Exec[i]);
				}
			}
			InvokeCommand(Selected, 0, te.hwnd, null, null, null, SW_SHOWNORMAL, 0, 0, FV, CMF_DEFAULTONLY);
		}
	}
	return S_OK;
}

SendCommand = function (Ctrl, nVerb) {
	if (nVerb) {
		var hwnd = Ctrl.hwndView;
		if (!hwnd) {
			var FV = te.Ctrl(CTRL_FV);
			if (FV) {
				hwnd = FV.hwndView;
			}
		}
		api.SendMessage(hwnd, WM_COMMAND, nVerb, 0);
	}
}

IncludeObject = function (FV, Item) {
	if (api.PathMatchSpec(Item.Name, FV.Data.Filter)) {
		return S_OK;
	}
	return S_FALSE;
}

EnableDragDrop = function () {
}

SetFolderViewData = function (FV, FVD) {
	if (FV && FVD) {
		var t = FV.Type;
		for (var i in FVD) {
			try {
				FV[i] = FVD[i];
			} catch (e) { }
		}
		if (t == FVD.Type) {
			FV.Refresh();
		}
	}
}

SetTreeViewData = function (FV, TVD) {
	var TV = FV.TreeView;
	if (TV && TVD) {
		TV.Align = TVD.Align;
		TV.Style = TVD.Style;
		TV.Width = TVD.Width;
		TV.SetRoot(TVD.Root, TVD.EnumFlags, TVD.RootStyle);
	}
}

FixIconSpacing = function (Ctrl)
{
	var hwnd = Ctrl.hwndList;
	if (hwnd) {
		if (api.SendMessage(hwnd, LVM_GETVIEW, 0, 0) == 0) {
			var s = Ctrl.IconSize;
			var cx = s * 96 / screen.deviceYDPI + (10 + (255 - s) / 10) * screen.deviceYDPI / 96;
			var cz = s < 96 ? (s - 96) / 5 : (s > 96 ? (96 - s) / 9 : 0);
			var cy = s * 96 / screen.deviceYDPI + (46 + cz) * screen.deviceYDPI / 96;
			api.SendMessage(hwnd, LVM_SETICONSPACING, 0, cx + (cy << 16));
			api.InvalidateRect(hwnd, null, true);
		}
	}
}

//Events

te.OnCreate = function (Ctrl) {
	if (!Ctrl.Data) {
		Ctrl.Data = api.CreateObject("Object");
	}
	RunEvent1("Create", Ctrl);
}

te.OnClose = function (Ctrl) {
	return RunEvent2("Close", Ctrl);
}

AddEvent("Close", function (Ctrl) {
	switch (Ctrl.Type) {
		case CTRL_TE:
			Finalize();
			if (api.GetThreadCount() && MessageBox("File is in operation.", TITLE, MB_ABORTRETRYIGNORE) != IDIGNORE) {
				return S_FALSE;
			}
			eventTE = { Environment: {} };
			api.SHChangeNotifyDeregister(te.Data.uRegisterId);
			DeleteTempFolder();
			break;
		case CTRL_WB:
			break;
		case CTRL_SB:
		case CTRL_EB:
			return CanClose(Ctrl) || api.ILIsEqual(Ctrl, "about:blank") && Ctrl.Parent.Count < 2 ? S_FALSE : CloseView(Ctrl);
		case CTRL_TC:
			SetDisplay("Panel_" + Ctrl.Id, "none");
			break;
	}
});

te.OnViewCreated = function (Ctrl) {
	switch (Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
			ListViewCreated(Ctrl);
			break;
		case CTRL_TC:
			TabViewCreated(Ctrl);
			break;
		case CTRL_TV:
			TreeViewCreated(Ctrl);
			break;
	}
	RunEvent1("ViewCreated", Ctrl);
}

te.OnBeforeNavigate = function (Ctrl, fs, wFlags, Prev) {
	if (g_.tid_rf[Ctrl.Id]) {
		UI.clearTimeout(g_.tid_rf[Ctrl.Id]);
		delete g_.tid_rf[Ctrl.Id];
	}
	if (Ctrl.Data) {
		Ctrl.Data.Setting = void 0;
	}
	var path = api.GetDisplayNameOf(Ctrl.FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL);
	var res = /javascript:(.*)/im.exec(path);
	if (res) {
		try {
			new Function(res[1])(Ctrl);
		} catch (e) {
			ShowError(e, res[1]);
		}
		return E_NOTIMPL;
	}
	var hr = RunEvent2("BeforeNavigate", Ctrl, fs, wFlags, Prev);
	if (hr == S_OK && IsUseExplorer(Ctrl.FolderItem)) {
		setTimeout(async function (Path) {
			OpenInExplorer(Path);
		}, 99, Ctrl.FolderItem);
		return E_ABORT;
	}
	if (GetLock(Ctrl) && (wFlags & SBSP_NEWBROWSER) == 0 && !api.ILIsEqual(Prev, "about:blank")) {
		hr = E_ACCESSDENIED;
	}
	return hr;
}

te.OnNavigateComplete = function (Ctrl) {
	if (g_.tid_rf[Ctrl.Id] || !Ctrl.FolderItem) {
		return S_OK;
	}
	Ctrl.Data.AccessTime = new Date().getTime();
	Ctrl.NavigateComplete();
	RunEvent1("NavigateComplete", Ctrl);
	ChangeView(Ctrl);
	FocusFV();
	if (g_.focused) {
		g_.focused.Focus();
		if (--g_.fTCs <= 0) {
			delete g_.focused;
		}
	}
	return S_OK;
}

te.OnKeyMessage = function (Ctrl, hwnd, msg, key, keydata) {
	if (!te.Data) {
		return S_OK;
	}
	var hr = RunEvent3("KeyMessage", Ctrl, hwnd, msg, key, keydata);
	if (isFinite(hr)) {
		return hr;
	}
	if (g_.mouse.str.length > 1) {
		SetGestureText(Ctrl, GetGestureKey() + g_.mouse.str);
	}
	if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN) {
		var nKey = (((keydata >> 16) & 0x17f) || (api.MapVirtualKey(key, 0) | ((key >= 33 && key <= 46 || key >= 91 && key <= 93 || key == 111 || key == 144) ? 256 : 0))) | GetKeyShift();
		if (nKey == 0x15d) {
			g_.mouse.CancelContextMenu = false;
		}
		te.Data.cmdKey = nKey;
		te.Data.cmdKeyF = true;
		switch (Ctrl.Type) {
			case CTRL_SB:
			case CTRL_EB:
				var strClass = api.GetClassName(hwnd);
				if (api.PathMatchSpec(strClass, WC_LISTVIEW + ";DirectUIHWND")) {
					if (KeyExecEx(Ctrl, "List", nKey, hwnd) === S_OK) {
						return S_OK;
					}
				}
				if (strClass == WC_EDIT) {
					if (KeyExecEx(Ctrl, "Edit", nKey, hwnd) === S_OK) {
						return S_OK;
					}
					if (key == VK_TAB && Ctrl.hwndList) {
						SelectNext(Ctrl);
					}
					return S_FALSE;
				}
				break;
			case CTRL_TV:
				var strClass = api.GetClassName(hwnd);
				if (api.PathMatchSpec(strClass, WC_TREEVIEW)) {
					if (KeyExecEx(Ctrl, "Tree", nKey, hwnd) === S_OK) {
						return S_OK;
					}
					if (key == VK_DELETE) {
						InvokeCommand(Ctrl.SelectedItem, 0, te.hwnd, CommandID_DELETE - 1, null, null, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
						return S_OK;
					}
				}
				if (strClass == WC_EDIT) {
					if (KeyExecEx(Ctrl, "Edit", nKey, hwnd) === S_OK) {
						return S_OK;
					}
					return S_FALSE;
				}
				break;
			case CTRL_WB:
				if (KeyExecEx(Ctrl, "Browser", nKey, hwnd) === S_OK) {
					return S_OK;
				}
				break;
			default:
				if (key == VK_RETURN && window.g_menu_click === true) {
					var hSubMenu = api.GetSubMenu(g_.menu_handle, g_.menu_pos);
					if (hSubMenu) {
						var mii = api.Memory("MENUITEMINFO");
						mii.cbSize = mii.Size;
						mii.fMask = MIIM_SUBMENU;
						api.SetMenuItemInfo(g_.menu_handle, g_.menu_pos, true, mii);
						api.DestroyMenu(hSubMenu);
						api.PostMessage(hwnd, WM_CHAR, VK_LBUTTON, 0);
					}
				}
				if (g_.menu_loop && key == VK_TAB) {
					var wParam = api.GetKeyState(VK_SHIFT) < 0 ? VK_UP : VK_DOWN;
					api.PostMessage(hwnd, WM_KEYDOWN, wParam, 0);
					api.PostMessage(hwnd, WM_KEYUP, wParam, 0);
				}
				window.g_menu_button = api.GetKeyState(VK_CONTROL) < 0 ? 3 : api.GetKeyState(VK_SHIFT) < 0 ? 2 : 1;
				if (FolderMenu.MenuLoop) {
					if (KeyExecEx(Ctrl, "Menus", nKey, hwnd) === S_OK) {
						return S_OK;
					}
				}
				break;
		}
		if (Ctrl.Type != CTRL_TE) {
			if (KeyExecEx(Ctrl, "All", nKey, hwnd) === S_OK) {
				return S_OK;
			}
		}
	}
	if (msg == WM_KEYUP || msg == WM_SYSKEYUP) {
		if (g_.menu_loop) {
			if ((g_.menu_state & 0x2001) == 0x2001 && api.GetKeyState(VK_CONTROL) >= 0) {
				g_.menu_state = 0;
				api.PostMessage(hwnd, WM_KEYDOWN, VK_RETURN, 0);
				api.PostMessage(hwnd, WM_KEYUP, VK_RETURN, 0);
			}
		}
	}
	return S_FALSE;
}

te.OnMouseMessage = function (Ctrl, hwnd, msg, wParam, pt) {
	if (!te.Data) {
		return S_OK;
	}
	if (g_.mouse.Capture) {
		if (msg == WM_LBUTTONUP || api.GetKeyState(VK_LBUTTON) >= 0) {
			api.ReleaseCapture();
			g_.mouse.Capture = 0;
			te.Data.bSaveConfig = true;
		}
		var pt2 = pt.Clone();
		api.ScreenToClient(te.hwnd, pt2);
		if (pt2.x < 1) {
			pt2.x = 1;
		}
		InvokeUI("MoveSplitter", [pt2.x, g_.mouse.Capture]);
		return S_OK;
	}
	if (msg != WM_MOUSEMOVE) {
		te.Data.cmdKeyF = false;
	}
	var hr = RunEvent3("MouseMessage", Ctrl, hwnd, msg, wParam, pt);
	if (isFinite(hr)) {
		return hr;
	}
	var strClass = api.GetClassName(hwnd);
	if (strClass == WC_EDIT) {
		return S_FALSE;
	}
	var bLV = Ctrl.Type <= CTRL_EB && api.PathMatchSpec(strClass, WC_LISTVIEW + ";DirectUIHWND");
	if (msg == WM_MOUSEWHEEL) {
		var Ctrl2 = te.CtrlFromPoint(pt);
		if (Ctrl2) {
			g_.mouse.str = GetGestureButton() + (wParam > 0 ? "8" : "9");
			if (api.GetKeyState(VK_RBUTTON) < 0) {
				g_.mouse.CancelContextMenu = true;
			}
			if (g_.mouse.Exec(Ctrl2, hwnd, pt) == S_OK) {
				return S_OK;
			}
			var hwnd2 = api.WindowFromPoint(pt);
			if (hwnd2 && hwnd != hwnd2) {
				api.SetFocus(hwnd2);
				api.SendMessage(hwnd2, msg, wParam & 0xffff0000, pt.x + (pt.y << 16));
				return S_OK;
			}
		}
	}
	if (msg == WM_LBUTTONUP || msg == WM_RBUTTONUP || msg == WM_MBUTTONUP || msg == WM_XBUTTONUP) {
		if (g_.mouse.str.length) {
			var hr = S_FALSE;
			var bButton = false;
			if (msg == WM_RBUTTONUP) {
				if (g_.mouse.RButton >= 0) {
					g_.mouse.RButtonDown(true);
					bButton = (g_.mouse.str == "2");
				} else if (bLV) {
					var iItem = Ctrl.HitTest(pt, LVHT_ONITEM);
					if (iItem < 0 && !IsDrag(pt, te.Data.pt)) {
						Ctrl.SelectItem(null, SVSI_DESELECTOTHERS);
					}
				}
			}
			if (g_.mouse.str.length >= 2 || /[45]/.test(g_.mouse.str) || (!IsDrag(pt, te.Data.pt) && strClass != WC_HEADER)) {
				if (msg != WM_RBUTTONUP || g_.mouse.str.length < 2) {
					hr = g_.mouse.Exec(te.CtrlFromWindow(g_.mouse.hwndGesture), g_.mouse.hwndGesture, pt);
					if (msg == WM_LBUTTONUP) {
						hr = S_FALSE;
					}
				} else {
					setTimeout(function (Ctrl, hwnd, pt, str) {
						g_.mouse.Exec(Ctrl, hwnd, pt, str);
					}, 99, te.CtrlFromWindow(g_.mouse.hwndGesture), g_.mouse.hwndGesture, pt, g_.mouse.str);
					hr = S_OK;
				}
			}
			g_.mouse.EndGesture(false);
			if (g_.mouse.bCapture) {
				api.ReleaseCapture();
				g_.mouse.bCapture = false;
			}
			if (Ctrl.Type == CTRL_WB) {
				FocusFV();
			}
			if (bButton) {
				api.PostMessage(g_.mouse.hwndGesture, WM_CONTEXTMENU, g_.mouse.hwndGesture, pt.x + (pt.y << 16));
				return S_OK;
			}
			if (hr === S_OK) {
				return S_OK;
			}
		}
	}
	if (msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN || msg == WM_MBUTTONDOWN || msg == WM_XBUTTONDOWN) {
		var s = g_.mouse.GetButton(msg, wParam);
		if (g_.mouse.str.indexOf(s) >= 0) {
			g_.mouse.str = "";
		}
		if (g_.mouse.str.length == 0) {
			te.Data.pt = pt.Clone();
			g_.mouse.ptGesture = pt.Clone();
			g_.mouse.hwndGesture = hwnd;
			g_.mouse.ptDown = pt.Clone();
		}
		g_.mouse.str += s;
		if (msg == WM_MBUTTONDOWN) {
			if (bLV && te.Data.Conf_WheelSelect && api.GetKeyState(VK_SHIFT) >= 0 && api.GetKeyState(VK_CONTROL) >= 0) {
				var iItem = Ctrl.HitTest(pt, LVHT_ONITEM);
				if (iItem >= 0) {
					Ctrl.SelectItem(iItem, SVSI_SELECT | SVSI_FOCUSED | SVSI_DESELECTOTHERS);
				} else {
					Ctrl.SelectItem(null, SVSI_DESELECTOTHERS);
				}
			}
		}
		g_.mouse.StartGestureTimer();
		SetGestureText(Ctrl, GetGestureKey() + g_.mouse.str);
		if (msg == WM_RBUTTONDOWN) {
			g_.mouse.CancelContextMenu = api.GetKeyState(VK_LBUTTON) < 0 || api.GetKeyState(VK_MBUTTON) < 0 || api.GetKeyState(VK_XBUTTON1) < 0 || api.GetKeyState(VK_XBUTTON2) < 0;
			if (te.Data.Conf_Gestures >= 2) {
				var iItem = -1;
				if (bLV) {
					iItem = Ctrl.HitTest(pt, LVHT_ONITEM);
					if (iItem < 0) {
						api.ScreenToClient(hwnd, pt);
						return pt.y < screen.deviceYDPI / 4 ? S_FALSE : S_OK;
					}
				}
				if (te.Data.Conf_Gestures == 3 && Ctrl.Type != CTRL_WB) {
					g_.mouse.RButton = iItem;
					g_.mouse.StartGestureTimer();
					return S_OK;
				}
			}
		}
		if (Ctrl.Type == CTRL_SB && api.SendMessage(hwnd, LVM_GETEDITCONTROL, 0, 0)) {
			var iItem = Ctrl.HitTest(pt, LVHT_ONITEM);
			g_.mouse.Select = iItem >= 0 ? Ctrl.Items().Item(iItem) : null;
			g_.mouse.Deselect = Ctrl;
		}
	}
	if (msg == WM_LBUTTONDBLCLK || msg == WM_RBUTTONDBLCLK || msg == WM_MBUTTONDBLCLK || msg == WM_XBUTTONDBLCLK) {
		if (!IsHeader(Ctrl, pt, hwnd, strClass)) {
			var tm = new Date().getTime();
			if (tm - (g_.mouse.tmDblClick || 0) > sha.GetSystemInformation("DoubleClickTime")) {
				g_.mouse.tmDblClick = tm;
				te.Data.pt = pt.Clone();
				g_.mouse.str = g_.mouse.GetButton(msg, wParam);
				g_.mouse.str += g_.mouse.str;
				if (g_.mouse.Exec(Ctrl, hwnd, pt) == S_OK) {
					return S_OK;
				}
			}
		}
	}

	if (msg == WM_MOUSEMOVE && !/[45]/.test(g_.mouse.str)) {
		if (api.GetKeyState(VK_ESCAPE) < 0) {
			g_.mouse.EndGesture(false);
		}
		if (FolderMenu.MenuLoop) {
			if (api.GetKeyState(VK_LBUTTON) < 0 || api.GetKeyState(VK_RBUTTON) < 0) {
				if (g_.ptMenuDrag && IsDrag(pt, g_.ptMenuDrag)) {
					var hSubMenu = api.GetSubMenu(g_.menu_handle, g_.menu_pos);
					if (hSubMenu) {
						var mii = api.Memory("MENUITEMINFO");
						mii.cbSize = mii.Size;
						mii.fMask = MIIM_SUBMENU;
						api.SetMenuItemInfo(g_.menu_handle, g_.menu_pos, true, mii);
						api.DestroyMenu(hSubMenu);
					}
					window.g_menu_click = 4;
					window.g_menu_button = 4;
					var lParam = pt.x + (pt.y << 16);
					api.PostMessage(hwnd, WM_LBUTTONDOWN, 0, lParam);
					api.PostMessage(hwnd, WM_LBUTTONUP, 0, lParam);
				}
			} else {
				g_.ptMenuDrag = pt.Clone();
			}
		}
		if (g_.mouse.str.length && (te.Data.Conf_Gestures > 1 && api.GetKeyState(VK_RBUTTON) < 0) || (te.Data.Conf_Gestures && (api.GetKeyState(VK_MBUTTON) < 0 || api.GetKeyState(VK_XBUTTON1) < 0 || api.GetKeyState(VK_XBUTTON2) < 0))) {
			if (g_.mouse.ptGesture.x == -1 && g_.mouse.ptGesture.y == -1) {
				g_.mouse.ptGesture = pt.Clone();
			}
			var x = (pt.x - g_.mouse.ptGesture.x);
			var y = (pt.y - g_.mouse.ptGesture.y);
			if (Math.abs(x) + Math.abs(y) >= 20) {
				if (te.Data.Conf_TrailSize) {
					var hdc = api.GetWindowDC(te.hwnd);
					if (hdc) {
						var rc = api.Memory("RECT");
						api.GetWindowRect(te.hwnd, rc);
						api.MoveToEx(hdc, g_.mouse.ptGesture.x - rc.left, g_.mouse.ptGesture.y - rc.top, null);
						var pen1 = api.CreatePen(PS_SOLID, te.Data.Conf_TrailSize, te.Data.Conf_TrailColor);
						var hOld = api.SelectObject(hdc, pen1);
						api.LineTo(hdc, pt.x - rc.left, pt.y - rc.top);
						api.SelectObject(hdc, hOld);
						api.DeleteObject(pen1);
						g_.mouse.bTrail = true;
						api.ReleaseDC(te.hwnd, hdc);
					}
				}
				g_.mouse.ptGesture = pt.Clone();
				var s = (Math.abs(x) >= Math.abs(y)) ? ((x < 0) ? "L" : "R") : ((y < 0) ? "U" : "D");

				if (s != g_.mouse.str.charAt(g_.mouse.str.length - 1)) {
					g_.mouse.str += s;
					if (api.GetKeyState(VK_RBUTTON) < 0) {
						g_.mouse.CancelContextMenu = true;
					}
					SetGestureText(Ctrl, GetGestureKey() + g_.mouse.str);
				}
				if (!g_.mouse.bCapture) {
					api.SetCapture(g_.mouse.hwndGesture);
					g_.mouse.bCapture = true;
				}
				g_.mouse.StartGestureTimer();
			}
		} else {
			g_.mouse.ptGesture.x = -1;
			g_.mouse.ptGesture.y = -1;
		}
		if (g_.mouse.Deselect) {
			g_.mouse.Deselect.SelectItem(g_.mouse.Select, (g_.mouse.Select ? SVSI_SELECT | SVSI_FOCUSED : 0) | (api.GetKeyState(VK_CONTROL) < 0 ? 0 : SVSI_DESELECTOTHERS));
			delete g_.mouse.Deselect;
			delete g_.mouse.Select
		}
	}
	return g_.mouse.str.length >= 2 ? S_OK : S_FALSE;
}

te.OnCommand = function (Ctrl, hwnd, msg, wParam, lParam) {
	if (Ctrl.Type <= CTRL_EB) {
		switch ((wParam & 0xfff) + 1) {
			case CommandID_DELETE:
				var Items = Ctrl.SelectedItems();
				for (var i = Items.Count; i--;) {
					UnlockFV(Items.Item(i));
				}
				break;
			case CommandID_CUT:
			case CommandID_COPY:
				if (Ctrl.hwndList) {
					api.InvalidateRect(Ctrl.hwndList, null, true);
				}
				break;
			case CommandID_PASTE:
				var cFV = te.Ctrls(CTRL_FV);
				for (var i in cFV) {
					var FV = cFV[i];
					if (FV.hwndList) {
						var item = api.Memory("LVITEM");
						item.stateMask = LVIS_CUT;
						api.SendMessage(FV.hwndList, LVM_SETITEMSTATE, -1, item);
					}
				}
				break;
		}
		if (wParam === 0x7103) {
			Refresh();
			return S_OK;
		}
	}
	var hr = RunEvent3("Command", Ctrl, hwnd, msg, wParam, lParam);
	if (!isFinite(hr) && Ctrl.Type <= CTRL_EB && (wParam & 0xfff) + 1 == CommandID_PROPERTIES) {
		hr = InvokeCommand(Ctrl.SelectedItems(), 0, te.hwnd, "properties", null, null, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
	}
	RunEvent1("ConfigChanged", "Window");
	return hr;
}

te.OnInvokeCommand = function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon) {
	var path;
	var strVerb = (isFinite(Verb) ? ContextMenu.GetCommandString(Verb, GCS_VERB) : Verb).toLowerCase();
	if (strVerb == "open" && ContextMenu.GetCommandString(Verb, GCS_HELPTEXT) != api.LoadString(hShell32, 12850)) {
		strVerb = "";
	}
	var hr = RunEvent3("InvokeCommand", ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon, strVerb);
	if (isNaN(hr) && strVerb == "cut") {
		var FV = ContextMenu.FolderView;
		if (FV && FV.hwndView && SameFolderItems(ContextMenu.Items(), FV.SelectedItems())) {
			var hMenu = te.MainMenu(FCIDM_MENU_EDIT);
			var mii = api.Memory("MENUITEMINFO");
			mii.cbSize = mii.Size;
			mii.fMask = MIIM_ID;
			for (var i = api.GetMenuItemCount(hMenu); i-- > 0;) {
				if (api.GetMenuItemInfo(hMenu, i, true, mii)) {
					if ((mii.wID & 0xfff) == Verb) {
						api.SendMessage(FV.hwndView, WM_COMMAND, mii.wID, 0);
						hr = S_OK;
						break;
					}
				}
			}
		}
	}
	RunEvent1("ConfigChanged", "Window");
	if (isFinite(hr)) {
		return hr;
	}
	if ("string" === typeof Directory && !api.PathIsDirectory(Directory)) {
		return S_FALSE;
	}
	var Items = ContextMenu.Items();
	var Exec = [];
	if (api.PathMatchSpec(strVerb, "opennewwindow;opennewprocess")) {
		CancelWindowRegistered();
	}

	var NewTab = GetNavigateFlags();
	for (var i = 0; i < Items.Count; ++i) {
		if (Verb && strVerb != "runas") {
			var Item = Items.Item(i);
			path = Item.ExtendedProperty("linktarget") || Item.Path;
			var cmd = api.AssocQueryString(ASSOCF_NONE, ASSOCSTR_COMMAND, Item.ExtendedProperty("linktarget") || Item, strVerb == "default" ? null : strVerb);
			if (cmd) {
				if (strVerb == "open" && api.PathMatchSpec(cmd, "?:\\Windows\\Explorer.exe;*\\Explorer.exe /idlist,*;rundll32.exe *fldr.dll,RouteTheCall*")) {
					Navigate(Items.Item(i), NewTab);
					NewTab |= SBSP_NEWBROWSER;
					continue;
				}
				var cmd2 = ExtractMacro(te, cmd);
				if (!SameText(cmd, cmd2)) {
					ShellExecute(cmd2.replace(/"?%1"?|%L/g, api.PathQuoteSpaces(path)).replace(/%\*|%I/g, ""), null, nShow, Directory);
					continue;
				}
			}
		}
		Exec.push(Items.Item(i));
	}
	if (Items.Count != Exec.length) {
		if (Exec.length) {
			var Selected = api.CreateObject("FolderItems");
			for (var i = 0; i < Exec.length; ++i) {
				Selected.AddItem(Exec[i]);
			}
			InvokeCommand(Selected, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon, ContextMenu.FolderView);
		}
		return S_OK;
	}
	return S_FALSE;
}

AddEvent("InvokeCommand", function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon, strVerb) {
	if (strVerb == "opencontaining") {
		var Items = ContextMenu.Items();
		for (var i = 0; i < Items.Count; ++i) {
			var Item = Items.Item(i);
			var path = Item.ExtendedProperty("linktarget") || Item.Path;
			Navigate(GetParentFolderName(path), SBSP_NEWBROWSER);
			setTimeout(function (FV, path) {
				FV.SelectItem(path, SVSI_SELECT | SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS);
			}, 99, te.Ctrl(CTRL_FV), path);
			return S_OK;
		}
	}
	var FV = ContextMenu.FolderView;
	var Items = ContextMenu.Items();
	if (strVerb == "delete") {
		for (var i = Items.Count; i--;) {
			UnlockFV(Items.Item(i));
		}
	}
	if (FV) {
		if (te.OnBeforeGetData(FV, Items, 2)) {
			return S_OK;
		}
	}
});

te.OnDragEnter = function (Ctrl, dataObj, pgrfKeyState, pt, pdwEffect) {
	let hr = E_NOTIMPL;
	const dwEffect = pdwEffect[0];
	const en = "DragEnter";
	const eo = eventTE[en.toLowerCase()];
	for (let i in eo) {
		try {
			if (pdwEffect[0]) {
				pdwEffect[0] = dwEffect;
			}
			const hr2 = api.Invoke(eo[i], [Ctrl, dataObj, pgrfKeyState[0], pt, pdwEffect, pgrfKeyState]);
			if (isFinite(hr2) && hr != S_OK) {
				hr = hr2;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	g_.mouse.str = "";
	return hr;
}

te.OnDragOver = function (Ctrl, dataObj, pgrfKeyState, pt, pdwEffect) {
	const dwEffect = pdwEffect[0];
	const en = "DragOver";
	const eo = eventTE[en.toLowerCase()];
	for (let i in eo) {
		try {
			pdwEffect[0] = dwEffect;
			const hr = api.Invoke(eo[i], [Ctrl, dataObj, pgrfKeyState[0], pt, pdwEffect, pgrfKeyState]);
			if (isFinite(hr)) {
				return hr;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	return E_NOTIMPL;
}

te.OnDrop = function (Ctrl, dataObj, pgrfKeyState, pt, pdwEffect) {
	const dwEffect = pdwEffect[0];
	const en = "Drop";
	const eo = eventTE[en.toLowerCase()];
	for (let i in eo) {
		try {
			pdwEffect[0] = dwEffect;
			const hr = api.Invoke(eo[i], [Ctrl, dataObj, pgrfKeyState[0], pt, pdwEffect, pgrfKeyState]);
			if (isFinite(hr)) {
				return hr;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	return E_NOTIMPL;
}

te.OnDragLeave = function (Ctrl) {
	let hr = E_NOTIMPL;
	const en = "DragLeave";
	const eo = eventTE[en.toLowerCase()];
	for (let i in eo) {
		try {
			const hr2 = api.Invoke(eo[i], [Ctrl]);
			if (isFinite(hr2) && hr != S_OK) {
				hr = hr2;
			}
		} catch (e) {
			ShowError(e, en, i);
		}
	}
	g_.mouse.str = "";
	return hr;
}

te.OnSelectionChanged = function (Ctrl, uChange) {
	if (Ctrl.Type == CTRL_TC && Ctrl.SelectedIndex >= 0) {
		ShowStatusText(Ctrl.Selected, "", 0);
		ChangeView(Ctrl.Selected);
	}
	return RunEvent3("SelectionChanged", Ctrl, uChange);
}

te.OnFilterChanged = function (Ctrl) {
	if (/search\-ms:.*?crumb=[^&]+/.test(Ctrl.FilterView) || isFinite(RunEvent3("FilterChanged", Ctrl))) {
		return;
	}
	var res = /\/(.*)\/(.*)/.exec(Ctrl.FilterView);
	if (res) {
		try {
			Ctrl.Data.RE = new RegExp((window.migemo && migemo.query(res[1])) || res[1], res[2]);
			Ctrl.OnIncludeObject = function (Ctrl, Path1, Path2, Item) {
				return Ctrl.Data.RE.test(Path1) || (Path1 != Path2 && Ctrl.Data.RE.test(Path2)) ? S_OK : S_FALSE;
			}
			return;
		} catch (e) { }
	}
}

te.OnShowContextMenu = function (Ctrl, hwnd, msg, wParam, pt, ContextMenu) {
	if (g_.mouse.CancelContextMenu) {
		return S_OK;
	}
	var hr = RunEvent3("ShowContextMenu", Ctrl, hwnd, msg, wParam, pt, ContextMenu);
	if (isFinite(hr)) {
		return hr;
	}
	switch (Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
			var Selected = Ctrl.SelectedItems();
			if (Selected.Count) {
				if (ExecMenu(Ctrl, "Context", pt, 1) == S_OK) {
					return S_OK;
				}
			} else {
				if (ExecMenu(Ctrl, "Background", pt, 1) == S_OK) {
					return S_OK;
				}
			}
			break;
		case CTRL_TV:
			if (ExecMenu(Ctrl, "Tree", pt, 1, false, ContextMenu) == S_OK) {
				return S_OK;
			}
			break;
		case CTRL_TC:
			if (ExecMenu(Ctrl, "Tabs", pt, 1) == S_OK) {
				return S_OK;
			}
			break;
		case CTRL_WB:
			if (ShowContextMenu) {
				return ShowContextMenu(Ctrl, hwnd, msg, wParam, pt);
			}
			if (wParam == CONTEXT_MENU_DEFAULT) {
				return S_OK;
			}
			break;
		case CTRL_TE:
			api.GetSystemMenu(te.hwnd, true);
			var hMenu = api.GetSystemMenu(te.hwnd, false);
			var menus = te.Data.xmlMenus.getElementsByTagName("System");
			if (menus && menus.length) {
				var items = menus[0].getElementsByTagName("Item");
				var arMenu = OpenMenu(items, null);
				MakeMenus(hMenu, menus, arMenu, items, Ctrl, pt, 0, null, true);
			}
			api.DestroyMenu(hMenu);
			break;
	}
	return S_FALSE;
}

te.OnDefaultCommand = function (Ctrl) {
	if (api.GetKeyState(VK_MENU) < 0) {
		api.SendMessage(Ctrl.hwndView, WM_COMMAND, CommandID_PROPERTIES - 1, 0);
		return S_OK;
	}
	var Selected = Ctrl.SelectedItems();
	var hr = RunEvent3("DefaultCommand", Ctrl, Selected);
	if (isFinite(hr)) {
		return hr;
	}
	if (ExecMenu(Ctrl, "Default", null, 2) == S_OK) {
		return S_OK;
	}
	if (Selected.Count == 1) {
		var pid = api.ILCreateFromPath(api.GetDisplayNameOf(Selected.Item(0), SHGDN_FORPARSING | SHGDN_FORADDRESSBAR | SHGDN_ORIGINAL));
		if (pid.Enum) {
			Ctrl.Navigate(pid, GetNavigateFlags(Ctrl));
			return S_OK;
		}
	}
	return InvokeCommand(Selected, 0, te.hwnd, null, null, null, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
}

te.OnSystemMessage = function (Ctrl, hwnd, msg, wParam, lParam) {
	if (!te.Data) {
		return S_OK;
	}
	var hr = RunEvent3("SystemMessage", Ctrl, hwnd, msg, wParam, lParam);
	if (isFinite(hr)) {
		return hr;
	}
	switch (Ctrl.Type) {
		case CTRL_TE:
			switch (msg) {
				case WM_DESTROY:
					var pid = api.Memory("DWORD");
					api.GetWindowThreadProcessId(te.hwnd, pid);
					WmiProcess("WHERE ExecutablePath = '" + (api.GetModuleFileName(null).split("\\").join("\\\\")) + "' AND ProcessId!=" + pid[0], function (item) {
						var hwnd = GethwndFromPid(item.ProcessId);
						api.SetWindowLongPtr(hwnd, GWL_EXSTYLE, api.GetWindowLongPtr(hwnd, GWL_EXSTYLE) & ~0x80);
						api.ShowWindow(hwnd, SW_SHOWNORMAL);
						api.SetForegroundWindow(hwnd);
					});
					for (var i in te.Data.Fonts) {
						api.DeleteObject(te.Data.Fonts[i]);
					}
					te.Data.Fonts = null;
					for (var i = SHIL_JUMBO + 1; i--;) {
						api.ImageList_Destroy(te.Data.SHIL[i], true);
					}
					te.Data.SHIL.length = 0;
					break;
				case WM_DEVICECHANGE:
					if (wParam == DBT_DEVICEARRIVAL || wParam == DBT_DEVICEREMOVECOMPLETE) {
						DeviceChanged(Ctrl);
					}
					break;
				case WM_ACTIVATE:
					g_.focused = void 0;
					if (te.Data.bSaveMenus) {
						te.Data.bSaveMenus = false;
						SaveXmlEx("menus.xml", te.Data.xmlMenus);
					}
					if (te.Data.bSaveAddons) {
						te.Data.bSaveAddons = false;
						SaveXmlEx("addons.xml", te.Data.Addons);
					}
					if (te.Data.bSaveConfig) {
						te.Data.bSaveConfig = false;
						SaveConfigXML(BuildPath(te.Data.DataFolder, "config\\window.xml"));
					}
					if (te.Data.bReload) {
						ReloadCustomize();
					}
					if (g_.TEData) {
						var FV = te.Ctrl(CTRL_FV);
						var Layout = g_.TEData.Layout;
						if (FV) {
							FV.CurrentViewMode = g_.TEData.ViewMode;
						}
						te.Layout = Layout;
						delete g_.TEData;
					}
					if (g_.FVData) {
						if (g_.FVData.All) {
							delete g_.FVData.All;
							var cFV = te.Ctrls(CTRL_FV);
							for (var i in cFV) {
								SetFolderViewData(cFV[i], g_.FVData);
							}
						} else {
							SetFolderViewData(te.Ctrl(CTRL_FV), g_.FVData);
						}
						delete g_.FVData;
					}
					if (g_.TVData) {
						if (g_.TVData.All) {
							var cFV = te.Ctrls(CTRL_FV);
							for (var i in cFV) {
								SetTreeViewData(cFV[i], g_.TVData);
							}
						} else {
							SetTreeViewData(te.Ctrl(CTRL_FV), g_.TVData);
						}
						delete g_.TVData;
					}
					if (wParam & 0xffff) {
						if (g_.mouse.str == "") {
							setTimeout(function () {
								const hFocus = api.GetFocus();
								if (!hFocus || hFocus == te.hwnd) {
									const FV = te.Ctrl(CTRL_FV);
									if (FV) {
										FV.Focus();
									}
								}
							}, 99);
						}
					} else {
						g_.mouse.str = "";
						SetGestureText(Ctrl, "");
					}
					break;
				case WM_QUERYENDSESSION:
					SaveConfig();
					return true;
				case WM_SYSCOMMAND:
					if (wParam < 0xf000) {
						var menus = te.Data.xmlMenus.getElementsByTagName("System");
						if (menus && menus.length) {
							var items = menus[0].getElementsByTagName("Item");
							if (items) {
								var item = items[wParam - 1];
								if (item) {
									Exec(Ctrl, item.text, item.getAttribute("Type"), null);
									return 1;
								}
							}
						}
					} else {
						if (!api.IsZoomed(te.hwnd) && !api.IsIconic(te.hwnd)) {
							api.GetWindowRect(te.hwnd, te.Data.rcWindow);
						}
					}
					break;
				case WM_POWERBROADCAST:
					if (wParam == PBT_APMRESUMEAUTOMATIC) {
						var cFV = te.Ctrls(CTRL_FV);
						for (var i in cFV) {
							var FV = cFV[i];
							if (FV.hwndView) {
								if (api.PathIsNetworkPath(api.GetDisplayNameOf(FV, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING))) {
									FV.Suspend();
								}
							}
						}
					}
					break;
				case WM_SYSCOLORCHANGE:
					var cFV = te.Ctrls(CTRL_FV);
					for (var i in cFV) {
						var FV = cFV[i];
						if (FV.hwndView) {
							FV.Suspend();
						}
					}
					break;
			}
			break;
		case CTRL_TC:
			if (msg == WM_SHOWWINDOW && wParam == 0) {
				SetDisplay("Panel_" + Ctrl.Id, "none");
			}
			break;
		case CTRL_SB:
		case CTRL_EB:
			if (msg == WM_SHOWWINDOW && te.hwnd == api.GetFocus()) {
				Ctrl.Focus();
			}
			break;
	}
	if (msg == WM_COPYDATA) {
		var cd = api.Memory("COPYDATASTRUCT", 1, lParam);
		var hr = RunEvent3("CopyData", Ctrl, cd, wParam);
		if (isFinite(hr)) {
			return hr;
		}
		if (Ctrl.Type == CTRL_TE && cd.dwData == 0 && cd.cbData) {
			var strData = api.SysAllocStringByteLen(cd.lpData, cd.cbData, cd.cbData);
			RestoreFromTray();
			RunCommandLine(strData);
			return S_OK;
		}
	}
};

te.OnMenuMessage = function (Ctrl, hwnd, msg, wParam, lParam) {
	var pStatus = [null, "", 0];
	var hr = RunEvent3("MenuMessage", Ctrl, hwnd, msg, wParam, lParam, pStatus);
	if (isFinite(hr)) {
		ShowStatusText(pStatus[0], pStatus[1], pStatus[2], pStatus[3]);
		return hr;
	}
	switch (msg) {
		case WM_INITMENUPOPUP:
			var s = api.GetMenuString(wParam, 0, MF_BYPOSITION);
			if (api.PathMatchSpec(s, "\t*Script\t*")) {
				var ar = s.split("\t");
				api.DeleteMenu(wParam, 0, MF_BYPOSITION);
				ExecScriptEx(Ctrl, ar[2], ar[1], hwnd);
			}
			break;
		case WM_MENUSELECT:
			pStatus[0] = te;
			RunEvent1("MenuSelect", Ctrl, hwnd, msg, wParam, lParam, pStatus);
			if (lParam) {
				var nVerb = wParam & 0xffff;
				if (Ctrl) {
					for (var i in Ctrl) {
						var CM = Ctrl[i];
						if (CM && nVerb >= CM.idCmdFirst && nVerb <= CM.idCmdLast) {
							Text = CM.GetCommandString(nVerb - CM.idCmdFirst, GCS_HELPTEXT);
							if (Text) {
								pStatus[1] = Text;
								break;
							}
						}
					}
				}
				var mf = wParam >> 16;
				if (FolderMenu.MenuLoop) {
					var wId = nVerb;
					if (mf & MF_POPUP) {
						var mii = api.Memory("MENUITEMINFO");
						mii.cbSize = mii.Size;
						mii.fMask = MIIM_ID;
						api.GetMenuItemInfo(lParam, nVerb, true, mii);
						wId = mii.wId;
					}
					var pid = FolderMenu.Items[wId - 1];
					if (pid) {
						g_.MenuSelected = pid;
						pStatus = [pid, pid.Name, 0, 200];
					}
				}
				var hSubMenu = api.GetSubMenu(lParam, nVerb);
				if (mf & MF_POPUP) {
					if (hSubMenu) {
						g_.menu_handle = lParam;
						g_.menu_pos = nVerb;
					}
				}
				if (!(mf & MF_MOUSESELECT)) {
					g_.menu_state |= 1;
				}
				window.g_menu_string = api.GetMenuString(lParam, nVerb, hSubMenu ? MF_BYPOSITION : MF_BYCOMMAND);
			}
			break;
		case WM_ENTERMENULOOP:
			g_.menu_loop = true;
			g_.menu_state = te.Data.cmdKey & ((te.Data.cmdKeyF && (te.Data.cmdKey & 0x17f) == 15) ? 0xf000 : 0);
			RunEvent1("EnterMenuLoop", Ctrl, hwnd, msg, wParam, lParam);
			break;
		case WM_EXITMENULOOP:
			window.g_menu_click = false;
			var ar = ["ExitMenuLoop", "EnterMenuLoop", "MenuSelect", "MenuChar"];
			RunEvent1(ar[0], Ctrl, hwnd, msg, wParam, lParam);
			for (var i = ar.length; i--;) {
				var eo = eventTE[ar[i].toLowerCase()];
				if (eo) {
					eo.length = 0;
				}
			}
			for (var i in g_arBM) {
				api.DeleteObject(g_arBM[i]);
			}
			g_arBM = [];
			g_.menu_loop = false;
			delete g_.ptMenuDrag;
			delete g_.MenuSelected
			pStatus = [GetFolderView(), "", 0];
			break;
		case WM_MENUCHAR:
			var hr = RunEvent4("MenuChar", Ctrl, hwnd, msg, wParam, lParam);
			if (hr !== void 0) {
				ShowStatusText(pStatus[0], pStatus[1], pStatus[2], pStatus[3]);
				return hr;
			}
			if (window.g_menu_click && (wParam & 0xffff) == VK_LBUTTON) {
				ShowStatusText(pStatus[0], pStatus[1], pStatus[2], pStatus[3]);
				return MNC_EXECUTE << 16 + g_.menu_pos;
			}
			break;
	}
	ShowStatusText(pStatus[0], pStatus[1], pStatus[2], pStatus[3]);
	return S_FALSE;
};

te.OnAppMessage = function (Ctrl, hwnd, msg, wParam, lParam) {
	var hr = RunEvent3("AppMessage", Ctrl, hwnd, msg, wParam, lParam);
	if (isFinite(hr)) {
		return hr;
	}
	if (msg == TWM_CHANGENOTIFY) {
		var pidls = {};
		var hLock = api.SHChangeNotification_Lock(wParam, lParam, pidls);
		if (hLock) {
			api.SHChangeNotification_Unlock(hLock);
			ChangeNotifyFV(pidls.lEvent, pidls[0], pidls[1]);
			RunEvent1("ChangeNotify", Ctrl, pidls, wParam, lParam);
			if (pidls.lEvent & (SHCNE_UPDATEITEM | SHCNE_RENAMEITEM)) {
				var n = pidls.lEvent & SHCNE_RENAMEITEM ? 1 : 0;
				var path = api.GetDisplayNameOf(pidls[n], SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL);
				RunEvent1("ChangeNotifyItem:" + path, api.ILCreateFromPath(path) || pidls[n]);
			}
		}
		return S_OK;
	}
	return S_FALSE;
}

te.OnNewWindow = function (Ctrl, dwFlags, UrlContext, Url) {
	var hr = RunEvent3("NewWindow", Ctrl, dwFlags, UrlContext, Url);
	if (isFinite(hr)) {
		return hr;
	}
	var FolderItem = api.ILCreateFromPath(api.PathCreateFromUrl(Url));
	if (FolderItem.IsFolder) {
		Navigate(FolderItem, SBSP_NEWBROWSER);
		return S_OK;
	}
	return S_FALSE;
}

te.OnClipboardText = function (Items) {
	if (!Items) {
		return "";
	}
	var r = RunEvent4("ClipboardText", Items);
	if (r !== void 0) {
		return r;
	}
	var s = [];
	for (var i = Items.Count; i-- > 0;) {
		s.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(Items.Item(i), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL)))
	}
	return s.join(" ");
}

te.OnILGetParent = function (FolderItem) {
	var r = RunEvent4("ILGetParent", FolderItem);
	if (r !== void 0) {
		return r;
	}
	var res = IsSearchPath(FolderItem);
	if (res) {
		return decodeURIComponent(res[1]);
	}
	if (api.ILIsEqual(FolderItem.Alt, ssfRESULTSFOLDER)) {
		var path = api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
		var ar = path.split && path.slice(3).split("\\") || [];
		if (ar.pop() && ar.join("\\")) {
			return path.slice(0, 3) + ar.join("\\");
		}
		return /^[A-Z]:\\/i.test(path) ? ssfDRIVES : /^\\\\\w/i.test(path) ? ssfNETWORK : ssfDESKTOP;
	}
}

te.OnReplacePath = function (Ctrl, Path) {
	if (/^[A-Z]:\\.+?\\|^\\\\.+?\\/i.test(Path)) {
		var i = Path.indexOf("\\/");
		if (i > 0) {
			var fn = Path.slice(i + 1);
			if (/\//.test(fn)) {
				Ctrl.FilterView = fn;
				return Path.slice(0, i);
			}
		}
		var fn = GetFileName(Path);
		if (/[\*\?]/.test(fn)) {
			Ctrl.FilterView = fn;
			return GetParentFolderName(Path);
		}
	}
	var res = IsSearchPath(Path);
	if (res) {
		Ctrl.FilterView = Path.replace(/displayname=[^&]+&|&crumb=location:[^&]+/g, "");
		return;
	}
	return RunEvent4("ReplacePath", Ctrl, Path);
}

te.OnGetAlt = function (dwSessionId, s) {
	var cFV = te.Ctrls(CTRL_FV);
	for (var i in cFV) {
		var FV = cFV[i];
		if (dwSessionId == FV.SessionId) {
			return BuildPath(FV.FolderItem.Path, GetFileName(s));
		}
	}
}

te.OnFilterView = function (FV, s)
{
	if (GetLock(FV) && !IsSearchPath(FV)) {
		var fn = function (strMatch, ref) {
			return encodeURIComponent(ref)
		};
		FV.Navigate(["search-ms:crumb=", s.replace(/([ -\\]+)/g, fn), "&crumb=location:", FV.FolderItem.Path.replace(/([ -\\]+)/g, fn)].join(""), SBSP_NEWBROWSER);
		return S_OK;
	}
	return S_FALSE;
}

te.OnVisibleChanged = function (Ctrl) {
	RunEvent1("VisibleChanged", Ctrl);
}

ShowStatusText = function (Ctrl, Text, iPart, tm) {
	if (!Ctrl) {
		return;
	}
	if (tm) {
		ShowStatusTextEx(Ctrl, Text, iPart, tm);
		return S_OK;
	}
	RunEvent1("StatusText", Ctrl, Text, iPart);
	return S_OK;
}

g_.event.statustext = ShowStatusText;

g_.event.beginnavigate = function (Ctrl) {
	ShowStatusText(Ctrl, "", 0);
	return RunEvent2("BeginNavigate", Ctrl);
}

g_.event.selectionchanging = function (Ctrl) {
	return RunEvent3("SelectionChanging", Ctrl);
}

g_.event.viewmodechanged = function (Ctrl) {
	RunEvent1("ViewModeChanged", Ctrl);
}

g_.event.columnschanged = function (Ctrl) {
	RunEvent1("ColumnsChanged", Ctrl);
}

g_.event.contentschanged = function (Ctrl) {
	RunEvent1("ContentsChanged", Ctrl);
}

g_.event.iconsizechanged = function (Ctrl) {
	RunEvent1("IconSizeChanged", Ctrl);
}

g_.event.itemclick = function (Ctrl, Item, HitTest, Flags) {
	return RunEvent3("ItemClick", Ctrl, Item, HitTest, Flags);
}

g_.event.columnclick = function (Ctrl, iItem) {
	return RunEvent3("ColumnClick", Ctrl, iItem);
}

g_.event.tooltip = function (Ctrl, Index) {
	return RunEvent4("ToolTip", Ctrl, Index);
}

g_.event.hittest = function (Ctrl, pt, flags) {
	return RunEvent3("HitTest", Ctrl, pt, flags);
}

g_.event.getpanestate = function (Ctrl, ep, peps) {
	return RunEvent3("GetPaneState", Ctrl, ep, peps);
}

g_.event.translatepath = function (Ctrl, Path) {
	return RunEvent4("TranslatePath", Ctrl, Path);
}

g_.event.begindrag = function (Ctrl) {
	return !isFinite(RunEvent3("BeginDrag", Ctrl));
}


g_.event.beforegetdata = function (Ctrl, Items, nMode) {
	return RunEvent2("BeforeGetData", Ctrl, Items, nMode);
}

g_.event.beginlabeledit = function (Ctrl) {
	return RunEvent4("BeginLabelEdit", Ctrl);
}

g_.event.endlabeledit = function (Ctrl, Name) {
	return RunEvent4("EndLabelEdit", Ctrl, Name);
}

g_.event.sort = function (Ctrl) {
	return RunEvent1("Sort", Ctrl);
}

g_.event.sorting = function (Ctrl, Name) {
	return RunEvent3("Sorting", Ctrl, Name);
}

g_.event.setname = function (pid, Name) {
	return RunEvent3("SetName", pid, Name);
}

g_.event.includeitem = function (Ctrl, pid) {
	return RunEvent2("IncludeItem", Ctrl, pid);
}

g_.event.endthread = function (Ctrl) {
	return RunEvent1("EndThread", api.GetThreadCount());
}

g_.event.itemprepaint = function (Ctrl, pid, nmcd, vcd, plRes) {
	RunEvent3("ItemPrePaint", Ctrl, pid, nmcd, vcd, plRes);
	RunEvent1("ItemPrePaint2", Ctrl, pid, nmcd, vcd, plRes);
}

g_.event.itempostpaint = function (Ctrl, pid, nmcd, vcd) {
	RunEvent3("ItemPostPaint", Ctrl, pid, nmcd, vcd);
	RunEvent1("ItemPostPaint2", Ctrl, pid, nmcd, vcd);
}

g_.event.handleicon = function (Ctrl, pid, iItem) {
	return RunEvent3("HandleIcon", Ctrl, pid, iItem);
}

//Tablacus Events

GetIconImage = function (Ctrl, BGColor, bSimple) {
	var nSize = api.GetSystemMetrics(SM_CYSMICON);
	var FolderItem = Ctrl.FolderItem || Ctrl;
	var r = GetNetworkIcon(FolderItem.Path);
	if (r) {
		if (FolderItem.Unavailable > 500) {
			return MakeImgDataEx("icon:shell32.dll,234", bSimple, nSize);
		}
		return MakeImgDataEx(r, bSimple, nSize);
	}
	var img = RunEvent4("GetIconImage", Ctrl, BGColor, bSimple);
	if (img) {
		return MakeImgDataEx(img, bSimple, nSize);
	}
	api.ILIsEmpty(FolderItem);
	if (FolderItem.Unavailable) {
		return MakeImgDataEx("icon:shell32.dll,234", bSimple, nSize);
	}
	if (g_.IEVer >= 8) {
		if (bSimple) {
			return bSimple != 2 ? api.GetDisplayNameOf(FolderItem, SHGDN_FORPARSING | SHGDN_ORIGINAL) : "";
		}
		var sfi = api.Memory("SHFILEINFO");
		api.SHGetFileInfo(FolderItem, 0, sfi, sfi.Size, SHGFI_SYSICONINDEX | SHGFI_PIDL);
		var hIcon = GetHICON(sfi.iIcon, nSize, ILD_NORMAL);
		if (hIcon) {
			img = api.CreateObject("WICBitmap").FromHICON(hIcon);
			api.DestroyIcon(hIcon);
			return img.DataURI();
		}
	}
	return MakeImgDataEx("icon:shell32.dll,3", bSimple, nSize);
}

UnlockFV = function (Item) {
	var cFV = te.Ctrls(CTRL_FV);
	for (var i in cFV) {
		var FV = cFV[i];
		if (FV.hwndView && api.ILIsParent(Item, FV, false)) {
			FV.Suspend();
		}
	}
}

ChangeNotifyFV = function (lEvent, item1, item2) {
	var fAdd = SHCNE_DRIVEADD | SHCNE_MEDIAINSERTED | SHCNE_NETSHARE | SHCNE_MKDIR | SHCNE_UPDATEDIR | SHCNE_UPDATEITEM;
	var fRemove = SHCNE_DRIVEREMOVED | SHCNE_MEDIAREMOVED | SHCNE_NETUNSHARE | SHCNE_RENAMEFOLDER | SHCNE_RMDIR | SHCNE_SERVERDISCONNECT | SHCNE_UPDATEDIR;
	if (lEvent & (SHCNE_DISKEVENTS | fAdd | fRemove) && item1.IsFileSystem) {
		var path1 = item1.Path;
		var bNetwork = api.ILIsEqual(item1, ssfNETWORK);
		var cFV = te.Ctrls(CTRL_FV);
		for (var i in cFV) {
			var FV = cFV[i];
			if (FV && FV.FolderItem) {
				var path = FV.FolderItem.Path;
				var bParent = api.PathMatchSpec(path, [path1.replace(/\\$/, ""), path1].join("\\*;")) || bNetwork && api.PathIsNetworkPath(path);
				if (lEvent == SHCNE_RENAMEFOLDER && CanClose(FV) == S_OK) {
					if (bParent) {
						FV.Navigate(path.replace(path1, api.GetDisplayNameOf(item2, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL)), SBSP_SAMEBROWSER);
						continue;
					}
				}
				if (FV.hwndView) {
					var bChild = SameText(GetParentFolderName(path1), path);
					if (bChild) {
						if (FV.hwndList) {
							var item = api.Memory("LVITEM");
							item.stateMask = LVIS_CUT;
							api.SendMessage(FV.hwndList, LVM_SETITEMSTATE, -1, item);
						}
						if (WINVER >= 0x600 && (lEvent & (SHCNE_DELETE | SHCNE_RMDIR))) {
							var nPos = FV.GetFocusedItem - 1;
							if (nPos > 0 && api.ILIsEqual(item1, FV.FocusedItem)) {
								var nCount = FV.ItemCount && FV.ItemCount(SVGIO_ALLVIEW);
								if (nCount) {
									FV.SelectItem(nPos < nCount ? nPos : nCount - 1, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | (FV.Id == te.Ctrl(CTRL_FV).Id ? 0 : SVSI_NOTAKEFOCUS));
								}
							}
						}
					}
					if (bChild || bParent) {
						if ((lEvent & fRemove) || ((lEvent & fAdd) && FV.FolderItem.Unavailable)) {
							RefreshEx(FV, 500, 500);
						}
					}
					FV.Notify(lEvent, item1, item2);
				}
			}
		}
	}
}

SetKeyExec = function (mode, strKey, path, type, bLast) {
	if (/,/.test(strKey) && !/,$/.test(strKey)) {
		var ar = strKey.split(",");
		for (var i in ar) {
			SetKeyExec(mode, ar[i], path, type, bLast);
		}
		return;
	}
	if (strKey) {
		strKey = GetKeyKey(strKey);
		var KeyMode = eventTE.Key[mode];
		if (KeyMode) {
			if (!KeyMode[strKey]) {
				KeyMode[strKey] = [];
			}
			KeyMode[strKey][bLast ? "push" : "unshift"]([path, type]);
		}
	}
}

SetGestureExec = function (mode, strGesture, path, type, bLast, Name) {
	if (/,/.test(strGesture) && !/,$/.test(strGesture)) {
		var ar = strGesture.split(",");
		for (var i in ar) {
			SetGestureExec(mode, ar[i], path, type, bLast, Name);
		}
		return;
	}
	if (strGesture) {
		strGesture = strGesture.toUpperCase();
		var MouseMode = eventTE.Mouse[mode];
		if (MouseMode) {
			if (!MouseMode[strGesture]) {
				MouseMode[strGesture] = [];
			}
			MouseMode[strGesture][bLast ? "push" : "unshift"]([path, type, Name]);
		}
	}
}

ArExec = function (Ctrl, ar, pt, hwnd) {
	for (var i in ar) {
		var cmd = ar[i];
		if (Exec(Ctrl, cmd[0], cmd[1], hwnd, pt) === S_OK) {
			return S_OK;
		}
	}
	return S_FALSE;
}

GestureExec = function (Ctrl, mode, str, pt, hwnd) {
	if (hwnd && Ctrl.Type != CTRL_TC && Ctrl.Type != CTRL_WB) {
		api.SetFocus(hwnd);
	}
	return ArExec(Ctrl, api.ObjGetI(eventTE.Mouse[mode], str), pt, hwnd || Ctrl.hwnd);
}

KeyExec = function (Ctrl, mode, str, hwnd) {
	return KeyExecEx(Ctrl, mode, GetKeyKey(str), hwnd || Ctrl.hwnd);
}

KeyExecEx = function (Ctrl, mode, nKey, hwnd) {
	var pt = api.Memory("POINT");
	if (Ctrl.Type <= CTRL_EB || Ctrl.Type == CTRL_TV) {
		var rc = api.Memory("RECT");
		Ctrl.GetItemRect(Ctrl.FocusedItem || Ctrl.SelectedItem, rc);
		pt.x = rc.left;
		pt.y = rc.top;
	}
	api.ClientToScreen(Ctrl.hwnd, pt);
	return ArExec(Ctrl, eventTE.Key[mode][nKey], pt, hwnd);
}

importScripts = function () {
	for (var i = 0; i < arguments.length; ++i) {
		importScript(arguments[i]);
	}
}

AddEvent("Arrange", function (Ctrl, rc) {
	if (Ctrl.Type == CTRL_TE && !api.IsIconic(te.hwnd)) {
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
				var w = (rcClient.right - rcClient.left) / 100;
				var h = (rcClient.bottom - rcClient.top) / 100;
				var s = TC.Left;
				if ("string" === typeof s) {
					rc.left = s.replace(/%$/g, "") * w + rcClient.left;
				} else {
					rc.left = s + (s >= 0 ? rcClient.left : rcClient.right);
				}
				if (s) {
					++rc.left;
				}
				s = TC.Top;
				if ("string" === typeof s) {
					rc.top = s.replace(/%$/g, "") * h + rcClient.top;
				} else {
					rc.top = s + (s >= 0 ? rcClient.top : rcClient.bottom);
				}
				if (s) {
					++rc.top;
				}
				s = TC.Width;
				if ("string" === typeof s) {
					rc.right = s.replace(/%$/g, "") * w + rc.left;
				} else {
					rc.right = s + (s >= 0 ? rc.left : rcClient.right);
				}
				if (rc.right >= rcClient.right) {
					rc.right = rcClient.right;
				} else {
					--rc.right;
				}
				s = TC.Height;
				if ("string" === typeof s) {
					rc.bottom = s.replace(/%$/g, "") * h + rc.top;
				} else {
					rc.bottom = s + (s >= 0 ? rc.top : rcClient.bottom);
				}
				if (rc.bottom >= rcClient.bottom) {
					rc.bottom = rcClient.bottom;
				} else {
					--rc.bottom;
				}
				te.OnArrange(TC, rc);
			}
		} catch (e) { }
		te.UnlockUpdate(1);
	}
});

AddEvent("Exec", function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop) {
	const FV = GetFolderView(Ctrl, pt);
	const fn = g_basic.FuncI(type);
	if (fn) {
		if (dataObj) {
			if (fn.Drop) {
				return api.Invoke(fn.Drop, [Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop]);
			}
			pdwEffect[0] = DROPEFFECT_NONE;
			return E_NOTIMPL;
		}
		if (FV && !/Background|Tabs|Tools|View|.*Script/i.test(type)) {
			const hr = te.OnBeforeGetData(FV, dataObj || FV.SelectedItems(), 3);
			if (hr) {
				return hr;
			}
		}
		if (fn.Exec) {
			const hr = api.Invoke(fn.Exec, [Ctrl, s, type, hwnd, pt]);
			return hr != void 0 ? hr : fn.Result;
		}
		return g_basic.Exec(Ctrl, s, type, hwnd, pt);
	}
});

AddEvent("MenuState:Tabs:Close Tab", function (Ctrl, pt, mii) {
	var FV = GetFolderView(Ctrl, pt);
	if (CanClose(FV)) {
		mii.fMask |= MIIM_STATE;
		mii.fState |= MFS_DISABLED;
	}
});

AddEvent("MenuState:Tabs:Close Tabs on Left", function (Ctrl, pt, mii) {
	var FV = GetFolderView(Ctrl, pt, true);
	if (FV && FV.Index == 0) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_DISABLED;
	}
});

AddEvent("MenuState:Tabs:Close Tabs on Right", function (Ctrl, pt, mii) {
	var FV = GetFolderView(Ctrl, pt, true);
	if (FV) {
		var TC = FV.Parent;
		if (FV.Index >= TC.Count - 1) {
			mii.fMask |= MIIM_STATE;
			mii.fState = MFS_DISABLED;
		}
	}
});

AddEvent("MenuState:Tabs:Up", function (Ctrl, pt, mii) {
	var FV = GetFolderView(Ctrl, pt);
	if (!FV || api.ILIsEmpty(FV)) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_DISABLED;
	}
});

AddEvent("MenuState:Tabs:Lock", function (Ctrl, pt, mii) {
	var FV = GetFolderView(Ctrl, pt);
	if (FV && FV.Data.Lock) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_CHECKED;
	}
});

AddEvent("MenuState:Tabs:Show frames", function (Ctrl, pt, mii) {
	if (WINVER < 0x600) {
		mii.fMask = 0;
		return S_OK;
	}
	var FV = GetFolderView(Ctrl, pt);
	if (!FV || FV.Type == CTRL_EB) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_CHECKED;
	}
});

AddEvent("ChangeNotify", function (Ctrl, pidls, wParam, lParam) {
	if (pidls.lEvent & (SHCNE_MKDIR | SHCNE_CREATE) && pidls[0].IsFileSystem) {
		var tm = new Date().getTime();
		if (g_.NewItemTime > tm) {
			delete g_.NewItemTime;
		} else if (g_.NewFolderTime > tm) {
			delete g_.NewFolderTime;
			g_.NewFolderTV.Expand(pidls[0], 0);
			wsh.SendKeys("{F2}");
		}
	}
});

AddEvent("AddType", function (arFunc) {
	for (var i in g_basic.Func) {
		arFunc.push(i);
	}
});

AddType = function (strType, o) {
	api.ObjPutI(g_basic.Func, strType.replace(/&|\.\.\.$/g, ""), o);
};

AddTypeEx = function (strType, strTitle, fn) {
	var type = g_basic.FuncI(strType);
	if (type && type.Cmd) {
		api.ObjPutI(type.Cmd, strTitle, fn);
	}
};

AddEvent("OptionRef", function (Id, s, pt) {
	var fn = g_basic.FuncI(Id);
	if (fn) {
		var r;
		if (fn.Ref) {
			return fn.Ref(s, pt);
		}
		if (fn.Cmd) {
			return g_basic.Popup(fn.Cmd, s, pt);
		}
	}
});

AddEvent("OptionEncode", function (Id, p) {
	if (Id === "") {
		var lines = p.s.split(/\r?\n/);
		for (var i in lines) {
			var res = /^([^,]+),(.*)$/.exec(lines[i]);
			if (res) {
				var p2 = { s: res[2] };
				Id = GetSourceText(res[1]);
				OptionEncode(Id, p2);
				lines[i] = [Id, p2.s].join(",");
			}
		}
		p.s = lines.join("\n");
		return S_OK;
	}
	var type = g_basic.FuncI(Id);
	if (type && type.Enc) {
		p.s = GetSourceText(p.s);
		return S_OK;
	}
});

AddEvent("OptionDecode", function (Id, p) {
	if (Id === "") {
		var lines = p.s.split(/\r?\n/);
		for (var i in lines) {
			var res = /^([^,]+),(.*)$/.exec(lines[i]);
			if (res) {
				var p2 = { s: res[2] };
				Id = res[1];
				OptionDecode(Id, p2);
				lines[i] = [GetText(Id), p2.s].join(",");
			}
		}
		p.s = lines.join("\n");
		return S_OK;
	}
	var type = g_basic.FuncI(Id);
	if (type && type.Enc) {
		var s = GetText(p.s);
		if (SameText(GetSourceText(s), p.s)) {
			p.s = GetText(p.s);
			return S_OK;
		}
		return S_OK;
	}
});

AddEvent("BeginNavigate", function (Ctrl) {
	var fn = Ctrl.FolderItem.Enum;
	if (fn) {
		var SessionId = Ctrl.SessionId;
		var Items = fn(Ctrl.FolderItem, Ctrl, function (Ctrl, Items) {
			if (Ctrl.SessionId == SessionId) {
				Ctrl.AddItems(Items, true, true);
			}
		}, SessionId);
		if (Items) {
			Ctrl.AddItems(Items, true, true);
		}
		return S_FALSE;
	}
});

AddEvent("UseExplorer", function (pid) {
	if (pid && pid.Path && !api.GetAttributesOf(pid.Alt || pid, SFGAO_FILESYSTEM | SFGAO_FILESYSANCESTOR | SFGAO_STORAGEANCESTOR | SFGAO_NONENUMERATED | SFGAO_DROPTARGET) && !api.ILIsParent(1, pid, false)) {
		return true;
	}
});

AddEvent("LocationPopup", function (hMenu) {
	FolderMenu.AddMenuItem(hMenu, api.ILCreateFromPath(ssfDESKTOP));
	FolderMenu.AddMenuItem(hMenu, api.ILCreateFromPath(ssfDRIVES));
	var Items = FolderMenu.Enum(api.ILCreateFromPath(ssfDRIVES));
	var path0 = api.GetDisplayNameOf(ssfDESKTOP, SHGDN_ORIGINAL | SHGDN_FORPARSING);
	for (var i = 0; i < Items.Count; ++i) {
		var Item = Items.Item(i);
		if (IsFolderEx(Item)) {
			var path = api.GetDisplayNameOf(Item, SHGDN_ORIGINAL | SHGDN_FORPARSING);
			if (path && path != path0) {
				FolderMenu.AddMenuItem(hMenu, Item);
			}
		}
	}
	FolderMenu.AddMenuItem(hMenu, api.ILCreateFromPath(ssfBITBUCKET), api.GetDisplayNameOf(ssfBITBUCKET, SHGDN_INFOLDER), true);
});

AddEvent("ReplaceMacroEx", [/%res:(.+)%/ig, function (strMatch, ref1) {
	return api.LoadString(hShell32, ref1) || GetTextR(ref1);
}]);

AddEvent("ReplaceMacroEx", [/%AddonStatus:([^%]*)%/ig, function (strMatch, ref1) {
	return GetNum(GetAddonElement(ref1).getAttribute("Enabled")) ? "on" : "off";
}]);

if (!window.chrome) {
	AddEvent("BrowserCreatedEx", 'MainWindow.RunEvent1("BrowserCreated", document);');
}

if (WINVER >= 0x600 && screen.deviceYDPI > 96) {
	AddEvent("NavigateComplete", FixIconSpacing);
	AddEvent("IconSizeChanged", FixIconSpacing);
	AddEvent("ViewModeChanged", FixIconSpacing);
}

AddEnv("Selected", function (Ctrl) {
	var ar = [];
	var Selected = GetSelectedItems(Ctrl);
	if (Selected) {
		for (var i = Selected.Count; i > 0; ar.unshift(api.PathQuoteSpaces(api.GetDisplayNameOf(Selected.Item(--i), SHGDN_FORPARSING | SHGDN_ORIGINAL)))) {
		}
	}
	return ar.join(" ");
});

AddEnv("Current", function (Ctrl) {
	var strSel = "";
	var FV = GetFolderView(Ctrl);
	if (FV) {
		strSel = api.PathQuoteSpaces(api.GetDisplayNameOf(FV, SHGDN_FORPARSING | SHGDN_ORIGINAL));
	}
	return strSel;
});

AddEnv("TreeSelected", function (Ctrl) {
	var strSel = "";
	if (!Ctrl || Ctrl.Type != CTRL_TV) {
		var FV = GetFolderView(Ctrl);
		if (FV) {
			Ctrl = FV.TreeView;
		}
	}
	if (Ctrl) {
		strSel = api.PathQuoteSpaces(api.GetDisplayNameOf(Ctrl.SelectedItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_ORIGINAL));
	}
	return strSel;
});

AddEnv("MenuSelected", function (Ctrl) {
	return g_.MenuSelected ? api.PathQuoteSpaces(g_.MenuSelected.Path) : "";
});

AddEnv("Installed", fso.GetDriveName(api.GetModuleFileName(null)));

AddEnv("TE_Config", function () {
	return BuildPath(te.Data.DataFolder, "config");
});

CreateUpdater = function (arg) {
	if (isFinite(RunEvent3("CreateUpdater", arg))) {
		return;
	}
	if (!IsExists(BuildPath(arg.temp, GetFileName(api.GetModuleFileName(null))))) {
		api.SHFileOperation(FO_MOVE, arg.temp + "\\*", GetParentFolderName(api.GetModuleFileName(null)), FOF_NOCONFIRMATION, false);
		ReloadCustomize();
		return;
	}
	g_.strUpdate = ['"', api.IsWow64Process(api.GetCurrentProcess()) ? wsh.ExpandEnvironmentStrings("%SystemRoot%\\Sysnative") : system32, "\\", "wscript.exe", '" "', arg.temp, "\\script\\update.js", '" "', api.GetModuleFileName(null), '" "', arg.temp, '" "', api.LoadString(hShell32, 12612), '" "', api.LoadString(hShell32, 12852), '"'].join("");
	DeleteTempFolder = PerformUpdate;
	WmiProcess("WHERE ExecutablePath='" + (api.GetModuleFileName(null).split("\\").join("\\\\")) + "' AND ProcessId!=" + arg.pid, function (item) {
		item.Terminate();
	});
	api.PostMessage(te.hwnd, WM_CLOSE, 0, 0);
}

IsHeader = function (Ctrl, pt, hwnd, strClass) {
	if (strClass == WC_HEADER) {
		return true;
	}
	if (strClass != "DirectUIHWND") {
		return false;
	}
	var iItem = Ctrl.HitTest(pt, LVHT_ONITEM);
	if (iItem >= 0) {
		return false;
	}
	var pt2 = pt.Clone();
	api.ScreenToClient(hwnd, pt2);
	return pt2.y < screen.deviceYDPI / 4;
}

AutocompleteThread = function () {
	var pid = api.ILCreateFromPath(path);
	if (!pid.IsFolder) {
		pid = api.ILCreateFromPath(api.CreateObject("fso").GetParentFolderName(path));
	}
	if (pid.IsFolder && pid.Path != Autocomplete.Path) {
		Autocomplete.Path = pid.Path;
		var Folder = pid.GetFolder;
		if (Folder) {
			var Items = Folder.Items();
			try {
				Items.Filter(fflag, "*");
			} catch (e) { }
			for (var i = 0; i < Items.Count; ++i) {
				if (Items.Item(i).IsFolder) {
					ar.push(Items.Item(i).Path);
				}
			}
		}
		api.Invoke(UI.Autocomplete, ar.join("\t"), pid.Path);
	}
}

AdjustAutocomplete = function (path) {
	if (te.Data.Conf_NoAutocomplete) {
		return;
	}
	var o = api.CreateObject("Object");
	o.Data = api.CreateObject("Object");
	o.Data.path = path;
	o.Data.Autocomplete = g_.Autocomplete;
	o.Data.fflag = SHCONTF_FOLDERS | ((te.Ctrl(CTRL_FV).ViewFlags & CDB2GVF_SHOWALLFILES) ? SHCONTF_INCLUDEHIDDEN : 0);
	o.Data.UI = UI;
	o.Data.api = api;
	o.Data.wsh = wsh;
	o.Data.ar = api.CreateObject("Array");
	api.ExecScript(AutocompleteThread.toString().replace(/^[^{]+{|}$/g, ""), "JScript", o, true);
}

FullscreenChanged = function (bFullscreen) {
	g_.Fullscreen = bFullscreen;
	if (bFullscreen) {
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
	RunEvent1("FullscreenChanged", bFullscreen);
}

ArrangeAddons1 = function (cl) {
	RunEvent1("Load");
	ClearEvent("Load");
	const cHwnd = [te.Ctrl(CTRL_WB).hwnd, te.hwnd];
	for (let i = cHwnd.length; i--;) {
		const hOld = api.SetClassLongPtr(cHwnd[i], GCLP_HBRBACKGROUND, api.CreateSolidBrush(cl));
		if (hOld) {
			api.DeleteObject(hOld);
		}
	}
	let hwnd, p = api.Memory("WCHAR", 11);
	p.Write(0, VT_LPWSTR, "ShellState");
	const cFV = te.Ctrls(CTRL_FV);
	for (let i in cFV) {
		if (hwnd = cFV[i].hwndView) {
			api.SendMessage(hwnd, WM_SETTINGCHANGE, 0, p);
		}
	}
}

InitCode = function () {
	const types = {
		Key: ["All", "List", "Tree", "Browser", "Edit", "Menus"],
		Mouse: ["All", "List", "List_Background", "Tree", "Tabs", "Tabs_Background", "Browser"]
	};
	for (let i = 0; i < 3; ++i) {
		g_.KeyState[i][0] = api.GetKeyNameText(g_.KeyState[i][0]);
	}
	let i = g_.KeyState.length;
	while (i-- > 4 && g_.KeyState[i][0] == g_.KeyState[i - 4][0]) {
		g_.KeyState.pop();
	}
	for (let j = 256; j >= 0; j -= 256) {
		for (let i = 128; i > 0; i--) {
			const v = api.GetKeyNameText((i + j) * 0x10000);
			if (v && v.charCodeAt(0) > 32) {
				g_.KeyCode[v.toUpperCase()] = i + j;
			}
		}
	}
	for (let i in types) {
		eventTE[i] = api.CreateObject("Object");
		for (let j in types[i]) {
			eventTE[i][types[i][j]] = api.CreateObject("Object");
		}
	}
	DefaultFont = api.Memory("LOGFONT");
	api.SystemParametersInfo(SPI_GETICONTITLELOGFONT, DefaultFont.Size, DefaultFont, 0);
	HOME_PATH = te.Data.Conf_NewTab || HOME_PATH;
	const cFV = te.Ctrls(CTRL_FV);
	for (let i in cFV) {
		cFV[i].ColumnsReplace = null;
	}
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
	te.DragIcon = !GetNum(te.Data.Conf_NoDragIcon);
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
		for (var i = menus.length; i--;) {
			items = menus[i].getElementsByTagName("Item");
			for (var j = api.ObjGetI(items, "length"); j--;) {
				a = items[j].getAttribute("Name").split(/\\t/);
				if (a.length > 1) {
					SetKeyExec("List", a[1], items[j].text, items[j].getAttribute("Type"), true);
				}
			}
		}
	}
}

InitWindow = function () {
	if (api.GetKeyState(VK_SHIFT) < 0 && api.GetKeyState(VK_CONTROL) < 0) {
		ShowOptions("Tab=Add-ons");
	}
	if (g_.xmlWindow && "string" !== typeof g_.xmlWindow) {
		LoadXml(g_.xmlWindow);
	}
	var cTC = te.Ctrls(CTRL_TC);
	if (cTC.Count == 0) {
		var TC = te.CreateCtrl(CTRL_TC, 0, 0, "100%", "100%", te.Data.Tab_Style, te.Data.Tab_Align, te.Data.Tab_TabWidth, te.Data.Tab_TabHeight);
		TC.Selected.Navigate2(HOME_PATH, SBSP_NEWBROWSER, te.Data.View_Type, te.Data.View_ViewMode, te.Data.View_fFlags, te.Data.View_Options, te.Data.View_ViewFlags, te.Data.View_IconSize, te.Data.Tree_Align, te.Data.Tree_Width, te.Data.Tree_Style, te.Data.Tree_EnumFlags, te.Data.Tree_RootStyle, te.Data.Tree_Root);
	} else if (te.Ctrls(CTRL_TC, true).Count == 0) {
		cTC[0].Visible = true;
	}
	g_.xmlWindow = void 0;
	setTimeout(function () {
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
			Resize();
			RunCommandLine(api.GetCommandLine());
		}, 500);
	}, 99);
}

Threads = api.CreateObject("Object");
Threads.Images = api.CreateObject("Array");
Threads.Data = api.CreateObject("Array");
Threads.nBase = 1;
Threads.nMax = 3;
Threads.nTI = 500;

Threads.GetImage = function (o) {
	o.OnGetAlt = te.OnGetAlt;
	Threads.Images.push(o);
	Threads.Run();
}

Threads.Run = function () {
	if (Threads.Data.length >= Threads.nMax) {
		return;
	}
	var tm = new Date().getTime();
	if (Threads.Data.length >= Threads.nBase && tm - Threads.Data[0].Data.tm < Threads.nTI) {
		return;
	}
	var o = api.CreateObject("Object");
	o.Data = api.CreateObject("Object");
	o.Data.Threads = Threads;
	o.Data.Id = Math.random();
	o.Data.tm = tm;
	o.Data.MainWindow = window;
	Threads.Data.unshift(o);
	if (!Threads.src) {
		Threads.src = ReadTextFile("script\\threads.js");
	}
	api.ExecScript(Threads.src, "JScript", o, true);
}

Threads.End = function (Id) {
	for (var i = Threads.Data.Count; i--;) {
		if (Id === Threads.Data[i].Data.Id) {
			Threads.Data.splice(i, 1);
			CollectGarbage();
			return;
		}
	}
	Threads.Data.pop();
	CollectGarbage();
}

Threads.Finalize = function () {
	Threads.GetImage = function () { };
	Threads.Images.Count = 0;
}

//Init

if (!te.Data) {
	te.Data = api.CreateObject("Object");
	te.Data.CustColors = api.Memory("int", 16);
	te.Data.AddonsData = api.CreateObject("Object");
	te.Data.Fonts = api.CreateObject("Object");
	te.Data.Exchange = api.CreateObject("Object");
	te.Data.rcWindow = api.Memory("RECT");
	//Default Value
	te.Data.Tab_Style = TCS_HOTTRACK | TCS_MULTILINE | TCS_RAGGEDRIGHT | TCS_SCROLLOPPOSITE | TCS_HOTTRACK | TCS_TOOLTIPS;
	te.Data.Tab_Align = TCA_TOP;
	te.Data.Tab_TabWidth = 96;
	te.Data.Tab_TabHeight = 0;

	te.Data.View_Type = CTRL_SB;
	te.Data.View_ViewMode = FVM_DETAILS;
	te.Data.View_fFlags = FWF_SHOWSELALWAYS;
	te.Data.View_Options = EBO_SHOWFRAMES | EBO_ALWAYSNAVIGATE;
	te.Data.View_ViewFlags = CDB2GVF_SHOWALLFILES;
	te.Data.View_IconSize = 0;

	te.Data.Tree_Align = 0;
	te.Data.Tree_Width = 200;
	te.Data.Tree_Style = NSTCS_HASEXPANDOS | NSTCS_SHOWSELECTIONALWAYS | NSTCS_HASLINES | NSTCS_BORDER | NSTCS_NOINFOTIP | NSTCS_HORIZONTALSCROLL;
	te.Data.Tree_EnumFlags = SHCONTF_FOLDERS;
	te.Data.Tree_RootStyle = NSTCRS_VISIBLE | NSTCRS_EXPANDED;
	te.Data.Tree_Root = 0;

	te.Data.Conf_TabDefault = true;
	te.Data.Conf_TreeDefault = true;
	te.Data.Conf_ListDefault = true;

	te.Data.Installed = GetParentFolderName(api.GetModuleFileName(null));
	te.Data.DataFolder = te.Data.Installed;
	te.Data.TempFolder = BuildPath(fso.GetSpecialFolder(2).Path, "tablacus");

	var fn = function () {
		te.Data.DataFolder = BuildPath(api.GetDisplayNameOf(ssfAPPDATA, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING), "Tablacus\\Explorer");
		var ParentFolder = GetParentFolderName(te.Data.DataFolder);
		if (!fso.FolderExists(ParentFolder)) {
			if (fso.CreateFolder(ParentFolder)) {
				CreateFolder2(te.Data.DataFolder);
			}
		}
	}

	var pf = [ssfPROGRAMFILES, ssfPROGRAMFILESx86];
	var x = api.sizeof("HANDLE") / 4;
	for (var i = 0; i < x; ++i) {
		var s = api.GetDisplayNameOf(pf[i], SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
		var l = s.replace(/\s*\(x86\)$/i, "").length;
		if (api.StrCmpNI(s, te.Data.DataFolder, l) == 0) {
			fn();
			break;
		}
	}
	var s = BuildPath(te.Data.DataFolder, "config");
	CreateFolder2(s);
	if (!fso.FolderExists(s)) {
		fn();
		CreateFolder2(BuildPath(te.Data.DataFolder, "config"));
	}
	delete fn;
	if (g_.IEVer < 8) {
		var s = BuildPath(te.Data.DataFolder, "cache");
		CreateFolder2(s);
		CreateFolder2(BuildPath(s, "bitmap"));
		CreateFolder2(BuildPath(s, "icon"));
		CreateFolder2(BuildPath(s, "file"));
	}
	te.Data.Conf_Lang = GetLangId();
	var SHIL = api.CreateObject("Array");
	var SHILS = api.CreateObject("Array");
	for (var i = 0; i < SHIL_JUMBO; ++i) {
		var il = api.SHGetImageList(i);
		SHIL.push(il);
		SHILS.push(api.Memory("SIZE"));
		api.ImageList_GetIconSize(SHIL[i], SHILS[i]);
	}
	te.Data.SHIL = SHIL;
	te.Data.SHILS = SHILS;
	var o = {};
	try {
		const sw = sha.Windows();
		for (let i = 0; i < sw.Count; ++i) {
			const x = sw.item(i);
			if (x && x.Document) {
				var w = x.Document.parentWindow;
				if (w && w.te && w.te.Data) {
					o[w.te.Data.WindowSetting] = 1;
				}
			}
		}
	} catch (e) { }
	api.SetWindowText(te.hwnd, TITLE);
	te.Data.WindowSetting = BuildPath(te.Data.DataFolder, "config\\window0.xml");
	for (let i = 1; i < 999; ++i) {
		const fn = BuildPath(te.Data.DataFolder, "config\\window" + i + ".xml");
		if (!o[fn]) {
			te.Data.WindowSetting = fn;
			break;
		}
	}
	if (api.GetKeyState(VK_SHIFT) < 0 && api.GetKeyState(VK_CONTROL) < 0) {
		g_.xmlWindow = "Init";
	} else {
		LoadConfig();
	}
	te.Data.uRegisterId = api.SHChangeNotifyRegister(te.hwnd, SHCNRF_InterruptLevel | SHCNRF_ShellLevel | SHCNRF_NewDelivery, SHCNE_ALLEVENTS & ~SHCNE_UPDATEIMAGE, TWM_CHANGENOTIFY, ssfDESKTOP, true);
} else {
	for (var i in te.Data) {
		if (/^xml/.test(i)) {
			delete te.Data[i];
		}
	}
	LoadConfig();
	delete g_.xmlWindow;
}
Exchange = te.Data.Exchange;

