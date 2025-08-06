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

g_.mouse = {
	str: "",
	CancelContextMenu: false,
	ptGesture: api.Memory("POINT"),
	ptDown: api.Memory("POINT"),
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
			const item = api.Memory("LVITEM");
			item.iItem = this.RButton;
			item.mask = LVIF_STATE;
			item.stateMask = LVIS_SELECTED;
			api.SendMessage(this.hwndGesture, LVM_GETITEM, 0, item);
			if (!(item.state & LVIS_SELECTED)) {
				if (mode) {
					const Ctrl = te.CtrlFromWindow(this.hwndGesture);
					Ctrl.SelectItem(this.RButton, SVSI_SELECT | SVSI_FOCUSED | SVSI_DESELECTOTHERS);
				} else {
					const ptc = te.Data.pt.Clone();
					api.ScreenToClient(this.hwndGesture, ptc);
					api.SendMessage(this.hwndGesture, WM_RBUTTONDOWN, 0, ptc.x + (ptc.y << 16));
				}
			}
		}
		this.RButton = -1;
	},

	GetButton: function (msg, wParam) {
		let s = "";
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
				if (api.GetClassName(api.WindowFromPoint(pt) || Ctrl.hwnd) == WC_TREEVIEW) {
					return;
				}
				return Ctrl.ItemCount && Ctrl.ItemCount(SVGIO_SELECTION) ? "List" : "List_Background";
			case CTRL_TV:
				return "Tree";
			case CTRL_TC:
				return pt.Target || Ctrl.HitTest(pt, TCHT_ONITEM) >= 0 ? "Tabs" : "Tabs_Background";
		}
	},

	Exec: function (Ctrl, hwnd, pt, str) {
		str = GetGestureKey() + (str || this.str);
		this.EndGesture(false);
		te.Data.cmdMouse = str;
		if (!Ctrl) {
			return S_FALSE;
		}
		const s = this.GetMode(Ctrl, pt);
		if (s) {
			if (GestureExec(Ctrl, s, str, pt, hwnd) === S_OK) {
				return S_OK;
			}
			const res = /(.*)_Background/i.exec(s);
			if (res) {
				if (GestureExec(Ctrl, res[1], str, pt, hwnd) === S_OK) {
					return S_OK;
				}
			}
		} else if (str.length) {
			window.g_menu_button = str;
			if (window.g_menu_click) {
				if (window.g_menu_click === true) {
					const hSubMenu = api.GetSubMenu(g_.menu_handle, g_.menu_pos);
					if (hSubMenu) {
						const rc = api.Memory("RECT");
						for (let i = g_.menu_pos; i >= 0; i--) {
							api.GetMenuItemRect(null, g_.menu_handle, i, rc);
							if (PtInRect(rc, pt)) {
								const mii = api.Memory("MENUITEMINFO");
								mii.fMask = MIIM_SUBMENU;
								if (api.SetMenuItemInfo(g_.menu_handle, g_.menu_pos, true, mii)) {
									api.DestroyMenu(hSubMenu);
								}
								break;
							}
						}
					}
				}
				if (str > 2) {
					window.g_menu_click = 4;
					const lParam = pt.x + (pt.y << 16);
					api.PostMessage(hwnd, WM_LBUTTONDOWN, 0, lParam);
					api.PostMessage(hwnd, WM_LBUTTONUP, 0, lParam);
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
		s = (s || "").replace(/&|\.\.\.$/g, "");
		return this.Func[s] || api.ObjGetI(this.Func, s);
	},

	CmdI: function (s, s2) {
		const type = this.FuncI(s);
		if (type) {
			return type.Cmd[s2] || api.ObjGetI(type.Cmd, s2);
		}
	},

	Func: {
		"": {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				const lines = s.split(/\r?\n/);
				for (let i in lines) {
					const cmd = lines[i].split(",");
					const Id = cmd.shift();
					const hr = Exec(Ctrl, cmd.join(","), Id, hwnd, pt);
					if (hr != S_OK) {
						break;
					}
					if (Ctrl.Type <= CTRL_EB) {
						Ctrl = GetFolderView();
					}
				}
				return S_OK;
			},

			Ref: function (s, pt) {
				const lines = s.split(/\r?\n/);
				const last = lines.length ? lines[lines.length - 1] : "";
				const res = /^([^,]+),$/.exec(last);
				if (res) {
					const Id = GetSourceText(res[1]);
					const r = OptionRef(Id, "", pt);
					if ("string" === typeof r) {
						return s + r + "\n";
					}
					return r;
				} else {
					const arFunc = [];
					RunEvent1("AddType", arFunc);
					const r = g_basic.Popup(arFunc, s, pt);
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
				const FV = GetFolderView(Ctrl, pt);
				if (FV) {
					s = ExtractMacro(Ctrl, s);
					SetFilterView(FV, s != "*" ? s : null);
				}
				return S_OK;
			}
		},

		Exec: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				const lines = ExtractMacro(Ctrl, s).split(/\r?\n/);
				for (let i in lines) {
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
				const re = /%Selected%/i;
				if (re.test(s)) {
					pdwEffect[0] = DROPEFFECT_LINK;
					if (bDrop) {
						const ar = [];
						for (let i = dataObj.Count; i > 0; ar.unshift(PathQuoteSpaces(dataObj.Item(--i).Path))) {
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
				const fn = g_basic.CmdI(type, s);
				if (fn) {
					return fn(Ctrl, pt);
				} else {
					return g_basic.Func.Exec.Exec(Ctrl, s + " %Selected%", "Exec", hwnd, pt);
				}
			},

			Drop: function (Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop) {
				const fn = g_basic.CmdI(type, s);
				if (fn) {
					pdwEffect[0] = DROPEFFECT_NONE;
					return S_OK;
				}
				return g_basic.Func.Exec.Drop(Ctrl, s + " %Selected%", "Exec", hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop)
			},

			Ref: function (s, pt) {
				let r = g_basic.Popup(g_basic.Func["Selected items"].Cmd, s, pt);
				if (SameText(r, GetText("Send to..."))) {
					const Folder = sha.NameSpace(ssfSENDTO);
					if (Folder) {
						const Items = Folder.Items();
						const hMenu = api.CreatePopupMenu();
						for (i = 0; i < Items.Count; ++i) {
							api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, i + 1, Items.Item(i).Name);
						}
						const nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD | (pt.width ? TPM_RIGHTALIGN : 0), pt.x + pt.width, pt.y, te.hwnd, null, null);
						api.DestroyMenu(hMenu);
						if (nVerb) {
							return PathQuoteSpaces(Items.Item(nVerb - 1).Path);
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
					const Selected = GetSelectedItems(Ctrl, pt);
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
				const fn = g_basic.CmdI(type, s);
				if (fn) {
					fn(Ctrl, pt);
					return S_OK;
				}
				const FV = GetFolderView(Ctrl, pt, true);
				if (FV) {
					const res = /^(\-?)(\d+)/.exec(s);
					if (res) {
						const TC = FV.Parent;
						TC.SelectedIndex = res[1] ? TC.Count - res[2] : res[2];
					}
				}
				return S_OK;
			},

			Cmd: {
				"Close tab": function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt, true);
					FV && FV.Close();
				},
				"Close other tabs": function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt);
					if (FV) {
						const TC = FV.Parent;
						const nIndex = GetFolderView(Ctrl, pt, true) ? FV.Index : -1;
						TC.LockUpdate();
						try {
							for (let i = TC.Count; i--;) {
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
					const FV = GetFolderView(Ctrl, pt, true);
					if (FV) {
						const TC = FV.Parent;
						TC.LockUpdate();
						try {
							for (let i = FV.Index; i--;) {
								TC[i].Close();
							}
						} catch (e) {
							ShowError(e);
						}
						TC.UnlockUpdate();
					}
				},
				"Close tabs on right": function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt, true);
					if (FV) {
						const TC = FV.Parent;
						const nIndex = FV.Index;
						TC.LockUpdate();
						try {
							for (let i = TC.Count; --i > nIndex;) {
								TC[i].Close();
							}
						} catch (e) {
							ShowError(e);
						}
						TC.UnlockUpdate();
					}
				},
				"Close all tabs": function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt);
					if (FV) {
						const TC = FV.Parent;
						TC.LockUpdate();
						try {
							for (let i = TC.Count; i--;) {
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
					const FV = GetFolderView(Ctrl, pt);
					FV && Lock(FV.Parent, FV.Index, true);
				},
				"Previous tab": function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt);
					FV && ChangeTab(FV.Parent, -1);
				},
				"Next tab": function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt);
					FV && ChangeTab(FV.Parent, 1);
				},
				Up: function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt);
					FV && FV.Navigate(null, SBSP_PARENT | GetNavigateFlags(FV, true));
				},
				Back: function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt);
					if (FV) {
						const wFlags = GetNavigateFlags(FV);
						if (wFlags & SBSP_NEWBROWSER) {
							const Log = FV.History;
							if (Log && Log.Index < Log.Count - 1) {
								FV.Navigate(Log[Log.Index + 1], wFlags);
							}
						} else {
							FV.Navigate(null, SBSP_NAVIGATEBACK);
						}
					}
				},
				Forward: function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt);
					if (FV) {
						const wFlags = GetNavigateFlags(FV);
						if (wFlags & SBSP_NEWBROWSER) {
							const Log = FV.History;
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
					const FV = GetFolderView(Ctrl, pt);
					if (FV) {
						FV.Type = (FV.Type == CTRL_SB) ? CTRL_EB : CTRL_SB;
					}
				},
				"Switch Explorer engine": function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt);
					if (FV) {
						FV.Type = (FV.Type == CTRL_SB) ? CTRL_EB : CTRL_SB;
					}
				},
				"Open in Explorer": function (Ctrl, pt) {
					OpenInExplorer(GetFolderView(Ctrl, pt));
				}
			},

			Enc: true
		},

		Edit: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				const hMenu = te.MainMenu(FCIDM_MENU_EDIT);
				const nVerb = GetCommandId(hMenu, s);
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
				const hMenu = te.MainMenu(FCIDM_MENU_VIEW);
				const nVerb = GetCommandId(hMenu, s);
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
				let Selected;
				let FV = Ctrl;
				if (Ctrl.Type <= CTRL_EB || Ctrl.Type == CTRL_TV) {
					Selected = Ctrl.SelectedItems();
				} else {
					FV = te.Ctrl(CTRL_FV);
					Selected = FV.SelectedItems();
				}
				if (Selected && Selected.Count) {
					const ContextMenu = api.ContextMenu(Selected, FV);
					if (ContextMenu) {
						const hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS | CMF_CANRENAME);
						const nVerb = GetCommandId(hMenu, s, ContextMenu);
						FV.Focus();
						ContextMenu.InvokeCommand(0, te.hwnd, nVerb ? nVerb - 1 : s, null, null, SW_SHOWNORMAL, 0, 0);
						api.DestroyMenu(hMenu);
					}
				}
				return S_OK;
			},

			Ref: function (s, pt) {
				const FV = te.Ctrl(CTRL_FV);
				if (FV) {
					let Selected = FV.SelectedItems();
					if (!Selected.Count) {
						Selected = api.GetModuleFileName(null);
					}
					const ContextMenu = api.ContextMenu(Selected, FV);
					if (ContextMenu) {
						const hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS | CMF_CANRENAME);
						return g_basic.PopupMenu(hMenu, ContextMenu, pt);
					}
				}
			}
		},

		Background: {
			Exec: function (Ctrl, s, type, hwnd, pt) {
				const FV = te.Ctrl(CTRL_FV);
				if (FV) {
					const ContextMenu = FV.ViewMenu();
					if (ContextMenu) {
						FV.Focus();
						const hMenu = api.CreatePopupMenu();
						ContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_EXTENDEDVERBS);
						const nVerb = GetCommandId(hMenu, s, ContextMenu);
						ContextMenu.InvokeCommand(0, te.hwnd, nVerb ? nVerb - 1 : s, null, null, SW_SHOWNORMAL, 0, 0);
						api.DestroyMenu(hMenu);
					}
				}
				return S_OK;
			},

			Ref: function (s, pt) {
				const FV = te.Ctrl(CTRL_FV);
				if (FV) {
					const ContextMenu = FV.ViewMenu();
					if (ContextMenu) {
						const hMenu = api.CreatePopupMenu();
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
					const Selected = GetSelectedItems(Ctrl, pt);
					let s = "";
					const nCount = Selected.Count;
					if (nCount) {
						s = te.OnClipboardText(Selected);
					} else {
						s = PathQuoteSpaces(api.GetDisplayNameOf(FV, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING));
					}
					api.SetClipboardData(s);
					return S_OK;
				},
				"Run dialog": function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt);
					api.ShRunDialog(te.hwnd, 0, FV ? FV.FolderItem.Path : null, null, null, 0);
					return S_OK;
				},
				Search: function (Ctrl, pt) {
					const FV = GetFolderView(Ctrl, pt);
					if (FV && FV.FolderItem && FV.FolderItem.IsFolder) {
						const res = IsSearchPath(FV.FolderItem, true);
						InputDialog("Search", res ? unescape(res[1]) : "", function (r) {
							if (r) {
								FV.Search(r);
							} else if (r === "") {
								CancelFilterView(FV);
							}
						});
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
		const fn = g_basic.CmdI(type, s);
		if (!pt) {
			pt = api.GetCursorPos();
		}
		if (fn) {
			InvokeFunc(fn, [Ctrl, pt]);
		}
		return S_OK;
	},

	Popup: function (Cmd, strDefault, pt) {
		let i, s;
		let ar = [];
		if (Cmd.length) {
			ar = Cmd;
		} else {
			for (i in Cmd) {
				ar.push(i);
			}
		}
		const hMenu = api.CreatePopupMenu();
		for (i = 0; i < ar.length; ++i) {
			if (ar[i]) {
				api.InsertMenu(hMenu, MAXINT, MF_BYPOSITION | MF_STRING, i + 1, GetTextR(ar[i]));
			}
		}
		const nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD | (pt.width ? TPM_RIGHTALIGN : 0), pt.x + pt.width, pt.y, te.hwnd, null, null);
		s = api.GetMenuString(hMenu, nVerb, MF_BYCOMMAND);
		api.DestroyMenu(hMenu);
		if (nVerb == 0) {
			return 1;
		}
		return s;
	},

	PopupMenu: function (hMenu, ContextMenu, pt) {
		let Verb;
		for (let i = api.GetMenuItemCount(hMenu); i--;) {
			if (api.GetMenuString(hMenu, i, MF_BYPOSITION)) {
				api.EnableMenuItem(hMenu, i, MF_ENABLED | MF_BYPOSITION);
			} else {
				api.DeleteMenu(hMenu, i, MF_BYPOSITION);
			}
		}
		window.g_menu_click = true;
		const nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD | (pt.width ? TPM_RIGHTALIGN : 0), pt.x + pt.width, pt.y, te.hwnd, null, ContextMenu);
		if (nVerb == 0) {
			api.DestroyMenu(hMenu);
			return 1;
		}
		if (ContextMenu) {
			Verb = ContextMenu.GetCommandString(nVerb - 1, GCS_VERB);
		}
		if (!Verb) {
			Verb = window.g_menu_string;
			const res = /\t(.*)$/.exec(Verb);
			if (res) {
				Verb = res[1];
			}
		}
		api.DestroyMenu(hMenu);
		return Verb;
	}
};

RefreshEx = function (FV, tm, df) {
	if (FV.Data && FV.FolderItem && IsWitness(FV.FolderItem) && /^[A-Z]:\\|^\\\\\w/i.test(FV.FolderItem.Path)) {
		if (RunEvent4("RefreshEx", FV, tm, df) == null) {
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
									if (RunEvent4("Error", FV) == null) {
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
	if (TC && !g_.LockUpdate && Ctrl && Ctrl.FolderItem && Ctrl.FolderItem.Path != null && te.OnArrange) {
		ChangeTabName(Ctrl);
		RunEvent1("ChangeView", Ctrl);
		if (Ctrl.Id == Ctrl.Parent.Selected.Id) {
			if (Ctrl.hwndView) {
				const tm = Math.max(5000, te.Data.Conf_NetworkTimeout);
				RefreshEx(Ctrl, tm, tm);
			}
			if (Ctrl.Parent.Id == TC.Id) {
				RunEvent1("ChangeView1", Ctrl);
				Ctrl.Focus();
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
		return RunEvent4("GetTabName", Ctrl) || GetFolderItemName(Ctrl.FolderItem);
	}
}

GetFolderItemName = function (pid) {
	return pid ? RunEvent4("GetFolderItemName", pid) || api.GetDisplayNameOf(pid, SHGDN_INFOLDER) : "";
}

IsUseExplorer = function (pid) {
	return RunEvent3("UseExplorer", pid);
}

CloseView = function (Ctrl) {
	return RunEvent2("CloseView", Ctrl);
}

DeviceChanged = function (Ctrl, wParam, lParam) {
	RunEvent1("DeviceChanged", Ctrl, wParam, lParam);
}

ListViewCreated = function (Ctrl) {
	ChangeTabName(Ctrl);
	Ctrl.Data.AccessTime = "#";
	RunEvent1("ListViewCreated", Ctrl);
	te.OnFilterChanged(Ctrl);
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
	if (g_.bFinalized) {
		return;
	}
	g_.bFinalized = true;
	RunEvent1("Finalize");
	FinalizeEx();
}

FinalizeEx = function () {
	SaveConfig();
	Threads.Finalize();

	for (let i in g_.dlgs) {
		const dlg = g_.dlgs[i];
		try {
			if (dlg) {
				if (dlg.oExec) {
					if (dlg.oExec.Status == 0) {
						dlg.oExec.Terminate();
					}
				} else if (dlg.Document) {
					dlg.CloseWindow();
				}
			}
		} catch (e) { }
	}
}

SetGestureText = function (Ctrl, Text) {
	const mode = g_.mouse.GetMode(Ctrl, g_.mouse.ptGesture);
	if (mode) {
		let s = eventTE.Mouse[mode][Text];
		if (!s) {
			const res = /(.*)_Background/i.exec(mode);
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
		const hdc = api.GetWindowDC(te.hwnd);
		if (hdc) {
			const rc = api.Memory("RECT");
			if (!g_.mouse.ptText) {
				g_.mouse.ptText = g_.mouse.ptGesture.Clone();
				api.ScreenToClient(te.hwnd, g_.mouse.ptText);
				g_.mouse.right = g_.mouse.bottom = -32767;
			}
			rc.left = g_.mouse.ptText.x;
			rc.top = g_.mouse.ptText.y;
			const hOld = api.SelectObject(hdc, CreateFont(DefaultFont));
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
			if (InvokeFunc(eo[i], args) === false) {
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
		RunEvent1("Lock", Ctrl, nIndex, GetLock(FV));
	}
}

GetLock = function (FV) {
	return FV && FV.Data && FV.Data.Lock;
}

CanClose = function (Ctrl) {
	if (Ctrl && Ctrl.Data) {
		switch (Ctrl.Type) {
			case CTRL_TE:
				if (api.GetThreadCount()) {
					return S_FALSE;
				}
				let hwnd1 = null;
				const pid = api.Memory("DWORD");
				api.GetWindowThreadProcessId(te.hwnd, pid);
				const pid0 = pid[0];
				while (hwnd1 = api.FindWindowEx(null, hwnd1, null, null)) {
					if (!/tablacus|_hidden/i.test(api.GetClassName(hwnd1)) && api.IsWindowVisible(hwnd1)) {
						api.GetWindowThreadProcessId(hwnd1, pid);
						if (pid[0] == pid0) {
							return S_FALSE;
						}
					}
				}
				break;
			case CTRL_SB:
			case CTRL_EB:
				if (GetLock(Ctrl)) {
					return S_FALSE;
				}
				break;
		}
		return RunEvent2("CanClose", Ctrl) || S_OK;
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
	const xml = OpenXml("window.xml", true, false);
	if (xml) {
		const items = xml.getElementsByTagName('Data');
		if (items.length && !bDog) {
			let path = items[0].getAttribute("Path");
			if (path) {
				path = ExtractMacro(te, path);
				if (api.PathIsDirectory(BuildPath(path, "config"))) {
					te.Data.DataFolder = path;
					LoadConfig(true);
					return;
				}
			}
		}
		const arKey = ["Conf", "Tab", "Tree", "View"];
		for (let j in arKey) {
			const key = arKey[j];
			let items = xml.getElementsByTagName(key);
			if (items.length == 0 && j == 0) {
				items = xml.getElementsByTagName('Config');
			}
			for (let i = 0; i < items.length; ++i) {
				const item = items[i];
				let s = item.text;
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
	if (api.PathFileExists(te.Data.WindowSetting)) {
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
		const items = xml.getElementsByTagName('Window');
		if (items.length) {
			const item = items[0];
			te.CmdShow = item.getAttribute("CmdShow");
			let x = GetNum(item.getAttribute("Left"));
			let y = GetNum(item.getAttribute("Top"));
			if (x > -30000 && y > -30000) {
				let w = GetNum(item.getAttribute("Width"));
				let h = GetNum(item.getAttribute("Height"));
				const z = GetNum(item.getAttribute("DPI")) / screen.deviceYDPI;
				if (z && z != 1) {
					x /= z;
					y /= z;
					w /= z;
					h /= z;
				}
				const pt = { x: x + w / 2, y: y };
				if (!api.MonitorFromPoint(pt, MONITOR_DEFAULTTONULL)) {
					const hMonitor = api.MonitorFromPoint(pt, MONITOR_DEFAULTTOPRIMARY);
					const mi = api.Memory("MONITORINFOEX");
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
		const root = te.Data.Addons.documentElement;
		if (root) {
			const items = root.childNodes;
			if (items) {
				const nLen = items.length;
				for (let i = 0; i < nLen; ++i) {
					const item = items[i];
					const Id = item.nodeName;
					if (Addons[Id] == null) {
						Addons[Id] = item.getAttribute("Enabled");
					}
				}
			}
		}
	}
	te.Data.bErrorAddons = false;
	const xml = CreateXml(true);
	for (let Id in Addons) {
		try {
			const items = te.Data.Addons.getElementsByTagName(Id);
			const item = items.length ? items[0].cloneNode(true) : xml.createElement(Id);
			let Enabled = Addons[Id];
			if (Enabled) {
				const AddonFolder = BuildPath(te.Data.Installed, "addons", Id);
				Enabled = api.PathIsDirectory(AddonFolder + "\\lang") ? 2 : 0;
				if (api.PathFileExists(AddonFolder + "\\script.vbs")) {
					Enabled |= 8;
				}
				if (api.PathFileExists(AddonFolder + "\\script.js")) {
					Enabled |= 1;
				}
				Enabled = (Enabled & 9) ? Enabled : 4;
			} else {
				AddonDisabled(Id);
			}
			item.setAttribute("Enabled", Enabled);
			const info = GetAddonInfo(Id);
			if (info.Level > 0) {
				item.setAttribute("Level", info.Level);
			}
			xml.documentElement.appendChild(item);
		} catch (e) { }
	}
	te.Data.Addons = xml;
	RunEvent1("ConfigChanged", "Addons");
}

AddonDisabled = function (Id) {
	RunEvent1("AddonDisabled", Id);
	if (eventTE.addondisabledex) {
		Id = Id.toLowerCase();
		if (Id == "tabgroups") {
			return;
		}
		const fn = eventTE.addondisabledex[Id];
		if (fn) {
			delete eventTE.addondisabledex[Id];
			AddEvent("Finalize", fn);
		}
	}
	CollectGarbage();
}

AddEvent("Refresh", function (Ctrl, pt) {
	const FV = GetFolderView(Ctrl, pt);
	if (FV) {
		FV.Focus();
		FV.Refresh();
	}
});

AddFavorite = function (FolderItem) {
	const xml = te.Data.xmlMenus;
	const menus = xml.getElementsByTagName("Favorites");
	if (menus && menus.length) {
		const item = xml.createElement("Item");
		if (!FolderItem) {
			const FV = te.Ctrl(CTRL_FV);
			if (FV) {
				FolderItem = FV.FolderItem;
			}
		}
		if (!FolderItem) {
			return;
		}
		InputDialog("Add Favorite", GetFolderItemName(FolderItem), function (s) {
			if (s) {
				item.setAttribute("Name", s.replace(/\\/g, "/"));
				item.setAttribute("Filter", "");
				let path = api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
				if ("string" === typeof path) {
					path = FolderItem.Path;
				}
				if (!FolderItem.Enum && api.PathFileExists(path) && !api.PathIsDirectory(path)) {
					path = PathQuoteSpaces(path);
					item.setAttribute("Type", "Exec");
				} else {
					item.setAttribute("Type", "Open");
				}
				item.text = path;
				menus[0].appendChild(item);
				SaveXmlEx("menus.xml", xml);
				FavoriteChanged();
			}
		});
	}
}

SetFilterView = function (FV, s) {
	if (SameText(s, FV.FilterView)) {
		return;
	}
	FV.FilterView = s || null;
	if (IsSearchPath(FV)) {
		return;
	}
	if (s) {
		FV.Refresh();
		return;
	}
	setTimeout(function () {
		FV.Refresh();
	}, 99);
}

CancelFilterView = function (FV) {
	if (IsSearchPath(FV)) {
		FV.Navigate(null, SBSP_PARENT);
		return S_OK;
	}
	if (FV.FilterView) {
		SetFilterView(FV);
		return S_OK;
	}
	return S_FALSE;
}

GetCommandId = function (hMenu, s, ContextMenu) {
	const arMenu = [hMenu];
	let wId = 0;
	if (s) {
		const sl = GetTextR(s);
		const mii = api.Memory("MENUITEMINFO");
		mii.fMask = MIIM_SUBMENU | MIIM_ID | MIIM_FTYPE | MIIM_STATE;
		while (arMenu.length) {
			hMenu = arMenu.pop();
			if (ContextMenu) {
				ContextMenu.HandleMenuMsg(WM_INITMENUPOPUP, hMenu, 0);
			}
			let i = api.GetMenuItemCount(hMenu);
			if (i == 1 && ContextMenu && WINVER >= 0x600) {
				setTimeout(function () {
					api.EndMenu();
				}, 200);
				api.TrackPopupMenuEx(hMenu, TPM_RETURNCMD, 0, 0, te.hwnd, null, ContextMenu);
				i = api.GetMenuItemCount(hMenu);
			}
			while (i-- > 0) {
				if (api.GetMenuItemInfo(hMenu, i, true, mii) && !(mii.fType & MFT_SEPARATOR) && !(mii.fState & MFS_DISABLED)) {
					const title = api.GetMenuString(hMenu, i, MF_BYPOSITION);
					if (title) {
						const a = title.split(/\t/);
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
	const ar = GetSelectedArray(Ctrl, pt);
	let Selected = ar[0];
	const FV = ar[2];
	if (Selected) {
		const Exec = [];
		if (Selected.Count > 1 && (NewTab & g_.OpenReverse)) {
			const n = api.ILIsEqual(FV, "about:blank") ? 1 : 0;
			if (n) {
				OpenSelected1(FV, Selected, 0, SBSP_SAMEBROWSER, Exec);
			}
			for (let i = Selected.Count; i-- > n;) {
				OpenSelected1(FV, Selected, i, NewTab, Exec);
			}
		} else {
			for (let i = 0; i < Selected.Count; ++i) {
				OpenSelected1(FV, Selected, i, NewTab, Exec);
			}
		}
		if (Exec.length) {
			if (Selected.Count != Exec.length) {
				Selected = api.CreateObject("FolderItems");
				for (let i = 0; i < Exec.length; ++i) {
					Selected.AddItem(Exec[i]);
				}
			}
			InvokeCommand(Selected, 0, te.hwnd, null, null, null, SW_SHOWNORMAL, 0, 0, FV, CMF_DEFAULTONLY);
		}
	}
	return S_OK;
}

OpenSelected1 = function (FV, Selected, i, NewTab, Exec) {
	const Item = Selected.Item(i);
	let bFolder = Item.IsFolder || (!Item.IsFileSystem && Item.IsBrowsable);
	if (!bFolder) {
		const path = Item.ExtendedProperty("linktarget");
		if (path) {
			bFolder = api.PathIsDirectory(path) === true;
		}
	}
	if (bFolder) {
		FV.Navigate(Item, NewTab);
		NewTab |= SBSP_NEWBROWSER;
	} else {
		if (NewTab & g_.OpenReverse) {
			Exec.unshift(Item);
		} else {
			Exec.push(Item);
		}
	}
}

SendCommand = function (Ctrl, nVerb) {
	if (nVerb) {
		let hwnd = Ctrl.hwndView;
		if (!hwnd) {
			const FV = te.Ctrl(CTRL_FV);
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

GetDropTargetItem = function (Ctrl, hList, pt) {
	const Dest = Ctrl.HitTest(pt);
	if (Dest) {
		if (hList) {
			if (api.SendMessage(hList, LVM_GETITEMSTATE, Ctrl.HitTest(pt, LVHT_ONITEM), LVIS_DROPHILITED)) {
				return Dest;
			}
			return Ctrl.FolderItem;
		}
		if (!api.PathIsDirectory(Dest.Path)) {
			if (api.DropTarget(Dest)) {
				return;
			}
			return Ctrl.FolderItem;
		}
		return Dest;
	}
	return Ctrl.FolderItem;
}

SetFolderViewData = function (FV, FVD) {
	if (FV && FVD) {
		const t = FV.Type;
		for (let i in FVD) {
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
	const TV = FV.TreeView;
	if (TV && TVD) {
		TV.Align = TVD.Align;
		TV.Style = TVD.Style;
		TV.Width = TVD.Width;
		TV.SetRoot(TVD.Root, TVD.EnumFlags, TVD.RootStyle);
	}
}

FixIconSpacing = function (Ctrl) {
	const hwnd = Ctrl.hwndList;
	if (hwnd) {
		if (api.SendMessage(hwnd, LVM_GETVIEW, 0, 0) == 0) {
			const s = Ctrl.IconSize;
			const cx = s * 96 / screen.deviceYDPI + (10 + (255 - s) / 10) * screen.deviceYDPI / 96;
			const cz = s < 96 ? (s - 96) / 5 : (s > 96 ? (96 - s) / 9 : 0);
			const cy = s * 96 / screen.deviceYDPI + (46 + cz) * screen.deviceYDPI / 96;
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
			if (CanClose(Ctrl) && api.GetKeyState(VK_SHIFT) >= 0) {
				setTimeout(function () {
					api.PostMessage(te.hwnd, WM_CLOSE, 0, 0);
				}, 999);
				return S_FALSE;
			}
			SetWindowAlpha(te.hwnd, 0);
			Finalize();
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
		clearTimeout(g_.tid_rf[Ctrl.Id]);
		delete g_.tid_rf[Ctrl.Id];
	}
	if (Ctrl.Data) {
		Ctrl.Data.Setting = void 0;
		Ctrl.Data.hwndBefore = api.GetFocus();
	}
	const res = /javascript:(.*)/im.exec(api.GetDisplayNameOf(Ctrl.FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING));
	if (res) {
		try {
			new AsyncFunction(res[1])(Ctrl);
		} catch (e) {
			ShowError(e, res[1]);
		}
		return E_NOTIMPL;
	}
	let hr = RunEvent2("BeforeNavigate", Ctrl, fs, wFlags, Prev);
	if (hr == S_OK && IsUseExplorer(Ctrl.FolderItem)) {
		setTimeout(OpenInExplorer, 99, Ctrl.FolderItem);
		return E_FAIL;
	}
	if (GetLock(Ctrl) || api.GetKeyState(VK_MBUTTON) < 0 || api.GetKeyState(VK_CONTROL) < 0) {
		if ((wFlags & SBSP_NEWBROWSER) == 0 && !api.ILIsEqual(Prev, "about:blank") && !api.ILIsEqual(Ctrl.FolderItem, "about:blank")  && !api.ILIsEqual(Ctrl.FolderItem, Prev)) {
			hr = E_ACCESSDENIED;
		}
	}
	return hr;
}

te.OnNavigateComplete = function (Ctrl) {
	if (g_.tid_rf[Ctrl.Id] || !Ctrl.FolderItem || !Ctrl.Data) {
		return S_OK;
	}
	Ctrl.Data.AccessTime = new Date().getTime();
	if (IsWitness(Ctrl.FolderItem)) {
		Ctrl.FolderFlags &= ~FWF_NOENUMREFRESH;
	} else {
		Ctrl.FolderFlags |= FWF_NOENUMREFRESH;
	}
	Ctrl.NavigateComplete();
	RunEvent1("NavigateComplete", Ctrl);
	ChangeView(Ctrl);
	if (api.GetClassName(Ctrl.Data.hwndBefore) == WC_TREEVIEW) {
		api.SetFocus(Ctrl.Data.hwndBefore);
	} else {
		FocusFV();
	}
	Ctrl.Data.hwndBefore = void 0;

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
	const hr = RunEvent3("KeyMessage", Ctrl, hwnd, msg, key, keydata);
	if (isFinite(hr)) {
		return hr;
	}
	if (g_.mouse.str.length > 1) {
		SetGestureText(Ctrl, GetGestureKey() + g_.mouse.str);
	}
	if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN) {
		const nKey = (((keydata >> 16) & 0x17f) || (api.MapVirtualKey(key, 0) | ((key >= 33 && key <= 46 || key >= 91 && key <= 93 || key == 111 || key == 144) ? 256 : 0))) | GetKeyShift();
		if (nKey == 0x15d) {
			g_.mouse.CancelContextMenu = false;
		}
		te.Data.cmdKey = nKey;
		te.Data.cmdKeyF = true;
		let strClass;
		switch (Ctrl.Type) {
			case CTRL_SB:
			case CTRL_EB:
				strClass = api.GetClassName(hwnd);
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
						const nCount = Ctrl.ItemCount(SVGIO_ALLVIEW);
						Ctrl.SelectItem((Ctrl.GetFocusedItem + (api.GetKeyState(VK_SHIFT) < 0 ? -1 : 1) + nCount) % nCount, SVSI_EDIT | SVSI_FOCUSED | SVSI_SELECT | SVSI_DESELECTOTHERS);
						setTimeout(function () {
							api.SetFocus(hwnd);
						}, 500);
						return S_OK;
					}
					return S_FALSE;
				}
				break;
			case CTRL_TV:
				strClass = api.GetClassName(hwnd);
				if (api.PathMatchSpec(strClass, WC_TREEVIEW)) {
					if (KeyExecEx(Ctrl, "Tree", nKey, hwnd) === S_OK) {
						return S_OK;
					}
					let nCmd;
					if (key == VK_DELETE) {
						nCmd = CommandID_DELETE;
					} else if (nKey == 0x202d) {
						nCmd = CommandID_CUT;
					} else if (nKey == 0x202e) {
						nCmd = CommandID_COPY;
					} else if (nKey == 0x202f) {
						nCmd = CommandID_PASTE;
					}
					if (nCmd) {
						InvokeCommand(Ctrl.SelectedItem, 0, te.hwnd, nCmd - 1, null, null, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
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
			default:
				if (key == VK_RETURN && window.g_menu_click === true) {
					const hSubMenu = api.GetSubMenu(g_.menu_handle, g_.menu_pos);
					if (hSubMenu) {
						const mii = api.Memory("MENUITEMINFO");
						mii.fMask = MIIM_SUBMENU;
						api.SetMenuItemInfo(g_.menu_handle, g_.menu_pos, true, mii);
						api.DestroyMenu(hSubMenu);
						api.PostMessage(hwnd, WM_CHAR, VK_LBUTTON, 0);
					}
				}
				if (g_.menu_loop && key == VK_TAB) {
					const wParam = api.GetKeyState(VK_SHIFT) < 0 ? VK_UP : VK_DOWN;
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
		if (Ctrl.Type != CTRL_TE && Ctrl.Type != CTRL_WB) {
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
		const pt2 = pt.Clone();
		api.ScreenToClient(te.hwnd, pt2);
		if (pt2.x < 1) {
			pt2.x = 1;
		}
		InvokeUI("MoveSplitter", [pt2.x, g_.mouse.Capture]);
		return S_OK;
	}
	if (msg != WM_MOUSEMOVE) {
		te.Data.cmdKeyF = false;
		if (!/^Chrome|^Internet/i.test(api.GetClassName(hwnd))) {
			if (/^Chrome|^Internet/i.test(api.GetClassName(api.GetFocus()))) {
				InvokeUI("FocusElement");
			}
		}
	}
	const hr = RunEvent3("MouseMessage", Ctrl, hwnd, msg, wParam, pt);
	if (isFinite(hr)) {
		return hr;
	}
	const strClass = api.GetClassName(hwnd);
	if (strClass == WC_EDIT) {
		return S_FALSE;
	}
	const bLV = Ctrl.Type <= CTRL_EB && api.PathMatchSpec(strClass, WC_LISTVIEW + ";DirectUIHWND");
	if (msg == WM_MOUSEWHEEL) {
		const Ctrl2 = te.CtrlFromPoint(pt);
		if (Ctrl2) {
			g_.mouse.str = GetGestureButton() + (wParam > 0 ? "8" : "9");
			if (api.GetKeyState(VK_RBUTTON) < 0) {
				g_.mouse.CancelContextMenu = true;
			}
			if (g_.mouse.Exec(Ctrl2, hwnd, pt) == S_OK) {
				return S_OK;
			}
			const hwnd2 = api.WindowFromPoint(pt);
			if (hwnd2 && hwnd != hwnd2) {
				api.SetFocus(hwnd2);
				api.SendMessage(hwnd2, msg, wParam & 0xffff0000, pt.x + (pt.y << 16));
				return S_OK;
			}
		}
	}
	if (msg == WM_LBUTTONUP || msg == WM_RBUTTONUP || msg == WM_MBUTTONUP || msg == WM_XBUTTONUP) {
		if (g_.mouse.str.length) {
			let hr = S_FALSE;
			let bButton = false;
			if (msg == WM_RBUTTONUP) {
				if (g_.mouse.RButton >= 0) {
					g_.mouse.RButtonDown(true);
					bButton = (g_.mouse.str == "2");
				} else if (bLV) {
					const iItem = Ctrl.HitTest(pt, LVHT_ONITEM);
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
		const s = g_.mouse.GetButton(msg, wParam);
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
				const iItem = Ctrl.HitTest(pt, LVHT_ONITEM);
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
				let iItem = -1;
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
			const iItem = Ctrl.HitTest(pt, LVHT_ONITEM);
			g_.mouse.Select = iItem >= 0 ? Ctrl.Items().Item(iItem) : null;
			g_.mouse.Deselect = Ctrl;
		}
	}
	if (msg == WM_LBUTTONDBLCLK || msg == WM_RBUTTONDBLCLK || msg == WM_MBUTTONDBLCLK || msg == WM_XBUTTONDBLCLK) {
		if (!IsHeader(Ctrl, pt, hwnd, strClass)) {
			const tm = new Date().getTime();
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
					const hSubMenu = api.GetSubMenu(g_.menu_handle, g_.menu_pos);
					if (hSubMenu) {
						const mii = api.Memory("MENUITEMINFO");
						mii.fMask = MIIM_SUBMENU;
						api.SetMenuItemInfo(g_.menu_handle, g_.menu_pos, true, mii);
						api.DestroyMenu(hSubMenu);
					}
					window.g_menu_click = 4;
					window.g_menu_button = 4;
					const lParam = pt.x + (pt.y << 16);
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
			const x = (pt.x - g_.mouse.ptGesture.x);
			const y = (pt.y - g_.mouse.ptGesture.y);
			if (Math.abs(x) + Math.abs(y) >= 20) {
				if (te.Data.Conf_TrailSize) {
					const hdc = api.GetWindowDC(te.hwnd);
					if (hdc) {
						const rc = api.Memory("RECT");
						api.GetWindowRect(te.hwnd, rc);
						api.MoveToEx(hdc, g_.mouse.ptGesture.x - rc.left, g_.mouse.ptGesture.y - rc.top, null);
						const pen1 = api.CreatePen(PS_SOLID, te.Data.Conf_TrailSize, te.Data.Conf_TrailColor);
						const hOld = api.SelectObject(hdc, pen1);
						api.LineTo(hdc, pt.x - rc.left, pt.y - rc.top);
						api.SelectObject(hdc, hOld);
						api.DeleteObject(pen1);
						g_.mouse.bTrail = true;
						api.ReleaseDC(te.hwnd, hdc);
					}
				}
				g_.mouse.ptGesture = pt.Clone();
				const s = (Math.abs(x) >= Math.abs(y)) ? ((x < 0) ? "L" : "R") : ((y < 0) ? "U" : "D");

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
				const Items = Ctrl.SelectedItems();
				for (let i = Items.Count; i--;) {
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
				const cFV = te.Ctrls(CTRL_FV);
				for (let i in cFV) {
					const FV = cFV[i];
					if (FV.hwndList) {
						const item = api.Memory("LVITEM");
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
	let hr = RunEvent3(msg + "!", Ctrl, Ctrl.Type, hwnd, msg, wParam, lParam);
	if (isFinite(hr)) {
		return hr;
	}
	hr = RunEvent3("Command", Ctrl, hwnd, msg, wParam, lParam);
	if (!isFinite(hr) && Ctrl.Type <= CTRL_EB) {
		if ((wParam & 0xfff) + 1 == CommandID_PROPERTIES) {
			hr = InvokeCommand(Ctrl.SelectedItems(), 0, te.hwnd, "properties", null, null, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
		} else if (wParam == 0x7032) {
			hr = InvokeCommand(Ctrl.FolderItem, 0, te.hwnd, "properties", null, null, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
		}
	}
	RunEvent1("ConfigChanged", "Window");
	return hr;
}

te.OnInvokeCommand = function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon) {
	let path;
	let strVerb = (isFinite(Verb) ? ContextMenu.GetCommandString(Verb, GCS_VERB) : Verb).toLowerCase();
	if (strVerb == "open" && ContextMenu.GetCommandString(Verb, GCS_HELPTEXT) != api.LoadString(hShell32, 12850)) {
		strVerb = "";
	}
	let hr = RunEvent3("InvokeCommand", ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon, strVerb);
	if (isNaN(hr) && strVerb == "cut") {
		const FV = ContextMenu.FolderView;
		if (FV && FV.hwndView && SameFolderItems(ContextMenu.Items(), FV.SelectedItems())) {
			const hMenu = te.MainMenu(FCIDM_MENU_EDIT);
			const mii = api.Memory("MENUITEMINFO");
			mii.fMask = MIIM_ID;
			for (let i = api.GetMenuItemCount(hMenu); i-- > 0;) {
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
	const Items = ContextMenu.Items();
	const Exec = [];
	if (/opennewwindow|opennewprocess/i.test(strVerb)) {
		CancelWindowRegistered();
	}

	let NewTab = GetNavigateFlags();
	for (let i = 0; i < Items.Count; ++i) {
		if (Verb && strVerb != "runas") {
			const Item = Items.Item(i);
			path = Item.ExtendedProperty("linktarget") || Item.Path;
			const cmd = api.AssocQueryString(ASSOCF_NONE, ASSOCSTR_COMMAND, Item.ExtendedProperty("linktarget") || Item, strVerb == "default" ? null : strVerb);
			if (cmd) {
				if (strVerb == "open") {
					const exp = GetWindowsPath("explorer.exe");
					if (api.PathMatchSpec(cmd, exp + ";" + exp + " /idlist,*;rundll32.exe *fldr.dll,RouteTheCall*")) {
						Navigate(Items.Item(i), NewTab);
						NewTab |= SBSP_NEWBROWSER;
						continue;
					}
				}
				const cmd2 = ExtractMacro(te, cmd);
				if (!SameText(cmd, cmd2)) {
					ShellExecute(cmd2.replace(/"?%1"?|%L/g, PathQuoteSpaces(path)).replace(/%\*|%I/g, ""), null, nShow, Directory);
					continue;
				}
			}
		}
		Exec.push(Items.Item(i));
	}
	if (Items.Count != Exec.length) {
		if (Exec.length) {
			const Selected = api.CreateObject("FolderItems");
			for (let i = 0; i < Exec.length; ++i) {
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
		const Items = ContextMenu.Items();
		for (let i = 0; i < Items.Count; ++i) {
			const Item = Items.Item(i);
			const path = Item.ExtendedProperty("linktarget") || Item.Path;
			Navigate(GetParentFolderName(path), SBSP_NEWBROWSER);
			SelectItem(te.Ctrl(CTRL_FV), path, SVSI_SELECT | SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS | SVSI_DESELECTOTHERS, 99, true);
			return S_OK;
		}
	}
	const FV = ContextMenu.FolderView;
	const Items = ContextMenu.Items();
	if (strVerb == "delete") {
		for (let i = Items.Count; i--;) {
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
			const hr2 = InvokeFunc(eo[i], [Ctrl, dataObj, pgrfKeyState[0], pt, pdwEffect, pgrfKeyState]);
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
			const hr = InvokeFunc(eo[i], [Ctrl, dataObj, pgrfKeyState[0], pt, pdwEffect, pgrfKeyState]);
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
			const hr = InvokeFunc(eo[i], [Ctrl, dataObj, pgrfKeyState[0], pt, pdwEffect, pgrfKeyState]);
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
			const hr2 = InvokeFunc(eo[i], [Ctrl]);
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
	if (isFinite(RunEvent3("FilterChanged", Ctrl))) {
		return;
	}
	const res = /\/(.*)\/(.*)/.exec(Ctrl.FilterView);
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
	const hr = RunEvent3("ShowContextMenu", Ctrl, hwnd, msg, wParam, pt, ContextMenu);
	if (isFinite(hr)) {
		return hr;
	}
	switch (Ctrl.Type) {
		case CTRL_SB:
		case CTRL_EB:
			const Selected = Ctrl.SelectedItems();
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
			const menus = te.Data.xmlMenus.getElementsByTagName("System");
			if (menus && menus.length) {
				const hMenu = api.GetSystemMenu(te.hwnd, false);
				const items = menus[0].getElementsByTagName("Item");
				const arMenu = OpenMenu(items, null);
				MakeMenus(hMenu, menus, arMenu, items, Ctrl, pt, 0, null, true);
			}
			break;
	}
	return S_FALSE;
}

te.OnDefaultCommand = function (Ctrl) {
	if (api.GetKeyState(VK_MENU) < 0) {
		api.SendMessage(Ctrl.hwndView, WM_COMMAND, CommandID_PROPERTIES - 1, 0);
		return S_OK;
	}
	const Selected = Ctrl.SelectedItems();
	const hr = RunEvent3("DefaultCommand", Ctrl, Selected);
	if (isFinite(hr)) {
		return hr;
	}
	if (ExecMenu(Ctrl, "Default", null, 2) == S_OK) {
		return S_OK;
	}
	if (Selected.Count) {
		let path = api.GetDisplayNameOf(Selected.Item(0), SHGDN_FORPARSING | SHGDN_FORADDRESSBAR);
		if (Selected.Count == 1) {
			const pid = api.ILCreateFromPath(path);
			if (pid.Enum) {
				Ctrl.Navigate(pid, GetNavigateFlags(Ctrl));
				return S_OK;
			}
		}
		return InvokeCommand(Selected, 0, te.hwnd, null, null, (Selected.Item(-1) || {}).Path, SW_SHOWNORMAL, 0, 0, Ctrl, CMF_DEFAULTONLY);
	}
	return S_FALSE;
}

te.OnSystemMessage = function (Ctrl, hwnd, msg, wParam, lParam) {
	if (!te.Data) {
		return S_OK;
	}
	let hr = RunEvent3(msg + "!", Ctrl, Ctrl.Type, hwnd, msg, wParam, lParam);
	if (isFinite(hr)) {
		return hr;
	}
	hr = RunEvent3("SystemMessage", Ctrl, hwnd, msg, wParam, lParam, Ctrl.Type);
	if (isFinite(hr)) {
		return hr;
	}
	let cFV;
	switch (Ctrl.Type) {
		case CTRL_TE:
			switch (msg) {
				case WM_DESTROY:
					const pid = api.Memory("DWORD");
					api.GetWindowThreadProcessId(te.hwnd, pid);
					WmiProcess("WHERE ExecutablePath = '" + (api.GetModuleFileName(null).split("\\").join("\\\\")) + "' AND ProcessId!=" + pid[0], function (item) {
						const hwnd = GethwndFromPid(item.ProcessId);
						api.SetWindowLongPtr(hwnd, GWL_EXSTYLE, api.GetWindowLongPtr(hwnd, GWL_EXSTYLE) & ~0x80);
						api.ShowWindow(hwnd, SW_SHOWNORMAL);
						api.SetForegroundWindow(hwnd);
					});
					for (let i in te.Data.Fonts) {
						api.DeleteObject(te.Data.Fonts[i]);
					}
					te.Data.Fonts = null;
					for (let i = SHIL_JUMBO + 1; i--;) {
						api.ImageList_Destroy(te.Data.SHIL[i], true);
					}
					te.Data.SHIL.length = 0;
					break;
				case WM_DEVICECHANGE:
					if (wParam == DBT_DEVICEARRIVAL || wParam == DBT_DEVICEREMOVECOMPLETE) {
						DeviceChanged(Ctrl, wParam, lParam);
					}
					break;
				case WM_ACTIVATE:
					g_.focused = void 0;
					if (g_.hwndTooltip) {
						api.ShowWindow(g_.hwndTooltip, SW_HIDE);
						g_.hwndTooltip = void 0;
					}
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
						const FV = te.Ctrl(CTRL_FV);
						const Layout = g_.TEData.Layout;
						if (FV) {
							FV.CurrentViewMode = g_.TEData.ViewMode;
						}
						te.Layout = Layout;
						delete g_.TEData;
					}
					if (g_.FVData) {
						if (g_.FVData.All) {
							delete g_.FVData.All;
							cFV = te.Ctrls(CTRL_FV);
							for (let i in cFV) {
								SetFolderViewData(cFV[i], g_.FVData);
							}
						} else {
							SetFolderViewData(te.Ctrl(CTRL_FV), g_.FVData);
						}
						delete g_.FVData;
					}
					if (g_.TVData) {
						if (g_.TVData.All) {
							cFV = te.Ctrls(CTRL_FV);
							for (let i in cFV) {
								SetTreeViewData(cFV[i], g_.TVData);
							}
						} else {
							SetTreeViewData(te.Ctrl(CTRL_FV), g_.TVData);
						}
						delete g_.TVData;
					}
					if (wParam & 0xffff) {
						if (g_.mouse.str == "") {
							FocusFV();
						}
					} else {
						g_.mouse.str = "";
						SetGestureText(Ctrl, "");
						if (g_.hwndTT && api.IsWindowVisible(g_.hwndTT)) {
							api.ShowWindow(g_.hwndTT, SW_HIDE);
						}
					}
					break;
				case WM_MOVE:
					if (g_.Fullscreen && api.GetKeyState(VK_LBUTTON) >= 0) {
						const hMonitor = api.MonitorFromWindow(te.hwnd, MONITOR_DEFAULTTOPRIMARY);
						const mi = api.Memory("MONITORINFOEX");
						api.GetMonitorInfo(hMonitor, mi);
						const rc = mi.rcMonitor;
						const x = rc.left;
						const y = rc.top;
						const w = rc.right - x;
						const h = rc.bottom - y;
						if (w > 0 && h > 0) {
							api.MoveWindow(te.hwnd, x, y, w, h, true);
						}
					} else {
						RunEvent1("ConfigChanged", "Window");
					}
					break;
				case WM_QUERYENDSESSION:
					SaveConfig();
					return true;
				case WM_SYSCOMMAND:
					if (wParam < 0xf000) {
						const menus = te.Data.xmlMenus.getElementsByTagName("System");
						if (menus && menus.length) {
							const items = menus[0].getElementsByTagName("Item");
							if (items) {
								const item = items[wParam - 1];
								if (item) {
									Exec(Ctrl, item.text, item.getAttribute("Type"), null);
									return 1;
								}
							}
						}
					} else if (!g_.Fullscreen && !api.IsZoomed(te.hwnd) && !api.IsIconic(te.hwnd)) {
						api.GetWindowRect(te.hwnd, te.Data.rcWindow);
					}
					break;
				case WM_POWERBROADCAST:
					if (wParam == PBT_APMRESUMEAUTOMATIC) {
						cFV = te.Ctrls(CTRL_FV);
						for (let i in cFV) {
							const FV = cFV[i];
							if (FV.hwndView) {
								if (api.PathIsNetworkPath(api.GetDisplayNameOf(FV, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING))) {
									FV.Suspend();
								}
							}
						}
					}
					break;
				case WM_SYSCOLORCHANGE:
					cFV = te.Ctrls(CTRL_FV);
					for (let i in cFV) {
						const FV = cFV[i];
						if (FV.hwndView) {
							FV.Suspend();
						}
					}
					break;
				case WM_SETTINGCHANGE:
					te.Data.TempFolder = GetTempPath(1);
					cFV = te.Ctrls(CTRL_FV, true);
					for (let i in cFV) {
						const hList = cFV[i].hwndList;
						if (hList) {
							api.SendMessage(hwnd, LVM_SETTEXTCOLOR, 0, GetSysColor(COLOR_WINDOWTEXT));
							api.SendMessage(hwnd, LVM_SETBKCOLOR, 0, GetSysColor(COLOR_WINDOW));
							api.SendMessage(hwnd, LVM_SETTEXTBKCOLOR, 0, GetSysColor(COLOR_WINDOW));
						}
					}
					break;
				case WM_NCLBUTTONDOWN:
					ExitFullscreen();
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
		const cd = api.Memory("COPYDATASTRUCT", 1, lParam);
		const hr = RunEvent3("CopyData", Ctrl, cd, wParam);
		if (isFinite(hr)) {
			return hr;
		}
		if (Ctrl.Type == CTRL_TE && cd.dwData == 0 && cd.cbData) {
			const strData = api.SysAllocStringByteLen(cd.lpData, cd.cbData, cd.cbData);
			RestoreFromTray();
			RunCommandLine(strData);
			return S_OK;
		}
	}
};

te.OnMenuMessage = function (Ctrl, hwnd, msg, wParam, lParam) {
	let pStatus = [null, "", 0];
	let hr = RunEvent3(msg + "!", Ctrl, (Ctrl || te).Type, hwnd, msg, wParam, lParam, pStatus);
	if (isFinite(hr)) {
		ShowStatusText(pStatus[0], pStatus[1], pStatus[2], pStatus[3]);
		return hr;
	}
	hr = RunEvent3("MenuMessage", Ctrl, hwnd, msg, wParam, lParam, pStatus);
	if (isFinite(hr)) {
		ShowStatusText(pStatus[0], pStatus[1], pStatus[2], pStatus[3]);
		return hr;
	}
	switch (msg) {
		case WM_INITMENUPOPUP:
			const s = api.GetMenuString(wParam, 0, MF_BYPOSITION);
			if (api.PathMatchSpec(s, "\t*Script\t*")) {
				const ar = s.split("\t");
				const mii = api.Memory("MENUITEMINFO");
				mii.fMask = MIIM_ID | MIIM_STATE | MIIM_STRING;
				mii.wId = -2;
				mii.fState = MFS_DISABLED;
				mii.dwTypeData = GetTextR("@shell32,-13585[-4223]");
				api.SetMenuItemInfo(wParam, 0, true, mii);
				ExecScriptEx(Ctrl, ar[2], ar[1], hwnd);
			}
			break;
		case WM_MENUSELECT:
			pStatus[0] = te;
			RunEvent1("MenuSelect", Ctrl, hwnd, msg, wParam, lParam, pStatus);
			if (lParam) {
				const nVerb = wParam & 0xffff;
				if (Ctrl) {
					for (let i in Ctrl) {
						const CM = Ctrl[i];
						if (CM && nVerb >= CM.idCmdFirst && nVerb <= CM.idCmdLast) {
							Text = CM.GetCommandString(nVerb - CM.idCmdFirst, GCS_HELPTEXT);
							if (Text) {
								pStatus[1] = Text;
								break;
							}
						}
					}
				}
				const mf = wParam >> 16;
				if (FolderMenu.MenuLoop) {
					let wId = nVerb;
					if (mf & MF_POPUP) {
						const mii = api.Memory("MENUITEMINFO");
						mii.fMask = MIIM_ID;
						api.GetMenuItemInfo(lParam, nVerb, true, mii);
						wId = mii.wId;
					}
					const pid = FolderMenu.Items[wId - 1];
					if (pid) {
						g_.MenuSelected = pid;
						pStatus = [pid, pid.Name, 0, 200];
					}
				}
				const hSubMenu = api.GetSubMenu(lParam, nVerb);
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
			const ar = ["ExitMenuLoop", "EnterMenuLoop", "MenuSelect", "MenuChar"];
			RunEvent1(ar[0], Ctrl, hwnd, msg, wParam, lParam);
			for (let i = ar.length; i--;) {
				const eo = eventTE[ar[i].toLowerCase()];
				if (eo) {
					eo.length = 0;
				}
			}
			for (let i in g_arBM) {
				api.DeleteObject(g_arBM[i]);
			}
			g_arBM = [];
			g_.menu_loop = false;
			delete g_.ptMenuDrag;
			delete g_.MenuSelected
			pStatus = [GetFolderView(), "", 0];
			break;
		case WM_MENUCHAR:
			const hr = RunEvent4("MenuChar", Ctrl, hwnd, msg, wParam, lParam);
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
	let hr = RunEvent3(msg + "!", Ctrl, Ctrl.Type, hwnd, msg, wParam, lParam);
	if (isFinite(hr)) {
		return hr;
	}
	hr = RunEvent3("AppMessage", Ctrl, hwnd, msg, wParam, lParam, Ctrl.Type);
	if (isFinite(hr)) {
		return hr;
	}
	if (msg == TWM_CHANGENOTIFY) {
		const pidls = api.CreateObject("Object");
		const hLock = api.SHChangeNotification_Lock(wParam, lParam, pidls);
		if (hLock) {
			api.SHChangeNotification_Unlock(hLock);
			if (pidls[0] && IsWitness(pidls[0])) {
				const path = pidls[0].Path;
				if (/^[A-Z]:\\|^\\\\\w/i.test(path)) {
					for (let key in g_.Notify) {
						if (new Date().getTime() > g_.Notify[key]) {
							delete g_.Notify[key];
						}
					}
					const lEvent = pidls.lEvent;
					const key = lEvent + ":" + path;
					if (!g_.Notify[key] && !(lEvent == SHCNE_RENAMEFOLDER && api.ILIsEqual(pidls[0], pidls[1]))) {
						g_.Notify[key] = new Date().getTime() + 999;
						ChangeNotifyFV(lEvent, pidls[0], pidls[1]);
						RunEvent1("ChangeNotify", Ctrl, pidls, wParam, lParam);
						if (lEvent & (SHCNE_UPDATEITEM | SHCNE_UPDATEDIR | SHCNE_CREATE | SHCNE_MKDIR)) {
							RunEvent1("ChangeNotifyItem:" + path, pidls[0], 0, lEvent);
						}
						if ((lEvent & (SHCNE_RENAMEITEM | SHCNE_RENAMEFOLDER)) && pidls[1]) {
							RunEvent1("ChangeNotifyItem:" + pidls[1].Path, pidls[1], 1, lEvent);
						}
					}
				}
			}
		}
		return S_OK;
	}
	return S_FALSE;
}

te.OnNewWindow = function (Ctrl, FolderItem) {
	const hr = RunEvent3("NewWindow", Ctrl, FolderItem);
	if (isFinite(hr)) {
		return hr;
	}
	Navigate(FolderItem, SBSP_NEWBROWSER);
	return S_OK;
}

te.OnClipboardText = function (Items) {
	if (!Items) {
		return "";
	}
	const r = RunEvent4("ClipboardText", Items);
	if (r !== void 0) {
		return r;
	}
	const s = [];
	for (let i = Items.Count; i-- > 0;) {
		s.unshift(PathQuoteSpaces(api.GetDisplayNameOf(Items.Item(i), SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)))
	}
	return s.join(" ");
}

te.OnILGetParent = function (FolderItem) {
	const r = RunEvent4("ILGetParent", FolderItem);
	if (r !== void 0) {
		return r;
	}
	if (FolderItem.Unavailable) {
		const path = api.GetDisplayNameOf(FolderItem, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING | SHGDN_FORPARSINGEX);
		if (/^[A-Z]:\\/i.test(path)) {
			return api.PathIsNetworkPath(path) ? ssfDRIVES : GetParentFolderName(path) || ssfDRIVES;
		}
		return /^\\\\\w/i.test(path) ? ssfNETWORK : ssfDESKTOP;
	}
	const res = IsSearchPath(FolderItem);
	if (res) {
		const pid = api.ILRemoveLastID(FolderItem);
		const res1 = IsSearchPath(pid, true);
		if ((res1 && !res1[1]) || api.ILIsEmpty(pid)) {
			return unescape(res[1]);
		}
	}
}

te.OnReplacePath = function (Ctrl, Path) {
	if (/^[A-Z]:\\.+?\\|^\\\\.+?\\/i.test(Path)) {
		const i = Path.indexOf("\\/");
		if (i > 0) {
			const fn = Path.slice(i + 1);
			if (/\//.test(fn)) {
				Ctrl.FilterView = fn;
				return Path.slice(0, i);
			}
		}
		const fn = GetFileName(Path);
		if (/[\*\?]/.test(fn)) {
			Ctrl.FilterView = fn;
			return GetParentFolderName(Path);
		}
	}
	let r = RunEvent4("ReplacePath", Ctrl, Path);
	if (/\//.test(r || Path) && !/^[0-9A-Z]{2,}:/i.test(r || Path)) {
		r = (r || Path).replace(/\//g, "\\").replace(/^\\([A-Z])(\\.*)$/i, "$1:$2");
	}
	return r;
}

te.OnGetAlt = function (SessionId, s) {
	const pid = te.Ctrl(CTRL_FI, SessionId);
	if (pid) {
		return BuildPath(pid.Path, GetFileName(s));
	}
}

te.OnFilterView = function (FV, s) {
	return S_FALSE;
}

te.OnVisibleChanged = function (Ctrl) {
	RunEvent1("VisibleChanged", Ctrl, Ctrl.Visible, Ctrl.Type, Ctrl.Id);
}

te.OnColumnClick = function (Ctrl, iItem) {
	let hr = RunEvent3("ColumnClick", Ctrl, iItem);
	if (isFinite(hr)) {
		return hr;
	}
	const cColumns = api.CommandLineToArgv(Ctrl.Columns(1));
	if (HasCustomSort(Ctrl, cColumns[iItem * 2])) {
		return S_OK;
	}
}

te.OnEndThread = function () {
	FocusFV();
	return RunEvent1("EndThread", api.GetThreadCount());
}

te.OnShowError = function (e, s) {
	let hr = RunEvent3("ShowError", Ctrl, iItem);
	if (isFinite(hr)) {
		return hr;
	}
	return ShowError(e, s);
}

ShowStatusText = function (Ctrl, Text, iPart, tm) {
	if (!Ctrl) {
		return;
	}
	if (tm) {
		ShowStatusTextEx(Ctrl, Text, iPart, tm);
		return S_OK;
	}
	RunEvent1("StatusText", Ctrl, Text, iPart, Ctrl.Type);
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
	if (Ctrl.Data) {
		RunEvent1("ViewModeChanged", Ctrl);
	}
}

g_.event.columnschanged = function (Ctrl) {
	if (Ctrl.Data) {
		RunEvent1("ColumnsChanged", Ctrl);
	}
}

g_.event.contentschanged = function (Ctrl) {
	if (Ctrl.Data) {
		RunEvent1("ContentsChanged", Ctrl);
	}
}

g_.event.iconsizechanged = function (Ctrl) {
	if (Ctrl.Data) {
		RunEvent1("IconSizeChanged", Ctrl);
	}
}

g_.event.itemclick = function (Ctrl, Item, HitTest, Flags) {
	return RunEvent3("ItemClick", Ctrl, Item, HitTest, Flags);
}

g_.event.tooltip = function (Ctrl, Index, hwnd) {
	return RunEvent4("ToolTip", Ctrl, Index, hwnd);
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

g_.event.windowregistered = function (Ctrl) {
	if (!g_.tmWindowRegistered || new Date().getTime() > g_.tmWindowRegistered) {
		RunEvent1("WindowRegistered", Ctrl);
	} else {
		g_.tmWindowRegistered = void 0;
	}
}

//Tablacus Events

GetIconImage = function (Ctrl, clBk, bSimple) {
	if ("number" !== typeof clBk) {
		clBk = CLR_DEFAULT | COLOR_WINDOW;
	}
	if ((clBk & CLR_DEFAULT) == CLR_DEFAULT) {
		clBk = GetSysColor(clBk & ~CLR_DEFAULT);
	}
	const nSize = api.GetSystemMetrics(SM_CYSMICON);
	if (!Ctrl) {
		return MakeImgDataEx("icon:shell32.dll,234", bSimple, nSize, clBk);
	}
	const FolderItem = Ctrl.FolderItem || Ctrl;
	const r = GetNetworkIcon(FolderItem.Path);
	if (r) {
		if (FolderItem.Unavailable > 999) {
			return MakeImgDataEx("icon:shell32.dll,234", bSimple, nSize, clBk);
		}
		return MakeImgDataEx(r, bSimple, nSize, clBk);
	}
	let img = RunEvent4("GetIconImage", Ctrl, clBk, bSimple);
	if (img) {
		return MakeImgDataEx(ExtractMacro(te, img), bSimple, nSize, clBk);
	}
	api.ILIsEmpty(FolderItem);
	if (FolderItem.Unavailable > 999) {
		return MakeImgDataEx("icon:shell32.dll,234", bSimple, nSize, clBk);
	}
	if (g_.IEVer >= 8) {
		if (bSimple) {
			return bSimple != 2 ? api.GetDisplayNameOf(FolderItem, SHGDN_FORPARSING) : "";
		}
		const hIcon = GetHICON(SHGetFileInfo(FolderItem, 0, SHGFI_SYSICONINDEX | SHGFI_PIDL).iIcon, nSize, ILD_NORMAL);
		if (hIcon) {
			img = api.CreateObject("WICBitmap").FromHICON(hIcon);
			api.DestroyIcon(hIcon);
			return img.DataURI();
		}
	}
	return MakeImgDataEx("folder:closed", bSimple, nSize, clBk);
}

UnlockFV = function (Item) {
	const cFV = te.Ctrls(CTRL_FV);
	for (let i in cFV) {
		const FV = cFV[i];
		if (FV.hwndView && api.ILIsParent(Item, FV, false) && !api.ILIsEqual(FV.FolderItem.Alt, ssfRESULTSFOLDER)) {
			FV.Suspend();
		}
	}
}

ChangeNotifyFV = function (lEvent, item1, item2) {
	const fAdd = SHCNE_DRIVEADD | SHCNE_MEDIAINSERTED | SHCNE_NETSHARE | SHCNE_MKDIR | SHCNE_UPDATEDIR | SHCNE_UPDATEITEM;
	const fRemove = SHCNE_DRIVEREMOVED | SHCNE_MEDIAREMOVED | SHCNE_NETUNSHARE | SHCNE_RENAMEFOLDER | SHCNE_RMDIR | SHCNE_SERVERDISCONNECT | SHCNE_UPDATEDIR;
	if (lEvent & (SHCNE_DISKEVENTS | fAdd | fRemove) && item1.IsFileSystem) {
		const path1 = item1.Path;
		const bNetwork = api.ILIsEqual(item1, ssfNETWORK);
		const cFV = te.Ctrls(CTRL_FV, true);
		for (let i in cFV) {
			const FV = cFV[i];
			if (FV && FV.FolderItem && IsWitness(FV.FolderItem)) {
				const path = FV.FolderItem.Path;
				const bParent = api.PathMatchSpec(path, [path1.replace(/\\$/, ""), path1].join("\\*;")) || bNetwork && api.PathIsNetworkPath(path);
				if (lEvent == SHCNE_RENAMEFOLDER && CanClose(FV) == S_OK) {
					if (bParent) {
						FV.Navigate(path.replace(path1, api.GetDisplayNameOf(item2, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING)), SBSP_SAMEBROWSER);
						continue;
					}
				}
				if (FV.hwndView) {
					const bChild = SameText(GetParentFolderName(path1), path);
					if (bChild) {
						if (FV.hwndList) {
							const item = api.Memory("LVITEM");
							item.stateMask = LVIS_CUT;
							api.SendMessage(FV.hwndList, LVM_SETITEMSTATE, -1, item);
						}
						if (WINVER >= 0x600 && (lEvent & (SHCNE_DELETE | SHCNE_RMDIR))) {
							const nPos = FV.GetFocusedItem - 1;
							if (nPos > 0 && api.ILIsEqual(item1, FV.FocusedItem)) {
								const nCount = FV.ItemCount && FV.ItemCount(SVGIO_ALLVIEW);
								if (nCount) {
									FV.SelectItem(nPos < nCount ? nPos : nCount - 1, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | (FV.Id == te.Ctrl(CTRL_FV).Id ? 0 : SVSI_NOTAKEFOCUS));
								}
							}
						}
					}
					if (bChild || bParent) {
						if ((lEvent & fRemove) || ((lEvent & fAdd) && FV.FolderItem.Unavailable)) {
							const tm = Math.max(500, te.Data.Conf_NetworkTimeout);
							RefreshEx(FV, tm, tm);
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
		const ar = strKey.split(",");
		for (let i in ar) {
			SetKeyExec(mode, ar[i], path, type, bLast);
		}
		return;
	}
	if (strKey) {
		strKey = GetKeyKey(strKey);
		const KeyMode = eventTE.Key[mode];
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
		const ar = strGesture.split(",");
		for (let i in ar) {
			SetGestureExec(mode, ar[i], path, type, bLast, Name);
		}
		return;
	}
	if (strGesture) {
		strGesture = strGesture.toUpperCase();
		const MouseMode = eventTE.Mouse[mode];
		if (MouseMode) {
			if (!MouseMode[strGesture]) {
				MouseMode[strGesture] = [];
			}
			MouseMode[strGesture][bLast ? "push" : "unshift"]([path, type, Name]);
		}
	}
}

ArExec = function (Ctrl, ar, pt, hwnd) {
	for (let i in ar) {
		const cmd = ar[i];
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
	const pt = api.Memory("POINT");
	if (Ctrl.Type <= CTRL_EB || Ctrl.Type == CTRL_TV) {
		const rc = api.Memory("RECT");
		Ctrl.GetItemRect(Ctrl.FocusedItem || Ctrl.SelectedItem, rc);
		pt.x = rc.left;
		pt.y = rc.top;
	}
	api.ClientToScreen(Ctrl.hwnd, pt);
	return ArExec(Ctrl, eventTE.Key[mode][nKey], pt, hwnd);
}

CancelWindowRegistered = function () {
	g_.tmWindowRegistered = new Date().getTime() + 9999;
}

importScripts = function () {
	for (let i = 0; i < arguments.length; ++i) {
		importScript(arguments[i]);
	}
}

AddEvent("Arrange", function (Ctrl, rc, Type, Id, FV) {
	if (Type == CTRL_TE && !api.IsIconic(te.hwnd)) {
		const rcClient = api.Memory("RECT");
		api.GetClientRect(te.hwnd, rcClient);
		rcClient.left += te.offsetLeft;
		rcClient.top += te.offsetTop;
		rcClient.right -= te.offsetRight;
		rcClient.bottom -= te.offsetBottom;
		te.LockUpdate(1);
		try {
			const cTC = te.Ctrls(CTRL_TC, true);
			for (let i = 0; i < cTC.Count; ++i) {
				const TC = cTC[i];
				const rc = api.Memory("RECT");
				const w = (rcClient.right - rcClient.left) / 100;
				const h = (rcClient.bottom - rcClient.top) / 100;
				let s = TC.Left;
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
				te.OnArrange(TC, rc, TC.Type, TC.Id, TC.Selected);
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
				return InvokeFunc(fn.Drop, [Ctrl, s, type, hwnd, pt, dataObj, grfKeyState, pdwEffect, bDrop]);
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
			const hr = InvokeFunc(fn.Exec, [Ctrl, s, type, hwnd, pt]);
			return hr != null ? hr : fn.Result;
		}
		return g_basic.Exec(Ctrl, s, type, hwnd, pt);
	}
});

AddEvent("GetFolderItemName", function (pid) {
	const res = /^ftp:\/\/([^:\/]+)[^\/]*\/$/i.exec(pid.Path);
	return res && res[1];
});

AddEvent("MenuState:Tabs:Close Tab", function (Ctrl, pt, mii) {
	const FV = GetFolderView(Ctrl, pt);
	if (CanClose(FV)) {
		mii.fMask |= MIIM_STATE;
		mii.fState |= MFS_DISABLED;
	}
});

AddEvent("MenuState:Tabs:Close Tabs on Left", function (Ctrl, pt, mii) {
	const FV = GetFolderView(Ctrl, pt, true);
	if (FV && FV.Index == 0) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_DISABLED;
	}
});

AddEvent("MenuState:Tabs:Close Tabs on Right", function (Ctrl, pt, mii) {
	const FV = GetFolderView(Ctrl, pt, true);
	if (FV) {
		const TC = FV.Parent;
		if (FV.Index >= TC.Count - 1) {
			mii.fMask |= MIIM_STATE;
			mii.fState = MFS_DISABLED;
		}
	}
});

AddEvent("MenuState:Tabs:Up", function (Ctrl, pt, mii) {
	const FV = GetFolderView(Ctrl, pt);
	if (!FV || api.ILIsEmpty(FV)) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_DISABLED;
	}
});

AddEvent("MenuState:Tabs:Lock", function (Ctrl, pt, mii) {
	const FV = GetFolderView(Ctrl, pt);
	if (GetLock(FV)) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_CHECKED;
	}
});

AddEvent("MenuState:Tabs:Show frames", function (Ctrl, pt, mii) {
	if (WINVER < 0x600) {
		mii.fMask = 0;
		return S_OK;
	}
	const FV = GetFolderView(Ctrl, pt);
	if (!FV || FV.Type == CTRL_EB) {
		mii.fMask |= MIIM_STATE;
		mii.fState = MFS_CHECKED;
	}
});

AddEvent("ChangeNotify", function (Ctrl, pidls, wParam, lParam) {
	if (pidls.lEvent & (SHCNE_MKDIR | SHCNE_CREATE) && pidls[0].IsFileSystem) {
		const tm = new Date().getTime();
		if (g_.NewItemTime > tm) {
			delete g_.NewItemTime;
		} else if (g_.NewFolderTime > tm) {
			delete g_.NewFolderTime;
			g_.NewFolderTV.Expand(pidls[0], 0);
			setTimeout(function () {
				SetModifierKeys(0);
				wsh.SendKeys("{F2}");
			}, 99);
		}
	}
});

AddEvent("ReplacePath", function (FolderItem, Path) {
	if (/^file:/.test(Path)) {
		return decodeURI(Path);
	}
	if (/^ftp:.*[^\/]$/.test(Path)) {
		return Path + "/";
	}
});

AddEvent("ToolTip", function (Ctrl, Index, hwnd) {
	if (Index == WM_PAINT && Ctrl.Type == CTRL_TE) {
		g_.hwndTooltip = hwnd;
	}
});

AddEvent("AddType", function (arFunc) {
	for (let i in g_basic.Func) {
		arFunc.push(i);
	}
});

AddType = function (strType, o) {
	api.ObjPutI(g_basic.Func, strType.replace(/&|\.\.\.$/g, ""), o);
};

AddTypeEx = function (strType, strTitle, fn) {
	const type = g_basic.FuncI(strType);
	if (type && type.Cmd) {
		api.ObjPutI(type.Cmd, strTitle, fn);
	}
};

AddEvent("OptionRef", function (Id, s, pt) {
	const fn = g_basic.FuncI(Id);
	if (fn) {
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
		const lines = p.s.split(/\r?\n/);
		for (let i in lines) {
			const res = /^([^,]+),(.*)$/.exec(lines[i]);
			if (res) {
				const p2 = api.CreateObject("Object");
				p2.s = res[2];
				Id = GetSourceText(res[1]);
				OptionEncode(Id, p2);
				lines[i] = [Id, p2.s].join(",");
			}
		}
		p.s = lines.join("\n");
		return S_OK;
	}
	const type = g_basic.FuncI(Id);
	if (type && type.Enc) {
		p.s = GetSourceText(p.s);
		return S_OK;
	}
});

AddEvent("OptionDecode", function (Id, p) {
	if (Id === "") {
		const lines = p.s.split(/\r?\n/);
		for (let i in lines) {
			const res = /^([^,]+),(.*)$/.exec(lines[i]);
			if (res) {
				const p2 = api.CreateObject("object");
				p2.s = res[2];
				Id = res[1];
				OptionDecode(Id, p2);
				lines[i] = [GetText(Id), p2.s].join(",");
			}
		}
		p.s = lines.join("\n");
		return S_OK;
	}
	const type = g_basic.FuncI(Id);
	if (type && type.Enc) {
		if (SameText(GetSourceText(GetText(p.s)), p.s)) {
			p.s = GetText(p.s);
			return S_OK;
		}
		return S_OK;
	}
});

AddEvent("BeginNavigate", function (Ctrl) {
	const fn = Ctrl.FolderItem.Enum;
	if (fn) {
		const SessionId = Ctrl.SessionId;
		const Items = fn(Ctrl.FolderItem, Ctrl, function (Ctrl, Items) {
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
	if (pid && pid.Path && !/^ftp:|^https?:/i.test(pid.Path) && !pid.Unavailable && !api.GetAttributesOf(pid.Alt || pid, SFGAO_FILESYSTEM | SFGAO_FILESYSANCESTOR | SFGAO_STORAGEANCESTOR | SFGAO_NONENUMERATED | SFGAO_DROPTARGET | SFGAO_STORAGE) && !api.ILIsParent(1, pid, false) && !IsSearchPath(pid) && api.GetDisplayNameOf(pid, SHGDN_FORPARSING) != "::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}") {
		return true;
	}
});

AddEvent("LocationPopup", function (hMenu) {
	FolderMenu.AddMenuItem(hMenu, api.ILCreateFromPath(ssfDESKTOP));
	FolderMenu.AddMenuItem(hMenu, api.ILCreateFromPath(ssfDRIVES));
	const Items = FolderMenu.Enum(api.ILCreateFromPath(ssfDRIVES));
	const path0 = api.GetDisplayNameOf(ssfDESKTOP, SHGDN_FORPARSING);
	for (let i = 0; i < Items.Count; ++i) {
		const Item = Items.Item(i);
		if (IsFolderEx(Item)) {
			const path = api.GetDisplayNameOf(Item, SHGDN_FORPARSING);
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

AddEvent("ReplaceMacroEx", [/"%Current%/ig, function (strMatch, ref1) {
	let strSel = "";
	const FV = GetFolderView(Ctrl);
	if (FV) {
		strSel = '"' + api.GetDisplayNameOf(FV, SHGDN_FORPARSING);
	}
	return strSel;
}]);

AddEvent("ConfigChanged", function (s) {
	te.Data["bSave" + s] = true;
});

AddEvent("ToolTip", function (Ctrl, Index, hwnd) {
	if (Ctrl.Type <= CTRL_EB && Index >= 0) {
		g_.hwndTT = hwnd;
	}
});

if (!window.chrome) {
	AddEvent("BrowserCreatedEx", 'MainWindow.RunEvent1("BrowserCreated", document);');
}

if (WINVER >= 0x600 && screen.deviceYDPI > 96) {
	AddEvent("NavigateComplete", FixIconSpacing);
	AddEvent("IconSizeChanged", FixIconSpacing);
	AddEvent("ViewModeChanged", FixIconSpacing);
}

AddEnv("Selected", function (Ctrl) {
	const ar = [];
	const Selected = GetSelectedItems(Ctrl);
	if (Selected) {
		for (let i = Selected.Count; i > 0; ar.unshift(PathQuoteSpaces(api.GetDisplayNameOf(Selected.Item(--i), SHGDN_FORPARSING)))) {
		}
	}
	return ar.join(" ");
});

AddEnv("Current", function (Ctrl) {
	const FV = GetFolderView(Ctrl);
	return FV ? PathQuoteSpaces(api.GetDisplayNameOf(FV, SHGDN_FORPARSING)) : '';
});

AddEnv("TreeSelected", function (Ctrl) {
	if (!Ctrl || Ctrl.Type != CTRL_TV) {
		const FV = GetFolderView(Ctrl);
		if (FV) {
			Ctrl = FV.TreeView;
		}
	}
	return Ctrl ? PathQuoteSpaces(Ctrl.SelectedItem.Path) : '';
});

AddEnv("MenuSelected", function (Ctrl) {
	return g_.MenuSelected ? PathQuoteSpaces(g_.MenuSelected.Path) : "";
});

AddEnv("Installed", GetDriveName(api.GetModuleFileName(null)));

AddEnv("TE_Config", function (Ctrl) {
	return BuildPath(te.Data.DataFolder, "config");
});

CreateUpdater = function (arg) {
	let te_new = BuildPath(arg.temp, "te64.exe");
	let nDog = 300;
	while (!api.PathFileExists(te_new)) {
		if (wsh.Popup(GetText("Please wait."), 1, TITLE, MB_OKCANCEL) == IDCANCEL || nDog-- == 0) {
			return;
		}
	}
	const arDel = [];
	const addons = BuildPath(arg.temp, "addons");
	if (api.PathIsDirectory(BuildPath(arg.temp, "config"))) {
		arDel.push(BuildPath(arg.temp, "config"));
	}
	const wfd = api.Memory("WIN32_FIND_DATA");
	for (let i = 32; i <= 64; i += 32) {
		te_new = BuildPath(arg.temp, "te" + i + ".exe");
		let te_old = BuildPath(te.Data.Installed, "te" + i + ".exe");
		const hFind = api.FindFirstFile(BuildPath(arg.temp, "lib", "*" + i + ".dll"), wfd);
		if (!api.PathFileExists(te_old)) {
			arDel.push(te_new);
			for (let bFind = hFind != INVALID_HANDLE_VALUE; bFind; bFind = api.FindNextFile(hFind, wfd)) {
				arDel.push(BuildPath(arg.temp, "lib", wfd.cFileName));
			}
		} else {
			if (api.GetFileVersionInfo(te_new) == api.GetFileVersionInfo(te_old)) {
				arDel.push(te_new);
			}
			for (let bFind = hFind != INVALID_HANDLE_VALUE; bFind; bFind = api.FindNextFile(hFind, wfd)) {
				const n = wfd.cFileName;
				if (!/^\.$|^\.\.$/i.test(n)) {
					te_new = BuildPath(arg.temp, "lib", n);
					te_old = BuildPath(te.Data.Installed, "lib", n);
					if (api.PathFileExists(te_old) && api.GetFileVersionInfo(te_new) == api.GetFileVersionInfo(te_old)) {
						arDel.push(te_new);
					}
				}
			}
		}
		api.FindClose(hFind);
	}

	const hFind = api.FindFirstFile(BuildPath(addons, "*"), wfd);
	for (let bFind = hFind != INVALID_HANDLE_VALUE; bFind; bFind = api.FindNextFile(hFind, wfd)) {
		if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
			const n = wfd.cFileName;
			if (!/^\.$|^\.\.$/i.test(n)) {
				const items = te.Data.Addons.getElementsByTagName(n);
				if (!items || items.length == 0) {
					arDel.push(BuildPath(addons, n));
				}
			}
		}
	}
	api.FindClose(hFind);

	if (arDel.length) {
		api.SHFileOperation(FO_DELETE, arDel, null, FOF_SILENT | FOF_NOCONFIRMATION, false);
	}
	const ppid = api.Memory("DWORD");
	api.GetWindowThreadProcessId(te.hwnd, ppid);
	arg.pid = ppid[0];

	if (isFinite(RunEvent3("CreateUpdater", arg))) {
		return;
	}
	setTimeout(UpdateAndReload, 500, arg);
};

UpdateAndReload = function (arg) {
	if (arg.pcRef && arg.pcRef[0] > 0) {
		if (!arg.tm || new Date().getTime() < arg.tm) {
			setTimeout(UpdateAndReload, 999, arg);
			return;
		}
	}
	if (CanClose(te)) {
		setTimeout(UpdateAndReload, 999, arg);
		return;
	}
	if (!arg.Boot) {
		if (!IsExists(BuildPath(arg.temp, GetFileName(api.GetModuleFileName(null))))) {
			const wfd = api.Memory("WIN32_FIND_DATA");
			const hFind = api.FindFirstFile(BuildPath(arg.temp, "lib", "*" + g_.bit + ".dll"), wfd);
			if (hFind == INVALID_HANDLE_VALUE) {
				api.SHFileOperation(FO_MOVE, arg.temp + "\\*", te.Data.Installed, FOF_NOCONFIRMATION, false);
				ReloadCustomize();
				return;
			}
			api.FindClose(hFind);
		}
	}
	let update = BuildPath(arg.temp, "script\\update.js");
	if (!api.PathFileExists(update)) {
		update = BuildPath(te.Data.Installed, "script\\update.js");
	}
	g_.strUpdate = [PathQuoteSpaces(BuildPath(api.IsWow64Process(api.GetCurrentProcess()) ? GetWindowsPath("Sysnative") : system32, "wscript.exe")), '/E:JScript', PathQuoteSpaces(update), PathQuoteSpaces(api.GetModuleFileName(null)), PathQuoteSpaces(arg.temp), PathQuoteSpaces(api.LoadString(hShell32, 12612)), PathQuoteSpaces(api.LoadString(hShell32, 12852))].join(" ");
	WmiProcess("WHERE ExecutablePath='" + (api.GetModuleFileName(null).split("\\").join("\\\\")) + "' AND ProcessId!=" + arg.pid, function (item) {
		if (!/\/run/i.test(item.CommandLine)) {
			item.Terminate();
		}
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
	const iItem = Ctrl.HitTest(pt, LVHT_ONITEM);
	if (iItem >= 0) {
		return false;
	}
	const pt2 = pt.Clone();
	api.ScreenToClient(hwnd, pt2);
	return pt2.y < screen.deviceYDPI / 4;
}

AutocompleteThread = function () {
	let pid = api.ILCreateFromPath(path);
	if (!pid.IsFolder) {
		var res = /^(.*)([\\\/])/.exec(path);
		var d = res && /^[A-Z]:/i.test(res[1]) ? 3 : 1;
		var r = res ? res[1].length < d ? res[1] + res[2] : res[1] : "";
		var s = r != path && r != "\\" && r.length >= d ? r : "";
		pid = api.ILCreateFromPath(s);
	}
	if (pid.IsFolder && pid.Path != Autocomplete.Path) {
		Autocomplete.Path = pid.Path;
		const Items = pid.GetFolder.Items();
		try {
			Items.Filter(fflag, "*");
		} catch (e) { }
		for (let i = 0; i < Items.Count; ++i) {
			if (Items.Item(i).IsFolder) {
				ar.push(Items.Item(i).Path);
			}
		}
		api.Invoke(UI.Autocomplete, [ar.join("\t"), pid.Path]);
	}
}

AdjustAutocomplete = function (path) {
	if (te.Data.Conf_NoAutocomplete) {
		return;
	}
	const o = api.CreateObject("Object");
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

FullscreenChanged = function (bFullscreen, bBody) {
	if (bFullscreen && !bBody) {
		const cTC = te.Ctrls(CTRL_TC, true);
		for (let i in cTC) {
			const TC = cTC[i];
			g_.stack_TC.push(TC);
			TC.Visible = false;
		}
	}
	if (!g_.Fullscreen != !bFullscreen) {
		let rc;
		const dwStyle = api.GetWindowLongPtr(te.hwnd, GWL_STYLE);
		if (bFullscreen) {
			if (!bBody) {
				const cTC = te.Ctrls(CTRL_TC, true);
				for (let i in cTC) {
					const TC = cTC[i];
					g_.stack_TC.push(TC);
					TC.Visible = false;
				}
			}
			g_.FullscreenWS = dwStyle;
			api.SetWindowLongPtr(te.hwnd, GWL_STYLE, g_.FullscreenWS & ~0xCF0000);
			const hMonitor = api.MonitorFromWindow(te.hwnd, MONITOR_DEFAULTTOPRIMARY);
			const mi = api.Memory("MONITORINFOEX");
			api.GetMonitorInfo(hMonitor, mi);
			rc = mi.rcMonitor;
		} else {
			while (g_.stack_TC.length) {
				g_.stack_TC.pop().Visible = true;
			}
			if (g_.FullscreenWS != null) {
				api.SetWindowLongPtr(te.hwnd, GWL_STYLE, g_.FullscreenWS);
			}
			rc = te.Data.rcWindow;
		}
		const x = rc.left;
		const y = rc.top;
		const w = rc.right - x;
		const h = rc.bottom - y;
		g_.Fullscreen = bFullscreen;
		if (w > 0 && h > 0) {
			api.MoveWindow(te.hwnd, x, y, w, h, true);
		}
		RunEvent1("FullscreenChanged", bFullscreen, bBody);
	}
}

GetAddonLocation = function (strName) {
	const items = te.Data.Addons.getElementsByTagName(strName);
	return items.length && items[0].getAttribute("Location");
}

GetAddonIcon = function (Id) {
	return g_.AddonsIcon[Id] || GetMiscIcon(Id);
}

SetAddonIcon = function (Id, s) {
	g_.AddonsIcon[Id] = s;
}

GetAddonIconImg = function (Id, json, n, text) {
	const o = JSON.parse(json || {});
	o.src = GetAddonIcon(Id);
	if (!o.src) {
		if (!text) {
			return "";
		}
		AddClass(o, text);
	}
	return GetImgTag(o, GetIconSize(0, n));
}

SetAddonLocation = function (Location, strName) {
	if (!g_.Locations[Location]) {
		g_.Locations[Location] = api.CreateObject("Array");
	}
	const ar = strName.split(/\t/);
	g_.Locations[Location].push(ar[0]);
	if (ar.length > 1 && !g_.AddonsIcon[ar[0]]) {
		SetAddonIcon(ar[0], ar[1]);
	}
}

FocusFV2 = function (Id) {
	const hFocus = api.GetFocus();
	if (hFocus == te.hwnd || api.IsChild(te.hwnd, hFocus)) {
		const FV = Id ? te.Ctrl(CTRL_TC, Id).Selected : GetFolderView();
		if (FV) {
			if (!api.PathMatchSpec(api.GetClassName(hFocus), WC_EDIT + ";" + WC_TREEVIEW)) {
				FV.Focus();
			}
		}
	}
}

SetMenuExec = function (n, strName, strMenu, nPos, strExec) {
	if (!Common[n]) {
		Common[n] = api.CreateObject("Object");
	}
	Common[n].strName = strName;
	Common[n].strMenu = strMenu;
	Common[n].nPos = GetNum(nPos);
	new AsyncFunction(['AddEvent(Common.', n, '.strMenu, function (Ctrl, hMenu, nPos) {\n',
		'api.InsertMenu(hMenu, Common.', n, '.nPos, MF_BYPOSITION | MF_STRING, ++nPos, Common.', n, '.strName);\n',
		'ExtraMenuCommand[nPos] = function () {\n',
		'	InvokeUI("', strExec || "Addons." + n + ".Exec", '", Array.apply(null, arguments));\n',
		'	return S_OK;\n',
		'}\n',
		'return nPos;\n',
	'});'].join(""))();
}

InitAddonsXML = function () {
	const s = ReadXmlFile("addons.xml", false, true);
	const xml = api.CreateObject("Msxml2.DOMDocument");
	xml.async = false;
	xml.loadXML(s);
	te.Data.Addons = xml;
	return window.chrome && s;
}

InitCode = function () {
	api.ShowWindow(te.hwnd, te.CmdShow);
	if (te.CmdShow == SW_SHOWNOACTIVATE) {
		api.SetWindowPos(te.hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
	}
	te.CmdShow = SW_SHOWNORMAL;
	const types = {
		Key: ["All", "List", "Tree", "Edit", "Menus"],
		Mouse: ["All", "List", "List_Background", "Tree", "Tabs", "Tabs_Background"]
	};
	for (let i = 0; i < 3; ++i) {
		g_.KeyState[i][0] = api.GetKeyNameText(g_.KeyState[i][0]);
	}
	let i = g_.KeyState.length;
	while (i-- > 4 && g_.KeyState[i][0] == g_.KeyState[i - 4][0]) {
		g_.KeyState.pop();
	}
	for (let i = 384; i > 0; i--) {
		if (!(i & 128)) {
			const v = api.GetKeyNameText(i * 0x10000);
			if (v && v.charCodeAt(0) > 32) {
				g_.KeyCode[v.toUpperCase()] = i;
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
	te.Data.Conf_Gestures = GetNum(te.Data.Conf_Gestures, 2);
	if ("string" === typeof te.Data.Conf_TrailColor) {
		te.Data.Conf_TrailColor = GetWinColor(te.Data.Conf_TrailColor);
	}
	te.Data.Conf_TrailColor = GetNum(te.Data.Conf_TrailColor, 0xff00);
	te.Data.Conf_TrailSize = GetNum(te.Data.Conf_TrailSize, 2);
	te.Data.Conf_GestureTimeout = GetNum(te.Data.Conf_GestureTimeout, 3000);
	te.Data.Conf_Layout = GetNum(te.Data.Conf_Layout, 0x80);
	te.Data.Conf_NetworkTimeout = GetNum(te.Data.Conf_NetworkTimeout, 2000);
	te.Data.Conf_WheelSelect = GetNum(te.Data.Conf_WheelSelect, 1);
	te.SizeFormat = (te.Data.Conf_SizeFormat || "").replace(/^0x/i, "");
	te.HiddenFilter = ExtractFilter(te.Data.Conf_HiddenFilter);
	te.DragIcon = !GetNum(te.Data.Conf_NoDragIcon);
	const ar = ['AutoArrange', 'DateTimeFormat', 'Layout', 'LibraryFilter', 'NetworkTimeout', 'ShowInternet', 'ViewOrder'];
	for (let i = ar.length; i--;) {
		te[ar[i]] = te.Data['Conf_' + ar[i]];
	}
	OpenMode = te.Data.Conf_OpenMode ? SBSP_NEWBROWSER : SBSP_SAMEBROWSER;
	const s = ReadTextFile(BuildPath(te.Data.DataFolder, "icons\\config.json"));
	if (s) {
		const json = JSON.parse(s);
		g_.IconExt = json.info.ext || ".png";
	}
}

InitCloud = function () {
	const cloud = ["HKCU\\Environment\\OneDrive", "HKCR\\CLSID\\{E31EA727-12ED-4702-820C-4B6445F28E1A}\\Instance\\InitPropertyBag\\TargetFolderPath", "HKCU\\SOFTWARE\\Box\\Box\\preferences\\sync_directory_path"];
	const r = [];
	for (let i = cloud.length; i--;) {
		try {
			const s = wsh.RegRead(cloud[i]);
			if (s) {
				r.unshift(PathUnquoteSpaces(s) + "*");
			}
		} catch (e) { }
	}
	ForEachWmi("winmgmts:\\\\.\\root\\cimv2", "SELECT * FROM Win32_LogicalDisk", function (item) {
		if (/Google Drive|\@gmail\.com/i.test(item.VolumeName)) {
			r.push(item.DeviceID + "\\*");
		} else if (item.DriveType == 3 && r.length) {
			const h = api.CreateFile(item.DeviceID, 0x80000000, 7, null, 3, 0x02000000, null);
			if (h != INVALID_HANDLE_VALUE) {
				if (api.PathMatchSpec(api.GetFinalPathNameByHandle(h, 0), r.join(";"))) {
					r.push(item.DeviceID + "\\*");
				}
				api.CloseHandle(h);
			}
		}
	});
	AddEnv('cloud', r.join(";"));
}

InitMenus = function () {
	te.Data.xmlMenus = OpenXml("menus.xml", false, true);
	const root = te.Data.xmlMenus.documentElement;
	if (root) {
		menus = root.childNodes;
		for (let i = menus.length; i--;) {
			items = menus[i].getElementsByTagName("Item");
			for (let j = api.ObjGetI(items, "length"); j--;) {
				a = items[j].getAttribute("Name").split(/\\t/);
				if (a.length > 1) {
					SetKeyExec("List", a[1], items[j].text, items[j].getAttribute("Type"), true);
				}
			}
		}
	}
}

InitBG = function (cl, bWC) {
	if (bWC) {
		cl = GetWinColor(cl);
	}
	const cHwnd = [te.Ctrl(CTRL_WB).hwnd, te.hwnd];
	for (let i = cHwnd.length; i--;) {
		const hOld = api.SetClassLongPtr(cHwnd[i], GCLP_HBRBACKGROUND, api.CreateSolidBrush(cl));
		if (hOld) {
			api.DeleteObject(hOld);
		}
	}
	api.SendMessage(te.hwnd, 0x2001, 0, cl);
}

InitWindow = function () {
	if (api.GetKeyState(VK_SHIFT) < 0 && api.GetKeyState(VK_CONTROL) < 0) {
		ShowOptions("Tab=Add-ons");
	}
	if (g_.xmlWindow && "string" !== typeof g_.xmlWindow) {
		LoadXml(g_.xmlWindow);
	}
	const cTC = te.Ctrls(CTRL_TC);
	if (cTC.Count == 0) {
		const TC = te.CreateCtrl(CTRL_TC, 0, 0, "100%", "100%", te.Data.Tab_Style, te.Data.Tab_Align, te.Data.Tab_TabWidth, te.Data.Tab_TabHeight);
		TC.Selected.Navigate2(HOME_PATH || "about:blank", SBSP_NEWBROWSER, te.Data.View_Type, te.Data.View_ViewMode, te.Data.View_fFlags, te.Data.View_Options, te.Data.View_ViewFlags, te.Data.View_IconSize, te.Data.Tree_Align, te.Data.Tree_Width, te.Data.Tree_Style, te.Data.Tree_EnumFlags, te.Data.Tree_RootStyle, te.Data.Tree_Root);
	} else if (te.Ctrls(CTRL_TC, true).Count == 0) {
		cTC[0].Visible = true;
	}
	g_.xmlWindow = void 0;
	if (te.Data.Load < 2) {
		RunCommandLine(api.GetCommandLine());
	} else {
		const cFV = te.Ctrls(CTRL_FV, true);
		for (let i in cFV) {
			ChangeView(cFV[i]);
		}
	}
	WebBrowser.DropMode = 1;
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
	const tm = new Date().getTime();
	if (Threads.Data.length >= Threads.nBase && tm - Threads.Data[0].Data.tm < Threads.nTI) {
		return;
	}
	const o = api.CreateObject("Object");
	o.Data = api.CreateObject("Object");
	o.Data.Threads = Threads;
	o.Data.Id = Math.random();
	o.Data.tm = tm;
	o.Data.MainWindow = window;
	Threads.Data.unshift(o);
	if (!Threads.src) {
		Threads.src = FixScript(ReadTextFile("script\\threads.js"));
	}
	api.ExecScript(Threads.src, "JScript", o, true);
}

Threads.End = function (Id) {
	for (let i = Threads.Data.Count; i--;) {
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
	te.Data.Load = 1;
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
	te.Data.Tree_RootStyle = NSTCRS_EXPANDED;
	te.Data.Tree_Root = 0;

	te.Data.Conf_TabDefault = true;
	te.Data.Conf_TreeDefault = true;
	te.Data.Conf_ListDefault = true;

	te.Data.Installed = GetParentFolderName(api.GetModuleFileName(null));
	te.Data.DataFolder = te.Data.Installed;
	te.Data.TempFolder = GetTempPath(1);
	let fn = function () {
		te.Data.DataFolder = BuildPath(api.GetDisplayNameOf(ssfAPPDATA, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING), "Tablacus\\Explorer");
		const ParentFolder = GetParentFolderName(te.Data.DataFolder);
		if (!api.PathIsDirectory(ParentFolder)) {
			try {
				if (api.CreateDirectory(ParentFolder)) {
					CreateFolder2(te.Data.DataFolder);
				}
			} catch (e) { }
		}
	}

	const pf = [ssfPROGRAMFILES, ssfPROGRAMFILESx86];
	const x = api.sizeof("HANDLE") / 4;
	for (let i = 0; i < x; ++i) {
		const s = api.GetDisplayNameOf(pf[i], SHGDN_FORADDRESSBAR | SHGDN_FORPARSING);
		const l = s.replace(/\s*\(x86\)$/i, "").length;
		if (api.StrCmpNI(s, te.Data.DataFolder, l) == 0) {
			fn();
			break;
		}
	}
	let s = BuildPath(te.Data.DataFolder, "config");
	try {
		api.CreateDirectory(s);
	} catch (e) { }
	if (!api.PathIsDirectory(s)) {
		fn();
		CreateFolder2(BuildPath(te.Data.DataFolder, "config"));
	}
	delete fn;
	if (g_.IEVer < 8) {
		s = BuildPath(te.Data.DataFolder, "cache");
		CreateFolder2(s);
		CreateFolder2(BuildPath(s, "bitmap"));
		CreateFolder2(BuildPath(s, "icon"));
		CreateFolder2(BuildPath(s, "file"));
	}
	te.Data.Conf_Lang = GetLangId();
	const SHIL = api.CreateObject("Array");
	const SHILS = api.CreateObject("Array");
	for (let i = 0; i < SHIL_JUMBO; ++i) {
		const il = api.SHGetImageList(i);
		SHIL.push(il);
		SHILS.push(api.Memory("SIZE"));
		api.ImageList_GetIconSize(SHIL[i], SHILS[i]);
	}
	te.Data.SHIL = SHIL;
	te.Data.SHILS = SHILS;
	const o = {};
	try {
		const sw = sha.Windows();
		for (let i = 0; i < sw.Count; ++i) {
			const x = sw.item(i);
			if (x && x.Document) {
				const w = x.Document.parentWindow;
				if (w && w.te && w.te.Data) {
					o[w.te.Data.WindowSetting] = 1;
				}
			}
		}
	} catch (e) { }
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
	++te.Data.Load;
	for (let i in te.Data) {
		if (/^xml/.test(i)) {
			delete te.Data[i];
		}
	}
	LoadConfig();
	delete g_.xmlWindow;
}
Exchange = te.Data.Exchange;
