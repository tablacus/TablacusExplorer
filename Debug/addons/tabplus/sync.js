Common.TabPlus.rc = api.CreateObject("Object");
Common.TabPlus.rcItem = api.CreateObject("Object");

Sync.TabPlus = {
	DropTo: !GetAddonOption("tabplus", "NoDropTo"),

	FromPt: function (Id, pt) {
		const ptc = pt.Clone();
		api.ScreenToClient(WebBrowser.hwnd, ptc);
		const Items = Common.TabPlus.rcItem[Id];
		const TC = te.Ctrl(CTRL_TC, Id);
		if (TC) {
			for (let i = Math.min(Items.length, TC.Count); i-- > 0;) {
				if (PtInRect(Items[i], ptc)) {
					return i;
				}
			}
		}
		return -1;
	},

	TCFromPt: function (pt) {
		const ptc = pt.Clone();
		api.ScreenToClient(WebBrowser.hwnd, ptc);
		for (let Id in Common.TabPlus.rc) {
			const TC = te.Ctrl(CTRL_TC, Id);
			if (TC.Visible && PtInRect(Common.TabPlus.rc[Id], ptc)) {
				return TC;
			}
		}
	},

	DropTab: function (FV, hwnd, pt) {
		const sw = sha.Windows();
		for (let i = 0; i < sw.Count; ++i) {
			const x = sw.item(i);
			if (x) {
				const w = x.Document.parentWindow;
				if (w && w.te && w.te.hwnd) {
					if (w.te.hwnd || hwnd, api.IsChild(w.te.hwnd, hwnd)) {
						const ptc = pt.Clone();
						api.ScreenToClient(w.WebBrowser.hwnd, ptc);
						const rc = api.Memory("RECT");
						api.GetClientRect(w.te.hwnd, rc);
						rc.left += w.te.offsetLeft;
						rc.top += w.te.offsetTop;
						rc.right -= w.te.offsetRight;
						rc.bottom -= w.te.offsetBottom;
						const cTC = w.te.Ctrls(CTRL_TC, true);
						let destFV;
						for (let i = 0; i < cTC.length; ++i) {
							const TC = cTC[i];
							const left = Sync.TabPlus.GetData(rc, TC.Left);
							const top = Sync.TabPlus.GetData(rc, TC.Top);
							const right = left + Sync.TabPlus.GetData(rc, TC.Width);
							const bottom = top + Sync.TabPlus.GetData(rc, TC.Height);
							if (ptc.x >= left && ptc.x < right && ptc.y >= top && ptc.y < bottom) {
								destFV = TC.Selected;
								break;
							}
						}
						w.NavigateFV(destFV || te.Ctrl(CTRL_FV), FV.FolderItem.Path, SBSP_NEWBROWSER);
						if (api.GetKeyState(VK_CONTROL) >= 0) {
							FV.Close();
						}
						return;
					}
				}
			}
		}
		($.OpenInNewWindow || OpenInExplorer)(FV);
		if (api.GetKeyState(VK_CONTROL) >= 0) {
			FV.Close();
		}
	},

	GetData: function (rc, i) {
		const res = /^([\d\.]+)%$/.exec(i);
		if (res) {
			i = (rc.right - rc.left) * res[1] / 100;
		}
		return i;
	}
}

AddEvent("HitTest", function (Ctrl, pt, flags) {
	if (Ctrl.Type == CTRL_TC) {
		return Sync.TabPlus.FromPt(Ctrl.Id, pt);
	}
});

AddEvent("DragEnter", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	if (Ctrl.Type == CTRL_WB) {
		InvokeUI("Addons.TabPlus.SetRects");
		pdwEffect[0] = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
		return S_OK;
	}
});

AddEvent("DragOver", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	if (Ctrl.Type == CTRL_WB) {
		const TC = Sync.TabPlus.TCFromPt(pt);
		if (TC) {
			if (Common.TabPlus.Drag5 || /\ntabplus_\d+_\d+$/.test(dataObj.GetData())) {
				pdwEffect[0] = DROPEFFECT_MOVE;
				return S_OK;
			}
			const nIndex = Sync.TabPlus.FromPt(TC.Id, pt);
			if (nIndex >= 0) {
				const FV = TC[nIndex];
				if (FV) {
					if (!g_.ptDrag || IsDrag(pt, g_.ptDrag)) {
						if (!Common.TabPlus.opt.NoDragOpen || (!FV.hwndView && FV.FolderItem.Enum)) {
							g_.ptDrag = pt.Clone();
							InvokeUI("Addons.TabPlus.DragOver", TC.Id);
						}
					}
					if (dataObj.Count && Sync.TabPlus.DropTo) {
						const Target = FV.FolderItem;
						if (!api.ILIsEqual(dataObj.Item(-1), Target)) {
							let DropTarget = api.DropTarget(Target);
							if (DropTarget) {
								return DropTarget.DragOver(dataObj, grfKeyState, pt, pdwEffect);
							}
						}
					}
				}
				if (Sync.TabPlus.DropTo) {
					pdwEffect[0] = DROPEFFECT_NONE;
					return S_OK;
				}
			}
			if (dataObj.Count && dataObj.Item(0).IsFolder) {
				pdwEffect[0] = DROPEFFECT_LINK;
				return S_OK;
			}
		}
	}
});

AddEvent("Drop", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	if (Ctrl.Type == CTRL_WB) {
		const TC = Sync.TabPlus.TCFromPt(pt);
		if (TC) {
			let nIndex = Sync.TabPlus.FromPt(TC.Id, pt);
			if (Common.TabPlus.Drag5) {
				const res = /^tabplus_(\d+)_(\d+)/.exec(Common.TabPlus.Drag5);
				if (res) {
					if (nIndex < 0) {
						nIndex = TC.Count;
					}
					if (res[1] != TC.Id || res[2] != nIndex) {
						const TC1 = te.Ctrl(CTRL_TC, res[1]);
						TC1.Move(res[2], nIndex, TC);
						TC1.SelectedIndex = nIndex;
					}
				}
				Common.TabPlus.Drag5 = void 0;
				return S_OK;
			}
			if (nIndex >= 0 && Sync.TabPlus.DropTo) {
				let hr = S_FALSE;
				const DropTarget = TC[nIndex].DropTarget;
				if (DropTarget) {
					InvokeUI("Addons.TabPlus.DragLeave");
					Common.TabPlus.bDropping = true;
					hr = DropTarget.Drop(dataObj, grfKeyState, pt, pdwEffect);
					Common.TabPlus.bDropping = false;
				}
				return hr;
			}
			if (dataObj.Count) {
				for (let i = 0; i < dataObj.Count; ++i) {
					const FV = TC.Selected.Navigate(dataObj.Item(i), SBSP_NEWBROWSER);
					TC.Move(FV.Index, nIndex >= 0 ? nIndex : TC.Count - 1);
				}
				return S_OK;
			}
		}
	}
});

AddEvent("DragLeave", function (Ctrl) {
	InvokeUI("Addons.TabPlus.DragLeave");
	return S_OK;
});

AddEvent("SystemMessage", function (Ctrl, hwnd, msg, wParam, lParam) {
	if (Ctrl.Type == CTRL_TE && (msg == WM_ACTIVATE || msg == WM_ACTIVATEAPP)) {
		InvokeUI("Addons.TabPlus.SetRects");
	}
});
