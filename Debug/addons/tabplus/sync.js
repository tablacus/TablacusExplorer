Common.TabPlus.rc = api.CreateObject("Object");
Common.TabPlus.rcItem = api.CreateObject("Object");

Sync.TabPlus = {
	FromPt: function (Id, pt) {
		var ptc = pt.Clone();
		api.ScreenToClient(WebBrowser.hwnd, ptc);
		var Items = Common.TabPlus.rcItem[Id];
		var TC = te.Ctrl(CTRL_TC, Id);
		if (TC) {
			for (var i = Math.min(TC.Count, Items.length); i-- > 0;) {
				if (PtInRect(Items[i], ptc)) {
					return i;
				}
			}
		}
		return -1;
	},

	TCFromPt: function (pt) {
		var ptc = pt.Clone();
		api.ScreenToClient(WebBrowser.hwnd, ptc);
		for (var Id in Common.TabPlus.rc) {
			var TC = te.Ctrl(CTRL_TC, Id);
			if (TC.Visible && PtInRect(Common.TabPlus.rc[Id], ptc)) {
				return TC;
			}
		}
	}
}

AddEvent("HitTest", function (Ctrl, pt, flags) {
	if (Ctrl.Type == CTRL_TC) {
		return Sync.TabPlus.FromPt(Ctrl.Id, pt);
	}
});

AddEvent("DragEnter", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	if (Ctrl.Type == CTRL_WB) {
		pdwEffect[0] = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
		return S_OK;
	}
});

AddEvent("DragOver", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	if (Ctrl.Type == CTRL_WB) {
		var TC = Sync.TabPlus.TCFromPt(pt);
		if (TC) {
			var nIndex = Sync.TabPlus.FromPt(TC.Id, pt);
			if (nIndex >= 0) {
				if (!g_.ptDrag || IsDrag(pt, g_.ptDrag)) {
					var FV = TC[nIndex];
					if (!Common.TabPlus.opt.NoDragOpen || (!FV.hwndView && FV.FolderItem.Enum)) {
						g_.ptDrag = pt.Clone();
						InvokeUI("Addons.TabPlus.DragOver", TC.Id, pt);
					}
				}
				if (Common.TabPlus.Drag5) {
					pdwEffect[0] = DROPEFFECT_MOVE;
					return S_OK;
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
		} else if (Common.TabPlus.Drag5) {
			pdwEffect[0] = DROPEFFECT_NONE;
			return S_OK;
		}
	}
});

AddEvent("Drop", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
	if (Ctrl.Type == CTRL_WB) {
		var TC = Sync.TabPlus.TCFromPt(pt);
		if (TC) {
			var nIndex = Sync.TabPlus.FromPt(TC.Id, pt);
			if (Common.TabPlus.Drag5) {
				var res = /^tabplus_(\d+)_(\d+)/.exec(Common.TabPlus.Drag5);
				if (res) {
					if (nIndex < 0) {
						nIndex = TC.Count - 1;
					}
					if (res[1] != TC.Id || res[2] != nIndex) {
						var TC1 = te.Ctrl(CTRL_TC, res[1]);
						TC1.Move(res[2], nIndex, TC);
						TC1.SelectedIndex = nIndex;
					}
				}
				Common.TabPlus.Drag5 = void 0;
				return S_OK;
			}
			if (nIndex >= 0) {
				var hr = S_FALSE;
				var DropTarget = TC[nIndex].DropTarget;
				if (DropTarget) {
					InvokeUI("Addons.TabPlus.DragLeave");
					Common.TabPlus.bDropping = true;
					hr = DropTarget.Drop(dataObj, grfKeyState, pt, pdwEffect);
					Common.TabPlus.bDropping = false;
				}
				return hr;
			} else if (dataObj.Count) {
				for (var i = 0; i < dataObj.Count; ++i) {
					var FV = TC.Selected.Navigate(dataObj.Item(i), SBSP_NEWBROWSER);
					TC.Move(FV.Index, TC.Count - 1);
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
	if (Ctrl.Type == CTRL_TE && msg == WM_ACTIVATE) {
		InvokeUI("Addons.TabPlus.GetRects");
	}
});
