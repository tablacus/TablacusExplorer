var Addon_Id = "addressbar";
var item = GetAddonElement(Addon_Id);

//Menu
if (Common.AddressBar.MenuExec) {
    AddEvent(item.getAttribute("Menu"), function (Ctrl, hMenu, nPos) {
        api.InsertMenu(hMenu, Common.AddressBar.nPos, MF_BYPOSITION | MF_STRING, ++nPos, Common.AddressBar.strName);
		ExtraMenuCommand[nPos] = function () {
			InvokeUI("Addons.AddressBar.Exec", arguments);
			return S_OK;
		}
		return nPos;
    });
}

AddEvent("MouseMessage", function (Ctrl, hwnd, msg, mouseData, pt, wHitTestCode, dwExtraInfo) {
	if (msg == WM_MOUSEMOVE && Ctrl.Type == CTRL_TE && Common.AddressBar.rcItem) {
		var Ctrl2 = te.CtrlFromPoint(pt);
		if (Ctrl2 && Ctrl2.Type == CTRL_WB) {
			var ptc = pt.Clone();
			api.ScreenToClient(WebBrowser.hwnd, ptc);
			for (var i = Common.AddressBar.rcItem.length; i-- > 0;) {
				if (PtInRect(Common.AddressBar.rcItem[i], ptc)) {
					api.PostMessage(hwnd, WM_KEYDOWN, VK_ESCAPE, 0);
					InvokeUI("Addons.AddressBar.ChangeMenu", i);
					break;
				}
			}
		}
	}
});

if (!window.chrome) {
	AddEvent("DragEnter", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
		if (Ctrl.Type == CTRL_WB) {
			return S_OK;
		}
	});

	AddEvent("DragOver", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
		if (Ctrl.Type == CTRL_WB && dataObj.Count) {
			var ptc = pt.Clone();
			api.ScreenToClient(api.GetWindow(document), ptc);
			var el = document.elementFromPoint(ptc.x, ptc.y);
			if (el) {
				var res = /^addressbar(\d+)_$/.exec(el.id);
				if (res) {
					var Target = Common.AddressBar.GetPath(res[1] - 0);
					if (!api.ILIsEqual(dataObj.Item(-1), Target)) {
						var DropTarget = api.DropTarget(Target);
						if (DropTarget) {
							return DropTarget.DragOver(dataObj, grfKeyState, pt, pdwEffect);
						}
					}
					pdwEffect[0] = DROPEFFECT_NONE;
					return S_OK;
				}
			}
		}
	});

	AddEvent("Drop", function (Ctrl, dataObj, grfKeyState, pt, pdwEffect) {
		if (Ctrl.Type == CTRL_WB && dataObj.Count) {
			var ptc = pt.Clone();
			api.ScreenToClient(api.GetWindow(document), ptc);
			var el = document.elementFromPoint(ptc.x, ptc.y);
			if (el) {
				var res = /^addressbar(\d+)_$/.exec(el.id);
				if (res) {
					var hr = S_FALSE;
					var Target = Common.AddressBar.GetPath(res[1] - 0);
					if (!api.ILIsEqual(dataObj.Item(-1), Target)) {
						var DropTarget = api.DropTarget(Target);
						if (DropTarget) {
							hr = DropTarget.Drop(dataObj, grfKeyState, pt, pdwEffect);
						}
					}
					return hr;
				}
			}
		}
	});

	AddEvent("DragLeave", function (Ctrl) {
		return S_OK;
	});
}
