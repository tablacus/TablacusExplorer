AddEvent("MouseMessage", function (Ctrl, hwnd, msg, mouseData, pt, wHitTestCode, dwExtraInfo) {
	if (msg == WM_MOUSEMOVE) {
		if (Common.MainMenu.bLoop && Ctrl.Type == CTRL_TE) {
			const Ctrl2 = te.CtrlFromPoint(pt);
			if (Ctrl2 && Ctrl2.Type == CTRL_WB) {
				if (!PtInRect(Common.MainMenu.Item, pt)) {
					for (let i = Common.MainMenu.Items.length; i--;) {
						if (PtInRect(Common.MainMenu.Items[i], pt)) {
							Common.MainMenu.bClose = false;
							api.PostMessage(hwnd, WM_KEYDOWN, VK_ESCAPE, 0);
							Common.MainMenu.Popup(Common.MainMenu.Menu[i]);
							break;
						}
					}
				}
			}
 		}
	}
});
