if (window.Addon == 1) {
	AddEvent("ChangeView", function(Ctrl)
	{
		try {
			setTimeout(function ()
			{
				if (Ctrl.Id == Ctrl.Parent.Selected.Id && Ctrl.Parent.Id == te.Ctrl(CTRL_TC).Id) {
					api.SetWindowText(te.hwnd, Ctrl.Title + ' - Tablacus Explorer');
				}
			}, 99);
		} catch (e) {}
	});
}
