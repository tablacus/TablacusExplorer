if (window.Addon == 1) {
	AddEvent("ChangeView", function (Ctrl) {
		if (Ctrl.Id == Ctrl.Parent.Selected.Id && Ctrl.Parent.Id == te.Ctrl(CTRL_TC).Id) {
			document.title = Ctrl.Title + ' - ' + TITLE;
		}
	});
}
