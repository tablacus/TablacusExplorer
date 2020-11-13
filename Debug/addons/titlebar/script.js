if (window.Addon == 1) {
	AddEvent("ChangeView", async function (Ctrl) {
		if (await Ctrl.Id == await Ctrl.Parent.Selected.Id && await Ctrl.Parent.Id == await te.Ctrl(CTRL_TC).Id) {
			document.title = await Ctrl.Title + ' - ' + TITLE;
		}
	});
}
