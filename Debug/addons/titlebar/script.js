if (window.Addon == 1) {
	AddEvent("ChangeView1", async function (Ctrl) {
		document.title = await Ctrl.Title + ' - ' + TITLE;
	});
}
