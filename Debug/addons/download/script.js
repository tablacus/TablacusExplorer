const Addon_Id = 'download';
if (window.Addon == 1) {
	$.importScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	SetTabContents(0, "", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
	const el = document.getElementById("_curl");
	if (await fso.FileExists(BuildPath(system32, "curl.exe"))) {
		el.style.display = "none";
	}
	SetDL = function (el) {
		ConfirmThenExec(el.innerText || GetText("Select"), function () {
			document.F.Path.value = el.title;
		});
	}
}
