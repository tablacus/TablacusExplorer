var Addon_Id = 'download';
if (window.Addon == 1) {
	importJScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	SetTabContents(0, "", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
	const el = document.getElementById("_curl");
	if (await fso.FileExists(BuildPath(system32, "curl.exe"))) {
		el.style.display = "none";
	} else {
		el.value = await api.sprintf(99, await GetText("Get %s..."), "cURL");
	}
	document.getElementById("_wget").value = await api.sprintf(99, await GetText("Get %s..."), "Wget");

	SetDL = async function (o) {
		if (await confirmOk(await GetText())) {
			document.F.Path.value = o.title;
		}
	}
}
