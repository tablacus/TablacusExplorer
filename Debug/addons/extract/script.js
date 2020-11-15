var Addon_Id = 'extract';

if (window.Addon == 1) {
	importJScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	SetTabContents(0, "", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
	document.getElementById("_7zip").value = await api.sprintf(99, await GetText("Get %s..."), "7-Zip");

	SetExtract = async function (o) {
		if (await confirmOk()) {
			document.F.Path.value = o.title;
		}
	}

	SetExe = async function (o) {
		if (await confirmOk()) {
			var ar = o.title.split(" ")
			var path = 'C:\\Program Files\\' + ar[0];
			var path2 = 'C:\\Program Files (x86)\\' + ar[0];
			ar[0] = await api.PathQuoteSpaces(!await $.fso.FileExists(path) && await $.fso.FileExists(path2) ? path2 : path);
			document.F.Path.value = ar.join(" ");
		}
	}
}
