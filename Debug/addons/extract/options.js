SetTabContents(0, "", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
document.getElementById("_7zip").value = await api.sprintf(99, await GetText("Get %s..."), "7-Zip");

SetExtract = async function (o) {
	if (await confirmOk()) {
		document.F.Path.value = o.title;
	}
}

SetExe = async function (o) {
	if (await confirmOk()) {
		const ar = o.title.split(" ")
		const path = 'C:\\Program Files\\' + ar[0];
		const path2 = 'C:\\Program Files (x86)\\' + ar[0];
		ar[0] = await PathQuoteSpaces(!await fso.FileExists(path) && await fso.FileExists(path2) ? path2 : path);
		document.F.Path.value = ar.join(" ");
	}
}
