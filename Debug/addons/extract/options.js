SetTabContents(0, "", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));

SetExe = async function (o) {
	ConfirmThenExec(o.innerText, async function () {
		const ar = o.title.split(" ")
		const path = 'C:\\Program Files\\' + ar[0];
		const path2 = 'C:\\Program Files (x86)\\' + ar[0];
		ar[0] = PathQuoteSpaces(!await fso.FileExists(path) && await fso.FileExists(path2) ? path2 : path);
		document.F.Path.value = ar.join(" ");
	});
}
