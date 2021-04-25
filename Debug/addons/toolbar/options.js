const ar = (await ReadTextFile("addons\\" + Addon_Id + "\\options.html")).split("<!--panel-->");
SetTabContents(0, "View", ar[0]);
SetTabContents(4, "General", ar[1]);

SaveLocation = async function () {
	if (g_bChanged && document.E.Name.value) {
		await ReplaceTB('List');
	}
	if (g_Chg["List"]) {
		const xml = await CreateXml();
		const root = await xml.createElement("TablacusExplorer");
		const o = document.E.List;
		for (let i = 0; i < o.length; i++) {
			const item = await xml.createElement("Item");
			const a = o[i].value.split(g_sep);
			item.text = a[1];
			await Promise.all([item.setAttribute("Name", a[0]), item.setAttribute("Type", a[2]), item.setAttribute("Icon", a[3]), item.setAttribute("Height", a[4])]);
			await root.appendChild(item);
		}
		await xml.appendChild(root);
		await SaveXmlEx(Addon_Id + ".xml", xml);
		te.Data["xml" + AddonName] = xml;
	}
}

EditTB = async function () {
	if (g_x.List.selectedIndex < 0) {
		return;
	}
	ClearX("List");
	const a = g_x.List[g_x.List.selectedIndex].value.split(g_sep);
	SetType(document.E.Type, a[2]);
	const ix = ["Name", "Path", "Type", "Icon", "Height"];
	for (let i = ix.length - 1; i >= 0; i--) {
		if (i != 2) {
			document.E.elements[ix[i]].value = a[i];
		}
	}
	const p = await api.CreateObject("Object");
	p.s = a[1];
	await MainWindow.OptionDecode(a[2], p);
	document.E.Path.value = await p.s;
	SetImage(document.E, "_IconE");
}

ReplaceTB = async function (mode) {
	ClearX();
	if (g_x.List.selectedIndex < 0) {
		g_x.List.selectedIndex = ++g_x.List.length - 1;
		EnableSelectTag(g_x.List);
	}
	const sel = g_x.List[g_x.List.selectedIndex];
	const o = document.E.Type;
	const p = await api.CreateObject("Object");
	p.s = document.E.Path.value;
	await MainWindow.OptionEncode(o[o.selectedIndex].value, p);
	SetData(sel, [document.E.Name.value, await p.s, o[o.selectedIndex].value, document.E.Icon.value, document.E.Height.value]);
	g_Chg[mode] = true;
}

LoadX("List", await dialogArguments.Data.nEdit ? async function () {
	g_x.List.selectedIndex = await dialogArguments.Data.nEdit - 1;
	EditTB();
} : null, document.E);
