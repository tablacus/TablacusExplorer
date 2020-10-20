var ado = OpenAdodbFromTextFile("addons\\" + Addon_Id + "\\options.html");
if (ado) {
	var ar = ado.ReadText(adReadAll).split("<!--panel-->");;
	SetTabContents(0, "View", ar[0]);
	SetTabContents(4, "General", ar[1]);
	ado.Close();
}

SaveLocation = function () {
	if (g_Chg["List"]) {
		var xml = CreateXml();
		var root = xml.createElement("TablacusExplorer");
		var o = document.E.List;
		for (var i = 0; i < o.length; i++) {
			var item = xml.createElement("Item");
			var a = o[i].value.split(g_sep);
			item.setAttribute("Name", a[0]);
			item.text = a[1];
			item.setAttribute("Type", a[2]);
			item.setAttribute("Icon", a[3]);
			item.setAttribute("Height", a[4]);
			root.appendChild(item);
		}
		xml.appendChild(root);
		SaveXmlEx(AddonName.toLowerCase() + ".xml", xml);
		te.Data["xml" + AddonName] = xml;
		xml = null;
	}
}

EditTB = function (mode) {
	if (g_x.List.selectedIndex < 0) {
		return;
	}
	ClearX("List");
	var a = g_x.List[g_x.List.selectedIndex].value.split(g_sep);
	SetType(document.E.Type, a[2]);
	var ix = ["Name", "Path", "Type", "Icon", "Height"];
	for (var i = ix.length - 1; i >= 0; i--) {
		if (i != 2) {
			document.E.elements[ix[i]].value = a[i];
		}
	}
	var p = { s: a[1] };
	MainWindow.OptionDecode(a[2], p);
	document.E.Path.value = p.s;
	SetImage(document.E, "_IconE");
}

ReplaceTB = function (mode) {
	ClearX();
	if (g_x.List.selectedIndex < 0) {
		g_x.List.selectedIndex = ++g_x.List.length - 1;
		EnableSelectTag(g_x.List);
	}
	var sel = g_x.List[g_x.List.selectedIndex];
	var o = document.E.Type;
	var p = { s: document.E.Path.value };
	MainWindow.OptionEncode(o[o.selectedIndex].value, p);
	SetData(sel, [document.E.Name.value, p.s, o[o.selectedIndex].value, document.E.Icon.value, document.E.Height.value]);
	g_Chg[mode] = true;
}

AddonName = Addon_Id;
LoadX("List", dialogArguments.Data.nEdit ? function () {
	g_x.List.selectedIndex = dialogArguments.Data.nEdit - 1;
	EditTB();
} : null, document.E);
