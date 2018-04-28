g_nResult = 3;
g_bChanged = false;

function ShowLocation()
{
	ShowLocationEx({id: AddonName.toLowerCase(), show: "6", index: "6"});
}

ShowIcon = function ()
{
	var s = ShowIconEx();
	if (s) {
		document.F.Icon.value = s;
		ChangeX("List");
		SetImage();
	}
}

function SaveTB(mode)
{
	if (g_Chg[mode]) {
		var xml = CreateXml();
		var root = xml.createElement("TablacusExplorer");
		var o = document.F.List;
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

function EditTB(mode)
{
	if (g_x.List.selectedIndex < 0) {
		return;
	}
	ClearX("List");
	var a = g_x.List[g_x.List.selectedIndex].value.split(g_sep);
	SetType(document.F.Type, a[2]);
	var ix = ["Name", "Path", "Type", "Icon", "Height"];
	for (var i = ix.length - 1; i >= 0; i--) {
		if (i != 2) {
			document.F.elements[ix[i]].value = a[i];
		}
	}
	var p = { s: a[1] };
	MainWindow.OptionDecode(a[2], p);
	document.F.Path.value = p.s;
	SetImage();
}

function ReplaceTB(mode)
{
	ClearX();
	if (g_x.List.selectedIndex < 0) {
		g_x.List.selectedIndex = ++g_x.List.length - 1;
		EnableSelectTag(g_x.List);
	}
	var sel = g_x.List[g_x.List.selectedIndex];
	var o = document.F.Type;
	var p = { s: document.F.Path.value };
	MainWindow.OptionEncode(o[o.selectedIndex].value, p);
	SetData(sel, [document.F.Name.value, p.s, o[o.selectedIndex].value, document.F.Icon.value, document.F.Height.value]);
	g_Chg[mode] = true;
}

function SetToolOptions(n)
{
	g_bChanged |= g_Chg.List;
	if (n === 1) {
		g_nResult = 1;
	}
	if (SetOptions(function () {
		SetChanged(ReplaceTB);
		SaveTB("List");
		TEOk();
	}, null, n === 1 ? 2 : 0)) {
		if (n === 1) {
			window.close();
		}
	}
}

ApplyLang(document);
var info = GetAddonInfo(Addon_Id);
document.title = info.Name;
LoadX("List", dialogArguments.Data.nEdit ? function () {
	document.F.List.selectedIndex = dialogArguments.Data.nEdit - 1;
	EditTB();
} : null);

AddEventEx(window, "beforeunload", SetToolOptions);
