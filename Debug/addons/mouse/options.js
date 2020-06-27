g_Types = { Mouse: ["All", "List", "List_Background", "Tree", "Tabs", "Tabs_Background", "Browser"] };

var ado = OpenAdodbFromTextFile("addons\\" + Addon_Id + "\\options.html");
if (ado) {
	var ar = [];
	var s = "CSA";
	for (var i = s.length; i--;) {
		ar.unshift('<input type="button" value="', MainWindow.g_.KeyState[i][0], '" title="', s.charAt(i), '" onclick="AddMouse(this)" />');
	}
	SetTabContents(4, "", ado.ReadText(adReadAll).replace("%s", ar.join("")));
	ado.Close();
}

AddMouse = function (o) {
	document.E.MouseMouse.value += o.title;
	ChangeX("Mouse");
}

SaveLocation = function () {
	SetChanged(null, document.E);
	SaveX("Mouse", document.E);
}

LoadX("Mouse", null, document.E);
