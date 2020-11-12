g_Types = { Mouse: ["All", "List", "List_Background", "Tree", "Tabs", "Tabs_Background", "Browser"] };

var src = await ReadTextFile(BuildPath("addons", Addon_Id, "options.html"));
var ar = [];
var s = "CSA";
for (var i = s.length; i--;) {
	ar.unshift('<input type="button" value="', await MainWindow.g_.KeyState[i][0], '" title="', s.charAt(i), '" onclick="AddMouse(this)">');
}
await SetTabContents(4, "", src.replace("%s", ar.join("")));

AddMouse = function (o) {
	document.E.MouseMouse.value += o.title;
	ChangeX("Mouse");
}

SaveLocation = function () {
	SetChanged(null, document.E);
	SaveX("Mouse", document.E);
}

LoadX("Mouse", null, document.E);
