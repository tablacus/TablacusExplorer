var g_x = {Mouse: null};
var g_Chg = {Mouse: false, Data: null};
g_Types = {Mouse: ["All", "List", "List_Background", "Tree", "Tabs", "Tabs_Background", "Browser"]};
g_nResult = 3;
g_bChanged = false;

var ado = OpenAdodbFromTextFile(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\mouse\\options.html"));
if (ado) {
	SetTabContents(4, "General", ado.ReadText(adReadAll));
	ado.Close();
}

AddMouse = function(o)
{
	document.E.MouseMouse.value += o.title;
	ChangeX("Mouse");
}

SaveLocation = function ()
{
	SetChanged(null, document.E);
	SaveX("Mouse", document.E);
}

var ar = [];
var s = "CSA";
for (var i = 0; i < s.length; i++) {
	ar.push('<input type="button" value="', MainWindow.g_.KeyState[i][0],'" title="', s.charAt(i), '" onclick="AddMouse(this)" />');
}
document.getElementById("__MOUSEDATA").innerHTML = ar.join("");
LoadLang2(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\mouse\\lang\\" + te.Data.Conf_Lang + ".xml"));
ApplyLang(document);
LoadX("Mouse", null, document.E);