g_Types = { Key: ["All", "List", "Tree", "Browser"] };

var ado = OpenAdodbFromTextFile("addons\\" + Addon_Id + "\\options.html");
if (ado) {
	SetTabContents(4, "", ado.ReadText(adReadAll));
	ado.Close();
}

SaveLocation = function ()
{
	SetChanged(null, document.E);
	SaveX("Key", document.E);
}

LoadX("Key", null, document.E);
MakeKeySelect();
