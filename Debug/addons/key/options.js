g_Types = { Key: ["All", "List", "Tree", "Browser", "Edit", "Menus"] };

var src = await ReadTextFile(BuildPath("addons", Addon_Id, "options.html"));
if (src) {
	await SetTabContents(4, "", src);
}

SaveLocation = function () {
	SetChanged(null, document.E);
	SaveX("Key", document.E);
}

LoadX("Key", null, document.E);
MakeKeySelect();
