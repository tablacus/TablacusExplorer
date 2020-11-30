g_Types = { Key: ["All", "List", "Tree", "Browser", "Edit", "Menus"] };

var src = await ReadTextFile(BuildPath("addons", Addon_Id, "options.html"));
if (src) {
	await SetTabContents(4, "", src);
}

SaveLocation = async function () {
	SetChanged(null, document.E);
	await SaveX("Key", document.E);
}

LoadX("Key", null, document.E);
MakeKeySelect();
