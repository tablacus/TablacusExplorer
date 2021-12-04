g_Types = { Mouse: ["All", "List", "List_Background", "Tree", "Tabs", "Tabs_Background"] };

const src = await ReadTextFile("addons\\" + Addon_Id + "\\options.html");
const ar = [];
const s = "CSA";
for (let i = s.length; i--;) {
	ar.unshift('<input type="button" value="', await MainWindow.g_.KeyState[i][0], '" title="', s.charAt(i), '" onclick="AddMouse(this)">');
}
await SetTabContents(4, "", src.replace("%s", ar.join("")));

AddMouse = function (o) {
	document.E.MouseMouse.value += o.title;
	ChangeX("Mouse");
}

SaveLocation = async function () {
	SetChanged(null, document.E);
	await SaveX("Mouse", document.E);
}

LoadX("Mouse", null, document.E);
