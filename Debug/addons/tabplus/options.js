const Icon = document.F.Icon;
if (Icon) {
	Icon.name = "Icon_0";
}
await SetTabContents(0, "", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
document.getElementById("_Drive").innerHTML = (await api.LoadString(hShell32, 4122)).replace(/ %c:?/, "");
document.getElementById("_DropTo").innerHTML = (await GetTextR("@SRH.dll,-8110[Drop to %1]")).replace(/%1/, await GetText("Folder"));
