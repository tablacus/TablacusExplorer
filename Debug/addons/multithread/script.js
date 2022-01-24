const Addon_Id = "multithread";
const item = GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("Copy", 1);
	item.setAttribute("Move", 1);
	item.setAttribute("Delete", 1);
}
if (window.Addon == 1) {
	$.importScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	SetTabContents(0, "", '<label><input type="checkbox" id="Copy">Copy</label><br><label><input type="checkbox" id="Move">Move</label><br><label><input type="checkbox" id="Delete">Delete</label><br><label><input type="checkbox" id="!NoTemp">@shell32.dll,-21815[Temporary Burn Folder]</label><br>');
}
