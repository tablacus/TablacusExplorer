var Addon_Id = "undoclosetab";
var Default = "None";

var item = await GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("MenuExec", 1);
	item.setAttribute("Menu", "Tabs");
	item.setAttribute("MenuPos", 0);

	item.setAttribute("KeyExec", 1);
	item.setAttribute("Key", "Shift+Ctrl+T");
	item.setAttribute("KeyOn", "All");

	item.setAttribute("MouseExec", 1);
	item.setAttribute("Mouse", "3");
	item.setAttribute("MouseOn", "Tabs_Background");
}
if (window.Addon == 1) {
	Addons.UndoCloseTab = {
		Exec: async function (el) {
			debugger;
			Sync.UndoCloseTab.Exec(await GetFolderViewEx(el));
			return S_OK;
		},

		KillTimer: function () {
			if (Addons.UndoCloseTab.tid) {
				clearTimeout(Addons.UndoCloseTab.tid);
				delete Addons.UndoCloseTab.tid;
			}
		},

		Save: function () {
			Addons.UndoCloseTab.KillTimer();
			Common.UndoCloseTab.bSave = true;
			Addons.UndoCloseTab.tid = setTimeout(function () {
				Sync.UndoCloseTab.SaveEx()
			}, 999);
		}
	}

	AddTypeEx("Add-ons", "Undo close tab", Addons.UndoCloseTab.Exec);

	var h = GetIconSize(item.getAttribute("IconSize"), item.getAttribute("Location") == "Inner" && 16);
	var s = item.getAttribute("Icon");
	SetAddon(Addon_Id, Default, ['<span class="button" onclick="Addons.UndoCloseTab.Exec(this)" oncontextmenu="return false;" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({ title: item.getAttribute("MenuName") || await GetAddonInfo(Addon_Id).Name, src: s }, h), '</span>']);

	Common.UndoCloseTab = await api.CreateObject("Object");
	importJScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	EnableInner();
	SetTabContents(0, "General", '<label>Number of items</label><br><input type="text" name="Save" size="4">');
}
