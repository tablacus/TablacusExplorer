const Addon_Id = "undoclosetab";
const Default = "None";
const item = await GetAddonElement(Addon_Id);
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
		Popup: async function (Ctrl, pt) {
			if (Addons.RecentlyClosedTabs) {
				Addons.RecentlyClosedTabs.Exec(Ctrl, pt);
			}
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

	AddEvent("Layout", async function () {
		SetAddon(Addon_Id, Default, ['<span class="button" onclick="SyncExec(Sync.UndoCloseTab.Exec, this)" oncontextmenu="SyncExec(Addons.UndoCloseTab.Popup, this, 9); return false" onmouseover="MouseOver(this)" onmouseout="MouseOut()">', await GetImgTag({
			title: item.getAttribute("MenuName") || await GetAddonInfo(Addon_Id).Name,
			src: item.getAttribute("Icon") || "icon:browser,16"
		}, GetIconSizeEx(item)), '</span>']);
	});

	$.importScript("addons\\" + Addon_Id + "\\sync.js");
} else {
	EnableInner();
	SetTabContents(0, "General", '<label>Number of items</label><br><input type="text" name="Save" size="4">');
}
