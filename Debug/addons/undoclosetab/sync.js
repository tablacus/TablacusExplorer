const Addon_Id = "undoclosetab";
const item = GetAddonElement(Addon_Id);

Common.UndoCloseTab = api.CreateObject("Object");
Sync.UndoCloseTab = {
	Items: item.getAttribute("Save") || 30,
	strName: item.getAttribute("MenuName") || GetAddonInfo(Addon_Id).Name,
	nPos: GetNum(item.getAttribute("MenuPos")),
	CONFIG: fso.BuildPath(te.Data.DataFolder, "config\\closedtabs.xml"),

	Exec: function (Ctrl, pt) {
		const FV = GetFolderView(Ctrl, pt);
		if (FV) {
			Sync.UndoCloseTab.bLock = true;
			while (Common.UndoCloseTab.db.length) {
				Sync.UndoCloseTab.bFail = false;
				Sync.UndoCloseTab.Open(FV, 0);
				if (!Sync.UndoCloseTab.bFail) {
					break;
				}
			}
			Sync.UndoCloseTab.bLock = false;
		}
		return S_OK;
	},

	Open: function (FV, i) {
		if (FV) {
			const Items = Sync.UndoCloseTab.Get(i);
			Common.UndoCloseTab.db.splice(i, 1);
			FV.Navigate(Items, SBSP_NEWBROWSER);
			InvokeUI("Addons.UndoCloseTab.Save");
		}
	},

	Get: function (nIndex) {
		Common.UndoCloseTab.db.splice(Sync.UndoCloseTab.Items, MAXINT);
		let s = Common.UndoCloseTab.db[nIndex];
		if ("string" === typeof s) {
			const a = s.split(/\n/);
			s = api.CreateObject("FolderItems");
			s.Index = a.pop();
			for (let i in a) {
				s.AddItem(a[i]);
			}
			Common.UndoCloseTab.db[nIndex] = s;
		}
		return s;
	},

	Load: function () {
		Common.UndoCloseTab.db = api.CreateObject("Array");
		const xml = OpenXml("closedtabs.xml", true, false);
		if (xml) {
			const items = xml.getElementsByTagName('Item');
			for (let i = items.length; i--;) {
				Common.UndoCloseTab.db.unshift(items[i].text);
			}
		}
		Sync.UndoCloseTab.ModifyDate = api.ILCreateFromPath(Sync.UndoCloseTab.CONFIG).ModifyDate;
	},

	SaveEx: function () {
		if (Common.UndoCloseTab.bSave) {
			Common.UndoCloseTab.bSave = false;
			InvokeUI("Addons.UndoCloseTab.KillTimer");
			const xml = CreateXml();
			const root = xml.createElement("TablacusExplorer");

			const db = Common.UndoCloseTab.db;
			for (let i = 0; i < db.length; i++) {
				const item = xml.createElement("Item");
				let s = db[i];
				if (s) {
					if ("string" !== typeof s) {
						const a = [];
						for (let j in s) {
							a.push(api.GetDisplayNameOf(s[j], SHGDN_FORPARSING | SHGDN_FORPARSINGEX));
						}
						a.push(s.Index);
						s = a.join("\n");
					}
					item.text = s;
					root.appendChild(item);
				}
			}
			xml.appendChild(root);
			SaveXmlEx("closedtabs.xml", xml, true);
			Sync.UndoCloseTab.ModifyDate = api.ILCreateFromPath(Sync.UndoCloseTab.CONFIG).ModifyDate;
		}
	}
}
Sync.UndoCloseTab.Load();

AddEvent("CloseView", function (Ctrl) {
	if (Ctrl.FolderItem) {
		if (Sync.UndoCloseTab.bLock) {
			Sync.UndoCloseTab.bFail = true;
		} else if (Ctrl.History.Count) {
			Common.UndoCloseTab.db.unshift(Ctrl.History);
			Common.UndoCloseTab.db.splice(Sync.UndoCloseTab.Items, MAXINT);
			InvokeUI("Addons.UndoCloseTab.Save");
		}
	}
	return S_OK;
});

AddEvent("SaveConfig", Sync.UndoCloseTab.SaveEx);

AddEvent("ChangeNotifyItem:" + Sync.UndoCloseTab.CONFIG, function (pid) {
	if (pid.ModifyDate - Sync.UndoCloseTab.ModifyDate) {
		Sync.UndoCloseTab.Load();
	}
});

//Menu
if (item.getAttribute("MenuExec")) {
	AddEvent(item.getAttribute("Menu"), function (Ctrl, hMenu, nPos) {
		api.InsertMenu(hMenu, Sync.UndoCloseTab.nPos, MF_BYPOSITION | MF_STRING | ((Common.UndoCloseTab.db.length) ? MF_ENABLED : MF_DISABLED), ++nPos, GetText(Sync.UndoCloseTab.strName));
		ExtraMenuCommand[nPos] = Sync.UndoCloseTab.Exec;
		return nPos;
	});
}
//Key
if (item.getAttribute("KeyExec")) {
	SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Sync.UndoCloseTab.Exec, "Func");
}
//Mouse
if (item.getAttribute("MouseExec")) {
	SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Sync.UndoCloseTab.Exec, "Func");
}

AddTypeEx("Add-ons", "Undo close tab", Sync.UndoCloseTab.Exec);
