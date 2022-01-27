const Addon_Id = "aboutblank";

Sync.AboutBlank = {
	dir: [ssfDRIVES, "shell:downloads", ssfPERSONAL, "shell:my music", "shell:my pictures", "shell:my video"],

	IsHandle: function (Ctrl) {
		return SameText("string" === typeof Ctrl ? Ctrl : api.GetDisplayNameOf(Ctrl, SHGDN_FORADDRESSBAR | SHGDN_FORPARSING), "about:blank");
	},

	ProcessMenu: function (Ctrl, hMenu, nPos, Selected, item, ContextMenu) {
		const FV = GetFolderView(Ctrl);
		if (Sync.AboutBlank.IsHandle(FV)) {
			RemoveCommand(hMenu, ContextMenu, "delete;rename");
		}
		return nPos;
	}
}

AddEvent("TranslatePath", function (Ctrl, Path) {
	if (Sync.AboutBlank.IsHandle(Path)) {
		Ctrl.Enum = function (pid, Ctrl, fncb) {
			const Items = api.CreateObject("FolderItems");
			for (let i = 0; i < Sync.AboutBlank.dir.length; ++i) {
				Items.AddItem(Sync.AboutBlank.dir[i]);
			}
			for (const e = api.CreateObject("Enum", fso.Drives); !e.atEnd(); e.moveNext()) {
				Items.AddItem(e.item().Path);
			}
			return Items;
		};
		return ssfRESULTSFOLDER;
	}
}, true);

AddEvent("GetTabName", function (Ctrl) {
	if (Sync.AboutBlank.IsHandle(Ctrl)) {
		return GetText("New tab");
	}
}, true);

AddEvent("GetIconImage", function (Ctrl, clBk, bSimple) {
	if (Sync.AboutBlank.IsHandle(Ctrl)) {
		return MakeImgDataEx("folder:closed", bSimple, 16, clBk);
	}
});

AddEvent("Context", Sync.AboutBlank.ProcessMenu);

AddEvent("File", Sync.AboutBlank.ProcessMenu);

AddEvent("Command", function (Ctrl, hwnd, msg, wParam, lParam) {
	if (Ctrl.Type == CTRL_SB || Ctrl.Type == CTRL_EB) {
		if (Sync.AboutBlank.IsHandle(Ctrl)) {
			if ((wParam & 0xfff) == CommandID_DELETE - 1) {
				return S_OK;
			}
		}
	}
}, true);

AddEvent("InvokeCommand", function (ContextMenu, fMask, hwnd, Verb, Parameters, Directory, nShow, dwHotKey, hIcon) {
	if (Verb == CommandID_DELETE - 1) {
		const FV = ContextMenu.FolderView;
		if (FV && Sync.AboutBlank.IsHandle(FV)) {
			return S_OK;
		}
	}
}, true);

AddEvent("BeginLabelEdit", function (Ctrl, Name) {
	if (Ctrl.Type <= CTRL_EB) {
		if (Sync.AboutBlank.IsHandle(Ctrl)) {
			return 1;
		}
	}
});

