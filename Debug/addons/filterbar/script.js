var Addon_Id = "filterbar";
var Default = "ToolBar2Right";

var item = await GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("MenuPos", -1);

	item.setAttribute("KeyExec", 1);
	item.setAttribute("KeyOn", "All");
	item.setAttribute("Key", "Ctrl+E");
}

if (window.Addon == 1) {
	Addons.FilterBar = {
		tid: null,
		filter: null,
		iCaret: -1,

		KeyDown: function (ev, o) {
			var k = ev.keyCode;
			if (k != VK_PROCESSKEY) {
				this.filter = o.value;
				clearTimeout(this.tid);
				if (k == VK_RETURN) {
					this.Change();
					return false;
				} else {
					this.tid = setTimeout(this.Change, 500);
				}
			}
		},

		KeyUp: function (ev) {
			var k = ev.keyCode;
			if (k == VK_UP || k == VK_DOWN) {
				(async function () {
					var FV = await GetFolderView();
					if (FV) {
						FV.Focus();
					}
				})();
				return false;
			}
		},

		Change: async function () {
			Addons.FilterBar.ShowButton();
			var FV = await GetFolderView();
			s = document.F.filter.value;
			if (s) {
				if (Addons.FilterBar.RE && !/^\*|\//.test(s)) {
					s = "/" + s + "/i";
				} else {
					if (!/^\//.test(s)) {
						var ar = s.split(/;/);
						for (var i in ar) {
							var res = /^([^\*\?]+)$/.exec(ar[i]);
							if (res) {
								ar[i] = "*" + res[1] + "*";
							}
						}
						s = ar.join(";");
					}
				}
			}
			if (!SameText(s, await FV.FilterView)) {
				FV.FilterView = s || null;
				FV.Refresh();
			}
		},

		Focus: function (o) {
			o.select();
			if (this.iCaret >= 0) {
				var range = o.createTextRange();
				range.move("character", this.iCaret);
				range.select();
				this.iCaret = -1;
			}
			this.ShowButton();
		},

		Clear: async function (flag) {
			document.F.filter.value = "";
			this.ShowButton();
			if (flag) {
				var FV = await GetFolderView();
				FV.FilterView = null;
				FV.Refresh();
			}
		},

		ShowButton: function () {
			if (WINVER < 0x602 || window.chrome) {
				document.getElementById("ButtonFilterClear").style.display = document.F.filter.value.length ? "" : "none";
			}
		},

		Exec: function () {
			WebBrowser.Focus();
			document.F.filter.focus();
			return S_OK;
		},

		GetFilter: async function (Ctrl) {
			if (await Ctrl.Type <= CTRL_EB && await Ctrl.Id == await Ctrl.Parent.Selected.Id && await Ctrl.Parent.Id == await te.Ctrl(CTRL_TC).Id) {
				clearTimeout(Addons.FilterBar.tid);
				var s = Addons.FilterBar.GetString(await Ctrl.FilterView);
				if (s != Addons.FilterBar.GetString(document.F.filter.value)) {
					document.F.filter.value = s;
					Addons.FilterBar.ShowButton();
				}
			}
		},

		GetString: function (s) {
			if (Addons.FilterBar.RE) {
				var res = /^\/(.*)\/i/.exec(s);
				if (res) {
					s = res[1];
				}
			} else if (s && !/^\//.test(s)) {
				var ar = s.split(/;/);
				for (var i in ar) {
					var res = /^\*([^/?/*]+)\*$/.exec(ar[i]);
					if (res) {
						ar[i] = res[1];
					}
				}
				s = ar.join(";");
			}
			return s;
		},

		FilterList: function (ev, o, chk) {
			if (chk) {
				var x = ev.screenX * ui_.Zoom;
				var y = ev.screenY * ui_.Zoom;
				var p = GetPos(o, 1);
				if (x < p.x || x >= p.x + o.offsetWidth || y < p.y || y >= p.y + o.offsetHeight) {
					return;
				}
			}
			if (Addons.FilterList) {
				Addons.FilterList.Exec(o);
			}
			return false;
		}

	};

	AddEvent("ChangeView1", Addons.FilterBar.GetFilter);
	AddEvent("Command", Addons.FilterBar.GetFilter);

	var width = "176px";

	var s = item.getAttribute("Width");
	if (s) {
		width = GetNum(s) == s ? (s + "px") : s;
	}
	var icon = item.getAttribute("Icon") ? EncodeSC(await ExtractMacro(te, await api.PathUnquoteSpaces(item.getAttribute("Icon")))) : await MakeImgSrc("bitmap:comctl32.dll,140,13,0", 0, false, 13);
	Addons.FilterBar.RE = item.getAttribute("RE");
	//Menu
	if (item.getAttribute("MenuExec")) {
		Common.FilterBar = await api.CreateObject("Object");
		Common.FilterBar.strMenu = item.getAttribute("Menu");
		Common.FilterBar.strName = item.getAttribute("MenuName") || await GetAddonInfo(Addon_Id).Name;
		Common.FilterBar.nPos = GetNum(item.getAttribute("MenuPos"));
		$.importScript("addons\\" + Addon_Id + "\\sync.js");
	}
	//Key
	if (item.getAttribute("KeyExec")) {
		SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Addons.FilterBar.Exec, "Async");
	}
	//Mouse
	if (item.getAttribute("MouseExec")) {
		SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Addons.FilterBar.Exec, "Async");
	}
	AddTypeEx("Add-ons", "Filter Bar", Addons.FilterBar.Exec);

	var z = screen.deviceYDPI / 96;
	var nSize = item.getAttribute("Icon") ? 16 : 13;
	SetAddon(Addon_Id, Default, ['<input type="text" name="filter" placeholder="Filter" onkeydown="return Addons.FilterBar.KeyDown(event, this)" onkeyup="return Addons.FilterBar.KeyUp(event)" onfocus="Addons.FilterBar.Focus(this)" onblur="Addons.FilterBar.ShowButton()" onmouseup="Addons.FilterBar.KeyDown(event, this)" ondblclick="return Addons.FilterBar.FilterList(event, this)" style="width:', EncodeSC(width), '; padding-right:', nSize * z, 'px; vertical-align: middle"><span style="position: relative"><input type="image" src="', icon, '" id="ButtonFilter" hidefocus="true" style="position: absolute; left:', -(nSize + 2) * z, 'px; top:', (18 - nSize) / 2 * z, 'px; width:', nSize * z, 'px; height:', nSize * z, 'px" onclick="return Addons.FilterBar.FilterList(event, this,1)" oncontextmenu="return Addons.FilterBar.FilterList(event, this)"><span id="ButtonFilterClear" class="button" style="font-family: marlett; font-size:', 9 * z, 'px; display: none; position: absolute; left:', -(nSize + 12) * z, 'px; top:', 4 * z, 'px" onclick="Addons.FilterBar.Clear(true)">r</span></span>'], "middle");
} else {
	SetTabContents(0, "General", '<table style="width: 100%"><tr><td><label>Width</label></td></tr><tr><td><input type="text" name="Width" size="10"></td><td><input type="button" value="Default" onclick="document.F.Width.value=\'\'"></td></tr><tr><td><label>Filter</label></td></tr><tr><td><input type="checkbox" id="RE" name="RE"><label for="RE">Regular Expression</label>/<label for="RE">Migemo</label></td></tr></table>');
	ChangeForm([["__IconSize", "style/display", "none"]]);
}
