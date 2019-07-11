var Addon_Id = "filterbar";
var Default = "ToolBar2Right";

var item = GetAddonElement(Addon_Id);
if (!item.getAttribute("Set")) {
	item.setAttribute("MenuPos", -1);

	item.setAttribute("KeyExec", 1);
	item.setAttribute("KeyOn", "All");
	item.setAttribute("Key", "Ctrl+E");
}

if (window.Addon == 1) {
	Addons.FilterBar =
	{
		tid: null,
		filter: null,
		iCaret: -1,
		strName: "Filter Bar",

		KeyDown: function (o)
		{
			var k = window.event.keyCode;
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

		KeyUp: function (o)
		{
			var k = window.event.keyCode;
			if (k == VK_UP || k == VK_DOWN) {
				var FV = te.Ctrl(CTRL_FV);
				if (FV) {
					FV.Focus();
					return false;
				}
			}
		},

		Change: function ()
		{
			Addons.FilterBar.ShowButton();
			var FV = te.Ctrl(CTRL_FV);
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
			if (String(s).toLowerCase() != FV.FilterView.toLowerCase()) {
				FV.FilterView = s || null;
				FV.Refresh();
			}
		},

		Focus: function (o)
		{
			o.select();
			if (this.iCaret >= 0) {
				var range = o.createTextRange();
				range.move("character", this.iCaret);
				range.select();
				this.iCaret = -1;
			}
			this.ShowButton();
		},

		Clear: function (flag)
		{
			document.F.filter.value = "";
			this.ShowButton();
			if (flag) {
				var FV = te.Ctrl(CTRL_FV);
				FV.FilterView = null;
				FV.Refresh();
			}
		},

		ShowButton: function ()
		{
			if (WINVER < 0x602) {
				document.getElementById("ButtonFilterClear").style.display = document.F.filter.value.length ? "inline" : "none";
			}
		},

		Exec: function ()
		{
			document.F.filter.focus();
			return S_OK;
		},

		GetFilter: function (Ctrl)
		{
			if (Ctrl.Type <= CTRL_EB) {
				clearTimeout(Addons.FilterBar.tid);
				var s = Addons.FilterBar.GetString(Ctrl.FilterView);
				if (s != Addons.FilterBar.GetString(document.F.filter.value)) {
					document.F.filter.value = s;
					Addons.FilterBar.ShowButton();
				}
			}
		},

		GetString: function (s)
		{
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

		FilterList: function (o)
		{
			if (Addons.FilterList) {
				Addons.FilterList.Exec(o);
			}
			return false;
		}

	};

	AddEvent("ChangeView", Addons.FilterBar.GetFilter);
	AddEvent("Command", Addons.FilterBar.GetFilter);

	var width = "176px";

	var s = item.getAttribute("Width");
	if (s) {
		width = (api.QuadPart(s) == s) ? (s + "px") : s;
	}
	var icon = ExtractMacro(te, api.PathUnquoteSpaces(item.getAttribute("Icon"))) || "../addons/filterbar/filter.png";
	Addons.FilterBar.RE = item.getAttribute("RE");
	//Menu
	if (item.getAttribute("MenuExec")) {
		Addons.FilterBar.nPos = api.LowPart(item.getAttribute("MenuPos"));
		var s = item.getAttribute("MenuName");
		if (s) {
			Addons.FilterBar.strName = s;
		}
		AddEvent(item.getAttribute("Menu"), function (Ctrl, hMenu, nPos)
		{
			api.InsertMenu(hMenu, Addons.FilterBar.nPos, MF_BYPOSITION | MF_STRING, ++nPos, GetText(Addons.FilterBar.strName));
			ExtraMenuCommand[nPos] = Addons.FilterBar.Exec;
			return nPos;
		});
	}
	//Key
	if (item.getAttribute("KeyExec")) {
		SetKeyExec(item.getAttribute("KeyOn"), item.getAttribute("Key"), Addons.FilterBar.Exec, "Func");
	}
	//Mouse
	if (item.getAttribute("MouseExec")) {
		SetGestureExec(item.getAttribute("MouseOn"), item.getAttribute("Mouse"), Addons.FilterBar.Exec, "Func");
	}
	AddTypeEx("Add-ons", "Filter Bar", Addons.FilterBar.Exec);

	SetAddon(Addon_Id, Default, ['<input type="text" name="filter" placeholder="Filter" onkeydown="return Addons.FilterBar.KeyDown(this)" onkeyup="return Addons.FilterBar.KeyUp(this)" onfocus="Addons.FilterBar.Focus(this)" onblur="Addons.FilterBar.ShowButton()" onmouseup="Addons.FilterBar.KeyDown(this)" ondblclick="return Addons.FilterBar.FilterList(this)" style="width:', EncodeSC(width), '; padding-right: 12pt; vertical-align: middle"><span style="position: relative"><input type="image" src="', EncodeSC(icon), '" id="ButtonFilter" hidefocus="true" style="position: absolute; left: -13.5pt; top: 1pt; width: 12pt; height: 12pt" onclick="return Addons.FilterBar.FilterList(this)" oncontextmenu="return Addons.FilterBar.FilterList(this)"><span id="ButtonFilterClear"  class="button" style="font-family: marlett; font-size: 7pt; display: none; position: absolute; left: -21pt; top: 3pt" onclick="Addons.FilterBar.Clear(true)">r</span></span>'], "middle");
} else {
	SetTabContents(0, "General", '<table style="width: 100%"><tr><td><label>Width</label></td></tr><tr><td><input type="text" name="Width" size="10"></td><td><input type="button" value="Default" onclick="document.F.Width.value=\'\'"></td></tr><tr><td><label>Filter</label></td></tr><tr><td><input type="checkbox" id="RE" name="RE"><label for="RE">Regular Expression</label>/<label for="RE">Migemo</label></td></tr></table>');
}
