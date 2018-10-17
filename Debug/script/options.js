//Tablacus Explorer

nTabMax = 0;
TabIndex = -1;
g_x = {Menu: null, Addons: null};
g_Chg = {Menus: false, Addons: false, Tab: false, Tree: false, View: false, Data: null};
g_arMenuTypes = ["Default", "Context", "Background", "Tabs", "Tree", "File", "Edit", "View", "Favorites", "Tools", "Help", "Systray", "System", "Alias"];
g_MenuType = "";
g_dlgAddons = null;
g_bDrag = false;
g_pt = {x: 0, y: 0};
g_Gesture = null;
g_tidResize = null;
g_drag5 = false;
g_nResult = 0;
g_bChanged = true;
g_bClosed = false;
g_nSort2 = 1;

arLangs = [GetLangId()];
var res = /(\w+)_/.test(arLangs[0]);
if (res) {
	arLangs.push(res[1]);
}
if (!/^en/.test(arLangs[0])) {
	arLangs.push("en");
}
arLangs.push("General");
g_ovPanel = null;

urlAddons = "https://tablacus.github.io/TablacusExplorerAddons/";
xhr = null;
xmlAddons = null;

function SetDefaultLangID()
{
	SetDefault(document.F.Conf_Lang, GetLangId(true));
}

function SetDefault(o, v)
{
	setTimeout(function ()
	{
		if (confirmOk()) {
			SetValue(o, v);
		}
	}, 99);
}

function OpenGroup(id)
{
	var o = document.getElementById(id);
	o.style.display = /block/i.test(o.style.display) ? "none" : "block";
}

function LoadChecked(form)
{
	for(var i = 0; i < form.length; i++) {
		var o = form[i];
		var ar = o.id.split("=");
		if (ar.length > 1 && form[ar[0]].value == (api.sscanf(ar[1], "0x%x") || ar[1])) {
			form[i].checked = true;
		}
		ar = o.id.split(":");
		if (ar.length > 1 && form[ar[0]].value & (api.sscanf(ar[1], "0x%x") || ar[1])) {
			form[i].checked = true;
		}
	}
}

function ResetForm()
{
	var TV = te.Ctrl(CTRL_TV);
	if (TV) {
		document.F.Tree_Align.value = TV.Align;
		document.F.Tree_Style.value = TV.Style;
		document.F.Tree_EnumFlags.value = TV.EnumFlags;
		document.F.Tree_RootStyle.value = TV.RootStyle;
		if (TV.Align & 2) {
			document.F.Tree_Root.value = TV.Root;
		}
		document.F.Tree_Width.value = TV.Width;
	}
	var TC = te.Ctrl(CTRL_TC);
	if (TC) {
		document.F.Tab_Style.value = TC.Style;
		document.F.Tab_Align.value = TC.Align;
		document.F.Tab_TabWidth.value = TC.TabWidth;
		document.F.Tab_TabHeight.value = TC.TabHeight;

		document.F.Tab_Left.value = TC.Left;
		document.F.Tab_Top.value = TC.Top;
		document.F.Tab_Width.value = TC.Width;
		document.F.Tab_Height.value = TC.Height;
	}
	var FV = te.Ctrl(CTRL_FV);
	if (FV) {
		document.F.View_Type.value = FV.Type;
		document.F.View_ViewMode.value = FV.CurrentViewMode;
		document.F.View_fFlags.value = FV.FolderFlags;
		document.F.View_Options.value = FV.Options;
		document.F.View_ViewFlags.value = FV.ViewFlags;
	}
	document.F.Conf_SizeFormat.value = te.Data.Conf_SizeFormat || 0;

	for(i = 0; i < document.F.length; i++) {
		o = document.F[i];
		if (String(o.type).toLowerCase() == 'checkbox') {
			if (!/^Conf_/.test(o.id)) {
				o.checked = false;
			}
		}
	}
	LoadChecked(document.F);
	var s = GetWebColor(document.F.Conf_TrailColor.value);
	document.F.Conf_TrailColor.value = s;
	document.F.Color_Conf_TrailColor.style.backgroundColor = s;

	var o = document.getElementById("_DateTimeFormat");
	if (/\@/.test(o.innerHTML)) {
		o.innerHTML = api.PSGetDisplayName("System.ItemDate");
	}
	document.getElementById("_EDIT").checked = true;
}

function ResizeTabPanel()
{
	var h = 4.8;
	if (g_.Inline && !/\w/.test(document.getElementById('toolbar').innerHTML)) {
		h = 2.4;
	}
	CalcElementHeight(g_ovPanel, h);
}

function ClickTab(o, nMode)
{
	nMode = api.LowPart(nMode);
	if (o && o.id) {
		nTabIndex = o.id.replace(new RegExp('tab', 'g'), '') - 0;
	}
	var i = 0;
	var ovTab;
	while (ovTab = document.getElementById('tab' + i)) {
		var ovPanel = document.getElementById('panel' + i);
		if (i == nTabIndex) {
			try {
				ovTab.focus();
			} catch (e) {}
			ovTab.className = 'activetab';
			ovPanel.style.display = 'block';
			g_ovPanel = ovPanel;
			ResizeTabPanel();
			if (window.OnTabChanged) {
				OnTabChanged(i);
			}
		} else {
			ovTab.className = 'tab';
			ovPanel.style.display = 'none';
		}
		i++;
	}
	nTabMax = i;
}

function ClickTree(o, nMode, strChg, bForce)
{
	if (strChg) {
		g_Chg[strChg] = true;
	}
	nMode = api.LowPart(nMode);
	var newTab = TabIndex != -1 ? TabIndex : 0;
	if (o && o.id) {
		var res = /tab([^_]+)(_?)(.*)/.exec(o.id);
		if (res) {
			newTab = res[1] + res[2] + res[3];
			if (res[3]) {
				ClickButton(res[1], true);
				if (res[1] == 1) {
					(function (Id) { setTimeout(function () {
						if (g_.elAddons[Id]) {
							if (!g_.elAddons[Id].contentWindow || !g_.elAddons[Id].contentWindow.document.body.innerHTML) {
								AddonOptions(Id, null, null, true);
							}
						}
					}, 999);}) (res[3]);
				}
			}
			ShowButtons(/^1$|^1_1$/.test(newTab), res[1] == 2, newTab);
			if (nMode == 0) {
				switch (res[1] - 0) {
					case 1:
						setTimeout(LoadAddons, 9);
						if (newTab == '1_1') {
							setTimeout(GetAddons, 9);
						}
						break;
					case 2:
						LoadMenus(res[3] - 0);
						break;
				}
			}
		}
	}
	if (newTab != TabIndex || bForce) {
		if (newTab == "0") {
			var o = document.getElementById("DefaultLangID");
			if (o && o.innerHTML == "") {
				var s = GetLangId(1);
				var s2 = GetLangId(2);
				if (s != s2) {
					s += ' (' + s2 + ')';
				}
				o.innerHTML = s;
			}
		}
		var ovTab = document.getElementById('tab' + TabIndex);
		if (ovTab) {
			var ovPanel = document.getElementById('panel' + TabIndex) || document.getElementById('panel' + TabIndex.replace(/_\w+/, ""));
			ovTab.className = 'button';
			ovPanel.style.display = 'none';
		}
		TabIndex = newTab;
		ovTab = document.getElementById('tab' + TabIndex);
		if (ovTab) {
			ovPanel = document.getElementById('panel' + TabIndex) || document.getElementById('panel' + TabIndex.replace(/_\w+/, ""));
			var res = /2_(.+)/.exec(TabIndex);
			if (res) {
				document.F.Menus.selectedIndex = res[1];
				SwitchMenus(document.F.Menus);
			}
			ovTab.className = 'hoverbutton';
			ovPanel.style.display = 'block';
			CalcElementHeight(ovPanel, 3);
			var o = document.getElementById("tab_");
			CalcElementHeight(o, 3);
			var i = ovTab.offsetTop - o.scrollTop;
			if (i < 0 || i >= o.offsetHeight - ovTab.offsetHeight) {
				ovTab.scrollIntoView(i < 0);
			}
		}
	}
}

function ClickButton(n, f)
{
	var o = document.getElementById("tabbtn" + n);
	var op = document.getElementById("tab" + n + "_");
	if (f || op.style.display.toLowerCase() == "none") {
		o.innerHTML = BUTTONS.opened;
		op.style.display = "block";
	} else {
		o.innerHTML = BUTTONS.closed;
		op.style.display = "none";
	}
}

function SetRadio(o)
{
	var ar = o.id.split("=");
	o.form[ar[0]].value = ar[1];
	var res = /^(Tab|Tree|View|Conf)/.exec(ar[0]);
	if (res) {
		MainWindow.g_.OptionsWindow.g_Chg[res[1]] = true;
		MainWindow.g_.OptionsWindow.g_bChanged = true;
	}
}

function SetCheckbox(o)
{
	var ar = o.id.split(":");
	if (o.checked) {
		o.form[ar[0]].value |= (api.sscanf(ar[1], "0x%x") || ar[1]);
	} else {
		o.form[ar[0]].value &= ~(api.sscanf(ar[1], "0x%x") || ar[1]);
	}
	var res = /^(Tab|Tree|View|Conf)/.exec(ar[0]);
	if (res) {
		MainWindow.g_.OptionsWindow.g_Chg[res[1]] = true;
		MainWindow.g_.OptionsWindow.g_bChanged = true;
	}
}

function SetValue(n, v)
{
	if (v.value != '!') {
		n.value = v.value.replace(/\\n/, "\n");
	}
}

function AddValue(name, i, min, max)
{
	var o = document.F[name];
	i += api.LowPart(o.value);
	i = (i < min) ? min : i;
	o.value = (i < max) ? i : max;
}

function ChooseColor1(o)
{
	setTimeout(function ()
	{
		var o2 = o.form[o.id.replace("Color_", "")];
		var c = ChooseColor(o2.value);
		if (isFinite(c)) {
			o2.value = c;
			o.style.backgroundColor = GetWebColor(c);
		}
	}, 99);
}

function ChooseColor2(o)
{
	setTimeout(function ()
	{
		var o2 = o.form[o.id.replace("Color_", "")];
		var c = ChooseColor(GetWinColor(o2.value));
		if (isFinite(c)) {
			c = GetWebColor(c);
			o2.value = c;
			o.style.backgroundColor = c;
		}
	}, 99);
}

function ChooseFolder1(o)
{
	setTimeout(function ()
	{
		var r = BrowseForFolder(o.value);
		if (r) {
			o.value = r;
		}
	}, 99);
}

function AddTabControl()
{
	if (document.F.Tab_Width.value == 0) {
		MessageBox("Please enter the width.", TITLE, MB_ICONEXCLAMATION);
		return;
	}
	if (document.F.Tab_Height.value == 0) {
		MessageBox("Please enter the height.", TITLE, MB_ICONEXCLAMATION);
		return;
	}
	var TC = te.CreateCtrl(CTRL_TC, document.F.Tab_Left.value, document.F.Tab_Top.value, document.F.Tab_Width.value, document.F.Tab_Height.value, document.F.Tab_Style.value, document.F.Tab_Align.value, document.F.Tab_TabWidth.value, document.F.Tab_TabHeight.value);
	TC.Selected.Navigate2("C:\\", SBSP_NEWBROWSER, document.F.View_Type.value, document.F.View_ViewMode.value, document.F.View_fFlags.value, 0, document.F.View_Options.value, document.F.View_ViewFlags.value);
}

function DelTabControl()
{
	var TC = te.Ctrl(CTRL_TC);
	if (TC) {
		TC.Close();
	}
}

function SetTabControls()
{
	if (g_Chg.Tab) {
		if (document.getElementById("Conf_TabDefault").checked) {
			var cTC = te.Ctrls(CTRL_TC);
			for (i = 0; i < cTC.Count; i++) {
				SetTabControl(cTC.Item(i));
			}
		} else {
			var TC = te.Ctrl(CTRL_TC);
			if (TC) {
				SetTabControl(TC);
			}
		}
	}
}

function SetFolderViews()
{
	if (g_Chg.View) {
		var o = te.Object();
		o.Layout = te.Data.Conf_Layout;
		o.ViewMode = document.F.View_ViewMode.value;
		MainWindow.g_.TEData = o;
		o = te.Object();
		o.All = document.getElementById("Conf_ListDefault").checked;
		o.FolderFlags = document.F.View_fFlags.value;
		o.Options = document.F.View_Options.value;
		o.ViewFlags = document.F.View_ViewFlags.value;
		o.Type = document.F.View_Type.value;
		MainWindow.g_.FVData = o;
	}
}

function SetTreeControls()
{
	if (g_Chg.Tree) {
		var o = te.Object();
		o.All = document.getElementById("Conf_TreeDefault").checked;
		o.Align = document.F.Tree_Align.value;
		o.Style = document.F.Tree_Style.value;
		o.Width = document.F.Tree_Width.value;
		o.Root = document.F.Tree_Root.value;
		o.EnumFlags = document.F.Tree_EnumFlags.value;
		o.RootStyle = document.F.Tree_RootStyle.value;
		MainWindow.g_.TVData = o;
	}
}

function SetTabControl(TC)
{
	if (TC) {
		TC.Style = document.F.Tab_Style.value;
		TC.Align = document.F.Tab_Align.value;
		TC.TabWidth = document.F.Tab_TabWidth.value;
		TC.TabHeight = document.F.Tab_TabHeight.value;
	}
}

function GetTabControl()
{
	var TC = te.Ctrl(CTRL_TC);
	if (TC) {
		document.F.Tab_Left.value = TC.Left;
		document.F.Tab_Top.value = TC.Top;
		document.F.Tab_Width.value = TC.Width;
		document.F.Tab_Height.value = TC.Height;
	}
}

function MoveTabControl()
{
	var TC = te.Ctrl(CTRL_TC);
	if (TC) {
		if (document.F.Tab_Width.value != "" && document.F.Tab_Height.value != "") {
			TC.Left = document.F.Tab_Left.value;
			TC.Top = document.F.Tab_Top.value;
			TC.Width = document.F.Tab_Width.value;
			TC.Height = document.F.Tab_Height.value;
		}
	}
}

function InitConfig(o)
{
	var InstallPath = fso.GetParentFolderName(api.GetModuleFileName(null));
	if (InstallPath == te.Data.DataFolder) {
		return;
	}
	if (!confirmOk()) {
		return;
	}
	api.SHFileOperation(FO_MOVE, fso.BuildPath(InstallPath, "layout"), te.Data.DataFolder, 0, false);
	o.disabled = true;
}

function SelectPos(o, s)
{
	var v = o[o.selectedIndex].value;
	if (v != "") {
		o.form[s].value = v;
	}
}

function SwitchMenus(o)
{
	if (g_x.Menus) {
		g_x.Menus.style.display = "none";
		var o = o || document.F.elements.Menus;
		for (var i = o.length; i-- > 0;) {
			var a = o[i].value.split(",");
			if ("Menus_" + a[0] == g_x.Menus.name) {
				s = a[0] + "," + o.form["Menus_Base"].selectedIndex + "," + o.form["Menus_Pos"].value;
				if (s != o[i].value) {
					g_Chg.Menus = true;
					o[i].value = s;
				}
				break;
			}
		}
	}
	if (o) {
		var a = o.value.split(",");
		g_x.Menus = o.form["Menus_" + a[0]];
		g_x.Menus.style.display = "inline";
		o.form["Menus_Base"].selectedIndex = a[1];
		o.form["Menus_Pos"].value = api.LowPart(a[2]);
		CancelX("Menus");
	}
}

function SwitchX(mode, o, form)
{
	g_x[mode].style.display = "none";
	g_x[mode] = (form || o.form)[mode + o.value];
	g_x[mode].style.display = "inline";
	CancelX(mode);
}

function ClearX(mode)
{
	g_Chg.Data = null;
}

function CancelX(mode)
{
	g_x[mode].selectedIndex = -1;
	EnableSelectTag(g_x[mode]);
}

ChangeX = function (mode)
{
	g_Chg.Data = mode;
	g_bChanged = true;
}

function ConfirmX(bCancel, fn)
{
	g_bChanged |= g_Chg.Data;
	return SetOptions(function() {
		SetChanged(fn);
	},
	function() {
		g_bChanged = false;
		ClearX(g_Chg.Data);
	}, !bCancel, true);
}

function SetOptions(fnYes, fnNo)
{
	if (g_nResult == 2 || document.activeElement && api.StrCmpI(document.activeElement.value, GetText("Cancel")) == 0) {
		if (fnNo) {
			fnNo();
		}
		return false;
	}
	if (document.activeElement) {
		document.activeElement.blur();
	}
	if (g_nResult == 1 && !g_Chg.Data) {
		if (fnYes) {
			fnYes();
		}
		return true;
	}
 	if (g_bChanged || g_Chg.Data) {
		switch (MessageBox("Do you want to replace?", TITLE, MB_ICONQUESTION | MB_YESNO)) {
			case IDYES:
				if (fnYes) {
					fnYes();
				}
				return true;
			case IDNO:
				if (fnNo) {
					fnNo();
				}
				if (g_nResult == 1) {
					ClearX();
					if (fnYes) {
						fnYes();
					}
					return true;
				}
				return false;
			default:
				return false;
		}
	}
	if (fnNo) {
		fnNo();
	}
	return false;
}

function EditMenus()
{
	if (g_x.Menus.selectedIndex < 0) {
		return;
	}
	ClearX("Menus");
	var a = g_x.Menus[g_x.Menus.selectedIndex].value.split(g_sep);
	var a2 = a[0].split(/\\t/);
	if (!a[5]) {
		a2[0] = GetText(a2[0]);
	}
	document.F.Menus_Key.value = a2.length > 1 ? GetKeyName(a2.pop()) : "";
	document.F.Menus_Name.value = a2.join("\\t");
	document.F.Menus_Filter.value = a[1];
	var p = { s: a[2] };
	MainWindow.OptionDecode(a[3], p);
	document.F.Menus_Path.value = p.s;
	SetType(document.F.Menus_Type, a[3]);
	document.F.Icon.value = a[4] || "";
	SetImage();
}

function EditXEx(o, s, form)
{
	if (document.getElementById("_EDIT").checked) {
		o(s, form);
	}
}

EditX = function (mode, form)
{
	if (g_x[mode].selectedIndex < 0) {
		return;
	}
	if (!form) {
		form = document.F;
	}
	ClearX(mode);
	var a = g_x[mode][g_x[mode].selectedIndex].value.split(g_sep);
	form[mode + mode].value = a[0];
	var p = { s: a[1] };
	MainWindow.OptionDecode(a[2], p);
	form[mode + "Path"].value = p.s;
	SetType(form[mode + "Type"], a[2]);
	if (api.StrCmpI(mode, "key") == 0) {
	 	SetKeyShift();
	}
	var o = form[mode + "Name"];
	if (o) {
		o.value = GetText(a[3] || "");
	}
}

function SetType(o, value)
{
	value = value.replace(/&|\.\.\.$/g, "");
	var i = o.length;
	while (--i >= 0) {
		if (o[i].value.toLowerCase() == value.toLowerCase()) {
			o.selectedIndex = i;
			break;
		}
	}
	if (i < 0) {
		i = o.length++;
		o[i].value = value;
		o[i].innerText = value;
		o.selectedIndex = i;
	}
}

function InsertX(sel)
{
	if (!sel) {
		return;
	}
	sel.length++;
	if (sel.selectedIndex < 0) {
		sel.selectedIndex = sel.length - 1;
	} else {
		sel.selectedIndex++;
		for (var i = sel.length; i-- > sel.selectedIndex;) {
			sel[i].text = sel[i - 1].text;
			sel[i].value = sel[i - 1].value;
		}
	}
}

function AddX(mode, fn, form)
{
	InsertX(g_x[mode]);
	(fn || ReplaceX)(mode, form);
	EnableSelectTag(g_x[mode]);
}

function ReplaceMenus()
{
	ClearX("Menus");
	if (g_x.Menus.selectedIndex < 0) {
		InsertX(g_x.Menus);
	}
	var sel = g_x.Menus[g_x.Menus.selectedIndex];
	var o = document.F.Menus_Type;
	var s = GetSourceText(document.F.Menus_Name.value.replace(/\\/g, "/"));
	var org = (s.toLowerCase() == document.F.Menus_Name.value.toLowerCase() && api.GetKeyState(VK_SHIFT) >= 0) ? "1" : ""
	if (document.F.Menus_Key.value.length) {
		var n = GetKeyKey(document.F.Menus_Key.value);
		s += "\\t" + (n ? api.sprintf(8, "$%x", n) : document.F.Menus_Key.value);
	}
	var p = { s: document.F.Menus_Path.value };
	MainWindow.OptionEncode(o[o.selectedIndex].value, p);
	SetMenus(sel, [s, document.F.Menus_Filter.value, p.s, o[o.selectedIndex].value, document.F.Icon.value, org]);
	g_Chg.Menus = true;
}

function ReplaceX(mode, form)
{
	if (!g_x[mode]) {
		return;
	}
	if (!form) {
		form = document.F;
	}
	ClearX(mode);
	if (g_x[mode].selectedIndex < 0) {
		InsertX(g_x[mode]);
		EnableSelectTag(g_x[mode]);
	}
	var sel = g_x[mode][g_x[mode].selectedIndex];
	var o = form[mode + "Type"];
	var p = { s: form[mode + "Path"].value };
	MainWindow.OptionEncode(o[o.selectedIndex].value, p);
	var o2 = form[mode + "Name"];
	SetData(sel, [form[mode + mode].value, p.s, o[o.selectedIndex].value, o2 ? GetSourceText(o2.value) : ""]);
	g_Chg[mode] = true;
	g_bChanged = true;
}

function RemoveX(mode)
{
	ClearX(mode);
	if (g_x[mode].selectedIndex < 0 || !confirmOk()) {
		return;
	}
	var i = g_x[mode].selectedIndex;
	g_x[mode][i] = null;
	g_Chg[mode] = true;
	g_bChanged = true;
	if (i >= g_x[mode].length) {
		i = g_x[mode].length - 1;
	}
	if (i >= 0) {
		g_x[mode].selectedIndex = i;
		FireEvent(g_x[mode][i], "change");
	}
}

function MoveX(mode, n)
{
	if (g_x[mode].selectedIndex < 0 || g_x[mode].selectedIndex + n < 0 || g_x[mode].selectedIndex + n >= g_x[mode].length) {
		return;
	}
	var src = g_x[mode][g_x[mode].selectedIndex];
	var dist = g_x[mode][g_x[mode].selectedIndex + n];
	var text = dist.text;
	var value = dist.value;
	dist.text = src.text;
	dist.value = src.value;
	src.text = text;
	src.value = value;
	g_x[mode].selectedIndex += n;
	g_Chg[mode] = true;
	g_bChanged = true;
	EnableSelectTag(g_x[mode]);
}

function SetMenus(sel, a)
{
	sel.value = PackData(a);
	var a2 = a[0].split(/\\t/);
	sel.text = [a[5] ? a2[0] : GetText(a2[0]), a[1]].join(" ").replace(/[\r\n].*/, "");
	if (!sel.text.length) {
		sel.text = "********";
	}
}

function LoadMenus(nSelected)
{
	var oa = document.F.Menus_Type;
	if (!g_x.Menus) {
		var arFunc = [];
		MainWindow.RunEvent1("AddType", arFunc);
		for (var i = 0; i < arFunc.length; i++) {
			var o = oa[++oa.length - 1];
			o.value = arFunc[i];
			o.innerText = GetText(arFunc[i]);
		}

		oa = document.F.Menus;
		oa.length = 0;

		for (var j in g_arMenuTypes) {
			var s = g_arMenuTypes[j];
			document.getElementById("Menus_List").insertAdjacentHTML("BeforeEnd", ['<select name="Menus_', s, '" size="17" style="width: 12em; height: 34em; height: calc(100vh - 6em); min-height: 20em; display: none" onchange="EditXEx(EditMenus)" ondblclick="EditMenus()" oncontextmenu="CancelX(\'Menus\')"></select>'].join(""));
			var menus = teMenuGetElementsByTagName(s);
			if (menus && menus.length) {
				oa[++oa.length - 1].value = s + "," + menus[0].getAttribute("Base") + "," + menus[0].getAttribute("Pos");
				var o = document.F["Menus_" + s];
				var items = menus[0].getElementsByTagName("Item");
				if (items) {
					var i = items.length;
					o.length = i;
					while (--i >= 0) {
						var item = items[i];
						SetMenus(o[i], [item.getAttribute("Name"), item.getAttribute("Filter"), item.text, item.getAttribute("Type"), item.getAttribute("Icon"), item.getAttribute("Org")]);
					}
				}
			} else {
				oa[++oa.length - 1].value = s;
			}
			oa[oa.length - 1].text = GetText(s);
		}
		SwitchMenus(oa[nSelected]);
	}
	for (var j in g_arMenuTypes) {
		var ar = String(g_MenuType).split(",");
		if (api.StrCmpI(ar[0], g_arMenuTypes[j]) == 0) {
			nSelected = oa.length - 1;
			oa[nSelected].selected = true;
			g_MenuType = j;
			(function (o, v) { setTimeout(function () {
				ClickTree(document.getElementById("tab2_" + g_MenuType));
				EditMenus();
				g_MenuType = "";
				if (isFinite(v)) {
					o.selectedIndex = v;
					EnableSelectTag(o);
					FireEvent(o, "change");
				}
			}, 99);}) (document.F["Menus_" + ar[0]], ar[1]);
		}
	}
}

function LoadX(mode, fn, form)
{
	if (!g_x[mode]) {
        if (!form) {
            form = document.F;
        }
		var arFunc = [];
		MainWindow.RunEvent1("AddType", arFunc);
		var oa = form[mode + "Type"] || form.Type;
		while (oa.length) {
			oa.removeChild(oa[0]);
		}
		for (var i = 0; i < arFunc.length; i++) {
			var o = oa[++oa.length - 1];
			o.value = arFunc[i];
			o.innerText = GetText(arFunc[i]);
		}
		g_x[mode] = form[mode + "All"];
		if (g_x[mode]) {
			oa = form[mode];
			oa.length = 0;
			var xml = OpenXml(mode + ".xml", false, true);
			for (var j in g_Types[mode]) {
				var o = oa[++oa.length - 1];
				o.text = GetTextEx(g_Types[mode][j]);
				o.value = g_Types[mode][j];
				o = form[mode + g_Types[mode][j]];
				var items = xml.getElementsByTagName(g_Types[mode][j]);
				var i = items.length;
				if (i == 0 && g_Types[mode][j] == "List") {
					items = xml.getElementsByTagName("Folder");
					i = items.length;
				}
				o.length = i;
				while (--i >= 0) {
					var item = items[i];
					var s = item.getAttribute(mode);
					if (api.StrCmpI(mode, "Key") == 0) {
						var ar = /,$/.test(s) ? [s] : s.split(",");
						for (var k = ar.length; k--;) {
							ar[k] = GetKeyName(ar[k]);
						}
						s = ar.join(",");
					}
					SetData(o[i], [s, item.text, item.getAttribute("Type"), item.getAttribute("Name")]);
				}
			}
		} else {
			g_x[mode] = form.List;
			g_x[mode].length = 0;
			var path = fso.GetParentFolderName(api.GetModuleFileName(null));
			var xml = te.Data["xml" + AddonName];
			if (!xml) {
				xml = te.CreateObject("Msxml2.DOMDocument");
				xml.async = false;
				xml.load(fso.BuildPath(path, "config\\" + AddonName + ".xml"));
				te.Data["xml" + AddonName] = xml;
			}

			var items = xml.getElementsByTagName("Item");
			var i = items.length;
			g_x[mode].length = i;
			while (--i >= 0) {
				var item = items[i];
				SetData(g_x[mode][i], [item.getAttribute("Name"), item.text, item.getAttribute("Type"), item.getAttribute("Icon"), item.getAttribute("Height")]);
			}
			xml = null;
		}
		EnableSelectTag(g_x[mode]);
		fn && fn();
	}
}

function SaveMenus()
{
	if (g_Chg.Menus) {
		var xml = CreateXml();

		var root = xml.createElement("TablacusExplorer");
		for (var j in g_arMenuTypes) {
			var o = document.F["Menus_" + g_arMenuTypes[j]];
			var items = xml.createElement(g_arMenuTypes[j]);
			var a = document.F.elements.Menus[j].value.split(",");
			items.setAttribute("Base", api.LowPart(a[1]));
			items.setAttribute("Pos", api.LowPart(a[2]));
			for (var i = 0; i < o.length; i++) {
				var item = xml.createElement("Item");
				var a = o[i].value.split(g_sep);
				item.setAttribute("Name", a[0]);
				item.setAttribute("Filter", a[1]);
				item.text = a[2];
				item.setAttribute("Type", a[3]);
				item.setAttribute("Icon", a[4]);
				if (a[5]) {
					item.setAttribute("Org", 1);
				}
				items.appendChild(item);
			}
			root.appendChild(items);
		}
		xml.appendChild(root);
		te.Data.xmlMenus = xml;
		MainWindow.RunEvent1("ConfigChanged", "Menus");
	}
}

function SaveX(mode, form)
{
	if (g_Chg[mode]) {
		var xml = CreateXml();
		var root = xml.createElement("TablacusExplorer");
		for (var j in g_Types[mode]) {
			var o = (form || document.F)[mode + g_Types[mode][j]];
			for (var i = 0; i < o.length; i++) {
				var item = xml.createElement(g_Types[mode][j]);
				var a = o[i].value.split(g_sep);
				var s = a[0];
				if (mode.toLowerCase() == "key") {
					var ar = /,$/.test(s) ? [s] : s.split(",");
					for (var k = ar.length; k--;) {
						var n = GetKeyKey(ar[k]);
						if (n) {
							ar[k] = api.sprintf(8, "$%x", n);
						}
					}
					s = ar.join(",");
				} else {
					item.setAttribute("Name", a[3]);
				}
				item.setAttribute(mode, s);
				item.text = a[1];
				item.setAttribute("Type", a[2]);
				root.appendChild(item);
			}
		}
		xml.appendChild(root);
		SaveXmlEx(mode.toLowerCase() + ".xml", xml);
	}
}

function SaveAddons()
{
	try {
		if (g_Chg.Addons || te.Data.bErrorAddons) {
			te.Data.bErrorAddons = false;
			var xml = CreateXml();
			var root = xml.createElement("TablacusExplorer");
			var table = document.getElementById("Addons");
			for (var j = 0; j < table.rows.length; j++) {
				var div = table.rows(j).cells(0).firstChild;
				var Id = div.id.replace("Addons_", "").toLowerCase();
				var item = null;
				var items = te.Data.Addons.getElementsByTagName(Id);
				if (items.length) {
					item = items[0].cloneNode(true);
				}
				if (!item) {
					item = xml.createElement(Id);
				}
				var Enabled = document.getElementById("enable_" + Id).checked;
				if (Enabled) {
					var AddonFolder = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\" + Id);
					Enabled = fso.FolderExists(AddonFolder + "\\lang") ? 2 : 0;
					if (fso.FileExists(AddonFolder + "\\script.vbs")) {
						Enabled |= 8;
					}
					if (fso.FileExists(AddonFolder + "\\script.js")) {
						Enabled |= 1;
					}
					Enabled = (Enabled & 9) ? Enabled : 4;
				} else {
					MainWindow.AddonDisabled(Id);
				}
				item.setAttribute("Enabled", Enabled);
				root.appendChild(item);
			}
			xml.appendChild(root);
			te.Data.Addons = xml;
			MainWindow.RunEvent1("ConfigChanged", "Addons");
		}
	} catch (e) {}
}

function SetChanged(fn, form)
{
	if (g_Chg.Data) {
		g_bChanged = true;
		if (g_x[g_Chg.Data] && g_x[g_Chg.Data].selectedIndex >= 0) {
			(fn || ReplaceX)(g_Chg.Data, form);
		} else {
			AddX(g_Chg.Data, fn, form);
		}
	}
}

function SetData(sel, a, t)
{
	sel.value = PackData(a);
	if (/boolean/.test(typeof t)) {
		t = (t ? String.fromCharCode(9745, 32) : String.fromCharCode(9744, 32)) + a[0];
	}
	sel.text = t || a[0];
}

function PackData(a)
{
	for (var i = a.length; i-- > 0;) {
		a[i] = String(a[i] || "").replace(g_sep, "`  ~");
	}
	return a.join(g_sep);
}

function LoadAddons()
{
	if (g_x.Addons) {
		return;
	}
	g_x.Addons = true;

	var AddonId = [];
	var wfd = api.Memory("WIN32_FIND_DATA");
	var path = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\");
	var hFind = api.FindFirstFile(path + "*", wfd);
	var bFind = hFind != INVALID_HANDLE_VALUE;
	while (bFind) {
		var Id = wfd.cFileName;
		if (Id != "." && Id != ".." && !AddonId[Id]) {
			AddonId[Id] = 1;
		}
		bFind = api.FindNextFile(hFind, wfd);
	}
	api.FindClose(hFind);

	var table = document.getElementById("Addons");
	table.ondragover = Over5;
	table.ondrop = Drop5;
	table.deleteRow(0);
	var root = te.Data.Addons.documentElement;
	if (root) {
		var items = root.childNodes;
		if (items) {
			for (var i = 0; i < items.length; i++) {
				var item = items[i];
				var Id = item.nodeName;
				if (AddonId[Id]) {
					AddAddon(table, Id, api.LowPart(item.getAttribute("Enabled")));
					delete AddonId[Id];
				}
			}
		}
	}
	for (var Id in AddonId) {
		if (fso.FileExists(path + Id + "\\config.xml")) {
			AddAddon(table, Id, false);
		}
	}
}

function AddAddon(table, Id, bEnable, Alt)
{
	SetAddon(Id, bEnable, table.insertRow().insertCell(), Alt);
	if (!Alt) {
		var sorted = document.getElementById("SortedAddons");
		if (sorted.rows.length) {
			SetAddon(Id, bEnable, sorted.insertRow().insertCell(), "Sorted_");
		}
	}
}

function SetAddon(Id, bEnable, td, Alt)
{
	if (!td) {
		td = document.getElementById(Alt || "Addons_" + Id).parentNode;
	}
	var info = GetAddonInfo(Id);
	var bMinVer = info.MinVersion && te.Version < CalcVersion(info.MinVersion);
	if (bMinVer) {
		bEnable = false;
	}
	var s = ['<div ', (Alt ? '' : 'draggable="true" ondragstart="Start5(this)" ondragend="End5(this)"'), ' title="', Id, '" Id="', Alt || "Addons_", Id, '" style="color: ', bEnable ? "": "gray", '">'];
	s.push('<table><tr style="border-top: 1px solid buttonshadow"><td>', (Alt ? '&nbsp;' : '<input type="radio" name="AddonId" id="_' + Id+ '">'), '</td><td style="width: 100%"><label for="_', Id, '">', info.Name, "&nbsp;", info.Version, '<br /><a href="#" onclick="return AddonInfo(\'', Id, '\', this)" style="font-size: .9em">', GetText('Details'), ' (', Id, ')</a>');
	if (bMinVer) {
		s.push('</td><td style="color: red; align: right; white-space: nowrap; vertical-align: middle">', info.MinVersion.replace(/^20/, "Version ").replace(/\.0/g, '.'), ' ', GetText("is required."), '</td>');
	} else if (info.Options) {
		s.push('</td><td style="white-space: nowrap; vertical-align: middle; padding-right: 1em"><a href="#" onclick="AddonOptions(\'', Id, '\'); return false;">', GetText('Options'), '</td>');
	}
	s.push('<td style="vertical-align: middle', bMinVer ? ';display: none"' : "", '"><input type="checkbox" ', (Alt ? "" : 'id="enable_' + Id + '"'), ' onclick="AddonEnable(this, \'', Id, '\')" ', bEnable ? " checked" : "", '></td>');
	s.push('<td style="vertical-align: middle', bMinVer ? ';display: none"' : "", '"><label for="enable_', Id, '" style="display: block; width: 6em; white-space: nowrap">', GetText(bEnable ? "Enabled" : "Enable"), '</label></td>');
	s.push('<td style="vertical-align: middle; padding-right: 1em"><input type="image" src="bitmap:ieframe.dll,216,16,10" title="', GetText('Remove'), '" onclick="AddonRemove(\'', Id, '\')" style="width: 12pt"></td>');
	s.push('</tr></table></label></div>');
	td.innerHTML = s.join("");
	ApplyLang(td);
	if (!Alt) {
		var div = document.getElementById("Sorted_" + Id);
		if (div) {
			SetAddon(Id, bEnable, div.parentNode, "Sorted_");
		}
	}
}

function Start5(o)
{
	if (api.GetKeyState(VK_LBUTTON) < 0) {
		event.dataTransfer.effectAllowed = 'move';
		g_drag5 = o.id;
		return true;
	}
	return false;
}

function End5()
{
	g_drag5 = false;
}

function Over5()
{
	if (g_drag5) {
		if (event.preventDefault) {
			event.preventDefault();
		} else {
 			event.returnValue = false;
		}
	}
}

function Drop5(e)
{
	if (g_drag5) {
		var o = document.elementFromPoint(e.clientX, e.clientY);
		do {
			if (/Addons_/i.test(o.id)) {
				(function (src, dest) { setTimeout(function () {
					AddonMoveEx(src, dest);
				}, 99);}) (GetRowIndexById(g_drag5) , GetRowIndexById(o.id));
				break;
			}
		} while (o = o.parentNode);
	}
}

function GetRowIndexById(id)
{
	try {
		var o = document.getElementById(id);
		if (o) {
			while (o = o.parentNode) {
				if (o.rowIndex !== undefined) {
					return o.rowIndex;
				}
			}
		}
	} catch (e) {}
}

function AddonInfo(Id, o)
{
	o.style.textDecoration = "none";
	o.style.color = "windowtext";
	o.style.cursor = "default";
	o.onclick = null;
	var info = GetAddonInfo(Id);
	o.innerHTML = info.Description;
	return false;
}

function AddonWebsite(Id)
{
	var info = GetAddonInfo(Id);
	wsh.run(info.URL);
}

function AddonEnable(o, Id)
{
	var div = document.getElementById("Addons_" + Id);
	SetAddon(Id, o.checked);
	document.getElementById("enable_" + Id).checked = o.checked;
	g_Chg.Addons = true;
}

function OptionMove(dir)
{
	if (/^1/.test(TabIndex)) {
		var r = document.F.AddonId;
		for (i = 0; i < r.length; i++) {
			if (r[i].checked) {
				if (api.GetKeyState(VK_CONTROL) < 0) {
					if (dir < 0) {
						dir = -i;
					} else {
						dir = document.getElementById("Addons").rows.length - i - 1;
					}
				}
				try {
					AddonMoveEx(i, i + dir);
				} catch (e) {}
	 			break;
			}
		}
	} else if (/^2/.test(TabIndex)) {
		if (g_x.Menus.selectedIndex < 0 || g_x.Menus.selectedIndex + dir < 0 || g_x.Menus.selectedIndex + dir >= g_x.Menus.length) {
			return;
		}
		var src = g_x.Menus[g_x.Menus.selectedIndex];
		var dist = g_x.Menus[g_x.Menus.selectedIndex + dir];
		var text = dist.text;
		var value = dist.value;
		dist.text = src.text;
		dist.value = src.value;
		src.text = text;
		src.value = value;
		g_x.Menus.selectedIndex += dir;
		g_Chg.Menus = true;
	}
}

function AddonMoveEx(src, dest)
{
	var table = document.getElementById("Addons");
	if (dest < 0 || dest >= table.rows.length || src == dest) {
		return false;
	}
	var tr = table.rows(src);
	var td = tr.cells(0);

	var s = td.innerHTML
	var md = td.onmousedown;
	var mu = td.onmouseup;
	var mm = td.onmousemove;

	table.deleteRow(src);

	tr = table.insertRow(dest);
	td = tr.insertCell();
	td.innerHTML = s;
	td.onmousedown = md;
	td.onmouseup = mu;
	td.onmousemove = mm;
	var o = document.getElementById('panel1');
	var i = td.offsetTop - o.scrollTop;
	if (i < 0 || i >= o.offsetHeight - td.offsetHeight) {
		td.scrollIntoView(i < 0);
	}
	document.F.AddonId[dest].checked = true;
	g_Chg.Addons = true;
	return false;
}

function AddonRemove(Id)
{
	if (!confirmOk()) {
		return;
	}
	MainWindow.AddonDisabled(Id);
	if (AddonBeforeRemove(Id) < 0) {
		return;
	}
	var sf = api.Memory("SHFILEOPSTRUCT");
	sf.hwnd = api.GetForegroundWindow();
	sf.wFunc = FO_DELETE;
	sf.fFlags = 0;
	sf.pFrom = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\"+ Id) + "\0";
	if (api.SHFileOperation(sf) == 0) {
		if (!sf.fAnyOperationsAborted) {
			var table = document.getElementById("Addons");
			table.deleteRow(GetRowIndexById("Addons_" + Id));
			var table = document.getElementById("SortedAddons");
			if (table.rows.length) {
				table.deleteRow(GetRowIndexById("Sorted_" + Id));
			}
			g_Chg.Addons = true;
		}
	}
}

function ApplyOptions()
{
	var o = api.GetKeyState(VK_SHIFT) >= 0 ? document.documentElement || document.body : {};
	var hwnd = api.GetWindow(document);
	var hwnd1 = hwnd;
	while (hwnd1 = api.GetParent(hwnd)) {
		hwnd = hwnd1;
	}
	if (!api.IsZoomed(hwnd) && !api.IsIconic(hwnd)) {
		var r = 12 / (Math.abs(MainWindow.DefaultFont.lfHeight) || 12);
		te.Data.Conf_OptWidth = api.LowPart(o.offsetWidth) * r;
		te.Data.Conf_OptHeight = api.LowPart(o.offsetHeight) * r;
	}
	SetChanged(ReplaceMenus);
	for (var i in document.F.elements) {
		if (!/=|:/.test(i)) {
			if (/^Tab_|^Tree_|^View_|^Conf_/.test(i) && !/_$/.test(i)) {
				te.Data[i] = GetElementValue(document.F[i]);
			}
		}
	}
	SaveAddons();
	SaveMenus();
	SetTabControls();
	SetTreeControls();
	SetFolderViews();
	te.Data.bReload = true;
	api.EnableWindow(te.Ctrl(CTRL_WB).hwnd, false);
}

function CancelOptions()
{
	if (te.Data.bErrorAddons) {
		SaveAddons();
		te.Data.bReload = true;
		api.EnableWindow(te.Ctrl(CTRL_WB).hwnd, false);
	}
}

InitOptions = function ()
{
	ApplyLang(document);
	MainWindow.g_.OptionsWindow = window;
	var InstallPath = fso.GetParentFolderName(api.GetModuleFileName(null));
	document.F.ButtonInitConfig.disabled = (InstallPath == te.Data.DataFolder) | !fso.FolderExists(fso.BuildPath(InstallPath, "layout"));
	for (i in document.F.elements) {
		if (!/=|:/.test(i)) {
			if (/^Tab_|^Tree_|^View_|^Conf_/.test(i)) {
				if (te.Data[i] !== undefined) {
					SetElementValue(document.F[i], te.Data[i]);
				}
			}
		}
	}

	ResetForm();
	var s = [];
	for (var i in g_arMenuTypes) {
		var j = g_arMenuTypes[i];
		s.push('<label id="tab2_' + i + '" class="button" style="width: 100%" onmousedown="ClickTree(this, null, \'Menus\');">' + GetText(j) + '</label><br />');
	}
	document.getElementById("tab2_").innerHTML = s.join("");

	AddEventEx(window, "load", function ()
	{
		SetTab(dialogArguments.Data);
	});

	AddEventEx(window, "resize", function ()
	{
		clearTimeout(g_tidResize);
		g_tidResize = setTimeout(function ()
		{
			ClickTree(null, null, null, true);
		}, 500);
	});

	SetOnChangeHandler();
	AddEventEx(window, "beforeunload", function ()
	{
		for (var i in g_.elAddons) {
			try {
				var w = g_.elAddons[i].contentWindow;
				w.g_nResult = g_nResult;
				if (g_nResult == 1) {
					w.TEOk();
				}
				w.close();
			} catch (e) {}
		}
		g_.elAddons = {};
		g_bChanged |= g_Chg.Addons || g_Chg.Menus || g_Chg.Tab || g_Chg.Tree || g_Chg.View;
		SetOptions(ApplyOptions, CancelOptions);
	});
}

OpenIcon = function (o)
{
	setTimeout(function ()
	{
		var data = [];
		var a = o.id.split(/,/);
		if (a[0] == "b") {
			var dllpath = fso.BuildPath(system32, "ieframe.dll");
			var image = api.CreateObject("WICBitmap");
			a[0] = fso.GetFileName(dllpath);
			var a1 = a[1];
			var hModule = LoadImgDll(a, 0);
			if (hModule) {
				var himl = api.ImageList_LoadImage(hModule, isFinite(a[1]) ? a[1] - 0 : a[1], a[2], 0, CLR_DEFAULT, IMAGE_BITMAP, LR_CREATEDIBSECTION);
				if (himl) {
					a[1] = a1;
					var nCount = api.ImageList_GetImageCount(himl);
					a[0] = fso.GetFileName(dllpath);
					for (a[3] = 0; a[3] < nCount; a[3]++) {
						var s = "bitmap:" + a.join(",");
						var src = MakeImgSrc(s, 0, false, a[2]);
						data.push('<img src="' + src + '" class="button" onclick="SelectIcon(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()" title="' + s + '"> ');
					}
					api.ImageList_Destroy(himl);
				}
				api.FreeLibrary(hModule);
			}
		} else {
			dllpath = fso.BuildPath(system32, a[1]);
			var nCount = api.ExtractIconEx(dllpath, -1, null, null, 0);
			for (var i = 0; i < nCount; i++) {
				var s = ["icon:" + a[1], i].join(",");
				var src = MakeImgSrc(s, 0, false, 32);
				data.push('<img src="' + src + '" class="button" onclick="SelectIcon(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()" title="' + s + '"> ');
			}
		}
		o.innerHTML = data.join("");
		o.cursor = "";
		o.onclick = null;
		document.body.style.cursor = "auto";
	}, 1);
	document.body.style.cursor = "wait";
}

InitDialog = function ()
{
	var Query = String(dialogArguments.Query || location.search.replace(/\?/, "")).toLowerCase();
	if (Query == "icon") {
		var a =
		{
			"16px ieframe,206" : "b,206,16",
			"24px ieframe,204" : "b,204,24",
			"16px ieframe,216" : "b,216,16",
			"24px ieframe,214" : "b,214,24",
			"16px ieframe,699" : "b,699,16",
			"24px ieframe,697" : "b,697,24",

			"shell32" : "i,shell32.dll",
			"imageres" : "i,imageres.dll",
			"wmploc" : "i,wmploc.dll",
			"setupapi" : "i,setupapi.dll",
			"dsuiext" : "i,dsuiext.dll",
			"inetcpl" : "i,inetcpl.cpl",

			"25px TRAVEL_ENABLED_XP" : "b,TRAVEL_ENABLED_XP.BMP,25",
			"30px TRAVEL_ENABLED_XP" : "b,TRAVEL_ENABLED_XP_120.BMP,30"
		};
		var s = [];
		for (var i in a) {
			s.push('<div id="' + a[i] + '" onclick="OpenIcon(this)" style="cursor: pointer"><span class="tab">' + i + '</span></div>');
		}
		document.getElementById("Content").innerHTML = s.join("");
	}
	if (Query == "mouse") {
		document.getElementById("Content").innerHTML = '<div id="Gesture" style="width: 100%; height: 100%; text-align: center" onmousedown="return MouseDown()" onmouseup="return MouseUp()" onmousemove="return MouseMove()" ondblclick="MouseDbl()" onmousewheel="return MouseWheel()"></div>';
		document.getElementById("Selected").innerHTML = '<input type="text" name="q" style="width: 100%" autocomplete="off" onkeydown="setTimeout(\'returnValue=document.F.q.value\',100)" />';
	}
	if (Query == "key") {
		returnValue = false;
		document.getElementById("Content").innerHTML = '<div style="padding: 8px;" style="display: block;"><label>Key</label><br /><input type="text" name="q" autocomplete="off" style="width: 100%; ime-mode: disabled" onfocus="this.blur()" /></div>';
		var fn = function (e)
		{
			var key = (e || event).keyCode;
			var k = api.MapVirtualKey(key, 0) | ((key >= 33 && key <= 46 || key >= 91 && key <= 93 || key == 111 || key == 144) ? 256 : 0);
			if (k == 42 || k == 29 || k == 56 || k == 347) {
				return false;
			}
			var s = api.sprintf(10, "$%x", k | GetKeyShift());
			returnValue = GetKeyName(s);
			if (/^\$\w02a$|^\$\w01d$|^\$\w038$/i.test(returnValue)) {
				returnValue = GetKeyName(s.substr(0, 3) + "1e").replace(/\+A$/, "");
			}
			document.F.q.value = returnValue;
			document.F.q.title = s;
			document.F.ButtonOk.disabled = false;
			return false;
		}
		AddEventEx(document.body, "keydown", fn);
		AddEventEx(document.body, "keyup", fn);
	}
	if (Query == "new") {
		returnValue = false;
		var s = [];
		s.push('<div style="padding: 8px;" style="display: block;"><input type="radio" name="mode" id="folder" onclick="document.F.path.focus()"><label for="folder">New Folder</label> <input type="radio" name="mode" id="file" onclick="document.F.path.focus()"><label for="file">New File</label><br />', dialogArguments.path ,'<br /><input type="text" name="path" style="width: 100%" /></div>');
		document.getElementById("Content").innerHTML = s.join("");
		AddEventEx(document.body, "keydown", function (e)
		{
			setTimeout(function ()
			{
				document.F.ButtonOk.disabled = !document.F.path.value;
			}, 99);
			var key = (e || event).keyCode;
			if (key == VK_RETURN && document.F.path.value) {
				SetResult(1);
			}
			if (key == VK_ESCAPE) {
				SetResult(2);
			}
			return true;
		});

		AddEventEx(document.body, "paste", function ()
		{
			setTimeout(function ()
			{
				document.F.ButtonOk.disabled = !document.F.path.value;
			}, 99);
		});

		setTimeout(function ()
		{
			document.F[dialogArguments.Mode].checked = true;
			document.F.path.focus();
		}, 99);

		AddEventEx(window, "beforeunload", function ()
		{
			if (g_nResult == 1) {
				path = document.F.path.value;
				if (path) {
					if (!/^[A-Z]:\\|^\\/i.test(path)) {
						path = fso.BuildPath(dialogArguments.path, path.replace(/^\s+/, ""));
					}
					if (document.getElementById("folder").checked) {
						CreateFolder(path);
					} else if (document.getElementById("file").checked) {
						CreateFile(path);
					}
				}
			}
		});
	}
	if (Query == "about") {
		returnValue = false;
		var s = ['<table style="border-spacing: 2em; border-collapse: separate; width: 100%"><tr><td>'];
		var src = MakeImgSrc(api.GetModuleFileName(null), 0, 32);
		s.push('<img src="', src , '"></td><td><span style="font-weight: bold; font-size: 120%">', te.About, '</span> (', (api.sizeof("HANDLE") * 8), 'bit)<br>');
		s.push('<br><a href="javascript:Run(0)">', api.GetModuleFileName(null), '</a><br>');
		s.push('<br><a href="javascript:Run(1)">', fso.BuildPath(te.Data.DataFolder, "config"), '</a><br>');
		s.push('<br><label>Information</label><input type="text" value="', GetTEInfo(), '" style="width: 100%" onclick="this.select()" readonly /><br>');
		var root = te.Data.Addons.documentElement;
		if (root) {
			var ar = [];
			var items = root.childNodes;
			if (items) {
				for (var i = 0; i < items.length; i++) {
					if (api.LowPart(items[i].getAttribute("Enabled"))) {
						ar.push(items[i].nodeName);
					}
				}
			}
			s.push('<br><label>Add-ons</label><input type="text" value="', ar.join(","), '" style="width: 100%" onclick="this.select()" /><br>');
		}
		s.push('<br><input type="button" value="Visit website" onclick="Run(2)">');
		s.push('&nbsp;<input type="button" value="Check for updates" onclick="Run(3)">');
		s.push('</td></tr></table>');
		document.getElementById("Content").innerHTML = s.join("");
		document.F.ButtonOk.disabled = false;
		document.getElementById("buttonCancel").style.display = "none";

		Run = function (n)
		{
			var ar = [
				'Navigate(api.GetModuleFileName(null) + "\\\\..", SBSP_NEWBROWSER)',
				'Navigate(fso.BuildPath(te.Data.DataFolder, "config"), SBSP_NEWBROWSER)',
				'wsh.Run("http://www.eonet.ne.jp/~gakana/tablacus/explorer_en.html")',
				'CheckUpdate()'
			];
			MainWindow.setTimeout(ar[n], 500);
			window.close();
		}
	}
	if (dialogArguments.element) {
		AddEventEx(window, "beforeunload", function ()
		{
			try {
				if (g_nResult == 1) {
					dialogArguments.element.value = returnValue;
					if (dialogArguments.element.onchange) {
						dialogArguments.element.onchange();
					}
				}
			} catch (e) {}
		});
	}
	DialogResize = function ()
	{
		CalcElementHeight(document.getElementById("panel0"), 3);
	};
	AddEventEx(window, "resize", function ()
	{
		clearTimeout(g_tidResize);
		g_tidResize = setTimeout(DialogResize, 500);
	});
	ApplyLang(document);
	DialogResize();
}

MouseDown = function ()
{
	if (g_Gesture) {
		var n = 1;
		var ar = [0, 0, 2, 1];
		for (i = 1; i < 6; i++) {
			if (g_Gesture.indexOf(i + "") < 0) {
				if ((event.buttons && event.buttons & n) || event.button == ar[i]) {
					returnValue += i + "";
				}
			}
			n *= 2;
		}
	} else {
		returnValue = GetGestureKey() + GetGestureButton();
		api.RedrawWindow(api.GetWindow(document), null, 0, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
	}
	document.F.q.value = returnValue;
	g_Gesture = returnValue;
	g_pt.x = event.clientX;
	g_pt.y = event.clientY;
	document.F.ButtonOk.disabled = false;
	var o = document.getElementById("Gesture");
	var s = o.style.height;
	o.style.height = "1px";
	o.style.height = s;
	return false;
}

MouseUp = function ()
{
	g_Gesture = GetGestureButton();
	return false;
}

MouseMove = function ()
{
	if (api.GetKeyState(VK_XBUTTON1) < 0 || api.GetKeyState(VK_XBUTTON2) < 0) {
		returnValue = GetGestureKey() + GetGestureButton();
		document.F.q.value = returnValue;
	}
	if (document.F.q.value.length && (api.GetKeyState(VK_RBUTTON) < 0 || (te.Data.Conf_Gestures && (api.GetKeyState(VK_MBUTTON) < 0)))) {
		var pt = {x: event.clientX, y: event.clientY};
		var x = (pt.x - g_pt.x);
		var y = (pt.y - g_pt.y);
		if (Math.abs(x) + Math.abs(y) >= 20) {
			if (te.Data.Conf_TrailSize) {
				var hwnd = api.GetWindow(document);
				var hdc = api.GetWindowDC(hwnd);
				if (hdc) {
					api.MoveToEx(hdc, g_pt.x, g_pt.y, null);
					var pen1 = api.CreatePen(PS_SOLID, te.Data.Conf_TrailSize, te.Data.Conf_TrailColor);
					var hOld = api.SelectObject(hdc, pen1);
					api.LineTo(hdc, pt.x, pt.y);
					api.SelectObject(hdc, hOld);
					api.DeleteObject(pen1);
					api.ReleaseDC(hwnd, hdc);
				}
			}
			g_pt = pt;
			var s = (Math.abs(x) >= Math.abs(y)) ? ((x < 0) ? "L" : "R") :  ((y < 0) ? "U" : "D");
			if (s != document.F.q.value.charAt(document.F.q.value.length - 1)) {
				returnValue += s;
				document.F.q.value = returnValue;
			}
		}
	}
	return false;
}

MouseDbl = function ()
{
	returnValue += returnValue.replace(/\D/g, "");
	document.F.q.value = returnValue;
	return false;
}

MouseWheel = function ()
{
	returnValue = GetGestureKey() + GetGestureButton() + (event.wheelDelta > 0 ? "8" : "9");
	document.F.q.value = returnValue;
	document.F.ButtonOk.disabled = false;
	return false;
}

InitLocation = function ()
{
	var ar = [], param = [];
	Addon_Id = dialogArguments.Data.id;
	LoadAddon("js", Addon_Id, ar, param);
	if (ar.length) {
		(function (ar) { setTimeout(function () {
			MessageBox(ar.join("\n\n"), TITLE, MB_OK);
		}, 500);}) (ar);
	}
	ar = [];
	var s = "CSA";
	for (var i = 0; i < s.length; i++) {
		ar.push('<input type="button" value="', MainWindow.g_.KeyState[i][0],'" title="', s.charAt(i), '" onclick="AddMouse(this)" />');
	}
	document.getElementById("__MOUSEDATA").innerHTML = ar.join("");
	ApplyLang(document);
	var info = GetAddonInfo(dialogArguments.Data.id);
	document.title = info.Name;
	var items = te.Data.Addons.getElementsByTagName(dialogArguments.Data.id);
	var item = null;
	if (items.length) {
		item = items[0];
		var Location = item.getAttribute("Location");
		if (!Location) {
			Location = param.Default;
		}
		for (var i = document.L.elements.length; i--;) {
			if (api.StrCmpI(Location, document.L[i].value) == 0) {
				document.L[i].checked = true;
			}
		}
	}
	var locs = [];
	items = te.Data.Locations;
	for (var i in items) {
		locs[i] = [];
		for (var j in items[i]) {
			var ar = items[i][j].split("\0");
			info = GetAddonInfo(ar[0]);
			if (ar[1]) {
				locs[i].push('<img src="', ar[1], '" title="', EncodeSC(info.Name || ""), '" class="img1"> ');
			} else {
				locs[i].push('<span class="text1">', info.Name, '</span> ');
			}
		}
	}
	for (var i in locs) {
		var s = locs[i].join("");
		try {
			var o = document.getElementById('_' + i);
			ApplyLang(o);
			o.parentNode.title = o.innerHTML.replace(/<[^>]*>|[\r\n]|\s\s+/g, "");
			o.innerHTML = s;
		} catch (e) {}
	}
	var oa = document.F.Menu;
	oa.length = 0;
	var o = oa[++oa.length - 1];
	o.value = "";
	o.text = GetText("Select");
	for (var j in g_arMenuTypes) {
		var s = g_arMenuTypes[j];
		if (!/Default|Alias/.test(s)) {
			o = oa[++oa.length - 1];
			o.value = s;
			o.text = GetText(s);
		}
	}
	var ar = ["Key", "Mouse"];
	for (i in ar) {
		var mode = ar[i];
		var oa = document.F[mode + "On"];
		oa.length = 0;
		o = oa[++oa.length - 1];
		o.value = "";
		o.text = GetText("Select");
		for (var j in MainWindow.eventTE[mode]) {
			o = oa[++oa.length - 1];
			o.text = GetTextEx(j);
			o.value = j;
		}
	}
	if (item) {
		var ele = document.F.elements;
		for (var i = ele.length; i--;) {
			var n = ele[i].id || ele[i].name;
			if (n && !/=/.test(n)) {
				s = item.getAttribute(n);
				if (/Name$/.test(n)) {
					s = GetText(s);
				}
				if (n == "Key") {
					s = GetKeyName(s);
				}
				if (s || s === 0) {
					SetElementValue(ele[n], s);
				}
			}
		}
	}
	LoadChecked(document.F);

	if (!dialogArguments.Data.show) {
		dialogArguments.Data.show = "6";
		dialogArguments.Data.index = 6;
	}
	if (/[8]/.test(dialogArguments.Data.show)) {
		MakeKeySelect();
		SetKeyShift();
	}
	var a = document.F.MenuName.value.split(/\t/);
	document.F._MenuName.value = GetText(a[0]);

	try {
		var ar = dialogArguments.Data.show.split(/,/);
		for (var i in ar) {
			document.getElementById("tab" + ar[i]).style.display = "inline";
		}
		nTabIndex = dialogArguments.Data.index;
	} catch (e) {}
	SetImage();

	SetOnChangeHandler();
	AddEventEx(window, "load", function ()
	{
		ApplyLang(document);
		ClickTab(null, 1);
	});
	AddEventEx(window, "resize", function ()
	{
		clearTimeout(g_tidResize);
		g_tidResize = setTimeout(ResizeTabPanel, 500);
	});

	AddEventEx(window, "beforeunload", function ()
	{
		SetOptions(function () {
			if (window.SaveLocation) {
				window.SaveLocation();
			}
			delete MainWindow.g_.OptionsWindow;
			var items = te.Data.Addons.getElementsByTagName(dialogArguments.Data.id);
			if (items.length) {
				var item = items[0];
				item.removeAttribute("Location");
				for (var i = document.L.elements.length; i--;) {
					if (document.L[i].checked) {
						item.setAttribute("Location", document.L[i].value);
						te.Data.bReload = true;
						MainWindow.RunEvent1("ConfigChanged", "Addons");
						break;
					}
				}
				var ele = document.F.elements;
				ele.MenuName.value = GetSourceText(ele._MenuName.value);
				if (dialogArguments.Data.show == "6") {
					ele.Set.value = "";
				}
				for (var i = ele.length; i--;) {
					var n = ele[i].id || ele[i].name;
					if (n && n.charAt(0) != "_") {
						if (n == "Key") {
							var s = GetKeyKey(document.F[n].value);
							if (s) {
								document.F[n].value = api.sprintf(10, "$%x", s);
							}
						}
						if (SetAttribEx(item, document.F, n)) {
							te.Data.bReload = true;
							MainWindow.RunEvent1("ConfigChanged", "Addons");
						}
					}
				}
			}
		});
	});
	if (item) {
		InitColor1(item);
	}
}

function SetAttrib(item, n, s)
{
	if (s) {
		item.setAttribute(n, s);
	} else {
		item.removeAttribute(n);
	}
}

function GetElementValue(o)
{
	if (o.type) {
		if (/checkbox/i.test(o.type)) {
			return o.checked ? 1 : 0;
		}
		if (/hidden|text|number|url|password|range|color|date|time/i.test(o.type)) {
			return o.value;
		}
		if (/select/i.test(o.type)) {
			return o[o.selectedIndex].value;
		}
	}
}

function SetElementValue(o, s)
{
	if (o && o.type) {
		if (/checkbox/i.test(o.type)) {
			o.checked = api.LowPart(s);
			return;
		}
		if (/select/i.test(o.type)) {
			for (var i = o.options.length; i-- > 0;) {
				if (o.options[i].value == s) {
					o.selectedIndex = i;
					break;
				}
			}
			return;
		}
		o.value = s;
	}
}

function SetAttribEx(item, f, n)
{
	var s = GetElementValue(f[n]);
	if (s != GetAttribEx(item, f, n)) {
		SetAttrib(item, n, s);
		return true;
	}
	return false;
}

function GetAttribEx(item, f, n)
{
	var res  = /([^=]*)=(.*)/.exec(n);
	if (res) {
		s = item.getAttribute(res[1]);
		if (s == res[2]) {
			document.getElementById(n).checked = true;
		}
		return;
	}
	s = item.getAttribute(n);
	if (s || s === 0) {
		if (n == "Key") {
			s = GetKeyName(s);
		}
		SetElementValue(f[n], s);
	}
}

function RefX(Id, bMultiLine, oButton, bFilesOnly, Filter)
{
	setTimeout(function () {
		if (/Path/.test(Id)) {
			var s = Id.replace("Path", "Type");
			var o = GetElement(s);
			if (o) {
				var pt;
				if (oButton) {
					pt = GetPos(oButton, true);
					pt = {x: pt.x, y: pt.y + oButton.offsetHeight, width: oButton.offsetWidth };
				} else {
					pt = api.Memory("POINT");
					api.GetCursorPos(pt);
				}
				var r = MainWindow.OptionRef(o[o.selectedIndex].value, GetElement(Id).value, pt);
				if (/string/i.test(typeof r)) {
					var p = { s: r };
					MainWindow.OptionDecode(o[o.selectedIndex].value, p);
					if (bMultiLine && api.GetKeyState(VK_CONTROL) < 0 && api.ILCreateFromPath(p.s)) {
						AddPath(Id, p.s);
					} else {
						SetValue(GetElement(Id), p.s);
					}
				}
				return;
			}
		}

		var o = GetElement(Id);
		var path = OpenDialogEx(o.value, Filter, api.LowPart(bFilesOnly));
		if (path) {
			if (bMultiLine) {
				AddPath(Id, path);
			} else {
				SetValue(o, path);
			}
		}
	}, 99);
}

function PortableX(Id)
{
	if (!confirmOk()) {
		return;
	}
	var o = GetElement(Id);
	var s = fso.GetDriveName(api.GetModuleFileName(null));
	SetValue(o, o.value.replace(wsh.ExpandEnvironmentStrings("%UserProfile%"), "%UserProfile%").replace(new RegExp('^("?)' + s, "igm"), "$1%Installed%").replace(new RegExp('( "?)' + s, "igm"), "$1%Installed%"));
}

function GetElement(Id)
{
	return (document.F && document.F[Id]) || document.getElementById(Id) || (document.E && document.E[Id]);
}

function AddPath(Id, strValue)
{
	var o = GetElement(Id);
	if (o) {
		var s = o.value;
		if (/\n$/.test(s) || s == "") {
			s += strValue;
		} else {
			s += "\n" + strValue;
		}
		SetValue(o, s);
	}
}

function SetValue(o, s)
{
	if (o.value != s) {
		o.value = s;
		FireEvent(o, "change");
	}
}

function GetCurrentSetting(s)
{
	var FV = te.Ctrl(CTRL_FV);

	if (confirmOk()) {
		AddPath(s, api.PathQuoteSpaces(api.GetDisplayNameOf(FV.FolderItem, SHGDN_FORPARSINGEX | SHGDN_FORPARSING)));
	}
}

function SetTab(s)
{
	var o = null;
	var arg = String(s).split(/&/);
	for (var i in arg) {
		var ar = arg[i].split(/=/);
		if (ar[0].toLowerCase() == "tab") {
			if (api.StrCmpI(ar[1], "Get Addons") == 0) {
				o = document.getElementById('tab1_1');
			}
			var s = GetText(ar[1]);
			var ovTab;
			for (var j = 0; ovTab = document.getElementById('tab' + j); j++) {
				if (api.StrCmpI(s, ovTab.innerText) == 0) {
					o = ovTab;
					break;
				}
			}
		} else if (api.StrCmpI(ar[0], "menus") == 0) {
			g_MenuType = ar[1];
		}
	}
	ClickTree(o);
}

function AddMouse(o)
{
	(document.F["MouseMouse"] || document.F["Mouse"]).value += o.title;
}

function InitAddonOptions(bFlag)
{
	returnValue = false;
	LoadLang2(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\" + Addon_Id + "\\lang\\" + GetLangId() + ".xml"));

	ApplyLang(document);
	info = GetAddonInfo(Addon_Id);
	document.title = info.Name;
	SetOnChangeHandler();
	AddEventEx(window, "beforeunload", function ()
	{
		SetOptions(SetAddonOptions);
	});
	var items = te.Data.Addons.getElementsByTagName(Addon_Id);
	if (items.length) {
		InitColor1(items[0]);
	}
}

function SetOnChangeHandler()
{
	g_nResult = 3;
	g_bChanged = false;
	var ar = ["input", "select", "textarea"];
	for (var j in ar) {
		var o = document.getElementsByTagName(ar[j]);
		if (o) {
			for (var i = o.length; i--;) {
				if ((o[i].name || o[i].id) && o[i].name != "List" && !/^_/.test(o[i].id)) {
					AddEventEx(o[i], "change", function ()
					{
						g_bChanged = true;
					});
				}
			}
		}
	}
}
function SetAddonOptions()
{
	if (!g_bClosed) {
		var items = te.Data.Addons.getElementsByTagName(Addon_Id);
		if (items.length) {
			var item = items[0];
			var ele = document.F.elements;
			for (var i = ele.length; i--;) {
				var n = ele[i].id || ele[i].name;
				if (n) {
					if (SetAttribEx(item, document.F, n)) {
						returnValue = true;
					}
				}
			}
		}
		g_bClosed = true;
		if (returnValue) {
			TEOk();
		}
		window.close();
	}
}

function SelectIcon(o)
{
	returnValue = o.title;
	document.F.ButtonOk.disabled = false;
	document.getElementById("Selected").innerHTML = o.outerHTML;
	DialogResize();
}

TestX = function (id)
{
	if (confirmOk()) {
		var o = document.F[id + "Type"];
		var p = { s: document.F[id + "Path"].value };
		MainWindow.OptionEncode(o[o.selectedIndex].value, p);
		MainWindow.focus();
		MainWindow.Exec(te.Ctrl(CTRL_FV), p.s, o[o.selectedIndex].value);
		focus();
	}
}

SetImage = function ()
{
	var o = document.getElementById("_Icon");
	if (o) {
		var h = api.LowPart(document.F.IconSize ? document.F.IconSize.value : document.F.Height.value);
		if (!h) {
			h = window.IconSize ? window.IconSize : 24;
		}
		var src = MakeImgSrc(api.PathUnquoteSpaces(document.F.Icon.value), 0, true, h);
		o.innerHTML = src ? '<img src="' + src + '" ' + (h ? 'height="' + h + 'px"' : "") + '>' : "";
	}
}

ShowIcon = ShowIconEx;

function SelectLangID(o)
{
	var i = 0;
	var Langs = [];
	var wfd = api.Memory("WIN32_FIND_DATA");
	var hFind = api.FindFirstFile(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "lang\\*.xml"), wfd);
	var bFind = hFind != INVALID_HANDLE_VALUE;
	while (bFind) {
		Langs.push(wfd.cFileName.replace(/\..*$/, ""));
		bFind = api.FindNextFile(hFind, wfd);
	}
	api.FindClose(hFind);
	Langs.sort();
	var path = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "lang\\");
	var hMenu = api.CreatePopupMenu();
	for (i in Langs) {
		var xml = te.CreateObject("Msxml2.DOMDocument");
		xml.async = false;
		var title = Langs[i];
		xml.load(path + title + '.xml');
		var items = xml.getElementsByTagName('lang');
		if (items && items.length) {
			var item = items[0];
			var en = item.getAttribute("en");
			en = (en && api.StrCmpI(item.text, en)) ? ' / ' + en : "";
			title = item.text + en + " (" + title + ")\t" + item.getAttribute("author");
		}
		api.InsertMenu(hMenu, i, MF_BYPOSITION | MF_STRING, api.LowPart(i) + 1, title);
	}
	var pt = GetPos(o, true);
	var nVerb = api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y + o.offsetHeight, te.hwnd, null, null);
	if (nVerb) {
		document.F.Conf_Lang.value = Langs[nVerb - 1];
		g_bChanged = true;
	}
	api.DestroyMenu(hMenu);
}

function GetTextEx(s)
{
	var ar = s.split(/_/);
	var s = GetText(ar.shift());
	if (ar && ar.length) {
		s += "(" + GetText(ar.join(" ")) + ")";
	}
	return s;
}

function GetAddons()
{
	MainWindow.OpenHttpRequest(urlAddons + "index.xml", "http", AddonsList);
}

function UpdateAddon(Id, o)
{
	if (!o) {
		AddAddon(document.getElementById("Addons"), Id, "Disable");
		g_Chg.Addons = true;
	}
}

function CheckAddon(Id)
{
	return fso.FileExists(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\" + Id + "\\config.xml"));
}

function AddonsSearch()
{
	if (xmlAddons) {
		AddonsAppend()
	} else {
		GetAddons();
	}
	return true;
}

function AddonsList(xhr2)
{
	if (xmlAddons) {
		return;
	}
	if (xhr2) {
		xhr = xhr2;
	}
	var xml = xhr.get_responseXML ? xhr.get_responseXML() : xhr.responseXML;
	if (xml) {
		xmlAddons = xml.getElementsByTagName("Item");
		AddonsAppend()
	}
}

function AddonsAppend()
{
	var Progress = te.ProgressDialog;
	var i = 0, td = [];
	Progress.StartProgressDialog(te.hwnd, null, 2);
	try {
		Progress.SetAnimation(hShell32, 150);
		Progress.SetLine(1, api.LoadString(hShell32, 13585) || api.LoadString(hShell32, 6478), true);
		while (!Progress.HasUserCancelled() && xmlAddons[i]) {
			ArrangeAddon(xmlAddons[i++], td, Progress);
			Progress.SetTitle(Math.floor(100 * i / xmlAddons.length) + "%");
			Progress.SetProgress(i, xmlAddons.length);
		}
		if (g_nSort2 == 1) {
			td.sort(function (a, b) {
				return b[0] - a[0];
			});
		} else {
			td.sort();
		}
		var table = document.getElementById("Addons1");
		if (table) {
			while (table.rows.length > 0) {
				table.deleteRow(0);
			}
			for (i = 0; i < td.length; i++) {
				var tr = table.insertRow(i);
				var td1 = tr.insertCell(0);
				td[i].shift();
				td1.innerHTML = td[i].join("");
				ApplyLang(td1);
				td1.className = (i & 1) ? "oddline" : "";
				td1.style.borderTop = "1px solid buttonshadow";
				td1.style.padding = "3px";
			}
		}
	} catch (e) {
		ShowError(e);
	}
	Progress.StopProgressDialog();
}

function ArrangeAddon(xml, td, Progress)
{
	var Id = xml.getAttribute("Id");
	var s = [];
	var strUpdate = "";
	var nInsert;
	if (Search(xml)) {
		var info = [];
		for (var i = arLangs.length; i--;) {
			GetAddonInfo2(xml, info, arLangs[i]);
		}
		var pubDate = "";
		var dt = new Date(info.pubDate);
		Progress.SetLine(2, info.Name, true);
		if (info.pubDate) {
			pubDate = api.GetDateFormat(LOCALE_USER_DEFAULT, 0, dt, api.GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE)) + " ";
		}
		s.push('<table width="100%"><tr><td width="100%"><b style="font-size: 1.3em">', info.Name, "</b>&nbsp;", info.Version, "&nbsp;", info.Creator, "<br />", info.Description, "<br />");
		if (info.Details) {
			s.push('<a href="#" title="', info.Details, '" onclick="wsh.run(this.title); return false;">', GetText('Details'), '</a><br />');
		}
		s.push(pubDate, '</td><td align="right">');
		var filename = info.filename;
		if (!filename) {
			filename = Id + '_' + info.Version.replace(/\D/, '') + '.zip';
		}
		var dt2 = (dt.getTime() / (24 * 60 * 60 * 1000)) - info.Version;
		var bUpdate = false;
		if (CheckAddon(Id)) {
			installed = GetAddonInfo(Id);
			if (installed.Version >= info.Version) {
				try {
					if (!installed.DllVersion) {
						return;
					}
					var path = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\" + Id + "\\t" + Id + (api.sizeof("HANDLE") * 8) + ".dll");
					if (CalcVersion(installed.DllVersion) <= CalcVersion(fso.GetFileVersion(path))) {
						return;
					}
				} catch (e) {
					return;
				}
			}
			strUpdate = '<br /><b id="_Addons_' + Id + '" style="color: red; white-space: nowrap;">' + GetText('Update available') + '</b>';
			dt2 += MAXINT * 2;
			bUpdate = true;
		} else {
			dt2 += MAXINT;
		}
		if (info.MinVersion && te.Version >= CalcVersion(info.MinVersion)) {
			s.push('<input type="button" onclick="Install(this,', bUpdate, ')" title="', Id, '_', info.Version, '" value="', GetText("Install"), '">');
		} else {
			s.push('<input type="button" style="color: red" onclick="MainWindow.CheckUpdate()" value="', info.MinVersion.replace(/^20/, "Version ").replace(/\.0/g, '.'), ' ', GetText("is required."), '">');
		}
		s.push(strUpdate, '</td></tr></table>');
		s.unshift(g_nSort2 == 1 ? dt2 : g_nSort2 ? Id : info.Name);
		td.push(s);
	}
}

function Search(xml)
{
	var q = document.getElementById('_GetAddonsQ');
	if (!q) {
		return true;
	}
	q = q.value.toUpperCase();
	if (q == "") {
		return true;
	}
	for (var k = arLangs.length; k-- > 0;) {
		var items = xml.getElementsByTagName(arLangs[k]);
		if (items.length) {
			var item = items[0].childNodes;
			for (var i = item.length; i-- > 0;) {
				var item1 = item[i];
				if (item1.tagName) {
					if ((item1.textContent || item1.text).toUpperCase().indexOf(q) >= 0) {
						return true;
					}
				}
			}
		}
	}
	return false;
}

function Install(o, bUpdate)
{
	if (!bUpdate && !confirmOk("Do you want to install it now?")) {
		return;
	}
	var Id = o.title.replace(/_.*$/, "");
	MainWindow.AddonDisabled(Id);
	if (AddonBeforeRemove(Id) < 0) {
		return;
	}
	var file = o.title.replace(/\./, "") + '.zip';
	MainWindow.OpenHttpRequest(urlAddons + Id + '/' + file, "http", Install2, o);
}

function Install2(xhr, url, o)
{
	var Id = o.title.replace(/_.*$/, "");
	var file = o.title.replace(/\./, "") + '.zip';
	var temp = fso.BuildPath(wsh.ExpandEnvironmentStrings("%TEMP%"), "tablacus");
	CreateFolder(temp);
	var dest = fso.BuildPath(temp, Id);
	DeleteItem(dest);
	var hr = MainWindow.Extract(fso.BuildPath(temp, file), temp, xhr);
	if (hr) {
		MessageBox([api.LoadString(hShell32, 4228).replace(/^\t/, "").replace("%d", api.sprintf(99, "0x%08x", hr)), GetText("Extract"), file].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
		return;
	}
	var configxml = dest + "\\config.xml";
	var nDog = 300;
	while (!fso.FileExists(configxml)) {
		if (wsh.Popup(GetText("Please wait."), 1, TITLE, MB_ICONINFORMATION | MB_OKCANCEL) == IDCANCEL || nDog-- == 0) {
			return;
		}
	}
	api.SHFileOperation(FO_MOVE, dest, fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons"), FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR, false);
	o.disabled = true;
	o.value = GetText("Installed");
	o = document.getElementById('_Addons_' + Id);
	if (o) {
		o.style.display = "none";
	}
	UpdateAddon(Id, o);
}

function EnableSelectTag(o)
{
	if (o) {
		var hwnd = api.GetWindow(document);
		api.SendMessage(hwnd, WM_SETREDRAW, false, 0);
		o.style.visibility = "hidden";
		setTimeout(function () {
			o.style.visibility = "visible";
			api.SendMessage(hwnd, WM_SETREDRAW, true, 0);
			api.RedrawWindow(hwnd, null, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
		}, 99);
	}
}

SetResult = function (i)
{
	g_nResult = i;
	window.close();
}

function InitColor1(item)
{
	var ele = document.F.elements;
	for (var i = ele.length; i--;) {
		var n = ele[i].id || ele[i].name;
		if (n) {
			GetAttribEx(item, document.F, n);
		}
	}
	for (var i = ele.length; i--;) {
		var n = ele[i].id || ele[i].name;
		if (n) {
			var res = /^Color_(.*)/.exec(n);
			if (res) {
				var o = document.F[res[1]];
				if (o) {
					ele[i].style.backgroundColor = GetWebColor(o.value);
				}
			}
		}
	}
}

function ChangeColor1(ele)
{
	var o = document.getElementById("Color_" + (ele.id || ele.name));
	if (o) {
		o.style.backgroundColor = GetWebColor(ele.value);
	}
}

function EnableInner()
{
	ChangeForm([["__Inner", "disabled", false]]);
}

function ChangeForm(ar)
{
	AddEventEx(window, "load", function ()
	{
		for (var i in ar) {
			var o = document.getElementById(ar[i][0]);
			for (var s = ar[i][1].split("/"); s.length > 1;) {
				o = o[s.shift()];
			}
			o[s[0]] = ar[i][2];
		}
	});
}

function SetTabContents(id, name, value)
{
	document.getElementById("tab" + id).value = GetText(name);
	document.getElementById("panel" + id).innerHTML = value.join ? value.join('') : value;
}

function ShowButtons(b1, b2, SortMode)
{
	if (SortMode) {
		g_SortMode = SortMode;
	}
	var o = document.getElementById("SortButton");
	o.style.display = b1 ? "inline-block" : "none";
	if (g_SortMode == 1) {
		var table = document.getElementById("Addons");
		var bSorted = /none/i.test(table.style.display);
		document.getElementById("MoveButton").style.display = (b1 || b2) && !bSorted ? "inline-block" : "none";
		for (var i = 3; i--;) {
			o = document.getElementById("SortButton_" + i);
			o.style.border = bSorted && g_nSort == i ? "1px solid highlight" : "";
			o.style.padding = bSorted && g_nSort == i ? "0" : "";
		}
	} else {
		document.getElementById("MoveButton").style.display = b2 ? "inline-block" : "none";
		for (var i = 3; i--;) {
			o = document.getElementById("SortButton_" + i);
			o.style.border = g_nSort2 == i ? "1px solid highlight" : "";
			o.style.padding = g_nSort2 == i ? "0" : "";
		}
	}
}

function SortAddons(n)
{
	if (g_SortMode == 1) {
		var table = document.getElementById("Addons");
		if (table.rows.length < 2) {
			return;
		}
		var sorted = document.getElementById("SortedAddons");
		while (sorted.rows.length > 0) {
			sorted.deleteRow(0);
		}
		if (/none/i.test(table.style.display) && g_nSort == n) {
			table.style.display = "";
			sorted.style.display = "none";
		} else {
			g_nSort = n;
			var s, ar = [];
			for (var j = table.rows.length; j--;) {
				var div = table.rows(j).cells(0).firstChild;
				var Id = div.id.replace("Addons_", "").toLowerCase();
				if (g_nSort == 0) {
					s = table.rows(j).cells(0).innerText;
				} else if (g_nSort == 1) {
					s = "";
					var info = GetAddonInfo(Id);
					if (info.pubDate) {
						var dt = new Date(info.pubDate);
						s = api.GetDateFormat(LOCALE_USER_DEFAULT, 0, dt, "yyyyMMdd");
					}
				} else {
					s = Id;
				}
				ar.push(s + "\t" + Id);
			}
			ar.sort();
			if (g_nSort == 1) {
				ar = ar.reverse();
			}
			var i = 0, td = [];
			var Progress = te.ProgressDialog;
			Progress.SetAnimation(hShell32, 150);
			Progress.StartProgressDialog(te.hwnd, null, 2);
			try {
				for (var i in ar) {
					if (Progress.HasUserCancelled()) {
						break;
					}
					Progress.SetTitle(Math.floor(100 * i / ar.length) + "%");
					Progress.SetProgress(i, ar.length);
					var data = ar[i].split("\t");
					var Id = data[data.length - 1];
					Progress.SetLine(1, Id, true);
					AddAddon(sorted, Id, document.getElementById("enable_" + Id).checked, "Sorted_");
				}
			} catch (e) {
				ShowError(e);
			}
			Progress.StopProgressDialog();
			table.style.display = Progress.HasUserCancelled() ? "" : "none";
			sorted.style.display = Progress.HasUserCancelled() ? "none" : "";
		}
	} else if (g_SortMode == "1_1") {
		if (n != g_nSort2) {
			g_nSort2 = n;
			AddonsSearch();
		}
	}
	ShowButtons(true);
}