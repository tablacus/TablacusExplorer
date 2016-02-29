//Tablacus Explorer

nTabMax = 0;
TabIndex = -1;
g_x = {Menu: null, Addons: null};
g_Chg = {Menus: false, Addons: false, Tab: false, Tree: false, View: false, Data: null};
g_arMenuTypes = ["Default", "Context", "Background", "Tabs", "Tree", "File", "Edit", "View", "Favorites", "Tools", "Help", "Systray", "System", "Alias"];
g_MenuType = "";
g_dlgAddons = null;
g_tdDown = null;
g_bDrag = false;
g_pt = {x: 0, y: 0};
g_Gesture = null;
g_tid = null;
g_tidResize = null;
g_drag5 = false;
g_nResult = 0;
g_bChanged = true;
g_bClosed = false;
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

urlAddons = "https://www.eonet.ne.jp/~gakana/tablacus/addons/";
nCount = 0;
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
		if (confirmOk("Are you sure?")) {
			SetValue(o, v);
		}
	}, 99);
}

function OpenGroup(id)
{
	var o = document.getElementById(id);
	o.style.display = String(o.style.display).toLowerCase() == "block" ? "none" : "block";
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
		document.F.View_SizeFormat.value = FV.SizeFormat > 9 ? api.sprintf(99, "%#x", FV.SizeFormat): FV.SizeFormat;
	}

	for(i = 0; i < document.F.length; i++) {
		o = document.F.elements[i];
		if (String(o.type).toLowerCase() == 'checkbox') {
			if (!/^Conf_/.test(o.id)) {
				o.checked = false;
			}
		}
	}
	for(i = 0; i < document.F.length; i++) {
		o = document.F.elements[i];
		if (/=/.test(o.id)) {
			var ar = o.id.split("=");
			if (document.F.elements[ar[0]].value == eval(ar[1])) {
				document.F.elements[i].checked = true;
			}
		}
		if (/:/.test(o.id)) {
			var ar = o.id.split(":");
			if (document.F.elements[ar[0]].value & eval(ar[1])) {
				document.F.elements[i].checked = true;
			}
		}
	}
	var s = GetWebColor(document.F.Conf_TrailColor.value);
	document.F.Conf_TrailColor.value = s;
	document.F.Color_Conf_TrailColor.style.backgroundColor = s;
	
	document.getElementById("_EDIT").checked = true;
	document.getElementById("_TEInfo").value = GetTEInfo();
}

function ResizeTabPanel()
{
	if (g_ovPanel) {
		var h = document.documentElement.clientHeight || document.body.clientHeight;
		h += MainWindow.DefaultFont.lfHeight * 5;
		if (h > 0) {
			g_ovPanel.style.height = h + 'px';
			g_ovPanel.style.height = 2 * h - g_ovPanel.offsetHeight + "px";
		}
	}
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
			document.getElementById("MoveButton").style.display = newTab == "1" || res[1] == 2 ? "inline-block" : "none";
			if (nMode == 0) {
				switch (res[1] - 0) {
					case 1:
						document.body.style.cursor = "wait";
						setTimeout(function () {
							LoadAddons();
							document.body.style.cursor = "auto";
						}, 9);
						if (newTab == '1_1') {
						setTimeout(function () {
							GetAddons();
							document.body.style.cursor = "auto";
						}, 9);
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
			var ovPanel = document.getElementById('panel' + TabIndex) || document.getElementById('panel' + TabIndex.replace(/_\d+/, ""));
			ovTab.className = 'button';
			ovPanel.style.display = 'none';
		}
		TabIndex = newTab;
		ovTab = document.getElementById('tab' + TabIndex);
		if (ovTab) {
			ovPanel = document.getElementById('panel' + TabIndex) || document.getElementById('panel' + TabIndex.replace(/_\d+/, ""));
			var res = /2_(.+)/.exec(TabIndex);
			if (res) {
				document.F.Menus.selectedIndex = res[1];
				SwitchMenus(document.F.Menus);
			}
			ovTab.className = 'hoverbutton';
			ovPanel.style.display = 'block';
			if (!dialogArguments.width) {
				var w = document.documentElement.clientWidth || document.body.clientWidth;
				ovPanel.style.width = document.documentMode >= 9 ? 'calc(100vw - 12em)' : (w - 12 * Math.abs(MainWindow.DefaultFont.lfHeight)) + "px";
			}
			var h = document.documentElement.clientHeight || document.body.clientHeight;
			h += MainWindow.DefaultFont.lfHeight * 2.7;
			if (h > 0) {
				ovPanel.style.height = h + 'px';
				ovPanel.style.height = 2 * h - ovPanel.offsetHeight + "px";
			}
			var o = document.getElementById("tab_");
			o.style.height = h + 'px';
			o.style.height = 2 * h - o.offsetHeight + "px";
		}
	}
}

function ClickButton(o, n, f)
{
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
	document.F.elements[ar[0]].value = ar[1];
}

function SetCheckbox(o)
{
	var ar = o.id.split(":");
	if (o.checked) {
		document.F.elements[ar[0]].value |= eval(ar[1]);
	} else {
		document.F.elements[ar[0]].value &= ~eval(ar[1]);
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
	var o = document.F.elements[name];
	i += api.LowPart(o.value);
	i = (i < min) ? min : i;
	o.value = (i < max) ? i : max;
}

function ChooseColor1(o)
{
	setTimeout(function ()
	{
		var o2 = document.F.elements[o.id.replace("Color_", "")];
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
		var o2 = document.F.elements[o.id.replace("Color_", "")];
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

function SetTreeControls()
{
	if (g_Chg.Tree) {
		if (document.getElementById("Conf_TreeDefault").checked) {
			var cTV = te.Ctrls(CTRL_TV);
			for (i = 0; i < cTV.Count; i++) {
				SetTreeControl(cTV.Item(i));
			}
		} else {
			var TV = te.Ctrl(CTRL_TV);
			if (TV) {
				SetTreeControl(TV);
			}
		}
	}
}

function SetTreeControl(TV)
{
	if (TV) {
		var Selected = TV.SelectedItem;
		TV.Align = document.F.Tree_Align.value;
		TV.Style = document.F.Tree_Style.value;
		TV.Width = document.F.Tree_Width.value;
		TV.SetRoot(document.F.Tree_Root.value, document.F.Tree_EnumFlags.value, document.F.Tree_RootStyle.value);
		TV.Expand(Selected, 1);
	}
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
	TC.Selected.Navigate2("c:\\", SBSP_NEWBROWSER, document.F.View_Type.value, document.F.View_ViewMode.value, document.F.View_fFlags.value, 0, document.F.View_Options.value, document.F.View_ViewFlags.value, Number(document.F.View_SizeFormat.value));
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

function SetFolderViews()
{
	FV = te.Ctrl(CTRL_FV);
	if (g_Chg.View) {
		if (document.getElementById("Conf_ListDefault").checked) {
			var cFV = te.Ctrls(CTRL_FV);
			for (i = 0; i< cFV.Count; i++) {
				SetFolderView(cFV.Item(i));
			}
		} else if (FV) {
			SetFolderView(FV);
		}
	}
	if (FV) {
		FV.CurrentViewMode = document.F.View_ViewMode.value;
	}
}

function SetFolderView(FV)
{
	if (FV) {
		FV.FolderFlags = document.F.View_fFlags.value;
		FV.Options = document.F.View_Options.value;
		FV.ViewFlags = document.F.View_ViewFlags.value;
		FV.SizeFormat = Number(document.F.View_SizeFormat.value);
		if (FV.Type != document.F.View_Type.value) {
			FV.Type = document.F.View_Type.value;
		} else {
			FV.Refresh();
		}
	}
}

function InitConfig(o)
{
	var InstallPath = fso.GetParentFolderName(api.GetModuleFileName(null));
	if (InstallPath == te.Data.DataFolder) {
		return;
	}
	if (!confirmOk("Are you sure?")) {
		return;
	}
	var Dist = sha.NameSpace(te.Data.DataFolder);
	Dist.MoveHere(fso.BuildPath(InstallPath, "layout"), 0);
	o.disabled = true;
}

function SelectPos(o, s)
{
	var v = o[o.selectedIndex].value;
	if (v != "") {
		document.F.elements[s].value = v;
	}
}

function SwitchMenus(o)
{
	if (g_x.Menus) {
		g_x.Menus.style.display = "none";
		var o = document.F.elements.Menus;
		for (var i = o.length; i-- > 0;) {
			var a = o[i].value.split(",");
			if ("Menus_" + a[0] == g_x.Menus.name) {
				s = a[0] + "," + document.F.elements["Menus_Base"].selectedIndex + "," + document.F.elements["Menus_Pos"].value;
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
		g_x.Menus = document.F.elements["Menus_" + a[0]];
		g_x.Menus.style.display = "inline";
		document.F.elements["Menus_Base"].selectedIndex = a[1];
		document.F.elements["Menus_Pos"].value = api.LowPart(a[2]);
		CancelX("Menus");
	}
}

function SwitchX(mode, o)
{
	g_x[mode].style.display = "none";
	g_x[mode] = document.F.elements[mode + o.value];
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

function SetOptions(fnYes, fnNo, NoCancel, bNoDef)
{
	if (g_nResult == 2 || api.StrCmpI(document.activeElement.value, GetText("Cancel")) == 0) {
		if (fnNo) {
			fnNo();
		}
		return false;
	}
	document.activeElement.blur();
	if (g_nResult == 1 && !g_Chg.Data) {
		if (fnYes) {
			fnYes();
		}
		return true;
	}
 	if (g_bChanged || g_Chg.Data) {
		switch (MessageBox("Do you want to replace?", TITLE, MB_ICONQUESTION | ((NoCancel || dialogArguments.width) && NoCancel !== 2 ? MB_YESNO : MB_YESNOCANCEL))) {
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
				}
				return true;
			default:
				if (g_nResult && NoCancel !== 2) {
					if (event.preventDefault) {
						event.preventDefault();
					} else {
			 			event.returnValue = GetText('Are you sure?');
						wsh.SendKeys("{Esc}");
					}
				}
				return false;
		}
		if (!bNoDef && fnYes) {
			fnYes();
		}
		return true;
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

function EditXEx(o, s)
{
	if (document.getElementById("_EDIT").checked) {
		o(s);
	}
}

EditX = function (mode)
{
	if (g_x[mode].selectedIndex < 0) {
		return;
	}
	ClearX(mode);
	var a = g_x[mode][g_x[mode].selectedIndex].value.split(g_sep);
	document.F.elements[mode + mode].value = a[0];
	var p = { s: a[1] };
	MainWindow.OptionDecode(a[2], p);
	document.F.elements[mode + "Path"].value = p.s;
	SetType(document.F.elements[mode + "Type"], a[2]);
	if (api.StrCmpI(mode, "Key") == 0) {
	 	SetKeyShift();
	}
}

function SetType(o, value)
{
	var i = o.length;
	while (--i >= 0) {
		if (o[i].value == value) {
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

function AddX(mode, fn)
{
	InsertX(g_x[mode]);
	(fn || ReplaceX)(mode);
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
	var s = GetSourceText(document.F.Menus_Name.value);
	var org = (s == document.F.Menus_Name.value && api.GetKeyState(VK_SHIFT) >= 0) ? "1" : ""
	if (document.F.Menus_Key.value.length) {
		var n = GetKeyKey(document.F.Menus_Key.value);
		s += "\\t" + (n ? api.sprintf(8, "$%x", n) : document.F.Menus_Key.value);
	}
	var p = { s: document.F.Menus_Path.value };
	MainWindow.OptionEncode(o[o.selectedIndex].value, p);
	SetMenus(sel, [s, document.F.Menus_Filter.value, p.s, o[o.selectedIndex].value, document.F.Icon.value, org]);
	g_Chg.Menus = true;
}

function ReplaceX(mode)
{
	ClearX(mode);
	if (g_x[mode].selectedIndex < 0) {
		InsertX(g_x[mode]);
		EnableSelectTag(g_x[mode]);
	}
	var sel = g_x[mode][g_x[mode].selectedIndex];
	var o = document.F.elements[mode + "Type"];
	var p = { s: document.F.elements[mode + "Path"].value };
	MainWindow.OptionEncode(o[o.selectedIndex].value, p);
	SetData(sel, [document.F.elements[mode + mode].value, p.s, o[o.selectedIndex].value]);
	g_Chg[mode] = true;
	g_bChanged = true;
}

function RemoveMenus()
{
	ClearX("Menus");
	if (g_x.Menus.selectedIndex < 0 || !confirmOk("Are you sure?")) {
		return;
	}
	g_x.Menus[g_x.Menus.selectedIndex] = null;
	g_Chg.Menus = true;
	g_bChanged = true;
}

function RemoveX(mode)
{
	ClearX(mode);
	if (g_x[mode].selectedIndex < 0 || !confirmOk("Are you sure?")) {
		return;
	}
	g_x[mode][g_x[mode].selectedIndex] = null;
	g_Chg[mode] = true;
	g_bChanged = true;
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
		for (var i in MainWindow.eventTE.AddType) {
			MainWindow.eventTE.AddType[i](arFunc);
		}
		for (var i = 0; i < arFunc.length; i++) {
			var o = oa[++oa.length - 1];
			o.value = arFunc[i];
			o.innerText = GetText(arFunc[i]).replace(/&|\.\.\.$/g, "").replace(/\(\w\)/, "");
		}

		oa = document.F.Menus;
		oa.length = 0;

		for (var j in g_arMenuTypes) {
			var s = g_arMenuTypes[j];
			document.getElementById("Menus_List").insertAdjacentHTML("BeforeEnd", ['<select name="Menus_', s, '" size="17" style="width: 12em; height: 32em; height: calc(100vh - 6em); display: none; font-family:', document.F.elements["Menus_Pos"].style.fontFamily, '" onclick="EditXEx(EditMenus)" ondblclick="EditMenus()" oncontextmenu="CancelX(\'Menus\')"></select>'].join(""));
			var menus = teMenuGetElementsByTagName(s);
			if (menus && menus.length) {
				oa[++oa.length - 1].value = s + "," + menus[0].getAttribute("Base") + "," + menus[0].getAttribute("Pos");
				var o = document.F.elements["Menus_" + s];
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
					FireEvent(o, "click");
				}
			}, 99);}) (document.F.elements["Menus_" + ar[0]], ar[1]);
		}
	}
}

function LoadX(mode, fn)
{
	if (!g_x[mode]) {
		var arFunc = [];
		for (var i in MainWindow.eventTE.AddType) {
			MainWindow.eventTE.AddType[i](arFunc);
		}
		var oa = document.F.elements[mode + "Type"] || document.F.Type;
		while (oa.length) {
			oa.removeChild(oa[0]);
		}
		for (var i = 0; i < arFunc.length; i++) {
			var o = oa[++oa.length - 1];
			o.value = arFunc[i];
			o.innerText = GetText(arFunc[i]).replace(/&|\.\.\.$/g, "").replace(/\(\w\)/, "");
		}
		g_x[mode] = document.F.elements[mode + "All"];
		if (g_x[mode]) {
			oa = document.F.elements[mode];
			oa.length = 0;
			xml = OpenXml(mode + ".xml", false, true);
			for (var j in g_Types[mode]) {
				var o = oa[++oa.length - 1];
				o.text = GetTextEx(g_Types[mode][j]);
				o.value = g_Types[mode][j];
				o = document.F.elements[mode + g_Types[mode][j]];
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
					SetData(o[i], [s, item.text, item.getAttribute("Type")]);
				}
			}
		} else {
			g_x[mode] = document.F.List;
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
			var o = document.F.elements["Menus_" + g_arMenuTypes[j]];
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

function SaveX(mode)
{
	if (g_Chg[mode]) {
		var xml = CreateXml();
		var root = xml.createElement("TablacusExplorer");
		for (var j in g_Types[mode]) {
			var o = document.F.elements[mode + g_Types[mode][j]];
			for (var i = 0; i < o.length; i++) {
				var item = xml.createElement(g_Types[mode][j]);
				var a = o[i].value.split(g_sep);
				var s = a[0];
				if (api.StrCmpI(mode, "Key") == 0) {
					var ar = /,$/.test(s) ? [s] : s.split(",");
					for (var k = ar.length; k--;) {
						var n = GetKeyKey(ar[k]);
						if (n) {
							ar[k] = api.sprintf(8, "$%x", n);
						}
					}
					s = ar.join(",");
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
				var Enabled = api.StrCmpI(div.style.color, "gray") ? 1 : 0;
				if (Enabled) {
					var AddonFolder = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\" + Id);
					Enabled = 0;
					if (fso.FolderExists(AddonFolder + "\\lang")) {
						Enabled = 2;
					}
					if (fso.FileExists(AddonFolder + "\\script.vbs")) {
						Enabled |= 8;
					}
					if (fso.FileExists(AddonFolder + "\\script.js")) {
						Enabled |= 1;
					}
					Enabled = (Enabled & 9) ? Enabled : 4;
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

function SetChanged(fn)
{
	if (g_Chg.Data) {
		g_bChanged = true;
		if (g_x[g_Chg.Data] && g_x[g_Chg.Data].selectedIndex >= 0) {
			(fn || ReplaceX)(g_Chg.Data);
		} else {
			AddX(g_Chg.Data, fn);
		}
	}
}

function SetData(sel, a)
{
	sel.value = PackData(a);
	sel.text = GetText(a[0]);
}

function PackData(a)
{
	var i = a.length;
	while (--i >= 0) {
		a[i] = (a[i] || "").replace(g_sep, "`  ~");
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
					AddAddon(table, Id, (api.LowPart(item.getAttribute("Enabled"))) ? "Disable" : "Enable");
					delete AddonId[Id];
				}
			}
		}
	}
	for (var Id in AddonId) {
		if (fso.FileExists(path + Id + "\\config.xml")) {
			AddAddon(table, Id, "Enable");
		}
	}
}

function AddAddon(table, Id, Enable)
{
	var tr = table.insertRow();
	tr.className = (tr.rowIndex & 1) ? "oddline" : "";
	SetAddon(tr.insertCell(), Id, Enable);
}

function SetAddon(td, Id, Enable)
{
	var info = GetAddonInfo(Id);

	if (info.Description && info.Description.length > 80) {
		info.Description = info.Description.substr(0, 80) + "...";
	}
	var s = ['<div draggable="true" ondragstart="Start5(this)" ondragend="End5(this)" Id="Addons_' + Id + '" style="color: '];
	s.push((Enable == "Enable") ? "gray" : "black");
	s.push('"><input type="radio" name="AddonId" id="_', Id, '"><label for="_', Id, '"><b>', info.Name, "</b>&nbsp;", info.Version, "&nbsp;", info.Creator, "<br />", info.Description, "<br />");
	s.push('<table><tr><td><input type="button" value="', GetText('Remove'), '" onclick="AddonRemove(\'', Id, '\')"></td>');
	s.push('<td><input type="button" value="', GetText(Enable), '" onclick="AddonEnable(\'', Id, '\', this)"');
	if (info.MinVersion && te.Version < CalcVersion(info.MinVersion)) {
		s.push(" disabled");
	}
	s.push('></td>');
	s.push('<td><input type="button" value="', GetText('Info...'), '" onclick="AddonInfo(\'', Id, '\')"></td>');
	s.push('<td><input type="button" value="', GetText('Options...'), '" onclick="AddonOptions(\'', Id, '\')"');
	if (!info.Options) {
		s.push(" disabled");
	}
	s.push('></td></tr></table></label></div>');
	td.innerHTML = s.join("");

	td.onmousedown = function (e)
	{
		var o = document.elementFromPoint((e || window.event).clientX, (e || window.event).clientY);
		if (!o || api.StrCmpI(o.tagName, "input")) {
			g_tdDown = (e ? e.currentTarget : window.event.srcElement).firstChild.id;
		}
	}

	td.onmouseup = function (e)
	{
		if (g_bDrag) {
			g_bDrag = false;
			SetCursor(document.getElementById("Addons"), "auto");
			if (g_tdDown) {
				(function (src, dest) { setTimeout(function () {
					AddonMoveEx(src, dest);
				}, 99);}) (GetRowIndexById(g_tdDown) , GetRowIndexById((e ? e.currentTarget : window.event.srcElement).firstChild.id));
			}
		}
		g_tdDown = null;
	}

	td.onmousemove = function (e)
	{
		if (g_tdDown) {
			if (api.GetKeyState(VK_LBUTTON) < 0 && !g_drag5) {
				(e ? e.currentTarget : window.event.srcElement).style.cursor = "move";
				g_bDrag = true;
			} else {
				td.onmouseup(e);
			}
		}
	}

	ApplyLang(td);
}

function Start5(o)
{
	if (api.GetKeyState(VK_LBUTTON) < 0) {
		event.dataTransfer.effectAllowed = 'move';
		g_drag5 = o.id;
		if (g_tdDown) {
			SetCursor(document.getElementById("Addons"), "auto");
			g_tdDown = null;
		}
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
	} catch (e) {
	}
}

function AddonInfo(Id)
{
	var info = GetAddonInfo(Id);
	var pubDate = "";
	if (info.pubDate) {
		pubDate = new Date(info.pubDate).toLocaleString() + "\n";
	}
	MessageBox(info.Name + " " + info.Version + " " + info.Creator + "\n\n" + info.Description + "\n\n" + pubDate + info.URL, Id, MB_ICONINFORMATION);
}

function AddonWebsite(Id)
{
	var info = GetAddonInfo(Id);
	wsh.run(info.URL);
}

function AddonEnable(Id, o)
{
	var div = document.getElementById("Addons_" + Id);
	if (o.value != GetText('Enable')) {
		MainWindow.AddonDisabled(Id);
		o.value = GetText('Enable');
		div.style.color = "gray";
	} else {
		var info = GetAddonInfo(Id);
		if (!info.MinVersion || te.Version >= api.LowPart(info.MinVersion.replace(/\D/g, ""))) {
			o.value = GetText('Disable');
			div.style.color = "black";
		}
	}
	g_Chg.Addons = true;
}

function OptionMove(dir)
{
	if (/^1/.test(TabIndex)) {
		var r = document.F.AddonId;
		for (i = 0; i < r.length; i++) {
			if (r[i].checked) {
				try {
					AddonMoveEx(i, i + dir);
					document.getElementById("Addons").rows(i + dir).scrollIntoView(false);
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
	var id2 = tr.id;
	var td = tr.cells(0);

	var s = td.innerHTML
	var md = td.onmousedown;
	var mu = td.onmouseup;
	var mm = td.onmousemove;

	table.deleteRow(src);

	tr = table.insertRow(dest);
	td = tr.insertCell();
	tr.id = id2;
	td.innerHTML = s;
	td.onmousedown = md;
	td.onmouseup = mu;
	td.onmousemove = mm;

	var i = src > dest ? src : dest;
	var j = src > dest ? dest : src;
	while (i >= j) {
		table.rows(i).className = (i & 1) ? "oddline" : "";
		i--;
	}
	document.F.AddonId[dest].checked = true;
	g_Chg.Addons = true;
	return false;
}

function AddonRemove(Id)
{
	if (!confirmOk("Are you sure?")) {
		return;
	}
	MainWindow.AddonDisabled(Id);
	sf = api.Memory("SHFILEOPSTRUCT");
	sf.hwnd = api.GetForegroundWindow();
	sf.wFunc = FO_DELETE;
	sf.fFlags = 0;
	sf.pFrom = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\"+ Id) + "\0";
	if (api.SHFileOperation(sf) == 0) {
		if (!sf.fAnyOperationsAborted) {
			var table = document.getElementById("Addons");
			var i = GetRowIndexById("Addons_" + Id);
			table.deleteRow(i);
			while (i < table.rows.length) {
				table.rows(i).className = (i & 1) ? "oddline" : "";
				i++;
			}
			g_Chg.Addons = true;
		}
	}
}

InitOptions = function ()
{
	ApplyLang(document);

	var InstallPath = fso.GetParentFolderName(api.GetModuleFileName(null));
	document.F.ButtonInitConfig.disabled = (InstallPath == te.Data.DataFolder) | !fso.FolderExists(fso.BuildPath(InstallPath, "layout"));
	for (i in document.F.elements) {
		if (!/=|:/.test(i)) {
			if (/^Tab_|^Tree_|^View_|^Conf_/.test(i)) {
				if (te.Data[i] !== undefined) {
					SetElementValue(document.F.elements[i], te.Data[i]);
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
		g_bChanged |= g_Chg.Addons || te.Data.bErrorAddons || g_Chg.Menus || g_Chg.Tab || g_Chg.Tree || g_Chg.View;
		SaveAddons();
		SetOptions(function () {
			SetChanged(ReplaceMenus);
			for (var i in document.F.elements) {
				if (!/=|:/.test(i)) {
					if (/^Tab_|^Tree_|^View_|^Conf_/.test(i)) {
						te.Data[i] = GetElementValue(document.F.elements[i]);
					}
				}
			}
			te.Layout = te.Data.Conf_Layout;
			SaveMenus();
			SetTabControls();
			SetTreeControls();
			SetFolderViews();
			te.Data.bReload = true;
		});
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
			var image = te.GdiplusBitmap;
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
				var s = ["icon:" + a[1],i ,a[2]].join(",");
				var src = MakeImgSrc(s, 0, false, a[2]);
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

			"16px shell32" : "i,shell32.dll,16",
			"32px shell32" : "i,shell32.dll,32",
			"16px imageres" : "i,imageres.dll,16",
			"32px imageres" : "i,imageres.dll,32",
			"16px wmploc" : "i,wmploc.dll,16",
			"32px wmploc" : "i,wmploc.dll,32",
			"16px setupapi" : "i,setupapi.dll,16",
			"32px setupapi" : "i,setupapi.dll,32",
			"16px dsuiext" : "i,dsuiext.dll,16",
			"32px dsuiext" : "i,dsuiext.dll,32",
			"16px inetcpl" : "i,inetcpl.cpl,16",
			"32px inetcpl" : "i,inetcpl.cpl,32",
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
		document.getElementById("Content").innerHTML = '<div style="padding: 8px;" style="display: block;"><label>Key</label><br /><input type="text" name="q" autocomplete="off" style="width: 100%; ime-mode: disabled" /></div>';
		document.body.onkeydown = function (e)
		{
			var key = (e || event).keyCode;
			var s = api.sprintf(10, "$%x", (api.MapVirtualKey(key, 0) | ((key >= 33 && key <= 46 || key >= 91 && key <= 93 || key == 111 || key == 144) ? 256 : 0) | GetKeyShift()));
			returnValue = GetKeyName(s);
			document.F.q.value = returnValue;
			document.F.q.title = s;
			document.F.ButtonOk.disabled = false;
			return false;
		}
	}
	if (Query == "new") {
		returnValue = false;
		var s = [];
		s.push('<div style="padding: 8px;" style="display: block;"><input type="radio" name="mode" id="folder" onclick="document.F.path.focus()"><label for="folder">New Folder</label> <input type="radio" name="mode" id="file" onclick="document.F.path.focus()"><label for="file">New File</label><br />', dialogArguments.path ,'<br /><input type="text" name="path" style="width: 100%" /></div>');
		document.getElementById("Content").innerHTML = s.join("");
		document.body.onkeydown = function (e)
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
		}
		document.body.onpaste = function ()
		{
			setTimeout(function ()
			{
				document.F.ButtonOk.disabled = !document.F.path.value;
			}, 99);
		}
		setTimeout(function ()
		{
			document.F.elements[dialogArguments.Mode].checked = true;
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
					dialogArguments.FV.SelectItem(path, SVSI_SELECT | SVSI_DESELECTOTHERS | SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_NOTAKEFOCUS);
				}
			}
		});
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
			} catch (e) {
			}
		});
	}
	DialogResize = function ()
	{
		var h = document.documentElement.clientHeight || document.body.clientHeight;
		var i = document.getElementById("buttons").offsetHeight * screen.deviceYDPI / screen.logicalYDPI + 6;
		h -= i > 34 ? i : 34;
		if (h > 0) {
			document.getElementById("panel0").style.height = h + 'px';
		}
	};
	AddEventEx(window, "resize", function ()
	{
		clearTimeout(g_tidResize);
		g_tidResize = setTimeout(function ()
		{
			DialogResize();
		}, 500);
	});
	ApplyLang(document);
	DialogResize();
}

MouseDown = function ()
{
	if (g_Gesture) {
		var c = returnValue.charAt(returnValue.length - 1);
		var n = 1;
		for (i = 1; i < 4; i++) {
			if (event.button & n && g_Gesture.indexOf(i + "") < 0) {
				returnValue += i + "";
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
		ar.push('<input type="button" value="', MainWindow.g_KeyState[i][0],'" title="', s.charAt(i), '" onclick="AddMouse(this)" />');
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
			if (api.StrCmpI(Location, document.L.elements[i].value) == 0) {
				document.L.elements[i].checked = true;
			}
		}
	}
	var locs = [];
	items = te.Data.Locations;
	for (var i in items) {
		locs[i] = [];
		for (var j in items[i]) {
			info = GetAddonInfo(items[i][j]);
			locs[i].push(info.Name);
		}
	}
	for (var i in locs) {
		var s = locs[i].join(", ").replace('"', "");
		try {
			var o = document.getElementById('_' + i);
			ApplyLang(o);
			var s2 = o.innerHTML.replace(/<[^>]*>|[\r\n]|\s\s+/g, "");
			o.innerHTML = ['<input type="text" value="', s, '" title="', s2, '" placeholder="', s2, '" style="width: 85%" readonly="readonly" />'].join("");
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
		var oa = document.F.elements[mode + "On"];
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
			if (n) {
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
	document.F._MenuKey.value = GetKeyName(a[1]) || "";

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
		g_tidResize = setTimeout(function ()
		{
			ResizeTabPanel();
		}, 500);
	});
	
	AddEventEx(window, "beforeunload", function ()
	{
		SetOptions(function () {
			if (window.SaveLocation) {
				window.SaveLocation();
			}
			var items = te.Data.Addons.getElementsByTagName(dialogArguments.Data.id);
			if (items.length) {
				var item = items[0];
				item.removeAttribute("Location");
				for (var i = document.L.elements.length; i--;) {
					if (document.L.elements[i].checked) {
						item.setAttribute("Location", document.L.elements[i].value);
						te.Data.bReload = true;
						MainWindow.RunEvent1("ConfigChanged", "Addons");
						break;
					}
				}
				var ele = document.F.elements;
				var a = [GetSourceText(ele._MenuName.value)];
				if (ele._MenuKey.value) {
					var s = GetKeyKey(ele._MenuKey.value);
					if (s) {
						a.push(api.sprintf(10, "$%x", s));
					}
				}
				ele.MenuName.value = a.join("\t");
				if (dialogArguments.Data.show == "6") {
					ele.Set.value = "";
				}
				for (var i = ele.length; i--;) {
					var n = ele[i].id || ele[i].name;
					if (n && n.charAt(0) != "_") {
						if (n == "Key") {
							var s = GetKeyKey(document.F.elements[n].value);
							if (s) {
								document.F.elements[n].value = api.sprintf(10, "$%x", s);
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
		if (api.StrCmpI(o.type, 'checkbox') == 0) {
			return o.checked ? 1 : 0;
		}
		if (/hidden|text/i.test(o.type)) {
			return o.value;
		}
		if (/select/i.test(o.type)) {
			return o[o.selectedIndex].value;
		}
	}
}

function SetElementValue(o, s)
{
	if (o.type) {
		if (api.StrCmpI(o.type, "checkbox") == 0) {
			o.checked = api.LowPart(s);
			return;
		}
		if (/text/i.test(o.type)) {
			o.value = s;
			return;
		}
		if (/select/i.test(o.type)) {
			var i = o.length;
			while (--i >= 0) {
				if (o(i).value == s) {
					o.selectedIndex = i;
					break;
				}
			}
		}
	}
}

function SetAttribEx(item, f, n)
{
	var s = GetElementValue(f.elements[n]);
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
		SetElementValue(f.elements[n], s);
	}
}

function RefX(Id, bMultiLine, oButton, bFilesOnly)
{
	setTimeout(function () {
		if (/Path/.test(Id)) {
			var s = Id.replace("Path", "Type");
			var o = GetElement(s);
			if (o) {
				var pt;
				if (oButton) {
					pt = GetPos(oButton, true);
					pt.y = pt.y + o.offsetHeight * screen.deviceYDPI / screen.logicalYDPI;
				} else {
					pt = api.Memory("POINT");
					api.GetCursorPos(pt);
				}
				var r = MainWindow.OptionRef(o[o.selectedIndex].value, GetElement(Id).value, pt);
				if (typeof r == "string") {
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
		var path = OpenDialog(o.value, api.LowPart(bFilesOnly));
		if (path) {
			if (bMultiLine) {
				AddPath(Id, path);
			} else {
				SetValue(o, path);
			}
		}
	}, 99);
	g_Chg.Data = true;
}

function PortableX(Id)
{
	if (!confirmOk("Are you sure?")) {
		return;
	}
	var o = GetElement(Id);
	var s = fso.GetDriveName(api.GetModuleFileName(null));
	SetValue(o, o.value.replace(new RegExp('^("?)' + s, "igm"), "$1%Installed%").replace(new RegExp('( "?)' + s, "igm"), "$1%Installed%"));
}

function GetElement(Id)
{
	var o = document.F.elements[Id];
	return o ? o : document.getElementById(Id);
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

	if (confirmOk("Are you sure?")) {
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
	(document.F.elements["MouseMouse"] || document.F.elements["Mouse"]).value += o.title;
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
				if (o[i].name && !/^_/.test(o[i].id)) {
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
}

TestX = function (id)
{
	if (confirmOk("Are you sure?")) {
		var o = document.F.elements[id + "Type"];
		var p = { s: document.F.elements[id + "Path"].value };
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
		api.InsertMenu(hMenu, i, MF_BYPOSITION | MF_STRING, api.QuadPart(i) + 1, title);
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
	if (nCount) {
		return;
	}
	xhr = createHttpRequest();
	xhr.onreadystatechange = function()
	{
		if (xhr.readyState == 4) {
			if (xhr.status == 200) {
				setTimeout(AddonsList, 99);
			}
		}
	}
	xhr.open("GET", urlAddons + "/index.xml?" + Math.floor(new Date().getTime() / 60000));
	xhr.setRequestHeader('Content-Type', 'text/xml');
	xhr.setRequestHeader('Pragma', 'no-cache');
	xhr.setRequestHeader('Cache-Control', 'no-cache');
	xhr.setRequestHeader('If-Modified-Since', 'Thu, 01 Jun 1970 00:00:00 GMT');
	try {
		xhr.send(null);
	} catch (e) {
		ShowError(e);
	}
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
	clearTimeout(g_tid);
	if (nCount != xmlAddons.length) {
		AddonsList();
		return true;
	}
	nCount = 0;
	var q =
	{
		td: [],
		ts: [],
		i: 0
	}
	if (AddonsSub(q)) {
		document.body.style.cursor = "wait";
		AddonsAppend(q)
	}
	return true;
}

function AddonsSub(q)
{
	if (document.getElementById) {
		var table = document.getElementById("Addons1");
		if (table) {
			while (table.rows.length > 0) {
				table.deleteRow(0);
			}
			for (var i = 0; i < xmlAddons.length; i++) {
				var tr = table.insertRow(i);
				q.td[i] = tr.insertCell(0);
			}
			q.ts = new Array(xmlAddons.length);
			return true;
		}
	}
	return false;
}

function AddonsList()
{
	clearTimeout(g_tid);
	nCount = 0;
	var q =
	{
		td: [],
		ts: [],
		i: 0
	}
	var xml = xhr.responseXML;
	if (xml) {
		xmlAddons = xml.getElementsByTagName("Item");
		if (AddonsSub(q)) {
			document.body.style.cursor = "wait";
			AddonsAppend(q)
		}
	}
}

function AddonsAppend(q)
{
	try {
		if (xmlAddons[q.i]) {
			ArrangeAddon(xmlAddons[q.i++], q.td, q.ts);
			g_tid = setTimeout(function () {
				AddonsAppend(q);
			}, 1);
			document.getElementById('STATUS').innerText = Math.floor(100 * q.i / q.ts.length) + " %";
		} else {
			document.body.style.cursor = "auto";
			document.getElementById('STATUS').innerText = "";
		}
	} catch (e) {}
}

function ArrangeAddon(xml, td, ts)
{
	var Id = xml.getAttribute("Id");
	var s = [];
	if (Search(xml)) {
		var info = [];
		for (var i = arLangs.length; i--;) {
			GetAddonInfo2(xml, info, arLangs[i]);
		}
		var pubDate = "";
		var dt = new Date(info.pubDate);
		if (info.pubDate) {
			pubDate = api.GetDateFormat(LOCALE_USER_DEFAULT, 0, dt, api.GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE)) + " ";
		}
		s.push('<b>', info.Name, "</b>&nbsp;", info.Version, "&nbsp;", info.Creator, "<br>", info.Description, "<br>");

		s.push('<table width="100%"><tr><td>', pubDate, '</td><td align="right">');
		var filename = info.filename;
		if (!filename) {
			filename = Id + '_' + info.Version.replace(/\D/, '') + '.zip';
		}
		var dt2 = (dt.getTime() / (24 * 60 * 60 * 1000)) - info.Version;
		var bUpdate = false;
		if (CheckAddon(Id)) {
			installed = GetAddonInfo(Id);
			if (installed.Version >= info.Version) {
				return;
			} else {
				s.push('<b id="_Addons_', Id,'" style="color: red; white-space: nowrap;">', GetText('Update available'), "</b>");
				dt2 += MAXINT * 2;
				bUpdate = true;
			}
		} else {
			dt2 += MAXINT;
		}
		if (info.MinVersion && te.Version >= CalcVersion(info.MinVersion)) {
			s.push('<input type="button" onclick="Install(this,', bUpdate, ')" title="', Id, '_', info.Version, '" value="', GetText("Install"), '">');
		} else {
			s.push('<input type="button" style="color: red" onclick="CheckUpdate()" value="', info.MinVersion.replace(/^20/, "Version ").replace(/\.0/g, '.'), ' ', GetText("is required."), '">');
		}
		s.push('</td></tr></table>');
		var nInsert = 0;
		while (nInsert <= nCount && dt2 < ts[nInsert]) {
			nInsert++;
		}
		for (j = nCount; j > nInsert; j--) {
			td[j].innerHTML = td[j - 1].innerHTML;
			td[j].className = (j & 1) ? "oddline" : "";
			ts[j] = ts[j - 1];
		}
		td[nInsert].className = (nInsert & 1) ? "oddline" : "";
		td[nInsert].innerHTML = s.join("");
		ApplyLang(td[nInsert]);
		ts[nInsert] = dt2;
		nCount++;
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
					if (item1.text.toUpperCase().match(q)) {
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
	var Id = o.title.replace(/_.*/, "");

	MainWindow.AddonDisabled(Id);
	document.body.style.cursor = "wait";
	setTimeout(function ()
	{
		var Id = o.title.replace(/_.*$/, "");
		var file = o.title.replace(/\./, "") + '.zip';
		var temp = fso.BuildPath(wsh.ExpandEnvironmentStrings("%TEMP%"), "tablacus");
		DeleteItem(temp);
		CreateFolder(temp);
		var zipfile = fso.BuildPath(temp, file);
		if (DownloadFile(urlAddons + Id + '/' + file, zipfile) != S_OK || MainWindow.Extract(zipfile, temp) != S_OK) {
			document.body.style.cursor = "auto";
			return;
		}

		var configxml = fso.BuildPath(temp, Id) + "\\config.xml";
		var nDog = 300;
		while (!fso.FileExists(configxml)) {
			if (wsh.Popup(GetText("Please wait."), 1, TITLE, MB_ICONINFORMATION | MB_OKCANCEL) == IDCANCEL || nDog-- == 0) {
				document.body.style.cursor = "auto";
				return;
			}
		}
		var oSrc = sha.NameSpace(fso.BuildPath(temp, Id));
		if (oSrc) {
			var Items = oSrc.Items();
			var dest = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\" + Id);
			CreateFolder(dest);
			var oDest = sha.NameSpace(dest);
			if (oDest) {
				oDest.MoveHere(Items, FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR);
				o.disabled = true;
				o.value = GetText("Installed");
				o = document.getElementById('_Addons_' + Id);
				if (o) {
					o.style.display = "none";
				}
				UpdateAddon(Id, o);
			}
		}
		document.body.style.cursor = "auto";
	}, 99);
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
				var o = document.F.elements[res[1]];
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
	AddEventEx(window, "load", function ()
	{
		document.getElementById("__Inner").disabled  = false;;
	});
}

function SetTabContents(id, name, value)
{
	document.getElementById("tab" + id).value = GetText(name);
	document.getElementById("panel" + id).innerHTML = value;
}
