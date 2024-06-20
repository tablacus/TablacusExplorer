//Tablacus Explorer

RunEventUI("BrowserCreatedEx");

nTabMax = 0;
TabIndex = -1;
g_x = { Menu: null, Addons: null };
g_Chg = { Menus: false, Addons: false, Tab: false, Tree: false, View: false, Data: null };
g_arMenuTypes = ["Default", "Context", "Background", "Tabs", "Tree", "File", "Edit", "View", "Favorites", "Tools", "Help", "Systray", "System", "Alias"];
g_MenuType = "";
g_Id = "";
g_dlgAddons = null;
g_bDrag = false;
g_pt = { x: 0, y: 0 };
g_Gesture = null;
g_drag5 = false;
g_nResult = 0;
g_bChanged = true;
g_bClosed = false;
g_nSort = {
	"1_1" : 1,
	"1_3" : 1
};
g_ovPanel = null;
g_.ShowError = true;
ui_.elAddons = {};

urlAddons = "https://tablacus.github.io/TablacusExplorerAddons/";
urlIcons = urlAddons + "te/iconpacks/";
urlLang = urlAddons + "te/lang/";
xhr = null;
xmlAddons = null;
arLangs = ["General"];

Promise.all([GetLangId(), MainWindow.IconSize]).then(function (r) {
	ui_.IconSize = r[1];
	if (!/^en/.test(r[0])) {
		arLangs.unshift("en");
	}
	const res = /(\w+)_/.exec(r[0]);
	if (res && !/zh_cn/i.test(r[0])) {
		arLangs.unshift(res[1]);
	}
	arLangs.unshift(r[0]);
});

CloseWB = async function (WB, bForce) {
	if (bForce || g_nResult != 4) {
		FireEvent(window, "unload");
		WB.Close();
	}
	g_nResult = 0;
}

async function SetDefaultLangID(el) {
	SetDefault(document.F.Conf_Lang, await GetLangId(true), el);
}

function SetDefault(o, v) {
	ConfirmThenExec(GetText("Set default"), async function () {
		SetValue(o, "function" === typeof v ? await v : v);
	});
}

function OpenGroup(id) {
	const o = document.getElementById(id);
	o.style.display = /block/i.test(o.style.display) ? "none" : "block";
}

function LoadChecked(form) {
	for (let i = 0; i < form.length; ++i) {
		const o = form[i];
		let ar = o.id.split("=");
		if (ar.length > 1 && form[ar[0]].value == Number(ar[1])) {
			form[i].checked = true;
		}
		ar = o.id.split(/:!?/);
		if (ar.length > 1) {
			if ((form[ar[0]].value & Number(ar[1]) ? 1 : 0) ^ (/!/.test(o.id) ? 1 : 0)) {
				form[i].checked = true;
			}
		}
	}
}

async function ResetForm() {
	const Ctrl = await Promise.all([te.Ctrl(CTRL_FV), te.Ctrl(CTRL_TC), te.Ctrl(CTRL_TV), te.Data.Conf_SizeFormat]);
	const TV = Ctrl[2];
	if (TV) {
		const r = await Promise.all([TV.Align, TV.Style, TV.EnumFlags, TV.RootStyle, TV.Align, TV.Root, TV.Width]);
		document.F.Tree_Align.value = r[0];
		document.F.Tree_Style.value = r[1];
		document.F.Tree_EnumFlags.value = r[2];
		document.F.Tree_RootStyle.value = r[3];
		if (r[4] & 2) {
			document.F.Tree_Root.value = r[5];
		}
		document.F.Tree_Width.value = r[6];
	}
	const TC = Ctrl[1];
	if (TC) {
		const r = await Promise.all([TC.Style, TC.Align, TC.TabWidth, TC.TabHeight, TC.Left, TC.Top, TC.Width, TC.Height]);
		document.F.Tab_Style.value = r[0];
		document.F.Tab_Align.value = r[1];
		document.F.Tab_TabWidth.value = r[2];
		document.F.Tab_TabHeight.value = r[3];

		document.F.Tab_Left.value = r[4];
		document.F.Tab_Top.value = r[5];
		document.F.Tab_Width.value = r[6];
		document.F.Tab_Height.value = r[7];
	}
	const FV = Ctrl[0];
	if (FV) {
		const r = await Promise.all([FV.Type, FV.CurrentViewMode, FV.FolderFlags, FV.Options, FV.ViewFlags]);
		document.F.View_Type.value = r[0];
		document.F.View_ViewMode.value = r[1];
		document.F.View_fFlags.value = r[2];
		document.F.View_Options.value = r[3];
		document.F.View_ViewFlags.value = r[4];
	}
	document.F.Conf_SizeFormat.value = Ctrl[3] || 0;
	for (let i = 0; i < document.F.length; ++i) {
		o = document.F[i];
		if (SameText(o.type, 'checkbox')) {
			if (!/^!?Conf_/.test(o.id)) {
				o.checked = false;
			}
		}
	}
	LoadChecked(document.F);
	const s = GetWebColor(document.F.Conf_TrailColor.value);
	document.F.Conf_TrailColor.value = s;
	document.F.Color_Conf_TrailColor.style.backgroundColor = s;
	document.getElementById("_EDIT").checked = true;
}

async function ResizeTabPanel() {
	let h = 4.8;
	if (window.g_Inline && !/\w/.test(document.getElementById('toolbar').innerHTML)) {
		h -= 2.4;
	}
	if (window.g_NoTab) {
		h -= 2.4;
	}
	CalcElementHeight(g_ovPanel, h);
}

async function ClickTab(o, nMode) {
	nMode = GetNum(nMode);
	if (o && o.id) {
		nTabIndex = o.id.replace(new RegExp('tab', 'g'), '') - 0;
	}
	let i = 0;
	let ovTab;
	while (ovTab = document.getElementById('tab' + i)) {
		const ovPanel = document.getElementById('panel' + i);
		if (i == nTabIndex) {
			try {
				ovTab.focus();
			} catch (e) { }
			ovTab.className = 'activetab';
			ovTab.style.zIndex = 11;
			ovPanel.style.display = 'block';
			g_ovPanel = ovPanel;
			await ResizeTabPanel();
			if (window.OnTabChanged) {
				OnTabChanged(i);
			}
		} else {
			ovTab.className = i < nTabIndex ? 'tab' : 'tab2';
			ovTab.style.zIndex = 10 - i;
			ovPanel.style.display = 'none';
		}
		++i;
	}
	nTabMax = i;
}

async function ClickTree(o, nMode, strChg, bForce) {
	CloseFindDialog();
	if ("string" === typeof o) {
		o = document.getElementById(o);
	}
	if (strChg) {
		g_Chg[strChg] = true;
	}
	nMode = GetNum(nMode);
	let newTab = TabIndex != -1 ? TabIndex : 0;
	if (o && o.id) {
		const res = /tab([^_]+)(_?)(.*)/.exec(o.id);
		if (res) {
			newTab = res[1] + res[2] + res[3];
			ClickButton(res[1], true);
			if (res[3]) {
				if (res[1] == 1) {
					setTimeout(function (Id) {
						if (ui_.elAddons[Id]) {
							if (!ui_.elAddons[Id].contentWindow || !ui_.elAddons[Id].contentWindow.document.body.innerHTML) {
								AddonOptions(Id, null, null, true);
							}
						}
					}, 999, res[3]);
				}
			}
			ShowButtons(/^1$|^1_1|^1_3$/.test(newTab), res[1] == 2, newTab);
			if (nMode == 0) {
				switch (res[1] - 0) {
					case 1:
						setTimeout(LoadAddons);
						if (newTab == '1_1') {
							setTimeout(GetAddons);
						}
						if (newTab == '1_3') {
							setTimeout(GetIconPacks);
						}
						if (newTab == '1_4') {
							setTimeout(GetLangPacks);
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
			const o = document.getElementById("DefaultLangID");
			if (o && o.innerHTML == "") {
				let s = await GetLangId(1);
				const s2 = await GetLangId(2);
				if (s != s2) {
					s += ' (' + s2 + ')';
				}
				o.innerHTML = s;
			}
		}
		let ovTab = document.getElementById('tab' + TabIndex);
		if (ovTab) {
			const ovPanel = document.getElementById('panel' + TabIndex) || document.getElementById('panel' + TabIndex.replace(/_\w+/, ""));
			ovTab.className = 'button';
			ovPanel.style.display = 'none';
		}
		TabIndex = newTab;
		ovTab = document.getElementById('tab' + TabIndex);
		if (ovTab) {
			ovPanel = document.getElementById('panel' + TabIndex) || document.getElementById('panel' + TabIndex.replace(/_\w+/, ""));
			const res = /2_(.+)/.exec(TabIndex);
			if (res) {
				document.F.Menus.selectedIndex = res[1];
				SwitchMenus(document.F.Menus);
			}
			ovTab.className = 'hoverbutton';
			ovPanel.style.display = 'block';
			CalcElementHeight(ovPanel, 3);
			const o = document.getElementById("tab_");
			CalcElementHeight(o, 3);
			const i = ovTab.offsetTop - o.scrollTop;
			if (i < 0 || i >= o.offsetHeight - ovTab.offsetHeight) {
				ovTab.scrollIntoView(i < 0);
			}
		}
	}
}

function ClickButton(n, f) {
	const o = document.getElementById("tabbtn" + n);
	const op = document.getElementById("tab" + n + "_");
	if (f || (f == null && /none/i.test(op.style.display))) {
		o.innerHTML = BUTTONS.opened;
		op.style.display = "block";
	} else {
		o.innerHTML = BUTTONS.closed;
		op.style.display = "none";
	}
}

function SetRadio(o) {
	const ar = o.id.split("=");
	const el = o.form[ar[0]];
	el.value = ar[1];
	FireEvent(el, "change");
}

async function SetCheckbox(o) {
	const ar = o.id.split(/:!?/);
	const el = o.form[ar[0]];
	if ((o.checked ? 1 : 0) ^ (/!/.test(o.id) ? 1 : 0)) {
		el.value |= Number(ar[1]);
	} else {
		el.value &= ~Number(ar[1]);
	}
	FireEvent(el, "change");
}

function SetValue(n, v) {
	if (v.value != '!') {
		n.value = v.value.replace(/\\n/, "\n");
	}
}

function AddValue(name, i, min, max) {
	const o = document.F[name];
	o.value = Math.min(Math.max(i + GetNum(o.value), min), max);
}

function ChooseColor1(o) {
	setTimeout(async function () {
		const o2 = o.form[o.id.replace("Color_", "")];
		const c = await ChooseColor(o2.value || o2.placeholder);
		if (c != null) {
			o2.value = c;
			o.style.backgroundColor = GetWebColor(c);
		}
	}, 99);
}

function ChooseColor2(o) {
	setTimeout(async function () {
		const o2 = o.form[o.id.replace("Color_", "")];
		let c = await ChooseColor(await GetWinColor(o2.value || o2.placeholder));
		if (c != null) {
			c = GetWebColor(c);
			o2.value = c;
			o.style.backgroundColor = c;
		}
	}, 99);
}

function ChooseFolder1(o) {
	setTimeout(async function () {
		const r = await BrowseForFolder(o.value);
		if (r) {
			if (await api.GetKeyState(VK_CONTROL) < 0 && /textarea/i.test(o.tagName)) {
				const ar = o.value.replace(/\s+$/g, "").split(/\n/);
				ar.push(r);
				o.value = ar.join("\n");
			} else {
				o.value = r;
			}
		}
	}, 99);
}

async function AddTabControl() {
	if (document.F.Tab_Width.value == 0) {
		MessageBox("Please enter the width.", TITLE, MB_ICONEXCLAMATION);
		return;
	}
	if (document.F.Tab_Height.value == 0) {
		MessageBox("Please enter the height.", TITLE, MB_ICONEXCLAMATION);
		return;
	}
	const TC = await te.CreateCtrl(CTRL_TC, document.F.Tab_Left.value, document.F.Tab_Top.value, document.F.Tab_Width.value, document.F.Tab_Height.value, document.F.Tab_Style.value, document.F.Tab_Align.value, document.F.Tab_TabWidth.value, document.F.Tab_TabHeight.value);
	TC.Selected.Navigate2("C:\\", SBSP_NEWBROWSER, document.F.View_Type.value, document.F.View_ViewMode.value, document.F.View_fFlags.value, 0, document.F.View_Options.value, document.F.View_ViewFlags.value);
}

async function DelTabControl() {
	const TC = await te.Ctrl(CTRL_TC);
	if (TC) {
		TC.Close();
	}
}

async function SetTabControls() {
	if (g_Chg.Tab) {
		if (document.getElementById("Conf_TabDefault").checked) {
			const cTC = await te.Ctrls(CTRL_TC, false, window.chrome);
			for (let i = 0; i < cTC.length; ++i) {
				SetTabControl(cTC[i]);
			}
		} else {
			const TC = await te.Ctrl(CTRL_TC);
			if (TC) {
				SetTabControl(TC);
			}
		}
	}
}

async function SetFolderViews() {
	if (g_Chg.View) {
		const r = await Promise.all([api.CreateObject("Object"), api.CreateObject("Object"), te.Data.Conf_Layout]);
		r[0].Layout = r[2];
		r[0].ViewMode = document.F.View_ViewMode.value;
		MainWindow.g_.TEData = r[0];
		r[1].All = document.getElementById("Conf_ListDefault").checked;
		r[1].FolderFlags = document.F.View_fFlags.value;
		r[1].Options = document.F.View_Options.value;
		r[1].ViewFlags = document.F.View_ViewFlags.value;
		r[1].Type = document.F.View_Type.value;
		MainWindow.g_.FVData = r[1];
	}
}

async function SetTreeControls() {
	if (g_Chg.Tree) {
		const o = await api.CreateObject("Object");
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

function SetTabControl(TC) {
	if (TC) {
		TC.Style = document.F.Tab_Style.value;
		TC.Align = document.F.Tab_Align.value;
		TC.TabWidth = document.F.Tab_TabWidth.value;
		TC.TabHeight = document.F.Tab_TabHeight.value;
	}
}

async function GetTabControl() {
	const TC = await te.Ctrl(CTRL_TC);
	if (TC) {
		const r = await Promise.all([TC.Left, TC.Top, TC.Width, TC.Height]);
		document.F.Tab_Left.value = r[0];
		document.F.Tab_Top.value = r[1];
		document.F.Tab_Width.value = r[2];
		document.F.Tab_Height.value = r[3];
	}
}

async function MoveTabControl() {
	const TC = await te.Ctrl(CTRL_TC);
	if (TC) {
		if (document.F.Tab_Width.value != "" && document.F.Tab_Height.value != "") {
			TC.Left = document.F.Tab_Left.value;
			TC.Top = document.F.Tab_Top.value;
			TC.Width = document.F.Tab_Width.value;
			TC.Height = document.F.Tab_Height.value;
		}
	}
}

async function SwapTabControl() {
	const TC1 = await te.Ctrl(CTRL_TC);
	const cTC = await te.Ctrls(CTRL_TC, true, window.chrome);
	const nLen = cTC.Count;
	for (let i = 0; i < nLen; ++i) {
		const TC = cTC[i];
		if (await TC.Id != await TC1.Id && await TC.Left == document.F.Tab_Left.value && await TC.Top == document.F.Tab_Top.value &&
			await TC.Width == document.F.Tab_Width.value && await TC.Height == document.F.Tab_Height.value) {
			TC.Left = await TC1.Left;
			TC.Top = await TC1.Top;
			TC.Width = await TC1.Width;
			TC.Height = await TC1.Height;
			TC1.Left = document.F.Tab_Left.value;
			TC1.Top = document.F.Tab_Top.value;
			TC1.Width = document.F.Tab_Width.value;
			TC1.Height = document.F.Tab_Height.value;
			break;
		}
	}
}

async function InitConfig(o) {
	if (ui_.Installed == await te.Data.DataFolder) {
		return;
	}
	if (!await confirmOk()) {
		return;
	}
	await api.SHFileOperation(FO_MOVE, BuildPath(ui_.Installed, "layout"), await te.Data.DataFolder, 0, false);
	o.disabled = true;
}

function SelectPos(o, s) {
	const v = o[o.selectedIndex].value;
	if (v != "") {
		o.form[s].value = v;
	}
}

function SwitchMenus(o) {
	if (g_x.Menus) {
		g_x.Menus.style.display = "none";
		if (!o) {
			o = document.F.Menus;
		}
		for (let i = o.length; i-- > 0;) {
			const a = o[i].value.split(",");
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
	if (o && o.value) {
		const a = o.value.split(",");
		g_x.Menus = o.form["Menus_" + a[0]];
		g_x.Menus.style.display = "inline";
		o.form["Menus_Base"].selectedIndex = a[1];
		o.form["Menus_Base"].disabled = /Default|Alias/.test(a[0]);
		o.form["Menus_Pos"].value = GetNum(a[2]);
		CancelX("Menus");
	}
}

function SwitchX(mode, o, form) {
	g_x[mode].style.display = "none";
	g_x[mode] = (form || o.form)[mode + o.value];
	g_x[mode].style.display = "inline";
	CancelX(mode);
}

function ClearX(mode) {
	g_Chg.Data = null;
}

function CancelX(mode) {
	g_x[mode].selectedIndex = -1;
	EnableSelectTag(g_x[mode]);
}

ChangeX = function (mode) {
	g_Chg.Data = mode;
	g_bChanged = true;
}

function ConfirmX(bCancel, fn) {
	g_bChanged |= g_Chg.Data;
	return SetOptions(function () {
		SetChanged(fn);
	},
	function () {
		g_bChanged = false;
		ClearX(g_Chg.Data);
	}, !bCancel, true);
}

async function SetOptions(fnYes, fnNo, fnCancel) {
	if (document.activeElement && SameText(document.activeElement.value, await GetText("Cancel"))) {
		g_nResult = 2;
	}
	if (g_nResult == 2) {
		if (fnNo) {
			await fnNo();
		}
		return false;
	}
	if (document.activeElement) {
		document.activeElement.blur();
	}
	if (g_nResult == 1 && !g_Chg.Data) {
		if (fnYes) {
			await fnYes();
		}
		return true;
	}
	if (g_bChanged || g_Chg.Data) {
		switch (await MessageBox("Do you want to replace?", TITLE, MB_ICONQUESTION | (fnCancel ? MB_YESNOCANCEL : MB_YESNO))) {
			case IDYES:
				g_nResult = 1;
				if (fnYes) {
					await fnYes();
				}
				return true;
			case IDNO:
				g_nResult = 2;
				if (fnNo) {
					await fnNo();
				}
				if (g_nResult == 2) {
					ClearX();
					if (fnYes) {
						fnYes();
					}
					return true;
				}
				return false;
			case IDCANCEL:
				g_nResult = 4;
				if (fnCancel) {
					await fnCancel();
				}
				return;
			default:
				return false;
		}
	}
	g_nResult = 2;
	if (fnNo) {
		await fnNo();
	}
	return false;
}

async function EditMenus() {
	if (g_x.Menus.selectedIndex < 0) {
		return;
	}
	ClearX("Menus");
	const a = g_x.Menus[g_x.Menus.selectedIndex].value.split(g_sep);
	const a2 = a[0].split(/\\t/);
	if (!a[5]) {
		a2[0] = await GetText(a2[0]);
	}
	document.F.Menus_Key.value = a2.length > 1 ? await GetKeyName(a2.pop()) : "";
	document.F.Menus_Name.value = a2.join("\\t");
	document.F.Menus_Filter.value = a[1];
	document.F.Menus_Path.value = await MainWindow.OptionDecode(a[3], a[2]);
	SetType(document.F.Menus_Type, a[3]);
	document.F.Icon.value = a[4] || "";
	document.F.IconSize.value = a[6] || "";
	SetImage();
}

function EditXEx(o, s, form) {
	if (document.getElementById("_EDIT").checked) {
		o(s, form);
	}
}

EditX = async function (mode, form) {
	if (g_x[mode].selectedIndex < 0) {
		return;
	}
	if (!form) {
		form = document.F;
	}
	ClearX(mode);
	const a = g_x[mode][g_x[mode].selectedIndex].value.split(g_sep);
	form[mode + mode].value = a[0];
	form[mode + "Path"].value = await MainWindow.OptionDecode(a[2], a[1]);
	SetType(form[mode + "Type"], a[2]);
	if (SameText(mode, "key")) {
		SetKeyShift();
	}
	const o = form[mode + "Name"];
	if (o) {
		o.value = a[3] ? await GetText(a[3]) : "";
	}
}

function SetType(o, value) {
	value = value.replace(/&|\.\.\.$/g, "");
	let i = o.length;
	while (--i >= 0) {
		if (SameText(o[i].value, value)) {
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

function InsertX(sel) {
	if (!sel) {
		return;
	}
	++sel.length;
	if (sel.selectedIndex < 0) {
		sel.selectedIndex = sel.length - 1;
	} else {
		++sel.selectedIndex;
		for (let i = sel.length; i-- > sel.selectedIndex;) {
			sel[i].text = sel[i - 1].text;
			sel[i].value = sel[i - 1].value;
		}
	}
}

function AddX(mode, fn, form) {
	InsertX(g_x[mode]);
	(fn || ReplaceX)(mode, form);
	EnableSelectTag(g_x[mode]);
}

async function ReplaceMenus() {
	ClearX("Menus");
	if (g_x.Menus.selectedIndex < 0) {
		InsertX(g_x.Menus);
	}
	const sel = g_x.Menus[g_x.Menus.selectedIndex];
	const o = document.F.Menus_Type;
	let s = await GetSourceText(document.F.Menus_Name.value.replace(/\\/g, "/"));
	const org = (SameText(s, document.F.Menus_Name.value) && await api.GetKeyState(VK_SHIFT) >= 0) ? "1" : ""
	if (document.F.Menus_Key.value.length) {
		s += "\\t" + await GetKeyKeyG(document.F.Menus_Key.value);
	}
	SetMenus(sel, [s, document.F.Menus_Filter.value, await MainWindow.OptionEncode(o[o.selectedIndex].value, document.F.Menus_Path.value), o[o.selectedIndex].value, document.F.Icon.value, org, document.F.IconSize.value]);
	EnableSelectTag(g_x.Menus);
	g_Chg.Menus = true;
}

async function ReplaceX(mode, form) {
	if (!g_x[mode]) {
		return;
	}
	if (!form) {
		form = document.F;
	}
	ClearX(mode);
	if (g_x[mode].selectedIndex < 0) {
		InsertX(g_x[mode]);
	}
	const sel = g_x[mode][g_x[mode].selectedIndex];
	const o = form[mode + "Type"];
	const o2 = form[mode + "Name"];
	SetData(sel, [form[mode + mode].value, await MainWindow.OptionEncode(o[o.selectedIndex].value, form[mode + "Path"].value), o[o.selectedIndex].value, o2 ? await GetSourceText(o2.value) : ""]);
	EnableSelectTag(g_x[mode]);
	g_Chg[mode] = true;
	g_bChanged = true;
}

function RemoveX(mode) {
	ClearX(mode);
	if (g_x[mode].selectedIndex < 0) {
		return;
	}
	ConfirmThenExec(GetText("Remove"), function () {
		let i = g_x[mode].selectedIndex;
		let j = i;
		while (j >= 0 && g_x[mode][j]) {
			g_x[mode][j] = null;
			j = g_x[mode].selectedIndex;
		}
		g_Chg[mode] = true;
		g_bChanged = true;
		if (i >= g_x[mode].length) {
			i = g_x[mode].length - 1;
		}
		if (i >= 0) {
			g_x[mode].selectedIndex = i;
			FireEvent(g_x[mode][i], "change");
		}
		g_bChanged = false;
		EnableSelectTag(g_x[mode]);
	});
}

function MoveX(mode, n) {
	if (n < 0) {
		for (let i = 0; i < g_x[mode].length + n; ++i) {
			if (!g_x[mode][i].selected && g_x[mode][i + 1].selected) {
				const ar = [g_x[mode][i].text, g_x[mode][i].value];
				g_x[mode][i].text = g_x[mode][i + 1].text;
				g_x[mode][i].value = g_x[mode][i + 1].value;
				g_x[mode][i + 1].text = ar[0];
				g_x[mode][i + 1].value = ar[1];
				g_x[mode][i + 1].selected = false;
				g_x[mode][i].selected = true;
			}
		}
	} else {
		for (let i = g_x[mode].length; i-- > n;) {
			if (!g_x[mode][i].selected && g_x[mode][i - 1].selected) {
				const ar = [g_x[mode][i].text, g_x[mode][i].value];
				g_x[mode][i].text = g_x[mode][i - 1].text;
				g_x[mode][i].value = g_x[mode][i - 1].value;
				g_x[mode][i - 1].text = ar[0];
				g_x[mode][i - 1].value = ar[1];
				g_x[mode][i - 1].selected = false;
				g_x[mode][i].selected = true;
			}
		}
	}
	g_Chg[mode] = true;
	g_bChanged = true;
	EnableSelectTag(g_x[mode]);
}

async function CreateMenuX(mode) {
	const sel = g_x[mode];
	if (sel && sel.selectedIndex >= 0) {
		InputDialog((await Promise.all([GetText("Menus"), GetText("Name")])).join("\n"), "", async function (sName, mode) {
			WebBrowser.Focus();
			if (sName) {
				const sOrg = (await api.GetKeyState(VK_SHIFT) >= 0) ? "1" : "";
				CreateMenu(mode, sName, sOrg);
			}
		}, mode);
	}
}

function CreateMenu(mode, sName, sOrg) {
	const sel = g_x[mode];
	if (sel && sel.selectedIndex >= 0) {

		ClearX(mode);
		let iFirst = 0;
		let iLast  = sel.length - 1;
		while (!sel[iFirst].selected) {++iFirst;}
		while (!sel[iLast].selected)  {--iLast; }

		sel.length += 2;
		for (let i = sel.length - 3; i > iLast; --i) {
			sel[i + 2].text  = sel[i].text;
			sel[i + 2].value = sel[i].value;
		}
		for (let i = iLast; i >= iFirst; --i) {
			sel[i + 1].text  = sel[i].text;
			sel[i + 1].value = sel[i].value;
			sel[i + 1].selected = false;
		}

		iLast += 2;
		if (mode == "Menus") {
			SetMenus(sel[iFirst], [sName, "", "Open",  "Menus", "", sOrg, ""]);
			SetMenus(sel[iLast],  [sName, "", "Close", "Menus", "", sOrg, ""]);
		} else {
			SetData(sel[iFirst], [sName, "Open",  "Menus", "", ""]);
			SetData(sel[iLast],  [sName, "Close", "Menus", "", ""]);
		}
		g_Chg[mode] = true;
		g_bChanged = true;
		sel[iLast].selected  = true;
		sel[iFirst].selected = true;
		FireEvent(sel[iFirst], "change");
		EnableSelectTag(sel);
	}
}

async function SetMenus(sel, a) {
	sel.value = PackData(a);
	const a2 = a[0].split(/\\t/);
	sel.text = [a[5] ? a2[0] : await GetText(a2[0]), a[1]].join(" ").replace(/[\r\n].*/, "");
	if (!sel.text.length) {
		sel.text = "********";
	}
}

async function LoadMenus(nSelected) {
	let oa = document.F.Menus_Type;
	if (!g_x.Menus) {
		let arFunc = await api.CreateObject("Array");
		await MainWindow.RunEvent1("AddType", arFunc);
		if (window.chrome) {
			arFunc = await api.CreateObject("SafeArray", arFunc);
		}
		for (let i = 0; i < arFunc.length; ++i) {
			const o = oa[++oa.length - 1];
			o.value = arFunc[i];
			o.innerText = await GetText(arFunc[i]);
		}

		oa = document.F.Menus;
		oa.length = 0;

		for (let j in g_arMenuTypes) {
			const s = g_arMenuTypes[j];
			document.getElementById("Menus_List").insertAdjacentHTML("beforeend", ['<select name="Menus_', s, '" size="17" style="width: 12em; height: 32em; height: calc(100vh - 6em); min-height: 20em; display: none" onchange="EditXEx(EditMenus)" ondblclick="EditMenus()" oncontextmenu="CancelX(\'Menus\')" data-menu multiple></select>'].join(""));
			const menus = await MenuGetElementsByTagNameUI(s);
			if (menus && menus.length) {
				oa[++oa.length - 1].value = s + "," + menus[0].getAttribute("Base") + "," + menus[0].getAttribute("Pos");
				const o = document.F["Menus_" + s];
				const items = menus[0].getElementsByTagName("Item");
				if (items) {
					let i = items.length;
					o.length = i;
					while (--i >= 0) {
						const item = items[i];
						SetMenus(o[i], [item.getAttribute("Name"), item.getAttribute("Filter"), item.text || item.textContent, item.getAttribute("Type"), item.getAttribute("Icon"), item.getAttribute("Org"), item.getAttribute("Height")]);
					}
				}
			} else {
				oa[++oa.length - 1].value = s;
			}
			oa[oa.length - 1].text = await GetText(s);
		}
		SwitchMenus(oa[nSelected]);
	}
	for (let j in g_arMenuTypes) {
		const ar = String(g_MenuType).split(",");
		if (SameText(ar[0], g_arMenuTypes[j])) {
			nSelected = oa.length - 1;
			oa[nSelected].selected = true;
			g_MenuType = j;
			setTimeout(async function (o, v) {
				await ClickTree("tab2_" + g_MenuType);
				await EditMenus();
				g_MenuType = "";
				if (isFinite(v)) {
					o.selectedIndex = v;
					EnableSelectTag(o);
					FireEvent(o, "change");
				}
			}, 99, document.F["Menus_" + ar[0]], ar[1]);
		}
	}
}

async function LoadX(mode, fn, form) {
	if (!g_x[mode]) {
		if (!form) {
			form = document.F;
		}
		let arFunc = await api.CreateObject("Array");
		await MainWindow.RunEvent1("AddType", arFunc);
		if (window.chrome) {
			arFunc = await api.CreateObject("SafeArray", arFunc);
		}
		let oa = form[mode + "Type"] || form.Type;
		while (oa.length) {
			oa.removeChild(oa[0]);
		}
		for (let i = 0; i < arFunc.length; ++i) {
			const o = oa[++oa.length - 1];
			o.value = arFunc[i];
			o.innerText = await GetText(arFunc[i]);
		}
		g_x[mode] = form[mode + "All"];
		if (g_x[mode]) {
			oa = form[mode];
			oa.length = 0;
			const xml = await OpenXmlUI(mode + ".xml", false, true);
			for (let j in g_Types[mode]) {
				let o = oa[++oa.length - 1];
				o.text = await GetTextEx(g_Types[mode][j]);
				o.value = g_Types[mode][j];
				o = form[mode + g_Types[mode][j]];
				let items = xml.getElementsByTagName(g_Types[mode][j]);
				let i = items.length;
				if (i == 0 && g_Types[mode][j] == "List") {
					items = xml.getElementsByTagName("Folder");
					i = items.length;
				}
				o.length = i;
				while (--i >= 0) {
					const item = items[i];
					let s = item.getAttribute(mode);
					if (SameText(mode, "Key")) {
						const ar = /,$/.test(s) ? [s] : s.split(",");
						for (let k = ar.length; k--;) {
							ar[k] = await GetKeyNameG(ar[k]);
						}
						s = ar.join(",");
					}
					SetData(o[i], [s, item.text || item.textContent, item.getAttribute("Type"), item.getAttribute("Name")]);
				}
			}
		} else {
			g_x[mode] = form.List;
			g_x[mode].length = 0;
			const xml = await OpenXmlUI(AddonName + ".xml", false, true);
			const items = xml.getElementsByTagName("Item");
			let i = items.length;
			g_x[mode].length = i;
			while (--i >= 0) {
				const item = items[i];
				SetData(g_x[mode][i], [item.getAttribute("Name"), item.text || item.textContent, item.getAttribute("Type"), item.getAttribute("Icon"), item.getAttribute("Height")]);
			}
		}
		EnableSelectTag(g_x[mode]);
		fn && fn();
	}
}

async function SaveMenus() {
	if (g_Chg.Menus) {
		const xml = CreateXmlUI(true);
		for (let j in g_arMenuTypes) {
			const o = document.F["Menus_" + g_arMenuTypes[j]];
			const items = xml.createElement(g_arMenuTypes[j]);
			let a = document.F.Menus[j].value.split(",");
			items.setAttribute("Base", GetNum(a[1]));
			items.setAttribute("Pos", GetNum(a[2]));
			for (let i = 0; i < o.length; ++i) {
				const item = xml.createElement("Item");
				a = o[i].value.split(g_sep);
				SetXmlText(item, a[2]);
				item.setAttribute("Name", a[0]);
				item.setAttribute("Filter", a[1]);
				item.setAttribute("Type", a[3]);
				item.setAttribute("Icon", a[4]);
				if (a[5]) {
					item.setAttribute("Org", 1);
				}
				if (a[6]) {
					item.setAttribute("Height", a[6]);
				}
				items.appendChild(item);
			}
			xml.documentElement.appendChild(items);
		}
		await SaveXmlExUI("menus.xml", xml);
	}
}

async function GetKeyKeyEx(s) {
	const n = await MainWindow.GetKeyKey(s);
	return n & 0xff ? "$" + n.toString(16) : s;
}

async function GetKeyKeyG(s) {
	const n = await GetKeyKeyEx(s);
	s = await GetKeyName(n, true);
	return /^\w$|\+\w$/i.test(s) ? s : n;
}

async function GetKeyNameG(s) {
	return await GetKeyName(/^$/.test(s) ? s : await GetKeyKeyEx(s));
}

async function SaveX(mode, form) {
	if (g_Chg[mode]) {
		const xml = CreateXmlUI(true);
		for (let j in g_Types[mode]) {
			const o = (form || document.F)[mode + g_Types[mode][j]];
			for (let i = 0; i < o.length; ++i) {
				const item = xml.createElement(g_Types[mode][j]);
				const a = o[i].value.split(g_sep);
				SetXmlText(item, a[1]);
				let s = a[0];
				if (SameText(mode, "key")) {
					const ar = /,$/.test(s) ? [s] : s.split(",");
					for (let k = ar.length; k--;) {
						ar[k] = await GetKeyKeyG(ar[k]);
					}
					s = ar.join(",");
				} else {
					item.setAttribute("Name", a[3]);
				}
				item.setAttribute(mode, s);
				item.setAttribute("Type", a[2]);
				xml.documentElement.appendChild(item);
			}
		}
		SaveXmlExUI(mode.toLowerCase() + ".xml", xml);
	}
}

async function SaveAddons() {
	if (g_Chg.Addons || await te.Data.bErrorAddons) {
		te.Data.bErrorAddons = false;
		const Addons = await api.CreateObject("Object");
		const table = document.getElementById("Addons");
		for (let j = 0; j < table.rows.length; ++j) {
			const div = table.rows[j].cells[0].firstChild;
			if (div) {
				const Id = div.id.replace("Addons_", "").toLowerCase();
				Addons[Id] = document.getElementById("enable_" + Id).checked;
			}
		}
		await MainWindow.SaveAddons(Addons, window.g_bAddonLoading);
	}
}

function SetChanged(fn, form) {
	if (g_Chg.Data) {
		g_bChanged = true;
		if (g_x[g_Chg.Data] && g_x[g_Chg.Data].selectedIndex >= 0) {
			(fn || ReplaceX)(g_Chg.Data, form);
		} else {
			AddX(g_Chg.Data, fn, form);
		}
	}
}

function SetData(sel, a, t) {
	if (sel) {
		sel.value = PackData(a);
		if ("boolean" === typeof t) {
			t = (t ? String.fromCharCode(9745, 32) : String.fromCharCode(9744, 32)) + a[0];
		}
		sel.text = t || a[0];
	}
}

function PackData(a) {
	for (let i = a.length; i-- > 0;) {
		a[i] = String(a[i] || "").replace(g_sep, "`  ~");
	}
	return a.join(g_sep);
}

async function LoadAddons() {
	if (g_x.Addons) {
		OpenAddonsOptions();
		return;
	}
	g_x.Addons = true;

	const AddonId = {};
	const path = BuildPath(ui_.Installed, "addons");
	for (let Id, FileList = await GetFileList(BuildPath(path, "*"), false, window.chrome); Id = FileList.shift();) {
		if (Id != "." && Id != ".." && !AddonId[Id]) {
			AddonId[Id] = 1;
		}
	}
	const table = document.getElementById("Addons");
	table.ondragover = Over5;
	table.ondrop = Drop5;
	table.deleteRow(0);

	const xml = window.chrome ? new DOMParser().parseFromString(await te.Data.Addons.xml, "application/xml") : te.Data.Addons;
	if (xml.documentElement) {
		const items = xml.documentElement.childNodes;
		if (items) {
			const nLen = items.length;
			const sorted = document.getElementById("SortedAddons");
			const tcell = [];
			const scell = [];
			for (let i = 0; i < nLen; ++i) {
				tcell[i] = table.insertRow().insertCell();
				if (sorted.rows.length) {
					scell[i] = sorted.insertRow().insertCell();
				}
			}
			for (let i = 0; i < nLen; ++i) {
				const Id = items[i].nodeName;
				if (AddonId[Id]) {
					const bEnable = GetNum(items[i].getAttribute("Enabled"));
					SetAddon(Id, bEnable, tcell[i]);
					if (sorted.rows.length) {
						SetAddon(Id, bEnable, scell[i], "Sorted_");
					}
					delete AddonId[Id];
				}
			}
		}
	}
	let r = [];
	for (let Id in AddonId) {
		r.push(fso.FileExists(BuildPath(path, Id, "config.xml")));
	}
	if (window.chrome) {
		r = await Promise.all(r);
	}
	for (let Id in AddonId) {
		if (r.shift()) {
			AddAddon(table, Id, false);
		}
	}
	OpenAddonsOptions();
	const nSort = await te.Data.Conf_AddonsSort;
	if (nSort) {
		SortAddons1(nSort - 1);
	}
}

function OpenAddonsOptions() {
	if (g_Id) {
		if (document.getElementById("opt_" + g_Id)) {
			AddonOptions(g_Id);
		}
		g_Id = "";
	}
}

function AddAddon(table, Id, bEnable, Alt) {
	SetAddon(Id, bEnable, table.insertRow().insertCell(), Alt);
	if (!Alt) {
		const sorted = document.getElementById("SortedAddons");
		if (sorted.rows.length) {
			SetAddon(Id, bEnable, sorted.insertRow().insertCell(), "Sorted_");
		}
	}
}

async function SetAddon(Id, bEnable, td, Alt) {
	if (!td) {
		td = document.getElementById(Alt || "Addons_" + Id).parentNode;
	}
	const info = GetAddonInfo(Id);
	const r = await Promise.all([info.Level, info.MinVersion, info.Config, AboutTE(0), info.Name, info.Version, info.Options, MainWindow.GetAddonIconImg(Id, '{ "style": "vertical-align: bottom" }', ui_.IconSize)]);
	const bLevel = (!window.chrome || r[0] > 1);
	const bConfig = GetNum(r[2]);
	const bMinVer = !bLevel || r[1] && r[3] < CalcVersion(r[1]);
	if (bMinVer || bConfig) {
		bEnable = false;
	}
	const s = ['<div ', (Alt ? '' : 'draggable="true" ondragstart="Start5(event, this)" ondragend="End5()"'), ' title="', Id, '" Id="', Alt || "Addons_", Id, '">'];
	s.push('<table><tr class="', bEnable || bConfig ? "box" : 'box disabled', '"><td class="middle">', (Alt ? '&nbsp;' : '<input type="radio" name="AddonId" id="_' + Id + '">'), '</td><td style="width: 100%; padding: .3em 0"><label for="_', Id, '">');
	if (r[7]) {
		s.push(r[7], "&nbsp;");
	}
	s.push('<b>', r[4], "</b>&nbsp;", r[5], '<br><button class="addonbutton" onclick="AddonInfo(\'', Id, '\',\'', Alt || "", '\')">Details</button>');
	s.push(' <button class="addonbutton" onclick="AddonRemove(\'', Id, '\');">Delete</button>&nbsp;(', Id, ')<div id="', Alt || "", "i_", Id, '"></div></td>');
	if (!bLevel) {
		if (!ui_.strNotSupported) {
			ui_.strNotSupported = await GetTextR("@comres.dll,-1845");
		}
		s.push('<td class="danger middle nowrap right">', ui_.strNotSupported, '&nbsp;</td>');
	} else if (bMinVer) {
		if (!ui_.strVersion) {
			ui_.strVersion = await GetTextR("@mstask.dll,-319");
		}
		if (!ui_.strIsRequired) {
			ui_.strIsRequired = await GetText("%s is required.");
		}
		s.push('<td class="danger middle nowrap right">', ui_.strIsRequired.replace("%s", ui_.strVersion + " " + (r[1].replace(/^20/, ""))), '</td>');
	} else if (r[6]) {
		s.push('<td class="nowrap middle" style="padding-right: 1em"><button onclick="AddonOptions(\'', Id, '\')" class="addonbuttonopt" id="opt_', Id, '">Options</button></td>');
	}
	const strEnable = bMinVer || bConfig ? 'visibility: hidden' : "";
	s.push('<td class="middle"><input type="checkbox" ', (Alt ? "" : 'id="enable_' + Id + '"'), ' onclick="AddonEnable(this, \'', Id, '\')" ', bEnable ? " checked" : "", ' style="', strEnable, '"></td>');
	s.push('<td class="middle"><label for="enable_', Id, '" class="nowrap" style="display: block; width: 6em;', strEnable, '">', bEnable ? "Enabled" : "Enable", '</label></td>');
	s.push('</tr></table></label></div>');
	td.style.visibility = "hidden";
	td.innerHTML = s.join("");
	Promise.all([ApplyLang(td), td]).then(function (r) {
		r[1].style.visibility = "";
	});
	if (!Alt) {
		const div = document.getElementById("Sorted_" + Id);
		if (div) {
			SetAddon(Id, bEnable, div.parentNode, "Sorted_");
		}
	}
}

function Start5(ev, o) {
	ev.dataTransfer.effectAllowed= 'move';
	g_drag5 = o.id;
	return true;
}

function End5() {
	g_drag5 = false;
}

function Over5(ev) {
	if (g_drag5) {
		PreventDefault(ev);
	}
}

function Drop5(ev) {
	if (g_drag5) {
		let o = document.elementFromPoint(ev.clientX, ev.clientY);
		do {
			if (/Addons_/i.test(o.id)) {
				setTimeout(function (src, dest) {
					AddonMoveEx(src, dest);
				}, 99, GetRowIndexById(g_drag5), GetRowIndexById(o.id));
				break;
			}
		} while (o = o.parentNode);
	}
}

function GetRowIndexById(id) {
	try {
		let o = document.getElementById(id);
		if (o) {
			while (o = o.parentNode) {
				if (o.rowIndex != null) {
					return o.rowIndex;
				}
			}
		}
	} catch (e) { }
}

async function AddonInfo(Id, Alt) {
	const o = document.getElementById(Alt + "i_" + Id);
	if (o.innerHTML) {
		if (!o.style.display) {
			o.style.display = "none";
			return;
		}
		o.style.display = "";
	} else {
		o.innerHTML = await GetAddonInfo(Id).Description;
	}
	if (o.getBoundingClientRect().bottom > document.getElementById("panel1").offsetHeight) {
		o.scrollIntoView(false);
	}
}

function AddonEnable(o, Id) {
	SetAddon(Id, o.checked);
	document.getElementById("enable_" + Id).checked = o.checked;
	g_Chg.Addons = true;
}

async function OptionMove(dir) {
	CloseFindDialog();
	if (/^1/.test(TabIndex)) {
		const r = document.F.AddonId;
		for (let i = 0; i < r.length; ++i) {
			if (r[i].checked) {
				if (await api.GetKeyState(VK_CONTROL) < 0) {
					if (dir < 0) {
						dir = -i;
					} else {
						dir = document.getElementById("Addons").rows.length - i - 1;
					}
				}
				try {
					AddonMoveEx(i, i + dir);
				} catch (e) { }
				break;
			}
		}
	} else if (/^2/.test(TabIndex)) {
		if (g_x.Menus.selectedIndex < 0 || g_x.Menus.selectedIndex + dir < 0 || g_x.Menus.selectedIndex + dir >= g_x.Menus.length) {
			return;
		}
		MoveX("Menus", dir);
	}
}

function AddonMoveEx(src, dest) {
	const table = document.getElementById("Addons");
	if (dest < 0 || dest >= table.rows.length || src == dest) {
		return false;
	}
	let tr = table.rows[src];
	let td = tr.cells[0];

	const s = td.innerHTML
	const md = td.onmousedown;
	const mu = td.onmouseup;
	const mm = td.onmousemove;

	table.deleteRow(src);

	tr = table.insertRow(dest);
	td = tr.insertCell();
	td.innerHTML = s;
	td.onmousedown = md;
	td.onmouseup = mu;
	td.onmousemove = mm;
	const o = document.getElementById('panel1');
	const i = td.offsetTop - o.scrollTop;
	if (i < 0 || i >= o.offsetHeight - td.offsetHeight) {
		td.scrollIntoView(i < 0);
	}
	document.F.AddonId[dest].checked = true;
	g_Chg.Addons = true;
	return false;
}

function AddonRemove(Id) {
	ConfirmThenExec(GetText("Delete"), async function () {
		MainWindow.SaveConfig();
		MainWindow.AddonDisabled(Id);
		if (await AddonBeforeRemove(Id) < 0) {
			return;
		}
		const path = BuildPath(ui_.Installed, "addons", Id);
		await DeleteItem(path, 0);
		setTimeout(async function () {
			if (!await IsExists(path)) {
				let table = document.getElementById("Addons");
				table.deleteRow(GetRowIndexById("Addons_" + Id));
				table = document.getElementById("SortedAddons");
				if (table.rows.length) {
					table.deleteRow(GetRowIndexById("Sorted_" + Id));
				}
				g_Chg.Addons = true;
			}
		}, 500);
	});
}

async function SetAddonsRssults() {
	for (let i in ui_.elAddons) {
		const w = ui_.elAddons[i].contentWindow;
		if (g_nResult == 1) {
			await w.TEOk();
		}
	}
	ui_.elAddons = {};
}

async function OkOptions() {
	await SetAddonsRssults();
	const hwnd = await GetTopWindow();
	if (await api.GetKeyState(VK_SHIFT) >= 0 && !await api.IsZoomed(hwnd) && !await api.IsIconic(hwnd)) {
		const r = 12 / (Math.abs(await MainWindow.DefaultFont.lfHeight) || 12);
		te.Data.Conf_OptWidth = GetNum(document.documentElement.offsetWidth || document.body.offsetWidth) * r;
		te.Data.Conf_OptHeight = GetNum(document.documentElement.offsetHeight || document.body.offsetHeight) * r;
	}
	SetChanged(ReplaceMenus);
	SetWindowAlpha(hwnd, 0);
	for (let i = 0; i < document.F.length; ++i) {
		const o = document.F[i];
		const Id = o.name || o.id;
		if (!/=|:/.test(Id)) {
			const res = /^(!?)(Tab_.+|Tree_.+|View_.+|Conf_.+)/.exec(Id);
			if (res && !/_$/.test(Id)) {
				const v = GetElementValue(o);
				te.Data[res[2]] = res[1] ? !v : v;
			}
		}
	}
	await SaveAddons();
	await SaveMenus();
	await SetTabControls();
	await SetTreeControls();
	await SetFolderViews();
	await MainWindow.RunEvent1("ConfigChanged", "Config");

	te.Data.bReload = true;
	MainWindow.g_.dlgs.Options = void 0;
	CloseWB(WebBrowser, true);
}

async function CancelOptions() {
	g_nResult = 0;
	if (await te.Data.bErrorAddons) {
		await SaveAddons();
		te.Data.bReload = true;
	}
	MainWindow.g_.dlgs.Options = void 0;
	CloseWB(WebBrowser, true);
}

async function ContinueOptions() {
	WebBrowser.PreventClose();
}

InitOptions = async function () {
	ApplyLang(document);
	Promise.all([GetText("Get %s"), GetTextR("@UIAutomationCore.dll,-220[Icons]"), GetTextR("@docprop.dll,-107"), GetText("File"), te.Data.DataFolder, fso.FolderExists(BuildPath(ui_.Installed, "layout")), GetLangId(), GetAltText("Get Icons"), GetAltText("Get Language file"), GetText("Options")]).then(function (r) {
		if (/^zh/i.test(r[6])) {
			r[0] = r[0].replace(" ", "");
		}
		document.getElementById("tab1_3").innerHTML = r[7] || r[0].replace("%s", r[1]);
		const sl = r[8] || r[0].replace("%s", r[2] + (/^ja|^zh/i.test(r[6]) ? "" : " ") + (r[3].toLowerCase()));
		document.getElementById("tab1_4").innerHTML = sl;
		setTimeout(function (sl) {
			document.getElementById("AddLang").innerHTML = sl;
		}, 99, sl);
		document.title = r[9] + " - " + TITLE;
		document.F.ButtonInitConfig.disabled = (ui_.Installed == r[4]) || !r[5];
	});
	MainWindow.g_.OptionsWindow = $;
	let data = [];
	const data1 = [];
	for (let i = 0; i < document.F.length; ++i) {
		const o = document.F[i];
		const Id = o.name || o.id;
		if (!/=|:/.test(Id)) {
			const res = /^(!?)(Tab_.+|Tree_.+|View_.+|Conf_.+)/.exec(Id);
			if (res) {
				data.push(te.Data[res[2]]);
				data1.push({ neg: res[1], el: o });
			}
		}
	}
	data = await Promise.all(data);
	while (data1.length) {
		const o = data1.shift();
		const v = data.shift();
		if (v != null || o.neg) {
			SetElementValue(o.el, o.neg ? !v : v);
		}
	}
	await ResetForm();
	const s = [];
	for (let i in g_arMenuTypes) {
		const j = g_arMenuTypes[i];
		s.push('<label id="tab2_' + i + '" class="button" style="width: 100%" onmousedown="ClickTree(this, null, \'Menus\');">' + await GetText(j) + '</label><br>');
	}
	document.getElementById("tab2_").innerHTML = s.join("");

	if (ui_.IEVer < 9) {
		window.addEventListener("resize", function () {
			clearTimeout(g_tidResize);
			g_tidResize = setTimeout(function () {
				ClickTree(null, null, null, true);
			}, 500);
		});
	}
	SetOnChangeHandler();
	for (let i = 6; i--;) {
		ClickButton(i, false);
	}
	SetTab(await dialogArguments.Data);
	document.F.Conf_IconSize.placeholder = screen.deviceYDPI / 4;
	document.F.Conf_InnerIconSize.placeholder = screen.deviceYDPI / 6;
	document.F.style.display = "";
	WebBrowser.OnClose = async function (WB) {
		g_bChanged |= g_Chg.Addons || g_Chg.Menus || g_Chg.Tab || g_Chg.Tree || g_Chg.View;
		if (!g_bChanged) {
			for (let i in ui_.elAddons) {
				const w = ui_.elAddons[i].contentWindow;
				if (!await w.IsChanged || await w.IsChanged()) {
					g_bChanged = true;
					break;
				}
			}
		}
		SetOptions(OkOptions, CancelOptions, ContinueOptions);
	};
}

OpenIcon = function (o) {
	setTimeout(async function () {
		o.cursor = "";
		o.onclick = null;
		const a = o.id.split(/,/);
		const px = screen.deviceYDPI / 3;
		const srcs = [];
		if (a[0] == "b") {
			const dllpath = BuildPath(system32, "ieframe.dll");
			a[0] = GetFileName(dllpath);
			const a1 = a[1];
			const hModule = await LoadImgDll(a, 0);
			if (hModule) {
				const lpbmp = isFinite(a[1]) ? a[1] - 0 : a[1];
				const himl = await api.ImageList_LoadImage(hModule, lpbmp, a[2], CLR_NONE, CLR_NONE, IMAGE_BITMAP, LR_CREATEDIBSECTION | LR_LOADTRANSPARENT);
				api.FreeLibrary(hModule);
				a[1] = a1;
				let nCount = 0;
				if (himl) {
					nCount = await api.ImageList_GetImageCount(himl);
					api.ImageList_Destroy(himl);
				}
				if (nCount == 0) {
					if (lpbmp == 204) {
						nCount = 20;
					} else if (lpbmp == 214) {
						nCount = 32;
					}
				}
				a[0] = GetFileName(dllpath);
				for (a[3] = 0; a[3] < nCount; a[3]++) {
					const tag = {
						"class": "button",
						onclick: "SelectIcon(this)",
						src: ("bitmap:" + a.join(",")).replace(/^bitmap:ieframe\.dll,214,24,/i, "icon:general,").replace(/^bitmap:ieframe\.dll,204,24,/i, "icon:browser,")
					}
					tag.title = tag.src;
					srcs.push(GetImgTag(tag, a[2]));
				}
			}
		} else if (a[0] == "i") {
			let dllPath = await ExtractMacro(te, a[1]);
			if (!/^[A-Z]:\\|^\\\\/i.test(dllPath)) {
				dllPath = BuildPath(system32, a[1]);
			}
			const nCount = await api.ExtractIconEx(dllPath, -1, null, null, 0);
			for (let i = 0; i < nCount; ++i) {
				const tag = {
					"class": "button",
					onclick: "SelectIcon(this)",
					src: ["icon:" + a[1], i].join(",")
				}
				tag.title = tag.src;
				srcs.push(GetImgTag(tag, px));
			}
		} else {
			for (let i = 0; i < a[3]; ++i) {
				const tag = {
					"class": "button",
					onclick: "SelectIcon(this)",
					src: "font:" + a[1] + ",0x" + (parseInt(a[2]) + i).toString(16)
				}
				tag.title = tag.src;
				srcs.push(GetImgTag(tag, px));
			}
		}
		if (srcs.length) {
			o.innerHTML = (await Promise.all(srcs)).join("");
		}
		document.body.style.cursor = "auto";
	});
	document.body.style.cursor = "wait";
}

async function SearchIcon(o) {
	document.F.ButtonSearch.disabled = true;
	InputDialog(system32, "*.exe;*.cpl", async function(sFilter, id1) {
		WebBrowser.Focus();
		if (sFilter) {
			if (!/^[A-Z]:\\|^\\\\/i.test(sFilter)) {
				dllPath = BuildPath(system32, sFilter);
			}
			const eFN = document.getElementById("footnote");
			const str = await GetTextR("@shell32.dll,-13581");
			eFN.innerHTML = str;
			for (let fn, FileList = await GetFileList(BuildPath(system32, sFilter), false, window.chrome); fn = FileList.shift();) {
				eFN.innerText = str + " " + FileList.length;
				const nCount = await api.ExtractIconEx(BuildPath(system32, fn), -1, null, null, 0);
				if (nCount) {
					const id = "i," + fn.toLowerCase();
					if (!document.getElementById(id)) {
						document.getElementById(id1).insertAdjacentHTML("beforeend", '<div id="' + id + '" onclick="OpenIcon(this)" style="cursor: pointer"><span class="tab">' + fn + ' : ' + nCount + '</span></div>');
						document.getElementById(id).scrollIntoView(false);
					}
				}
			}
			eFN.innerText = "";
		}
		document.F.ButtonSearch.disabled = false;
	}, o.id)
}

ReturnDialogResult = async function (WB) {
	if (g_nResult == 1) {
		dialogArguments.InvokeUI("SetElement", (await dialogArguments.Id).replace(/\t.*$/, ""), returnValue);
	}
	WB.Close();
}

InitDialog = async function () {
	document.title = TITLE;
	const Query = String(await window.dialogArguments.Query || location.search.replace(/\?/, "")).toLowerCase();
	let res = /^icon(.*)/.exec(Query);
	if (res) {
		const a = {
			"ieframe,214": "b,214,24",
			"ieframe,204": "b,204,24",

			"shell32": "i,shell32.dll"
		};
		const fontDir = await api.ILCreateFromPath(ssfFONTS).Path;
		if (await fso.FileExists(BuildPath(fontDir, "SegoeIcons.ttf")) || await fso.FileExists(await ExtractMacro(te, "%AppData%\\..\\Local\\Microsoft\\Windows\\Fonts\\Segoe Fluent Icons.ttf"))) {
			for (let i = 0xe700; i < 0xf8ff; i += 256) {
				a["Segoe Fluent Icons " + i.toString(16)] = "f,Segoe Fluent Icons," + i + ",256";
			}
		}
		if (await fso.FileExists(BuildPath(fontDir, "segmdl2.ttf"))) {
			for (let i = 0xe700; i < 0xf8ff; i += 256) {
				a["Segoe MDL2 Assets " + i.toString(16)] = "f,Segoe MDL2 Assets," + i + ",256";
			}
		}
		const sue = [0x2600, 0x2700];
		if (await fso.FileExists(BuildPath(fontDir, "seguiemj.ttf"))) {
			sue.unshift(0x1f300, 0x1f400, 0x1f500, 0x1f600, 0x1f700, 0x1f900, 0x1fa00);
		}
		for (let i = 0; i < sue.length; ++i) {
			a["Segoe UI Emoji " + sue[i].toString(16)] = "f,Segoe UI Emoji," + sue[i] + ",256";
		}
		if (await fso.FileExists(BuildPath(fontDir, "seguisym.ttf"))) {
			for (let i = 0xe000; i < 0xe1ff; i += 256) {
				a["Segoe UI Symbol " + i.toString(16)] = "f,Segoe UI Symbol," + i + ",256";
			}
		}
		a.Webdings = "f,Webdings,33,223";
		a.Wingdings = "f,Wingdings,33,224";
		a.Marlett = "f,Marlett,48,74";
		a["ieframe,697"] = "b,697,24";
		a.imageres = "i,imageres.dll";
		a.wmploc = "i,wmploc.dll";
		a.setupapi = "i,setupapi.dll";
		a.dsuiext = "i,dsuiext.dll";
		a.inetcpl = "i,inetcpl.cpl";

		const s = [];
		const path = BuildPath(await te.Data.DataFolder, "icons");
		for (let fn, FileList = await GetFileList(BuildPath(path, "*"), true, window.chrome); fn = FileList.shift();) {
			const wfd = fn.split("\n");
			if ((wfd[1] & FILE_ATTRIBUTE_DIRECTORY) && /^[a-z]/i.test(wfd[0])) {
				const arfn = [];
				const path2 = BuildPath(path, wfd[0]);
				for (let fn2, FileList2 = await GetFileList(BuildPath(path2, "*"), true, window.chrome); fn2 = FileList2.shift();) {
					const wfd2 = fn2.split("\n");
					if (!(wfd2[1] & FILE_ATTRIBUTE_DIRECTORY)) {
						arfn.push(wfd2[0].replace(/\..*$/, ""));
					}
				}
				if (arfn.length) {
					arfn.sort(function (a, b) {
						return a - b || a.localeCompare(b);
					});
					const px = screen.deviceYDPI / 3;
					let n;
					for (let i = 0; i < arfn.length; ++i) {
						if (n != arfn[i]) {
							n = arfn[i];
							const src = ["icon:" + GetFileName(path2), n].join(",");
							s.push(await GetImgTag({ src: src, title: src, "class": "button", onclick: "SelectIcon(this)" }, px));
						}
					}
					s.push("<br>");
				}
			}
		}
		const r = await Promise.all([GetText("General"), GetText("Browser")]);
		for (let i in a) {
			if (a[i].charAt(0) == "i" || res[1] != "2") {
				s.push('<div id="', a[i], '" onclick="OpenIcon(this)" style="cursor: pointer"><span class="tab">', i.replace(/ieframe,214/, r[0]).replace(/ieframe,204/, r[1]), '</span></div>');
			}
		}
		document.getElementById("Content").innerHTML = s.join("");
		document.getElementById("ButtonSearch").style.display = "inline-block";
		document.F.ButtonSearch.disabled = false;
		WebBrowser.OnClose = ReturnDialogResult;
	}
	if (Query == "mouse") {
		document.body.oncontextmenu = DetectProcessTag;
		document.getElementById("Content").innerHTML = '<canvas id="Gesture" style="width: 100%; height: 99%; text-align: center;"></canvas>';
		const el = document.getElementById("Gesture");
		el.onmousedown = async function (ev) {
			ev = ev || event;
			if (g_Gesture) {
				let n = 1;
				for (let i = 1; i < 6; ++i) {
					if (g_Gesture.indexOf(i + "") < 0) {
						if ((ev.buttons != null ? ev.buttons : ev.button) & n) {
							returnValue += i + "";
						}
					}
					n *= 2;
				}
			} else {
				returnValue = await GetGestureKey() + await GetGestureButton();
				api.RedrawWindow(await WebBrowser.hwnd, null, 0, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
				const el = document.getElementById("Gesture");
				el.width = el.offsetWidth;
				el.height = el.offsetHeight;
				const ctx = el.getContext('2d');
				ctx.beginPath();
				ctx.clearRect(0, 0, el.width, el.height);
			}
			document.F.q.value = returnValue;
			g_Gesture = returnValue;
			g_pt.x = ev.clientX;
			g_pt.y = ev.clientY;
			document.F.ButtonOk.disabled = false;
			const o = document.getElementById("Gesture");
			const s = o.style.height;
			o.style.height = "1px";
			o.style.height = s;
			return false;
		}

		el.onmouseup = async function () {
			g_Gesture = await GetGestureButton();
			return false;
		}

		el.onmousemove = async function (ev) {
			ev = ev || event;
			if (await api.GetKeyState(VK_XBUTTON1) < 0 || await api.GetKeyState(VK_XBUTTON2) < 0) {
				returnValue = await GetGestureKey() + await GetGestureButton();
				document.F.q.value = returnValue;
			}
			const buttons = ev.buttons != null ? ev.buttons : ev.button;
			if (document.F.q.value.length && (buttons & 2 || (await te.Data.Conf_Gestures && buttons & 4))) {
				const pt = { x: ev.clientX, y: ev.clientY };
				const x = (pt.x - g_pt.x);
				const y = (pt.y - g_pt.y);
				if (Math.abs(x) + Math.abs(y) >= 20) {
					const nTrail = await te.Data.Conf_TrailSize;
					if (nTrail) {
						if (ui_.IEVer > 8) {
							const el = document.getElementById("Gesture");
							const ctx = el.getContext('2d');
							ctx.beginPath();
							ctx.strokeStyle = GetWebColor(await te.Data.Conf_TrailColor);
							ctx.lineWidth = nTrail;
							ctx.moveTo(g_pt.x, g_pt.y);
							ctx.lineTo(pt.x, pt.y);
							ctx.stroke();
						} else {
							const hwnd = WebBrowser.hwnd;
							const hdc = api.GetWindowDC(hwnd);
							if (hdc) {
								api.MoveToEx(hdc, g_pt.x, g_pt.y, null);
								const pen1 = api.CreatePen(PS_SOLID, nTrail, te.Data.Conf_TrailColor);
								const hOld = api.SelectObject(hdc, pen1);
								api.LineTo(hdc, pt.x, pt.y);
								api.SelectObject(hdc, hOld);
								api.DeleteObject(pen1);
								api.ReleaseDC(hwnd, hdc);
							}
						}
					}
					g_pt = pt;
					const s = (Math.abs(x) >= Math.abs(y)) ? ((x < 0) ? "L" : "R") : ((y < 0) ? "U" : "D");
					if (s != document.F.q.value.charAt(document.F.q.value.length - 1)) {
						returnValue += s;
						document.F.q.value = returnValue;
					}
				}
			}
			return false;
		}

		el.ondblclick = function () {
			returnValue += returnValue.replace(/\D/g, "");
			document.F.q.value = returnValue;
			return false;
		}

		el.onmousewheel = async function (ev) {
			ev = ev || event;
			returnValue = await GetGestureKey() + await GetGestureButton() + (ev.wheelDelta > 0 ? "8" : "9");
			document.F.q.value = returnValue;
			document.F.ButtonOk.disabled = false;
			return false;
		}

		document.getElementById("Selected").innerHTML = '<input type="text" name="q" class="full" autocomplete="off" onkeydown="setTimeout(\'returnValue=document.F.q.value\',100)">';
		WebBrowser.OnClose = ReturnDialogResult;
	}
	if (Query == "key") {
		returnValue = false;
		document.getElementById("Content").innerHTML = '<div style="padding: 8px;" style="display: block;"><label>Key</label><br><input type="text" name="q" autocomplete="off" class="full" style="ime-mode: disabled" onfocus="this.blur()"></div>';
		const fn = async function (ev) {
			const key = ev.keyCode;
			const k = await api.MapVirtualKey(key, 0) | ((key >= 33 && key <= 46 || key >= 91 && key <= 93 || key == 111 || key == 144) ? 256 : 0);
			if (k == 42 || k == 29 || k == 56 || k == 347) {
				return false;
			}
			const s = "$" + ((k | await GetKeyShift()).toString(16));
			returnValue = await GetKeyName(s);
			if (/^\$\w02a$|^\$\w01d$|^\$\w038$/i.test(returnValue)) {
				returnValue = await GetKeyName(s.slice(0, 3) + "1e").replace(/\+A$/, "");
			}
			document.F.q.value = returnValue;
			document.F.q.title = s;
			document.F.ButtonOk.disabled = false;
			return false;
		}
		window.addEventListener("keydown", fn);
		window.addEventListener("keyup", fn);
		setTimeout(FocusElement, 9, document.F.q);
		WebBrowser.OnClose = ReturnDialogResult;
	}
	if (Query == "new") {
		returnValue = false;
		ui_.ArgPath = await dialogArguments.path;
		const s = ['<div style="padding: 8px;" style="display: block;"><label><input type="radio" name="mode" id="folder" onclick="document.F.path.focus()">New Folder</label> <label><input type="radio" name="mode" id="file" onclick="document.F.path.focus()">New File</label><br>', await dialogArguments.path, '<br><input type="text" name="path" style="width: 100%"></div>'];
		document.getElementById("Content").innerHTML = s.join("");
		document.body.addEventListener("keydown", function (ev) {
			setTimeout(function () {
				document.F.ButtonOk.disabled = !document.F.path.value;
			}, 99);
			return KeyDownEvent(ev, document.F.path.value && function () {
				SetResult(1);
			}, function () {
				SetResult(2);
			});
		});

		document.body.addEventListener("paste", function () {
			setTimeout(function () {
				document.F.ButtonOk.disabled = !document.F.path.value;
			}, 99);
		});

		setTimeout(async function () {
			document.F[await dialogArguments.Mode].checked = true;
			FocusElement(document.F.path);
		}, 99);

		WebBrowser.OnClose = function (WB) {
			if (g_nResult == 1) {
				let path = document.F.path.value;
				if (path) {
					if (!/^[A-Z]:\\|^\\/i.test(path)) {
						path = BuildPath(ui_.ArgPath, path.trim());
					}
					if (GetElement("folder").checked) {
						MainWindow.CreateFolder(path);
					} else if (GetElement("file").checked) {
						MainWindow.CreateFile(path);
					}
				}
			}
			WB.Close();
		};
	}
	if (Query == "fileicon") {
		const s = PathUnquoteSpaces((await window.dialogArguments.Id).replace(/^.*\t/, ""));
		document.title = s;
		GetElement("Content").innerHTML = '<div id="i,' + s + '" style="cursor: pointer"></div>';
		await OpenIcon(GetElement("i," + s));
		WebBrowser.OnClose = ReturnDialogResult;
	}
	if (Query == "about") {
		const r = [MakeImgSrc(ui_.TEPath, 0, true, 48), AboutTE(2), GetTextR(ui_.bit + '-bit'), te.Data.DataFolder, AboutTE(3), GetLangId()];
		const s = ['<table style="border-spacing: 2em; border-collapse: separate; width: 100%"><tr><td>'];
		s.push('<img id="img1"></td><td style="width: 100%"><div class="nowrap"><span style="font-weight: bold; font-size: 120%" id="about2"></span> (<span id="bit1"></span>)</div>');
		s.push('<br><div id="lib"></div>');
		s.push('<br><label>@sqlsrv32.rll,-40028</label><br><a class="link" onclick="Run(1, this); return false" id="df1"></a><br>');
		s.push('<br><label>Information</label><input id="about3" type="text" style="width: 100%" onclick="this.select()" readonly><br>');
		const nAddon = r.length;
		const root = await te.Data.Addons.documentElement;
		if (root) {
			const items = await root.childNodes;
			if (items) {
				const nLen = await GetLength(items);
				for (let i = 0; i < nLen; ++i) {
					r.push(items[i].getAttribute("Enabled"), items[i].nodeName, GetAddonInfo(i).Version);
				}
			}
			s.push('<br><label>Add-ons</label><input id="UsedAddons" type="text" style="width: 100%" onclick="this.select()" readonly><br>');
		}
		s.push('<br><button id="ws1" onclick="Run(2)">Visit website</button>');
		s.push('&nbsp;<button onclick="Run(3)">Check for updates</button>');
		s.push('</td></tr></table>');
		document.getElementById("Content").innerHTML = s.join("");
		let website = "https://tablacus.github.io/explorer_en.html";
		if (r.length) {
			Promise.all(r).then(async function (r) {
				document.getElementById("img1").src = r[0];
				document.getElementById("about2").innerHTML = r[1];
				document.getElementById("bit1").innerHTML = r[2];
				document.getElementById("about3").value = r[4];
				if (/^ja/i.test(r[5])) {
					website = website.replace("_en", "");
				}
				document.getElementById("ws1").title = website;

				const el = document.getElementById("df1");
				el.innerHTML = el.href = BuildPath(r[3], "config");
				const ar = [];
				for (let i = nAddon; i < r.length; i += 3) {
					if (GetNum(r[i]) && r[i + 2]) {
						ar.push(r[i + 1] + " " + r[i + 2]);
					}
				}
				document.getElementById("UsedAddons").value = ar.join(",");
			});
		}
		document.F.ButtonOk.disabled = false;
		const el = document.getElementById("buttonCancel");
		el.innerHTML = await GetText("Copy");
		el.onclick = function () {
			clipboardData.setData("text", document.getElementById("about3").value + "\n" + document.getElementById("UsedAddons").value);
		}

		Run = async function (n, el) {
			setTimeout(function (n, path) {
				if (n == 2) {
					wsh.Run(website);
				} else if (n == 3) {
					if (window.chrome) {
						MainWindow.CheckUpdate();
					} else {
						MainWindow.setTimeout(MainWindow.CheckUpdate);
					}
				} else {
					MainWindow.Navigate(path + (n ? "" : "\\.."), SBSP_NEWBROWSER);
				}
				CloseWindow();
			}, 500, n, el && el.innerHTML);
		}

		s.length = 0;
		if (await api.GetModuleHandle(ui_.TEPath)) {
			s.push('<a href="', ui_.TEPath,'" class="link" onclick="Run(0, this);return false">', ui_.TEPath, "</a> (", await fso.GetFileVersion(ui_.TEPath), ")<br>");
		}
		for (let fn, FileList = await GetFileList(BuildPath(ui_.Installed, "lib", "*.dll"), false, window.chrome); fn = FileList.shift();) {
			const dll = BuildPath(ui_.Installed, "lib", fn);
			if (await api.GetModuleHandle(dll)) {
				s.push('<a href="', dll, '" class="link" onclick="Run(0, this); return false">', dll, "</a> (", await fso.GetFileVersion(dll), ")<br>");
			}
		}
		document.getElementById("lib").innerHTML = s.join("");
	}
	if (Query == "input") {
		returnValue = false;
		document.title = await window.dialogArguments.Id || TITLE;
		document.getElementById("Content").innerHTML = ['<div style="padding: 8px;" style="display: block;"><label>', EncodeSC(await dialogArguments.text).replace(/\r?\n/g, "<br>"), '<br><input type="text" name="text" class="full"></div>'].join("");
		document.body.addEventListener("keydown", function (ev) {
			return KeyDownEvent(ev, function () {
				SetResult(1);
			}, function () {
				SetResult(2);
			});
		});

		setTimeout(async function () {
			const el = document.F.text;
			el.value = await dialogArguments.defaultText;
			el.select();
			FocusElement(el);
		}, 99);

		WebBrowser.OnClose = async function (WB) {
			dialogArguments.callback(g_nResult == 1 ? document.F.text.value : null);
			WB.Close();
		};
		document.F.ButtonOk.disabled = false;
	}
	res = /^importScript:(.*)/.exec(Query);
	if (res) {
		importScript(res[1]);
	}

	DialogResize = function () {
		CalcElementHeight(document.getElementById("panel0"), 3);
	};

	if (ui_.IEVer < 9) {
		window.addEventListener("resize", function () {
			clearTimeout(g_tidResize);
			g_tidResize = setTimeout(DialogResize, 500);
		});
	}
	await ApplyLang(document);
	document.F.style.display = "";
	DialogResize();
}

SetLocations = async function () {
	const locs = {};
	const items = await MainWindow.g_.Locations;
	const r = await ObjectKeys(items, window.chrome);
	for (let i = 0; i < r.length; ++i) {
		locs[r[i]] = [];
		const dup = {};
		const item1 = window.chrome ? await api.CreateObject("SafeArray", await items[r[i]]) : items[r[i]];
		for (let j = item1.length; j--;) {
			const ar = item1[j].split("\t");
			if (dup[ar[0]]) {
				continue;
			}
			dup[ar[0]] = true;
			locs[r[i]].unshift(await MainWindow.GetAddonIconImg(ar[0], JSON.stringify({ title: await GetAddonInfo(ar[0]).Name }), 16, "text1") + '<span style="font-size: 1px"> </span>');
		}
	}
	for (let i in locs) {
		const s = locs[i].join("");
		try {
			document.getElementById('_' + i).innerHTML = s;
		} catch (e) { }
	}
}

InitLocation = function () {
	Promise.all([api.CreateObject("Array"), api.CreateObject("Object"), dialogArguments.Data.id, te.Data.DataFolder, te.Data.Addons.xml, GetLangId()]).then(async function (r) {
		let ar = r[0];
		const param = r[1];
		Addon_Id = r[2];
		ui_.DataFolder = r[3];
		for (let i = 10; i--;) {
			const o = document.getElementById('tab' + i);
			o.className = "tab";
			o.hidefocus = true;
			o.style.display = "none";
			(function (o) {
				o.onmousedown = function () {
					ClickTab(o, 1);
				};
				o.onfocus = function () {
					o.blur()
				};
			})(o);
		}
		if (window.chrome) {
			ui_.Addons = new DOMParser().parseFromString(r[4], "application/xml");
		}
		await LoadLang2(BuildPath("addons", Addon_Id, "lang", r[5] + ".xml"));
		await LoadAddon("js", Addon_Id, ar, param);
		if (window.chrome) {
			ar = await api.CreateObject("SafeArray", ar);
		}
		if (ar.length) {
			setTimeout(function (ar) {
				MessageBox(ar.join("\n\n"), TITLE, MB_OK);
			}, 500, ar);
		}
		ar = [];
		const s = "CSA";
		for (let i = 0; i < s.length; ++i) {
			ar.push('<button " title="', s.charAt(i), '" onclick="AddMouse(this)">', await MainWindow.g_.KeyState[i][0], "</button>");
		}
		document.getElementById("__MOUSEDATA").innerHTML = ar.join("");
		document.title = await GetAddonInfo(Addon_Id).Name;
		await ApplyLang(document);
		const item = GetAddonElement(Addon_Id);
		const Location = item.getAttribute("Location") || await param.Default;
		for (let i = document.L.length; i--;) {
			if (SameText(Location, document.L[i].value)) {
				document.L[i].disabled = false;
				document.L[i].checked = true;
				break;
			}
		}
		SetLocations();
		let oa = document.F.Menu;
		oa.length = 0;
		let o = oa[++oa.length - 1];
		o.value = "";
		o.text = await GetText("Select");
		for (let j in g_arMenuTypes) {
			const s = g_arMenuTypes[j];
			if (!/Default|Alias/.test(s)) {
				o = oa[++oa.length - 1];
				o.value = s;
				o.text = await GetText(s);
			}
		}
		ar = ["Key", "Mouse"];
		for (let i in ar) {
			const mode = ar[i];
			oa = document.F[mode + "On"];
			oa.length = 0;
			o = oa[++oa.length - 1];
			o.value = "";
			o.text = await GetText("Select");
			const r = await ObjectKeys(await MainWindow.eventTE[mode], window.chrome);
			for (j = 0; j < r.length; ++j) {
				o = oa[++oa.length - 1];
				o.value = r[j];
				o.text = await GetTextEx(r[j]);
			}
		}
		const el = document.F;
		for (let i = el.length; i--;) {
			GetAttribEx(item, el, el[i].id || el[i].name);
		}
		SetImage();
		LoadChecked(document.F);

		if (!await dialogArguments.Data.show) {
			dialogArguments.Data.show = "6";
			dialogArguments.Data.index = 6;
		}
		if (!/,/.test(await dialogArguments.Data.show)) {
			g_NoTab = true;
		} else {
			setTimeout(function () {
				document.getElementById("tabs").style.display = "block";
			}, 99);
			document.getElementById("tabs").style.display = "block";
		}
		if (/[8]/.test(await dialogArguments.Data.show)) {
			await MakeKeySelect();
			await SetKeyShift();
		}
		ar = (await dialogArguments.Data.show).split(/,/);
		for (let i in ar) {
			document.getElementById("tab" + ar[i]).style.display = "inline";
		}
		nTabIndex = await dialogArguments.Data.index;

		await SetOnChangeHandler();

		if (ui_.IEVer < 9) {
			window.addEventListener("resize", function () {
				clearTimeout(g_tidResize);
				g_tidResize = setTimeout(ResizeTabPanel, 500);
			});
		}
		IsChanged = function () {
			return g_bChanged || g_Chg.Data;
		};
		const TEOk1 = window.TEOk;
		TEOk = async function () {
			if (window.SaveLocation) {
				await SaveLocation();
			}
			MainWindow.g_.OptionsWindow = void 0;
			const items = await te.Data.Addons.getElementsByTagName(Addon_Id);
			if (await GetLength(items)) {
				let bConfigChanged = false;
				const item = await items[0];
				item.removeAttribute("Location");
				for (let i = document.L.length; i--;) {
					if (document.L[i].checked) {
						item.setAttribute("Location", document.L[i].value);
						bConfigChanged = true;
						break;
					}
				}
				const el = document.F;
				if (await dialogArguments.Data.show == "6") {
					el.Set.value = "";
				}
				for (let i = el.length; i--;) {
					const n = el[i].id || el[i].name;
					if (n && n.charAt(0) != "_") {
						if (n == "Key") {
							document.F[n].value = await GetKeyKeyG(document.F[n].value);
						}
						if (await SetAttribEx(item, document.F, n)) {
							bConfigChanged = true;
						}
					}
				}
				if (TEOk1) {
					if ("string" === typeof TEOk1) {
						MainWindow.InvokeUI(TEOk1);
					} else {
						TEOk1();
					}
				} else if (bConfigChanged) {
					te.Data.bReload = true;
					MainWindow.RunEvent1("ConfigChanged", "Addons");
				}
			}
		};

		if (await WebBrowser.OnClose) {
			g_Inline = true;
			AddRule(".ok, .cancel { display: none }");
		} else {
			WebBrowser.OnClose = async function (WB) {
				await SetOptions(TEOk, null, ContinueOptions);
				CloseWB(WB);
			};
		}
		if (item) {
			InitColor1(item);
		}
		ClickTab(null, 1);
		document.getElementById("P").style.display = "";
	});
}

async function SetAttrib(item, n, s) {
	if (/^!/.test(n)) {
		n = n.slice(1);
		s = !s;
	}
	if (window.SetXmlAttr) {
		s = await SetXmlAttr(item, n, s);
	}
	if (n == "TextContent") {
		if ("string" === typeof (await item.text)) {
			item.text = s;
		} else {
			item.textContent = s;
		}
	} else if (s) {
		item.setAttribute(n, s);
	} else {
		item.removeAttribute(n);
	}
}

function GetElementValue(o) {
	if (o.type) {
		if (/checkbox/i.test(o.type)) {
			return o.checked ? 1 : 0;
		}
		if (/hidden|text|number|url|password|range|color|date|time|search/i.test(o.type)) {
			return o.value;
		}
		if (/select/i.test(o.type)) {
			return o[o.selectedIndex].value;
		}
	}
}

function SetElementValue(o, s) {
	if (o && o.type) {
		if (/checkbox/i.test(o.type)) {
			o.checked = GetNum(s);
			return;
		}
		if (/select/i.test(o.type)) {
			for (let i = o.options.length; i-- > 0;) {
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

async function SetAttribEx(item, f, n) {
	const s = GetElementValue(f[n]);
	if (s != await GetAttribEx(item, f, n)) {
		await SetAttrib(item, n, s);
		return true;
	}
	return false;
}

async function GetAttribEx(item, f, n) {
	if (!n) {
		return;
	}
	let s;
	const res = /([^=]*)=(.*)/.exec(n);
	if (res) {
		s = await item.getAttribute(res[1]);
		if (s == res[2]) {
			document.getElementById(n).checked = true;
		}
		return;
	}
	if (n == "TextContent") {
		s = await item.text;
		if ("string" !== typeof s) {
			s = item.textContent;
		}
	} else {
		s = /^!/.test(n) ? !await item.getAttribute(n.slice(1)) : await item.getAttribute(n);
	}
	if (window.GetXmlAttr) {
		s = await GetXmlAttr(item, n, s);
	}
	if (s || s === 0) {
		if (n == "Key") {
			s = await GetKeyName(s);
		}
		SetElementValue(f[n], s);
	}
}

function RefX(IdList, bMultiLine, oButton, bFilesOnly, Filter, f) {
	setTimeout(async function () {
		const ar = IdList.split(";");
		const Id = ar[0];
		const s = (ar.length > 1) ? ar[1] : Id.replace("Path", "Type");
		const o = GetElement(Id, f);
		if ((/Path/.test(Id) || (ar.length > 1)) && "string" !== typeof Filter) {
			if (o) {
				let pt;
				if (oButton) {
					pt = await api.CreateObject("Object");
					const pt1 = GetPos(oButton, 9);
					pt.x = pt1.x;
					pt.y = pt1.y;
					pt.width = oButton.offsetWidth;
				} else {
					pt = await api.GetCursorPos();
				}
				const oType = o.form[s];
				const optId = oType ? oType[oType.selectedIndex].value : "exec";
				const r = await MainWindow.OptionRef(optId, o.value, pt);
				if ("string" === typeof r) {
					const ps = await MainWindow.OptionDecode(optId, r);
					if (bMultiLine && await api.GetKeyState(VK_CONTROL) < 0 && await api.ILCreateFromPath(ps)) {
						AddPath(Id, ps, f);
					} else {
						SetValue(o, ps);
					}
				}
				return;
			}
		}

		let path = o.value || o.getAttribute("placeholder") || "";
		const res = /^icon:([^,]*)|^bitmap:([^,]*)/i.exec(path) || [];
		path = await OpenDialogEx(res[1] || res[2] || path, Filter, GetNum(bFilesOnly));
		if (path) {
			if (bMultiLine) {
				AddPath(Id, path);
				return;
			}
			SetValue(o, path);
			if (/Icon|Large|Small/i.test(Id)) {
				const s = await ExtractPath(te, path);
				if (await api.ExtractIconEx(s, -1, null, null, 0) > 1) {
					ShowDialogEx("fileicon", 640, 480, [GetElementIdEx(o), o.value].join("\t"));
					return;
				}
			}
		}
	}, 99);
}

function PortableX(Id) {
	ConfirmThenExec(GetText("Portable"), async function () {
		const o = GetElement(Id);
		const s = await fso.GetDriveName(ui_.TEPath);
		SetValue(o, o.value.replace(await wsh.ExpandEnvironmentStrings("%UserProfile%"), "%UserProfile%").replace(new RegExp('^("?)' + s, "igm"), "$1%Installed%").replace(new RegExp('( "?)' + s, "igm"), "$1%Installed%").replace(new RegExp('(:)' + s, "igm"), "$1%Installed%"));
	});
}

function GetElement(Id, o) {
	return (o && o[Id]) || (document.F && document.F[Id]) || document.getElementById(Id) || (document.E && document.E[Id]);
}

function AddPath(Id, strValue, f) {
	const o = GetElement(Id, f);
	if (o) {
		let s = o.value;
		if (/\n$/.test(s) || s == "") {
			s += strValue;
		} else {
			s += "\n" + strValue;
		}
		SetValue(o, s);
	}
}

function SetValue(el, s) {
	if (el && el.value != s) {
		el.value = s;
		FireEvent(el, "change");
	}
}

SetElement = async function (Id, v) {
	SetValue(GetElementEx(Id), v);
}

async function GetCurrentSetting(s) {
	const FV = await te.Ctrl(CTRL_FV);

	if (await confirmOk()) {
		AddPath(s, PathQuoteSpaces(await api.GetDisplayNameOf(await FV.FolderItem, SHGDN_FORPARSINGEX | SHGDN_FORPARSING)));
	}
}

async function SetTab(s) {
	let o = null;
	const arg = String(s).split(/&/);
	for (let i in arg) {
		const ar = arg[i].split(/=/);
		if (SameText(ar[0], "tab")) {
			if (SameText(ar[1], "Get Addons")) {
				o = document.getElementById('tab1_1');
			}
			if (SameText(ar[1], "Get Icons")) {
				o = document.getElementById('tab1_3');
			}
			if (SameText(ar[1], "Get Lang")) {
				o = document.getElementById('tab1_4');
			}
			const s = await GetText(ar[1]);
			let ovTab;
			for (let j = 0; ovTab = document.getElementById('tab' + j); ++j) {
				if (SameText(s, ovTab.innerText)) {
					o = ovTab;
					break;
				}
			}
		} else if (SameText(ar[0], "menus")) {
			g_MenuType = ar[1];
		} else if (SameText(ar[0], "id")) {
			g_Id = ar[1];
		}
	}
	ClickTree(o);
}

function AddMouse(o) {
	(document.F["MouseMouse"] || document.F["Mouse"]).value += o.title;
}

async function InitAddonOptions(bFlag) {
	returnValue = false;
	await LoadLang2(BuildPath(ui_.Installed, "addons", Addon_Id, "lang", await GetLangId() + ".xml"));
	await ApplyLang(document);
	document.title = await GetAddonInfo(Addon_Id).Name;
	SetOnChangeHandler();
	IsChanged = function () {
		return g_bChanged || g_Chg.Data;
	};
	TEOk = SetAddonOptions;
	if (!await WebBrowser.OnClose) {
		WebBrowser.OnClose = async function (WB) {
			await SetOptions(TEOk, null, ContinueOptions);
			CloseWB(WB);
		};
	}
	const items = await te.Data.Addons.getElementsByTagName(Addon_Id);
	if (GetLength(items)) {
		InitColor1(await items[0]);
	}
}

function SetOnChangeHandler() {
	g_nResult = 3;
	g_bChanged = false;
	const ar = ["input", "select", "textarea"];
	for (let j in ar) {
		const o = document.getElementsByTagName(ar[j]);
		if (o) {
			for (let i = o.length; i--;) {
				if ((o[i].name || o[i].id) && o[i].name != "List" && !/^_/.test(o[i].id)) {
					AddEventEx(o[i], "change", function (ev) {
						ev = ev || event;
						g_bChanged = true;
						const target = ev.target || ev.srcElement;
						if (target) {
							const res = /^(Tab|Tree|View|Conf)/.exec(target.name || target.id);
							if (res) {
								g_Chg[res[1]] = true;
							}
						}
					});
				}
			}
		}
	}
}

async function SetAddonOptions() {
	if (!g_bClosed) {
		const items = await te.Data.Addons.getElementsByTagName(Addon_Id);
		if (GetLength(items)) {
			const item = await items[0];
			const el = document.F;
			for (let i = el.length; i--;) {
				const n = el[i].id || el[i].name;
				if (n) {
					if (await SetAttribEx(item, document.F, n)) {
						returnValue = true;
					}
				}
			}
		}
		g_bClosed = true;
		if (returnValue) {
			TEOk();
		}
		CloseWindow();
	}
}

SelectIcon = function (o) {
	o = o.target || o.srcElement || o;
	returnValue = o.title;
	document.F.ButtonOk.disabled = false;
	document.F.ButtonOk.focus();
	document.getElementById("Selected").innerHTML = o.outerHTML;
	DialogResize();
}

TestX = function (id, f) {
	ConfirmThenExec(GetText("Test"), async function () {
		if (!f) {
			f = document.F;
		}
		const o = f[id + "Type"];
		await MainWindow.InvokeUI("window.focus");
		await MainWindow.Exec(await te.Ctrl(CTRL_FV), await MainWindow.OptionEncode(o[o.selectedIndex].value, f[id + "Path"].value), o[o.selectedIndex].value);
		focus();
	});
}

SetImage = async function (f, n) {
	const o = document.getElementById(n || "_Icon");
	if (o) {
		if (!f) {
			f = document.F;
		}
		const px = screen.deviceYDPI / 3;
		const h = Math.min(GetNum(f.IconSize || f.Height) || ui_.IconSize, px);
		o.innerHTML = await GetImgTag({ src: await ExtractPath(te, f.Icon.value), "max-width": px + "px", "max-height": px + "px" }, h);
	}
}

ShowIcon = ShowIconEx;

async function SelectLangID(o) {
	const el = document.activeElement;
	if (!el || el == ui_.elConfirm && ui_.ConfirmMenu && (new Date().getTime() - ui_.ConfirmMenu) < 500) {
		return;
	}
	const Langs = await GetFileList(BuildPath(ui_.Installed, "lang", "*.xml"), false, window.chrome);
	Langs.sort(function (a, b) {
		return a.localeCompare(b);
	});
	const path = BuildPath(ui_.Installed, "lang");
	const hMenu = await api.CreatePopupMenu();
	for (let i in Langs) {
		const xml = await api.CreateObject("Msxml2.DOMDocument");
		xml.async = false;
		let title = Langs[i];
		await xml.load(BuildPath(path, title));
		Langs[i] = title = title.replace(/\..*$/, "");
		const items = await xml.getElementsByTagName('lang');
		if (items && await GetLength(items)) {
			const item = await items[0];
			let en = await item.getAttribute("en");
			en = (en && SameText(await item.text, en)) ? "" : ' / ' + en;
			title = await item.text + en + " (" + title + ")\t" + await item.getAttribute("author");
		}
		await api.InsertMenu(hMenu, i, MF_BYPOSITION | MF_STRING, GetNum(i) + 1, title);
	}
	const pt = GetPos(o, 9);
	const nVerb = await api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RIGHTALIGN | TPM_RETURNCMD, pt.x + o.offsetWidth, pt.y, ui_.hwnd, null, null);
	if (nVerb) {
		document.F.Conf_Lang.value = Langs[nVerb - 1];
		g_bChanged = true;
	} else {
		ui_.ConfirmMenu = new Date().getTime();
		ui_.elConfirm = el;
	}
	api.DestroyMenu(hMenu);
}

async function GetTextEx(s) {
	const ar = s.split(/_/);
	s = await GetText(ar.shift());
	if (ar && ar.length) {
		s += "(" + await GetText(ar.join(" ")) + ")";
	}
	return s;
}

async function GetAddons() {
	OpenHttpRequest(urlAddons + "index.xml", "http", AddonsList);
}

function GetIconPacks() {
	OpenHttpRequest(urlIcons + "index.json", "http", IconPacksList);
}

function GetLangPacks() {
	OpenHttpRequest(urlLang + "index.json", "http", LangPacksList);
}

async function InstallLang(el) {
	const arg = await api.CreateObject("Object");
	arg.responseType = "blob";
	arg.title = el.title
	OpenHttpRequest(urlLang + el.title.replace(/\n.*/, ""), "http", InstallLang2, arg);
}

function UpdateAddon(Id, o) {
	if (!o) {
		AddAddon(document.getElementById("Addons"), Id, "Disable");
		g_Chg.Addons = true;
	}
}

async function CheckAddon(Id) {
	return fso.FileExists(BuildPath(ui_.Installed, "addons", Id, "config.xml"));
}

function AddonsSearch() {
	if (g_SortMode == 1) {
		SortAddons1(g_nSort[1]);
		document.getElementById("_AddonsQ").focus();
	} else if (g_SortMode == "1_1") {
		if (xmlAddons) {
			AddonsAppend(xmlAddons.getElementsByTagName("Item"))
		} else {
			GetAddons();
		}
	}
	return true;
}

function ClearSearch(el, Id) {
	const elText = document.getElementById(Id);
	if (elText.value) {
		ConfirmThenExec(GetTextR("@dinput8.dll,-6300"), async function () {
			elText.value = "";
			FireEvent(el, "click");
		});
	}
}

async function AddonsList(xhr2) {
	if (xmlAddons) {
		return;
	}
	if (await xhr2) {
		xhr = await xhr2;
	}
	let xml = await xhr.responseXML;
	if (!xml) {
		if (window.DOMParser) {
			xml = new DOMParser().parseFromString(await xhr.get_responseText(), "application/xml");
		} else {
			xml = api.CreateObject("Msxml2.DOMDocument");
			xml.async = false;
			xml.loadXML(xhr.get_responseText());
		}
	}
	if (xml) {
		xmlAddons = xml;
		AddonsAppend(xml.getElementsByTagName("Item"));
	}
}

function SetTable(table, td) {
	if (table) {
		while (table.rows.length > 0) {
			table.deleteRow(0);
		}
		for (let i = 0; i < td.length; ++i) {
			const tr = table.insertRow(i);
			const td1 = tr.insertCell(0);
			td[i].shift();
			td1.innerHTML = td[i].join("");
			ApplyLang(td1);
			td1.className = (i & 1) ? "box oddline" : "box";
			td1.style.padding = "3px";
		}
	}
}

async function AddonsAppend(Items) {
	const Progress = await api.CreateObject("ProgressDialog");
	const td = [];
	Progress.StartProgressDialog(ui_.hwnd, null, 2);
	try {
		if (!ui_.nTEVer) {
			ui_.nTEVer = await AboutTE(0);
		}
		Progress.SetAnimation(hShell32, 150);
		Progress.SetLine(1, await api.LoadString(hShell32, 13585) || await api.LoadString(hShell32, 6478), true);
		let bCancelled = false;
		for (let i = 0; !bCancelled && i < Items.length; ++i) {
			Promise.all([Progress.HasUserCancelled(i, Items.length, 2)]).then(function (r) {
				bCancelled = r[0];
			});
			await ArrangeAddon(Items[i], td);
		}
		td.sort(function (a, b) {
			return a[0] - b[0] || String(a[0]).localeCompare(b[0]);
		});
		SetTable(document.getElementById("Addons1"), td);
	} catch (e) {
		ShowError(e);
	}
	Progress.StopProgressDialog();
}

async function ArrangeAddon(xml, td) {
	const Id = xml.getAttribute("Id");
	const s = [];
	let strUpdate = "";
	const q = document.getElementById("_GetAddonsQ").value.toLowerCase();
	if (!q || Id.toLowerCase().indexOf(q) >= 0 || await Search(xml, q)) {
		const info = {};
		for (let i = arLangs.length; i--;) {
			await GetAddonInfo2(xml, info, arLangs[i]);
		}
		const dt = new Date(info.pubDate);
		s.push('<table width="100%"><tr><td width="100%"><b style="font-size: 1.3em">', info.Name, "</b>&nbsp;");
		s.push(info.Version, "&nbsp;");
		s.push(info.Creator, "<br>")
		s.push(info.Description, "<br>");
		if (info.Details) {
			if (!ui_.strDetails) {
				ui_.strDetails = await GetText('Details');
			}
			s.push('<a href="', info.Details, '" title="', info.Details, '" class="link" onclick="wsh.run(this.title); return false;">', ui_.strDetails, '</a><br>');
		}
		s.push(dt.toLocaleDateString(), '</td><td align="right">');
		let dt2 = (dt.getTime() / (24 * 60 * 60 * 1000)) - info.Version;
		let bUpdate = false;
		if (await CheckAddon(Id)) {
			const installed = GetAddonInfo(Id);
			if (await installed.Version >= info.Version) {
				try {
					if (!await installed.DllVersion) {
						return;
					}
					const path = BuildPath(ui_.Installed, "addons", Id);
					const cFileName = await IsExists(BuildPath(path, "*" + ui_.bit + ".dll"), "cFileName");
					if (!cFileName) {
						return;
					}
					if (CalcVersion(await installed.DllVersion) <= CalcVersion(await fso.GetFileVersion(BuildPath(path, cFileName)))) {
						return;
					}
				} catch (e) {
					return;
				}
			}
			strUpdate = '<br><b id="_Addons_' + Id + '"  class="danger nowrap">' + await GetText('Update available') + '</b>';
			dt2 += MAXINT * 2;
			bUpdate = true;
		} else {
			dt2 += MAXINT;
		}
		if (info.MinVersion && ui_.nTEVer >= CalcVersion(info.MinVersion)) {
			if (!ui_.strInstall) {
				ui_.strInstall = await GetText("Install");
			}
			s.push('<button onclick="Install(this,', bUpdate, ')" title="', Id, '_', info.Version, '">', ui_.strInstall, '</button>');
		} else {
			if (!ui_.strVersion) {
				ui_.strVersion = await GetTextR("@mstask.dll,-319");
			}
			if (!ui_.strIsRequired) {
				ui_.strIsRequired = await GetText("%s is required.");
			}
			s.push('<button class="danger" onclick="MainWindow.CheckUpdate()">', ui_.strIsRequired.replace("%s", ui_.strVersion + " " + (info.MinVersion.replace(/^20/, "").replace(/\.0/g, '.'))), '</button>');
		}
		s.push(strUpdate, '</td></tr></table>');
		s.unshift(g_nSort["1_1"] == 1 ? -dt2 : g_nSort["1_1"] ? Id : info.Name);
		td.push(s);
	}
}

async function Search(xml, q) {
	for (let k = arLangs.length; k-- > 0;) {
		const items = xml.getElementsByTagName(arLangs[k]);
		if (items.length) {
			const item = items[0].childNodes;
			for (let i = item.length; i-- > 0;) {
				const item1 = item[i];
				if (/Name|Description/.test(item1.tagName)) {
					const s = item1.text || item1.textContent;
					if (s.toLowerCase().indexOf(q) >= 0) {
						return true;
					}
					if (arLangs[k] == "en" && (await GetAltText(s, true)).toLowerCase().indexOf(q) >= 0) {
						return true;
					}
				}
			}
		}
	}
	return false;
}

async function Install(o, bUpdate) {
	if (!bUpdate && !await confirmOk("Do you want to install it now?")) {
		return;
	}
	const Id = o.title.replace(/_.*$/, "");
	await MainWindow.SaveConfig();
	await MainWindow.AddonDisabled(Id);
	if (await AddonBeforeRemove(Id) < 0) {
		return;
	}
	OpenHttpRequest(urlAddons + Id + '/' + o.title.replace(/\./, "") + '.zip', "http", Install2, o);
}

async function Install2(xhr, url, o) {
	const Id = o.title.replace(/_.*$/, "");
	const file = o.title.replace(/\./, "") + '.zip';
	const temp = await GetTempPath(3);
	await CreateFolder(temp);
	const dest = BuildPath(temp, Id);
	await DeleteItem(dest);
	const hr = await Extract(BuildPath(temp, file), temp, xhr);
	if (hr) {
		ShowExtractError(hr, file);
		return;
	}
	const configxml = dest + "\\config.xml";
	let nDog = 300;
	while (!await fso.FileExists(configxml)) {
		if (await wsh.Popup(await GetText("Please wait."), 1, TITLE, MB_ICONINFORMATION | MB_OKCANCEL) == IDCANCEL || nDog-- == 0) {
			return;
		}
	}
	await api.SHFileOperation(FO_MOVE, dest, BuildPath(ui_.Installed, "addons"), FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR, false);
	o.disabled = true;
	if (!ui_.strInstalled) {
		ui_.strInstalled = await GetText("Installed");
	}
	o.value = ui_.strInstalled;
	o = document.getElementById('_Addons_' + Id);
	if (o) {
		o.style.display = "none";
	}
	UpdateAddon(Id, o);
}

async function InstallIcon(o) {
	if (!await confirmOk("Do you want to install it now?")) {
		return;
	}
	const Id = o.title.replace(/_[^_]*$/, "");
	OpenHttpRequest(urlIcons + Id + '/' + o.title.replace(/\./g, "") + '.zip', "http", InstallIcon2, o);
}

async function InstallIcon2(xhr, url, o) {
	const file = o.title.replace(/\./, "") + '.zip';
	const temp = await GetTempPath(3);
	await CreateFolder(temp);
	const dest = BuildPath(await te.Data.DataFolder, "icons");
	await CreateFolder(dest);
	const hr = await Extract(BuildPath(temp, file), dest, xhr);
	if (hr) {
		ShowExtractError(hr, file);
		return;
	}
	IconPacksList();
}

function JsonSearch(info, q) {
	if (!q) {
		return true;
	}
	let s = "";
	for (let i in info) {
		s += info[i] + "\n";
	}
	return s.toLowerCase().indexOf(q) >= 0
}

async function IconPacksList1(s, Id, info, json) {
	const q = document.getElementById('_GetIconsQ').value.toLowerCase();
	if (!json && !JsonSearch(info, q)) {
		return false;
	}
	const langId = await GetLangId();
	s.push('<img src="', urlIcons, Id, '/preview.png" align="left" style="margin-right: 8px"><b style="font-size: 1.3em">', info.name[langId] || info.name.en, '</b> ');
	s.push(info.version, " ");
	if (info.URL) {
		s.push('<a href="', info.URL, '" title="', info.URL, '" class="link" onclick="wsh.run(this.title); return false;">');
	}
	s.push(info.creator[langId] || info.creator.en);
	if (info.URL) {
		s.push('</a>');
	}
	s.push("<br>", info.descprition[langId] || info.descprition.en, "<br>");
	if (json) {
		if (!ui_.strInstalled) {
			ui_.strInstalled = await GetText("Installed");
		}
		s.push(ui_.strInstalled, '<button onclick="DeleteIconPacks()" style="float: right;">Delete</button>');
		if (json[Id] && Number(json[Id].info.version) > Number(info.version)) {
			s.push('<hr><b class="danger nowrap">', await GetText('Update available'), '</b> ', json[Id].info.version);
			info = json[Id].info;
			json = false;
		}
	}
	if (!json) {
		if (!ui_.strInstall) {
			ui_.strInstall = await GetText("Install");
		}
		s.push('<button onclick="InstallIcon(this)" title="', Id, '_', info.version, '" style="float: right;">', ui_.strInstall, '</button>');
	}
	s.push("<br>", new Date(info.pubDate).toLocaleDateString());
	return true;
}

async function IconPacksList(xhr) {
	if (xhr) {
		g_xhrIcons = xhr;
	} else {
		xhr = window.g_xhrIcons;
	}
	if (!xhr) {
		return;
	}
	let s = await ReadTextFile(BuildPath(await te.Data.DataFolder, "icons\\config.json"));
	const json1 = JSON.parse(s || '{}');
	const json = JSON.parse(await xhr.get_responseText ? await xhr.get_responseText() : xhr.responseText);
	const td = [];
	let Installed = "";
	if (json1.info) {
		Installed = json1.info.id || json1.info.name.en.replace(/\W/g, "_");
	}
	const LangId = await GetLangId();
	for (let n in json) {
		if (n != Installed) {
			let s1;
			const s = [];
			info = json[n].info;
			if (await IconPacksList1(s, n, info)) {
				if (g_nSort["1_3"] == 0) {
					s1 = info.name[LangId] || info.name.en;
				} else if (g_nSort["1_3"] == 1) {
					if (json.pubDate) {
						s1 = ("0000000" + (99999999 - Math.floor(new Date(json.pubDate).getTime() / 86400000))).slice(-8);
					}
				} else {
					s1 = n;
				}
				td.push([s1 + "\t" + n, s.join("")]);
			}
		}
	}
	td.sort(function (a, b) {
		a[0].localeCompare(b[0]);
	});
	if (json1.info) {
		s = [];
		if (await IconPacksList1(s, Installed, json1.info, json)) {
			td.unshift(["", s.join("")]);
		}
	}
	SetTable(document.getElementById("IconPacks1"), td);
}

async function DeleteIconPacks() {
	if (await confirmOk()) {
		await DeleteItem(BuildPath(await te.Data.DataFolder, "icons"));
		IconPacksList();
	}
}

async function LangPacksList(xhr) {
	if (xhr) {
		g_xhrLang = xhr;
	} else {
		xhr = window.g_xhrLang;
	}
	if (!xhr) {
		return;
	}
	const text = await xhr.get_responseText ? await xhr.get_responseText() : xhr.responseText;
	const json = JSON.parse(text);
	const td = [];
	const bt = await GetText("Install");
	const q = document.getElementById('_GetLangQ').value.toLowerCase();
	for (let n in json) {
		const info = json[n];
		if (JsonSearch(info, q)) {
			const tm = new Date(info.pubDate).getTime();
			let strUpdate = "";
			const ftLastWriteTime = await IsExists(BuildPath(ui_.Installed, "lang", n), "ftLastWriteTime");
			if (ftLastWriteTime) {
				if (tm < new Date(ftLastWriteTime).getTime() + 2000) {
					continue;
				}
				strUpdate = ['<b class="danger nowrap" style="float: right">', await GetText('Update available'), '&nbsp;&nbsp;</b>'].join("");
			}
			const ar = [tm, '<b style="font-size: 1.3em">', info.name, " / ", info.en, "</b><br>", info.author, '<button onclick="InstallLang(this)" title="', n, "\n", info.pubDate, '" style="float: right">', bt, '</button>', strUpdate, '<br>', new Date(tm).toLocaleDateString()];
			if (info.size) {
				ar.push("<br>", (info.size / 1024).toFixed(1), " KB");
			}
			if (strUpdate) {
				td.unshift(ar);
			} else {
				td.push(ar);
			}
		}
	}
	SetTable(document.getElementById("LangPacks1"), td);
}

async function InstallLang2(xhr, url, arg) {
	if (!xhr) {
		return;
	}
	const temp = await GetTempPath(3);
	await CreateFolder(temp);
	const ar = (await arg.title).split("\n");
	const src = BuildPath(temp, ar[0]);
	await DownloadFile(xhr, src);
	await SetFileTime(src, null, null, new Date(ar[1]).getTime());
	await api.SHFileOperation(FO_MOVE, src, BuildPath(ui_.Installed, "lang"), FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR, false);
	document.F.Conf_Lang.value = ar[0].replace(/\.xml$/i, "");
	g_bChanged = true;
	await MessageBox("Completed.", TITLE, MB_OK);
	ClickTree("tab0");
}

function repeat(strIn, nTimes) {
	let strOut = "";
	for (let n = 0; n < nTimes; ++n) {
		strOut += strIn;
	}
	return strOut;
}

function IndentMenuItems(o) {
	if (o && o.hasAttribute("data-menu")) {
		const INDENT = String.fromCharCode(8195);	//EMSP
		const MenusOpen  = new RegExp(g_sep + "Open"  + g_sep + "Menus" + g_sep, "i");
		const MenusClose = new RegExp(g_sep + "Close" + g_sep + "Menus" + g_sep, "i");
		let nLevel = 0;
		for (let i = 0; i < o.length; ++i) {
			o[i].style.color = "";
			if (MenusClose.test(o[i].value)) {
				if (nLevel) {
					--nLevel;
				} else {
					o[i].style.color = "red";
				}
			}
			o[i].label = repeat(INDENT, nLevel) + o[i].text;
			if (MenusOpen.test(o[i].value)) {
				++nLevel;
			}
		}
	}
}

async function EnableSelectTag(o) {
	await IndentMenuItems(o);
	if (o && !window.chrome) {
		const hwnd = await WebBrowser.hwnd;
		api.SendMessage(hwnd, WM_SETREDRAW, false, 0);
		o.style.visibility = "hidden";
		setTimeout(async function () {
			o.style.visibility = "visible";
			await api.SendMessage(hwnd, WM_SETREDRAW, true, 0);
			api.RedrawWindow(hwnd, null, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
		}, 99);
	}
}

SetResult = function (i) {
	g_nResult = i;
	CloseWindow();
}

function InitColor1(item) {
	const el = document.F;
	for (let i = el.length; i--;) {
		const n = el[i].id || el[i].name;
		if (n) {
			GetAttribEx(item, document.F, n);
		}
	}
	for (let i = el.length; i--;) {
		const n = el[i].id || el[i].name;
		if (n) {
			const res = /^Color_(.*)/.exec(n);
			if (res) {
				const o = document.F[res[1]];
				if (o) {
					el[i].style.backgroundColor = GetWebColor(o.value || o.placeholder);
				}
			}
		}
	}
}

function ChangeColor1(el) {
	const o = document.getElementById("Color_" + (el.id || el.name));
	if (o) {
		o.style.backgroundColor = GetWebColor(el.value || el.placeholder);
	}
}

function EnableInner() {
	ChangeForm([
		["__Inner", "disabled", false],
		["__InnerRight", "disabled", false],
		["__InnerBottom", "disabled", false],
		["__InnerBottomRight", "disabled", false]
	]);
}

function ChangeForm(ar) {
	const fn = function () {
		for (let i in ar) {
			let o = document.getElementById(ar[i][0]);
			if (o) {
				let s = ar[i][1].split("/");
				while (s.length > 1) {
					o = o[s.shift()];
				}
				o[s[0]] = ar[i][2];
			}
		}
	}
	window.addEventListener("load", fn);
	fn();
}

function SetTabContents(id, name, value) {
	if (name) {
		Promise.all([GetText(name)]).then(function (r) {
			document.getElementById("tab" + id).innerHTML = r[0];
		});
	}
	document.getElementById("panel" + id).innerHTML = (value.join ? value.join("") : value).replace(/<input type="button"([^>]*?)value="([^"]*)"\s?([^>]*)>/ig, "<button$1$3>$2</button>");
}

async function ShowButtons(b1, b2, SortMode) {
	if (SortMode) {
		g_SortMode = SortMode;
	}
	const o = document.getElementById("SortButton");
	o.style.display = b1 ? "inline-block" : "none";
	if (!document.getElementById("SortButton_0")) {
		const h = screen.deviceYDPI / 4;
		o.innerHTML = await GetImgTag({
			id: "SortButton_0",
			src: "icon:general,24",
			title: await GetTextR("@wmploc.dll,-29080"),
			onclick: "SortAddons(this)",
			"class": "button"
		}, h);
	}
	if (g_SortMode == 1) {
		const table = document.getElementById("Addons");
		const bSorted = /none/i.test(table.style.display);
		document.getElementById("MoveButton").style.display = (b1 || b2) && !bSorted ? "inline-block" : "none";
	} else {
		document.getElementById("MoveButton").style.display = b2 ? "inline-block" : "none";
	}
}

async function SortAddons(el) {
	if (el == ui_.elConfirm && ui_.ConfirmMenu && (new Date().getTime() - ui_.ConfirmMenu) < 500) {
		return;
	}
	let ar = [8976, 8980, 8977];
	if (g_SortMode == 1) {
		ar.push(9808);
	}
	for (let i = ar.length; i--;) {
		ar[i] = api.LoadString(hShell32, ar[i]);
	}
	ar.push(api.CreatePopupMenu());
	ar = await Promise.all(ar);
	const hMenu = ar.pop();
	for (let i = 0; i < ar.length; ++i) {
		let uFlags = MF_BYPOSITION | MF_STRING;
		if (g_SortMode == 1) {
			const table = document.getElementById("Addons");
			if (/none/i.test(table.style.display)) {
				if (g_nSort[1] == i) {
					uFlags |= MF_CHECKED | MFT_RADIOCHECK;
				}
			} else if (i == 3) {
				uFlags |= MF_CHECKED | MFT_RADIOCHECK;
			}
		} else {
			if (g_nSort[g_SortMode] == i) {
				uFlags |= MF_CHECKED | MFT_RADIOCHECK;
			}
		}
		await api.InsertMenu(hMenu, i, uFlags, i + 1, ar[i]);
	}
	const pt = GetPos(el, 1);
	const n = await api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y + o.offsetHeight, ui_.hwnd, null, null) - 1;
	api.DestroyMenu(hMenu);
	if (n >= 0) {
		SortAddons1(n);
	} else {
		ui_.ConfirmMenu = new Date().getTime();
		ui_.elConfirm = el;
	}
}

async function SortAddons1(n) {
	if ("number" !== typeof n) {
		n = 3;
	}
	if (g_SortMode == 1) {
		const table = document.getElementById("Addons");
		if (table.rows.length < 2) {
			return;
		}
		const sorted = document.getElementById("SortedAddons");
		while (sorted.rows.length > 0) {
			sorted.deleteRow(0);
		}
		const q = document.getElementById("_AddonsQ").value.toLowerCase();
		if (n == 3 && !q) {
			table.style.display = "";
			sorted.style.display = "none";
		} else {
			let s, ar = [], bCancelled = false, nTotal = table.rows.length;
			const Progress = await api.CreateObject("ProgressDialog");
			try {
				Progress.SetAnimation(hShell32, 150);
				Progress.SetLine(1, await GetTextR("@shell32.dll,-13585[-6478]"), true);
				Progress.StartProgressDialog(ui_.hwnd, null, 2);
				for (let j = 0; !bCancelled && j < table.rows.length; ++j) {
					Promise.all([Progress.HasUserCancelled(j, nTotal, 2)]).then(function (r) {
						bCancelled = r[0];
					});
					const div = table.rows[j].cells[0].firstChild || {};
					const Id = (div.id || "").replace("Addons_", "").toLowerCase(); if (Id) {
						if (!q || Id.indexOf(q) >= 0 || await FindAddonInfo(Id, q)) {
							if (n == 0) {
								s = table.rows[j].cells[0].innerText;
							} else if (n == 1) {
								s = "";
								const pubDate = await GetAddonInfo(Id).pubDate;
								if (pubDate) {
									s = ("0000000" + (99999999 - Math.floor(new Date(pubDate).getTime() / 86400000))).slice(-8);
								}
							} else {
								s = Id;
							}
							ar.push(s + "\t" + Id);
							++nTotal;
						}
					}
				}
				if (n != 3) {
					ar.sort(function (a, b) {
						return a.localeCompare(b);
					});
				}
				for (let i = 0; !bCancelled && i < ar.length; ++i) {
					Promise.all([Progress.HasUserCancelled(table.rows.length + i, nTotal, 2)]).then(function (r) {
						bCancelled = r[0];
					});
					const data = ar[i].split("\t");
					const Id = data[data.length - 1];
					AddAddon(sorted, Id, document.getElementById("enable_" + Id).checked, "Sorted_");
				}
			} catch (e) {
				ShowError(e);
			}
			Progress.StopProgressDialog();
			table.style.display = bCancelled ? "" : "none";
			sorted.style.display = bCancelled ? "none" : "";
		}
		if (g_nSort[1] !== n) {
			g_nSort[1] = n;
			te.Data.Conf_AddonsSort = (n + 1) % 4;
			g_bChanged = true;
		}
	} else if (g_SortMode == "1_1") {
		if (n != g_nSort["1_1"]) {
			g_nSort["1_1"] = n;
			AddonsSearch();
		}
	} else if (g_SortMode == "1_3") {
		if (n != g_nSort["1_3"]) {
			g_nSort["1_3"] = n;
			IconPacksList();
		}
	}
	ShowButtons(true);
}
