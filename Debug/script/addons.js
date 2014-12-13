var urlAddons = "http://www.eonet.ne.jp/~gakana/tablacus/addons/";
nCount = 0;
returnValue = false;
xhr = null;
xmlAddons = null;
g_tid = null;

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
	var q =
	{
		td: [],
		ts: [],
		i: 0
	}
	if (AddonsSub(q)) {
		document.F.b.disabled = true;
		document.body.style.cursor = "wait";
		document.getElementById('c').style.display = "inline";
		AddonsAppend(q)
	}
	return true;
}

function AddonsSub(q)
{
	if (document.getElementById) {
		var table = document.getElementById("addons");
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

	xmlAddons = xhr.responseXML.getElementsByTagName("Item");
	if (AddonsSub(q)) {
		document.F.b.disabled = true;
		document.body.style.cursor = "wait";
		document.getElementById('c').style.display = "inline";
		AddonsAppend(q)
	}
}

function AddonsAppend(q)
{
	if (xmlAddons[q.i]) {
		ArrangeAddon(xmlAddons[q.i++], q.td, q.ts);
		document.getElementById('c').innerText = Math.floor(100 * q.i / q.ts.length) + " %";

		g_tid = setTimeout(function () {
			AddonsAppend(q);
		}, 1);
	}
	else {
		document.F.b.disabled = false;
		document.body.style.cursor = "auto";
		document.getElementById('c').style.display = "none";
	}
}

function ArrangeAddon(xml, td, ts)
{
	var Id = xml.getAttribute("Id");
	var s = [];
	if (Search(xml)) {
		var info = [];
		GetAddonInfo2(xml, info, "General");
		GetAddonInfo2(xml, info, "en");
		GetAddonInfo2(xml, info, GetLangId());

		var pubDate = "";
		var dt = new Date(info.pubDate);
		if (info.pubDate) {
			pubDate = dt.toLocaleString() + " ";
		}
		s.push('<b>', info.Name, "</b>&nbsp;", info.Version, "&nbsp;", info.Creator, "<br>", info.Description, "<br>");

		s.push('<table width="100%"><tr><td>', pubDate, '</td><td align="right">');
		var filename = info.filename;
		if (!filename) {
			filename = Id + '_' + info.Version.replace(/\D/, '') + '.zip';
		}
		var dt2 = (dt.getTime() / (24 * 60 * 60 * 1000)) - info.Version;
		if (fso) {
			var bInstall = true;
			if (CheckAddon(Id)) {
				installed = GetAddonInfo(Id);
				if (installed.Version >= info.Version) {
					s.push(GetText('Installed'));
					bInstall = false;
				}
				else {
					s.push('<b id="_', Id,'" style="color: red">', GetText('Update available'), "</b>");
					dt2 += MAXINT * 2;
				}
			}
			else {
				dt2 += MAXINT;
			}
			if (bInstall) {
				if (info.MinVersion && te.Version >= CalcVersion(info.MinVersion)) {
					s.push('<input type="button" onclick="Install(this)" title="', Id, '_', info.Version, '" value="', GetText("Install"), '">');
				}
				else {
					s.push('<input type="button" style="color: red" onclick="CheckUpdate()" value="', info.MinVersion.replace(/^20/, "Version ").replace(/\.0/g, '.'), ' ', GetText("is required."), '">');
				}
			}
		}
		else {
			s.push('<a href="', Id, '/', filename, '">', 'Download', '</a>');
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

function GetAddonInfo2(xml, info, Tag)
{
	var items = xml.getElementsByTagName(Tag);
	if (items.length) {
		var item = items[0].childNodes;
		for (var i = 0; i < item.length; i++) {
			if (item[i].tagName) {
				if (item[i].textContent) {
					info[item[i].tagName] = item[i].textContent;
				}
				else {
					info[item[i].tagName] = item[i].text;
				}
			}
		}
	}
	return info;
}

function Search(xml)
{
	var q =  document.F.q.value.toUpperCase();
	if (q == "") {
		return true;
	}
	var Tags = ["General", "en", "es", "ja", "zh"];

	for (var k = 0; k < Tags.length; k++) {
		var items = xml.getElementsByTagName(Tags[k]);
		if (items.length) {
			var item = items[0].childNodes;
			for (var i = 0; i < item.length; i++) {
				if (item[i].tagName) {
					var s = '';
					if (item[i].textContent) {
						s = item[i].textContent + "";
					}
					else {
						s = item[i].text + "";
					}
					if (s.toUpperCase().match(q)) {
						return true;
					}
				}
			}
		}
	}
	return false;
}

function Install(o)
{
	if (!confirmOk("Do you want to install it now?")) {
		return;
	}
	document.body.style.cursor = "wait";
	setTimeout(function ()
	{
		var Id = o.title.replace(/_.*$/, "");
		var file = o.title.replace(/\./, "") + '.zip';
		var temp = fso.BuildPath(wsh.ExpandEnvironmentStrings("%TEMP%"), "tablacus");
		DeleteItem(temp);
		CreateFolder(temp);

		var zipfile = fso.BuildPath(temp, file);
		DownloadFile(urlAddons + Id + '/' + file, zipfile);
		if (MainWindow.Extract(zipfile, temp) != S_OK) {
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
			var dist = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\" + Id);
			CreateFolder(dist);
			var oDist = sha.NameSpace(dist);
			if (oDist) {
				oDist.MoveHere(Items, FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR);
				o.disabled = true;
				o.value = GetText("Installed");
				o = document.getElementById('_' + Id);
				if (o) {
					o.style.display = "none";
				}
				UpdateAddon(Id, o);
			}
		}
		document.body.style.cursor = "auto";
	}, 100);
}

function Resize()
{
	var w = document.documentElement.clientWidth || document.body.clientWidth;
	var h = document.documentElement.clientHeight || document.body.clientHeight;
	var o = document.getElementById("panel0");
	if (o) {
		var p = GetPos(o, false, true);
		h -= p.y;
		o.style.width = w + 'px';
		o.style.width = 2 * w - o.offsetWidth + "px";
		o.style.height = h + 'px';
		o.style.height = 2 * h - o.offsetHeight + "px";
	}
}

AddEventEx(window, "load", function ()
{
	ApplyLang(document);
	Resize();
	if (location.search.match(/q=(.*)/)) {
		document.F.q.value = decodeURI(RegExp.$1);
	}
	xhr = createHttpRequest();
	xhr.onreadystatechange = function()
	{
		if (xhr.readyState == 4) {
			if (xhr.status == 200) {
				setTimeout(AddonsList, 100);
			}
		}
	}
	xhr.open("GET", urlAddons + "/index.xml?" + Math.floor(new Date().getTime() / 60000));
	xhr.setRequestHeader('Pragma', 'no-cache');
	xhr.setRequestHeader('Cache-Control', 'no-cache');
	xhr.setRequestHeader('If-Modified-Since', 'Thu, 01 Jun 1970 00:00:00 GMT');
	xhr.send(null);
});

AddEventEx(window, "resize", Resize);
