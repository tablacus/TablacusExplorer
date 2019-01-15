var Addon_Id = 'download';
var item = GetAddonElement(Addon_Id);
if (window.Addon == 1) {
	Addons.Download = {
		Path: item.getAttribute("Path"),
		Show: api.LowPart(item.getAttribute("Visible")) ? SW_SHOWNORMAL : SW_HIDE
	}
	if (Addons.Download.Path) {
		createHttpRequest = function ()
		{
			return {
				open: function (method, url)
				{
					this.URL = url;
				},

				send: function ()
				{
					var temp = fso.BuildPath(fso.GetSpecialFolder(2).Path, "tablacus");
					CreateFolder(temp);
					this.fn = temp + "\\" + fso.GetTempName();
					DeleteItem(this.fn);
					wsh.Run(ExtractMacro(te, Addons.Download.Path.replace(/%url%/ig, this.URL).replace(/%file%/ig, this.fn)), Addons.Download.Show, true);
					this.readyState = 4;
					var wfd = api.Memory("WIN32_FIND_DATA");
					var hFind = api.FindFirstFile(this.fn, wfd);
					api.FindClose(hFind);
					this.status = (hFind != INVALID_HANDLE_VALUE) ? wfd.nFileSizeLow ? 200 : 403 : 404;
					if (this.status != 200) {
						DeleteItem(this.fn);
					}
					if (this.onreadystatechange) {
						this.onreadystatechange();
					}
				},

				get_responseText: function ()
				{
					var s;
					var ado = OpenAdodbFromTextFile(this.fn);
					if (ado) {
						s = ado.ReadText();
						ado.Close();
					}
					return s;
				},

				get_responseXML: function ()
				{
					var xml = te.CreateObject("Msxml2.DOMDocument");
					xml.async = false;
					xml.load(this.fn);
					return xml;
				}
			}
		}

		DownloadFile = function (url, fn)
		{
			if (url.fn) {
				DeleteItem(fn);
				fso.MoveFile(url.fn, fn);
				return 0;
			}
			return api.URLDownloadToFile(null, url, fn);
		}
	}
} else {
	var ado = OpenAdodbFromTextFile(fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\"+ Addon_Id + "\\options.html"));
	if (ado) {
		SetTabContents(0, "General", ado.ReadText(adReadAll));
		ado.Close();
	}
	document.getElementById("_curl").value = api.sprintf(99, GetText("Get %s..."), "cURL");
	document.getElementById("_wget").value = api.sprintf(99, GetText("Get %s..."), "Wget");
	SetDL = function (o)
	{
		if (confirmOk(GetText("Are you sure?"))) {
			document.F.Path.value = o.title;
		}
	}
}