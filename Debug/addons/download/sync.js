var Addon_Id = 'download';
var item = GetAddonElement(Addon_Id);

Common.Download = {
	Path: item.getAttribute("Path"),
	Show: api.LowPart(item.getAttribute("Visible")) ? SW_SHOWNORMAL : SW_HIDE
}
if (Common.Download.Path) {
	AddEvent("createHttpRequest", function () {
		var o = api.CreateObject("Object");

		o.open = function (method, url) {
			o.URL = url;
		}

		o.send = function () {
			var temp = BuildPath(fso.GetSpecialFolder(2).Path, "tablacus");
			CreateFolder(temp);
			o.fn = temp + "\\" + fso.GetTempName();
			DeleteItem(o.fn);
			wsh.Run(ExtractMacro(te, Common.Download.Path.replace(/%url%/ig, o.URL).replace(/%file%/ig, o.fn)), Common.Download.Show, true);
			o.readyState = 4;
			var wfd = api.Memory("WIN32_FIND_DATA");
			var hFind = api.FindFirstFile(o.fn, wfd);
			api.FindClose(hFind);
			o.status = (hFind != INVALID_HANDLE_VALUE) ? wfd.nFileSizeLow ? 200 : 403 : 404;
			if (o.status != 200) {
				DeleteItem(o.fn);
			}
			if (o.onload) {
				api.Invoke(o.onload, o);
			} else if (o.onreadystatechange) {
				api.Invoke(o.onreadystatechange, o);
			}
		}

		o.get_responseText = function () {
			return ReadTextFile(o.fn);
		}

		o.get_responseXML = function () {
			var xml = te.CreateObject("Msxml2.DOMDocument");
			xml.async = false;
			xml.load(o.fn);
			return xml;
		}

		return o;
	});

	AddEvent("DownloadFile", function (url, fn) {
		if (url.fn) {
			DeleteItem(fn);
			fso.MoveFile(url.fn, fn);
			return 0;
		}
		return api.URLDownloadToFile(null, url, fn);
	});
}
