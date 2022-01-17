const Addon_Id = "download";
const item = GetAddonElement(Addon_Id);

Common.Download = {
	Path: item.getAttribute("Path"),
	Show: GetNum(item.getAttribute("Visible")) ? SW_SHOWNORMAL : SW_HIDE
}
if (Common.Download.Path) {
	AddEvent("createHttpRequest", function () {
		const o = api.CreateObject("Object");

		o.open = function (method, url) {
			o.URL = url;
		}

		o.send = function () {
			o.fn = GetTempPath(3);
			wsh.Run(ExtractMacro(te, Common.Download.Path.replace(/%url%/ig, o.URL).replace(/%file%/ig, o.fn)), Common.Download.Show, true);
			o.readyState = 4;
			const wfd = api.Memory("WIN32_FIND_DATA");
			const hFind = api.FindFirstFile(o.fn, wfd);
			api.FindClose(hFind);
			o.status = (hFind != INVALID_HANDLE_VALUE) ? wfd.nFileSizeLow ? 200 : 403 : 404;
			if (o.status != 200) {
				DeleteItem(o.fn);
			}
			InvokeFunc(o.onload || o.onreadystatechange, [o]);
		}

		o.get_responseText = function () {
			return ReadTextFile(o.fn);
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
