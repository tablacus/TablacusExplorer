var Addon_Id = 'extract';

if (window.Addon == 1) {
	AddEvent("Extract", function (Src, Dest) {
		var s = GetAddonOption("extract", "Path");
		if (s) {
			var r = api.CreateProcess(ExtractMacro(te, s.replace(/%src%/i, api.PathQuoteSpaces(Src)).replace(/%dest%|%dist%/i, api.PathQuoteSpaces(Dest))), api.PathUnquoteSpaces(Dest), 0, 0, 0, true);
			if ("number" === typeof r) {
				return r;
			}
			var bWait;
			do {
				bWait = false;
				WmiProcess("WHERE ProcessId=" + r.ProcessId, function (item) {
					bWait = true;
					api.Sleep(500);
				});
			} while (bWait);
			return S_OK;
		}
	});
} else {
	var ado = OpenAdodbFromTextFile("addons\\" + Addon_Id + "\\options.html");
	if (ado) {
		SetTabContents(0, "", ado.ReadText(adReadAll));
		ado.Close();
	}
	document.getElementById("_7zip").value = api.sprintf(99, GetText("Get %s..."), "7-Zip");

	SetExtract = function (o) {
		if (confirmOk()) {
			document.F.Path.value = o.title;
		}
	}

	SetExe = function (o) {
		if (confirmOk()) {
			var ar = o.title.split(" ")
			var path = 'C:\\Program Files\\' + ar[0];
			var path2 = 'C:\\Program Files (x86)\\' + ar[0];
			ar[0] = api.PathQuoteSpaces(!fso.FileExists(path) && fso.FileExists(path2) ? path2 : path);
			document.F.Path.value = ar.join(" ");
		}
	}
}
