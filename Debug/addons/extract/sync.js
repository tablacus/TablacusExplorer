var Addon_Id = 'extract';
var item = GetAddonElement(Addon_Id);

if (Sync.Extract = item.getAttribute("Path")) {
	AddEvent("Extract", function (Src, Dest) {
		var r = api.CreateProcess(ExtractMacro(te, Sync.Extract.replace(/%src%/i, api.PathQuoteSpaces(Src)).replace(/%dest%|%dist%/i, api.PathQuoteSpaces(Dest))), api.PathUnquoteSpaces(Dest), 0, 0, 0, true);
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
	});
}
