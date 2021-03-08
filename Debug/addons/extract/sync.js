const Addon_Id = 'extract';
const item = GetAddonElement(Addon_Id);

if (Sync.Extract = item.getAttribute("Path")) {
	AddEvent("Extract", function (Src, Dest) {
		const r = api.CreateProcess(ExtractMacro(te, Sync.Extract.replace(/%src%/i, PathQuoteSpaces(Src)).replace(/%dest%|%dist%/i, PathQuoteSpaces(Dest))), PathUnquoteSpaces(Dest), 0, 0, 0, true);
		if ("number" === typeof r) {
			return r;
		}
		let bWait;
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
