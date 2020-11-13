AddEvent(Common.FilterBar.strMenu, function (Ctrl, hMenu, nPos) {
	api.InsertMenu(hMenu, Common.FilterBar.nPos, MF_BYPOSITION | MF_STRING, ++nPos, Common.FilterBar.strName);
	ExtraMenuCommand[nPos] = function () {
		InvokeUI("Addons.FilterBar.Exec", arguments);
		return S_OK;
	}
	return nPos;
});
