//Menu
AddEvent(Common.Up.strMenu, function (Ctrl, hMenu, nPos) {
	api.InsertMenu(hMenu, Common.Up.nPos, MF_BYPOSITION | MF_STRING, ++nPos, Common.Up.strName);
	ExtraMenuCommand[nPos] = function () {
		InvokeUI("Addons.Up.Exec", arguments);
		return S_OK;
	}
	return nPos;
});
