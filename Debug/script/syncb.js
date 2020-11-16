// Tablacus Explorer

if (window.UI) {
	BlurId = function () {
		InvokeUI("BlurId", arguments);
	}

	clearTimeout = function () {
		InvokeUI("clearTimeout", arguments);
	}

	CancelWindowRegistered = function () {
		InvokeUI("CancelWindowRegistered");
	}

	CloseWindow = function () {
		InvokeUI("CloseWindow");
	}

	ExitFullscreen = function () {
		InvokeUI("ExitFullscreen");
	}

	FocusFV = function () {
		InvokeUI("FocusFV");
	}

	MoveSplitter = function () {
		InvokeUI("MoveSplitter", arguments);
	}

	OpenHttpRequest = function () {
		InvokeUI("OpenHttpRequest", arguments);
	}

	ReloadCustomize = function () {
		InvokeUI("ReloadCustomize");
		return S_OK;
	}

	Resize = function () {
		InvokeUI("Resize");
	}

	SelectItem = function () {
		InvokeUI("SelectItem", arguments);
	}

	SelectNext = function () {
		InvokeUI("SelectNext", arguments);
	}

	SetDisplay = function () {
		InvokeUI("SetDisplay", arguments);
	}

	setTimeout = function () {
		api.Invoke(UI.setTimeoutAsync, arguments);
	}

	ShowStatusTextEx = function () {
		InvokeUI("ShowStatusTextEx", arguments);
	}
}
