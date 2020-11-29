// Tablacus Explorer

if (window.UI) {
	InvokeUI = function () {
		if (arguments.length == 2 && arguments[1].unshift) {
			const args = Array.apply(null, arguments[1]);
			args.unshift(arguments[0]);
			api.Invoke(UI.Invoke, args);
			return S_OK;
		}
		api.Invoke(UI.Invoke, arguments);
		return S_OK;
	}

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
