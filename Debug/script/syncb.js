// Tablacus Explorer

if (window.UI) {
	InvokeUI = function () {
		if (arguments.length == 2 && arguments[1].unshift) {
			const args = Array.apply(null, arguments[1]);
			args.unshift(arguments[0]);
			api.Invoke(UI.Invoke, args);
			return S_OK;
		}
		api.Invoke(UI.Invoke, Array.apply(null, arguments));
		return S_OK;
	}

	BlurId = function () {
		InvokeUI("BlurId", Array.apply(null, arguments));
	}

	clearTimeout = function () {
		InvokeUI("clearTimeout", Array.apply(null, arguments));
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

	MouseOut = function () {
		InvokeUI("MouseOut", Array.apply(null, arguments));
	}

	MouseOver = function () {
		InvokeUI("MouseOver", Array.apply(null, arguments));
	}

	OpenHttpRequest = function () {
		InvokeUI("OpenHttpRequest", Array.apply(null, arguments));
	}

	ReloadCustomize = function () {
		InvokeUI("ReloadCustomize");
		return S_OK;
	}

	Resize = function () {
		InvokeUI("Resize");
	}

	SelectItem = function () {
		InvokeUI("SelectItem", Array.apply(null, arguments));
	}

	SelectNext = function () {
		InvokeUI("SelectNext", Array.apply(null, arguments));
	}

	SetDisplay = function () {
		InvokeUI("SetDisplay", Array.apply(null, arguments));
	}

	setTimeout = function () {
		InvokeUI("setTimeout", Array.apply(null, arguments));
	}

	ShowStatusTextEx = function () {
		InvokeUI("ShowStatusTextEx", Array.apply(null, arguments));
	}
}
