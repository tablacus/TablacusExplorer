// Tablacus Explorer

if (window.UI) {
	BlurId = function (Id) {
		InvokeUI("BlurId", arguments);
	}

	clearTimeout = function () {
		InvokeUI("clearTimeout", arguments);
	}

	CloseWindow = function () {
		InvokeUI("CloseWindow");
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

	setTimeout = function () {
		api.Invoke(UI.setTimeoutAsync, arguments);
	}

}
