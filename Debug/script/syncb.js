// Tablacus Explorer

InvokeUI = function () {
	if (arguments.length == 2 && arguments[1].unshift) {
		const args = Array.apply(null, arguments[1]);
		args.unshift(arguments[0]);
		InvokeFunc(UI.Invoke, args);
		return S_OK;
	}
	InvokeFunc(UI.Invoke, Array.apply(null, arguments));
	return S_OK;
}

BlurId = function () {
	InvokeUI("BlurId", Array.apply(null, arguments));
}

clearTimeout = function () {
	InvokeUI("clearTimeout", Array.apply(null, arguments));
}

clipboardData = {
	setData: function (format, data) {
		api.SetClipboardData(data);
		return true;
	},
	getData: api.GetClipboardData
}

CloseWindow = function () {
	InvokeUI("CloseWindow");
}

ExitFullscreen = function () {
	FullscreenChanged(false);
	InvokeUI("ExitFullscreen");
}

FocusFV = function () {
	InvokeUI("FocusFV");
}

GetFolderView = function (Ctrl, pt, bStrict) {
	if (!Ctrl) {
		return te.Ctrl(CTRL_FV);
	}
	const nType = Ctrl.Type;
	if (nType <= CTRL_EB) {
		return Ctrl;
	}
	if (nType == CTRL_TV) {
		return Ctrl.FolderView;
	}
	if (nType != CTRL_TC) {
		return te.Ctrl(CTRL_FV);
	}
	if (pt) {
		const FV = pt.Target || Ctrl.HitTest(pt);
		if (FV) {
			return FV;
		}
	}
	if (!bStrict || !pt) {
		return Ctrl.Selected;
	}
};

ImgBase64 = function (el, index, h) {
	return MakeImgSrc(ExtractMacro(te, el.src), index, false, h || el.height);
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

SetDisplay = function () {
	InvokeUI("SetDisplay", Array.apply(null, arguments));
}

setTimeout = function () {
	InvokeUI("setTimeout", Array.apply(null, arguments));
}

ShowStatusTextEx = function () {
	InvokeUI("ShowStatusTextEx", Array.apply(null, arguments));
}

WebBrowser = te.Ctrl(CTRL_WB);
