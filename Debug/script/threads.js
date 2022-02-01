try {
	while (Threads.Images.length) {
		var o = Threads.Images.pop();
		var image = api.CreateObject("WICBitmap");
		image.OnGetAlt = o.OnGetAlt;
		if (image.FromFile(o.path, o.cx)) {
			if (o.cx) {
				if (!o.anime || image.GetFrameCount() < 2) {
					image = GetThumbnail(image, o.cx, o.f);
				}
			}
			if (o.mix) {
				image.AlphaBlend(o.rc, o.mix, o.max || 100);
			}
			if ("string" === typeof o.type) {
				o.out = image.DataURI(o.type, o.anime && o.quality != -2 && image.GetFrameCount() > 1 ? -2 : o.quality);
			} else if ("number" === typeof o.type) {
				o.out = image.GetHBITMAP(o.type);
			} else {
				o.out = MainWindow.api.CreateObject("WICBitmap").FromStream(image.GetStream("image/png", -2));
			}
			api.Invoke(o.onload || o.callback, o);
		} else if (o.onerror) {
			api.Invoke(o.onerror, o);
		}
	}
} catch (e) { }
try {
	Threads.End(Id);
} catch (e) { }

function GetThumbnail (image, m, f) {
	var w = image.GetWidth(), h = image.GetHeight(), z = m / Math.max(w, h);
	if (z == 1 || (f && z > 1)) {
		return image;
	}
	return image.GetThumbnailImage(w * z, h * z);
}
