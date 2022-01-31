try {
	while (Threads.Images.length) {
		var o = Threads.Images.pop();
		var image = api.CreateObject("WICBitmap");
		image.OnGetAlt = o.OnGetAlt;
		if (image.FromFile(o.path, o.cx)) {
			if (o.cx) {
				image = GetThumbnail(image, o.cx, o.f);
			}
			o.out = o.type ? image.DataURI(o.type) : MainWindow.api.CreateObject("WICBitmap").FromSource(image, o.meta);
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
