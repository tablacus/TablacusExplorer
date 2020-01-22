try {
	while (Threads.Images.length) {
		var o = Threads.Images.pop();
		var image = api.CreateObject("WICBitmap").FromFile(o.path, o.cx);
		if (image) {
			if (o.cx) {
				image = GetThumbnail(image, o.cx, o.f);
			}
			o.out = isFinite(o.cl) ? image.GetHBITMAP(o.cl) : image.GetStream();
			o.callback(o);
			if (isFinite(o.cl)) {
				api.DeleteObject(o.out);
			}
		}
	}
} catch (e) {}
try {
	Threads.End(Id);
} catch (e) { }
