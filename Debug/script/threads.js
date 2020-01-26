try {
	while (Threads.Images.length) {
		var o = Threads.Images.pop();
		var image = api.CreateObject("WICBitmap");
		image.OnFromFile = o.OnFromFile;
		image.OnFromStream = o.OnFromStream;
		if (image.FromFile(o.path, o.cx)) {
			if (o.cx) {
				image = GetThumbnail(image, o.cx, o.f);
			}
			o.out = image;
			o.callback(o);
		}
	}
} catch (e) {}
try {
	Threads.End(Id);
} catch (e) { }
