const Addon_Id = "multiprocess";
const item = GetAddonElement(Addon_Id);

Addons.MultiProcess = {
	File: item.getAttribute("File"),

	Player: async function (autoplay) {
		if (Addons.MultiProcess.tid) {
			clearTimeout(Addons.MultiProcess.tid);
			delete Addons.MultiProcess.tid;
		}
		const src = await ExtractPath(te, (autoplay === true) ? Addons.MultiProcess.File : document.F.File.value);
		if (window.g_nInit) {
			if (src) {
				g_nInit = 0;
			} else {
				--g_nInit;
				Addons.MultiProcess.tid = setTimeout(Addons.MultiProcess.Player, 500);
				return;
			}
		}
		let el;
		if (ui_.IEVer >= 11 && (window.chrome ? /\.mp3$|\.m4a$|\.webm$|\.mp4$|\.wav$|\.ogg$/i : /\.mp3$|\.m4a$|\.mp4$/i).test(src)) {
			el = document.createElement('audio');
			if (autoplay === true) {
				el.setAttribute("autoplay", "true");
			} else {
				el.setAttribute("controls", "true");
			}
		} else if (/\.wav$/i.test(src)) {
			if (autoplay === true) {
				api.PlaySound(src, null, 1);
				return;
			}
			el = document.createElement('button');
			el.innerHTML = "&#x25B6;";
			el.title = await GetText("Play");
			el.onclick = function () {
				Addons.MultiProcess.File = src;
				Addons.MultiProcess.Player(true);
			}
		} else {
			el = document.createElement('embed');
			el.setAttribute("volume", "0");
			el.setAttribute("autoplay", autoplay === true);
		}
		el.src = src;
		el.setAttribute("style", "width: 100%; height: 3.5em");
		const o = document.getElementById('multiprocess_player');
		while (o.firstChild) {
			o.removeChild(o.firstChild);
		}
		if (src) {
			o.appendChild(el);
		}
	}
};

if (window.Addon == 1) {
	$.importScript("addons\\" + Addon_Id + "\\sync.js");
	document.getElementById('None').insertAdjacentHTML("BeforeEnd", '<div id="multiprocess_player"></div>');
} else {
	SetTabContents(0, "", await ReadTextFile("addons\\" + Addon_Id + "\\options.html"));
	g_nInit = 5;
	Addons.MultiProcess.tid = setTimeout(Addons.MultiProcess.Player, 9);
}
