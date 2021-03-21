TITLE = "Tablacus Explorer";
fso = new ActiveXObject("Scripting.FileSystemObject");
sha = new ActiveXObject('Shell.Application');
wsh = new ActiveXObject('WScript.Shell');
args = WScript.Arguments;

var server = GetObject("winmgmts:\\\\.\\root\\cimv2");
var t = new Date().getTime();
if (server) {
	for (;;) {
		var df = new Date().getTime() - t;
		var cols = server.ExecQuery('SELECT * FROM Win32_Process WHERE ExecutablePath="' + (args(0).split("\\").join("\\\\")) + '"');
		if (!cols.Count) {
			break;
		}
		if (df > 30000) {
			for (var list = new Enumerator(cols); !list.atEnd(); list.moveNext()) {
				if (list.item().Terminate() == 0) {
					continue;
				}
			}
		}
		if (df < 6000) {
			WScript.Sleep(500);
		} else if (wsh.Popup(args(2), 5, TITLE, 1) == 2) {
			WScript.Quit();
		}
	}
} else {
	wsh.Popup(args(2), 9, TITLE, 0);
}
if (args.length > 5 && args(5)) {
	var f = args.length > 6 ? parseInt(args(6)) : 0x0210;
	if (/^Move$/i.test(args(5))) {
		sha.NameSpace(args(4)).MoveHere(args(1), f);
	} else if (/^Copy$/i.test(args(5))) {
		sha.NameSpace(args(4)).CopyHere(args(1), f);
	} else if (/^MoveItems$/i.test(args(5))) {
		sha.NameSpace(args(4)).MoveHere(sha.NameSpace(args(1)).Items(), f);
	} else if (/^CopyItems$/i.test(args(5))) {
		sha.NameSpace(args(4)).CopyHere(sha.NameSpace(args(1)).Items(), f);
	}
} else if (args.length > 4 && args(4)) {
	sha.NameSpace(args(4)).CopyHere(sha.NameSpace(args(1)).Items(), 0x0210);
} else {
	sha.NameSpace(fso.GetParentFolderName(args(0))).MoveHere(sha.NameSpace(args(1)).Items(), 0x0210);
}
if (!args(3) || sha.NameSpace(args(1)).Items().Count == 0 || wsh.Popup(args(3), 0, TITLE, 0x21) == 1) {
	wsh.Run('"' + args(0) + '"');
}
