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
var v = fso.GetFileVersion(args(0));
sha.NameSpace(fso.GetParentFolderName(args(0))).CopyHere(sha.NameSpace(args(1)).Items(), 0x0210);
if (v != fso.GetFileVersion(args(0)) || wsh.Popup(args(3), 0, TITLE, 0x21) == 1) {
	wsh.Run('"' + args(0) + '"');
}
