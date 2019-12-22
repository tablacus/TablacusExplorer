var hEvent = MainWindow.Threads.hEvent;
for(;;) {
	api.WaitForSingleObject(hEvent, -1);
	while(MainWindow.Threads.Codes.length) {
		new Function(MainWindow.Threads.Codes.shift())();
	}
	api.ResetEvent(hEvent);
}
api.CloseHandle(hEvent);
