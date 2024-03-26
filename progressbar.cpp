std::vectorstd::string GetFilesInDirectory(const std::string& directory) {
std::vectorstd::string files;
for (const auto& entry : std::filesystem::directory_iterator(directory)) {
files.push_back(entry.path().string());
}
return files;
}
void DisplayProgressBar(const std::vectorstd::string& files) {
int totalFiles = files.size();
int currentFile = 0;
for (const auto& file : files) {
// Load the file and perform operations.

// progress bar update
currentFile++;
std::cout << "\r" << "[" << std::string(currentFile * 50 / totalFiles, '=') << "]" << std::flush;
}
std::cout << std::endl;
}

std::thread fileLoaderThread(& {
files = GetFilesInDirectory(directory);
});

// View progress bar
DisplayProgressBar(files);

fileLoaderThread.join(); 
