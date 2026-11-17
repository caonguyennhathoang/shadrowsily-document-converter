#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include <string>

namespace {

int notify(const wchar_t* message, bool success = false) {
    MessageBoxW(nullptr, message, L"Cài đặt Shadrovert",
                MB_OK | (success ? MB_ICONINFORMATION : MB_ICONERROR));
    return success ? 0 : 1;
}

bool write_package(const wchar_t* path) {
    HRSRC resource = FindResourceW(nullptr, MAKEINTRESOURCEW(101), RT_RCDATA);
    if (!resource) return false;
    HGLOBAL data = LoadResource(nullptr, resource);
    if (!data) return false;
    const DWORD size = SizeofResource(nullptr, resource);
    HANDLE file = CreateFileW(path, GENERIC_WRITE, 0, nullptr, TRUNCATE_EXISTING,
                              FILE_ATTRIBUTE_TEMPORARY, nullptr);
    if (file == INVALID_HANDLE_VALUE) return false;
    DWORD written = 0;
    const bool ok = WriteFile(file, LockResource(data), size, &written, nullptr) && written == size;
    CloseHandle(file);
    return ok;
}

bool install_package(const wchar_t* path) {
    wchar_t system[MAX_PATH];
    const UINT length = GetSystemDirectoryW(system, MAX_PATH);
    if (!length || length >= MAX_PATH || !SetEnvironmentVariableW(L"CONVERTER_SETUP_PACKAGE", path)) {
        return false;
    }
    const std::wstring powershell = std::wstring(system) + L"\\WindowsPowerShell\\v1.0\\powershell.exe";
    std::wstring command = L"\"" + powershell + L"\" -NoLogo -NoProfile -NonInteractive -Command "
        L"\"$ErrorActionPreference = 'Stop'; Add-AppxPackage -Path $env:CONVERTER_SETUP_PACKAGE -AllowUnsigned\"";
    STARTUPINFOW startup{};
    startup.cb = sizeof(startup);
    PROCESS_INFORMATION process{};
    if (!CreateProcessW(powershell.c_str(), command.data(), nullptr, nullptr, FALSE,
                        CREATE_NO_WINDOW, nullptr, nullptr, &startup, &process)) return false;
    CloseHandle(process.hThread);
    DWORD code = 1;
    const bool ok = WaitForSingleObject(process.hProcess, INFINITE) == WAIT_OBJECT_0
        && GetExitCodeProcess(process.hProcess, &code) && code == 0;
    CloseHandle(process.hProcess);
    return ok;
}

}  // namespace

int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR, int) {
    wchar_t folder[MAX_PATH];
    wchar_t temporary[MAX_PATH];
    const DWORD length = GetTempPathW(MAX_PATH, folder);
    if (!length || length >= MAX_PATH || !GetTempFileNameW(folder, L"CDT", 0, temporary)) {
        return notify(L"Không thể tạo tệp cài đặt tạm.");
    }
    // Give Add-AppxPackage the expected file extension and fail if that path already exists.
    const std::wstring package = std::wstring(temporary) + L".msix";
    if (!MoveFileW(temporary, package.c_str())) {
        DeleteFileW(temporary);
        return notify(L"Không thể chuẩn bị gói cài đặt.");
    }
    const bool installed = write_package(package.c_str()) && install_package(package.c_str());
    DeleteFileW(package.c_str());
    if (!installed) return notify(L"Cài đặt thất bại. Hãy đóng ứng dụng và thử lại.");
    return notify(L"Đã cài Shadrovert và menu chuột phải.", true);
}
