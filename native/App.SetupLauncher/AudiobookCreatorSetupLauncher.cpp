#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <urlmon.h>

#include <string>
#include <vector>

#pragma comment(lib, "urlmon.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")

namespace
{
    constexpr wchar_t kProductName[] = L"Audiobook Creator Setup";
    constexpr wchar_t kDotNetRuntimeUrl[] = L"https://aka.ms/dotnet/7.0/windowsdesktop-runtime-win-x64.exe";
    constexpr int kManagedSetupResourceId = 201;
    constexpr wchar_t kManagedSetupTempFolderName[] = L"AudiobookCreatorSetup";
    constexpr wchar_t kManagedSetupFileName[] = L"AudiobookCreatorSetup.exe";

    std::wstring GetExeDirectory()
    {
        std::wstring buffer(MAX_PATH, L'\0');
        DWORD length = 0;

        while (true)
        {
            length = GetModuleFileNameW(nullptr, &buffer[0], static_cast<DWORD>(buffer.size()));
            if (length == 0)
            {
                return L"";
            }

            if (length < buffer.size() - 1)
            {
                buffer.resize(length);
                break;
            }

            buffer.resize(buffer.size() * 2);
        }

        if (!PathRemoveFileSpecW(&buffer[0]))
        {
            return L"";
        }

        buffer.resize(wcslen(buffer.c_str()));
        return buffer;
    }

    std::wstring CombinePath(const std::wstring& left, const std::wstring& right)
    {
        std::wstring result = left;
        if (!result.empty() && result.back() != L'\\' && result.back() != L'/')
        {
            result += L'\\';
        }
        result += right;
        return result;
    }

    void ShowMessage(const std::wstring& text, UINT flags)
    {
        MessageBoxW(nullptr, text.c_str(), kProductName, flags | MB_OK);
    }

    bool HasRuntimeDirectoryVersionPrefix(const std::wstring& rootPath, const std::wstring& prefix)
    {
        const std::wstring searchPattern = CombinePath(rootPath, prefix + L"*");
        WIN32_FIND_DATAW findData{};
        HANDLE findHandle = FindFirstFileW(searchPattern.c_str(), &findData);
        if (findHandle == INVALID_HANDLE_VALUE)
        {
            return false;
        }

        do
        {
            if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
            {
                FindClose(findHandle);
                return true;
            }
        } while (FindNextFileW(findHandle, &findData));

        FindClose(findHandle);
        return false;
    }

    bool HasDotNetDesktopRuntime7OnDisk()
    {
        wchar_t programFilesPath[MAX_PATH];
        DWORD length = GetEnvironmentVariableW(L"ProgramW6432", programFilesPath, MAX_PATH);
        if (length == 0 || length >= MAX_PATH)
        {
            length = GetEnvironmentVariableW(L"ProgramFiles", programFilesPath, MAX_PATH);
        }

        if (length == 0 || length >= MAX_PATH)
        {
            return false;
        }

        const std::wstring runtimeRoot = CombinePath(
            programFilesPath,
            L"dotnet\\shared\\Microsoft.WindowsDesktop.App");

        return HasRuntimeDirectoryVersionPrefix(runtimeRoot, L"7.");
    }

    bool HasDotNetDesktopRuntime7()
    {
        HKEY key = nullptr;
        constexpr wchar_t subkey[] = L"SOFTWARE\\dotnet\\Setup\\InstalledVersions\\x64\\sharedfx\\Microsoft.WindowsDesktop.App";
        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, subkey, 0, KEY_READ | KEY_WOW64_64KEY, &key) != ERROR_SUCCESS)
        {
            return HasDotNetDesktopRuntime7OnDisk();
        }

        DWORD index = 0;
        wchar_t name[256];
        DWORD nameLength = 0;

        while (true)
        {
            nameLength = static_cast<DWORD>(std::size(name));
            const auto status = RegEnumKeyExW(
                key,
                index,
                name,
                &nameLength,
                nullptr,
                nullptr,
                nullptr,
                nullptr);

            if (status == ERROR_NO_MORE_ITEMS)
            {
                break;
            }

            if (status == ERROR_SUCCESS && nameLength >= 2 && name[0] == L'7' && name[1] == L'.')
            {
                RegCloseKey(key);
                return true;
            }

            ++index;
        }

        RegCloseKey(key);
        return HasDotNetDesktopRuntime7OnDisk();
    }

    bool DownloadDotNetRuntime(const std::wstring& destinationPath)
    {
        DeleteFileW(destinationPath.c_str());
        const HRESULT hr = URLDownloadToFileW(nullptr, kDotNetRuntimeUrl, destinationPath.c_str(), 0, nullptr);
        return SUCCEEDED(hr);
    }

    bool RunProcessAndWait(const std::wstring& filePath, const std::wstring& parameters, bool elevate, DWORD* exitCode)
    {
        SHELLEXECUTEINFOW execInfo{};
        execInfo.cbSize = sizeof(execInfo);
        execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
        execInfo.hwnd = nullptr;
        execInfo.lpVerb = elevate ? L"runas" : L"open";
        execInfo.lpFile = filePath.c_str();
        execInfo.lpParameters = parameters.empty() ? nullptr : parameters.c_str();
        execInfo.lpDirectory = nullptr;
        execInfo.nShow = SW_SHOWNORMAL;

        if (!ShellExecuteExW(&execInfo))
        {
            return false;
        }

        if (execInfo.hProcess != nullptr)
        {
            WaitForSingleObject(execInfo.hProcess, INFINITE);
            if (exitCode != nullptr)
            {
                GetExitCodeProcess(execInfo.hProcess, exitCode);
            }
            CloseHandle(execInfo.hProcess);
        }

        return true;
    }

    bool LaunchProcessNoWait(const std::wstring& filePath, const std::wstring& workingDirectory)
    {
        SHELLEXECUTEINFOW execInfo{};
        execInfo.cbSize = sizeof(execInfo);
        execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
        execInfo.hwnd = nullptr;
        execInfo.lpVerb = L"open";
        execInfo.lpFile = filePath.c_str();
        execInfo.lpDirectory = workingDirectory.c_str();
        execInfo.nShow = SW_SHOWNORMAL;

        if (!ShellExecuteExW(&execInfo))
        {
            return false;
        }

        if (execInfo.hProcess != nullptr)
        {
            CloseHandle(execInfo.hProcess);
        }

        return true;
    }

    bool WriteBufferToFile(const std::wstring& filePath, const void* buffer, DWORD size)
    {
        HANDLE fileHandle = CreateFileW(
            filePath.c_str(),
            GENERIC_WRITE,
            0,
            nullptr,
            CREATE_ALWAYS,
            FILE_ATTRIBUTE_NORMAL,
            nullptr);

        if (fileHandle == INVALID_HANDLE_VALUE)
        {
            return false;
        }

        DWORD bytesWritten = 0;
        const bool success = WriteFile(fileHandle, buffer, size, &bytesWritten, nullptr) != FALSE && bytesWritten == size;
        CloseHandle(fileHandle);
        return success;
    }

    bool ExtractManagedSetupFromResource(std::wstring* extractedPath, std::wstring* extractedDirectory)
    {
        HRSRC resourceInfo = FindResourceW(nullptr, MAKEINTRESOURCEW(kManagedSetupResourceId), RT_RCDATA);
        if (resourceInfo == nullptr)
        {
            return false;
        }

        HGLOBAL loadedResource = LoadResource(nullptr, resourceInfo);
        if (loadedResource == nullptr)
        {
            return false;
        }

        const DWORD resourceSize = SizeofResource(nullptr, resourceInfo);
        const void* resourceBytes = LockResource(loadedResource);
        if (resourceBytes == nullptr || resourceSize == 0)
        {
            return false;
        }

        wchar_t tempPathBuffer[MAX_PATH];
        if (GetTempPathW(MAX_PATH, tempPathBuffer) == 0)
        {
            return false;
        }

        const std::wstring setupDirectory = CombinePath(tempPathBuffer, kManagedSetupTempFolderName);
        if (!CreateDirectoryW(setupDirectory.c_str(), nullptr))
        {
            const DWORD lastError = GetLastError();
            if (lastError != ERROR_ALREADY_EXISTS)
            {
                return false;
            }
        }

        const std::wstring setupPath = CombinePath(setupDirectory, kManagedSetupFileName);
        if (!WriteBufferToFile(setupPath, resourceBytes, resourceSize))
        {
            return false;
        }

        *extractedPath = setupPath;
        *extractedDirectory = setupDirectory;
        return true;
    }

    bool EnsureDotNetRuntimeInstalled()
    {
        if (HasDotNetDesktopRuntime7())
        {
            return true;
        }

        const auto response = MessageBoxW(
            nullptr,
            L"Audiobook Creator Setup needs Microsoft .NET Desktop Runtime 7.x before it can continue.\n\n"
            L"The launcher will now download the official runtime from Microsoft and finish that step first.",
            kProductName,
            MB_OKCANCEL | MB_ICONINFORMATION);

        if (response != IDOK)
        {
            return false;
        }

        wchar_t tempPathBuffer[MAX_PATH];
        if (GetTempPathW(MAX_PATH, tempPathBuffer) == 0)
        {
            ShowMessage(L"Could not resolve the Windows temp folder.", MB_ICONERROR);
            return false;
        }

        const std::wstring installerPath = CombinePath(tempPathBuffer, L"windowsdesktop-runtime-win-x64.exe");
        if (!DownloadDotNetRuntime(installerPath))
        {
            ShowMessage(
                L"Audiobook Creator Setup could not download Microsoft .NET Desktop Runtime 7.x.\n\n"
                L"Check your internet connection and try again.",
                MB_ICONERROR);
            return false;
        }

        DWORD exitCode = 1;
        if (!RunProcessAndWait(installerPath, L"/install /quiet /norestart", true, &exitCode) || exitCode != 0)
        {
            ShowMessage(
                L"Microsoft .NET Desktop Runtime 7.x did not install successfully.",
                MB_ICONERROR);
            return false;
        }

        if (!HasDotNetDesktopRuntime7())
        {
            ShowMessage(
                L"Audiobook Creator Setup could not verify Microsoft .NET Desktop Runtime 7.x after installation.",
                MB_ICONERROR);
            return false;
        }

        return true;
    }
}

int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR, int)
{
    if (!EnsureDotNetRuntimeInstalled())
    {
        return 1;
    }

    std::wstring managedSetupPath;
    std::wstring managedSetupDir;
    if (!ExtractManagedSetupFromResource(&managedSetupPath, &managedSetupDir))
    {
        ShowMessage(
            L"Audiobook Creator Setup could not extract the embedded setup application.",
            MB_ICONERROR);
        return 1;
    }

    if (!LaunchProcessNoWait(managedSetupPath, managedSetupDir))
    {
        ShowMessage(
            L"Audiobook Creator Setup could not start the main setup application.",
            MB_ICONERROR);
        return 1;
    }

    return 0;
}
