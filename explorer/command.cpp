#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shobjidl.h>
#include <shlwapi.h>

#include <atomic>
#include <iterator>
#include <new>
#include <string>

namespace {

constexpr CLSID to_pdf = {0x7b657271, 0xe79a, 0x4db5, {0x95, 0xf3, 0x29, 0x45, 0xad, 0xfb, 0xa3, 0xe1}};
constexpr CLSID from_pdf = {0x19787c67, 0x4ed8, 0x4945, {0xb4, 0x3d, 0x36, 0xef, 0x4c, 0x1d, 0xd7, 0xd5}};

struct Target {
    const wchar_t* argument;
    const wchar_t* title;
};

constexpr Target pdf_target[] = {{L"pdf", L"Sang PDF"}};
constexpr Target pdf_outputs[] = {
    {L"docx", L"Sang Word (DOCX)"}, {L"png", L"Sang ảnh PNG"},
    {L"jpeg", L"Sang ảnh JPEG"}, {L"tiff", L"Sang ảnh TIFF"},
    {L"bmp", L"Sang ảnh BMP"}, {L"webp", L"Sang ảnh WebP"},
};

HMODULE module;
std::atomic_long module_refs{0};

class Lifetime {
public:
    Lifetime() { ++module_refs; }
    virtual ~Lifetime() { --module_refs; }
    ULONG add_ref() { return ++refs; }
    ULONG release() {
        const ULONG remaining = --refs;
        if (!remaining) delete this;
        return remaining;
    }

private:
    std::atomic_ulong refs{1};
};

HRESULT selected_path(IShellItemArray* items, PWSTR* path) {
    *path = nullptr;
    DWORD count = 0;
    if (!items || FAILED(items->GetCount(&count)) || count != 1) return E_INVALIDARG;
    IShellItem* item = nullptr;
    HRESULT result = items->GetItemAt(0, &item);
    if (FAILED(result)) return result;
    SFGAOF attributes = 0;
    result = item->GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM, &attributes);
    if (SUCCEEDED(result)) {
        result = attributes == SFGAO_FILESYSTEM
            ? item->GetDisplayName(SIGDN_FILESYSPATH, path) : E_INVALIDARG;
    }
    item->Release();
    return result;
}

bool supports_path(bool source_is_pdf, PCWSTR path) {
    const wchar_t* suffix = PathFindExtensionW(path);
    if (source_is_pdf) return _wcsicmp(suffix, L".pdf") == 0;
    constexpr const wchar_t* extensions[] = {
        L".doc", L".docx", L".jpg", L".jpeg", L".png", L".webp",
        L".bmp", L".tif", L".tiff", L".gif", L".jp2", L".avif",
    };
    for (const wchar_t* extension : extensions) {
        if (_wcsicmp(suffix, extension) == 0) return true;
    }
    return false;
}

std::wstring app_path() {
    wchar_t path[32768];
    const DWORD length = GetModuleFileNameW(module, path, std::size(path));
    if (!length || length == std::size(path)) return {};
    std::wstring executable(path, length);
    executable.resize(executable.find_last_of(L'\\') + 1);
    return executable + L"ChuyenDoiTaiLieu.exe";
}

class Command final : public IExplorerCommand, private Lifetime {
public:
    explicit Command(bool source_is_pdf, const Target* target = nullptr)
        : source_is_pdf(source_is_pdf), target(target) {}

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        *out = nullptr;
        if (iid != IID_IUnknown && iid != IID_IExplorerCommand) return E_NOINTERFACE;
        *out = static_cast<IExplorerCommand*>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return add_ref(); }
    ULONG STDMETHODCALLTYPE Release() override { return release(); }

    HRESULT STDMETHODCALLTYPE GetTitle(IShellItemArray*, LPWSTR* title) override {
        return SHStrDupW(target ? target->title : L"Shadrovert", title);
    }
    HRESULT STDMETHODCALLTYPE GetIcon(IShellItemArray*, LPWSTR* icon) override {
        *icon = nullptr;
        try {
            const auto executable = app_path();
            return executable.empty() ? E_FAIL : SHStrDupW((executable + L",0").c_str(), icon);
        } catch (const std::bad_alloc&) {
            return E_OUTOFMEMORY;
        }
    }
    HRESULT STDMETHODCALLTYPE GetToolTip(IShellItemArray*, LPWSTR* tip) override {
        *tip = nullptr;
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE GetCanonicalName(GUID* name) override {
        *name = source_is_pdf ? from_pdf : to_pdf;
        if (target) {
            const Target* first = source_is_pdf ? pdf_outputs : pdf_target;
            name->Data1 += static_cast<unsigned long>(target - first) + 1;
        }
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetState(IShellItemArray* items, BOOL, EXPCMDSTATE* state) override {
        PWSTR path;
        const HRESULT result = selected_path(items, &path);
        *state = SUCCEEDED(result) && path && supports_path(source_is_pdf, path) ? ECS_ENABLED : ECS_HIDDEN;
        CoTaskMemFree(path);
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Invoke(IShellItemArray* items, IBindCtx*) override {
        if (!target) return E_NOTIMPL;
        PWSTR path;
        const HRESULT result = selected_path(items, &path);
        if (FAILED(result) || !path || !supports_path(source_is_pdf, path)) {
            CoTaskMemFree(path);
            return E_INVALIDARG;
        }

        HRESULT started;
        try {
            const auto executable = app_path();
            if (executable.empty()) {
                started = E_FAIL;
            } else {
                std::wstring command_line = L"\"" + executable + L"\" --convert " + target->argument + L" \"" + path + L"\"";
                STARTUPINFOW startup{};
                startup.cb = sizeof(startup);
                PROCESS_INFORMATION process{};
                if (CreateProcessW(executable.c_str(), command_line.data(), nullptr, nullptr, FALSE,
                                   0, nullptr, nullptr, &startup, &process)) {
                    CloseHandle(process.hThread);
                    CloseHandle(process.hProcess);
                    started = S_OK;
                } else {
                    started = HRESULT_FROM_WIN32(GetLastError());
                }
            }
        } catch (const std::bad_alloc&) {
            started = E_OUTOFMEMORY;
        }
        CoTaskMemFree(path);
        return started;
    }
    HRESULT STDMETHODCALLTYPE GetFlags(EXPCMDFLAGS* flags) override {
        *flags = target ? ECF_DEFAULT : ECF_HASSUBCOMMANDS;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE EnumSubCommands(IEnumExplorerCommand** out) override;

private:
    bool source_is_pdf;
    const Target* target;
};

class CommandList final : public IEnumExplorerCommand, private Lifetime {
public:
    explicit CommandList(bool source_is_pdf, ULONG index = 0)
        : source_is_pdf(source_is_pdf), index(index) {}

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        *out = nullptr;
        if (iid != IID_IUnknown && iid != IID_IEnumExplorerCommand) return E_NOINTERFACE;
        *out = static_cast<IEnumExplorerCommand*>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return add_ref(); }
    ULONG STDMETHODCALLTYPE Release() override { return release(); }

    HRESULT STDMETHODCALLTYPE Next(ULONG count, IExplorerCommand** commands, ULONG* fetched) override {
        if (fetched) *fetched = 0;
        if (!commands || (!fetched && count != 1)) return E_POINTER;
        const Target* targets = source_is_pdf ? pdf_outputs : pdf_target;
        ULONG done = 0;
        while (done < count && index < size()) {
            auto command = new (std::nothrow) Command(source_is_pdf, &targets[index]);
            if (!command) {
                index -= done;
                while (done) commands[--done]->Release();
                return E_OUTOFMEMORY;
            }
            commands[done++] = command;
            ++index;
        }
        if (fetched) *fetched = done;
        return done == count ? S_OK : S_FALSE;
    }
    HRESULT STDMETHODCALLTYPE Skip(ULONG count) override {
        const ULONG remaining = size() - index;
        index += count < remaining ? count : remaining;
        return count <= remaining ? S_OK : S_FALSE;
    }
    HRESULT STDMETHODCALLTYPE Reset() override {
        index = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Clone(IEnumExplorerCommand** out) override {
        *out = new (std::nothrow) CommandList(source_is_pdf, index);
        return *out ? S_OK : E_OUTOFMEMORY;
    }

private:
    ULONG size() const { return source_is_pdf ? std::size(pdf_outputs) : std::size(pdf_target); }
    bool source_is_pdf;
    ULONG index;
};

HRESULT Command::EnumSubCommands(IEnumExplorerCommand** out) {
    *out = nullptr;
    if (target) return E_NOTIMPL;
    *out = new (std::nothrow) CommandList(source_is_pdf);
    return *out ? S_OK : E_OUTOFMEMORY;
}

class Factory final : public IClassFactory, private Lifetime {
public:
    explicit Factory(bool source_is_pdf) : source_is_pdf(source_is_pdf) {}

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        *out = nullptr;
        if (iid != IID_IUnknown && iid != IID_IClassFactory) return E_NOINTERFACE;
        *out = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return add_ref(); }
    ULONG STDMETHODCALLTYPE Release() override { return release(); }
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID iid, void** out) override {
        *out = nullptr;
        if (outer) return CLASS_E_NOAGGREGATION;
        auto command = new (std::nothrow) Command(source_is_pdf);
        if (!command) return E_OUTOFMEMORY;
        const HRESULT result = command->QueryInterface(iid, out);
        command->Release();
        return result;
    }
    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override {
        if (lock) ++module_refs;
        else --module_refs;
        return S_OK;
    }

private:
    bool source_is_pdf;
};

}  // namespace

BOOL WINAPI DllMain(HMODULE instance, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) module = instance;
    return TRUE;
}

extern "C" HRESULT WINAPI DllGetClassObject(REFCLSID clsid, REFIID iid, void** out) {
    *out = nullptr;
    if (clsid != to_pdf && clsid != from_pdf) return CLASS_E_CLASSNOTAVAILABLE;
    auto factory = new (std::nothrow) Factory(clsid == from_pdf);
    if (!factory) return E_OUTOFMEMORY;
    const HRESULT result = factory->QueryInterface(iid, out);
    factory->Release();
    return result;
}

extern "C" HRESULT WINAPI DllCanUnloadNow() {
    return module_refs == 0 ? S_OK : S_FALSE;
}
