#pragma once

#include <vector>
#include <string>
#include <functional>
// #include <Guiddef.h> // for IID HRESULT CoInitialize CoUninitialize

namespace utils::Utils {
    std::vector<unsigned short> ToLittleEndian(const unsigned short* pWords, int nFirstIndex, int nLastIndex);
    std::vector<unsigned short> TrimSpace(std::vector<unsigned short> const& v);
    std::wstring ToWString(std::vector<unsigned short> const& v);
    std::wstring ProgramIdToClassId(std::wstring const& programId);
    std::string Encrypt(std::string const& S, unsigned short key);
    std::string Decrypt(std::string const& S, unsigned short key);

    template <typename TInterface>
    bool COMProcess(IID const& classId, int classContext, IID const& interfaceId, std::function<bool(TInterface* anInterface)>&& process) {
        auto result = false;
        HRESULT hr = CoInitialize(nullptr);
        if (SUCCEEDED(hr)) {
            TInterface* anInterface = nullptr;
            hr = CoCreateInstance(classId, nullptr, classContext, interfaceId, (void**)&anInterface);
            if (SUCCEEDED(hr) && anInterface != nullptr) {
                result = process(anInterface);
                anInterface->Release();
            }
            CoUninitialize();
        }
        return result;
    }
}
