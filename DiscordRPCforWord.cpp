#pragma comment(lib, "discord-rpc.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "advapi32.lib")

#include <iostream>
#include <windows.h>
#include <string>
#include <cmath>
#include "discord_rpc.h"

// =====================================================
// Helper: Panggil property dari IDispatch (COM late binding)
// Tidak perlu #import sama sekali!
// =====================================================

// Ambil property bertipe string dari COM object
std::wstring GetStringProperty(IDispatch* pDisp, const wchar_t* propName) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(propName);
    if (FAILED(pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return L"";

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pDisp->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    std::wstring ret;
    if (result.vt == VT_BSTR && result.bstrVal)
        ret = result.bstrVal;
    VariantClear(&result);
    return ret;
}

// Ambil property bertipe long dari COM object
long GetLongProperty(IDispatch* pDisp, const wchar_t* propName) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(propName);
    if (FAILED(pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return 0;

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pDisp->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    long ret = 0;
    if (result.vt == VT_I4) ret = result.lVal;
    else if (result.vt == VT_I2) ret = result.iVal;
    VariantClear(&result);
    return ret;
}

// Ambil property bertipe IDispatch* (sub-object) dari COM object
IDispatch* GetDispatchProperty(IDispatch* pDisp, const wchar_t* propName) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(propName);
    if (FAILED(pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return nullptr;

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pDisp->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    IDispatch* ret = nullptr;
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        ret = result.pdispVal;
        ret->AddRef();
    }
    VariantClear(&result);
    return ret;
}

// Panggil method dengan 1 parameter integer, hasil IDispatch*
IDispatch* InvokeWithLongParam(IDispatch* pDisp, const wchar_t* methodName, long param) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(methodName);
    if (FAILED(pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return nullptr;

    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_I4;
    arg.lVal = param;

    DISPPARAMS dp = { &arg, nullptr, 1, 0 };
    VARIANT result;
    VariantInit(&result);
    pDisp->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, nullptr, nullptr);

    IDispatch* ret = nullptr;
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        ret = result.pdispVal;
        ret->AddRef();
    }
    VariantClear(&result);
    return ret;
}

// Ambil property Information[index] dari Selection
long GetSelectionInfo(IDispatch* pSelection, long infoIndex) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(L"Information");
    if (FAILED(pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return 0;

    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_I4;
    arg.lVal = infoIndex; // wdActiveEndPageNumber = 3

    DISPPARAMS dp = { &arg, nullptr, 1, 0 };
    VARIANT result;
    VariantInit(&result);
    pSelection->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    long ret = 0;
    if (result.vt == VT_I4) ret = result.lVal;
    else if (result.vt == VT_I2) ret = result.iVal;
    VariantClear(&result);
    return ret;
}

// Konversi wide string ke UTF-8
std::string WideToUTF8(const std::wstring& wstr) {
    if (wstr.empty()) return "";
    int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string result(size - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &result[0], size, nullptr, nullptr);
    return result;
}

// =====================================================
// Discord callback
// =====================================================
void handleDiscordReady(const DiscordUser* user) {
    std::cout << "Discord siap! Terhubung sebagai: " << user->username << std::endl;
}

// =====================================================
// MAIN
// =====================================================
int main() {
    // Init Discord RPC
    DiscordEventHandlers handlers;
    memset(&handlers, 0, sizeof(handlers));
    handlers.ready = handleDiscordReady;
    Discord_Initialize("1503785206417592330", &handlers, 1, NULL);

    // Init COM
    if (FAILED(CoInitialize(NULL))) {
        std::cout << "Gagal inisialisasi COM!" << std::endl;
        return 1;
    }

    std::cout << "========================================" << std::endl;
    std::cout << "   Word RPC for Discord sedang jalan... " << std::endl;
    std::cout << "========================================" << std::endl;

    static int64_t startTime = (int64_t)time(0);

    // Konstanta Word (tidak perlu #import)
    // wdActiveEndPageNumber = 3
    // wdStatisticPages      = 2
    const long WD_ACTIVE_END_PAGE_NUMBER = 3;
    const long WD_STATISTIC_PAGES = 2;

    while (true) {
        CLSID clsid;
        IUnknown* pUnk = nullptr;
        IDispatch* pWordApp = nullptr;

        if (SUCCEEDED(CLSIDFromProgID(L"Word.Application", &clsid)) &&
            SUCCEEDED(GetActiveObject(clsid, NULL, &pUnk)) &&
            SUCCEEDED(pUnk->QueryInterface(IID_IDispatch, (void**)&pWordApp))) {

            try {
                // Ambil ActiveDocument
                IDispatch* pDoc = GetDispatchProperty(pWordApp, L"ActiveDocument");
                // Ambil Selection
                IDispatch* pSel = GetDispatchProperty(pWordApp, L"Selection");

                if (pDoc && pSel) {
                    // Nama dokumen
                    std::wstring wDocName = GetStringProperty(pDoc, L"Name");
                    std::string  docName = WideToUTF8(wDocName);

                    // Halaman sekarang
                    long curPage = GetSelectionInfo(pSel, WD_ACTIVE_END_PAGE_NUMBER);

                    // Total halaman via ComputeStatistics(wdStatisticPages=2)
                    IDispatch* pStatResult = InvokeWithLongParam(pDoc, L"ComputeStatistics", WD_STATISTIC_PAGES);
                    // ComputeStatistics langsung return long, bukan IDispatch
                    // Pakai cara langsung:
                    DISPID dispId2;
                    LPOLESTR statName = const_cast<LPOLESTR>(L"ComputeStatistics");
                    long totalPage = 0;
                    if (SUCCEEDED(pDoc->GetIDsOfNames(IID_NULL, &statName, 1, LOCALE_USER_DEFAULT, &dispId2))) {
                        VARIANT arg2;
                        VariantInit(&arg2);
                        arg2.vt = VT_I4;
                        arg2.lVal = WD_STATISTIC_PAGES;
                        DISPPARAMS dp2 = { &arg2, nullptr, 1, 0 };
                        VARIANT res2;
                        VariantInit(&res2);
                        pDoc->Invoke(dispId2, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp2, &res2, nullptr, nullptr);
                        if (res2.vt == VT_I4) totalPage = res2.lVal;
                        else if (res2.vt == VT_I2) totalPage = res2.iVal;
                        VariantClear(&res2);
                    }

                    DiscordRichPresence presence;
                    memset(&presence, 0, sizeof(presence));
                    presence.details = docName.c_str();
                    std::string stateLabel = "Page " + std::to_string(curPage) + " of " + std::to_string(totalPage);
                    presence.state = stateLabel.c_str();


                    presence.largeImageKey = "word_logo";
                    presence.largeImageText = "Microsoft Word";

                    presence.startTimestamp = startTime;

                    Discord_UpdatePresence(&presence);

                    std::cout << "\r[Status] " << docName
                        << " (" << curPage << "/" << totalPage << ")     " << std::flush;
                }
                else {
                    DiscordRichPresence presence;
                    memset(&presence, 0, sizeof(presence));
                    presence.details = "Idle (Tidak ada dokumen)";
                    presence.largeImageKey = "word_logo";
                    Discord_UpdatePresence(&presence);
                }

                if (pDoc) pDoc->Release();
                if (pSel) pSel->Release();
            }
            catch (...) {
                DiscordRichPresence presence;
                memset(&presence, 0, sizeof(presence));
                presence.details = "Idle (Tidak ada dokumen)";
                presence.largeImageKey = "word_logo";
                Discord_UpdatePresence(&presence);
            }

            if (pWordApp) pWordApp->Release();
            if (pUnk)     pUnk->Release();
        }
        else {
            // Word tidak berjalan
            Discord_ClearPresence();
            std::cout << "\r[Status] Word tidak terdeteksi...        " << std::flush;
        }

        Discord_RunCallbacks();
        Sleep(2000);
    }

    Discord_Shutdown();
    CoUninitialize();
    return 0;
}