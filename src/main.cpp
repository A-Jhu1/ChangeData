#include <windows.h>
#include <comdef.h>
#if __has_include(<filesystem>)
#include <filesystem>
namespace fs = std::filesystem;
#elif __has_include(<experimental/filesystem>)
#include <experimental/filesystem>
namespace fs = std::experimental::filesystem;
#else
#error "C++ filesystem support is required. Please enable C++17 or upgrade the compiler."
#endif
#include <iostream>

// NOTE: Update the MSWORD.OLB path to match your installed Office version.
// Common paths:
//   C:\\Program Files\\Microsoft Office\\root\\Office16\\MSWORD.OLB
//   C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\MSWORD.OLB
#import "C:\\Program Files\\Microsoft Office\\root\\Office16\\MSWORD.OLB" \
    rename("ExitWindows", "WordExitWindows") \
    rename_namespace("Word")

struct ReplaceStats {
    size_t files_scanned = 0;
    size_t files_updated = 0;
};

bool HasWordExtension(const fs::path &path) {
    auto ext = path.extension().wstring();
    for (auto &ch : ext) {
        ch = static_cast<wchar_t>(towlower(ch));
    }
    return ext == L".doc" || ext == L".docx";
}

bool ReplaceInDocument(Word::_DocumentPtr &document,
                       const std::wstring &needle,
                       const std::wstring &replacement) {
    Word::RangePtr range = document->Content;
    Word::FindPtr finder = range->Find;

    finder->ClearFormatting();
    finder->Replacement->ClearFormatting();

    finder->Text = _bstr_t(needle.c_str());
    finder->Replacement->Text = _bstr_t(replacement.c_str());
    finder->Forward = VARIANT_TRUE;
    finder->Wrap = Word::WdFindWrap::wdFindContinue;
    finder->Format = VARIANT_FALSE;
    finder->MatchCase = VARIANT_FALSE;
    finder->MatchWholeWord = VARIANT_FALSE;
    finder->MatchWildcards = VARIANT_FALSE;
    finder->MatchSoundsLike = VARIANT_FALSE;
    finder->MatchAllWordForms = VARIANT_FALSE;

    VARIANT_BOOL replaced = finder->Execute(
        _variant_t(needle.c_str()),
        _variant_t(VARIANT_FALSE),
        _variant_t(VARIANT_FALSE),
        _variant_t(VARIANT_FALSE),
        _variant_t(VARIANT_FALSE),
        _variant_t(VARIANT_FALSE),
        _variant_t(VARIANT_TRUE),
        _variant_t(Word::WdFindWrap::wdFindContinue),
        _variant_t(VARIANT_FALSE),
        _variant_t(replacement.c_str()),
        _variant_t(Word::WdReplace::wdReplaceAll),
        _variant_t(VARIANT_FALSE),
        _variant_t(VARIANT_FALSE),
        _variant_t(VARIANT_FALSE),
        _variant_t(VARIANT_FALSE));

    return replaced == VARIANT_TRUE;
}

int wmain(int argc, wchar_t *argv[]) {
    if (argc < 4) {
        std::wcout << L"Usage: WordBatchReplace.exe <folder> <find> <replace>\n";
        return 1;
    }

    fs::path root = argv[1];
    std::wstring needle = argv[2];
    std::wstring replacement = argv[3];

    if (!fs::exists(root) || !fs::is_directory(root)) {
        std::wcerr << L"Error: folder does not exist: " << root << L"\n";
        return 1;
    }

    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr)) {
        std::wcerr << L"Error: CoInitialize failed.\n";
        return 1;
    }

    ReplaceStats stats{};

    try {
        Word::_ApplicationPtr app;
        hr = app.CreateInstance(L"Word.Application");
        if (FAILED(hr)) {
            std::wcerr << L"Error: cannot create Word.Application. Is Word installed?\n";
            CoUninitialize();
            return 1;
        }

        app->Visible = VARIANT_FALSE;
        app->DisplayAlerts = Word::WdAlertLevel::wdAlertsNone;

        Word::DocumentsPtr documents = app->Documents;

        for (const auto &entry : fs::recursive_directory_iterator(root)) {
            if (!entry.is_regular_file()) {
                continue;
            }

            const fs::path &path = entry.path();
            if (!HasWordExtension(path)) {
                continue;
            }

            stats.files_scanned++;

            _variant_t filename(path.c_str());
            _variant_t read_only(VARIANT_FALSE);
            _variant_t visible(VARIANT_FALSE);
            _variant_t confirm_conversions(VARIANT_FALSE);

            Word::_DocumentPtr document = documents->Open(
                filename,
                confirm_conversions,
                read_only,
                _variant_t(VARIANT_FALSE),
                _variant_t(),
                _variant_t(),
                _variant_t(),
                _variant_t(),
                _variant_t(),
                _variant_t(),
                _variant_t(),
                visible,
                _variant_t(),
                _variant_t(),
                _variant_t(),
                _variant_t());

            bool replaced = ReplaceInDocument(document, needle, replacement);
            if (replaced) {
                document->Save();
                stats.files_updated++;
            }

            document->Close(_variant_t(VARIANT_FALSE));
        }

        app->Quit(_variant_t(VARIANT_FALSE));
    } catch (const _com_error &error) {
        std::wcerr << L"COM error: " << error.ErrorMessage() << L"\n";
        CoUninitialize();
        return 1;
    }

    CoUninitialize();

    std::wcout << L"Done. Files scanned: " << stats.files_scanned
               << L", files updated: " << stats.files_updated << L"\n";
    return 0;
}
