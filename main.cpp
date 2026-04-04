#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <richedit.h>
#include <shellapi.h>
#include <string>
#include <thread>
#include <atomic>
#include <vector>
#include <algorithm>

// Menu IDs
#define IDM_FILE_NEW    101
#define IDM_FILE_OPEN   102
#define IDM_FILE_SAVE   103
#define IDM_FILE_SAVEAS 104
#define IDM_FILE_EXIT   105
#define IDM_FORMAT_FONT 106

// Control IDs
#define ID_EDIT   201
#define ID_STATUS 202

// Custom messages
#define WM_LOAD_DONE  (WM_USER + 1)
#define WM_LOAD_ERROR (WM_USER + 2)

// ---- Globals ----
static HINSTANCE g_hInst;
static HWND      g_hwndMain, g_hwndEdit, g_hwndStatus;
static HMODULE   g_hRichEdit;
static HFONT     g_hFont;
static LOGFONTW  g_logFont;
static std::wstring g_currentFile;
static bool      g_dirty   = false;
static bool      g_loading = false;

static const wchar_t* REG_KEY = L"Software\\longpad";

// ---- Forward declarations ----
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
static void UpdateTitle();
static void UpdateStatusBar();
static void MarkDirty();
static void MarkClean();
static bool ConfirmDiscard();
static void DoNew();
static void DoOpen();
static void DoSave();
static void DoSaveAs();
static void StartLoadFile(const std::wstring& path);
static void SaveToPath(const std::wstring& path);

// ---- Utility ----
static std::wstring BaseName(const std::wstring& path) {
    size_t pos = path.find_last_of(L"\\/");
    return (pos == std::wstring::npos) ? path : path.substr(pos + 1);
}

// ---- Title bar ----
static void UpdateTitle() {
    std::wstring name = g_currentFile.empty() ? L"Untitled" : BaseName(g_currentFile);
    std::wstring t = (g_dirty ? L"*" : L"") + name + L" - longpad";
    SetWindowTextW(g_hwndMain, t.c_str());
}

// ---- Status bar ----
static void UpdateStatusBar() {
    if (!g_hwndStatus || !g_hwndEdit) return;
    std::wstring fname = g_currentFile.empty() ? L"Untitled" : BaseName(g_currentFile);
    if (g_loading) {
        std::wstring msg = L"Loading " + fname + L"...";
        SendMessageW(g_hwndStatus, SB_SETTEXT, 0, (LPARAM)msg.c_str());
        SendMessageW(g_hwndStatus, SB_SETTEXT, 1, (LPARAM)L"");
        return;
    }
    CHARRANGE cr{};
    SendMessageW(g_hwndEdit, EM_EXGETSEL, 0, (LPARAM)&cr);
    LONG charIdx   = cr.cpMin;
    LONG line      = (LONG)SendMessageW(g_hwndEdit, EM_LINEFROMCHAR, (WPARAM)charIdx, 0);
    LONG lineStart = (LONG)SendMessageW(g_hwndEdit, EM_LINEINDEX, (WPARAM)line, 0);
    LONG col       = charIdx - lineStart;
    wchar_t pos[64];
    swprintf_s(pos, L"Ln %ld, Col %ld", line + 1, col + 1);
    SendMessageW(g_hwndStatus, SB_SETTEXT, 0, (LPARAM)fname.c_str());
    SendMessageW(g_hwndStatus, SB_SETTEXT, 1, (LPARAM)pos);
}

static void SetStatusBarParts(int totalWidth) {
    int colW = 150;
    int p[] = { totalWidth - colW, -1 };
    SendMessageW(g_hwndStatus, SB_SETPARTS, 2, (LPARAM)p);
}

// ---- Dirty state ----
static void MarkDirty() {
    if (!g_dirty && !g_loading) {
        g_dirty = true;
        UpdateTitle();
    }
}
static void MarkClean() {
    g_dirty = false;
    UpdateTitle();
}

static bool ConfirmDiscard() {
    if (!g_dirty) return true;
    std::wstring name = g_currentFile.empty() ? L"Untitled" : BaseName(g_currentFile);
    std::wstring msg  = L"Save changes to \"" + name + L"\"?";
    int r = MessageBoxW(g_hwndMain, msg.c_str(), L"longpad", MB_YESNOCANCEL | MB_ICONQUESTION);
    if (r == IDCANCEL) return false;
    if (r == IDYES)    DoSave();
    return !g_dirty; // if save succeeded, dirty is cleared
}

// ---- Background file load ----
static void LoadThread(HWND hwnd, std::wstring path) {
    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ,
        nullptr, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        PostMessageW(hwnd, WM_LOAD_ERROR, GetLastError(), 0);
        return;
    }

    LARGE_INTEGER sz{};
    if (!GetFileSizeEx(hFile, &sz)) {
        CloseHandle(hFile);
        PostMessageW(hwnd, WM_LOAD_ERROR, GetLastError(), 0);
        return;
    }

    // Refuse files over 1 GB to avoid silent OOM
    if (sz.QuadPart > (LONGLONG)1 * 1024 * 1024 * 1024) {
        CloseHandle(hFile);
        PostMessageW(hwnd, WM_LOAD_ERROR, ERROR_FILE_TOO_LARGE, 0);
        return;
    }

    DWORD fileSize = (DWORD)sz.QuadPart;

    // BOM detection
    BYTE bom[3]{};
    DWORD bomRead = 0;
    ReadFile(hFile, bom, 3, &bomRead, nullptr);
    UINT  cp        = CP_UTF8;
    DWORD dataStart = 0;
    if (bomRead >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) {
        dataStart = 3; // skip UTF-8 BOM
    }
    SetFilePointer(hFile, (LONG)dataStart, nullptr, FILE_BEGIN);
    DWORD dataSize = fileSize - dataStart;

    // Read entire file into memory (disk I/O on this thread, not main thread)
    std::string raw(dataSize, '\0');
    DWORD totalRead = 0;
    while (totalRead < dataSize) {
        DWORD toRead = (std::min)(dataSize - totalRead, (DWORD)(1024 * 1024));
        DWORD read   = 0;
        if (!ReadFile(hFile, raw.data() + totalRead, toRead, &read, nullptr) || read == 0)
            break;
        totalRead += read;
    }
    CloseHandle(hFile);
    raw.resize(totalRead);

    // Convert to UTF-16 (wstring)
    int wlen = MultiByteToWideChar(cp, 0, raw.c_str(), (int)raw.size(), nullptr, 0);
    if (wlen < 0) wlen = 0;
    auto* wtext = new std::wstring(wlen, L'\0');
    if (wlen > 0)
        MultiByteToWideChar(cp, 0, raw.c_str(), (int)raw.size(), wtext->data(), wlen);

    PostMessageW(hwnd, WM_LOAD_DONE, 0, (LPARAM)wtext);
}

// ---- In-memory EM_STREAMIN callback ----
struct MemStream { const BYTE* ptr; LONG size; LONG pos; };

static DWORD CALLBACK MemStreamInCb(DWORD_PTR cookie, LPBYTE buf, LONG cb, LONG* pcb) {
    auto* ms   = reinterpret_cast<MemStream*>(cookie);
    LONG  avail = ms->size - ms->pos;
    LONG  n     = (std::min)(cb, avail);
    if (n > 0) {
        memcpy(buf, ms->ptr + ms->pos, n);
        ms->pos += n;
    }
    *pcb = n;
    return 0;
}

// ---- EM_STREAMOUT → UTF-8 file callback ----
struct FileStream { HANDLE hFile; bool ok; };

static DWORD CALLBACK FileStreamOutCb(DWORD_PTR cookie, LPBYTE buf, LONG cb, LONG* pcb) {
    auto* fs    = reinterpret_cast<FileStream*>(cookie);
    int   wchars = cb / 2;
    if (wchars <= 0) { *pcb = 0; return 0; }
    const wchar_t* wbuf = reinterpret_cast<const wchar_t*>(buf);
    int u8len = WideCharToMultiByte(CP_UTF8, 0, wbuf, wchars, nullptr, 0, nullptr, nullptr);
    std::string u8(u8len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, wbuf, wchars, u8.data(), u8len, nullptr, nullptr);
    DWORD written = 0;
    if (!WriteFile(fs->hFile, u8.data(), (DWORD)u8.size(), &written, nullptr)
        || written != (DWORD)u8.size()) {
        fs->ok = false;
        return 1; // signal error to RichEdit
    }
    *pcb = cb;
    return 0;
}

// ---- Save ----
static void SaveToPath(const std::wstring& path) {
    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr,
        CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        MessageBoxW(g_hwndMain, L"Could not open file for writing.", L"longpad", MB_ICONERROR);
        return;
    }
    FileStream  fs{hFile, true};
    EDITSTREAM  es{(DWORD_PTR)&fs, 0, FileStreamOutCb};
    SendMessageW(g_hwndEdit, EM_STREAMOUT, (WPARAM)(SF_TEXT | SF_UNICODE), (LPARAM)&es);
    CloseHandle(hFile);
    if (!fs.ok) {
        MessageBoxW(g_hwndMain, L"Error writing file.", L"longpad", MB_ICONERROR);
        return;
    }
    g_currentFile = path;
    MarkClean();
    UpdateStatusBar();
}

static void DoSave() {
    if (g_currentFile.empty()) { DoSaveAs(); return; }
    SaveToPath(g_currentFile);
}

static void DoSaveAs() {
    wchar_t path[MAX_PATH]{};
    if (!g_currentFile.empty())
        wcscpy_s(path, g_currentFile.c_str());
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = g_hwndMain;
    ofn.lpstrFile   = path;
    ofn.nMaxFile    = MAX_PATH;
    ofn.lpstrFilter = L"Text Files\0*.txt\0All Files\0*.*\0\0";
    ofn.lpstrDefExt = L"txt";
    ofn.Flags       = OFN_OVERWRITEPROMPT | OFN_HIDEREADONLY;
    if (!GetSaveFileNameW(&ofn)) return;
    SaveToPath(path);
}

// ---- Load (kicks off background thread) ----
static void StartLoadFile(const std::wstring& path) {
    g_loading = true;
    // Suppress change notifications during reset
    SendMessageW(g_hwndEdit, EM_SETEVENTMASK, 0, 0);
    SendMessageW(g_hwndEdit, WM_SETTEXT, 0, (LPARAM)L"");
    SendMessageW(g_hwndEdit, EM_EMPTYUNDOBUFFER, 0, 0);
    g_currentFile = path;
    g_dirty       = false;
    UpdateTitle();
    // Make read-only during load (avoids grayed-out appearance from EnableWindow)
    SendMessageW(g_hwndEdit, EM_SETREADONLY, TRUE, 0);
    UpdateStatusBar(); // shows "Loading ..."
    std::thread(LoadThread, g_hwndMain, path).detach();
}

// ---- New / Open ----
static void DoNew() {
    if (!ConfirmDiscard()) return;
    g_loading = true;
    SendMessageW(g_hwndEdit, EM_SETEVENTMASK, 0, 0);
    SendMessageW(g_hwndEdit, EM_SETREADONLY, FALSE, 0);
    SendMessageW(g_hwndEdit, WM_SETTEXT, 0, (LPARAM)L"");
    SendMessageW(g_hwndEdit, EM_EMPTYUNDOBUFFER, 0, 0);
    g_loading = false;
    SendMessageW(g_hwndEdit, EM_SETEVENTMASK, 0, ENM_CHANGE | ENM_SELCHANGE);
    g_currentFile.clear();
    MarkClean();
    UpdateStatusBar();
    SetFocus(g_hwndEdit);
}

static void DoOpen() {
    if (!ConfirmDiscard()) return;
    wchar_t path[MAX_PATH]{};
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = g_hwndMain;
    ofn.lpstrFile   = path;
    ofn.nMaxFile    = MAX_PATH;
    ofn.lpstrFilter = L"All Files\0*.*\0Text Files\0*.txt\0\0";
    ofn.Flags       = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
    if (!GetOpenFileNameW(&ofn)) return;
    StartLoadFile(path);
}

// ---- Font persistence ----
static void LoadFontFromRegistry() {
    memset(&g_logFont, 0, sizeof(g_logFont));
    g_logFont.lfHeight         = -14;
    g_logFont.lfWeight         = FW_NORMAL;
    g_logFont.lfCharSet        = DEFAULT_CHARSET;
    g_logFont.lfOutPrecision   = OUT_DEFAULT_PRECIS;
    g_logFont.lfClipPrecision  = CLIP_DEFAULT_PRECIS;
    g_logFont.lfQuality        = CLEARTYPE_QUALITY;
    g_logFont.lfPitchAndFamily = FIXED_PITCH | FF_MODERN;
    wcscpy_s(g_logFont.lfFaceName, L"Consolas");

    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, REG_KEY, 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return;

    DWORD type, size;
    wchar_t faceName[LF_FACESIZE]{};
    size = sizeof(faceName);
    if (RegQueryValueExW(hKey, L"FontFaceName", nullptr, &type, (BYTE*)faceName, &size) == ERROR_SUCCESS && type == REG_SZ)
        wcscpy_s(g_logFont.lfFaceName, faceName);

    DWORD val;
    size = sizeof(val);
    if (RegQueryValueExW(hKey, L"FontHeight", nullptr, &type, (BYTE*)&val, &size) == ERROR_SUCCESS && type == REG_DWORD)
        g_logFont.lfHeight = (LONG)val;
    size = sizeof(val);
    if (RegQueryValueExW(hKey, L"FontWeight", nullptr, &type, (BYTE*)&val, &size) == ERROR_SUCCESS && type == REG_DWORD)
        g_logFont.lfWeight = (LONG)val;
    size = sizeof(val);
    if (RegQueryValueExW(hKey, L"FontItalic", nullptr, &type, (BYTE*)&val, &size) == ERROR_SUCCESS && type == REG_DWORD)
        g_logFont.lfItalic = (BYTE)val;
    size = sizeof(val);
    if (RegQueryValueExW(hKey, L"FontCharSet", nullptr, &type, (BYTE*)&val, &size) == ERROR_SUCCESS && type == REG_DWORD)
        g_logFont.lfCharSet = (BYTE)val;

    RegCloseKey(hKey);
}

static void SaveFontToRegistry() {
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, REG_KEY, 0, nullptr, 0, KEY_WRITE, nullptr, &hKey, nullptr) != ERROR_SUCCESS)
        return;

    RegSetValueExW(hKey, L"FontFaceName", 0, REG_SZ, (const BYTE*)g_logFont.lfFaceName,
        (DWORD)((wcslen(g_logFont.lfFaceName) + 1) * sizeof(wchar_t)));
    DWORD val = (DWORD)g_logFont.lfHeight;
    RegSetValueExW(hKey, L"FontHeight", 0, REG_DWORD, (const BYTE*)&val, sizeof(val));
    val = (DWORD)g_logFont.lfWeight;
    RegSetValueExW(hKey, L"FontWeight", 0, REG_DWORD, (const BYTE*)&val, sizeof(val));
    val = g_logFont.lfItalic;
    RegSetValueExW(hKey, L"FontItalic", 0, REG_DWORD, (const BYTE*)&val, sizeof(val));
    val = g_logFont.lfCharSet;
    RegSetValueExW(hKey, L"FontCharSet", 0, REG_DWORD, (const BYTE*)&val, sizeof(val));

    RegCloseKey(hKey);
}

static void ApplyFont() {
    HFONT hNew = CreateFontIndirectW(&g_logFont);
    if (!hNew) return;
    SendMessageW(g_hwndEdit, WM_SETFONT, (WPARAM)hNew, TRUE);
    if (g_hFont) DeleteObject(g_hFont);
    g_hFont = hNew;
}

static void DoChooseFont() {
    LOGFONTW lf = g_logFont;
    CHOOSEFONTW cf{};
    cf.lStructSize = sizeof(cf);
    cf.hwndOwner   = g_hwndMain;
    cf.lpLogFont   = &lf;
    cf.Flags       = CF_SCREENFONTS | CF_INITTOLOGFONTSTRUCT | CF_FORCEFONTEXIST;
    if (!ChooseFontW(&cf)) return;
    g_logFont = lf;
    ApplyFont();
    SaveFontToRegistry();
}

// ---- Accelerator table ----
static HACCEL CreateAccel() {
    ACCEL accels[] = {
        { FVIRTKEY | FCONTROL,         'N', IDM_FILE_NEW    },
        { FVIRTKEY | FCONTROL,         'O', IDM_FILE_OPEN   },
        { FVIRTKEY | FCONTROL,         'S', IDM_FILE_SAVE   },
        { FVIRTKEY | FCONTROL | FSHIFT,'S', IDM_FILE_SAVEAS },
        { FVIRTKEY | FCONTROL,         'Q', IDM_FILE_EXIT   },
    };
    return CreateAcceleratorTableW(accels, (int)ARRAYSIZE(accels));
}

// ---- Window procedure ----
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {

    case WM_CREATE: {
        if (!(g_hRichEdit = LoadLibraryW(L"riched20.dll"))) {
            MessageBoxW(nullptr, L"Could not load riched20.dll", L"longpad", MB_ICONERROR);
            return -1;
        }

        // Status bar
        g_hwndStatus = CreateWindowExW(0, STATUSCLASSNAME, nullptr,
            WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
            0, 0, 0, 0, hwnd, (HMENU)(INT_PTR)ID_STATUS, g_hInst, nullptr);
        SetStatusBarParts(800);

        // RichEdit control
        g_hwndEdit = CreateWindowExW(0, RICHEDIT_CLASS, nullptr,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL |
            ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_NOHIDESEL,
            0, 0, 0, 0,
            hwnd, (HMENU)(INT_PTR)ID_EDIT, g_hInst, nullptr);

        // Remove 64 KB limit
        SendMessageW(g_hwndEdit, EM_EXLIMITTEXT, 0, (LPARAM)0x7FFFFFFF);

        // Font: load saved preference or default to Consolas 10pt
        LoadFontFromRegistry();
        ApplyFont();

        // Enable RichEdit-specific notifications
        SendMessageW(g_hwndEdit, EM_SETEVENTMASK, 0, ENM_CHANGE | ENM_SELCHANGE);

        // Menu
        HMENU hMenu = CreateMenu();
        HMENU hFile = CreatePopupMenu();
        AppendMenuW(hFile, MF_STRING, IDM_FILE_NEW,    L"&New\tCtrl+N");
        AppendMenuW(hFile, MF_STRING, IDM_FILE_OPEN,   L"&Open...\tCtrl+O");
        AppendMenuW(hFile, MF_STRING, IDM_FILE_SAVE,   L"&Save\tCtrl+S");
        AppendMenuW(hFile, MF_STRING, IDM_FILE_SAVEAS, L"Save &As...\tCtrl+Shift+S");
        AppendMenuW(hFile, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(hFile, MF_STRING, IDM_FILE_EXIT,   L"E&xit");
        AppendMenuW(hMenu, MF_POPUP, (UINT_PTR)hFile, L"&File");
        HMENU hFormat = CreatePopupMenu();
        AppendMenuW(hFormat, MF_STRING, IDM_FORMAT_FONT, L"&Font...");
        AppendMenuW(hMenu, MF_POPUP, (UINT_PTR)hFormat, L"F&ormat");
        SetMenu(hwnd, hMenu);

        UpdateTitle();
        UpdateStatusBar();
        return 0;
    }

    case WM_SIZE: {
        if (g_hwndStatus) {
            SendMessageW(g_hwndStatus, WM_SIZE, 0, 0);
            SetStatusBarParts((int)LOWORD(lp));
            RECT rc{};
            GetWindowRect(g_hwndStatus, &rc);
            int statusH = rc.bottom - rc.top;
            if (g_hwndEdit)
                MoveWindow(g_hwndEdit, 0, 0, (int)LOWORD(lp), (int)HIWORD(lp) - statusH, TRUE);
        }
        return 0;
    }

    case WM_SETFOCUS:
        if (g_hwndEdit) SetFocus(g_hwndEdit);
        return 0;

    case WM_COMMAND:
        if (LOWORD(wp) == ID_EDIT && HIWORD(wp) == EN_CHANGE) {
            MarkDirty();
        } else {
            switch (LOWORD(wp)) {
            case IDM_FILE_NEW:    DoNew();       break;
            case IDM_FILE_OPEN:   DoOpen();      break;
            case IDM_FILE_SAVE:   DoSave();      break;
            case IDM_FILE_SAVEAS: DoSaveAs();    break;
            case IDM_FILE_EXIT:   SendMessageW(hwnd, WM_CLOSE, 0, 0); break;
            case IDM_FORMAT_FONT: DoChooseFont(); break;
            }
        }
        return 0;

    case WM_NOTIFY: {
        auto* nm = reinterpret_cast<NMHDR*>(lp);
        if (nm->idFrom == ID_EDIT && nm->code == EN_SELCHANGE)
            UpdateStatusBar();
        return 0;
    }

    case WM_LOAD_DONE: {
        auto* wtext = reinterpret_cast<std::wstring*>(lp);
        if (wtext && !wtext->empty()) {
            MemStream ms{
                reinterpret_cast<const BYTE*>(wtext->data()),
                (LONG)(wtext->size() * sizeof(wchar_t)),
                0
            };
            EDITSTREAM es{(DWORD_PTR)&ms, 0, MemStreamInCb};
            SendMessageW(g_hwndEdit, EM_STREAMIN, (WPARAM)(SF_TEXT | SF_UNICODE), (LPARAM)&es);
        }
        delete wtext;
        SendMessageW(g_hwndEdit, EM_EMPTYUNDOBUFFER, 0, 0);
        SendMessageW(g_hwndEdit, EM_SETSEL, 0, 0);
        SendMessageW(g_hwndEdit, EM_SETREADONLY, FALSE, 0);
        SendMessageW(g_hwndEdit, EM_SETEVENTMASK, 0, ENM_CHANGE | ENM_SELCHANGE);
        g_loading = false;
        SetFocus(g_hwndEdit);
        MarkClean();
        UpdateStatusBar();
        return 0;
    }

    case WM_LOAD_ERROR: {
        g_loading = false;
        SendMessageW(g_hwndEdit, EM_SETREADONLY, FALSE, 0);
        SendMessageW(g_hwndEdit, EM_SETEVENTMASK, 0, ENM_CHANGE | ENM_SELCHANGE);
        DWORD err = (DWORD)wp;
        wchar_t errMsg[256];
        if (err == ERROR_FILE_TOO_LARGE)
            wcscpy_s(errMsg, L"File exceeds the 1 GB size limit.");
        else
            swprintf_s(errMsg, L"Failed to open file (error %lu).", err);
        MessageBoxW(hwnd, errMsg, L"longpad", MB_ICONERROR);
        UpdateStatusBar();
        return 0;
    }

    case WM_CLOSE:
        if (!ConfirmDiscard()) return 0;
        DestroyWindow(hwnd);
        return 0;

    case WM_DESTROY:
        if (g_hFont)    { DeleteObject(g_hFont);         g_hFont    = nullptr; }
        if (g_hRichEdit){ FreeLibrary(g_hRichEdit);      g_hRichEdit = nullptr; }
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// ---- Entry point ----
int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, PWSTR, int nCmdShow) {
    g_hInst = hInst;

    INITCOMMONCONTROLSEX icc{ sizeof(icc), ICC_BAR_CLASSES };
    InitCommonControlsEx(&icc);

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInst;
    wc.hCursor       = LoadCursorW(nullptr, IDC_IBEAM);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"LongpadWnd";
    wc.hIcon         = LoadIconW(nullptr, IDI_APPLICATION);
    if (!RegisterClassExW(&wc)) return 1;

    HWND hwnd = CreateWindowExW(0, L"LongpadWnd", L"longpad",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 900, 650,
        nullptr, nullptr, hInst, nullptr);
    if (!hwnd) return 1;
    g_hwndMain = hwnd;

    // Show the window BEFORE any file I/O
    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    // Open file from command line if provided
    int   argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv && argc > 1)
        StartLoadFile(argv[1]);
    LocalFree(argv);

    HACCEL hAccel = CreateAccel();
    MSG    msg{};
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        if (!TranslateAcceleratorW(hwnd, hAccel, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    DestroyAcceleratorTable(hAccel);
    return (int)msg.wParam;
}
