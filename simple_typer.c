/*
 * simple_typer.c - Simple Typer v2.11
 *
 * Each button types its stored text into whatever window had focus
 * before the launcher was clicked.
 *
 * Features:
 *   - INI-configured buttons; each button types text into the last focused window
 *   - Optional per-button custom border color (via color picker)
 *   - Dark mode, configurable font size and window width
 *   - Always on top, minimize to system tray
 *   - Right-click buttons to Edit / Delete / Duplicate / Move Up / Move Down
 *   - Separator (divider) lines between buttons
 *   - Optional custom ICO icon per button
 *   - Remembers last window position
 *   - Window opacity and configurable title
 *   - Button tooltips showing a preview of stored text on hover
 *   - Variables: {date} {time} {clipboard} expand automatically when typed
 *   - Fill-in-the-blank: {?} in button text pops a small input box first
 *   - Global text prefix and suffix appended to every button's output
 *   - Search/filter bar - type to instantly filter buttons by name
 *   - Compact mode - icon-only grid palette for a tiny always-on-top layout
 *   - Categories - collapsible group headers to organise buttons
 *   - Keyboard shortcuts - optional global hotkey per button
 *   - Multiple profiles - switchable INI sets from tray or Profiles menu
 *   - System key tokens - {tab} {enter} {esc} etc. send keystrokes mid-text
 *   - Version 2.11
 *
 * Compile:
 *
 *   cl typer.c typer.res /link user32.lib shell32.lib comdlg32.lib gdi32.lib dwmapi.lib comctl32.lib /subsystem:windows /out:simple_typer.exe
 */

#include <windows.h>
#include <commctrl.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <commdlg.h>
#include <dwmapi.h>
#include <shellapi.h>

/* ── Forward declarations ────────────────────────────────────────────── */
static void RefreshMainWindow(void);
static void RebuildMenu(void);
static void ApplyDarkBackground(void);
static void ApplyOpacity(void);
static void SaveAll(void);
static void LoadSettings(void);
static void LoadButtons(void);
static void LoadButtonIcons(void);
static void FreeIcons(void);
static void RecreateFont(void);
static void SetTitleBarDark(HWND hwnd, int dark);
static void RegisterAllHotkeys(void);
static void UnregisterAllHotkeys(void);
static void FireButton(int idx);   /* shared by click and hotkey */
static void BuildFireActions(const char *text);
static void FireNextAction(void);
static void ScanProfiles(void);
static void SwitchProfile(int idx);

/* ── IDs ─────────────────────────────────────────────────────────────── */
#define ID_ADD_BTN            1
#define ID_BUTTON_BASE        100
#define MAX_BUTTONS           64
#define IDI_APPICON           200
#define MAX_TEXT              4096
#define COMPACT_BTN_SZ        28
#define COMPACT_BTN_GAP       4
#define ID_HOTKEY_BASE        500   /* RegisterHotKey IDs 500..563 */

/* Dialog controls */
#define IDC_NAME_EDIT         10
#define IDC_TEXT_EDIT         11
#define IDC_OK                14
#define IDC_CANCEL            15
#define IDC_DARK_CHECK        17
#define IDC_TOPMOST_CHECK     18
#define IDC_TRAY_CHECK        19
#define IDC_INFO_TEXT         30
#define IDC_INFO_OK           31
#define IDC_FONT_EDIT         32
#define IDC_WIDTH_EDIT        33
#define IDC_SEP_CHECK         35
#define IDC_ICON_CHECK        36
#define IDC_ICON_PATH_EDIT    37
#define IDC_ICON_BROWSE       38
#define IDC_OPACITY_EDIT      39
#define IDC_TITLE_EDIT        40
#define IDC_BTN_COLOR_CHECK   41
#define IDC_BTN_COLOR_BTN     42
#define IDC_PREFIX_EDIT       43
#define IDC_SUFFIX_EDIT       44
#define IDC_COMPACT_CHECK     45
#define IDC_CAT_CHECK         46   /* "Category header" checkbox */
#define IDC_HK_CTRL           47
#define IDC_HK_SHIFT          48
#define IDC_HK_ALT            49
#define IDC_SEARCH_EDIT       60
#define IDC_HK_KEY            61   /* combobox for hotkey key */
#define IDC_HK_CLEAR          62

/* Prompt dialog controls */
#define IDC_PROMPT_EDIT       50
#define IDC_PROMPT_OK         51
#define IDC_PROMPT_CANCEL     52

/* Menu IDs */
#define ID_HELP_INSTRUCTIONS  20
#define ID_HELP_ABOUT         21
#define ID_SETTINGS           22
#define ID_PROFILES_MENU      23

/* Right-click context menu */
#define IDM_MOVE_UP           300
#define IDM_MOVE_DOWN         301
#define IDM_EDIT_BTN          302
#define IDM_DELETE_BTN        303
#define IDM_DUPLICATE_BTN     304

/* System tray */
#define WM_TRAYICON           (WM_APP + 1)
#define ID_TRAY_ICON          1
#define IDM_TRAY_RESTORE      400
#define IDM_TRAY_EXIT         401

/* Profiles */
#define IDM_PROFILE_BASE      600   /* 600..615 → switch to profile N */
#define IDM_PROFILE_NEW       616
#define IDM_PROFILE_DELETE    617
#define MAX_PROFILES          16

/* ── Dark mode colors ────────────────────────────────────────────────── */
#define DK_BG       RGB( 28,  28,  28)
#define DK_BTN      RGB( 48,  48,  48)
#define DK_BTN_PRE  RGB( 68,  68,  68)
#define DK_BORDER   RGB( 85,  85,  85)
#define DK_TEXT     RGB(220, 220, 220)
#define DK_MENU_BG  RGB( 32,  32,  32)
#define DK_MENU_HOT RGB( 55,  55,  55)
#define DK_SEP      RGB( 70,  70,  70)
#define LT_SEP      RGB(160, 160, 160)
#define DK_SEARCH   RGB( 40,  40,  40)
#define DK_CAT_BG   RGB( 50,  75, 110)   /* category header — dark mode */
#define LT_CAT_BG   RGB(200, 220, 245)   /* category header — light mode */
#define DK_CAT_TEXT RGB(200, 225, 255)
#define LT_CAT_TEXT RGB( 20,  50, 100)

static const char *g_menuLabels[] = { "Instructions", "Settings", "Profiles", "About" };
static const UINT  g_menuIDs[]    = { ID_HELP_INSTRUCTIONS, ID_SETTINGS, ID_PROFILES_MENU, ID_HELP_ABOUT };

/* ── Hotkey key table ────────────────────────────────────────────────── */
static const char *g_hkNames[] = {
    "(none)",
    "A","B","C","D","E","F","G","H","I","J","K","L","M",
    "N","O","P","Q","R","S","T","U","V","W","X","Y","Z",
    "0","1","2","3","4","5","6","7","8","9",
    "F1","F2","F3","F4","F5","F6","F7","F8","F9","F10","F11","F12",
    NULL
};
static const int g_hkVKs[] = {
    0,
    'A','B','C','D','E','F','G','H','I','J','K','L','M',
    'N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
    '0','1','2','3','4','5','6','7','8','9',
    VK_F1,VK_F2,VK_F3,VK_F4,VK_F5,VK_F6,
    VK_F7,VK_F8,VK_F9,VK_F10,VK_F11,VK_F12
};
#define HK_KEY_COUNT 50   /* elements in above arrays */

/* ── Data ────────────────────────────────────────────────────────────── */
typedef struct {
    char     name[256];
    char     text[MAX_TEXT];
    char     iconPath[MAX_PATH];
    COLORREF btnColor;
    int      hasColor;
    int      isSeparator;
    int      showIcon;
    int      isCategory;   /* collapsible group header */
    int      hotkeyMod;    /* MOD_CONTROL | MOD_SHIFT | MOD_ALT  (0 = none) */
    int      hotkeyVk;     /* virtual key code                    (0 = none) */
} ButtonConfig;

static ButtonConfig g_buttons[MAX_BUTTONS];
static HICON        g_icons[MAX_BUTTONS];
static int          g_count         = 0;
static int          g_collapsed[MAX_BUTTONS]; /* 1 = category header is collapsed */

/* Settings */
static int          g_darkMode    = 0;
static int          g_alwaysOnTop = 0;
static int          g_minToTray   = 0;
static int          g_fontSize    = 9;
static int          g_winWidth    = 300;
static int          g_winX        = -1;
static int          g_winY        = -1;
static int          g_opacity     = 100;
static int          g_compactMode = 0;
static char         g_winTitle[256] = "Simple Typer";
static char         g_prefix[512]   = "";
static char         g_suffix[512]   = "";

/* Runtime */
static HWND         g_hwndMain;
static HWND         g_hwndBtns[MAX_BUTTONS];
static HINSTANCE    g_hInst;
static char         g_basePath[MAX_PATH];        /* always typer.ini */
static char         g_iniPath[MAX_PATH];          /* active profile */
static char         g_profileNames[MAX_PROFILES][64];
static char         g_profilePaths[MAX_PROFILES][MAX_PATH];
static int          g_profileCount  = 0;
static int          g_activeProfile = 0;
static HWND         g_hwndDlg     = NULL;
static HBRUSH       g_hbrDkBg     = NULL;
static HFONT        g_hFont       = NULL;
static HFONT        g_hFontBold   = NULL;   /* bold font for category headers */
static int          g_trayAdded   = 0;
static int          g_editIndex   = -1;
static int          g_ctxIndex    = -1;
static HWND         g_prevWindow  = NULL;
static int          g_pendingIdx  = -1;
static COLORREF     g_settingBtnColor;
static COLORREF     g_customColors[16];
static HWND         g_hwndTooltip = NULL;
static char         g_tooltipText[MAX_BUTTONS][128];
static HWND         g_hwndSearch  = NULL;
static char         g_filterText[256] = "";
static HBRUSH       g_hbrSearchDk = NULL;

/* Fill-in-the-blank prompt */
static char         g_promptResult[512];
static int          g_promptDone      = 0;
static int          g_promptCancelled = 0;

static const char  *g_infoDlgTitle   = NULL;
static const char  *g_infoDlgContent = NULL;

/* ── System-key token table ──────────────────────────────────────────── */
/* Tokens the user can embed in button text to send a keystroke.          */
static const struct { const char *tok; int tokLen; int vk; } g_keyTokens[] = {
    {"{tab}",        5,  VK_TAB},
    {"{enter}",      7,  VK_RETURN},
    {"{return}",     8,  VK_RETURN},
    {"{esc}",        5,  VK_ESCAPE},
    {"{escape}",     8,  VK_ESCAPE},
    {"{backspace}", 11,  VK_BACK},
    {"{delete}",     8,  VK_DELETE},
    {"{del}",        5,  VK_DELETE},
    {"{up}",         4,  VK_UP},
    {"{down}",       6,  VK_DOWN},
    {"{left}",       6,  VK_LEFT},
    {"{right}",      7,  VK_RIGHT},
    {"{home}",       6,  VK_HOME},
    {"{end}",        5,  VK_END},
    {"{pgup}",       6,  VK_PRIOR},
    {"{pgdn}",       6,  VK_NEXT},
    {NULL, 0, 0}
};

/* ── Fire-action list (text segments + key presses) ──────────────────── */
#define ACT_TEXT  0
#define ACT_KEY   1
#define MAX_FIRE_ACTIONS 64

typedef struct {
    int  type;           /* ACT_TEXT or ACT_KEY */
    char text[MAX_TEXT]; /* used when type == ACT_TEXT */
    int  vk;             /* used when type == ACT_KEY  */
} FireAction;

static FireAction g_fireActions[MAX_FIRE_ACTIONS];
static int        g_fireCount  = 0;
static int        g_fireIdx    = 0;
static int        g_pendingVk  = 0;

/* ── DWM dark title bar ──────────────────────────────────────────────── */
static void SetTitleBarDark(HWND hwnd, int dark)
{
    BOOL val = dark ? TRUE : FALSE;
    if (FAILED(DwmSetWindowAttribute(hwnd, 20, &val, sizeof(val))))
        DwmSetWindowAttribute(hwnd, 19, &val, sizeof(val));
}

/* ── Opacity ─────────────────────────────────────────────────────────── */
static void ApplyOpacity(void)
{
    LONG ex = GetWindowLong(g_hwndMain, GWL_EXSTYLE);
    if (g_opacity < 100) {
        SetWindowLong(g_hwndMain, GWL_EXSTYLE, ex | WS_EX_LAYERED);
        SetLayeredWindowAttributes(g_hwndMain, 0,
                                   (BYTE)(g_opacity * 255 / 100), LWA_ALPHA);
    } else {
        SetWindowLong(g_hwndMain, GWL_EXSTYLE, ex & ~WS_EX_LAYERED);
    }
}

/* ── Fonts ───────────────────────────────────────────────────────────── */
static void RecreateFont(void)
{
    if (g_hFont)     { DeleteObject(g_hFont);     g_hFont     = NULL; }
    if (g_hFontBold) { DeleteObject(g_hFontBold); g_hFontBold = NULL; }
    HDC hdc = GetDC(NULL);
    int h = -MulDiv(g_fontSize, GetDeviceCaps(hdc, LOGPIXELSY), 72);
    ReleaseDC(NULL, hdc);
    g_hFont = CreateFont(h, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                         DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                         CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, "Segoe UI");
    g_hFontBold = CreateFont(h, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
                         DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                         CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, "Segoe UI");
}

/* ── Hotkeys ─────────────────────────────────────────────────────────── */
static void UnregisterAllHotkeys(void)
{
    for (int i = 0; i < MAX_BUTTONS; i++)
        UnregisterHotKey(g_hwndMain, ID_HOTKEY_BASE + i);
}

static void RegisterAllHotkeys(void)
{
    UnregisterAllHotkeys();
    for (int i = 0; i < g_count; i++) {
        if (g_buttons[i].hotkeyVk && !g_buttons[i].isSeparator && !g_buttons[i].isCategory)
            RegisterHotKey(g_hwndMain, ID_HOTKEY_BASE + i,
                           (UINT)g_buttons[i].hotkeyMod, (UINT)g_buttons[i].hotkeyVk);
    }
}

/* Build a short human-readable hotkey string, e.g. "Ctrl+Shift+A" */
static void HotkeyToString(int mod, int vk, char *buf, int bufSize)
{
    buf[0] = '\0';
    if (!vk) { strcpy(buf, "(none)"); return; }
    if (mod & MOD_CONTROL) strcat(buf, "Ctrl+");
    if (mod & MOD_SHIFT)   strcat(buf, "Shift+");
    if (mod & MOD_ALT)     strcat(buf, "Alt+");
    /* find key name */
    for (int i = 0; i < HK_KEY_COUNT; i++) {
        if (g_hkVKs[i] == vk) { strcat(buf, g_hkNames[i]); return; }
    }
    /* fallback */
    char tmp[8]; sprintf(tmp, "0x%02X", vk); strcat(buf, tmp);
}

/* ── Tray ────────────────────────────────────────────────────────────── */
static void AddTrayIcon(void)
{
    if (g_trayAdded) return;
    NOTIFYICONDATA nid = {0};
    nid.cbSize = sizeof(nid); nid.hWnd = g_hwndMain; nid.uID = ID_TRAY_ICON;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.hIcon  = LoadIcon(g_hInst, MAKEINTRESOURCE(IDI_APPICON));
    nid.uCallbackMessage = WM_TRAYICON;
    strcpy(nid.szTip, "Simple Typer");
    Shell_NotifyIcon(NIM_ADD, &nid);
    g_trayAdded = 1;
}

static void RemoveTrayIcon(void)
{
    if (!g_trayAdded) return;
    NOTIFYICONDATA nid = {0};
    nid.cbSize = sizeof(nid); nid.hWnd = g_hwndMain; nid.uID = ID_TRAY_ICON;
    Shell_NotifyIcon(NIM_DELETE, &nid);
    g_trayAdded = 0;
}

/* ── Profiles ────────────────────────────────────────────────────────── */
static void ScanProfiles(void)
{
    g_profileCount = 0;
    /* Profile 0 is always "Default" = g_basePath */
    strcpy(g_profileNames[0], "Default");
    strcpy(g_profilePaths[0], g_basePath);
    g_profileCount = 1;

    /* Collect typer_*.ini files in the same directory */
    char dir[MAX_PATH];
    strcpy(dir, g_basePath);
    char *ls = strrchr(dir, '\\'); if (!ls) ls = strrchr(dir, '/');
    if (ls) *(ls + 1) = '\0'; else dir[0] = '\0';

    char pat[MAX_PATH];
    sprintf(pat, "%styper_*.ini", dir);
    WIN32_FIND_DATA fd;
    HANDLE h = FindFirstFile(pat, &fd);
    if (h != INVALID_HANDLE_VALUE) {
        do {
            if (g_profileCount >= MAX_PROFILES) break;
            /* strip "typer_" prefix (6 chars) and ".ini" suffix */
            char nm[64];
            strncpy(nm, fd.cFileName + 6, 63); nm[63] = '\0';
            char *dot = strrchr(nm, '.'); if (dot) *dot = '\0';
            if (!nm[0]) continue;
            strcpy(g_profileNames[g_profileCount], nm);
            sprintf(g_profilePaths[g_profileCount], "%s%s", dir, fd.cFileName);
            g_profileCount++;
        } while (FindNextFile(h, &fd));
        FindClose(h);
    }

    /* Read which profile is active from [Meta] in the base INI */
    char active[64];
    GetPrivateProfileString("Meta", "ActiveProfile", "Default",
                            active, sizeof(active), g_basePath);
    g_activeProfile = 0;
    for (int i = 0; i < g_profileCount; i++)
        if (_stricmp(g_profileNames[i], active) == 0) { g_activeProfile = i; break; }
    strcpy(g_iniPath, g_profilePaths[g_activeProfile]);
}

static void SwitchProfile(int idx)
{
    if (idx < 0 || idx >= g_profileCount || idx == g_activeProfile) return;
    SaveAll();
    g_activeProfile = idx;
    strcpy(g_iniPath, g_profilePaths[idx]);
    /* Persist the choice */
    WritePrivateProfileString("Meta", "ActiveProfile",
                              g_profileNames[idx], g_basePath);
    /* Tear down and rebuild everything */
    UnregisterAllHotkeys();
    FreeIcons();
    g_count = 0;
    memset(g_collapsed, 0, sizeof(g_collapsed));
    g_filterText[0] = '\0';
    if (g_hwndSearch) SetWindowText(g_hwndSearch, "");
    LoadSettings();
    LoadButtons();
    RecreateFont();
    LoadButtonIcons();
    ApplyDarkBackground();
    SetTitleBarDark(g_hwndMain, g_darkMode);
    SetWindowText(g_hwndMain, g_winTitle);
    SetWindowPos(g_hwndMain,
                 g_alwaysOnTop ? HWND_TOPMOST : HWND_NOTOPMOST,
                 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    ApplyOpacity();
    RebuildMenu();
    RegisterAllHotkeys();
    RefreshMainWindow();
}

/* Build and pop a Profiles context menu at screen point (x,y) */
static void ShowProfilesMenu(HWND hwnd, int x, int y)
{
    HMENU hM = CreatePopupMenu();
    for (int i = 0; i < g_profileCount; i++) {
        UINT flags = MF_STRING | (i == g_activeProfile ? MF_CHECKED : 0);
        AppendMenu(hM, flags, IDM_PROFILE_BASE + i, g_profileNames[i]);
    }
    AppendMenu(hM, MF_SEPARATOR, 0, NULL);
    AppendMenu(hM, MF_STRING, IDM_PROFILE_NEW, "New Profile...");
    AppendMenu(hM, MF_STRING | (g_activeProfile == 0 ? MF_GRAYED : 0),
               IDM_PROFILE_DELETE, "Delete Current Profile");
    SetForegroundWindow(hwnd);
    TrackPopupMenu(hM, TPM_LEFTBUTTON, x, y, 0, hwnd, NULL);
    DestroyMenu(hM);
}

/* ── INI helpers ─────────────────────────────────────────────────────── */
static void GetBasePath(void)
{
    GetModuleFileName(NULL, g_basePath, MAX_PATH);
    char *dot = strrchr(g_basePath, '.');
    if (dot) strcpy(dot, ".ini"); else strcat(g_basePath, ".ini");
    /* g_iniPath starts pointing at the base; ScanProfiles may redirect it */
    strcpy(g_iniPath, g_basePath);
}

static void EncodeNewlines(const char *src, char *dst, int dstSize)
{
    int j = 0;
    for (int i = 0; src[i] && j < dstSize - 3; i++) {
        if (src[i] == '\r') continue;
        if (src[i] == '\n') { dst[j++] = '\\'; dst[j++] = 'n'; }
        else                  dst[j++] = src[i];
    }
    dst[j] = '\0';
}

static void DecodeNewlines(const char *src, char *dst, int dstSize)
{
    int j = 0;
    for (int i = 0; src[i] && j < dstSize - 3; i++) {
        if (src[i] == '\\' && src[i + 1] == 'n') {
            dst[j++] = '\r'; dst[j++] = '\n'; i++;
        } else {
            dst[j++] = src[i];
        }
    }
    dst[j] = '\0';
}

static void LoadSettings(void)
{
    g_darkMode    = GetPrivateProfileInt("Settings", "DarkMode",    0,   g_iniPath);
    g_alwaysOnTop = GetPrivateProfileInt("Settings", "AlwaysOnTop", 0,   g_iniPath);
    g_minToTray   = GetPrivateProfileInt("Settings", "MinToTray",   0,   g_iniPath);
    g_fontSize    = GetPrivateProfileInt("Settings", "FontSize",    9,   g_iniPath);
    g_winWidth    = GetPrivateProfileInt("Settings", "WindowWidth", 300, g_iniPath);
    g_winX        = GetPrivateProfileInt("Settings", "WindowX",    -1,  g_iniPath);
    g_winY        = GetPrivateProfileInt("Settings", "WindowY",    -1,  g_iniPath);
    g_opacity     = GetPrivateProfileInt("Settings", "Opacity",    100, g_iniPath);
    g_compactMode = GetPrivateProfileInt("Settings", "CompactMode", 0,   g_iniPath);
    GetPrivateProfileString("Settings", "WindowTitle", "Simple Typer",
                            g_winTitle, sizeof(g_winTitle), g_iniPath);
    char encoded[512];
    GetPrivateProfileString("Settings", "Prefix", "", encoded, sizeof(encoded), g_iniPath);
    DecodeNewlines(encoded, g_prefix, sizeof(g_prefix));
    GetPrivateProfileString("Settings", "Suffix", "", encoded, sizeof(encoded), g_iniPath);
    DecodeNewlines(encoded, g_suffix, sizeof(g_suffix));

    if (g_fontSize < 6)   g_fontSize = 6;
    if (g_fontSize > 72)  g_fontSize = 72;
    if (g_winWidth < 150) g_winWidth = 150;
    if (g_winWidth > 800) g_winWidth = 800;
    if (g_opacity  < 10)  g_opacity  = 10;
    if (g_opacity  > 100) g_opacity  = 100;
}

static void LoadButtons(void)
{
    g_count = GetPrivateProfileInt("Buttons", "Count", 0, g_iniPath);
    if (g_count > MAX_BUTTONS) g_count = MAX_BUTTONS;
    for (int i = 0; i < g_count; i++) {
        char sec[32], encoded[MAX_TEXT];
        sprintf(sec, "Button%d", i + 1);
        GetPrivateProfileString(sec, "Name",     "", g_buttons[i].name,     256,      g_iniPath);
        GetPrivateProfileString(sec, "IconPath", "", g_buttons[i].iconPath, MAX_PATH, g_iniPath);
        GetPrivateProfileString(sec, "Text",     "", encoded,               MAX_TEXT, g_iniPath);
        DecodeNewlines(encoded, g_buttons[i].text, MAX_TEXT);
        g_buttons[i].btnColor    = (COLORREF)GetPrivateProfileInt(sec, "BtnColor",    0, g_iniPath);
        g_buttons[i].hasColor    = GetPrivateProfileInt(sec, "HasColor",    0, g_iniPath);
        g_buttons[i].isSeparator = GetPrivateProfileInt(sec, "Separator",   0, g_iniPath);
        g_buttons[i].showIcon    = GetPrivateProfileInt(sec, "ShowIcon",    0, g_iniPath);
        g_buttons[i].isCategory  = GetPrivateProfileInt(sec, "IsCategory",  0, g_iniPath);
        g_buttons[i].hotkeyMod   = GetPrivateProfileInt(sec, "HotkeyMod",   0, g_iniPath);
        g_buttons[i].hotkeyVk    = GetPrivateProfileInt(sec, "HotkeyVk",    0, g_iniPath);
        g_collapsed[i] = 0;
    }
}

static void SaveAll(void)
{
    if (g_hwndMain) {
        RECT rc; GetWindowRect(g_hwndMain, &rc);
        g_winX = rc.left; g_winY = rc.top;
    }
    FILE *f = fopen(g_iniPath, "w");
    if (!f) return;

    char enc_prefix[512], enc_suffix[512];
    EncodeNewlines(g_prefix, enc_prefix, sizeof(enc_prefix));
    EncodeNewlines(g_suffix, enc_suffix, sizeof(enc_suffix));

    fprintf(f, "[Settings]\r\n");
    fprintf(f, "DarkMode=%d\r\n",     g_darkMode);
    fprintf(f, "AlwaysOnTop=%d\r\n",  g_alwaysOnTop);
    fprintf(f, "MinToTray=%d\r\n",    g_minToTray);
    fprintf(f, "FontSize=%d\r\n",     g_fontSize);
    fprintf(f, "WindowWidth=%d\r\n",  g_winWidth);
    fprintf(f, "WindowX=%d\r\n",      g_winX);
    fprintf(f, "WindowY=%d\r\n",      g_winY);
    fprintf(f, "Opacity=%d\r\n",      g_opacity);
    fprintf(f, "CompactMode=%d\r\n",  g_compactMode);
    fprintf(f, "WindowTitle=%s\r\n",  g_winTitle);
    fprintf(f, "Prefix=%s\r\n",       enc_prefix);
    fprintf(f, "Suffix=%s\r\n",       enc_suffix);
    fprintf(f, "\r\n[Buttons]\r\nCount=%d\r\n", g_count);
    for (int i = 0; i < g_count; i++) {
        char encoded[MAX_TEXT];
        EncodeNewlines(g_buttons[i].text, encoded, MAX_TEXT);
        fprintf(f, "\r\n[Button%d]\r\n",   i + 1);
        fprintf(f, "Name=%s\r\n",          g_buttons[i].name);
        fprintf(f, "Text=%s\r\n",          encoded);
        fprintf(f, "IconPath=%s\r\n",      g_buttons[i].iconPath);
        fprintf(f, "BtnColor=%d\r\n",      (int)g_buttons[i].btnColor);
        fprintf(f, "HasColor=%d\r\n",      g_buttons[i].hasColor);
        fprintf(f, "Separator=%d\r\n",     g_buttons[i].isSeparator);
        fprintf(f, "ShowIcon=%d\r\n",      g_buttons[i].showIcon);
        fprintf(f, "IsCategory=%d\r\n",    g_buttons[i].isCategory);
        fprintf(f, "HotkeyMod=%d\r\n",     g_buttons[i].hotkeyMod);
        fprintf(f, "HotkeyVk=%d\r\n",      g_buttons[i].hotkeyVk);
    }
    fclose(f);
}

/* ── Icons ───────────────────────────────────────────────────────────── */
static void FreeIcons(void)
{
    for (int i = 0; i < MAX_BUTTONS; i++) {
        if (g_icons[i]) { DestroyIcon(g_icons[i]); g_icons[i] = NULL; }
    }
}

static void LoadButtonIcons(void)
{
    FreeIcons();
    for (int i = 0; i < g_count; i++) {
        if (!g_buttons[i].iconPath[0] || g_buttons[i].isSeparator || g_buttons[i].isCategory)
            continue;
        g_icons[i] = (HICON)LoadImage(NULL, g_buttons[i].iconPath,
                                      IMAGE_ICON, 16, 16, LR_LOADFROMFILE);
    }
}

static HGLOBAL g_hOldClip    = NULL;
static HWINEVENTHOOK g_hEventHook = NULL;

static void CALLBACK WinEventProc(HWINEVENTHOOK hHook, DWORD event,
    HWND hwnd, LONG idObject, LONG idChild,
    DWORD dwEventThread, DWORD dwmsEventTime)
{
    if (hwnd && hwnd != g_hwndMain)
        g_prevWindow = hwnd;
}

static void TypeText(const char *text)
{
    if (!text || !text[0]) return;
    if (!g_prevWindow || !IsWindow(g_prevWindow)) return;

    int wlen = MultiByteToWideChar(CP_ACP, 0, text, -1, NULL, 0);
    if (wlen <= 1) return;
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, wlen * sizeof(WCHAR));
    if (!hMem) return;
    WCHAR *pMem = (WCHAR *)GlobalLock(hMem);
    MultiByteToWideChar(CP_ACP, 0, text, -1, pMem, wlen);
    GlobalUnlock(hMem);

    if (!OpenClipboard(g_hwndMain)) { GlobalFree(hMem); return; }
    HGLOBAL hExist = (HGLOBAL)GetClipboardData(CF_UNICODETEXT);
    g_hOldClip = NULL;
    if (hExist) {
        SIZE_T sz = GlobalSize(hExist);
        g_hOldClip = GlobalAlloc(GMEM_MOVEABLE, sz);
        if (g_hOldClip) {
            void *pS = GlobalLock(hExist), *pD = GlobalLock(g_hOldClip);
            if (pS && pD) memcpy(pD, pS, sz);
            if (pS) GlobalUnlock(hExist);
            if (pD) GlobalUnlock(g_hOldClip);
        }
    }
    EmptyClipboard();
    SetClipboardData(CF_UNICODETEXT, hMem);
    CloseClipboard();
}

/* ── Variable expansion ──────────────────────────────────────────────── */
static void ExpandVariables(const char *in, char *out, int outSize, const char *clipText)
{
    SYSTEMTIME st; GetLocalTime(&st);
    char dateBuf[32], timeBuf[32];
    sprintf(dateBuf, "%02d/%02d/%04d", st.wMonth, st.wDay, st.wYear);
    sprintf(timeBuf, "%02d:%02d", st.wHour, st.wMinute);

    int i = 0, j = 0, inLen = (int)strlen(in);
    while (i < inLen && j < outSize - 1) {
        if (in[i] == '{') {
            if (strncmp(&in[i], "{date}", 6) == 0) {
                int l = (int)strlen(dateBuf);
                if (j + l < outSize - 1) { memcpy(&out[j], dateBuf, l); j += l; }
                i += 6;
            } else if (strncmp(&in[i], "{time}", 6) == 0) {
                int l = (int)strlen(timeBuf);
                if (j + l < outSize - 1) { memcpy(&out[j], timeBuf, l); j += l; }
                i += 6;
            } else if (strncmp(&in[i], "{clipboard}", 11) == 0) {
                int l = (int)strlen(clipText);
                if (j + l < outSize - 1) { memcpy(&out[j], clipText, l); j += l; }
                i += 11;
            } else { out[j++] = in[i++]; }
        } else { out[j++] = in[i++]; }
    }
    out[j] = '\0';
}

/* ── Action list builder ─────────────────────────────────────────────── */
/* Splits text into ACT_TEXT / ACT_KEY segments for sequential dispatch.  */
static void BuildFireActions(const char *text)
{
    g_fireCount = 0;
    g_fireIdx   = 0;

    char buf[MAX_TEXT];
    int  bi = 0, ti = 0, tlen = (int)strlen(text);

    while (ti < tlen) {
        if (text[ti] == '{') {
            int matched = 0;
            for (int k = 0; g_keyTokens[k].tok && g_fireCount < MAX_FIRE_ACTIONS - 1; k++) {
                int tl = g_keyTokens[k].tokLen;
                if (ti + tl > tlen) continue;
                char tmp[16]; int ci;
                for (ci = 0; ci < tl; ci++)
                    tmp[ci] = (char)tolower((unsigned char)text[ti + ci]);
                tmp[tl] = '\0';
                if (strcmp(tmp, g_keyTokens[k].tok) == 0) {
                    if (bi > 0 && g_fireCount < MAX_FIRE_ACTIONS) {
                        buf[bi] = '\0';
                        g_fireActions[g_fireCount].type = ACT_TEXT;
                        strncpy(g_fireActions[g_fireCount].text, buf, MAX_TEXT - 1);
                        g_fireActions[g_fireCount].text[MAX_TEXT - 1] = '\0';
                        g_fireActions[g_fireCount].vk = 0;
                        g_fireCount++; bi = 0;
                    }
                    if (g_fireCount < MAX_FIRE_ACTIONS) {
                        g_fireActions[g_fireCount].type    = ACT_KEY;
                        g_fireActions[g_fireCount].text[0] = '\0';
                        g_fireActions[g_fireCount].vk      = g_keyTokens[k].vk;
                        g_fireCount++;
                    }
                    ti += tl; matched = 1; break;
                }
            }
            if (!matched) { if (bi < MAX_TEXT - 1) buf[bi++] = text[ti]; ti++; }
        } else {
            if (bi < MAX_TEXT - 1) buf[bi++] = text[ti];
            ti++;
        }
    }
    if (bi > 0 && g_fireCount < MAX_FIRE_ACTIONS) {
        buf[bi] = '\0';
        g_fireActions[g_fireCount].type = ACT_TEXT;
        strncpy(g_fireActions[g_fireCount].text, buf, MAX_TEXT - 1);
        g_fireActions[g_fireCount].text[MAX_TEXT - 1] = '\0';
        g_fireActions[g_fireCount].vk = 0;
        g_fireCount++;
    }
}

/* Dispatch the next action. Called from FireButton and from WM_TIMER     */
/* (timers 2 and 4) when the previous action completes.                   */
static void FireNextAction(void)
{
    while (g_fireIdx < g_fireCount &&
           g_fireActions[g_fireIdx].type == ACT_TEXT &&
           !g_fireActions[g_fireIdx].text[0])
        g_fireIdx++;

    if (g_fireIdx >= g_fireCount) { g_pendingIdx = -1; return; }

    FireAction *a = &g_fireActions[g_fireIdx++];

    if (a->type == ACT_TEXT) {
        TypeText(a->text);
        SetForegroundWindow(g_prevWindow);
        SetTimer(g_hwndMain, 1, 80, NULL);
    } else {
        g_pendingVk = a->vk;
        SetForegroundWindow(g_prevWindow);
        SetTimer(g_hwndMain, 3, 50, NULL);
    }
}

/* ── Filter helper ───────────────────────────────────────────────────── */
static int ButtonMatchesFilter(int i)
{
    if (!g_filterText[0]) return 1;
    if (g_buttons[i].isSeparator || g_buttons[i].isCategory) return 0;
    char hay[256], ndl[256]; int hi = 0, ni = 0;
    for (const char *p = g_buttons[i].name; *p && hi < 255; p++)
        hay[hi++] = (char)tolower((unsigned char)*p);
    hay[hi] = '\0';
    for (const char *p = g_filterText; *p && ni < 255; p++)
        ndl[ni++] = (char)tolower((unsigned char)*p);
    ndl[ni] = '\0';
    return (strstr(hay, ndl) != NULL) ? 1 : 0;
}

/* ── Owner-draw button ───────────────────────────────────────────────── */
static void DrawButton(LPDRAWITEMSTRUCT dis, int idx)
{
    RECT rc      = dis->rcItem;
    BOOL pressed = (dis->itemState & ODS_SELECTED);

    /* ── Compact mode (non-Add button) ── */
    if (g_compactMode && idx >= 0) {
        if (g_darkMode) {
            HBRUSH hbr = CreateSolidBrush(pressed ? DK_BTN_PRE : DK_BTN);
            FillRect(dis->hDC, &rc, hbr); DeleteObject(hbr);
            HPEN hp = CreatePen(PS_SOLID, 1, DK_BORDER);
            HPEN hop = (HPEN)SelectObject(dis->hDC, hp);
            HBRUSH hn = (HBRUSH)SelectObject(dis->hDC, GetStockObject(NULL_BRUSH));
            Rectangle(dis->hDC, rc.left, rc.top, rc.right, rc.bottom);
            SelectObject(dis->hDC, hop); SelectObject(dis->hDC, hn); DeleteObject(hp);
        } else {
            FillRect(dis->hDC, &rc, (HBRUSH)(COLOR_BTNFACE + 1));
            DrawEdge(dis->hDC, &rc, pressed ? EDGE_SUNKEN : EDGE_RAISED, BF_RECT);
        }
        if (idx < g_count && g_buttons[idx].hasColor) {
            HPEN hp = CreatePen(PS_SOLID, 2, g_buttons[idx].btnColor);
            HPEN hop = (HPEN)SelectObject(dis->hDC, hp);
            HBRUSH hn = (HBRUSH)SelectObject(dis->hDC, GetStockObject(NULL_BRUSH));
            RECT inner = { rc.left+2, rc.top+2, rc.right-2, rc.bottom-2 };
            Rectangle(dis->hDC, inner.left, inner.top, inner.right, inner.bottom);
            SelectObject(dis->hDC, hop); SelectObject(dis->hDC, hn); DeleteObject(hp);
        }
        int sz = 16, ix = rc.left + (rc.right-rc.left-sz)/2, iy = rc.top + (rc.bottom-rc.top-sz)/2;
        if (idx < g_count && g_icons[idx]) {
            DrawIconEx(dis->hDC, ix, iy, g_icons[idx], sz, sz, 0, NULL, DI_NORMAL);
        } else if (idx < g_count && g_buttons[idx].name[0]) {
            char letter[2] = { (char)toupper((unsigned char)g_buttons[idx].name[0]), '\0' };
            SetBkMode(dis->hDC, TRANSPARENT);
            SetTextColor(dis->hDC, g_darkMode ? DK_TEXT : GetSysColor(COLOR_BTNTEXT));
            HFONT hof = g_hFont ? (HFONT)SelectObject(dis->hDC, g_hFont) : NULL;
            DrawText(dis->hDC, letter, 1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            if (hof) SelectObject(dis->hDC, hof);
        }
        if (dis->itemState & ODS_FOCUS) DrawFocusRect(dis->hDC, &rc);
        return;
    }

    /* ── Category header ── */
    if (idx >= 0 && idx < g_count && g_buttons[idx].isCategory) {
        COLORREF bg   = g_darkMode ? DK_CAT_BG   : LT_CAT_BG;
        COLORREF text = g_darkMode ? DK_CAT_TEXT  : LT_CAT_TEXT;
        HBRUSH hbr = CreateSolidBrush(pressed ? (g_darkMode ? DK_BTN_PRE : RGB(180,205,235)) : bg);
        FillRect(dis->hDC, &rc, hbr); DeleteObject(hbr);

        /* arrow: ▶ collapsed, ▼ expanded */
        const char *arrow = g_collapsed[idx] ? ">" : "v";
        SetBkMode(dis->hDC, TRANSPARENT);
        SetTextColor(dis->hDC, text);
        HFONT hof = g_hFontBold ? (HFONT)SelectObject(dis->hDC, g_hFontBold) : NULL;
        RECT arrowRc = { rc.left + 6, rc.top, rc.left + 20, rc.bottom };
        DrawText(dis->hDC, arrow, -1, &arrowRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        RECT textRc  = { rc.left + 22, rc.top, rc.right - 4, rc.bottom };
        DrawText(dis->hDC, g_buttons[idx].name, -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        if (hof) SelectObject(dis->hDC, hof);
        if (dis->itemState & ODS_FOCUS) DrawFocusRect(dis->hDC, &rc);
        return;
    }

    /* ── Separator ── */
    if (idx >= 0 && idx < g_count && g_buttons[idx].isSeparator) {
        HBRUSH hbr = CreateSolidBrush(g_darkMode ? DK_BG : GetSysColor(COLOR_BTNFACE));
        FillRect(dis->hDC, &rc, hbr); DeleteObject(hbr);
        int midY = (rc.top + rc.bottom) / 2;
        HPEN hp = CreatePen(PS_SOLID, 1, g_darkMode ? DK_SEP : LT_SEP);
        HPEN hop = (HPEN)SelectObject(dis->hDC, hp);
        MoveToEx(dis->hDC, rc.left + 6, midY, NULL);
        LineTo(dis->hDC, rc.right - 6, midY);
        SelectObject(dis->hDC, hop); DeleteObject(hp);
        return;
    }

    /* ── Normal button ── */
    if (g_darkMode) {
        HBRUSH hbr = CreateSolidBrush(pressed ? DK_BTN_PRE : DK_BTN);
        FillRect(dis->hDC, &rc, hbr); DeleteObject(hbr);
        HPEN hp = CreatePen(PS_SOLID, 1, DK_BORDER);
        HPEN hop = (HPEN)SelectObject(dis->hDC, hp);
        HBRUSH hn = (HBRUSH)SelectObject(dis->hDC, GetStockObject(NULL_BRUSH));
        Rectangle(dis->hDC, rc.left, rc.top, rc.right, rc.bottom);
        SelectObject(dis->hDC, hop); SelectObject(dis->hDC, hn); DeleteObject(hp);
    } else {
        FillRect(dis->hDC, &rc, (HBRUSH)(COLOR_BTNFACE + 1));
        DrawEdge(dis->hDC, &rc, pressed ? EDGE_SUNKEN : EDGE_RAISED, BF_RECT);
    }

    if (idx >= 0 && idx < g_count && g_buttons[idx].hasColor) {
        HPEN hp = CreatePen(PS_SOLID, 2, g_buttons[idx].btnColor);
        HPEN hop = (HPEN)SelectObject(dis->hDC, hp);
        HBRUSH hn = (HBRUSH)SelectObject(dis->hDC, GetStockObject(NULL_BRUSH));
        RECT inner = { rc.left+3, rc.top+3, rc.right-3, rc.bottom-3 };
        Rectangle(dis->hDC, inner.left, inner.top, inner.right, inner.bottom);
        SelectObject(dis->hDC, hop); SelectObject(dis->hDC, hn); DeleteObject(hp);
    }

    int textLeft = rc.left;
    if (idx >= 0 && idx < g_count && g_icons[idx] && g_buttons[idx].showIcon) {
        int sz = 16, ix = rc.left + 6, iy = rc.top + (rc.bottom - rc.top - sz) / 2;
        DrawIconEx(dis->hDC, ix, iy, g_icons[idx], sz, sz, 0, NULL, DI_NORMAL);
        textLeft = ix + sz + 4;
    }

    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, g_darkMode ? DK_TEXT : GetSysColor(COLOR_BTNTEXT));
    HFONT hof = g_hFont ? (HFONT)SelectObject(dis->hDC, g_hFont) : NULL;
    RECT textRc = { textLeft, rc.top, rc.right, rc.bottom };
    if (pressed) OffsetRect(&textRc, 1, 1);
    const char *label = (idx == -1) ? "+ Add Button" :
                        (idx >= 0 && idx < g_count) ? g_buttons[idx].name : "";
    DrawText(dis->hDC, label, -1, &textRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    if (hof) SelectObject(dis->hDC, hof);
    if (dis->itemState & ODS_FOCUS) DrawFocusRect(dis->hDC, &rc);
}

/* ── Dark background ─────────────────────────────────────────────────── */
static void ApplyDarkBackground(void)
{
    if (g_hbrDkBg) { DeleteObject(g_hbrDkBg); g_hbrDkBg = NULL; }
    if (g_darkMode) g_hbrDkBg = CreateSolidBrush(DK_BG);
    if (g_hbrSearchDk) { DeleteObject(g_hbrSearchDk); g_hbrSearchDk = NULL; }
    if (g_darkMode) g_hbrSearchDk = CreateSolidBrush(DK_SEARCH);
    if (g_hwndSearch) InvalidateRect(g_hwndSearch, NULL, TRUE);
}

/* ── Menu ────────────────────────────────────────────────────────────── */
static void RebuildMenu(void)
{
    HMENU hOld = GetMenu(g_hwndMain);
    if (hOld) { SetMenu(g_hwndMain, NULL); DestroyMenu(hOld); }
    HMENU hBar = CreateMenu();
    int n = sizeof(g_menuIDs) / sizeof(g_menuIDs[0]);
    for (int i = 0; i < n; i++) {
        if (g_darkMode) AppendMenu(hBar, MF_OWNERDRAW, g_menuIDs[i], (LPCSTR)g_menuLabels[i]);
        else            AppendMenu(hBar, MF_STRING,    g_menuIDs[i], g_menuLabels[i]);
    }
    SetMenu(g_hwndMain, hBar);
    DrawMenuBar(g_hwndMain);
}

/* ── Layout ──────────────────────────────────────────────────────────── */
static void RefreshMainWindow(void)
{
    int btnAreaW = g_winWidth - 20;

    for (int i = 0; i < MAX_BUTTONS; i++) {
        if (g_hwndBtns[i]) { DestroyWindow(g_hwndBtns[i]); g_hwndBtns[i] = NULL; }
    }
    if (g_hwndTooltip) { DestroyWindow(g_hwndTooltip); g_hwndTooltip = NULL; }

    LoadButtonIcons();

    /* Search bar */
    if (!g_hwndSearch) {
        g_hwndSearch = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", "",
            WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL,
            10, 10, btnAreaW, 22, g_hwndMain, (HMENU)IDC_SEARCH_EDIT, g_hInst, NULL);
        SendMessage(g_hwndSearch, EM_SETCUEBANNER, FALSE, (LPARAM)L"Search\u2026");
        if (g_hFont) SendMessage(g_hwndSearch, WM_SETFONT, (WPARAM)g_hFont, FALSE);
    } else {
        SetWindowPos(g_hwndSearch, HWND_TOP, 10, 10, btnAreaW, 22, SWP_SHOWWINDOW);
        if (g_hFont) SendMessage(g_hwndSearch, WM_SETFONT, (WPARAM)g_hFont, FALSE);
    }

    int y = 10 + 22 + 6;

    if (g_compactMode) {
        int sz = COMPACT_BTN_SZ, gap = COMPACT_BTN_GAP;
        int cols = (btnAreaW + gap) / (sz + gap);
        if (cols < 1) cols = 1;
        int col = 0;
        for (int i = 0; i < g_count; i++) {
            if (g_buttons[i].isSeparator || g_buttons[i].isCategory) continue;
            int bx = 10 + col * (sz + gap);
            g_hwndBtns[i] = CreateWindow("BUTTON", g_buttons[i].name,
                WS_VISIBLE | WS_CHILD | BS_OWNERDRAW, bx, y, sz, sz,
                g_hwndMain, (HMENU)(UINT_PTR)(ID_BUTTON_BASE + i), g_hInst, NULL);
            if (g_hFont) SendMessage(g_hwndBtns[i], WM_SETFONT, (WPARAM)g_hFont, FALSE);
            col++;
            if (col >= cols) { col = 0; y += sz + gap; }
        }
        if (col > 0) y += sz + gap;

    } else if (g_filterText[0]) {
        /* flat filtered list — no category structure */
        for (int i = 0; i < g_count; i++) {
            if (!ButtonMatchesFilter(i)) continue;
            g_hwndBtns[i] = CreateWindow("BUTTON", g_buttons[i].name,
                WS_VISIBLE | WS_CHILD | BS_OWNERDRAW, 10, y, btnAreaW, 26,
                g_hwndMain, (HMENU)(UINT_PTR)(ID_BUTTON_BASE + i), g_hInst, NULL);
            if (g_hFont) SendMessage(g_hwndBtns[i], WM_SETFONT, (WPARAM)g_hFont, FALSE);
            y += 26 + 5;
        }

    } else {
        /* full list with categories */
        int catCollapsed = 0;
        for (int i = 0; i < g_count; i++) {
            if (g_buttons[i].isCategory) {
                catCollapsed = g_collapsed[i];
                int h = 24;
                g_hwndBtns[i] = CreateWindow("BUTTON", g_buttons[i].name,
                    WS_VISIBLE | WS_CHILD | BS_OWNERDRAW, 10, y, btnAreaW, h,
                    g_hwndMain, (HMENU)(UINT_PTR)(ID_BUTTON_BASE + i), g_hInst, NULL);
                if (g_hFontBold) SendMessage(g_hwndBtns[i], WM_SETFONT, (WPARAM)g_hFontBold, FALSE);
                y += h + 3;

            } else if (catCollapsed) {
                /* skip — category is collapsed */
                continue;

            } else if (g_buttons[i].isSeparator) {
                g_hwndBtns[i] = CreateWindow("BUTTON", "", WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
                    10, y, btnAreaW, 14, g_hwndMain,
                    (HMENU)(UINT_PTR)(ID_BUTTON_BASE + i), g_hInst, NULL);
                y += 14 + 5;

            } else {
                g_hwndBtns[i] = CreateWindow("BUTTON", g_buttons[i].name,
                    WS_VISIBLE | WS_CHILD | BS_OWNERDRAW, 10, y, btnAreaW, 26,
                    g_hwndMain, (HMENU)(UINT_PTR)(ID_BUTTON_BASE + i), g_hInst, NULL);
                if (g_hFont) SendMessage(g_hwndBtns[i], WM_SETFONT, (WPARAM)g_hFont, FALSE);
                y += 26 + 5;
            }
        }
    }

    /* "+ Add Button" */
    HWND hAdd = GetDlgItem(g_hwndMain, ID_ADD_BTN);
    SetWindowLongPtr(hAdd, GWL_STYLE, WS_VISIBLE | WS_CHILD | BS_OWNERDRAW);
    SetWindowPos(hAdd, HWND_TOP, 10, y, btnAreaW, 26, SWP_SHOWWINDOW | SWP_FRAMECHANGED);
    if (g_hFont) SendMessage(hAdd, WM_SETFONT, (WPARAM)g_hFont, FALSE);

    int clientH = y + 31 + 10;
    if (clientH < 80) clientH = 80;
    RECT rc = { 0, 0, g_winWidth, clientH };
    AdjustWindowRect(&rc, WS_OVERLAPPEDWINDOW, TRUE);
    SetWindowPos(g_hwndMain, NULL, 0, 0,
                 rc.right - rc.left, rc.bottom - rc.top, SWP_NOMOVE | SWP_NOZORDER);
    InvalidateRect(g_hwndMain, NULL, TRUE);

    /* Tooltips */
    g_hwndTooltip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, NULL,
        WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        g_hwndMain, NULL, g_hInst, NULL);
    if (g_hwndTooltip) {
        SendMessage(g_hwndTooltip, TTM_SETMAXTIPWIDTH, 0, 320);
        SendMessage(g_hwndTooltip, TTM_SETDELAYTIME, TTDT_INITIAL, 600);
        for (int i = 0; i < g_count; i++) {
            if (!g_hwndBtns[i]) continue;

            const char *src = NULL;
            char tipbuf[160] = "";
            if (g_buttons[i].isCategory) {
                src = g_buttons[i].name;
                snprintf(tipbuf, sizeof(tipbuf), "Category: %s", src);
            } else if (g_buttons[i].isSeparator) {
                continue;
            } else if (g_compactMode) {
                /* in compact mode show name */
                src = g_buttons[i].name;
                snprintf(tipbuf, sizeof(tipbuf), "%s", src);
            } else {
                /* normal mode: text preview + hotkey if set */
                src = g_buttons[i].text;
                int k = 0;
                for (int c = 0; src[c] && k < 80; c++) {
                    if (src[c] == '\r') continue;
                    tipbuf[k++] = (src[c] == '\n') ? ' ' : src[c];
                }
                tipbuf[k] = '\0';
                if ((int)strlen(src) > 80) strcat(tipbuf, "...");
                if (g_buttons[i].hotkeyVk) {
                    char hkstr[64]; HotkeyToString(g_buttons[i].hotkeyMod, g_buttons[i].hotkeyVk, hkstr, sizeof(hkstr));
                    int tl = (int)strlen(tipbuf);
                    snprintf(tipbuf + tl, sizeof(tipbuf) - tl, "\n[%s]", hkstr);
                }
            }
            if (!tipbuf[0]) continue;

            /* copy into stable per-button buffer */
            strncpy(g_tooltipText[i], tipbuf, 127); g_tooltipText[i][127] = '\0';

            TOOLINFO ti; ZeroMemory(&ti, sizeof(ti));
            ti.cbSize   = sizeof(TOOLINFO);
            ti.uFlags   = TTF_IDISHWND | TTF_SUBCLASS;
            ti.hwnd     = g_hwndMain;
            ti.uId      = (UINT_PTR)g_hwndBtns[i];
            ti.lpszText = g_tooltipText[i];
            SendMessage(g_hwndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
        }
    }
}

/* ── Right-click context menu ────────────────────────────────────────── */
static void ShowButtonContextMenu(HWND hwnd, int idx, POINT pt)
{
    HMENU hMenu = CreatePopupMenu();
    AppendMenu(hMenu, MF_STRING | (idx == 0 ? MF_GRAYED : 0),         IDM_MOVE_UP,       "Move Up");
    AppendMenu(hMenu, MF_STRING | (idx == g_count-1 ? MF_GRAYED : 0), IDM_MOVE_DOWN,     "Move Down");
    AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
    if (!g_buttons[idx].isSeparator)
        AppendMenu(hMenu, MF_STRING, IDM_EDIT_BTN, "Edit...");
    AppendMenu(hMenu, MF_STRING | (g_count >= MAX_BUTTONS ? MF_GRAYED : 0),
               IDM_DUPLICATE_BTN, "Duplicate");
    AppendMenu(hMenu, MF_STRING, IDM_DELETE_BTN, "Delete");
    SetForegroundWindow(hwnd);
    TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
    DestroyMenu(hMenu);
}

/* ── Shared dark coloring helper for dialogs ─────────────────────────── */
static LRESULT HandleDlgDarkColor(HWND hwnd, UINT msg, WPARAM wParam, HBRUSH *phBr)
{
    if (!g_darkMode) return -1;
    if (msg == WM_ERASEBKGND) {
        HDC hdc = (HDC)wParam; RECT rc;
        GetClientRect(hwnd, &rc);
        if (!*phBr) *phBr = CreateSolidBrush(DK_BG);
        FillRect(hdc, &rc, *phBr); return 1;
    }
    if (msg == WM_CTLCOLORSTATIC || msg == WM_CTLCOLORBTN || msg == WM_CTLCOLOREDIT
        || msg == WM_CTLCOLORLISTBOX) {
        HDC hdc = (HDC)wParam;
        SetTextColor(hdc, DK_TEXT); SetBkColor(hdc, DK_BG);
        if (!*phBr) *phBr = CreateSolidBrush(DK_BG);
        return (LRESULT)*phBr;
    }
    return -1;
}

/* ── Settings dialog ─────────────────────────────────────────────────── */
static HBRUSH g_hbrSettingsBg = NULL;

static LRESULT CALLBACK SettingsDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    LRESULT dark = HandleDlgDarkColor(hwnd, msg, wParam, &g_hbrSettingsBg);
    if (dark != -1) return dark;

    switch (msg) {
    case WM_DESTROY:
        if (g_hbrSettingsBg) { DeleteObject(g_hbrSettingsBg); g_hbrSettingsBg = NULL; }
        g_hwndDlg = NULL; return 0;

    case WM_CREATE:
        CreateWindow("STATIC", "Appearance", WS_VISIBLE|WS_CHILD, 10,10,200,16, hwnd, NULL, g_hInst, NULL);
        CreateWindow("BUTTON", "Dark mode", WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 10,30,200,20, hwnd,(HMENU)IDC_DARK_CHECK,g_hInst,NULL);
        SendDlgItemMessage(hwnd, IDC_DARK_CHECK, BM_SETCHECK, g_darkMode ? BST_CHECKED : BST_UNCHECKED, 0);

        CreateWindow("STATIC", "Font size:", WS_VISIBLE|WS_CHILD, 10,60,80,18, hwnd,NULL,g_hInst,NULL);
        { char b[8]; sprintf(b,"%d",g_fontSize);
          CreateWindow("EDIT",b,WS_VISIBLE|WS_CHILD|WS_BORDER|ES_NUMBER, 95,58,40,20, hwnd,(HMENU)IDC_FONT_EDIT,g_hInst,NULL); }
        CreateWindow("STATIC","pt",WS_VISIBLE|WS_CHILD,140,60,20,18,hwnd,NULL,g_hInst,NULL);

        CreateWindow("STATIC","Layout",WS_VISIBLE|WS_CHILD,10,88,200,16,hwnd,NULL,g_hInst,NULL);
        CreateWindow("STATIC","Window width:",WS_VISIBLE|WS_CHILD,10,108,100,18,hwnd,NULL,g_hInst,NULL);
        { char b[8]; sprintf(b,"%d",g_winWidth);
          CreateWindow("EDIT",b,WS_VISIBLE|WS_CHILD|WS_BORDER|ES_NUMBER,115,106,55,20,hwnd,(HMENU)IDC_WIDTH_EDIT,g_hInst,NULL); }
        CreateWindow("STATIC","px",WS_VISIBLE|WS_CHILD,175,108,20,18,hwnd,NULL,g_hInst,NULL);

        CreateWindow("STATIC","Behaviour",WS_VISIBLE|WS_CHILD,10,136,200,16,hwnd,NULL,g_hInst,NULL);
        CreateWindow("BUTTON","Always on top",WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX,10,156,200,20,hwnd,(HMENU)IDC_TOPMOST_CHECK,g_hInst,NULL);
        SendDlgItemMessage(hwnd,IDC_TOPMOST_CHECK,BM_SETCHECK,g_alwaysOnTop?BST_CHECKED:BST_UNCHECKED,0);
        CreateWindow("BUTTON","Minimize to system tray",WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX,10,180,200,20,hwnd,(HMENU)IDC_TRAY_CHECK,g_hInst,NULL);
        SendDlgItemMessage(hwnd,IDC_TRAY_CHECK,BM_SETCHECK,g_minToTray?BST_CHECKED:BST_UNCHECKED,0);
        CreateWindow("BUTTON","Compact mode (icon grid, no labels)",WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX,10,204,265,20,hwnd,(HMENU)IDC_COMPACT_CHECK,g_hInst,NULL);
        SendDlgItemMessage(hwnd,IDC_COMPACT_CHECK,BM_SETCHECK,g_compactMode?BST_CHECKED:BST_UNCHECKED,0);

        CreateWindow("STATIC","Window",WS_VISIBLE|WS_CHILD,10,234,200,16,hwnd,NULL,g_hInst,NULL);
        CreateWindow("STATIC","Title:",WS_VISIBLE|WS_CHILD,10,254,40,18,hwnd,NULL,g_hInst,NULL);
        CreateWindow("EDIT",g_winTitle,WS_VISIBLE|WS_CHILD|WS_BORDER|ES_AUTOHSCROLL,55,252,220,20,hwnd,(HMENU)IDC_TITLE_EDIT,g_hInst,NULL);
        CreateWindow("STATIC","Opacity:",WS_VISIBLE|WS_CHILD,10,280,55,18,hwnd,NULL,g_hInst,NULL);
        { char b[8]; sprintf(b,"%d",g_opacity);
          CreateWindow("EDIT",b,WS_VISIBLE|WS_CHILD|WS_BORDER|ES_NUMBER,70,278,40,20,hwnd,(HMENU)IDC_OPACITY_EDIT,g_hInst,NULL); }
        CreateWindow("STATIC","% (10-100)",WS_VISIBLE|WS_CHILD,115,280,75,18,hwnd,NULL,g_hInst,NULL);

        CreateWindow("STATIC","Text output (optional)",WS_VISIBLE|WS_CHILD,10,310,265,16,hwnd,NULL,g_hInst,NULL);
        CreateWindow("STATIC","Prefix:",WS_VISIBLE|WS_CHILD,10,332,50,18,hwnd,NULL,g_hInst,NULL);
        CreateWindow("EDIT",g_prefix,WS_VISIBLE|WS_CHILD|WS_BORDER|ES_AUTOHSCROLL,65,330,210,20,hwnd,(HMENU)IDC_PREFIX_EDIT,g_hInst,NULL);
        CreateWindow("STATIC","Suffix:",WS_VISIBLE|WS_CHILD,10,358,50,18,hwnd,NULL,g_hInst,NULL);
        CreateWindow("EDIT",g_suffix,WS_VISIBLE|WS_CHILD|WS_BORDER|ES_AUTOHSCROLL,65,356,210,20,hwnd,(HMENU)IDC_SUFFIX_EDIT,g_hInst,NULL);

        CreateWindow("BUTTON","Save",WS_VISIBLE|WS_CHILD|BS_DEFPUSHBUTTON,108,390,80,26,hwnd,(HMENU)IDC_OK,g_hInst,NULL);
        CreateWindow("BUTTON","Cancel",WS_VISIBLE|WS_CHILD|BS_PUSHBUTTON,198,390,80,26,hwnd,(HMENU)IDC_CANCEL,g_hInst,NULL);
        return 0;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_OK: {
            int nd = (SendDlgItemMessage(hwnd,IDC_DARK_CHECK,   BM_GETCHECK,0,0)==BST_CHECKED)?1:0;
            int nt = (SendDlgItemMessage(hwnd,IDC_TOPMOST_CHECK,BM_GETCHECK,0,0)==BST_CHECKED)?1:0;
            int ny = (SendDlgItemMessage(hwnd,IDC_TRAY_CHECK,   BM_GETCHECK,0,0)==BST_CHECKED)?1:0;
            int nc = (SendDlgItemMessage(hwnd,IDC_COMPACT_CHECK,BM_GETCHECK,0,0)==BST_CHECKED)?1:0;
            char buf[256];
            GetDlgItemText(hwnd,IDC_FONT_EDIT,buf,16);
            int nf=atoi(buf); if(nf<6)nf=6; if(nf>72)nf=72;
            GetDlgItemText(hwnd,IDC_WIDTH_EDIT,buf,16);
            int nw=atoi(buf); if(nw<150)nw=150; if(nw>800)nw=800;
            GetDlgItemText(hwnd,IDC_OPACITY_EDIT,buf,16);
            int no=atoi(buf); if(no<10)no=10; if(no>100)no=100;
            char ntitle[256]; GetDlgItemText(hwnd,IDC_TITLE_EDIT,ntitle,256);
            if(!ntitle[0]) strcpy(ntitle,"Simple Typer");
            char npfx[512], nsfx[512];
            GetDlgItemText(hwnd,IDC_PREFIX_EDIT,npfx,sizeof(npfx));
            GetDlgItemText(hwnd,IDC_SUFFIX_EDIT,nsfx,sizeof(nsfx));

            int refresh=0;
            if(nd!=g_darkMode){ g_darkMode=nd; ApplyDarkBackground(); SetTitleBarDark(g_hwndMain,g_darkMode); RebuildMenu(); refresh=1; }
            if(nt!=g_alwaysOnTop){ g_alwaysOnTop=nt; SetWindowPos(g_hwndMain,nt?HWND_TOPMOST:HWND_NOTOPMOST,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE); }
            g_minToTray=ny;
            if(nc!=g_compactMode){ g_compactMode=nc; refresh=1; }
            if(nf!=g_fontSize){ g_fontSize=nf; RecreateFont(); refresh=1; }
            if(nw!=g_winWidth){ g_winWidth=nw; refresh=1; }
            if(no!=g_opacity){ g_opacity=no; ApplyOpacity(); }
            if(strcmp(ntitle,g_winTitle)!=0){ strcpy(g_winTitle,ntitle); SetWindowText(g_hwndMain,g_winTitle); }
            strcpy(g_prefix,npfx); strcpy(g_suffix,nsfx);
            SaveAll();
            if(refresh) RefreshMainWindow();
            DestroyWindow(hwnd); g_hwndDlg=NULL; break;
        }
        case IDC_CANCEL: DestroyWindow(hwnd); g_hwndDlg=NULL; break;
        }
        return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/* ── Add / Edit dialog ───────────────────────────────────────────────── */
static HBRUSH g_hbrAddBg = NULL;

/* Helper: enable/disable fields based on type checkboxes */
static void UpdateAddDlgFields(HWND hwnd)
{
    BOOL isSep = (SendDlgItemMessage(hwnd, IDC_SEP_CHECK, BM_GETCHECK, 0, 0) == BST_CHECKED);
    BOOL isCat = (SendDlgItemMessage(hwnd, IDC_CAT_CHECK, BM_GETCHECK, 0, 0) == BST_CHECKED);

    /* name: enabled for normal and category; disabled for separator */
    EnableWindow(GetDlgItem(hwnd, IDC_NAME_EDIT),       !isSep);
    /* text/icon/color/hotkey: only for normal buttons */
    BOOL normal = (!isSep && !isCat);
    EnableWindow(GetDlgItem(hwnd, IDC_TEXT_EDIT),       normal);
    EnableWindow(GetDlgItem(hwnd, IDC_ICON_CHECK),      normal);
    EnableWindow(GetDlgItem(hwnd, IDC_ICON_PATH_EDIT),  normal);
    EnableWindow(GetDlgItem(hwnd, IDC_ICON_BROWSE),     normal);
    EnableWindow(GetDlgItem(hwnd, IDC_BTN_COLOR_CHECK), normal);
    EnableWindow(GetDlgItem(hwnd, IDC_BTN_COLOR_BTN),   normal);
    EnableWindow(GetDlgItem(hwnd, IDC_HK_CTRL),         normal);
    EnableWindow(GetDlgItem(hwnd, IDC_HK_SHIFT),        normal);
    EnableWindow(GetDlgItem(hwnd, IDC_HK_ALT),          normal);
    EnableWindow(GetDlgItem(hwnd, IDC_HK_KEY),          normal);
    EnableWindow(GetDlgItem(hwnd, IDC_HK_CLEAR),        normal);
    /* mutually exclusive: checking one unchecks the other */
    if (isSep) SendDlgItemMessage(hwnd, IDC_CAT_CHECK, BM_SETCHECK, BST_UNCHECKED, 0);
    if (isCat) SendDlgItemMessage(hwnd, IDC_SEP_CHECK, BM_SETCHECK, BST_UNCHECKED, 0);
}

static LRESULT CALLBACK AddDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    LRESULT dark = HandleDlgDarkColor(hwnd, msg, wParam, &g_hbrAddBg);
    if (dark != -1) return dark;

    switch (msg) {
    case WM_DRAWITEM: {
        LPDRAWITEMSTRUCT dis = (LPDRAWITEMSTRUCT)lParam;
        if ((int)dis->CtlID == IDC_BTN_COLOR_BTN) {
            HBRUSH hbr = CreateSolidBrush(g_settingBtnColor);
            FillRect(dis->hDC, &dis->rcItem, hbr); DeleteObject(hbr);
            FrameRect(dis->hDC, &dis->rcItem, (HBRUSH)GetStockObject(BLACK_BRUSH));
            return TRUE;
        }
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    case WM_CREATE: {
        int edit = (g_editIndex >= 0);
        ButtonConfig *bc = edit ? &g_buttons[g_editIndex] : NULL;
        g_settingBtnColor = (edit && bc->hasColor) ? bc->btnColor : RGB(0, 120, 215);

        /* ── Name ── */
        CreateWindow("STATIC","Button display name:",WS_VISIBLE|WS_CHILD,10,10,200,18,hwnd,NULL,g_hInst,NULL);
        CreateWindow("EDIT",edit?bc->name:"",WS_VISIBLE|WS_CHILD|WS_BORDER|WS_TABSTOP|ES_AUTOHSCROLL,
                     10,30,390,22,hwnd,(HMENU)IDC_NAME_EDIT,g_hInst,NULL);
        /* ── Text ── */
        CreateWindow("STATIC","Text to type:  (use {date} {time} {clipboard} {?})",WS_VISIBLE|WS_CHILD,
                     10,62,390,16,hwnd,NULL,g_hInst,NULL);
        CreateWindow("STATIC","Keys: {tab} {enter} {esc} {backspace} {del}",WS_VISIBLE|WS_CHILD,
                     10,78,390,16,hwnd,NULL,g_hInst,NULL);
        CreateWindow("STATIC","{up} {down} {left} {right} {home} {end} {pgup} {pgdn}",WS_VISIBLE|WS_CHILD,
                     10,94,390,16,hwnd,NULL,g_hInst,NULL);
        CreateWindow("EDIT",edit?bc->text:"",WS_VISIBLE|WS_CHILD|WS_BORDER|WS_TABSTOP|WS_VSCROLL|
                     ES_MULTILINE|ES_AUTOVSCROLL|ES_WANTRETURN,
                     10,114,390,72,hwnd,(HMENU)IDC_TEXT_EDIT,g_hInst,NULL);
        /* ── Icon ── */
        CreateWindow("BUTTON","Show icon",WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX|WS_TABSTOP,
                     10,196,120,20,hwnd,(HMENU)IDC_ICON_CHECK,g_hInst,NULL);
        if(edit&&bc->showIcon) SendDlgItemMessage(hwnd,IDC_ICON_CHECK,BM_SETCHECK,BST_CHECKED,0);
        CreateWindow("STATIC","Icon file (.ico):",WS_VISIBLE|WS_CHILD,10,222,120,18,hwnd,NULL,g_hInst,NULL);
        CreateWindow("EDIT",edit?bc->iconPath:"",WS_VISIBLE|WS_CHILD|WS_BORDER|WS_TABSTOP|ES_AUTOHSCROLL,
                     10,242,290,22,hwnd,(HMENU)IDC_ICON_PATH_EDIT,g_hInst,NULL);
        CreateWindow("BUTTON","Browse...",WS_VISIBLE|WS_CHILD|BS_PUSHBUTTON|WS_TABSTOP,
                     310,242,90,22,hwnd,(HMENU)IDC_ICON_BROWSE,g_hInst,NULL);
        /* ── Color ── */
        CreateWindow("BUTTON","Custom border color",WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX|WS_TABSTOP,
                     10,274,160,20,hwnd,(HMENU)IDC_BTN_COLOR_CHECK,g_hInst,NULL);
        if(edit&&bc->hasColor) SendDlgItemMessage(hwnd,IDC_BTN_COLOR_CHECK,BM_SETCHECK,BST_CHECKED,0);
        CreateWindow("BUTTON","",WS_VISIBLE|WS_CHILD|BS_OWNERDRAW|WS_TABSTOP,
                     178,272,36,22,hwnd,(HMENU)IDC_BTN_COLOR_BTN,g_hInst,NULL);
        /* ── Hotkey ── */
        CreateWindow("STATIC","Hotkey:",WS_VISIBLE|WS_CHILD,10,306,50,18,hwnd,NULL,g_hInst,NULL);
        CreateWindow("BUTTON","Ctrl",WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX|WS_TABSTOP,
                     65,304,50,20,hwnd,(HMENU)IDC_HK_CTRL,g_hInst,NULL);
        CreateWindow("BUTTON","Shift",WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX|WS_TABSTOP,
                     120,304,55,20,hwnd,(HMENU)IDC_HK_SHIFT,g_hInst,NULL);
        CreateWindow("BUTTON","Alt",WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX|WS_TABSTOP,
                     180,304,45,20,hwnd,(HMENU)IDC_HK_ALT,g_hInst,NULL);
        /* Key combobox */
        HWND hCb = CreateWindow("COMBOBOX","",WS_VISIBLE|WS_CHILD|WS_TABSTOP|CBS_DROPDOWNLIST,
                     230,302,100,200,hwnd,(HMENU)IDC_HK_KEY,g_hInst,NULL);
        for(int k=0; g_hkNames[k]; k++) SendMessage(hCb, CB_ADDSTRING, 0, (LPARAM)g_hkNames[k]);
        CreateWindow("BUTTON","Clear",WS_VISIBLE|WS_CHILD|BS_PUSHBUTTON|WS_TABSTOP,
                     338,302,60,22,hwnd,(HMENU)IDC_HK_CLEAR,g_hInst,NULL);

        /* populate hotkey controls if editing */
        if(edit) {
            if(bc->hotkeyMod & MOD_CONTROL) SendDlgItemMessage(hwnd,IDC_HK_CTRL, BM_SETCHECK,BST_CHECKED,0);
            if(bc->hotkeyMod & MOD_SHIFT)   SendDlgItemMessage(hwnd,IDC_HK_SHIFT,BM_SETCHECK,BST_CHECKED,0);
            if(bc->hotkeyMod & MOD_ALT)     SendDlgItemMessage(hwnd,IDC_HK_ALT,  BM_SETCHECK,BST_CHECKED,0);
            int sel = 0; /* default "(none)" */
            for(int k=0; k<HK_KEY_COUNT; k++) if(g_hkVKs[k]==bc->hotkeyVk){ sel=k; break; }
            SendMessage(hCb, CB_SETCURSEL, sel, 0);
        } else {
            SendMessage(hCb, CB_SETCURSEL, 0, 0);
        }

        /* ── Type checkboxes ── */
        CreateWindow("BUTTON","Separator (divider line)",WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX|WS_TABSTOP,
                     10,336,200,20,hwnd,(HMENU)IDC_SEP_CHECK,g_hInst,NULL);
        CreateWindow("BUTTON","Category header (collapsible group)",WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX|WS_TABSTOP,
                     10,360,270,20,hwnd,(HMENU)IDC_CAT_CHECK,g_hInst,NULL);

        if(edit&&bc->isSeparator) SendDlgItemMessage(hwnd,IDC_SEP_CHECK,BM_SETCHECK,BST_CHECKED,0);
        if(edit&&bc->isCategory)  SendDlgItemMessage(hwnd,IDC_CAT_CHECK,BM_SETCHECK,BST_CHECKED,0);

        /* ── Save / Cancel ── */
        CreateWindow("BUTTON",edit?"Save":"Add",WS_VISIBLE|WS_CHILD|BS_DEFPUSHBUTTON,
                     220,390,85,28,hwnd,(HMENU)IDC_OK,g_hInst,NULL);
        CreateWindow("BUTTON","Cancel",WS_VISIBLE|WS_CHILD|BS_PUSHBUTTON,
                     315,390,85,28,hwnd,(HMENU)IDC_CANCEL,g_hInst,NULL);

        UpdateAddDlgFields(hwnd);
        return 0;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_SEP_CHECK:
        case IDC_CAT_CHECK:
            UpdateAddDlgFields(hwnd); break;

        case IDC_ICON_BROWSE: {
            OPENFILENAME ofn; char file[MAX_PATH]="";
            ZeroMemory(&ofn,sizeof(ofn)); ofn.lStructSize=sizeof(ofn); ofn.hwndOwner=hwnd;
            ofn.lpstrFile=file; ofn.nMaxFile=MAX_PATH;
            ofn.lpstrFilter="ICO files (*.ico)\0*.ico\0All Files\0*.*\0";
            ofn.Flags=OFN_FILEMUSTEXIST|OFN_PATHMUSTEXIST;
            if(GetOpenFileName(&ofn)) SetDlgItemText(hwnd,IDC_ICON_PATH_EDIT,file);
            break;
        }
        case IDC_BTN_COLOR_BTN: {
            CHOOSECOLOR cc={0}; cc.lStructSize=sizeof(cc); cc.hwndOwner=hwnd;
            cc.lpCustColors=g_customColors; cc.rgbResult=g_settingBtnColor;
            cc.Flags=CC_FULLOPEN|CC_RGBINIT;
            if(ChooseColor(&cc)){ g_settingBtnColor=cc.rgbResult; InvalidateRect(GetDlgItem(hwnd,IDC_BTN_COLOR_BTN),NULL,TRUE); }
            break;
        }
        case IDC_HK_CLEAR:
            SendDlgItemMessage(hwnd,IDC_HK_CTRL, BM_SETCHECK,BST_UNCHECKED,0);
            SendDlgItemMessage(hwnd,IDC_HK_SHIFT,BM_SETCHECK,BST_UNCHECKED,0);
            SendDlgItemMessage(hwnd,IDC_HK_ALT,  BM_SETCHECK,BST_UNCHECKED,0);
            SendDlgItemMessage(hwnd,IDC_HK_KEY,  CB_SETCURSEL,0,0);
            break;

        case IDC_OK: {
            BOOL isSep = (SendDlgItemMessage(hwnd,IDC_SEP_CHECK,BM_GETCHECK,0,0)==BST_CHECKED);
            BOOL isCat = (SendDlgItemMessage(hwnd,IDC_CAT_CHECK,BM_GETCHECK,0,0)==BST_CHECKED);
            char name[256], iconPath[MAX_PATH], rawText[MAX_TEXT];
            GetDlgItemText(hwnd,IDC_NAME_EDIT,     name,    256);
            GetDlgItemText(hwnd,IDC_ICON_PATH_EDIT,iconPath,MAX_PATH);
            GetDlgItemText(hwnd,IDC_TEXT_EDIT,     rawText, MAX_TEXT);

            if(!isSep && !name[0]){
                MessageBox(hwnd,"Please enter a display name.","Missing info",MB_OK|MB_ICONWARNING);
                break;
            }
            int ti = (g_editIndex>=0) ? g_editIndex : g_count;
            if(ti>=MAX_BUTTONS){ MessageBox(hwnd,"Maximum buttons reached.","Error",MB_OK|MB_ICONWARNING); break; }

            ButtonConfig *bc = &g_buttons[ti];
            memset(bc, 0, sizeof(ButtonConfig));

            if(isSep) {
                strcpy(bc->name,"---"); bc->isSeparator=1;
            } else if(isCat) {
                strcpy(bc->name,name); bc->isCategory=1;
            } else {
                strcpy(bc->name,     name);
                strcpy(bc->text,     rawText);
                strcpy(bc->iconPath, iconPath);
                bc->hasColor    = (SendDlgItemMessage(hwnd,IDC_BTN_COLOR_CHECK,BM_GETCHECK,0,0)==BST_CHECKED)?1:0;
                bc->btnColor    = g_settingBtnColor;
                bc->showIcon    = (SendDlgItemMessage(hwnd,IDC_ICON_CHECK,BM_GETCHECK,0,0)==BST_CHECKED)?1:0;
                /* hotkey */
                int hkmod = 0;
                if(SendDlgItemMessage(hwnd,IDC_HK_CTRL, BM_GETCHECK,0,0)==BST_CHECKED) hkmod|=MOD_CONTROL;
                if(SendDlgItemMessage(hwnd,IDC_HK_SHIFT,BM_GETCHECK,0,0)==BST_CHECKED) hkmod|=MOD_SHIFT;
                if(SendDlgItemMessage(hwnd,IDC_HK_ALT,  BM_GETCHECK,0,0)==BST_CHECKED) hkmod|=MOD_ALT;
                int keysel = (int)SendDlgItemMessage(hwnd,IDC_HK_KEY,CB_GETCURSEL,0,0);
                if(keysel<0) keysel=0;
                int hkvk = (keysel < HK_KEY_COUNT) ? g_hkVKs[keysel] : 0;
                bc->hotkeyMod = hkvk ? hkmod : 0;
                bc->hotkeyVk  = hkvk;
            }
            if(g_editIndex<0) g_count++;
            if(g_editIndex>=0 && g_editIndex<MAX_BUTTONS) g_collapsed[ti]=0;

            LoadButtonIcons();
            RegisterAllHotkeys();
            SaveAll();
            RefreshMainWindow();
            DestroyWindow(hwnd); g_hwndDlg=NULL; g_editIndex=-1;
            break;
        }
        case IDC_CANCEL:
            DestroyWindow(hwnd); g_hwndDlg=NULL; g_editIndex=-1; break;
        }
        return 0;

    case WM_DESTROY:
        if(g_hbrAddBg){ DeleteObject(g_hbrAddBg); g_hbrAddBg=NULL; }
        g_hwndDlg=NULL; g_editIndex=-1; return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/* ── Info dialog ─────────────────────────────────────────────────────── */
static HBRUSH g_hbrInfoBg = NULL;

static LRESULT CALLBACK InfoDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    LRESULT dark = HandleDlgDarkColor(hwnd, msg, wParam, &g_hbrInfoBg);
    if (dark != -1) return dark;
    switch (msg) {
    case WM_CREATE: {
        HWND hE = CreateWindowEx(WS_EX_CLIENTEDGE,"EDIT",g_infoDlgContent,
            WS_VISIBLE|WS_CHILD|WS_VSCROLL|ES_MULTILINE|ES_READONLY|ES_AUTOVSCROLL,
            10,10,460,300,hwnd,(HMENU)IDC_INFO_TEXT,g_hInst,NULL);
        if(g_darkMode) SendMessage(hE,0x0443,0,(LPARAM)DK_BG);
        CreateWindow("BUTTON","OK",WS_VISIBLE|WS_CHILD|BS_DEFPUSHBUTTON,
            195,320,90,28,hwnd,(HMENU)IDC_INFO_OK,g_hInst,NULL);
        return 0; }
    case WM_COMMAND:
        if(LOWORD(wParam)==IDC_INFO_OK){ DestroyWindow(hwnd); g_hwndDlg=NULL; }
        return 0;
    case WM_DESTROY:
        if(g_hbrInfoBg){ DeleteObject(g_hbrInfoBg); g_hbrInfoBg=NULL; }
        g_hwndDlg=NULL; return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

static void ShowInfoDialog(HWND parent, const char *title, const char *content)
{
    if(g_hwndDlg){ SetForegroundWindow(g_hwndDlg); return; }
    g_infoDlgTitle=title; g_infoDlgContent=content;
    g_hwndDlg = CreateWindowEx(WS_EX_DLGMODALFRAME,"InfoDlgClass",title,
        WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU, CW_USEDEFAULT,CW_USEDEFAULT,494,395,
        parent,NULL,g_hInst,NULL);
    SetTitleBarDark(g_hwndDlg,g_darkMode);
    ShowWindow(g_hwndDlg,SW_SHOW); UpdateWindow(g_hwndDlg);
}

/* ── Prompt dialog ───────────────────────────────────────────────────── */
static HBRUSH g_hbrPromptBg = NULL;

static LRESULT CALLBACK PromptDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    LRESULT dark = HandleDlgDarkColor(hwnd, msg, wParam, &g_hbrPromptBg);
    if (dark != -1) return dark;
    switch (msg) {
    case WM_CREATE:
        CreateWindow("STATIC","Enter value:",WS_VISIBLE|WS_CHILD,10,10,280,18,hwnd,NULL,g_hInst,NULL);
        CreateWindow("EDIT","",WS_VISIBLE|WS_CHILD|WS_BORDER|WS_TABSTOP|ES_AUTOHSCROLL,
            10,32,282,22,hwnd,(HMENU)IDC_PROMPT_EDIT,g_hInst,NULL);
        CreateWindow("BUTTON","OK",WS_VISIBLE|WS_CHILD|BS_DEFPUSHBUTTON|WS_TABSTOP,
            110,64,80,26,hwnd,(HMENU)IDC_PROMPT_OK,g_hInst,NULL);
        CreateWindow("BUTTON","Cancel",WS_VISIBLE|WS_CHILD|BS_PUSHBUTTON|WS_TABSTOP,
            200,64,80,26,hwnd,(HMENU)IDC_PROMPT_CANCEL,g_hInst,NULL);
        SetFocus(GetDlgItem(hwnd,IDC_PROMPT_EDIT));
        return 0;
    case WM_COMMAND:
        if(LOWORD(wParam)==IDC_PROMPT_OK){
            GetDlgItemText(hwnd,IDC_PROMPT_EDIT,g_promptResult,(int)sizeof(g_promptResult));
            g_promptCancelled=0; g_promptDone=1; DestroyWindow(hwnd);
        } else if(LOWORD(wParam)==IDC_PROMPT_CANCEL){
            g_promptResult[0]='\0'; g_promptCancelled=1; g_promptDone=1; DestroyWindow(hwnd);
        }
        return 0;
    case WM_CLOSE:
        g_promptResult[0]='\0'; g_promptCancelled=1; g_promptDone=1; DestroyWindow(hwnd); return 0;
    case WM_DESTROY:
        if(g_hbrPromptBg){ DeleteObject(g_hbrPromptBg); g_hbrPromptBg=NULL; }
        g_hwndDlg=NULL; return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/* ── FireButton — shared by click and WM_HOTKEY ─────────────────────── */
static void FireButton(int idx)
{
    if (idx < 0 || idx >= g_count) return;
    if (g_buttons[idx].isSeparator || g_buttons[idx].isCategory) return;
    if (!g_prevWindow || !IsWindow(g_prevWindow)) return;

    /* 1. capture clipboard */
    char clipContent[MAX_TEXT] = "";
    if (OpenClipboard(NULL)) {
        HANDLE hc = GetClipboardData(CF_UNICODETEXT);
        if (hc) {
            WCHAR *pw = (WCHAR *)GlobalLock(hc);
            if (pw) { WideCharToMultiByte(CP_ACP, 0, pw, -1, clipContent, MAX_TEXT, NULL, NULL); GlobalUnlock(hc); }
        }
        CloseClipboard();
    }

    /* 2. handle {?} */
    char working[MAX_TEXT];
    strncpy(working, g_buttons[idx].text, MAX_TEXT - 1); working[MAX_TEXT-1] = '\0';

    if (strstr(working, "{?}")) {
        if (g_hwndDlg) { SetForegroundWindow(g_hwndDlg); return; }
        g_promptResult[0]='\0'; g_promptDone=0; g_promptCancelled=0;
        HWND hPr = CreateWindowEx(WS_EX_DLGMODALFRAME,"PromptDlgClass","Fill in the Blank",
            WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU, CW_USEDEFAULT,CW_USEDEFAULT,320,132,
            g_hwndMain,NULL,g_hInst,NULL);
        g_hwndDlg = hPr;
        SetTitleBarDark(hPr, g_darkMode);
        ShowWindow(hPr, SW_SHOW); UpdateWindow(hPr);
        MSG pmsg;
        while (!g_promptDone && GetMessage(&pmsg, NULL, 0, 0)) {
            if (IsWindow(hPr) && IsDialogMessage(hPr, &pmsg)) continue;
            TranslateMessage(&pmsg); DispatchMessage(&pmsg);
        }
        if (g_promptCancelled) return;
        char replaced[MAX_TEXT];
        const char *p = working; char *q = replaced; int rem = MAX_TEXT - 1;
        while (*p && rem > 0) {
            if (strncmp(p, "{?}", 3) == 0) {
                int l = (int)strlen(g_promptResult); if (l > rem) l = rem;
                memcpy(q, g_promptResult, l); q += l; rem -= l; p += 3;
            } else { *q++ = *p++; rem--; }
        }
        *q = '\0';
        strncpy(working, replaced, MAX_TEXT-1); working[MAX_TEXT-1] = '\0';
    }

    /* 3. expand variables */
    char expanded[MAX_TEXT];
    ExpandVariables(working, expanded, MAX_TEXT, clipContent);

    /* 4. prefix / suffix */
    char final_text[MAX_TEXT];
    int pl=(int)strlen(g_prefix), el=(int)strlen(expanded), sl=(int)strlen(g_suffix);
    if (pl+el+sl < MAX_TEXT) {
        memcpy(final_text,        g_prefix, pl);
        memcpy(final_text+pl,     expanded, el);
        memcpy(final_text+pl+el,  g_suffix, sl);
        final_text[pl+el+sl]='\0';
    } else {
        strncpy(final_text, expanded, MAX_TEXT-1); final_text[MAX_TEXT-1]='\0';
    }

    /* 5. build action list and dispatch */
    g_pendingIdx = idx;
    BuildFireActions(final_text);
    FireNextAction();
}

/* ── Main window proc ────────────────────────────────────────────────── */
static LRESULT CALLBACK MainProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {

    case WM_CREATE:
        CreateWindow("BUTTON","+ Add Button",WS_VISIBLE|WS_CHILD|BS_OWNERDRAW,
            10,10,g_winWidth-20,26, hwnd,(HMENU)ID_ADD_BTN,g_hInst,NULL);
        return 0;

    case WM_CTLCOLOREDIT: {
        HWND hCtrl = (HWND)lParam;
        if (g_darkMode && hCtrl == g_hwndSearch) {
            HDC hdc = (HDC)wParam;
            SetTextColor(hdc, DK_TEXT); SetBkColor(hdc, DK_SEARCH);
            if (!g_hbrSearchDk) g_hbrSearchDk = CreateSolidBrush(DK_SEARCH);
            return (LRESULT)g_hbrSearchDk;
        }
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    case WM_HOTKEY: {
        int hotId = (int)wParam;
        if (hotId >= ID_HOTKEY_BASE && hotId < ID_HOTKEY_BASE + MAX_BUTTONS)
            FireButton(hotId - ID_HOTKEY_BASE);
        return 0;
    }

    case WM_TIMER:
        if (wParam == 1) {
            /* Timer 1: send Ctrl+V to paste the clipboard into the target window */
            KillTimer(hwnd, 1);
            if (g_pendingIdx >= 0 && g_pendingIdx < g_count) {
                INPUT inputs[4] = {0};
                inputs[0].type=INPUT_KEYBOARD; inputs[0].ki.wVk=VK_CONTROL;
                inputs[1].type=INPUT_KEYBOARD; inputs[1].ki.wVk='V';
                inputs[2].type=INPUT_KEYBOARD; inputs[2].ki.wVk='V';        inputs[2].ki.dwFlags=KEYEVENTF_KEYUP;
                inputs[3].type=INPUT_KEYBOARD; inputs[3].ki.wVk=VK_CONTROL; inputs[3].ki.dwFlags=KEYEVENTF_KEYUP;
                SendInput(4, inputs, sizeof(INPUT));
                SetTimer(hwnd, 2, 300, NULL);
            }
        } else if (wParam == 2) {
            /* Timer 2: restore previous clipboard contents, then fire next action */
            KillTimer(hwnd, 2);
            if (OpenClipboard(g_hwndMain)) {
                EmptyClipboard();
                if (g_hOldClip) { SetClipboardData(CF_UNICODETEXT, g_hOldClip); g_hOldClip=NULL; }
                CloseClipboard();
            }
            FireNextAction();
        } else if (wParam == 3) {
            /* Timer 3: send a system key (e.g. Tab, Enter, Esc) via SendInput */
            KillTimer(hwnd, 3);
            if (g_pendingVk) {
                INPUT ki[2] = {0};
                ki[0].type = INPUT_KEYBOARD; ki[0].ki.wVk = (WORD)g_pendingVk;
                ki[1].type = INPUT_KEYBOARD; ki[1].ki.wVk = (WORD)g_pendingVk;
                ki[1].ki.dwFlags = KEYEVENTF_KEYUP;
                SendInput(2, ki, sizeof(INPUT));
                g_pendingVk = 0;
            }
            SetTimer(hwnd, 4, 50, NULL);   /* brief pause before next action */
        } else if (wParam == 4) {
            /* Timer 4: post-key pause — fire the next action in the list */
            KillTimer(hwnd, 4);
            FireNextAction();
        }
        return 0;

    case WM_CLOSE:
        if (g_minToTray) { ShowWindow(hwnd, SW_HIDE); AddTrayIcon(); }
        else             { SaveAll(); DestroyWindow(hwnd); }
        return 0;

    case WM_TRAYICON:
        if (lParam == WM_LBUTTONDBLCLK) {
            RemoveTrayIcon(); ShowWindow(hwnd,SW_RESTORE); SetForegroundWindow(hwnd);
        } else if (lParam == WM_RBUTTONUP) {
            POINT pt; GetCursorPos(&pt);
            /* Build a Profiles sub-menu */
            HMENU hSub = CreatePopupMenu();
            for (int i = 0; i < g_profileCount; i++) {
                UINT f = MF_STRING | (i == g_activeProfile ? MF_CHECKED : 0);
                AppendMenu(hSub, f, IDM_PROFILE_BASE + i, g_profileNames[i]);
            }
            AppendMenu(hSub, MF_SEPARATOR, 0, NULL);
            AppendMenu(hSub, MF_STRING, IDM_PROFILE_NEW, "New Profile...");
            AppendMenu(hSub, MF_STRING | (g_activeProfile==0 ? MF_GRAYED : 0),
                       IDM_PROFILE_DELETE, "Delete Current Profile");
            HMENU hM = CreatePopupMenu();
            AppendMenu(hM, MF_POPUP, (UINT_PTR)hSub, "Profiles");
            AppendMenu(hM, MF_SEPARATOR, 0, NULL);
            AppendMenu(hM, MF_STRING, IDM_TRAY_RESTORE, "Restore");
            AppendMenu(hM, MF_SEPARATOR, 0, NULL);
            AppendMenu(hM, MF_STRING, IDM_TRAY_EXIT, "Exit");
            SetForegroundWindow(hwnd);
            int cmd = TrackPopupMenu(hM, TPM_RETURNCMD|TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
            DestroyMenu(hM); /* also destroys hSub */
            if      (cmd == IDM_TRAY_RESTORE) { RemoveTrayIcon(); ShowWindow(hwnd,SW_RESTORE); SetForegroundWindow(hwnd); }
            else if (cmd == IDM_TRAY_EXIT)    { RemoveTrayIcon(); SaveAll(); DestroyWindow(hwnd); }
            else if (cmd >= IDM_PROFILE_BASE && cmd < IDM_PROFILE_BASE + g_profileCount)
                SwitchProfile(cmd - IDM_PROFILE_BASE);
            else if (cmd == IDM_PROFILE_NEW)    PostMessage(hwnd, WM_COMMAND, IDM_PROFILE_NEW,    0);
            else if (cmd == IDM_PROFILE_DELETE) PostMessage(hwnd, WM_COMMAND, IDM_PROFILE_DELETE, 0);
        }
        return 0;

    case WM_CONTEXTMENU: {
        HWND hClicked = (HWND)wParam;
        POINT pt = { (short)LOWORD(lParam), (short)HIWORD(lParam) };
        g_ctxIndex = -1;
        for (int i = 0; i < g_count; i++)
            if (g_hwndBtns[i] == hClicked) { g_ctxIndex = i; break; }
        if (g_ctxIndex >= 0) ShowButtonContextMenu(hwnd, g_ctxIndex, pt);
        return 0;
    }

    case WM_ERASEBKGND:
        if (g_darkMode) {
            HDC hdc=(HDC)wParam; RECT rc;
            GetClientRect(hwnd,&rc);
            if(!g_hbrDkBg) g_hbrDkBg=CreateSolidBrush(DK_BG);
            FillRect(hdc,&rc,g_hbrDkBg); return 1;
        }
        return DefWindowProc(hwnd, msg, wParam, lParam);

    case WM_NCPAINT: {
        LRESULT res = DefWindowProc(hwnd, msg, wParam, lParam);
        if (g_darkMode) {
            HMENU hMenu = GetMenu(hwnd); int cnt = hMenu ? GetMenuItemCount(hMenu) : 0;
            MENUBARINFO mbiBar = { sizeof(mbiBar) };
            if (cnt > 0 && GetMenuBarInfo(hwnd, OBJID_MENU, 0, &mbiBar)) {
                RECT rcWin; GetWindowRect(hwnd,&rcWin);
                HDC hdc = GetWindowDC(hwnd);
                RECT rcBar = { mbiBar.rcBar.left-rcWin.left, mbiBar.rcBar.top-rcWin.top,
                               mbiBar.rcBar.right-rcWin.left, mbiBar.rcBar.bottom-rcWin.top };
                HBRUSH hbr = CreateSolidBrush(DK_MENU_BG); FillRect(hdc,&rcBar,hbr); DeleteObject(hbr);
                SetBkMode(hdc,TRANSPARENT); SetTextColor(hdc,DK_TEXT);
                HFONT hOld=(HFONT)SelectObject(hdc,(HFONT)GetStockObject(DEFAULT_GUI_FONT));
                int n=sizeof(g_menuLabels)/sizeof(g_menuLabels[0]);
                for (int i=0; i<cnt&&i<n; i++) {
                    MENUBARINFO mbi={sizeof(mbi)};
                    if(!GetMenuBarInfo(hwnd,OBJID_MENU,i+1,&mbi)) continue;
                    RECT rc2={mbi.rcBar.left-rcWin.left,mbi.rcBar.top-rcWin.top,
                              mbi.rcBar.right-rcWin.left,mbi.rcBar.bottom-rcWin.top};
                    DrawText(hdc,g_menuLabels[i],-1,&rc2,DT_CENTER|DT_VCENTER|DT_SINGLELINE);
                }
                SelectObject(hdc,hOld); ReleaseDC(hwnd,hdc);
            }
        }
        return res;
    }

    case WM_NCACTIVATE: {
        LRESULT res = DefWindowProc(hwnd, msg, wParam, lParam);
        if (g_darkMode) {
            HMENU hMenu = GetMenu(hwnd); int cnt = hMenu ? GetMenuItemCount(hMenu) : 0;
            MENUBARINFO mbiBar = { sizeof(mbiBar) };
            if (cnt > 0 && GetMenuBarInfo(hwnd, OBJID_MENU, 0, &mbiBar)) {
                RECT rcWin; GetWindowRect(hwnd,&rcWin);
                HDC hdc = GetWindowDC(hwnd);
                RECT rcBar = { mbiBar.rcBar.left-rcWin.left, mbiBar.rcBar.top-rcWin.top,
                               mbiBar.rcBar.right-rcWin.left, mbiBar.rcBar.bottom-rcWin.top };
                HBRUSH hbr = CreateSolidBrush(DK_MENU_BG); FillRect(hdc,&rcBar,hbr); DeleteObject(hbr);
                SetBkMode(hdc,TRANSPARENT); SetTextColor(hdc,DK_TEXT);
                HFONT hOld=(HFONT)SelectObject(hdc,(HFONT)GetStockObject(DEFAULT_GUI_FONT));
                int n=sizeof(g_menuLabels)/sizeof(g_menuLabels[0]);
                for (int i=0; i<cnt&&i<n; i++) {
                    MENUBARINFO mbi={sizeof(mbi)};
                    if(!GetMenuBarInfo(hwnd,OBJID_MENU,i+1,&mbi)) continue;
                    RECT rc2={mbi.rcBar.left-rcWin.left,mbi.rcBar.top-rcWin.top,
                              mbi.rcBar.right-rcWin.left,mbi.rcBar.bottom-rcWin.top};
                    DrawText(hdc,g_menuLabels[i],-1,&rc2,DT_CENTER|DT_VCENTER|DT_SINGLELINE);
                }
                SelectObject(hdc,hOld); ReleaseDC(hwnd,hdc);
            }
        }
        return res;
    }

    case WM_MEASUREITEM: {
        MEASUREITEMSTRUCT *mis=(MEASUREITEMSTRUCT*)lParam;
        if(mis->CtlType==ODT_MENU){
            const char *label=(const char*)mis->itemData;
            HFONT hMenuFont=(HFONT)GetStockObject(DEFAULT_GUI_FONT);
            HDC hdc=GetDC(hwnd);
            HFONT hOld=(HFONT)SelectObject(hdc,hMenuFont);
            SIZE sz; GetTextExtentPoint32(hdc,label,(int)strlen(label),&sz);
            SelectObject(hdc,hOld); ReleaseDC(hwnd,hdc);
            mis->itemWidth  = sz.cx + 20;
            mis->itemHeight = GetSystemMetrics(SM_CYMENU); /* match system bar height exactly */
            return TRUE;
        }
        return DefWindowProc(hwnd,msg,wParam,lParam);
    }

    case WM_DRAWITEM: {
        LPDRAWITEMSTRUCT dis=(LPDRAWITEMSTRUCT)lParam;
        if(dis->CtlType==ODT_MENU){
            const char *label=(const char*)dis->itemData;
            BOOL hot=(dis->itemState&(ODS_SELECTED|ODS_HOTLIGHT))!=0;
            HBRUSH hbr=CreateSolidBrush(hot?DK_MENU_HOT:DK_MENU_BG);
            FillRect(dis->hDC,&dis->rcItem,hbr); DeleteObject(hbr);
            SetBkMode(dis->hDC,TRANSPARENT); SetTextColor(dis->hDC,DK_TEXT);
            /* Explicitly select the same font used in WM_NCPAINT to prevent
               a font/size jump between the painted state and the hover state. */
            HFONT hOld=(HFONT)SelectObject(dis->hDC,(HFONT)GetStockObject(DEFAULT_GUI_FONT));
            DrawText(dis->hDC,label,-1,&dis->rcItem,DT_CENTER|DT_VCENTER|DT_SINGLELINE);
            SelectObject(dis->hDC,hOld);
            return TRUE;
        }
        int id=(int)dis->CtlID;
        if(id==ID_ADD_BTN)                                       { DrawButton(dis,-1);               return TRUE; }
        if(id>=ID_BUTTON_BASE && id<ID_BUTTON_BASE+g_count)     { DrawButton(dis,id-ID_BUTTON_BASE); return TRUE; }
        return DefWindowProc(hwnd,msg,wParam,lParam);
    }

    case WM_COMMAND: {
        int id    = LOWORD(wParam);
        int notif = HIWORD(wParam);

        /* Search bar */
        if(id==IDC_SEARCH_EDIT && notif==EN_CHANGE){
            GetWindowText(g_hwndSearch, g_filterText, sizeof(g_filterText));
            RefreshMainWindow(); return 0;
        }

        if(id==IDM_MOVE_UP && g_ctxIndex>0){
            ButtonConfig tmp=g_buttons[g_ctxIndex]; g_buttons[g_ctxIndex]=g_buttons[g_ctxIndex-1]; g_buttons[g_ctxIndex-1]=tmp;
            HICON hi=g_icons[g_ctxIndex]; g_icons[g_ctxIndex]=g_icons[g_ctxIndex-1]; g_icons[g_ctxIndex-1]=hi;
            int ci=g_collapsed[g_ctxIndex]; g_collapsed[g_ctxIndex]=g_collapsed[g_ctxIndex-1]; g_collapsed[g_ctxIndex-1]=ci;
            RegisterAllHotkeys(); SaveAll(); RefreshMainWindow(); g_ctxIndex=-1;

        } else if(id==IDM_MOVE_DOWN && g_ctxIndex>=0 && g_ctxIndex<g_count-1){
            ButtonConfig tmp=g_buttons[g_ctxIndex]; g_buttons[g_ctxIndex]=g_buttons[g_ctxIndex+1]; g_buttons[g_ctxIndex+1]=tmp;
            HICON hi=g_icons[g_ctxIndex]; g_icons[g_ctxIndex]=g_icons[g_ctxIndex+1]; g_icons[g_ctxIndex+1]=hi;
            int ci=g_collapsed[g_ctxIndex]; g_collapsed[g_ctxIndex]=g_collapsed[g_ctxIndex+1]; g_collapsed[g_ctxIndex+1]=ci;
            RegisterAllHotkeys(); SaveAll(); RefreshMainWindow(); g_ctxIndex=-1;

        } else if(id==IDM_EDIT_BTN && g_ctxIndex>=0){
            if(g_hwndDlg){ SetForegroundWindow(g_hwndDlg); return 0; }
            g_editIndex=g_ctxIndex; g_ctxIndex=-1;
            g_hwndDlg=CreateWindowEx(WS_EX_DLGMODALFRAME,"AddDlgClass","Edit Button",
                WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU, CW_USEDEFAULT,CW_USEDEFAULT,420,462,
                hwnd,NULL,g_hInst,NULL);
            SetTitleBarDark(g_hwndDlg,g_darkMode);
            ShowWindow(g_hwndDlg,SW_SHOW); UpdateWindow(g_hwndDlg);

        } else if(id==IDM_DUPLICATE_BTN && g_ctxIndex>=0){
            if(g_count>=MAX_BUTTONS){ MessageBox(hwnd,"Maximum buttons reached.","Error",MB_OK|MB_ICONWARNING); }
            else {
                for(int i=g_count; i>g_ctxIndex+1; i--){
                    g_buttons[i]=g_buttons[i-1]; g_icons[i]=NULL; g_collapsed[i]=g_collapsed[i-1];
                }
                g_buttons[g_ctxIndex+1]=g_buttons[g_ctxIndex];
                char *nm=g_buttons[g_ctxIndex+1].name;
                if((int)strlen(nm)+7<256) strcat(nm," (copy)");
                g_icons[g_ctxIndex+1]=NULL;
                g_collapsed[g_ctxIndex+1]=0;
                g_count++;
                LoadButtonIcons(); RegisterAllHotkeys(); SaveAll(); RefreshMainWindow();
            }
            g_ctxIndex=-1;

        } else if(id==IDM_DELETE_BTN && g_ctxIndex>=0){
            char confirm[300]; sprintf(confirm,"Delete \"%s\"?",g_buttons[g_ctxIndex].name);
            if(MessageBox(hwnd,confirm,"Confirm Delete",MB_YESNO|MB_ICONQUESTION)==IDYES){
                if(g_icons[g_ctxIndex]){ DestroyIcon(g_icons[g_ctxIndex]); g_icons[g_ctxIndex]=NULL; }
                for(int i=g_ctxIndex; i<g_count-1; i++){
                    g_buttons[i]=g_buttons[i+1]; g_icons[i]=g_icons[i+1]; g_collapsed[i]=g_collapsed[i+1];
                }
                g_icons[g_count-1]=NULL; g_count--;
                RegisterAllHotkeys(); SaveAll(); RefreshMainWindow();
            }
            g_ctxIndex=-1;

        } else if(id >= IDM_PROFILE_BASE && id < IDM_PROFILE_BASE + g_profileCount){
            SwitchProfile(id - IDM_PROFILE_BASE);

        } else if(id == ID_PROFILES_MENU){
            POINT pt; GetCursorPos(&pt);
            ShowProfilesMenu(hwnd, pt.x, pt.y);

        } else if(id == IDM_PROFILE_NEW){
            /* Reuse prompt dialog for the profile name */
            if(g_hwndDlg){ SetForegroundWindow(g_hwndDlg); return 0; }
            g_promptResult[0]='\0'; g_promptDone=0; g_promptCancelled=0;
            HWND hPr = CreateWindowEx(WS_EX_DLGMODALFRAME,"PromptDlgClass","New Profile Name",
                WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU,CW_USEDEFAULT,CW_USEDEFAULT,320,132,
                hwnd,NULL,g_hInst,NULL);
            g_hwndDlg=hPr; SetTitleBarDark(hPr,g_darkMode);
            ShowWindow(hPr,SW_SHOW); UpdateWindow(hPr);
            MSG pmsg;
            while(!g_promptDone && GetMessage(&pmsg,NULL,0,0)){
                if(IsWindow(hPr)&&IsDialogMessage(hPr,&pmsg)) continue;
                TranslateMessage(&pmsg); DispatchMessage(&pmsg);
            }
            if(!g_promptCancelled && g_promptResult[0]){
                /* Sanitise: strip chars illegal in filenames */
                char safe[64]; int si=0;
                for(int i=0; g_promptResult[i]&&si<63; i++){
                    char c=g_promptResult[i];
                    if(c!='/'&&c!='\\'&&c!=':'&&c!='*'&&c!='?'&&c!='"'&&c!='<'&&c!='>'&&c!='|')
                        safe[si++]=c;
                }
                safe[si]='\0';
                if(!safe[0] || _stricmp(safe,"Default")==0){
                    MessageBox(hwnd,"Invalid or reserved profile name.","Error",MB_OK|MB_ICONWARNING);
                } else {
                    /* Build typer_<name>.ini path next to exe */
                    char dir[MAX_PATH]; strcpy(dir,g_basePath);
                    char *ls=strrchr(dir,'\\'); if(!ls) ls=strrchr(dir,'/');
                    if(ls) *(ls+1)='\0'; else dir[0]='\0';
                    char newPath[MAX_PATH];
                    sprintf(newPath,"%styper_%s.ini",dir,safe);
                    /* Create an empty file if it doesn't exist yet */
                    FILE *tf=fopen(newPath,"a"); if(tf) fclose(tf);
                    ScanProfiles();
                    /* Switch to the new profile */
                    for(int i=0;i<g_profileCount;i++)
                        if(_stricmp(g_profileNames[i],safe)==0){ SwitchProfile(i); break; }
                }
            }

        } else if(id == IDM_PROFILE_DELETE){
            if(g_activeProfile==0){
                MessageBox(hwnd,"Cannot delete the Default profile.","Error",MB_OK|MB_ICONWARNING);
            } else {
                char confirm[200];
                sprintf(confirm,"Delete profile \"%s\"?",g_profileNames[g_activeProfile]);
                if(MessageBox(hwnd,confirm,"Confirm Delete",MB_YESNO|MB_ICONQUESTION)==IDYES){
                    char delPath[MAX_PATH]; strcpy(delPath,g_profilePaths[g_activeProfile]);
                    SwitchProfile(0);     /* switch to Default before deleting */
                    DeleteFile(delPath);
                    ScanProfiles();
                    RebuildMenu();
                }
            }

        } else if(id==ID_HELP_INSTRUCTIONS){
            ShowInfoDialog(hwnd,"Instructions",
                "HOW IT WORKS\r\n"
                "Simple Typer shows a list of buttons. Click a button to type "
                "its stored text into whatever window you were in before clicking.\r\n"
                "\r\n"
                "USAGE\r\n"
                "1. Click in a text field in any application.\r\n"
                "2. Click a button in Simple Typer.\r\n"
                "3. The stored text is pasted into your previous text field.\r\n"
                "\r\n"
                "VARIABLES\r\n"
                "  {date}      - today's date (MM/DD/YYYY)\r\n"
                "  {time}      - current time (HH:MM)\r\n"
                "  {clipboard} - your clipboard contents\r\n"
                "  {?}         - pops an input box; your answer replaces {?}\r\n"
                "\r\n"
                "SYSTEM KEYS\r\n"
                "Embed these tokens to send a keystroke at that point in the text:\r\n"
                "  {tab}       - Tab key\r\n"
                "  {enter}     - Enter / Return key\r\n"
                "  {esc}       - Escape key\r\n"
                "  {backspace} - Backspace key\r\n"
                "  {delete}    - Delete key\r\n"
                "  {up} {down} {left} {right} - Arrow keys\r\n"
                "  {home} {end} {pgup} {pgdn} - Navigation keys\r\n"
                "Example: \"Hello{tab}World{enter}\" types Hello, presses Tab,\r\n"
                "types World, then presses Enter.\r\n"
                "\r\n"
                "CATEGORIES\r\n"
                "Add a 'Category header' item to group buttons under a labelled "
                "collapsible section. Click the category bar to collapse or expand "
                "the buttons beneath it. A '>' means collapsed; 'v' means expanded. "
                "Categories are skipped in compact mode and search/filter mode.\r\n"
                "\r\n"
                "KEYBOARD SHORTCUTS\r\n"
                "Each button can have a global hotkey (Ctrl/Shift/Alt + key). "
                "Set it in the Add/Edit dialog. The hotkey fires the button even when "
                "Simple Typer is not focused. Hotkeys are shown in the button tooltip. "
                "At least one modifier (Ctrl, Shift, or Alt) is recommended to avoid "
                "conflicts with normal typing.\r\n"
                "\r\n"
                "SEARCH / FILTER\r\n"
                "Type in the search bar to filter buttons by name. Category headers "
                "and separators are hidden while a filter is active.\r\n"
                "\r\n"
                "COMPACT MODE\r\n"
                "Enable in Settings for a small icon-grid palette. Each tile shows "
                "its icon or first letter. Hover for the full button name.\r\n"
                "\r\n"
                "RIGHT-CLICK BUTTONS\r\n"
                "Edit, Delete, Duplicate, Move Up, Move Down.\r\n"
                "\r\n"
                "SETTINGS\r\n"
                "Dark mode, font size, window width, always on top, minimize to "
                "system tray, compact mode, window title, opacity, prefix/suffix.\r\n"
                "\r\n"
                "CONFIGURATION FILE\r\n"
                "All settings are saved to typer.ini (Default profile) or "
                "typer_<name>.ini for named profiles, in the same folder as typer.exe.\r\n"
                "\r\n"
                "PROFILES\r\n"
                "Profiles let you maintain separate sets of buttons in different INI files. "
                "Use the Profiles menu (menu bar or tray right-click) to switch profiles, "
                "create new ones, or delete the current one. "
                "The Default profile (typer.ini) cannot be deleted.");

        } else if(id==ID_SETTINGS){
            if(g_hwndDlg){ SetForegroundWindow(g_hwndDlg); return 0; }
            g_hwndDlg=CreateWindowEx(WS_EX_DLGMODALFRAME,"SettingsDlgClass","Settings",
                WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU, CW_USEDEFAULT,CW_USEDEFAULT,295,458,
                hwnd,NULL,g_hInst,NULL);
            SetTitleBarDark(g_hwndDlg,g_darkMode);
            ShowWindow(g_hwndDlg,SW_SHOW); UpdateWindow(g_hwndDlg);

        } else if(id==ID_HELP_ABOUT){
            ShowInfoDialog(hwnd,"About Simple Typer",
                "Simple Typer\r\nVersion 2.11\r\n\r\n"
                "Author:   UberGuidoZ\r\n"
                "Contact:  https://github.com/UberGuidoZ");

        } else if(id==ID_ADD_BTN){
            if(g_hwndDlg){ SetForegroundWindow(g_hwndDlg); return 0; }
            g_editIndex=-1;
            g_hwndDlg=CreateWindowEx(WS_EX_DLGMODALFRAME,"AddDlgClass","Add Button",
                WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU, CW_USEDEFAULT,CW_USEDEFAULT,420,462,
                hwnd,NULL,g_hInst,NULL);
            SetTitleBarDark(g_hwndDlg,g_darkMode);
            ShowWindow(g_hwndDlg,SW_SHOW); UpdateWindow(g_hwndDlg);

        } else if(id>=ID_BUTTON_BASE && id<ID_BUTTON_BASE+g_count){
            int idx = id - ID_BUTTON_BASE;

            /* Category header: toggle collapse */
            if(g_buttons[idx].isCategory){
                g_collapsed[idx] = !g_collapsed[idx];
                RefreshMainWindow();
                return 0;
            }

            FireButton(idx);
        }
        return 0;
    }

    case WM_DESTROY:
        UnregisterAllHotkeys();
        if(g_hEventHook) { UnhookWinEvent(g_hEventHook); g_hEventHook=NULL; }
        if(g_hwndTooltip){ DestroyWindow(g_hwndTooltip); g_hwndTooltip=NULL; }
        if(g_hwndSearch) { DestroyWindow(g_hwndSearch);  g_hwndSearch=NULL;  }
        if(g_hbrSearchDk){ DeleteObject(g_hbrSearchDk);  g_hbrSearchDk=NULL; }
        RemoveTrayIcon();
        FreeIcons();
        SaveAll();
        if(g_hbrDkBg)  { DeleteObject(g_hbrDkBg);   g_hbrDkBg=NULL;   }
        if(g_hFont)    { DeleteObject(g_hFont);      g_hFont=NULL;     }
        if(g_hFontBold){ DeleteObject(g_hFontBold);  g_hFontBold=NULL; }
        PostQuitMessage(0); return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/* ── Entry point ─────────────────────────────────────────────────────── */
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrev, LPSTR lpCmd, int nShow)
{
    g_hInst = hInstance;
    GetBasePath();
    ScanProfiles();
    LoadSettings();
    LoadButtons();
    RecreateFont();
    LoadButtonIcons();
    if (g_darkMode) {
        g_hbrDkBg     = CreateSolidBrush(DK_BG);
        g_hbrSearchDk = CreateSolidBrush(DK_SEARCH);
    }

    WNDCLASS wc = {0};
    wc.hInstance = hInstance; wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpfnWndProc=InfoDlgProc;     wc.lpszClassName="InfoDlgClass";     RegisterClass(&wc);
    wc.lpfnWndProc=SettingsDlgProc; wc.lpszClassName="SettingsDlgClass"; RegisterClass(&wc);
    wc.lpfnWndProc=AddDlgProc;      wc.lpszClassName="AddDlgClass";      RegisterClass(&wc);
    wc.lpfnWndProc=PromptDlgProc;   wc.lpszClassName="PromptDlgClass";   RegisterClass(&wc);

    WNDCLASSEX wcx = {0};
    wcx.cbSize=sizeof(WNDCLASSEX); wcx.lpfnWndProc=MainProc; wcx.hInstance=hInstance;
    wcx.lpszClassName="TyperMain"; wcx.hbrBackground=(HBRUSH)(COLOR_WINDOW+1);
    wcx.hCursor=LoadCursor(NULL,IDC_ARROW);
    wcx.hIcon  =LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APPICON));
    wcx.hIconSm=LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APPICON));
    RegisterClassEx(&wcx);

    int startX=(g_winX!=-1&&g_winY!=-1)?g_winX:CW_USEDEFAULT;
    int startY=(g_winX!=-1&&g_winY!=-1)?g_winY:CW_USEDEFAULT;

    g_hwndMain = CreateWindow("TyperMain", g_winTitle, WS_OVERLAPPEDWINDOW,
                              startX, startY, g_winWidth+22, 120,
                              NULL, NULL, hInstance, NULL);

    RebuildMenu();
    RefreshMainWindow();
    SetTitleBarDark(g_hwndMain, g_darkMode);
    ApplyOpacity();
    if(g_alwaysOnTop)
        SetWindowPos(g_hwndMain, HWND_TOPMOST, 0,0,0,0, SWP_NOMOVE|SWP_NOSIZE);

    g_hEventHook = SetWinEventHook(
        EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
        NULL, WinEventProc, 0, 0,
        WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);

    RegisterAllHotkeys();

    ShowWindow(g_hwndMain, nShow);
    UpdateWindow(g_hwndMain);

    MSG msg;
    while(GetMessage(&msg, NULL, 0, 0)){
        if(g_hwndDlg && IsDialogMessage(g_hwndDlg,&msg)) continue;
        TranslateMessage(&msg); DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}
