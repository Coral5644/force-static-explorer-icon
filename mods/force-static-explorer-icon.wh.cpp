// ==WindhawkMod==
// @id              force-static-explorer-icon
// @name            Force Static Explorer Icon
// @description     强制让 Windows 11 任务栏上的所有"文件资源管理器"窗口按钮显示统一的静态图标
// @version         3.1
// @author          You
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -lshell32 -lole32 -loleaut32 -lversion
// ==/WindhawkMod==

// ==WindhawkModReadme==
/*
# Force Static Explorer Icon

修复 Windows 11 任务栏在"不合并"模式下资源管理器图标混乱的问题。

## 功能
- 强制任务栏上的资源管理器窗口显示统一的静态图标
- 支持自定义 .ico 图标文件

## 兼容性
- Windows 11 22H2/23H2/24H2
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- iconSource: explorer
  $name: 图标来源
  $options:
  - explorer: explorer.exe (默认)
  - shell32: shell32.dll
  - custom: 自定义图标
- customIconPath: ""
  $name: 自定义图标路径
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>
#include <shlobj.h>

// 全局变量
HICON g_iconLarge = NULL;
HICON g_iconSmall = NULL;

struct {
    int source;
    WCHAR customPath[MAX_PATH];
} g_settings;

// ============================================================================
// 工具函数
// ============================================================================

// 检查是否为资源管理器窗口
inline bool IsExplorerWindow(HWND hWnd) {
    WCHAR className[32];
    return GetClassNameW(hWnd, className, 32) && wcscmp(className, L"CabinetWClass") == 0;
}

// 获取图标副本（防止被系统销毁）
inline HICON GetIconCopy(bool small) {
    HICON src = small ? (g_iconSmall ? g_iconSmall : g_iconLarge) : g_iconLarge;
    return src ? CopyIcon(src) : NULL;
}

// ============================================================================
// 图标加载
// ============================================================================

bool LoadIcon() {
    // 自定义图标
    if (g_settings.source == 2 && g_settings.customPath[0]) {
        g_iconLarge = (HICON)LoadImageW(NULL, g_settings.customPath, IMAGE_ICON,
                                        GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON),
                                        LR_LOADFROMFILE);
        g_iconSmall = (HICON)LoadImageW(NULL, g_settings.customPath, IMAGE_ICON,
                                        GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON),
                                        LR_LOADFROMFILE);
        return g_iconLarge != NULL;
    }

    // 系统图标
    WCHAR path[MAX_PATH];
    int index = 0;
    
    if (g_settings.source == 0) {
        GetWindowsDirectoryW(path, MAX_PATH);
        wcscat_s(path, L"\\explorer.exe");
    } else {
        GetSystemDirectoryW(path, MAX_PATH);
        wcscat_s(path, L"\\shell32.dll");
        index = 4;
    }

    SHDefExtractIconW(path, index, 0, &g_iconLarge, &g_iconSmall,
                      MAKELONG(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CXSMICON)));
    
    if (!g_iconLarge) {
        ExtractIconExW(path, index, &g_iconLarge, &g_iconSmall, 1);
    }
    
    return g_iconLarge != NULL;
}

void FreeIcon() {
    if (g_iconLarge) { DestroyIcon(g_iconLarge); g_iconLarge = NULL; }
    if (g_iconSmall) { DestroyIcon(g_iconSmall); g_iconSmall = NULL; }
}

// ============================================================================
// Taskbar Hooks
// ============================================================================

using CTaskGroup_GetAppID_t = PCWSTR(WINAPI*)(PVOID);
using CTaskGroup_GetNumItems_t = int(WINAPI*)(PVOID);
using CTaskGroup_GetShortcutIDList_t = const ITEMIDLIST*(WINAPI*)(PVOID);

CTaskGroup_GetAppID_t CTaskGroup_GetAppID_Orig;
CTaskGroup_GetNumItems_t CTaskGroup_GetNumItems_Orig;
CTaskGroup_GetShortcutIDList_t CTaskGroup_GetShortcutIDList_Orig;

// 核心 Hook：返回 nullptr 强制任务栏查询窗口图标
const ITEMIDLIST* WINAPI CTaskGroup_GetShortcutIDList_Hook(PVOID pThis) {
    if (g_iconLarge && CTaskGroup_GetAppID_Orig) {
        PCWSTR appId = CTaskGroup_GetAppID_Orig(pThis);
        if (appId && wcsstr(appId, L"Microsoft.Windows.Explorer")) {
            int n = CTaskGroup_GetNumItems_Orig ? CTaskGroup_GetNumItems_Orig(pThis) : 1;
            if (n > 0) return nullptr;
        }
    }
    return CTaskGroup_GetShortcutIDList_Orig(pThis);
}

// ============================================================================
// Windows API Hooks
// ============================================================================

using GetClassLongPtrW_t = decltype(&GetClassLongPtrW);
using DefWindowProcW_t = decltype(&DefWindowProcW);

GetClassLongPtrW_t GetClassLongPtrW_Orig;
DefWindowProcW_t DefWindowProcW_Orig;

// 拦截窗口类图标获取
ULONG_PTR WINAPI GetClassLongPtrW_Hook(HWND hWnd, int nIndex) {
    if ((nIndex == GCLP_HICON || nIndex == GCLP_HICONSM) && g_iconLarge && IsExplorerWindow(hWnd)) {
        HICON icon = GetIconCopy(nIndex == GCLP_HICONSM);
        if (icon) return (ULONG_PTR)icon;
    }
    return GetClassLongPtrW_Orig(hWnd, nIndex);
}

// 拦截 WM_SETICON/WM_GETICON（最重要）
LRESULT WINAPI DefWindowProcW_Hook(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam) {
    if (g_iconLarge && IsExplorerWindow(hWnd)) {
        bool small = (wParam == ICON_SMALL || wParam == ICON_SMALL2);
        
        if (Msg == WM_SETICON) {
            HICON icon = GetIconCopy(small);
            if (icon) lParam = (LPARAM)icon;
        }
        else if (Msg == WM_GETICON) {
            HICON icon = GetIconCopy(small);
            if (icon) return (LRESULT)icon;
        }
    }
    return DefWindowProcW_Orig(hWnd, Msg, wParam, lParam);
}

// ============================================================================
// 初始化
// ============================================================================

void LoadSettings() {
    PCWSTR src = Wh_GetStringSetting(L"iconSource");
    if (src) {
        g_settings.source = wcscmp(src, L"shell32") == 0 ? 1 : 
                            wcscmp(src, L"custom") == 0 ? 2 : 0;
        Wh_FreeStringSetting(src);
    }
    
    PCWSTR path = Wh_GetStringSetting(L"customIconPath");
    if (path) {
        wcscpy_s(g_settings.customPath, path);
        Wh_FreeStringSetting(path);
    }
}

bool HookTaskbar() {
    HMODULE module = LoadLibraryEx(L"taskbar.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!module) return false;

    WindhawkUtils::SYMBOL_HOOK hooks[] = {
        {{LR"(public: virtual unsigned short const * __cdecl CTaskGroup::GetAppID(void))"},
         &CTaskGroup_GetAppID_Orig, nullptr, false},
        {{LR"(public: virtual int __cdecl CTaskGroup::GetNumItems(void))"},
         &CTaskGroup_GetNumItems_Orig, nullptr, false},
        {{LR"(public: virtual struct _ITEMIDLIST_ABSOLUTE const * __cdecl CTaskGroup::GetShortcutIDList(void))"},
         &CTaskGroup_GetShortcutIDList_Orig, CTaskGroup_GetShortcutIDList_Hook, false}
    };

    return WindhawkUtils::HookSymbols(module, hooks, ARRAYSIZE(hooks));
}

// ============================================================================
// 模组入口
// ============================================================================

BOOL Wh_ModInit() {
    Wh_Log(L"Force Static Explorer Icon v3.1 - Init");

    LoadSettings();
    
    if (!LoadIcon()) {
        Wh_Log(L"Failed to load icon");
    }

    if (!HookTaskbar()) {
        Wh_Log(L"Failed to hook taskbar");
        return FALSE;
    }

    Wh_SetFunctionHook((void*)GetClassLongPtrW, (void*)GetClassLongPtrW_Hook, (void**)&GetClassLongPtrW_Orig);
    Wh_SetFunctionHook((void*)DefWindowProcW, (void*)DefWindowProcW_Hook, (void**)&DefWindowProcW_Orig);

    return TRUE;
}

void Wh_ModUninit() {
    FreeIcon();
}

void Wh_ModSettingsChanged() {
    FreeIcon();
    LoadSettings();
    LoadIcon();
}
