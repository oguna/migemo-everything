// リリースビルド時にコンソールウィンドウを非表示にする
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

// --- 依存クレート ---
use windows::{
    core::*,
    Win32::Foundation::*,
    Win32::Graphics::Gdi::*,
    Win32::System::Com::{
        CoInitializeEx, CoTaskMemFree, CoUninitialize, COINIT_APARTMENTTHREADED,
    },
    Win32::System::DataExchange::{CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData},
    Win32::System::LibraryLoader::GetModuleHandleA,
    Win32::System::Memory::{GlobalAlloc, GlobalLock, GlobalUnlock, GMEM_MOVEABLE},
    Win32::System::Ole::CF_UNICODETEXT,
    Win32::System::SystemServices::SFGAO_FILESYSTEM,
    Win32::System::Time::FileTimeToSystemTime,
    Win32::Storage::FileSystem::{FILE_ATTRIBUTE_DIRECTORY, FILE_ATTRIBUTE_NORMAL},
    Win32::UI::Controls::*,
    Win32::UI::HiDpi::{
        DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2,
        GetDpiForWindow,
        SetProcessDpiAwarenessContext,
    },
    Win32::UI::Input::KeyboardAndMouse::SetFocus,
    Win32::UI::Shell::{
        Common::ITEMIDLIST, ShellExecuteW, SHFILEINFOW, SHGFI_ICON, SHGFI_SMALLICON,
        SHGFI_SYSICONINDEX, SHGFI_USEFILEATTRIBUTES, SHGetFileInfoW, SHBindToParent,
        SHParseDisplayName, CMINVOKECOMMANDINFO, CMF_NORMAL, IContextMenu, IShellFolder,
    },
    Win32::UI::WindowsAndMessaging::*,
};

use everything_sdk::ergo::{global, RequestFlags};
use rustmigemo::migemo::{
    compact_dictionary::CompactDictionary, query::query, regex_generator::RegexOperator,
};

use std::ffi::c_void;
use std::fs::File;
use std::io::Read;
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::thread;

// --- 定数 ---

/// コントロールID: テキスト入力
const EDIT_ID: u16 = 1000;
/// コントロールID: REボタン
const RE_BUTTON_ID: u16 = 1001;
/// コントロールID: Miボタン
const MI_BUTTON_ID: u16 = 1002;
/// コントロールID: シェルコンテキストメニュー切り替えトグル
const SHELL_CONTEXT_TOGGLE_ID: u16 = 1003;

/// タイマーID
const TIMER_ID: usize = 1;

/// メニューID: 終了
const IDM_FILE_EXIT: u16 = 2001;
/// メニューID: 正規表現検索
const IDM_SEARCH_REGEX: u16 = 3001;
/// メニューID: Migemo検索
const IDM_SEARCH_MIGEMO: u16 = 3002;

/// アクセラレータID: 終了
const IDA_EXIT: u16 = 5001;
/// アクセラレータID: 正規表現検索
const IDA_REGEX: u16 = 5002;
/// アクセラレータID: Migemo検索
const IDA_MIGEMO: u16 = 5003;

/// コンテキストメニューID: 開く
const IDM_CONTEXT_OPEN: u16 = 4001;
/// コンテキストメニューID: フォルダを開く
const IDM_CONTEXT_OPEN_FOLDER: u16 = 4002;
/// コンテキストメニューID: フルパスをコピー
const IDM_CONTEXT_COPY_PATH: u16 = 4003;


// --- アプリケーションの状態管理 ---

/// 検索結果のファイル情報を格納する構造体
#[derive(Debug, Clone)]
pub struct FileResult {
    pub name: String,
    pub path: String,
    pub size: u64,
    pub modified_date: u64,
    pub highlighted_name: String,
    pub highlighted_path: String,
    pub is_folder: bool,
}

/// アプリケーションの状態をすべて保持する構造体
pub struct AppState {
    // --- UIハンドル ---
    pub main_hwnd: HWND,
    pub status_hwnd: HWND,
    pub edit_hwnd: HWND,
    pub listview_hwnd: HWND,
    pub re_button_hwnd: HWND,
    pub mi_button_hwnd: HWND,
    pub shell_context_toggle_hwnd: HWND,
    pub himagelist: HIMAGELIST,

    // --- DPI関連 ---
    pub current_dpi: u32,
    pub scale_factor: f32,

    // --- 検索オプション ---
    pub regex_enabled: bool,
    pub migemo_enabled: bool,
    pub shell_context_enabled: bool,

    // --- データ ---
    pub migemo_dict: Option<CompactDictionary>,
    pub search_results: Mutex<Vec<FileResult>>,

    // --- 仮想リストビュー関連 ---
    pub total_results: u32,
    pub current_search_term: String,
    pub page_size: usize,
    pub current_page_offset: usize, // 現在ロードされているページの開始オフセット

    // --- その他 ---
    // LVN_GETDISPINFOで使うための静的バッファ
    pub item_wide_buffer: [Vec<u16>; 4],
}

impl AppState {
    /// AppStateの新しいインスタンスを作成する
    pub fn new() -> Self {
        let migemo_dict = init_migemo_dict();
        Self {
            main_hwnd: HWND::default(),
            status_hwnd: HWND::default(),
            edit_hwnd: HWND::default(),
            listview_hwnd: HWND::default(),
            re_button_hwnd: HWND::default(),
            mi_button_hwnd: HWND::default(),
            shell_context_toggle_hwnd: HWND::default(),
            himagelist: HIMAGELIST::default(),
            current_dpi: 96,  // デフォルトDPI
            scale_factor: 1.0,  // デフォルトスケール
            regex_enabled: false,
            migemo_enabled: true,
            shell_context_enabled: false,
            migemo_dict,
            search_results: Mutex::new(Vec::new()),
            total_results: 0,
            current_search_term: String::new(),
            page_size: 100,  // 一度に読み込む件数（初回検索の件数と一致）
            current_page_offset: 0,
            item_wide_buffer: [Vec::new(), Vec::new(), Vec::new(), Vec::new()],
        }
    }
}

// --- main関数 ---

/// アプリケーションのエントリポイント
fn main() -> Result<()> {
    // COMライブラリの初期化
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
    }

    // DPI対応を有効にする
    unsafe {
        // プロセス全体でDPI認識を設定
        let _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    }

    // アプリケーションの状態を初期化
    let app_state = AppState::new();

    unsafe {
        let instance = GetModuleHandleA(None)?;
        let icon = LoadIconW(Some(instance.into()), PCWSTR(1 as _))?;

        // ウィンドウクラスの登録
        let wc = WNDCLASSW {
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            hInstance: instance.into(),
            lpszClassName: w!("window"),
            style: CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS,
            lpfnWndProc: Some(wndproc),
            hIcon: icon,
            ..Default::default()
        };

        let atom = RegisterClassW(&wc);
        debug_assert!(atom != 0);

        // メインウィンドウの作成
        // ここで Box<AppState> を作成し、WM_CREATE でウィンドウに渡す
        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("window"),
            w!("Migemo Everything"),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE | WS_CLIPCHILDREN,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            800, // 初期ウィンドウ幅
            600, // 初期ウィンドウ高さ
            None,
            None,
            Some(instance.into()),
            Some(Box::into_raw(Box::new(app_state)) as *const c_void), // AppStateを渡す
        )?;

        // アクセラレータテーブルの作成
        let accelerators = [
            ACCEL { fVirt: FCONTROL | FVIRTKEY, key: b'Q' as u16, cmd: IDA_EXIT },
            ACCEL { fVirt: FCONTROL | FVIRTKEY, key: b'R' as u16, cmd: IDA_REGEX },
            ACCEL { fVirt: FCONTROL | FSHIFT | FVIRTKEY, key: b'R' as u16, cmd: IDA_MIGEMO },
        ];
        let haccel = CreateAcceleratorTableW(&accelerators)?;

        // メッセージループ
        let mut message = MSG::default();
        while GetMessageW(&mut message, None, 0, 0).into() {
            if TranslateAcceleratorW(hwnd, haccel, &message) == 0 {
                let _ = TranslateMessage(&message);
                DispatchMessageW(&message);
            }
        }
    }

    // COMライブラリの解放
    unsafe { CoUninitialize() };
    Ok(())
}

// --- ウィンドウプロシージャ ---

/// メインウィンドウプロシージャ
/// 各メッセージを対応するハンドラ関数に振り分ける
pub extern "system" fn wndproc(
    window: HWND,
    message: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    // GWLP_USERDATAからAppStateのポインタを取得
    // WM_CREATEより前のメッセージではまだ設定されていないので注意
    let app_state_ptr =
        unsafe { GetWindowLongPtrW(window, GWLP_USERDATA) as *mut AppState };

    // ポインタがNULLでなければ、安全な参照に変換
    // WM_DESTROYでポインタは0に設定されるので、それ以降はNoneになる
    let state = if app_state_ptr.is_null() {
        None
    } else {
        Some(unsafe { &mut *app_state_ptr })
    };

    match message {
        WM_CREATE => handle_create(window, lparam),
        WM_DESTROY => handle_destroy(window),
        WM_COMMAND => handle_command(window, wparam, lparam, state.unwrap()),
        WM_TIMER => handle_timer(window, wparam, state.unwrap()),
        WM_NOTIFY => handle_notify(window, lparam, state.unwrap()),
        WM_SIZE => handle_size(window, lparam, state.unwrap()),
        WM_SETFOCUS => handle_setfocus(state.unwrap()),
        WM_DPICHANGED => handle_dpi_changed(window, wparam, lparam, state.unwrap()),
        WM_PAINT => {
            let _ = unsafe { ValidateRect(Some(window), None) };
            LRESULT(0)
        }
        _ => unsafe { DefWindowProcW(window, message, wparam, lparam) },
    }
}

// --- イベントハンドラ ---

/// WM_CREATE メッセージのハンドラ
fn handle_create(window: HWND, lparam: LPARAM) -> LRESULT {
    // CreateWindowExWから渡されたポインタを取得
    let create_struct = unsafe { &*(lparam.0 as *const CREATESTRUCTW) };
    let app_state_ptr = create_struct.lpCreateParams as *mut AppState;

    // ポインタをウィンドウのユーザーデータとして保存
    unsafe {
        SetWindowLongPtrW(window, GWLP_USERDATA, app_state_ptr as isize);
    }

    // AppStateへのミュータブルな参照を取得
    let state = unsafe { &mut *app_state_ptr };
    state.main_hwnd = window;

    // DPIを初期化
    unsafe {
        state.current_dpi = GetDpiForWindow(window);
        state.scale_factor = state.current_dpi as f32 / 96.0;
    }

    // UIコントロールの作成
    create_menu(window);
    create_controls(window, create_struct.hInstance, state);
    setup_listview(state);
    update_ui_states(state);

    LRESULT(0)
}

/// WM_DESTROY メッセージのハンドラ
fn handle_destroy(window: HWND) -> LRESULT {
    // ウィンドウのユーザーデータからポインタを取得
    let app_state_ptr =
        unsafe { GetWindowLongPtrW(window, GWLP_USERDATA) as *mut AppState };

    if !app_state_ptr.is_null() {
        // ポインタを0に設定して、ダングリングポインタを防ぐ
        unsafe {
            SetWindowLongPtrW(window, GWLP_USERDATA, 0);
            // Boxを再構築して、メモリを適切に解放する
            drop(Box::from_raw(app_state_ptr));
        }
    }
    unsafe { PostQuitMessage(0) };
    LRESULT(0)
}

/// WM_COMMAND メッセージのハンドラ (メニュー、ボタンクリック)
fn handle_command(window: HWND, wparam: WPARAM, lparam: LPARAM, state: &mut AppState) -> LRESULT {
    let control_id = loword(wparam.0 as u32);
    let notification_code = hiword(wparam.0 as u32);

    match control_id {
        // --- アクセラレータ ---
        IDA_EXIT => { let _ = unsafe { DestroyWindow(window) }; }
        IDA_REGEX => {
            state.regex_enabled = !state.regex_enabled;
            if state.regex_enabled { state.migemo_enabled = false; }
            update_ui_states(state);
            trigger_search(window);
        }
        IDA_MIGEMO => {
            state.migemo_enabled = !state.migemo_enabled;
            if state.migemo_enabled { state.regex_enabled = false; }
            update_ui_states(state);
            trigger_search(window);
        }
        // --- メニュー項目 ---
        IDM_FILE_EXIT => { let _ = unsafe { DestroyWindow(window) }; }
        IDM_SEARCH_REGEX => {
            state.regex_enabled = !state.regex_enabled;
            if state.regex_enabled { state.migemo_enabled = false; }
            update_ui_states(state);
            trigger_search(window);
        }
        IDM_SEARCH_MIGEMO => {
            state.migemo_enabled = !state.migemo_enabled;
            if state.migemo_enabled { state.regex_enabled = false; }
            update_ui_states(state);
            trigger_search(window);
        }
        // --- ボタン ---
        RE_BUTTON_ID => {
            state.regex_enabled = !state.regex_enabled;
            if state.regex_enabled { state.migemo_enabled = false; }
            update_ui_states(state);
            trigger_search(window);
        }
        MI_BUTTON_ID => {
            state.migemo_enabled = !state.migemo_enabled;
            if state.migemo_enabled { state.regex_enabled = false; }
            update_ui_states(state);
            trigger_search(window);
        }
        SHELL_CONTEXT_TOGGLE_ID => {
            let is_checked = unsafe { SendMessageW(state.shell_context_toggle_hwnd, BM_GETCHECK, None, None) } == LRESULT(BST_CHECKED.0 as isize);
            state.shell_context_enabled = is_checked;
        }
        // --- エディットボックス ---
        EDIT_ID if notification_code as u32 == EN_CHANGE => {
            // 500ミリ秒後に検索タイマーをセット
            unsafe { SetTimer(Some(window), TIMER_ID, 500, None) };
        }
        // --- コンテキストメニュー ---
        IDM_CONTEXT_OPEN => {
            let item_index = lparam.0 as usize;
            ensure_data_available(state, item_index);
            let results = state.search_results.lock().unwrap();
            let local_index = item_index - state.current_page_offset;
            if let Some(result) = results.get(local_index) {
                let full_path = Path::new(&result.path).join(&result.name);
                let path_w = str_to_wide(full_path.to_str().unwrap_or(""));
                thread::spawn(move || unsafe {
                    ShellExecuteW(None, w!("open"), PCWSTR(path_w.as_ptr()), None, None, SW_SHOW);
                });
            }
        }
        IDM_CONTEXT_OPEN_FOLDER => {
            let item_index = lparam.0 as usize;
            ensure_data_available(state, item_index);
            let results = state.search_results.lock().unwrap();
            let local_index = item_index - state.current_page_offset;
            if let Some(result) = results.get(local_index) {
                let full_path = Path::new(&result.path).join(&result.name);
                let params = format!("/select,\"{}\"", full_path.display());
                let params_w = str_to_wide(&params);
                thread::spawn(move || unsafe {
                    ShellExecuteW(None, w!("open"), w!("explorer.exe"), PCWSTR(params_w.as_ptr()), None, SW_SHOW);
                });
            }
        }
        IDM_CONTEXT_COPY_PATH => {
            let item_index = lparam.0 as usize;
            ensure_data_available(state, item_index);
            let results = state.search_results.lock().unwrap();
            let local_index = item_index - state.current_page_offset;
            if let Some(result) = results.get(local_index) {
                let full_path_str = Path::new(&result.path).join(&result.name).to_str().unwrap_or("").to_string();
                copy_text_to_clipboard(window, &full_path_str);
            }
        }
        _ => {}
    }
    LRESULT(0)
}

/// WM_TIMER メッセージのハンドラ
fn handle_timer(window: HWND, wparam: WPARAM, state: &mut AppState) -> LRESULT {
    if wparam.0 == TIMER_ID {
        let _ = unsafe { KillTimer(Some(window), TIMER_ID) };
        perform_search(state);
    }
    LRESULT(0)
}

/// WM_NOTIFY メッセージのハンドラ (主にListViewからの通知)
fn handle_notify(window: HWND, lparam: LPARAM, state: &mut AppState) -> LRESULT {
    let nmhdr = unsafe { &*(lparam.0 as *const NMHDR) };

    if nmhdr.hwndFrom == state.listview_hwnd {
        match nmhdr.code {
            LVN_GETDISPINFOW => handle_get_disp_info(lparam, state),
            NM_CUSTOMDRAW => return handle_custom_draw(lparam, state),
            NM_RCLICK => handle_right_click(window, lparam, state),
            NM_DBLCLK => {
                let item_activate = unsafe { &*(lparam.0 as *const NMITEMACTIVATE) };
                if item_activate.iItem != -1 {
                    unsafe {
                        SendMessageW(window, WM_COMMAND, Some(WPARAM(IDM_CONTEXT_OPEN as usize)), Some(LPARAM(item_activate.iItem as isize)));
                    }
                }
            }
            _ => {}
        }
    }
    LRESULT(0)
}

/// WM_SIZE メッセージのハンドラ
fn handle_size(_window: HWND, lparam: LPARAM, state: &AppState) -> LRESULT {
    let width = loword(lparam.0 as u32) as i32;
    let height = hiword(lparam.0 as u32) as i32;
    layout_controls(width, height, state);
    LRESULT(0)
}

/// WM_SETFOCUS メッセージのハンドラ
fn handle_setfocus(state: &AppState) -> LRESULT {
    let _ = unsafe { SetFocus(Some(state.edit_hwnd)) };
    LRESULT(0)
}

/// WM_DPICHANGED メッセージのハンドラ
fn handle_dpi_changed(window: HWND, wparam: WPARAM, lparam: LPARAM, state: &mut AppState) -> LRESULT {
    let new_dpi = hiword(wparam.0 as u32) as u32;
    state.current_dpi = new_dpi;
    state.scale_factor = new_dpi as f32 / 96.0;
    
    let suggested_rect = unsafe { &*(lparam.0 as *const RECT) };
    
    unsafe {
        let _ = SetWindowPos(
            window, None,
            suggested_rect.left, suggested_rect.top,
            suggested_rect.right - suggested_rect.left,
            suggested_rect.bottom - suggested_rect.top,
            SWP_NOZORDER | SWP_NOACTIVATE,
        );
        let _ = InvalidateRect(Some(window), None, true);
    }
    
    LRESULT(0)
}

// --- イベントハンドラ (WM_NOTIFY) のためのヘルパー関数 ---

fn handle_get_disp_info(lparam: LPARAM, state: &mut AppState) {
    let dispinfo = unsafe { &mut *(lparam.0 as *mut NMLVDISPINFOW) };
    let item = &mut dispinfo.item;
    let item_index = item.iItem as usize;

    if item_index >= state.total_results as usize { return; }

    ensure_data_available(state, item_index);

    let results = state.search_results.lock().unwrap();
    let local_index = item_index - state.current_page_offset;
    
    if let Some(result) = results.get(local_index) {
        if (item.mask & LVIF_TEXT) == LVIF_TEXT {
            let sub_item_index = item.iSubItem as usize;
            let text = match sub_item_index {
                0 => if !result.highlighted_name.is_empty() { parse_highlight_text(&result.highlighted_name).0 } else { result.name.clone() },
                1 => if !result.highlighted_path.is_empty() { parse_highlight_text(&result.highlighted_path).0 } else { result.path.clone() },
                2 => format_size(result.size),
                3 => format_date(result.modified_date),
                _ => String::new(),
            };
            state.item_wide_buffer[sub_item_index] = str_to_wide(&text);
            item.pszText = PWSTR(state.item_wide_buffer[sub_item_index].as_mut_ptr());
        }
        if item.iSubItem == 0 && (item.mask & LVIF_IMAGE) == LVIF_IMAGE && !result.name.is_empty() {
            item.iImage = get_icon_index(&result.name, result.is_folder, state.himagelist);
        }
    }
}

fn handle_custom_draw(lparam: LPARAM, state: &mut AppState) -> LRESULT {
    let custom_draw = unsafe { &mut *(lparam.0 as *mut NMLVCUSTOMDRAW) };

    match custom_draw.nmcd.dwDrawStage {
        CDDS_PREPAINT => LRESULT(CDRF_NOTIFYITEMDRAW as isize),
        CDDS_ITEMPREPAINT => LRESULT(CDRF_NOTIFYSUBITEMDRAW as isize),
        stage if stage.0 == (CDDS_SUBITEM.0 | CDDS_ITEMPREPAINT.0) => {
            let item_index = custom_draw.nmcd.dwItemSpec as usize;
            let sub_item_index = custom_draw.iSubItem as usize;

            if item_index >= state.total_results as usize { return LRESULT(CDRF_DODEFAULT as isize); }

            ensure_data_available(state, item_index);

            let results = state.search_results.lock().unwrap();
            let local_index = item_index - state.current_page_offset;
            
            if let Some(result) = results.get(local_index) {
                let (text_to_draw, highlight_ranges) = match sub_item_index {
                    0 if !result.highlighted_name.is_empty() => parse_highlight_text(&result.highlighted_name),
                    1 if !result.highlighted_path.is_empty() => parse_highlight_text(&result.highlighted_path),
                    _ => return LRESULT(CDRF_DODEFAULT as isize),
                };

                if highlight_ranges.is_empty() { return LRESULT(CDRF_DODEFAULT as isize); }

                let hdc = custom_draw.nmcd.hdc;
                let mut rect = custom_draw.nmcd.rc;
                let is_selected = (custom_draw.nmcd.uItemState & CDIS_SELECTED).0 != 0;

                let bg_color = if is_selected { unsafe { GetSysColor(COLOR_HIGHLIGHT) } } else { unsafe { GetSysColor(COLOR_WINDOW) } };
                let bg_brush = unsafe { CreateSolidBrush(COLORREF(bg_color)) };
                unsafe { FillRect(hdc, &rect, bg_brush) };
                let _ = unsafe { DeleteObject(bg_brush.into()) };

                if sub_item_index == 0 && !result.name.is_empty() {
                    let icon_index = get_icon_index(&result.name, result.is_folder, state.himagelist);
                    if state.himagelist.0 != 0 && icon_index >= 0 {
                        let icon_size = (16.0 * state.scale_factor) as i32;
                        let icon_padding = (2.0 * state.scale_factor) as i32;
                        let icon_y = rect.top + (rect.bottom - rect.top - icon_size) / 2;
                        
                        if icon_y >= rect.top && icon_y + icon_size <= rect.bottom {
                            let _ = unsafe { ImageList_Draw(state.himagelist, icon_index, hdc, rect.left + icon_padding, icon_y, ILD_TRANSPARENT) };
                        }
                    }
                    rect.left += (22.0 * state.scale_factor) as i32;
                } else {
                    rect.left += (4.0 * state.scale_factor) as i32;
                }
                rect.right -= (4.0 * state.scale_factor) as i32;

                let text_color = if is_selected { unsafe { GetSysColor(COLOR_HIGHLIGHTTEXT) } } else { unsafe { GetSysColor(COLOR_WINDOWTEXT) } };
                unsafe {
                    SetBkMode(hdc, TRANSPARENT);
                    SetTextColor(hdc, COLORREF(text_color));
                }

                let mut x = rect.left;
                let font_offset = (8.0 * state.scale_factor) as i32;
                let y = rect.top + (rect.bottom - rect.top) / 2 - font_offset;
                let chars: Vec<char> = text_to_draw.chars().collect();
                
                let full_text_wide = str_to_wide(&text_to_draw);
                let mut char_widths = vec![0i32; chars.len()];
                let mut max_fit_chars = chars.len();
                
                if chars.len() > 0 && full_text_wide.len() > 1 {
                    let mut fit_count = 0i32;
                    let mut size = SIZE::default();
                    let _ = unsafe {
                        GetTextExtentExPointW(hdc, PCWSTR(full_text_wide.as_ptr()), (full_text_wide.len() - 1) as i32, rect.right - rect.left, Some(&mut fit_count), Some(char_widths.as_mut_ptr()), &mut size)
                    };
                    max_fit_chars = fit_count as usize;
                }
                
                let ellipsis = "...";
                let ellipsis_wide = str_to_wide(ellipsis);
                let ellipsis_width = unsafe { let mut size = SIZE::default(); let _ = GetTextExtentPointW(hdc, &ellipsis_wide, &mut size); size.cx };
                
                let is_truncated = chars.len() > max_fit_chars;
                
                let effective_max_chars = if is_truncated {
                    let available_width_for_text = rect.right - rect.left - ellipsis_width;
                    if available_width_for_text > 0 {
                        let mut truncated_fit_count = 0i32;
                        let mut size = SIZE::default();
                        let _ = unsafe { GetTextExtentExPointW(hdc, PCWSTR(full_text_wide.as_ptr()), (full_text_wide.len() - 1) as i32, available_width_for_text, Some(&mut truncated_fit_count), None, &mut size) };
                        std::cmp::min(truncated_fit_count as usize, max_fit_chars)
                    } else { 0 }
                } else { max_fit_chars };
                
                let mut current_pos = 0;
                let mut last_drawn_pos = 0;
                
                while current_pos < chars.len() && current_pos < effective_max_chars {
                    let is_current_highlighted = highlight_ranges.iter().any(|(start, end)| current_pos >= *start && current_pos < *end);
                    let mut end_pos = current_pos + 1;
                    
                    while end_pos < chars.len() && end_pos <= effective_max_chars {
                        let is_next_highlighted = highlight_ranges.iter().any(|(start, end)| end_pos >= *start && end_pos < *end);
                        if is_current_highlighted == is_next_highlighted { end_pos += 1; } else { break; }
                    }
                    
                    end_pos = std::cmp::min(end_pos, effective_max_chars);
                    let text_segment: String = chars[current_pos..end_pos].iter().collect();
                    let text_wide = str_to_wide(&text_segment);
                    
                    let start_x = if current_pos == 0 { 0 } else { char_widths[current_pos - 1] };
                    let end_x = if end_pos > 0 && end_pos <= char_widths.len() { char_widths[end_pos - 1] } else { 0 };
                    let segment_width = end_x - start_x;
                    let available_space = rect.right - x;
                    let actual_segment_width = std::cmp::min(segment_width, available_space);
                    
                    if actual_segment_width <= 0 || x >= rect.right { break; }
                    
                    if is_current_highlighted && !is_selected {
                        let highlight_left = x;
                        let highlight_right = std::cmp::min(x + segment_width, rect.right);
                        
                        if highlight_right > highlight_left && highlight_left < rect.right {
                            let highlight_brush = unsafe { CreateSolidBrush(COLORREF(0x00FFFF)) };
                            let highlight_rect = RECT { left: highlight_left, top: rect.top, right: highlight_right, bottom: rect.bottom };
                            unsafe { FillRect(hdc, &highlight_rect, highlight_brush) };
                            let _ = unsafe { DeleteObject(highlight_brush.into()) };
                        }
                    }
                    
                    unsafe {
                        let clip_region = CreateRectRgn(rect.left, rect.top, rect.right, rect.bottom);
                        SelectClipRgn(hdc, Some(clip_region));
                        let _ = TextOutW(hdc, x, y, &text_wide);
                        SelectClipRgn(hdc, None);
                        let _ = DeleteObject(clip_region.into());
                    }
                    
                    x += actual_segment_width;
                    current_pos = end_pos;
                    last_drawn_pos = end_pos;
                    if x >= rect.right { break; }
                }
                
                if is_truncated && last_drawn_pos < chars.len() && x + ellipsis_width <= rect.right {
                    unsafe {
                        let clip_region = CreateRectRgn(rect.left, rect.top, rect.right, rect.bottom);
                        SelectClipRgn(hdc, Some(clip_region));
                        let _ = TextOutW(hdc, x, y, &ellipsis_wide);
                        SelectClipRgn(hdc, None);
                        let _ = DeleteObject(clip_region.into());
                    }
                }
                return LRESULT(CDRF_SKIPDEFAULT as isize);
            }
            LRESULT(CDRF_DODEFAULT as isize)
        }
        _ => LRESULT(CDRF_DODEFAULT as isize),
    }
}

fn handle_right_click(window: HWND, lparam: LPARAM, state: &mut AppState) {
    let item_activate = unsafe { &*(lparam.0 as *const NMITEMACTIVATE) };
    let item_index = item_activate.iItem;

    if item_index == -1 { return; }

    // デッドロックを避けるため、メニュー表示の前にファイルパスを取得し、Mutexロックを解放する
    let maybe_full_path: Option<PathBuf> = {
        ensure_data_available(state, item_index as usize);
        let results = state.search_results.lock().unwrap();
        let local_index = item_index as usize - state.current_page_offset;
        results.get(local_index).map(|result| {
            Path::new(&result.path).join(&result.name)
        })
    };

    // 有効なパスが取得できた場合のみ続行
    if let Some(full_path) = maybe_full_path {
        if state.shell_context_enabled {
            // --- Shell Context Menu Logic ---
            show_shell_context_menu(window, state.listview_hwnd, &full_path, item_activate.ptAction);
        } else {
            // --- Original Custom Menu Logic ---
            unsafe {
                let h_popup_menu = CreatePopupMenu().unwrap();
                let _ = AppendMenuW(h_popup_menu, MF_STRING, IDM_CONTEXT_OPEN as usize, w!("開く(&O)"));
                let _ = AppendMenuW(h_popup_menu, MF_STRING, IDM_CONTEXT_OPEN_FOLDER as usize, w!("フォルダを開く(&F)"));
                let _ = AppendMenuW(h_popup_menu, MF_STRING, IDM_CONTEXT_COPY_PATH as usize, w!("フルパスをコピー(&C)"));
                let _ = SetMenuDefaultItem(h_popup_menu, IDM_CONTEXT_OPEN as u32, 0);

                let mut pt = item_activate.ptAction;
                let _ = ClientToScreen(state.listview_hwnd, &mut pt);

                // コンパイルエラーを修正: 5番目の引数は Option<i32> 型であるため Some(0) を渡す
                let cmd = TrackPopupMenu(h_popup_menu, TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, Some(0), window, None);

                if cmd.as_bool() {
                    SendMessageW(window, WM_COMMAND, Some(WPARAM(cmd.0 as usize)), Some(LPARAM(item_index as isize)));
                }
                let _ = DestroyMenu(h_popup_menu);
            }
        }
    }
}


// --- UI関連の関数 ---

/// メニューを作成してウィンドウに設定する
fn create_menu(window: HWND) {
    unsafe {
        let h_menu = CreateMenu().unwrap();
        let h_file_submenu = CreatePopupMenu().unwrap();
        let _ = AppendMenuW(h_file_submenu, MF_STRING, IDM_FILE_EXIT as usize, w!("終了(&E)\tCtrl+Q"));
        let _ = AppendMenuW(h_menu, MF_POPUP, h_file_submenu.0 as usize, w!("ファイル(&F)"));

        let h_search_submenu = CreatePopupMenu().unwrap();
        let _ = AppendMenuW(h_search_submenu, MF_STRING, IDM_SEARCH_REGEX as usize, w!("正規表現で検索\tCtrl+R"));
        let _ = AppendMenuW(h_search_submenu, MF_STRING, IDM_SEARCH_MIGEMO as usize, w!("Migemoで検索\tCtrl+Shift+R"));
        let _ = AppendMenuW(h_menu, MF_POPUP, h_search_submenu.0 as usize, w!("検索(&S)"));
        let _ = SetMenu(window, Some(h_menu));
    }
}

/// すべてのUIコントロールを作成する（DPI対応）
fn create_controls(window: HWND, instance: HINSTANCE, state: &mut AppState) {
    let scale = state.scale_factor;
    let font_height = (-12.0 * scale) as i32;
    
    let h_font = unsafe {
        CreateFontW(font_height, 0, 0, 0, FW_NORMAL.0 as i32, 0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, (FF_DONTCARE.0 | VARIABLE_PITCH.0) as u32, w!("Segoe UI"))
    };

    unsafe {
        state.status_hwnd = CreateWindowExW(WINDOW_EX_STYLE::default(), w!("STATIC"), w!("Ready"), WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, Some(window), None, Some(instance), None).unwrap();
        state.edit_hwnd = CreateWindowExW(WS_EX_CLIENTEDGE, w!("EDIT"), w!(""), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_LEFT as u32 | ES_AUTOHSCROLL as u32), 0, 0, 0, 0, Some(window), Some(HMENU(EDIT_ID as isize as *mut c_void)), Some(instance), None).unwrap();
        state.re_button_hwnd = CreateWindowExW(WINDOW_EX_STYLE::default(), w!("BUTTON"), w!("RE"), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_PUSHBUTTON as u32), 0, 0, 0, 0, Some(window), Some(HMENU(RE_BUTTON_ID as isize as *mut c_void)), Some(instance), None).unwrap();
        state.mi_button_hwnd = CreateWindowExW(WINDOW_EX_STYLE::default(), w!("BUTTON"), w!("Mi"), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_PUSHBUTTON as u32), 0, 0, 0, 0, Some(window), Some(HMENU(MI_BUTTON_ID as isize as *mut c_void)), Some(instance), None).unwrap();
        state.listview_hwnd = CreateWindowExW(WINDOW_EX_STYLE::default(), w!("SysListView32"), w!(""), WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | WS_VSCROLL | WS_HSCROLL | WINDOW_STYLE(LVS_REPORT as u32 | LVS_OWNERDATA as u32), 0, 0, 0, 0, Some(window), None, Some(instance), None).unwrap();
        
        // シェルコンテキストメニュー切り替えボタン
        state.shell_context_toggle_hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(), w!("BUTTON"), w!("Shell Menu"),
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
            0, 0, 0, 0, Some(window), Some(HMENU(SHELL_CONTEXT_TOGGLE_ID as isize as *mut c_void)), Some(instance), None,
        ).unwrap();


        if !h_font.is_invalid() {
            SendMessageW(state.status_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
            SendMessageW(state.edit_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
            SendMessageW(state.re_button_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
            SendMessageW(state.mi_button_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
            SendMessageW(state.listview_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
            SendMessageW(state.shell_context_toggle_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
        }
    }
}

/// リストビューの初期設定（カラム、拡張スタイル、イメージリスト）（DPI対応）
fn setup_listview(state: &mut AppState) {
    unsafe {
        let ex_style = LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES;
        SendMessageW(state.listview_hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, Some(WPARAM(ex_style as usize)), Some(LPARAM(ex_style as isize)));

        let mut shfi: SHFILEINFOW = std::mem::zeroed();
        state.himagelist = HIMAGELIST(SHGetFileInfoW(w!(""), FILE_ATTRIBUTE_NORMAL, Some(&mut shfi as *mut _), std::mem::size_of::<SHFILEINFOW>() as u32, SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON) as isize);

        if state.himagelist.0 != 0 {
            SendMessageW(state.listview_hwnd, LVM_SETIMAGELIST, Some(WPARAM(LVSIL_SMALL as usize)), Some(LPARAM(state.himagelist.0)));
        }

        let scale = state.scale_factor;
        let columns = [
            (w!("名前"), (300.0 * scale) as i32),
            (w!("フォルダ"), (300.0 * scale) as i32),
            (w!("サイズ"), (80.0 * scale) as i32),
            (w!("更新日時"), (150.0 * scale) as i32)
        ];
        
        for (i, (text, width)) in columns.iter().enumerate() {
            let mut col = LVCOLUMNW {
                mask: LVCF_TEXT | LVCF_WIDTH, cx: *width, pszText: PWSTR(text.as_ptr() as *mut _), ..Default::default()
            };
            if i == 2 { col.mask |= LVCF_FMT; col.fmt = LVCFMT_RIGHT; }
            SendMessageW(state.listview_hwnd, LVM_INSERTCOLUMNW, Some(WPARAM(i)), Some(LPARAM(&mut col as *mut _ as isize)));
        }
    }
}

/// 状態に基づいてUI（メニューのチェック、ボタンのスタイル）を更新する
fn update_ui_states(state: &AppState) {
    unsafe {
        let h_menu = GetMenu(state.main_hwnd);
        if h_menu.0 != std::ptr::null_mut() {
            let re_flag = if state.regex_enabled { MF_CHECKED } else { MF_UNCHECKED };
            let _ = CheckMenuItem(h_menu, IDM_SEARCH_REGEX as u32, re_flag.0);
            let mi_flag = if state.migemo_enabled { MF_CHECKED } else { MF_UNCHECKED };
            let _ = CheckMenuItem(h_menu, IDM_SEARCH_MIGEMO as u32, mi_flag.0);
        }

        let re_style = if state.regex_enabled { BS_DEFPUSHBUTTON } else { BS_PUSHBUTTON };
        SetWindowLongW(state.re_button_hwnd, GWL_STYLE, (GetWindowLongW(state.re_button_hwnd, GWL_STYLE) & !(BS_DEFPUSHBUTTON as i32)) | re_style as i32);
        let _ = InvalidateRect(Some(state.re_button_hwnd), None, true);

        let mi_style = if state.migemo_enabled { BS_DEFPUSHBUTTON } else { BS_PUSHBUTTON };
        SetWindowLongW(state.mi_button_hwnd, GWL_STYLE, (GetWindowLongW(state.mi_button_hwnd, GWL_STYLE) & !(BS_DEFPUSHBUTTON as i32)) | mi_style as i32);
        let _ = InvalidateRect(Some(state.mi_button_hwnd), None, true);
    }
}

/// ウィンドウリサイズ時にコントロールを再配置する（DPI対応）
fn layout_controls(width: i32, height: i32, state: &AppState) {
    let scale = state.scale_factor;
    let bar_height = (25.0 * scale) as i32;
    let status_bar_height = (20.0 * scale) as i32;
    let button_width = (40.0 * scale) as i32;
    let toggle_button_width = (100.0 * scale) as i32;
    let total_button_width = button_width * 2;
    let list_y = bar_height;
    let toggle_x = width - toggle_button_width;

    unsafe {
        let _ = MoveWindow(state.edit_hwnd, 0, 0, width - total_button_width, bar_height, true);
        let _ = MoveWindow(state.re_button_hwnd, width - total_button_width, 0, button_width, bar_height, true);
        let _ = MoveWindow(state.mi_button_hwnd, width - button_width, 0, button_width, bar_height, true);
        let _ = MoveWindow(state.listview_hwnd, 0, list_y, width, height - list_y - status_bar_height, true);
        let _ = MoveWindow(state.status_hwnd, 0, height - status_bar_height, toggle_x, status_bar_height, true);
        let _ = MoveWindow(state.shell_context_toggle_hwnd, toggle_x, height - status_bar_height, toggle_button_width, status_bar_height, true);
    }
}

// --- 検索関連の関数 ---

/// Migemo辞書を初期化する
fn init_migemo_dict() -> Option<CompactDictionary> {
    use std::env;
    let paths = [
        PathBuf::from("migemo-compact-dict"),
        env::current_exe().ok().and_then(|p| p.parent().map(|d| d.join("migemo-compact-dict"))).unwrap_or_default(),
    ];

    for path in paths.iter() {
        if let Ok(mut f) = File::open(path) {
            let mut buf = Vec::new();
            if f.read_to_end(&mut buf).is_ok() {
                return Some(CompactDictionary::new(&buf));
            }
        }
    }
    None
}

/// Migemo検索を実行する
fn migemo_query(text: &str, dict: &Option<CompactDictionary>) -> Option<String> {
    dict.as_ref().map(|d| query(text.to_string(), d, &RegexOperator::Default))
}

/// 検索を即座に実行するためのタイマーをセットする
fn trigger_search(window: HWND) {
    unsafe { SetTimer(Some(window), TIMER_ID, 100, None) };
}

/// Everythingを使用して検索を実行し、結果を更新する
fn perform_search(state: &mut AppState) {
    let mut buffer: [u16; 512] = [0; 512];
    let len = unsafe { GetWindowTextW(state.edit_hwnd, &mut buffer) };
    let search_term = String::from_utf16_lossy(&buffer[..len as usize]);

    let window_title = if search_term.is_empty() { "Migemo Everything".to_string() } else { format!("{} - Migemo Everything", search_term) };
    unsafe {
        let title_wide = str_to_wide(&window_title);
        let _ = SetWindowTextW(state.main_hwnd, PCWSTR(title_wide.as_ptr()));
    }

    if search_term.is_empty() {
        state.search_results.lock().unwrap().clear();
        state.total_results = 0;
        state.current_search_term.clear();
        state.current_page_offset = 0;
        unsafe {
            let _ = SetWindowTextW(state.status_hwnd, w!("Ready"));
            SendMessageW(state.listview_hwnd, LVM_SETITEMCOUNT, Some(WPARAM(0)), Some(LPARAM(0)));
            let _ = InvalidateRect(Some(state.listview_hwnd), None, true);
        }
        return;
    }

    let final_search_term = if state.migemo_enabled { migemo_query(&search_term, &state.migemo_dict).unwrap_or(search_term) } else { search_term };

    if state.current_search_term != final_search_term {
        state.search_results.lock().unwrap().clear();
        state.current_search_term = final_search_term.clone();
        state.current_page_offset = 0;
    }

    let mut guard = global().lock().unwrap();
    let mut searcher = guard.searcher();
    
    searcher.set_search(&final_search_term);
    searcher.set_regex(state.regex_enabled || state.migemo_enabled);
    searcher.set_request_flags(
        RequestFlags::EVERYTHING_REQUEST_FILE_NAME | RequestFlags::EVERYTHING_REQUEST_PATH |
        RequestFlags::EVERYTHING_REQUEST_SIZE | RequestFlags::EVERYTHING_REQUEST_DATE_MODIFIED |
        RequestFlags::EVERYTHING_REQUEST_ATTRIBUTES |
        RequestFlags::EVERYTHING_REQUEST_HIGHLIGHTED_FILE_NAME |
        RequestFlags::EVERYTHING_REQUEST_HIGHLIGHTED_PATH
    );

    let query_results = searcher.set_max(100).query();
    state.total_results = query_results.total();

    let mut initial_results = Vec::new();
    for item in query_results.iter() {
        initial_results.push(FileResult {
            name: item.filename().unwrap_or_default().to_string_lossy().to_string(),
            path: item.path().unwrap_or_default().to_string_lossy().to_string(),
            size: item.size().unwrap_or(0),
            modified_date: item.date_modified().unwrap_or(0),
            highlighted_name: item.highlighted_filename().unwrap_or_default().to_string_lossy().to_string(),
            highlighted_path: item.highlighted_path().unwrap_or_default().to_string_lossy().to_string(),
            is_folder: item.is_folder(),
        });
    }

    state.current_page_offset = 0;
    *state.search_results.lock().unwrap() = initial_results;

    let status_text = format!("{} items found", state.total_results);
    unsafe {
        let _ = SetWindowTextW(state.status_hwnd, PCWSTR(str_to_wide(&status_text).as_ptr()));
        SendMessageW(state.listview_hwnd, LVM_SETITEMCOUNT, Some(WPARAM(state.total_results as usize)), Some(LPARAM(0)));
        let _ = InvalidateRect(Some(state.listview_hwnd), None, true);
    }
}

/// 指定されたアイテムインデックスのデータが利用可能かを確認し、必要に応じて読み込む
fn ensure_data_available(state: &mut AppState, item_index: usize) {
    if state.current_search_term.is_empty() { return; }
    
    let page_start = (item_index / state.page_size) * state.page_size;
    
    if state.current_page_offset == page_start {
        let results = state.search_results.lock().unwrap();
        let local_index = item_index - page_start;
        if local_index < results.len() { return; }
    }
    
    load_page(state, page_start);
}

/// 指定されたオフセットからページサイズ分のデータを読み込む
fn load_page(state: &mut AppState, offset: usize) {
    if state.current_search_term.is_empty() { return; }
    
    let mut guard = global().lock().unwrap();
    let mut searcher = guard.searcher();
    
    searcher.set_search(&state.current_search_term);
    searcher.set_regex(state.regex_enabled || state.migemo_enabled);
    searcher.set_offset(offset as u32);
    searcher.set_max(state.page_size as u32);
    searcher.set_request_flags(
        RequestFlags::EVERYTHING_REQUEST_FILE_NAME | RequestFlags::EVERYTHING_REQUEST_PATH |
        RequestFlags::EVERYTHING_REQUEST_SIZE | RequestFlags::EVERYTHING_REQUEST_DATE_MODIFIED |
        RequestFlags::EVERYTHING_REQUEST_ATTRIBUTES |
        RequestFlags::EVERYTHING_REQUEST_HIGHLIGHTED_FILE_NAME |
        RequestFlags::EVERYTHING_REQUEST_HIGHLIGHTED_PATH
    );

    let query_results = searcher.query();
    let mut new_results = Vec::new();
    
    for item in query_results.iter() {
        new_results.push(FileResult {
            name: item.filename().unwrap_or_default().to_string_lossy().to_string(),
            path: item.path().unwrap_or_default().to_string_lossy().to_string(),
            size: item.size().unwrap_or(0),
            modified_date: item.date_modified().unwrap_or(0),
            highlighted_name: item.highlighted_filename().unwrap_or_default().to_string_lossy().to_string(),
            highlighted_path: item.highlighted_path().unwrap_or_default().to_string_lossy().to_string(),
            is_folder: item.is_folder(),
        });
    }
    
    state.current_page_offset = offset;
    *state.search_results.lock().unwrap() = new_results;
}

// --- シェルコンテキストメニュー関連 ---

/// シェルのコンテキストメニューを表示する
fn show_shell_context_menu(owner: HWND, listview_hwnd: HWND, full_path: &Path, point: POINT) {
    if let Ok((shell_folder, _pidl_absolute, pidl_relative)) = get_shell_folder_and_pidl(full_path) {
        let context_menu: Result<IContextMenu> = unsafe { shell_folder.GetUIObjectOf(owner, &[pidl_relative], None) };

        if let Ok(context_menu) = context_menu {
            if let Ok(hmenu) = unsafe { CreatePopupMenu() } {
                if unsafe { context_menu.QueryContextMenu(hmenu, 0, 1, 0x7FFF, CMF_NORMAL) }.is_ok() {
                    let mut pt = point;
                    let _ = unsafe { ClientToScreen(listview_hwnd, &mut pt) };

                    // コンパイルエラーを修正: 5番目の引数は Option<i32> 型であるため Some(0) を渡す
                    let cmd = unsafe { TrackPopupMenu(hmenu, TPM_RETURNCMD, pt.x, pt.y, Some(0), owner, None) };
                    let cmd_u32 = cmd.0 as u32;

                    if cmd_u32 > 0 && cmd_u32 <= 0x7FFF {
                        let ci = CMINVOKECOMMANDINFO {
                            cbSize: std::mem::size_of::<CMINVOKECOMMANDINFO>() as u32,
                            hwnd: owner,
                            lpVerb: PCSTR::from_raw((cmd_u32 - 1) as *const u8),
                            nShow: SW_SHOWNORMAL.0 as i32,
                            ..Default::default()
                        };

                        if let Err(e) = unsafe { context_menu.InvokeCommand(&ci) } {
                            eprintln!("InvokeCommand failed: {:?}", e);
                        }
                    }
                }
                let _ = unsafe { DestroyMenu(hmenu) };
            }
        }
    }
}

/// ファイルパスからIShellFolderと相対PIDLを取得する
fn get_shell_folder_and_pidl(path: &Path) -> Result<(IShellFolder, OwningPidl, *const ITEMIDLIST)> {
    let path_wide: Vec<u16> = path.as_os_str().encode_wide().chain(Some(0)).collect();
    let mut pidl_absolute = OwningPidl::new();

    unsafe {
        let sfgao: u32 = SFGAO_FILESYSTEM.0;
        SHParseDisplayName(PCWSTR(path_wide.as_ptr()), None, pidl_absolute.as_mut_ptr(), sfgao, None)?;
    }

    let mut pidl_relative_ptr: *mut ITEMIDLIST = std::ptr::null_mut();
    let shell_folder: IShellFolder = unsafe { SHBindToParent(pidl_absolute.as_ptr(), Some(&mut pidl_relative_ptr))? };

    Ok((shell_folder, pidl_absolute, pidl_relative_ptr))
}

/// PIDLのメモリ解放を管理するラッパー構造体
struct OwningPidl {
    ptr: *mut ITEMIDLIST,
}

impl OwningPidl {
    fn new() -> Self { Self { ptr: std::ptr::null_mut() } }
    fn as_ptr(&self) -> *const ITEMIDLIST { self.ptr }
    fn as_mut_ptr(&mut self) -> *mut *mut ITEMIDLIST { &mut self.ptr }
}

impl Drop for OwningPidl {
    fn drop(&mut self) {
        if !self.ptr.is_null() {
            unsafe { CoTaskMemFree(Some(self.ptr as *const _)) };
        }
    }
}

// --- ユーティリティ関数 ---

/// Win32のHIWORDマクロ相当
fn hiword(val: u32) -> u16 { ((val >> 16) & 0xFFFF) as u16 }
/// Win32のLOWORDマクロ相当
fn loword(val: u32) -> u16 { (val & 0xFFFF) as u16 }

/// 文字列をUTF-16のワイド文字列（終端ナル文字付き）に変換する
fn str_to_wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

fn parse_highlight_text(highlighted_text: &str) -> (String, Vec<(usize, usize)>) {
    let mut plain_text = String::new();
    let mut ranges = Vec::new();
    let mut highlight_start = 0;
    let mut in_highlight = false;

    let chars: Vec<char> = highlighted_text.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '*' {
            if in_highlight {
                ranges.push((highlight_start, plain_text.chars().count()));
                in_highlight = false;
            } else {
                highlight_start = plain_text.chars().count();
                in_highlight = true;
            }
            i += 1;
        } else {
            plain_text.push(chars[i]);
            i += 1;
        }
    }
    (plain_text, ranges)
}

/// ファイル/フォルダのアイコンインデックスを取得する
fn get_icon_index(name: &str, is_folder: bool, himagelist: HIMAGELIST) -> i32 {
    let mut shfi: SHFILEINFOW = unsafe { std::mem::zeroed() };
    let file_name_w = str_to_wide(name);
    let mut flags = SHGFI_ICON | SHGFI_USEFILEATTRIBUTES;
    if himagelist.0 != 0 { flags |= SHGFI_SYSICONINDEX; }
    let attr = if is_folder { FILE_ATTRIBUTE_DIRECTORY } else { FILE_ATTRIBUTE_NORMAL };

    unsafe {
        SHGetFileInfoW(PCWSTR(file_name_w.as_ptr()), attr, Some(&mut shfi as *mut _), std::mem::size_of::<SHFILEINFOW>() as u32, flags);
    }
    shfi.iIcon
}

fn format_with_commas(n: u64) -> String {
    let s = n.to_string();
    let bytes = s.as_bytes();
    let mut result = Vec::new();
    let len = bytes.len();
    let first = len % 3;

    if first > 0 {
        result.extend_from_slice(&bytes[..first]);
        if len > first { result.push(b','); }
    }

    for (i, chunk) in bytes[first..].chunks(3).enumerate() {
        result.extend_from_slice(chunk);
        if i < (len - first) / 3 - 1 { result.push(b','); }
    }
    String::from_utf8(result).unwrap_or_default()
}

/// ファイルサイズをKB単位の文字列にフォーマットする
fn format_size(bytes: u64) -> String {
    if bytes == 0 { return "".to_string(); }
    let kb = (bytes + 1023) / 1024;
    format!("{} KB", format_with_commas(kb))
}

/// FILETIME(u64)を"YYYY-MM-DD HH:MM"形式の文字列に変換する
fn format_date(filetime: u64) -> String {
    if filetime == 0 { return String::new(); }
    let ft = FILETIME { dwLowDateTime: (filetime & 0xFFFFFFFF) as u32, dwHighDateTime: (filetime >> 32) as u32 };
    let mut st = SYSTEMTIME::default();
    if unsafe { FileTimeToSystemTime(&ft, &mut st).is_ok() } {
        format!("{:04}-{:02}-{:02} {:02}:{:02}", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute)
    } else {
        String::new()
    }
}

/// テキストをクリップボードにコピーする
fn copy_text_to_clipboard(window: HWND, text: &str) {
    let path_w = str_to_wide(text);
    unsafe {
        if OpenClipboard(Some(window)).is_ok() {
            let _ = EmptyClipboard();
            if let Ok(h_mem) = GlobalAlloc(GMEM_MOVEABLE, path_w.len() * std::mem::size_of::<u16>()) {
                let p_mem = GlobalLock(h_mem);
                if !p_mem.is_null() {
                    std::ptr::copy_nonoverlapping(path_w.as_ptr() as *const _, p_mem, path_w.len() * std::mem::size_of::<u16>());
                    let _ = GlobalUnlock(h_mem);
                    let _ = SetClipboardData(CF_UNICODETEXT.0 as u32, Some(HANDLE(h_mem.0 as *mut _)));
                }
            }
            let _ = CloseClipboard();
        }
    }
}

