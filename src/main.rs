// リリースビルド時にコンソールウィンドウを非表示にする
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

// --- 依存クレート ---
use windows::{
    core::*,
    Win32::Foundation::*,
    Win32::Graphics::Gdi::*,
    Win32::System::DataExchange::{CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData},
    Win32::System::LibraryLoader::GetModuleHandleA,
    Win32::System::Memory::{GlobalAlloc, GlobalLock, GlobalUnlock, GMEM_MOVEABLE},
    Win32::System::Ole::CF_UNICODETEXT,
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
        ShellExecuteW, SHFILEINFOW, SHGFI_ICON, SHGFI_SMALLICON, SHGFI_SYSICONINDEX,
        SHGFI_USEFILEATTRIBUTES, SHGetFileInfoW,
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
use std::path::Path;
use std::sync::Mutex;
use std::thread;

// --- 定数 ---

/// コントロールID: テキスト入力
const EDIT_ID: u16 = 1000;
/// コントロールID: REボタン
const RE_BUTTON_ID: u16 = 1001;
/// コントロールID: Miボタン
const MI_BUTTON_ID: u16 = 1002;

/// タイマーID
const TIMER_ID: usize = 1;

/// メニューID: 終了
const IDM_FILE_EXIT: u16 = 2001;
/// メニューID: 正規表現検索
const IDM_SEARCH_REGEX: u16 = 3001;
/// メニューID: Migemo検索
const IDM_SEARCH_MIGEMO: u16 = 3002;

/// コンテキストメニューID: 開く
const IDM_CONTEXT_OPEN: u16 = 4001;
/// コンテキストメニューID: フォルダーを開く
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
    pub himagelist: HIMAGELIST,

    // --- DPI関連 ---
    pub current_dpi: u32,
    pub scale_factor: f32,

    // --- 検索オプション ---
    pub regex_enabled: bool,
    pub migemo_enabled: bool,

    // --- データ ---
    pub migemo_dict: Option<CompactDictionary>,
    pub search_results: Mutex<Vec<FileResult>>,

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
            himagelist: HIMAGELIST::default(),
            current_dpi: 96,  // デフォルトDPI
            scale_factor: 1.0,  // デフォルトスケール
            regex_enabled: false,
            migemo_enabled: false,
            migemo_dict,
            search_results: Mutex::new(Vec::new()),
            item_wide_buffer: [Vec::new(), Vec::new(), Vec::new(), Vec::new()],
        }
    }
}

// --- main関数 ---

/// アプリケーションのエントリポイント
fn main() -> Result<()> {
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
        let _hwnd = CreateWindowExW(
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

        // メッセージループ
        let mut message = MSG::default();
        while GetMessageW(&mut message, None, 0, 0).into() {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
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
        // --- メニュー項目 ---
        IDM_FILE_EXIT => {
            let _ = unsafe { DestroyWindow(window) };
        }
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
        // --- エディットボックス ---
        EDIT_ID if notification_code as u32 == EN_CHANGE => {
            // 500ミリ秒後に検索タイマーをセット
            unsafe { SetTimer(Some(window), TIMER_ID, 500, None) };
        }
        // --- コンテキストメニュー ---
        IDM_CONTEXT_OPEN => {
            let item_index = lparam.0 as usize;
            let results = state.search_results.lock().unwrap();
            if let Some(result) = results.get(item_index) {
                let full_path = Path::new(&result.path).join(&result.name);
                let path_w = str_to_wide(full_path.to_str().unwrap_or(""));
                thread::spawn(move || unsafe {
                    ShellExecuteW(None, w!("open"), PCWSTR(path_w.as_ptr()), None, None, SW_SHOW);
                });
            }
        }
        IDM_CONTEXT_OPEN_FOLDER => {
            let item_index = lparam.0 as usize;
            let results = state.search_results.lock().unwrap();
            if let Some(result) = results.get(item_index) {
                let full_path = Path::new(&result.path).join(&result.name);
                let params = format!("/select,\"{}\"", full_path.display());
                let params_w = str_to_wide(&params);
                thread::spawn(move || unsafe {
                    ShellExecuteW(
                        None, w!("open"), w!("explorer.exe"),
                        PCWSTR(params_w.as_ptr()), None, SW_SHOW,
                    );
                });
            }
        }
        IDM_CONTEXT_COPY_PATH => {
            let item_index = lparam.0 as usize;
            let results = state.search_results.lock().unwrap();
            if let Some(result) = results.get(item_index) {
                let full_path_str = Path::new(&result.path)
                    .join(&result.name)
                    .to_str().unwrap_or("").to_string();
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
            NM_RCLICK => handle_right_click(window, lparam),
            NM_DBLCLK => {
                let item_activate = unsafe { &*(lparam.0 as *const NMITEMACTIVATE) };
                if item_activate.iItem != -1 {
                    unsafe {
                        SendMessageW(
                            window, WM_COMMAND,
                            Some(WPARAM(IDM_CONTEXT_OPEN as usize)),
                            Some(LPARAM(item_activate.iItem as isize)),
                        );
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
    // 新しいDPIを取得
    let new_dpi = hiword(wparam.0 as u32) as u32;
    
    // スケールファクターを計算
    state.current_dpi = new_dpi;
    state.scale_factor = new_dpi as f32 / 96.0;
    
    // 推奨ウィンドウサイズと位置を取得
    let suggested_rect = unsafe { &*(lparam.0 as *const RECT) };
    
    // ウィンドウサイズと位置を更新
    unsafe {
        let _ = SetWindowPos(
            window,
            None,
            suggested_rect.left,
            suggested_rect.top,
            suggested_rect.right - suggested_rect.left,
            suggested_rect.bottom - suggested_rect.top,
            SWP_NOZORDER | SWP_NOACTIVATE,
        );
    }
    
    // UIを再描画
    unsafe {
        let _ = InvalidateRect(Some(window), None, true);
    }
    
    LRESULT(0)
}

// --- イベントハンドラ (WM_NOTIFY) のためのヘルパー関数 ---

fn handle_get_disp_info(lparam: LPARAM, state: &mut AppState) {
    let dispinfo = unsafe { &mut *(lparam.0 as *mut NMLVDISPINFOW) };
    let item = &mut dispinfo.item;
    let item_index = item.iItem as usize;

    let results = state.search_results.lock().unwrap();
    if let Some(result) = results.get(item_index) {
        // テキスト情報
        if (item.mask & LVIF_TEXT) == LVIF_TEXT {
            let sub_item_index = item.iSubItem as usize;
            let text = match sub_item_index {
                0 => {
                    if !result.highlighted_name.is_empty() {
                        parse_highlight_text(&result.highlighted_name).0
                    } else {
                        result.name.clone()
                    }
                }
                1 => {
                    if !result.highlighted_path.is_empty() {
                        parse_highlight_text(&result.highlighted_path).0
                    } else {
                        result.path.clone()
                    }
                }
                2 => format_size(result.size),
                3 => format_date(result.modified_date),
                _ => String::new(),
            };
            state.item_wide_buffer[sub_item_index] = str_to_wide(&text);
            item.pszText = PWSTR(state.item_wide_buffer[sub_item_index].as_mut_ptr());
        }
        // アイコン情報
        if item.iSubItem == 0 && (item.mask & LVIF_IMAGE) == LVIF_IMAGE {
            item.iImage = get_icon_index(&result.name, result.is_folder, state.himagelist);
        }
    }
}

fn handle_custom_draw(lparam: LPARAM, state: &mut AppState) -> LRESULT {
    let custom_draw = unsafe { &mut *(lparam.0 as *mut NMLVCUSTOMDRAW) };

    match custom_draw.nmcd.dwDrawStage {
        CDDS_PREPAINT => {
            // アイテムごとの描画通知を要求
            LRESULT(CDRF_NOTIFYITEMDRAW as isize)
        }
        CDDS_ITEMPREPAINT => {
            // サブアイテムごとの描画通知を要求
            LRESULT(CDRF_NOTIFYSUBITEMDRAW as isize)
        }
        stage if stage.0 == (CDDS_SUBITEM.0 | CDDS_ITEMPREPAINT.0) => {
            let item_index = custom_draw.nmcd.dwItemSpec as usize;
            let sub_item_index = custom_draw.iSubItem as usize;

            let results = state.search_results.lock().unwrap();
            if let Some(result) = results.get(item_index) {
                // ハイライト対象の文字列と、そのプレーンテキスト/ハイライト範囲を取得
                let (text_to_draw, highlight_ranges) = match sub_item_index {
                    0 if !result.highlighted_name.is_empty() => {
                        parse_highlight_text(&result.highlighted_name)
                    }
                    1 if !result.highlighted_path.is_empty() => {
                        parse_highlight_text(&result.highlighted_path)
                    }
                    // ハイライトがない場合はデフォルト描画に任せる
                    _ => return LRESULT(CDRF_DODEFAULT as isize),
                };

                // ハイライト範囲がない場合もデフォルト描画
                if highlight_ranges.is_empty() {
                    return LRESULT(CDRF_DODEFAULT as isize);
                }

                // --- ここからカスタム描画処理 ---
                let hdc = custom_draw.nmcd.hdc;
                let mut rect = custom_draw.nmcd.rc;
                let is_selected = (custom_draw.nmcd.uItemState & CDIS_SELECTED).0 != 0;

                // 1. 背景を描画
                let bg_color = if is_selected {
                    unsafe { GetSysColor(COLOR_HIGHLIGHT) }
                } else {
                    unsafe { GetSysColor(COLOR_WINDOW) }
                };
                let bg_brush = unsafe { CreateSolidBrush(COLORREF(bg_color)) };
                unsafe { FillRect(hdc, &rect, bg_brush) };
                let _ = unsafe { DeleteObject(bg_brush.into()) };

                // 2. アイコンを描画 (1列目のみ)
                if sub_item_index == 0 {
                    let icon_index = get_icon_index(&result.name, result.is_folder, state.himagelist);
                    if state.himagelist.0 != 0 {
                        let icon_size = (16.0 * state.scale_factor) as i32;
                        let icon_padding = (2.0 * state.scale_factor) as i32;
                        let _ = unsafe {
                            ImageList_Draw(
                                state.himagelist,
                                icon_index,
                                hdc,
                                rect.left + icon_padding,
                                rect.top + (rect.bottom - rect.top - icon_size) / 2, // 中央揃え
                                ILD_TRANSPARENT,
                            )
                        };
                    }
                    // アイコンの分だけ描画開始位置をずらす
                    rect.left += (22.0 * state.scale_factor) as i32;
                } else {
                    rect.left += (4.0 * state.scale_factor) as i32; // パディング
                }
                rect.right -= (4.0 * state.scale_factor) as i32;

                // 3. テキストとハイライトを描画
                let text_color = if is_selected {
                    unsafe { GetSysColor(COLOR_HIGHLIGHTTEXT) }
                } else {
                    unsafe { GetSysColor(COLOR_WINDOWTEXT) }
                };
                unsafe {
                    SetBkMode(hdc, TRANSPARENT);
                    SetTextColor(hdc, COLORREF(text_color));
                }

                let mut x = rect.left;
                let font_offset = (8.0 * state.scale_factor) as i32;
                let y = rect.top + (rect.bottom - rect.top) / 2 - font_offset;
                let chars: Vec<char> = text_to_draw.chars().collect();
                
                // 全テキストの文字位置を事前に計算
                let full_text_wide = str_to_wide(&text_to_draw);
                let mut char_widths = vec![0i32; chars.len()];
                if chars.len() > 0 && full_text_wide.len() > 1 {
                    let mut fit_count = 0i32;
                    let mut size = SIZE::default();
                    let _ = unsafe {
                        GetTextExtentExPointW(
                            hdc,
                            PCWSTR(full_text_wide.as_ptr()),
                            (full_text_wide.len() - 1) as i32, // null終端を除く長さ
                            rect.right - rect.left,
                            Some(&mut fit_count),
                            Some(char_widths.as_mut_ptr()),
                            &mut size,
                        )
                    };
                }
                
                // 文字をグループ化して連続描画
                let mut current_pos = 0;
                while current_pos < chars.len() {
                    // 現在の位置からハイライト状態が同じ範囲を見つける
                    let is_current_highlighted = highlight_ranges.iter().any(|(start, end)| current_pos >= *start && current_pos < *end);
                    let mut end_pos = current_pos + 1;
                    
                    // 同じハイライト状態の文字が続く限り範囲を拡張
                    while end_pos < chars.len() {
                        let is_next_highlighted = highlight_ranges.iter().any(|(start, end)| end_pos >= *start && end_pos < *end);
                        if is_current_highlighted == is_next_highlighted {
                            end_pos += 1;
                        } else {
                            break;
                        }
                    }
                    
                    // この範囲のテキストを取得
                    let text_segment: String = chars[current_pos..end_pos].iter().collect();
                    let text_wide = str_to_wide(&text_segment);
                    
                    // 正確な幅を計算（前の文字位置からの差分）
                    let start_x = if current_pos == 0 { 0 } else { char_widths[current_pos - 1] };
                    let end_x = if end_pos > 0 { char_widths[end_pos - 1] } else { 0 };
                    let segment_width = end_x - start_x;
                    
                    // 描画範囲チェック
                    if x + segment_width > rect.right { break; }
                    
                    // ハイライト背景を描画
                    if is_current_highlighted && !is_selected {
                        let highlight_brush = unsafe { CreateSolidBrush(COLORREF(0x00FFFF)) }; // 黄色
                        let highlight_rect = RECT { left: x, top: rect.top, right: x + segment_width, bottom: rect.bottom };
                        unsafe { FillRect(hdc, &highlight_rect, highlight_brush) };
                        let _ = unsafe { DeleteObject(highlight_brush.into()) };
                    }
                    
                    // テキストを描画
                    let _ = unsafe { TextOutW(hdc, x, y, &text_wide) };
                    x += segment_width;
                    current_pos = end_pos;
                }

                // デフォルト描画をスキップ
                return LRESULT(CDRF_SKIPDEFAULT as isize);
            }
            LRESULT(CDRF_DODEFAULT as isize)
        }
        _ => LRESULT(CDRF_DODEFAULT as isize),
    }
}

fn handle_right_click(window: HWND, lparam: LPARAM) {
    let item_activate = unsafe { &*(lparam.0 as *const NMITEMACTIVATE) };
    let item_index = item_activate.iItem;

    if item_index != -1 {
        unsafe {
            let h_popup_menu = CreatePopupMenu().unwrap();
            let _ = AppendMenuW(h_popup_menu, MF_STRING, IDM_CONTEXT_OPEN as usize, w!("開く(&O)"));
            let _ = AppendMenuW(h_popup_menu, MF_STRING, IDM_CONTEXT_OPEN_FOLDER as usize, w!("フォルダーを開く(&F)"));
            let _ = AppendMenuW(h_popup_menu, MF_STRING, IDM_CONTEXT_COPY_PATH as usize, w!("フルパスをコピー(&C)"));
            let _ = SetMenuDefaultItem(h_popup_menu, IDM_CONTEXT_OPEN as u32, 0);

            let mut pt = POINT::default();
            let _ = GetCursorPos(&mut pt);

            let cmd = TrackPopupMenu(
                h_popup_menu,
                TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD,
                pt.x, pt.y, Some(0), window, None,
            );

            if cmd.as_bool() {
                SendMessageW(
                    window, WM_COMMAND,
                    Some(WPARAM(cmd.0 as usize)),
                    Some(LPARAM(item_index as isize)),
                );
            }
            let _ = DestroyMenu(h_popup_menu);
        }
    }
}


// --- UI関連の関数 ---

/// メニューを作成してウィンドウに設定する
fn create_menu(window: HWND) {
    unsafe {
        let h_menu = CreateMenu().unwrap();
        let h_file_submenu = CreatePopupMenu().unwrap();
        let _ = AppendMenuW(h_file_submenu, MF_STRING, IDM_FILE_EXIT as usize, w!("終了(&X)\tAlt+F4"));
        let _ = AppendMenuW(h_menu, MF_POPUP, h_file_submenu.0 as usize, w!("ファイル(&F)"));

        let h_search_submenu = CreatePopupMenu().unwrap();
        let _ = AppendMenuW(h_search_submenu, MF_STRING, IDM_SEARCH_REGEX as usize, w!("正規表現で検索(&R)"));
        let _ = AppendMenuW(h_search_submenu, MF_STRING, IDM_SEARCH_MIGEMO as usize, w!("Migemoで検索(&M)"));
        let _ = AppendMenuW(h_menu, MF_POPUP, h_search_submenu.0 as usize, w!("検索(&S)"));
        let _ = SetMenu(window, Some(h_menu));
    }
}

/// すべてのUIコントロールを作成する（DPI対応）
fn create_controls(window: HWND, instance: HINSTANCE, state: &mut AppState) {
    // DPIに基づいてフォントサイズを調整
    let scale = state.scale_factor;
    let font_height = (-12.0 * scale) as i32; // 負の値でピクセル単位指定
    
    let h_font = unsafe {
        CreateFontW(
            font_height,       // nHeight
            0,                 // nWidth
            0,                 // nEscapement
            0,                 // nOrientation
            FW_NORMAL.0 as i32,// nWeight
            0,                 // bItalic
            0,                 // bUnderline
            0,                 // bStrikeOut
            DEFAULT_CHARSET,   // nCharSet
            OUT_DEFAULT_PRECIS, // nOutPrecision
            CLIP_DEFAULT_PRECIS, // nClipPrecision
            DEFAULT_QUALITY,   // nQuality
            (FF_DONTCARE.0 | VARIABLE_PITCH.0) as u32, // nPitchAndFamily
            w!("Segoe UI"),    // lpszFacename
        )
    };

    unsafe {
        state.status_hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(), w!("STATIC"), w!("Ready"),
            WS_CHILD | WS_VISIBLE, 0, 0, 0, 0,
            Some(window), None, Some(instance), None,
        ).unwrap();
        state.edit_hwnd = CreateWindowExW(
            WS_EX_CLIENTEDGE, w!("EDIT"), w!(""),
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_LEFT as u32 | ES_AUTOHSCROLL as u32),
            0, 0, 0, 0, Some(window), Some(HMENU(EDIT_ID as isize as *mut c_void)), Some(instance), None,
        ).unwrap();
        state.re_button_hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(), w!("BUTTON"), w!("RE"),
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_PUSHBUTTON as u32),
            0, 0, 0, 0, Some(window), Some(HMENU(RE_BUTTON_ID as isize as *mut c_void)), Some(instance), None,
        ).unwrap();
        state.mi_button_hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(), w!("BUTTON"), w!("Mi"),
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_PUSHBUTTON as u32),
            0, 0, 0, 0, Some(window), Some(HMENU(MI_BUTTON_ID as isize as *mut c_void)), Some(instance), None,
        ).unwrap();
        state.listview_hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(), w!("SysListView32"), w!(""),
            WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | WS_VSCROLL | WS_HSCROLL |
            WINDOW_STYLE(LVS_REPORT as u32 | LVS_OWNERDATA as u32),
            0, 0, 0, 0, Some(window), None, Some(instance), None,
        ).unwrap();

        // フォント設定
        if !h_font.is_invalid() {
            SendMessageW(state.status_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
            SendMessageW(state.edit_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
            SendMessageW(state.re_button_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
            SendMessageW(state.mi_button_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
            SendMessageW(state.listview_hwnd, WM_SETFONT, Some(WPARAM(h_font.0 as usize)), Some(LPARAM(1)));
        }
    }
}

/// リストビューの初期設定（カラム、拡張スタイル、イメージリスト）（DPI対応）
fn setup_listview(state: &mut AppState) {
    unsafe {
        let ex_style = LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES;
        SendMessageW(state.listview_hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, Some(WPARAM(ex_style as usize)), Some(LPARAM(ex_style as isize)));

        let mut shfi: SHFILEINFOW = std::mem::zeroed();
        state.himagelist = HIMAGELIST(SHGetFileInfoW(
            w!(""), FILE_ATTRIBUTE_NORMAL, Some(&mut shfi as *mut _),
            std::mem::size_of::<SHFILEINFOW>() as u32,
            SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON,
        ) as isize);

        if state.himagelist.0 != 0 {
            SendMessageW(state.listview_hwnd, LVM_SETIMAGELIST, Some(WPARAM(LVSIL_SMALL as usize)), Some(LPARAM(state.himagelist.0)));
        }

        // DPIに基づいてカラム幅を調整
        let scale = state.scale_factor;
        let columns = [
            (w!("名前"), (300.0 * scale) as i32),
            (w!("フォルダー"), (300.0 * scale) as i32),
            (w!("サイズ"), (80.0 * scale) as i32),
            (w!("更新日時"), (150.0 * scale) as i32)
        ];
        
        for (i, (text, width)) in columns.iter().enumerate() {
            let mut col = LVCOLUMNW {
                mask: LVCF_TEXT | LVCF_WIDTH,
                cx: *width,
                pszText: PWSTR(text.as_ptr() as *mut _), ..Default::default()
            };
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
    // DPIに基づいてサイズを調整
    let scale = state.scale_factor;
    let bar_height = (25.0 * scale) as i32;
    let status_bar_height = (20.0 * scale) as i32;
    let button_width = (40.0 * scale) as i32;
    let total_button_width = button_width * 2;
    let list_y = bar_height;

    unsafe {
        let _ = MoveWindow(state.edit_hwnd, 0, 0, width - total_button_width, bar_height, true);
        let _ = MoveWindow(state.re_button_hwnd, width - total_button_width, 0, button_width, bar_height, true);
        let _ = MoveWindow(state.mi_button_hwnd, width - button_width, 0, button_width, bar_height, true);
        let _ = MoveWindow(state.listview_hwnd, 0, list_y, width, height - list_y - status_bar_height, true);
        let _ = MoveWindow(state.status_hwnd, 0, height - status_bar_height, width, status_bar_height, true);
    }
}

// --- 検索関連の関数 ---

/// Migemo辞書を初期化する
fn init_migemo_dict() -> Option<CompactDictionary> {
    use std::env;
    use std::path::PathBuf;
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

    if search_term.is_empty() {
        state.search_results.lock().unwrap().clear();
        unsafe {
            let _ = SetWindowTextW(state.status_hwnd, w!("Ready"));
            SendMessageW(state.listview_hwnd, LVM_SETITEMCOUNT, Some(WPARAM(0)), Some(LPARAM(0)));
            let _ = InvalidateRect(Some(state.listview_hwnd), None, true);
        }
        return;
    }

    let mut guard = global().lock().unwrap();
    let mut searcher = guard.searcher();
    let final_search_term = if state.migemo_enabled {
        migemo_query(&search_term, &state.migemo_dict).unwrap_or(search_term)
    } else { search_term };

    searcher.set_search(&final_search_term);
    searcher.set_regex(state.regex_enabled || state.migemo_enabled);
    searcher.set_request_flags(
        RequestFlags::EVERYTHING_REQUEST_FILE_NAME |
        RequestFlags::EVERYTHING_REQUEST_PATH |
        RequestFlags::EVERYTHING_REQUEST_SIZE |
        RequestFlags::EVERYTHING_REQUEST_DATE_MODIFIED |
        RequestFlags::EVERYTHING_REQUEST_ATTRIBUTES |
        RequestFlags::EVERYTHING_REQUEST_HIGHLIGHTED_FILE_NAME |
        RequestFlags::EVERYTHING_REQUEST_HIGHLIGHTED_PATH
    );

    // クエリを実行し、結果オブジェクトを保持する
    let query_results = searcher.query();
    //  APIから合計件数を取得する
    let total_results = query_results.total();

    let mut new_results = Vec::new();
    const MAX_DISPLAY_COUNT: usize = 10000;
    for item in query_results.iter().take(MAX_DISPLAY_COUNT) {
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

    let displayed_count = new_results.len();
    *state.search_results.lock().unwrap() = new_results;

    // ステータスバーのテキストを、総件数と表示件数に応じて変更
    let status_text = if total_results > displayed_count as u32 {
        format!(
            "Showing {} of {} items",
            displayed_count, total_results
        )
    } else {
        format!("{} items found", total_results)
    };

    unsafe {
        let _ = SetWindowTextW(state.status_hwnd, PCWSTR(str_to_wide(&status_text).as_ptr()));
        // ListViewにセットする件数は表示件数(displayed_count)のまま
        SendMessageW(state.listview_hwnd, LVM_SETITEMCOUNT, Some(WPARAM(displayed_count)), Some(LPARAM(0)));
        let _ = InvalidateRect(Some(state.listview_hwnd), None, true);
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
                let end = plain_text.chars().count();
                ranges.push((highlight_start, end));
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
        SHGetFileInfoW(
            PCWSTR(file_name_w.as_ptr()), attr, Some(&mut shfi as *mut _),
            std::mem::size_of::<SHFILEINFOW>() as u32, flags,
        );
    }
    shfi.iIcon
}

/// ファイルサイズをKB単位の文字列にフォーマットする
fn format_size(bytes: u64) -> String {
    if bytes == 0 { return "".to_string(); }
    let kb = (bytes + 1023) / 1024;
    format!("{} KB", kb)
}

/// FILETIME(u64)を"YYYY-MM-DD HH:MM:SS"形式の文字列に変換する
fn format_date(filetime: u64) -> String {
    if filetime == 0 { return String::new(); }
    let ft = FILETIME {
        dwLowDateTime: (filetime & 0xFFFFFFFF) as u32,
        dwHighDateTime: (filetime >> 32) as u32,
    };
    let mut st = SYSTEMTIME::default();
    if unsafe { FileTimeToSystemTime(&ft, &mut st).is_ok() } {
        format!(
            "{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
            st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond
        )
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
                    std::ptr::copy_nonoverlapping(
                        path_w.as_ptr() as *const _, p_mem,
                        path_w.len() * std::mem::size_of::<u16>(),
                    );
                    let _ = GlobalUnlock(h_mem);
                    let _ = SetClipboardData(CF_UNICODETEXT.0 as u32, Some(HANDLE(h_mem.0 as *mut _)));
                }
            }
            let _ = CloseClipboard();
        }
    }
}
