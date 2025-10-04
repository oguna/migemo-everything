#![allow(static_mut_refs)]
// column.rs

// Cargo.toml に以下を追加してください:
// [dependencies.windows]
// version = "0.52.0"
// features = [
//     "Win32_Foundation",
//     "Win32_System_LibraryLoader",
//     "Win32_System_Time",
//     "Win32_UI_Controls",
//     "Win32_UI_WindowsAndMessaging",
//     "Win32_Storage_FileSystem",
// ]

use std::cmp::Ordering;
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use std::path::PathBuf;
use std::time::SystemTime;

use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::LibraryLoader::GetModuleHandleW,
        System::Time::*,
        UI::Controls::*,
        UI::WindowsAndMessaging::*,
        Storage::FileSystem::{FileTimeToLocalFileTime},
    },
};

const ID_LISTVIEW: isize = 1000;

// ヘッダーのコンテキストメニュー用ID
const IDM_TOGGLE_TYPE: u16 = 101;
const IDM_TOGGLE_SIZE: u16 = 102;
const IDM_TOGGLE_MODIFIED: u16 = 103;

// 列のインデックス
const COLUMN_NAME: i32 = 0;
const COLUMN_TYPE: i32 = 1;
const COLUMN_SIZE: i32 = 2;
const COLUMN_MODIFIED: i32 = 3;

/// 各ファイル/フォルダの情報を保持する構造体
#[derive(Debug)]
struct FileInfo {
    path: PathBuf,
    is_dir: bool,
    size: u64,
    modified: SystemTime,
}

// アプリケーションの状態を管理するグローバル変数
static mut FILE_ITEMS: Vec<FileInfo> = Vec::new();
static mut SORT_COLUMN: i32 = 0;
static mut SORT_ASCENDING: bool = true;
// [Type, Size, Modified] の表示状態
static mut COLUMN_VISIBILITY: [bool; 3] = [true, true, true];

fn main() -> Result<()> {
    let instance = unsafe { GetModuleHandleW(None)? };
    let window_class_name = w!("ListViewColumnSample");

    let wc = WNDCLASSW {
        hCursor: unsafe { LoadCursorW(None, IDC_ARROW)? },
        hInstance: instance.into(),
        lpszClassName: window_class_name,
        style: CS_HREDRAW | CS_VREDRAW,
        lpfnWndProc: Some(wndproc),
        ..Default::default()
    };

    let _atom = unsafe { RegisterClassW(&wc) };

    let _hwnd = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            window_class_name,
            w!("Rust ListView Column Sample"),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            800,
            600,
            None,
            None,
            Some(instance.into()),
            None,
        )
    }?;

    let mut message = MSG::default();
    while unsafe { GetMessageW(&mut message, None, 0, 0) }.as_bool() {
        unsafe {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
    Ok(())
}

extern "system" fn wndproc(window: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match message {
            WM_CREATE => {
                let instance = GetModuleHandleW(None).unwrap();
                let listview_hwnd = create_listview(window, instance.into()).unwrap();
                setup_columns(listview_hwnd);
                populate_listview(listview_hwnd);
                LRESULT(0)
            }
            WM_SIZE => {
                if let Ok(listview_hwnd) = GetDlgItem(Some(window), ID_LISTVIEW as i32) {
                    let mut rect = RECT::default();
                    GetClientRect(window, &mut rect).unwrap();
                    SetWindowPos(
                        listview_hwnd,
                        None,
                        0,
                        0,
                        rect.right - rect.left,
                        rect.bottom - rect.top,
                        SWP_NOZORDER,
                    )
                    .unwrap();
                }
                LRESULT(0)
            }
            WM_NOTIFY => {
                let nmhdr: &NMHDR = &*(lparam.0 as *const NMHDR);
                
                // リストビューの通知を処理 (ソート)
                if nmhdr.idFrom as isize == ID_LISTVIEW {
                    if nmhdr.code == LVN_COLUMNCLICK {
                        let nmlv = &*(lparam.0 as *const NMLISTVIEW);
                        let clicked_column = nmlv.iSubItem;

                        if clicked_column == SORT_COLUMN {
                            SORT_ASCENDING = !SORT_ASCENDING;
                        } else {
                            SORT_COLUMN = clicked_column;
                            SORT_ASCENDING = true;
                        }
                        
                        SendMessageW(
                            nmhdr.hwndFrom,
                            LVM_SORTITEMS,
                            Some(WPARAM(SORT_COLUMN as usize)),
                            Some(LPARAM(compare_func as isize)),
                        );
                    }
                }

                // ヘッダーの通知を処理 (コンテキストメニュー)
                if let Ok(listview_hwnd) = GetDlgItem(Some(window), ID_LISTVIEW as i32) {
                    let header_hwnd = HWND(SendMessageW(listview_hwnd, LVM_GETHEADER, None, None).0 as *mut _);
                    if nmhdr.hwndFrom == header_hwnd && nmhdr.code == NM_RCLICK {
                        let mut pt = POINT::default();
                        GetCursorPos(&mut pt).unwrap();
                        show_header_context_menu(window, pt);
                    }
                }
                
                LRESULT(0)
            }
            WM_COMMAND => {
                let command_id = (wparam.0 & 0xFFFF) as u16;
                match command_id {
                    IDM_TOGGLE_TYPE | IDM_TOGGLE_SIZE | IDM_TOGGLE_MODIFIED => {
                        let index = (command_id - IDM_TOGGLE_TYPE) as usize;
                        COLUMN_VISIBILITY[index] = !COLUMN_VISIBILITY[index];

                        if let Ok(listview_hwnd) = GetDlgItem(Some(window), ID_LISTVIEW as i32) {
                            while SendMessageW(listview_hwnd, LVM_DELETECOLUMN, Some(WPARAM(0)), None) != LRESULT(0) {}
                            setup_columns(listview_hwnd);
                            populate_listview(listview_hwnd);
                        }
                    }
                    _ => {}
                }
                LRESULT(0)
            }
            WM_DESTROY => {
                PostQuitMessage(0);
                LRESULT(0)
            }
            _ => DefWindowProcW(window, message, wparam, lparam),
        }
    }
}

/// リストビューを作成する
fn create_listview(parent: HWND, instance: HINSTANCE) -> Result<HWND> {
    unsafe {
        let icex = INITCOMMONCONTROLSEX {
            dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
            dwICC: ICC_LISTVIEW_CLASSES,
        };
        let _ = InitCommonControlsEx(&icex);

        let style = WS_CHILD | WS_VISIBLE | WINDOW_STYLE(LVS_REPORT as u32);
        let listview_hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            WC_LISTVIEW,
            w!(""),
            style,
            0,
            0,
            0,
            0,
            Some(parent),
            Some(HMENU(ID_LISTVIEW as *mut _)),
            Some(instance),
            None,
        )?;
        Ok(listview_hwnd)
    }
}

/// リストビューの列を設定する
fn setup_columns(listview_hwnd: HWND) {
    let columns = [
        ("名前", 250, None),
        ("種類", 150, Some(unsafe { &COLUMN_VISIBILITY[0] })),
        ("サイズ(バイト)", 100, Some(unsafe { &COLUMN_VISIBILITY[1] })),
        ("更新日時", 150, Some(unsafe { &COLUMN_VISIBILITY[2] })),
    ];
    
    let mut display_index = 0;
    for (i, (name, width, visibility)) in columns.iter().enumerate() {
        if visibility.map_or(true, |v| *v) {
            let mut col_text: Vec<u16> = OsStr::new(name).encode_wide().chain(Some(0)).collect();
            let col = LVCOLUMNW {
                mask: LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM,
                cx: *width,
                pszText: PWSTR(col_text.as_mut_ptr()),
                iSubItem: i as i32,
                ..Default::default()
            };
            unsafe {
                SendMessageW(
                    listview_hwnd,
                    LVM_INSERTCOLUMNW,
                    Some(WPARAM(display_index)),
                    Some(LPARAM(&col as *const _ as isize)),
                );
            }
            display_index += 1;
        }
    }
}

/// リストビューにカレントディレクトリのファイル/フォルダを読み込む
fn populate_listview(listview_hwnd: HWND) {
    unsafe {
        SendMessageW(listview_hwnd, LVM_DELETEALLITEMS, None, None);
        FILE_ITEMS.clear();

        if let Ok(current_dir) = std::env::current_dir() {
            if let Ok(entries) = std::fs::read_dir(current_dir) {
                for entry in entries.flatten() {
                    if let Ok(metadata) = entry.metadata() {
                        FILE_ITEMS.push(FileInfo {
                            path: entry.path(),
                            is_dir: metadata.is_dir(),
                            size: metadata.len(),
                            modified: metadata.modified().unwrap_or(SystemTime::UNIX_EPOCH),
                        });
                    }
                }
            }
        }

        sort_file_items();

        for (i, file_info) in FILE_ITEMS.iter().enumerate() {
            let file_name = file_info.path.file_name().unwrap_or_default().to_string_lossy();
            let mut item_text: Vec<u16> = OsStr::new(&*file_name).encode_wide().chain(Some(0)).collect();
            
            let item = LVITEMW {
                mask: LVIF_TEXT | LVIF_PARAM,
                iItem: i as i32,
                pszText: PWSTR(item_text.as_mut_ptr()),
                lParam: LPARAM(i as isize),
                ..Default::default()
            };
            
            SendMessageW(listview_hwnd, LVM_INSERTITEMW, None, Some(LPARAM(&item as *const _ as isize)));

            let mut subitem_display_index = 1;
            if COLUMN_VISIBILITY[0] {
                let type_str = if file_info.is_dir { "ファイル フォルダー" } else { "ファイル" };
                set_subitem_text(listview_hwnd, i as i32, subitem_display_index, type_str);
                subitem_display_index += 1;
            }
            if COLUMN_VISIBILITY[1] {
                let size_str = if file_info.is_dir {
                    "".to_string()
                } else {
                    // ファイルサイズをそのまま文字列に変換
                    file_info.size.to_string()
                };
                set_subitem_text(listview_hwnd, i as i32, subitem_display_index, &size_str);
                subitem_display_index += 1;
            }
            if COLUMN_VISIBILITY[2] {
                let ft_u64 = systemtime_to_filetime_u64(file_info.modified);
                let modified_str = format_date(ft_u64);
                set_subitem_text(listview_hwnd, i as i32, subitem_display_index, &modified_str);
            }
        }
    }
}

/// リストビューのサブアイテムにテキストを設定するヘルパー関数
fn set_subitem_text(listview: HWND, item_index: i32, subitem_index: i32, text: &str) {
    let mut text_w: Vec<u16> = OsStr::new(text).encode_wide().chain(Some(0)).collect();
    let subitem = LVITEMW {
        iSubItem: subitem_index,
        pszText: PWSTR(text_w.as_mut_ptr()),
        ..Default::default()
    };
    unsafe {
        SendMessageW(
            listview,
            LVM_SETITEMTEXTW,
            Some(WPARAM(item_index as usize)),
            Some(LPARAM(&subitem as *const _ as isize)),
        );
    }
}

/// `FILE_ITEMS` ベクタを現在のソート設定でソートする
fn sort_file_items() {
    unsafe {
        FILE_ITEMS.sort_by(|a, b| {
            let ordering = match SORT_COLUMN {
                COLUMN_NAME => a.path.file_name().cmp(&b.path.file_name()),
                COLUMN_TYPE => a.is_dir.cmp(&b.is_dir).reverse(),
                COLUMN_SIZE => a.size.cmp(&b.size),
                COLUMN_MODIFIED => a.modified.cmp(&b.modified),
                _ => Ordering::Equal,
            };
            if SORT_ASCENDING {
                ordering
            } else {
                ordering.reverse()
            }
        });
    }
}

/// `ListView_SortItems` のための比較コールバック関数
extern "system" fn compare_func(lparam1: LPARAM, lparam2: LPARAM, lparam_sort: LPARAM) -> i32 {
    unsafe {
        let index1 = lparam1.0 as usize;
        let index2 = lparam2.0 as usize;
        let sort_column = lparam_sort.0 as i32;

        if let (Some(item1), Some(item2)) = (FILE_ITEMS.get(index1), FILE_ITEMS.get(index2)) {
            let ordering = match sort_column {
                COLUMN_NAME => item1.path.file_name().cmp(&item2.path.file_name()),
                COLUMN_TYPE => item1.is_dir.cmp(&item2.is_dir).reverse(),
                COLUMN_SIZE => item1.size.cmp(&item2.size),
                COLUMN_MODIFIED => item1.modified.cmp(&item2.modified),
                _ => Ordering::Equal,
            };

            let result = if SORT_ASCENDING { ordering } else { ordering.reverse() };
            
            match result {
                Ordering::Less => -1,
                Ordering::Equal => 0,
                Ordering::Greater => 1,
            }
        } else {
            0
        }
    }
}

/// ヘッダーの右クリックでコンテキストメニューを表示する
fn show_header_context_menu(owner: HWND, pt: POINT) {
    unsafe {
        let hmenu = CreatePopupMenu().unwrap();

        let items = [
            ("種類(&T)", IDM_TOGGLE_TYPE, COLUMN_VISIBILITY[0]),
            ("サイズ(&S)", IDM_TOGGLE_SIZE, COLUMN_VISIBILITY[1]),
            ("更新日時(&D)", IDM_TOGGLE_MODIFIED, COLUMN_VISIBILITY[2]),
        ];

        for (name, id, is_visible) in items {
            let mut flags = MF_STRING;
            if is_visible {
                flags |= MF_CHECKED;
            }
            let text: Vec<u16> = OsStr::new(name).encode_wide().chain(Some(0)).collect();
            AppendMenuW(hmenu, flags, id as usize, PCWSTR(text.as_ptr())).unwrap();
        }

        let _ = TrackPopupMenuEx(hmenu, (TPM_TOPALIGN | TPM_LEFTALIGN).0, pt.x, pt.y, owner, None);
        DestroyMenu(hmenu).unwrap();
    }
}

/// SystemTimeをFILETIME(u64)に変換する
fn systemtime_to_filetime_u64(st: SystemTime) -> u64 {
    const UNIX_EPOCH_AS_FILETIME: u64 = 116444736000000000;
    const HUNDRED_NS_PER_SEC: u64 = 10_000_000;

    match st.duration_since(SystemTime::UNIX_EPOCH) {
        Ok(d) => {
            UNIX_EPOCH_AS_FILETIME + d.as_secs() * HUNDRED_NS_PER_SEC + (d.subsec_nanos() / 100) as u64
        },
        Err(_) => 0,
    }
}

/// FILETIME(u64)を"YYYY-MM-DD HH:MM:SS"形式の文字列に変換する
fn format_date(filetime: u64) -> String {
    if filetime == 0 { return String::new(); }
    let mut ft = FILETIME {
        dwLowDateTime: (filetime & 0xFFFFFFFF) as u32,
        dwHighDateTime: (filetime >> 32) as u32,
    };
    let mut st = SYSTEMTIME::default();
    // ローカルタイムに変換
    unsafe { let _ = FileTimeToLocalFileTime(&ft, &mut ft); }
    if unsafe { FileTimeToSystemTime(&ft, &mut st).is_ok() } {
        format!(
            "{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
            st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond
        )
    } else {
        String::new()
    }
}
