// main.rs (dragdrop.rs)

// Cargo.toml に以下を追加してください:
// [dependencies.windows]
// version = "0.61.0"
// features = [
//     "Win32_Foundation",
//     "Win32_Graphics_Gdi",
//     "Win32_System_Com",
//     "Win32_System_LibraryLoader",
//     "Win32_System_Memory",
//     "Win32_System_Ole", // OLE (ドラッグ&ドロップ) のために追加
//     "Win32_System_SystemServices",
//     "Win32_UI_Controls",
//     "Win32_UI_Shell",
//     "Win32_UI_Shell_Common",
//     "Win32_UI_WindowsAndMessaging",
// ]

use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::{
            Com::{
                CoInitializeEx, CoUninitialize,
                COINIT_APARTMENTTHREADED,
            },
            LibraryLoader::GetModuleHandleW,
        },
        UI::{
            Controls::*,
            WindowsAndMessaging::*,
        },
    },
};

const ID_LISTVIEW: isize = 1000;

fn main() -> Result<()> {
    // 1. COMライブラリの初期化
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
    }

    let instance = unsafe { GetModuleHandleW(None)? };
    let window_class_name = w!("DragDropSample");

    let wc = WNDCLASSW {
        hCursor: unsafe { LoadCursorW(None, IDC_ARROW)? },
        hInstance: instance.into(),
        lpszClassName: window_class_name,
        style: CS_HREDRAW | CS_VREDRAW,
        lpfnWndProc: Some(wndproc),
        ..Default::default()
    };

    let atom = unsafe { RegisterClassW(&wc) };
    if atom == 0 {
        return Err(Error::from_win32());
    }

    let hwnd = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            window_class_name,
            w!("Rust Drag Drop Sample"),
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
    if hwnd.is_invalid() {
        return Err(Error::from_win32());
    }

    let mut message = MSG::default();
    while unsafe { GetMessageW(&mut message, None, 0, 0) }.as_bool() {
        unsafe {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    // COMライブラリの解放
    unsafe { CoUninitialize() };
    Ok(())
}

extern "system" fn wndproc(window: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match message {
            WM_CREATE => {
                let instance = GetModuleHandleW(None).unwrap();
                create_listview(window, instance.into()).unwrap();
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
                if nmhdr.idFrom as isize == ID_LISTVIEW {
                    // ドラッグ開始の通知を処理
                    if nmhdr.code == LVN_BEGINDRAG {
                        let nmlistview: &NMLISTVIEW = &*(lparam.0 as *const NMLISTVIEW);
                        handle_drag_begin(nmhdr.hwndFrom, nmlistview.iItem);
                    }
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

// リストビューを作成し、初期化する関数
fn create_listview(parent: HWND, instance: HINSTANCE) -> Result<()> {
    unsafe {
        let icex = INITCOMMONCONTROLSEX {
            dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
            dwICC: ICC_LISTVIEW_CLASSES,
        };
        let _ = InitCommonControlsEx(&icex);

        let style = WS_CHILD | WS_VISIBLE | WINDOW_STYLE(LVS_REPORT) | WINDOW_STYLE(LVS_SINGLESEL);

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

        let mut col = LVCOLUMNW::default();
        col.mask = LVCF_TEXT | LVCF_WIDTH;
        col.cx = 250;

        let mut col_text: Vec<u16> = OsStr::new("名前").encode_wide().chain(Some(0)).collect();
        col.pszText = PWSTR(col_text.as_mut_ptr());
        SendMessageW(
            listview_hwnd,
            LVM_INSERTCOLUMNW,
            Some(WPARAM(0)),
            Some(LPARAM(&col as *const _ as isize)),
        );

        populate_listview(listview_hwnd)?;
    }
    Ok(())
}

// リストビューにカレントディレクトリのファイル/フォルダを populate する関数
fn populate_listview(listview_hwnd: HWND) -> Result<()> {
    let current_dir = std::env::current_dir().unwrap();
    let mut item_index = 0;

    for entry in std::fs::read_dir(current_dir).unwrap() {
        let entry = entry.unwrap();
        let path = entry.path();
        let file_name = path.file_name().unwrap_or_default().to_string_lossy();
        let mut item_text: Vec<u16> = OsStr::new(&*file_name).encode_wide().chain(Some(0)).collect();

        let item = LVITEMW {
            mask: LVIF_TEXT,
            iItem: item_index,
            pszText: PWSTR(item_text.as_mut_ptr()),
            ..Default::default()
        };

        unsafe {
            SendMessageW(
                listview_hwnd,
                LVM_INSERTITEMW,
                Some(WPARAM(0)),
                Some(LPARAM(&item as *const _ as isize)),
            );
        }
        item_index += 1;
    }
    Ok(())
}

// ドラッグ開始を処理する関数
fn handle_drag_begin(listview_hwnd: HWND, item_index: i32) {
    if item_index < 0 {
        return;
    }

    // 1. ドラッグされるアイテムのフルパスを取得
    let mut text_buffer: [u16; MAX_PATH as usize] = [0; MAX_PATH as usize];
    let mut item = LVITEMW {
        mask: LVIF_TEXT,
        iItem: item_index,
        iSubItem: 0,
        pszText: PWSTR(text_buffer.as_mut_ptr()),
        cchTextMax: MAX_PATH as i32,
        ..Default::default()
    };
    unsafe {
        SendMessageW(
            listview_hwnd,
            LVM_GETITEMW,
            Some(WPARAM(0)),
            Some(LPARAM(&mut item as *mut _ as isize)),
        );
    }

    let file_name = unsafe { item.pszText.to_string().unwrap() };
    println!("Dragging: {}", file_name);
    
    // 簡略化: 実際のドラッグ&ドロップはここでは省略
    // 完全な実装にはより複雑なCOM操作が必要
}
