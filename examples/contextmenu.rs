// main.rs

// Cargo.toml に以下を追加してください:
// [dependencies.windows]
// version = "0.56.0" // Or a newer version
// features = [
//     "Win32_Foundation",
//     "Win32_Graphics_Gdi",
//     "Win32_System_Com",
//     "Win32_System_LibraryLoader",
//     "Win32_System_Memory",
//     "Win32_System_SystemServices",
//     "Win32_UI_Controls",
//     "Win32_UI_Shell",
//     "Win32_UI_Shell_Common",
//     "Win32_UI_WindowsAndMessaging",
// ]

use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use std::path::Path;
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::Gdi::ClientToScreen,
        System::{
            Com::{
                CoInitializeEx, CoUninitialize, CoTaskMemFree,
                COINIT_APARTMENTTHREADED,
            },
            LibraryLoader::GetModuleHandleW,
            SystemServices::SFGAO_FILESYSTEM,
        },
        UI::{
            Controls::*,
            Shell::{
                Common::ITEMIDLIST,
                SHBindToParent, SHParseDisplayName, CMINVOKECOMMANDINFO, IContextMenu,
                IShellFolder, CMF_NORMAL,
            },
            WindowsAndMessaging::*,
        },
    },
};

const ID_LISTVIEW: isize = 1000;
// カスタムメニューアイテムのコマンドID
const IDM_CUSTOM_COMMAND: u32 = 0x8000;

fn main() -> Result<()> {
    // 1. COMライブラリの初期化
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
    }

    let instance = unsafe { GetModuleHandleW(None)? };
    let window_class_name = w!("ExplorerContextMenuSample");

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
            w!("Rust Explorer Context Menu"),
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
            // 警告を解消するために戻り値を無視
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
                    if nmhdr.code == NM_RCLICK {
                        let nmitem: &NMITEMACTIVATE = &*(lparam.0 as *const NMITEMACTIVATE);
                        if nmitem.iItem != -1 {
                            show_context_menu(
                                window,
                                nmhdr.hwndFrom,
                                nmitem.iItem,
                                nmitem.ptAction,
                            );
                        }
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
        // 警告を解消するために戻り値を無視
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

// コンテキストメニューを表示するメインの関数
fn show_context_menu(owner: HWND, listview_hwnd: HWND, item_index: i32, point: POINT) {
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
    let current_dir = std::env::current_dir().unwrap();
    let full_path = current_dir.join(&file_name);

    // _pidl_absolute で所有権を持つことで、pidl_relative が指すメモリの生存期間を保証する
    if let Ok((shell_folder, _pidl_absolute, pidl_relative)) = get_shell_folder_and_pidl(&full_path) {
        let context_menu: Result<IContextMenu> =
            unsafe { shell_folder.GetUIObjectOf(owner, &[pidl_relative], None) };

        if let Ok(context_menu) = context_menu {
            let hmenu = unsafe { CreatePopupMenu().unwrap() };

            if unsafe { context_menu.QueryContextMenu(hmenu, 0, 1, 0x7FFF, CMF_NORMAL) }.is_ok() {
                unsafe {
                    let item_count = GetMenuItemCount(Some(hmenu));
                    InsertMenuW(
                        hmenu,
                        item_count as u32,
                        MF_BYPOSITION | MF_SEPARATOR,
                        0,
                        None,
                    )
                    .unwrap();

                    let menu_text: Vec<u16> =
                        OsStr::new("カスタムアクション(&C)").encode_wide().chain(Some(0)).collect();
                    InsertMenuW(
                        hmenu,
                        (item_count + 1) as u32,
                        MF_BYPOSITION | MF_STRING,
                        IDM_CUSTOM_COMMAND as usize,
                        PCWSTR(menu_text.as_ptr()),
                    )
                    .unwrap();
                }

                let mut pt = point;
                unsafe { ClientToScreen(listview_hwnd, &mut pt).unwrap() };

                let cmd = unsafe {
                    TrackPopupMenuEx(hmenu, TPM_RETURNCMD.0, pt.x, pt.y, owner, None)
                };

                let cmd_u32 = cmd.0 as u32;
                if cmd_u32 > 0 {
                    if cmd_u32 == IDM_CUSTOM_COMMAND {
                        let message =
                            format!("カスタムアクションがファイル「{}」に対して実行されました。", file_name);
                        let title = "カスタムコマンド";
                        let message_w: Vec<u16> =
                            OsStr::new(&message).encode_wide().chain(Some(0)).collect();
                        let title_w: Vec<u16> =
                            OsStr::new(title).encode_wide().chain(Some(0)).collect();

                        unsafe {
                            MessageBoxW(
                                Some(owner),
                                PCWSTR(message_w.as_ptr()),
                                PCWSTR(title_w.as_ptr()),
                                MB_OK | MB_ICONINFORMATION,
                            );
                        }
                    } else if cmd_u32 <= 0x7FFF {
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
            }
            unsafe { DestroyMenu(hmenu).unwrap() };
        }
    }
}

// 戻り値を変更: 所有権を持つ絶対PIDLと、それへのポインタである相対PIDLを返す
fn get_shell_folder_and_pidl(path: &Path) -> Result<(IShellFolder, OwningPidl, *const ITEMIDLIST)> {
    let path_wide: Vec<u16> = path.as_os_str().encode_wide().chain(Some(0)).collect();
    let mut pidl_absolute = OwningPidl::new();

    unsafe {
        let sfgao: u32 = SFGAO_FILESYSTEM.0;
        SHParseDisplayName(
            PCWSTR(path_wide.as_ptr()),
            None,
            pidl_absolute.as_mut_ptr(),
            sfgao,
            None,
        )?;
    }

    // SHBindToParent が書き込むポインタなので *mut ITEMIDLIST に変更
    let mut pidl_relative_ptr: *mut ITEMIDLIST = std::ptr::null_mut();

    let shell_folder: IShellFolder = unsafe {
        SHBindToParent(
            pidl_absolute.as_ptr(),
            Some(&mut pidl_relative_ptr),
        )?
    };

    Ok((shell_folder, pidl_absolute, pidl_relative_ptr))
}

// メモリ解放の責務を持つことを明確にするために名前を変更
struct OwningPidl {
    ptr: *mut ITEMIDLIST,
}

impl OwningPidl {
    fn new() -> Self {
        Self {
            ptr: std::ptr::null_mut(),
        }
    }

    fn as_ptr(&self) -> *const ITEMIDLIST {
        self.ptr
    }

    fn as_mut_ptr(&mut self) -> *mut *mut ITEMIDLIST {
        &mut self.ptr
    }
}

impl Drop for OwningPidl {
    fn drop(&mut self) {
        if !self.ptr.is_null() {
            unsafe { CoTaskMemFree(Some(self.ptr as *const _)) };
        }
    }
}
