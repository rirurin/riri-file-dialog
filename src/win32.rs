use std::{
    error::Error,
    path::{ Path, PathBuf },
    sync::{ Mutex, MutexGuard },
};
use windows::{
    core::{
        Error as WinError,
        PCWSTR,
    },
    Win32::{
        Foundation::HWND,
        System::Com::{
            CoCreateInstance,
            CoTaskMemFree,
            CLSCTX_ALL
        },
        UI::Shell::{
            Common::COMDLG_FILTERSPEC,
            FOS_PICKFOLDERS,
            FileOpenDialog,
            FileSaveDialog,
            IFileOpenDialog,
            IFileSaveDialog,
            IShellItem,
            SIGDN_FILESYSPATH,
            SHCreateItemFromParsingName
        }
    }
};

#[derive(Debug)]
pub struct FileTypeFilter {
    extension: String,
    description: String
}

impl FileTypeFilter {
    pub const fn new(extension: String, description: String) -> Self {
        Self { extension, description }
    }

    pub fn get_extension(&self) -> &str { &self.extension }
    pub fn get_description(&self) -> &str { &self.description }
}

#[derive(Debug)]
pub struct FileTypeFilterWin32 {
    extension: Vec<u16>,
    description: Vec<u16>
}

impl FileTypeFilterWin32 {
    pub fn new(extension: &str, description: &str) -> Self {
        let extension = format!("*.{}", extension);
        let extension = FileDialogUtils::to_win32_wide(&extension);
        let description = FileDialogUtils::to_win32_wide(description);
        Self { extension, description }
    }

    pub fn get_extension(&self) -> PCWSTR { PCWSTR(self.extension.as_ptr()) }
    pub fn get_description(&self) -> PCWSTR { PCWSTR(self.description.as_ptr()) }
}

#[derive(Debug)]
pub struct FileDialogManager {
    // see https://learn.microsoft.com/en-us/windows/win32/shell/common-file-dialog#controlling-the-default-folder
    default: PathBuf,
    window: HWND
}

unsafe impl Send for FileDialogManager {}
unsafe impl Sync for FileDialogManager {}

pub(crate) static FILE_DIALOG_MANAGER: Mutex<Option<FileDialogManager>> = Mutex::new(None);
type MgrBorrow = MutexGuard<'static, Option<FileDialogManager>>;

impl FileDialogManager {
    pub fn new(default: PathBuf, window: HWND) {
        let mut lock_dlg = FILE_DIALOG_MANAGER.lock().unwrap();
        *lock_dlg = Some(Self { default, window })

    }

    pub fn get() -> MgrBorrow {
        Self::try_get().unwrap()
    }

    pub fn try_get() -> Option<MgrBorrow> {
        let file_dlg = FILE_DIALOG_MANAGER.lock().unwrap();
        match file_dlg.as_ref().is_some() {
            true => Some(file_dlg),
            false => None
        }
    }

    pub fn get_or_set(default: PathBuf, window: HWND) -> MgrBorrow {
        Self::try_get().unwrap_or_else(|| {
            Self::new(default, window);
            Self::get()
        })
    }

    pub fn get_default_open(&self) -> &Path { self.default.as_path() }
    pub fn get_default_save(&self) -> &Path { self.default.as_path() }
    pub fn set_default_open<P>(&mut self, value: P) where P: AsRef<Path> { self.default = value.as_ref().to_owned() }
    pub fn set_default_save<P>(&mut self, value: P) where P: AsRef<Path> { self.default = value.as_ref().to_owned() }
    pub fn get_window_handle(&self) -> HWND { self.window }
}

pub trait FileDialog {
    fn get_default_title(&self) -> &'static str;
    fn get_title(&self, title: Option<&str>) -> Vec<u16> {
        match title {
            Some(v) => FileDialogUtils::to_win32_wide(v),
            None => FileDialogUtils::to_win32_wide(self.get_default_title())
        }
    }
    fn get_default_path(&self) -> &Path;
    fn set_default_path<P>(&mut self, file: P) where P: AsRef<Path>;
    fn get_window_handle(&self) -> HWND;
}
pub struct FileDialogUtils;
impl FileDialogUtils {
    pub(crate) fn to_win32_wide(s: &str) -> Vec<u16> {
        let mut alloc = Vec::with_capacity(s.len() + 1);
        alloc.extend(s.encode_utf16());
        alloc.push(0); // add null terminator
        alloc
    }
}
pub struct OpenDialog<'a> {
    manager: &'a mut FileDialogManager,
    handle: IFileOpenDialog
}
impl<'a> FileDialog for OpenDialog<'a> {
    fn get_default_title(&self) -> &'static str {
        "Open a file"
    }

    fn get_default_path(&self) -> &Path {
        self.manager.get_default_open()
    }

    fn set_default_path<P>(&mut self, file: P) where P: AsRef<Path> {
        self.manager.set_default_open(file)
    }

    fn get_window_handle(&self) -> HWND {
        self.manager.get_window_handle()
    }
}

impl<'a> OpenDialog<'a> {
    pub fn new(manager: &'a mut FileDialogManager) -> Result<Self, Box<dyn Error>> {
        Ok(Self {
            manager,
            handle: unsafe { CoCreateInstance(&FileOpenDialog, None, CLSCTX_ALL)? }
        })
    }

    fn open_inner(&mut self, title: Option<&str>) -> Result<Option<PathBuf>, WinError> {
        // Window Title
        let title = self.get_title(title);
        unsafe { self.handle.SetTitle(PCWSTR(title.as_ptr()))? }
        // Default folder
        let default_folder = FileDialogUtils::to_win32_wide(self.get_default_path().to_str().unwrap());
        let item: IShellItem = unsafe { SHCreateItemFromParsingName(PCWSTR(default_folder.as_ptr()), None)? };
        unsafe { self.handle.SetDefaultFolder(&item)? };
        // Run open dialog
        if unsafe { self.handle.Show(Some(self.get_window_handle())).is_ok() } {
            let res = unsafe { self.handle.GetResult()? };
            let path = unsafe { res.GetDisplayName(SIGDN_FILESYSPATH)? };
            let out = PathBuf::from(unsafe { path.to_string()? });
            self.set_default_path(out.as_path());
            unsafe { CoTaskMemFree(Some(path.0 as _)) }
            Ok(Some(out))
        } else {
            Ok(None)
        }
    }

    pub fn open(&mut self, filter: Option<&[FileTypeFilter]>, title: Option<&str>) -> Result<Option<PathBuf>, WinError> {
        // Provide owned allocation for file type strings
        let filter_platform: Option<Vec<FileTypeFilterWin32>> = filter.map(|filter| {
            filter.iter().map(|v| FileTypeFilterWin32::new(v.get_extension(), v.get_description())).collect()
        });
        if let Some(f) = filter_platform {
            let types: Vec<COMDLG_FILTERSPEC> = f.iter().map(|v| COMDLG_FILTERSPEC {
                pszName: v.get_description(),
                pszSpec: v.get_extension()
            }).collect();
            unsafe { self.handle.SetFileTypes(types.as_slice())? };
        }
        self.open_inner(title)
    }

    pub fn open_folder(&mut self, title: Option<&str>) -> Result<Option<PathBuf>, WinError> {
        let options = unsafe { self.handle.GetOptions()? };
        unsafe { self.handle.SetOptions(options | FOS_PICKFOLDERS)? };
        self.open_inner(title)
    }
}

pub struct SaveDialog<'a> {
    manager: &'a mut FileDialogManager,
    handle: IFileSaveDialog
}
impl<'a> FileDialog for SaveDialog<'a> {
    fn get_default_title(&self) -> &'static str {
        "Save a file"
    }

    fn get_default_path(&self) -> &Path {
        self.manager.get_default_save()
    }

    fn set_default_path<P>(&mut self, file: P) where P: AsRef<Path> {
        self.manager.set_default_save(file);
    }

    fn get_window_handle(&self) -> HWND {
        self.manager.get_window_handle()
    }
}

impl<'a> SaveDialog<'a> {
    pub fn new(manager: &'a mut FileDialogManager) -> Result<Self, Box<dyn Error>> {
        Ok(Self {
            manager,
            handle: unsafe { CoCreateInstance(&FileSaveDialog, None, CLSCTX_ALL)? }
        })
    }

    pub fn save(&mut self, filter: Option<&[FileTypeFilter]>, title: Option<&str>) -> Result<Option<PathBuf>, WinError> {
        // Provide owned allocation for file type strings
        let filter_platform: Option<Vec<FileTypeFilterWin32>> = filter.map(|filter| {
            filter.iter().map(|v| FileTypeFilterWin32::new(v.get_extension(), v.get_description())).collect()
        });
        if let Some(f) = filter_platform {
            let types: Vec<COMDLG_FILTERSPEC> = f.iter().map(|v| COMDLG_FILTERSPEC {
                pszName: v.get_description(),
                pszSpec: v.get_extension()
            }).collect();
            unsafe { self.handle.SetFileTypes(types.as_slice())? };
        }
        // Window Title
        let title = self.get_title(title);
        unsafe { self.handle.SetTitle(PCWSTR(title.as_ptr()))? }
        // Default folder
        let default_folder = FileDialogUtils::to_win32_wide(self.get_default_path().to_str().unwrap());
        let item: IShellItem = unsafe { SHCreateItemFromParsingName(PCWSTR(default_folder.as_ptr()), None)? };
        unsafe { self.handle.SetDefaultFolder(&item)? };
        // Run open dialog
        if unsafe { self.handle.Show(Some(self.get_window_handle())).is_ok() } {
            let res = unsafe { self.handle.GetResult()? };
            let path = unsafe { res.GetDisplayName(SIGDN_FILESYSPATH)? };
            let out =  PathBuf::from(unsafe { path.to_string()? });
            self.set_default_path(out.clone());
            unsafe { CoTaskMemFree(Some(path.0 as _)) }
            Ok(Some(out))
        } else {
            Ok(None)
        }
    }
}