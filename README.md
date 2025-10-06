# riri-file-dialog

A small utility crate for creating OS save/load dialogs.
This has been designed for personal needs, so it currently only supports Windows.

## Usage

Initializing `FileDialogManager` instance (should be done at startup):

```rust
// window handle uses Win32's HWND on Windows
FileDialogManager::new(default_path, hwnd);
```

Opening a single file using `OpenDialog`:

```rust
let mut dlg_lock = FileDialogManager::get();
if let Some(path) = OpenDialog::new(dlg_lock.as_mut().unwrap()).unwrap().open(
    Some(&[FileTypeFilter::new("p5path".to_owned(), "P5R Freecam Path".to_owned())]), // defines what extensions are supported
    Some("Open camera path") // dialog title
).unwrap() {
    // Code to handle selected file...
}
```

Saving a single file using `SaveDialog`:

```rust
let mut dlg_lock = FileDialogManager::get();
if let Some(path) = SaveDialog::new(dlg_lock.as_mut().unwrap()).unwrap().save(
    Some(&[FileTypeFilter::new("p5path".to_owned(), "P5R Freecam Path".to_owned())]), // defines what extensions are supported
    Some("Save camera path") // dialog title
).unwrap() {
    // Code to handle selected file...
}
```