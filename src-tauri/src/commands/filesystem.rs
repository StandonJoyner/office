use crate::error::Result;
use crate::filesystem;

#[tauri::command]
pub async fn read_file(path: String) -> Result<Vec<u8>> {
    filesystem::read_file(&path)
}

#[tauri::command]
pub async fn read_text_file(path: String) -> Result<String> {
    filesystem::read_text_file(&path)
}

#[tauri::command]
pub async fn write_file(path: String, contents: Vec<u8>) -> Result<()> {
    filesystem::write_file(&path, &contents)
}

#[tauri::command]
pub async fn write_text_file(path: String, contents: String) -> Result<()> {
    filesystem::write_text_file(&path, &contents)
}

#[tauri::command]
pub async fn remove_file(path: String) -> Result<()> {
    filesystem::remove_file(&path)
}

#[tauri::command]
pub async fn exists(path: String) -> bool {
    filesystem::exists(&path)
}

#[tauri::command]
pub async fn read_dir(path: String, recursive: bool) -> Result<Vec<filesystem::FileEntry>> {
    filesystem::read_dir(&path, recursive)
}

#[tauri::command]
pub async fn create_dir(path: String, recursive: bool) -> Result<()> {
    filesystem::create_dir(&path, recursive)
}

#[tauri::command]
pub async fn remove_dir(path: String, recursive: bool) -> Result<()> {
    filesystem::remove_dir(&path, recursive)
}

// Office file commands
#[tauri::command]
pub async fn parse_docx(path: String) -> Result<filesystem::DocxContent> {
    filesystem::parse_docx(&path)
}

#[tauri::command]
pub async fn parse_xlsx(path: String) -> Result<filesystem::XlsxContent> {
    filesystem::parse_xlsx(&path)
}

#[tauri::command]
pub async fn save_docx(path: String, content: filesystem::DocxContent) -> Result<()> {
    filesystem::save_docx(&path, content)
}

#[tauri::command]
pub async fn save_xlsx(path: String, content: filesystem::XlsxContent) -> Result<()> {
    filesystem::save_xlsx(&path, content)
}

// File watching
#[tauri::command]
pub async fn watch_file(path: String, watch_id: String, app: tauri::AppHandle) -> Result<()> {
    // TODO: Implement file watching
    tracing::info!("Watching file: {}", path);
    Ok(())
}

#[tauri::command]
pub async fn watch_directory(path: String, watch_id: String, app: tauri::AppHandle) -> Result<()> {
    // TODO: Implement directory watching
    tracing::info!("Watching directory: {}", path);
    Ok(())
}

#[tauri::command]
pub async fn unwatch(watch_id: String) -> Result<()> {
    // TODO: Implement unwatching
    tracing::info!("Unwatching: {}", watch_id);
    Ok(())
}
