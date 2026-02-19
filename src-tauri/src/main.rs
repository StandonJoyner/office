#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod commands;
mod error;
mod file_watcher;
mod filesystem;
mod storage;
mod system;

use tauri::Manager;

#[tokio::main]
async fn main() {
    // Initialize tracing
    #[cfg(debug_assertions)]
    tracing_subscriber::fmt()
        .with_max_level(tracing::Level::DEBUG)
        .init();

    #[cfg(not(debug_assertions))]
    tracing_subscriber::fmt()
        .with_max_level(tracing::Level::INFO)
        .init();

    tracing::info!("Starting Office Suite Tauri application");

    tauri::Builder::default()
        .plugin(tauri_plugin_shell::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_notification::init())
        .plugin(tauri_plugin_clipboard_manager::init())
        .setup(|app| {
            // Initialize file watcher
            let app_handle = app.handle().clone();
            tokio::spawn(async move {
                file_watcher::init_file_watcher(app_handle).await;
            });

            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            // File system commands
            commands::filesystem::read_file,
            commands::filesystem::read_text_file,
            commands::filesystem::write_file,
            commands::filesystem::write_text_file,
            commands::filesystem::remove_file,
            commands::filesystem::exists,
            commands::filesystem::read_dir,
            commands::filesystem::create_dir,
            commands::filesystem::remove_dir,
            // Office file commands
            commands::filesystem::parse_docx,
            commands::filesystem::parse_xlsx,
            commands::filesystem::save_docx,
            commands::filesystem::save_xlsx,
            // File watching
            commands::filesystem::watch_file,
            commands::filesystem::watch_directory,
            commands::filesystem::unwatch,
            // System commands
            commands::system::show_menu,
            commands::system::create_menu,
            commands::system::prompt,
            commands::system::get_theme,
            commands::system::listen_theme_change,
            commands::system::enable_drag_drop,
            commands::system::disable_drag_drop,
            commands::system::listen_notification,
            // Storage commands
            commands::storage::get_all_keys,
            commands::storage::get_storage_size,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
