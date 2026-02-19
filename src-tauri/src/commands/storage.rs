use crate::error::Result;
use std::sync::Mutex;
use std::collections::HashMap;
use tauri::State;

pub struct StorageState(pub Mutex<HashMap<String, String>>);

#[tauri::command]
pub async fn get_all_keys(state: State<'_, StorageState>) -> Result<Vec<String>> {
    let storage = state.0.lock().unwrap();
    Ok(storage.keys().cloned().collect())
}

#[tauri::command]
pub async fn get_storage_size(state: State<'_, StorageState>) -> Result<usize> {
    let storage = state.0.lock().unwrap();
    Ok(storage.len())
}
