use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::{Arc, Mutex};
use notify::{RecommendedWatcher, Watcher, RecursiveMode, Event, EventKind, Config};
use tokio::sync::mpsc;
use tauri::{AppHandle, Emitter};
use crate::error::Result;

type WatchCallback = Arc<Mutex<Box<dyn Fn(&str, EventKind) + Send>>;

pub struct FileWatcher {
    watcher: RecommendedWatcher,
    tx: mpsc::Sender<FileWatcherEvent>,
    callbacks: Arc<Mutex<HashMap<String, Vec<WatchCallback>>>>,
}

impl FileWatcher {
    pub fn new(tx: mpsc::Sender<FileWatcherEvent>) -> Result<Self> {
        let callbacks = Arc::new(Mutex::new(HashMap::new()));

        let callbacks_clone = callbacks.clone();
        let watcher = notify::recommended_watcher(move |res: Result<Event, _>| {
            if let Ok(event) = res {
                if let Some(path) = event.paths.first() {
                    if let Some(path_str) = path.to_str() {
                        let callbacks = callbacks_clone.lock().unwrap();
                        // Find all callbacks that match this path
                        for (watch_path, cb_list) in callbacks.iter() {
                            if path_str.starts_with(watch_path) {
                                for cb in cb_list.iter() {
                                    cb(path_str, event.kind.clone());
                                }
                            }
                        }
                    }
                }
            }
        })?;

        Ok(Self {
            watcher,
            tx,
            callbacks,
        })
    }

    pub fn watch(&mut self, path: String, watch_id: String, app_handle: AppHandle) -> Result<()> {
        let watch_id_clone = watch_id.clone();
        let app_handle_clone = app_handle.clone();

        let callback: WatchCallback = Arc::new(Mutex::new(Box::new(move |path: &str, kind: EventKind| {
            let watch_id = watch_id_clone.clone();
            let app_handle = app_handle_clone.clone();

            let event_type = match kind {
                EventKind::Create(_) => "created",
                EventKind::Modify(ModifyKind::Data(_) | ModifyKind::Any) => "modified",
                EventKind::Remove(_) => "deleted",
                _ => "other",
            };

            let _ = app_handle.emit("file-watch-event", FileWatchPayload {
                watch_id,
                path: path.to_string(),
                event_type: event_type.to_string(),
                timestamp: std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .unwrap()
                    .as_millis(),
            });
        }));

        // Add callback
        self.callbacks.lock().unwrap()
            .entry(path.clone())
            .or_insert_with(Vec::new)
            .push(callback);

        // Start actual watching
        self.watcher.watch(&PathBuf::from(&path), RecursiveMode::Recursive)?;
        Ok(())
    }

    pub fn unwatch(&mut self, path: &str) -> Result<()> {
        self.watcher.unwatch(PathBuf::from(path))?;
        self.callbacks.lock().unwrap().remove(path);
        Ok(())
    }
}

#[derive(serde::Serialize)]
pub struct FileWatcherEvent {
    pub watch_id: String,
    pub path: String,
}

#[derive(serde::Serialize)]
pub struct FileWatchPayload {
    pub watch_id: String,
    pub path: String,
    pub event_type: String,
    pub timestamp: u128,
}

pub async fn init_file_watcher(app_handle: tauri::AppHandle) {
    let (tx, mut rx) = mpsc::channel::<FileWatcherEvent>(100);

    // TODO: Store watcher in app state for proper lifecycle management
    if let Ok(mut watcher) = FileWatcher::new(tx) {
        // Watcher would be registered here
        tracing::info!("File watcher initialized");
    }
}
