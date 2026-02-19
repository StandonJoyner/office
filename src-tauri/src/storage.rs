use std::fs;
use std::path::PathBuf;
use crate::error::Result;
use crate::error::AppError;

pub struct Storage {
    base_path: PathBuf,
}

impl Storage {
    pub fn new(base_path: PathBuf) -> Self {
        Self { base_path }
    }

    fn get_key_path(&self, key: &str) -> PathBuf {
        self.base_path.join(format!("{}.json", key))
    }

    pub fn get(&self, key: &str) -> Result<Option<String>> {
        let path = self.get_key_path(key);
        if !path.exists() {
            return Ok(None);
        }
        fs::read_to_string(path).map(Some).map_err(AppError::from)
    }

    pub fn set(&self, key: &str, value: &str) -> Result<()> {
        let path = self.get_key_path(key);
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)?;
        }
        fs::write(path, value).map_err(AppError::from)
    }

    pub fn remove(&self, key: &str) -> Result<()> {
        let path = self.get_key_path(key);
        if path.exists() {
            fs::remove_file(path).map_err(AppError::from)
        } else {
            Ok(())
        }
    }

    pub fn exists(&self, key: &str) -> bool {
        self.get_key_path(key).exists()
    }

    pub fn list_keys(&self) -> Result<Vec<String>> {
        let mut keys = Vec::new();
        if !self.base_path.exists() {
            return Ok(keys);
        }

        for entry in fs::read_dir(&self.base_path)? {
            let entry = entry?;
            let path = entry.path();
            if path.is_file() && path.extension().map_or(false, |e| e == "json") {
                if let Some(name) = path.file_stem().and_then(|s| s.to_str()) {
                    keys.push(name.to_string());
                }
            }
        }
        Ok(keys)
    }

    pub fn size(&self) -> Result<u64> {
        let mut total = 0u64;
        if !self.base_path.exists() {
            return Ok(total);
        }

        for entry in fs::read_dir(&self.base_path)? {
            let entry = entry?;
            let path = entry.path();
            if path.is_file() {
                total += entry.metadata()?.len();
            }
        }
        Ok(total)
    }

    pub fn clear(&self) -> Result<()> {
        if !self.base_path.exists() {
            return Ok(());
        }
        fs::remove_dir_all(&self.base_path)?;
        fs::create_dir_all(&self.base_path)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn test_storage_basic() {
        let temp = TempDir::new().unwrap();
        let storage = Storage::new(temp.path().to_path_buf());

        assert!(!storage.exists("test"));
        assert_eq!(storage.get("test").unwrap(), None);

        storage.set("test", "value").unwrap();
        assert!(storage.exists("test"));
        assert_eq!(storage.get("test").unwrap(), Some("value".to_string()));

        storage.remove("test").unwrap();
        assert!(!storage.exists("test"));
    }

    #[test]
    fn test_storage_list() {
        let temp = TempDir::new().unwrap();
        let storage = Storage::new(temp.path().to_path_buf());

        storage.set("key1", "value1").unwrap();
        storage.set("key2", "value2").unwrap();

        let keys = storage.list_keys().unwrap();
        assert_eq!(keys.len(), 2);
        assert!(keys.contains(&"key1".to_string()));
        assert!(keys.contains(&"key2".to_string()));
    }
}
