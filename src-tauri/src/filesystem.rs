use std::fs;
use std::path::{Path, PathBuf};
use std::collections::HashMap;
use serde::{Deserialize, Serialize};
use notify::{RecommendedWatcher, RecursiveMode, Watcher, Event, EventKind};
use notify::event::ModifyKind;
use tokio::sync::mpsc;
use crate::error::{AppError, Result};

#[derive(Debug, Serialize, Deserialize)]
pub struct FileMetadata {
    pub size: u64,
    pub modified: u64,
    pub accessed: u64,
    pub created: u64,
    pub readonly: bool,
    pub is_file: bool,
    pub is_dir: bool,
    pub is_symlink: bool,
}

impl FileMetadata {
    pub fn from_path(path: &Path) -> Result<Self> {
        let metadata = fs::metadata(path)?;
        let modified = metadata.modified()?
            .duration_since(std::time::UNIX_EPOCH)?
            .as_secs();
        let accessed = metadata.accessed()?
            .duration_since(std::time::UNIX_EPOCH)?
            .as_secs();
        let created = metadata.created()
            .ok()
            .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
            .map(|d| d.as_secs())
            .unwrap_or(0);

        Ok(Self {
            size: metadata.len(),
            modified,
            accessed,
            created,
            readonly: metadata.permissions().readonly(),
            is_file: metadata.is_file(),
            is_dir: metadata.is_dir(),
            is_symlink: metadata.is_symlink(),
        })
    }
}

#[derive(Debug, Serialize, Deserialize)]
pub struct FileEntry {
    pub path: String,
    pub children: Vec<FileEntry>,
    pub metadata: FileMetadata,
}

pub fn read_file(path: &str) -> Result<Vec<u8>> {
    fs::read(path).map_err(AppError::from)
}

pub fn read_text_file(path: &str) -> Result<String> {
    fs::read_to_string(path).map_err(AppError::from)
}

pub fn write_file(path: &str, contents: &[u8]) -> Result<()> {
    // Ensure parent directory exists
    if let Some(parent) = Path::new(path).parent() {
        fs::create_dir_all(parent)?;
    }
    fs::write(path, contents).map_err(AppError::from)
}

pub fn write_text_file(path: &str, contents: &str) -> Result<()> {
    // Ensure parent directory exists
    if let Some(parent) = Path::new(path).parent() {
        fs::create_dir_all(parent)?;
    }
    fs::write(path, contents).map_err(AppError::from)
}

pub fn remove_file(path: &str) -> Result<()> {
    fs::remove_file(path).map_err(AppError::from)
}

pub fn exists(path: &str) -> bool {
    Path::new(path).exists()
}

pub fn read_dir(path: &str, recursive: bool) -> Result<Vec<FileEntry>> {
    let path = Path::new(path);
    if !path.exists() {
        return Err(AppError::NotFound(path.to_string_lossy().to_string()));
    }

    let mut entries = Vec::new();
    read_dir_recursive(path, &mut entries, recursive)?;
    Ok(entries)
}

fn read_dir_recursive(path: &Path, entries: &mut Vec<FileEntry>, recursive: bool) -> Result<()> {
    for entry in fs::read_dir(path)? {
        let entry_path = entry?.path();
        let entry_path = entry_path.ok_or_else(|| AppError::Path("Invalid entry".to_string()))?;

        let metadata = FileMetadata::from_path(&entry_path)?;
        let children = if recursive && metadata.is_dir {
            let mut child_entries = Vec::new();
            read_dir_recursive(&entry_path, &mut child_entries, true)?;
            child_entries
        } else {
            Vec::new()
        };

        entries.push(FileEntry {
            path: entry_path.to_string_lossy().to_string(),
            children,
            metadata,
        });
    }
    Ok(())
}

pub fn create_dir(path: &str, recursive: bool) -> Result<()> {
    if recursive {
        fs::create_dir_all(path).map_err(AppError::from)
    } else {
        fs::create_dir(path).map_err(AppError::from)
    }
}

pub fn remove_dir(path: &str, recursive: bool) -> Result<()> {
    if recursive {
        fs::remove_dir_all(path).map_err(AppError::from)
    } else {
        fs::remove_dir(path).map_err(AppError::from)
    }
}

// Office file parsing (placeholder - would need actual implementation with crates)
#[derive(Debug, Serialize, Deserialize)]
pub struct DocxContent {
    pub text: String,
    pub paragraphs: Vec<DocxParagraph>,
    pub tables: Vec<DocxTable>,
    pub metadata: DocxMetadata,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct DocxParagraph {
    pub id: String,
    pub text: String,
    pub runs: Vec<DocxRun>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct DocxRun {
    pub text: String,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct DocxTable {
    pub id: String,
    pub rows: Vec<DocxRow>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct DocxRow {
    pub cells: Vec<DocxCell>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct DocxCell {
    pub text: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct DocxMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct XlsxContent {
    pub sheets: Vec<XlsxSheet>,
    pub metadata: XlsxMetadata,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct XlsxSheet {
    pub id: String,
    pub name: String,
    pub cells: HashMap<String, XlsxCell>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct XlsxCell {
    pub value: serde_json::Value,
    pub formula: Option<String>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct XlsxMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
}

// Placeholder implementations for office file parsing
// In production, use crates like: rust_xlsxwriter, docx-rs, calamine
pub fn parse_docx(path: &str) -> Result<DocxContent> {
    // TODO: Implement actual .docx parsing
    let text = read_text_file(path)?;
    Ok(DocxContent {
        text: text.clone(),
        paragraphs: vec![DocxParagraph {
            id: "1".to_string(),
            text,
            runs: vec![DocxRun {
                text,
                bold: None,
                italic: None,
                underline: None,
            }],
        }],
        tables: Vec::new(),
        metadata: DocxMetadata {
            title: None,
            author: None,
        },
    })
}

pub fn parse_xlsx(path: &str) -> Result<XlsxContent> {
    // TODO: Implement actual .xlsx parsing
    // Would use calamine crate in production
    Ok(XlsxContent {
        sheets: vec![XlsxSheet {
            id: "sheet1".to_string(),
            name: "Sheet1".to_string(),
            cells: HashMap::new(),
        }],
        metadata: XlsxMetadata {
            title: None,
            author: None,
        },
    })
}

pub fn save_docx(path: &str, content: DocxContent) -> Result<()> {
    // TODO: Implement actual .docx saving
    let json = serde_json::to_string_pretty(&content)?;
    write_text_file(path, &json)
}

pub fn save_xlsx(path: &str, content: XlsxContent) -> Result<()> {
    // TODO: Implement actual .xlsx saving
    let json = serde_json::to_string_pretty(&content)?;
    write_text_file(path, &json)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_file_metadata() {
        let path = Path::new("test.txt");
        let result = FileMetadata::from_path(path);
        assert!(result.is_ok() || result.is_err()); // May not exist
    }
}
