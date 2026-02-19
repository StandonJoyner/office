# Tauri Desktop Application

This directory contains the Rust backend for the Office Suite Tauri desktop application.

## Directory Structure

```
src-tauri/
├── src/
│   ├── main.rs              # Application entry point
│   ├── error.rs             # Error types
│   ├── filesystem.rs        # File system operations
│   ├── storage.rs           # Storage implementation
│   ├── system.rs            # System integrations
│   ├── file_watcher.rs      # File watching
│   └── commands/           # Tauri IPC commands
├── Cargo.toml              # Rust dependencies
├── tauri.conf.json        # Tauri configuration
└── build.rs               # Build script
```

## Development

### Prerequisites

- Rust 1.70+ with toolchain `stable`
- Node.js 18+
- npm or yarn

### Setting up development environment

1. Install Rust:
   ```bash
   curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
   ```

2. Install system dependencies:
   - **Ubuntu/Debian**: `sudo apt install libwebkit2gtk-4.0-dev libssl-dev libgtk-3-dev libayatana-appindicator3-dev librsvg2-dev`
   - **Fedora**: `sudo dnf install webkit2gtk3-devel openssl-devel gtk3-devel librsvg2-devel`
   - **macOS**: No additional dependencies needed
   - **Windows**: No additional dependencies needed (Visual Studio C++ Build Tools required)

3. Install Node dependencies (already done):
   ```bash
   npm install
   ```

### Running in development mode

```bash
npm run tauri:dev
```

This will:
1. Start the Next.js development server (http://localhost:3000)
2. Build and launch the Tauri application

### Building for production

```bash
npm run tauri:build
```

This will build the application for your current platform in `src-tauri/target/release/bundle/`.

## Available Tauri Commands

### File System Commands

- `read_file(path)` - Read file as bytes
- `read_text_file(path)` - Read file as text
- `write_file(path, contents)` - Write file as bytes
- `write_text_file(path, contents)` - Write file as text
- `remove_file(path)` - Delete file
- `exists(path)` - Check if file exists
- `read_dir(path, recursive)` - List directory contents
- `create_dir(path, recursive)` - Create directory
- `remove_dir(path, recursive)` - Delete directory

### Office File Commands

- `parse_docx(path)` - Parse .docx file
- `parse_xlsx(path)` - Parse .xlsx file
- `save_docx(path, content)` - Save .docx file
- `save_xlsx(path, content)` - Save .xlsx file

### File Watching Commands

- `watch_file(path, watch_id)` - Watch file for changes
- `watch_directory(path, watch_id)` - Watch directory for changes
- `unwatch(watch_id)` - Stop watching

### System Commands

- `get_theme()` - Get current system theme
- `listen_theme_change()` - Listen for theme changes
- `show_menu()` - Show native menu
- `create_menu(items)` - Create custom menu
- `prompt(message, default_text)` - Show prompt dialog
- `enable_drag_drop()` - Enable drag and drop
- `disable_drag_drop()` - Disable drag and drop

### Storage Commands

- `get_all_keys()` - Get all storage keys
- `get_storage_size()` - Get storage size

## Platform-Specific Notes

### Windows
- Requires Visual Studio C++ Build Tools
- Builds `.msi` and `.exe` installers

### macOS
- Requires Xcode Command Line Tools
- Builds `.app` bundle and `.dmg` disk image

### Linux
- Requires webkit2gtk and other dependencies
- Builds `.deb` and `.AppImage` packages

## File Associations

The application is configured to open:
- `.docx` files (Word documents)
- `.xlsx` files (Excel spreadsheets)

## Troubleshooting

### Build fails with "linking with cc failed"
Install the appropriate system dependencies for your platform.

### Tauri dev mode shows blank window
Make sure Next.js is running on http://localhost:3000

### File permissions error
Ensure the application has permissions to access the file system

## Adding New Commands

1. Create the command function in `src/commands/`:
   ```rust
   #[tauri::command]
   pub async fn my_command(arg: String) -> Result<String> {
       Ok(format!("Hello, {}!", arg))
   }
   ```

2. Register the command in `src/main.rs`:
   ```rust
   .invoke_handler(tauri::generate_handler![
       // ... existing commands
       commands::my_module::my_command,
   ])
   ```

3. Call from frontend:
   ```typescript
   import { invoke } from '@tauri-apps/api/core';
   const result = await invoke('my_command', { arg: 'World' });
   ```
