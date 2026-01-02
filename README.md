# Golden Copy Manager

A lightweight, cross-platform clipboard manager built with Python and Tkinter.

Golden Copy Manager runs quietly in the background, records your clipboard history, lets you **pin important items**, and gives you quick access to past copies without clutter or distractions.

---

## ✨ Features

- 📋 **Clipboard history**
  - Automatically records copied text
  - Each entry includes a **timestamp**
- 📌 **Pin / Favorite**
  - Pin important items so they’re never removed
- ⏸ **Pause / Resume monitoring**
  - Temporarily stop capturing clipboard content
- 🗂 **History limit**
  - Unpinned items are capped (oldest removed first)
- 🖱 **Right-click context menu**
  - Copy, pin/unpin, delete
- 🖥 **System tray support (Windows)**
  - Hide to tray
  - Show / hide window
  - Pause / resume
  - Exit
- ⌨ **Global hotkey**
  - Works on Windows and macOS
- 🚀 **Optional start with Windows**
- 💾 **Persistent storage**
  - Clipboard history and settings saved locally

---

## 🧰 Tech Stack

- Python 3.9+
- Tkinter
- Pillow
- pystray (Windows tray)
- pynput (global hotkeys)

---

## 📂 Project Structure

```
Golden-Copy-Manager/
│
├─ Golden_Copy_Manager_.py
├─ assets/
│   ├─ icon.ico
│   └─ background.png
│
├─ data/
│   ├─ clipboard_history.json
│   └─ clipboard_manager_settings.json
│
└─ README.md
```

---

## ▶ Run from source

```bash
pip install pillow pystray pynput
python Golden_Copy_Manager.py
```

---

## 🏗 Build Windows .exe

```powershell
py -m pip install pyinstaller
py -m PyInstaller --noconsole --onedir --clean --name GoldenCopyManager `
  --icon .\assets\icon.ico `
  --add-data 'assets;assets' `
  --add-data 'data;data' `
  .\Golden_Copy_Manager.py
```

Output:
```
dist/GoldenCopyManager/GoldenCopyManager.exe
```

Zip the entire folder to share.

---

## 🔐 Privacy

- Runs locally only
- No network access
- Clipboard data never leaves your computer


