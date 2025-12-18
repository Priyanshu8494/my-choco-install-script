# 🚀 Ultimate PC Setup Toolkit

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?style=for-the-badge&logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D6?style=for-the-badge&logo=windows)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

**The only script you need to set up a new Windows PC.**  
Automated. Powerful. Self-Contained.

---

## 🔥 One-Line Install (The "Magic" Command)
Right-click your **PowerShell (Admin)** and paste this:

```powershell
irm https://raw.githubusercontent.com/Priyanshu8494/my-choco-install-script/main/setup-new-pc.ps1 | iex
```

*(If you have connection issues, see the [Secure Install](#-secure-installation-method) section below.)*

---

## ⚡ Features

### 🛠️ 1. Essential Software Installer
Skip the manual downloads. Installs the best utilities instantly:
- **Browsers:** Chrome, Firefox
- **Tools:** WinRAR, VLC, SumatraPDF
- **Remote:** AnyDesk, UltraViewer

### 💼 2. Microsoft Office Suite (Pro Plus 2021)
- **Online:** Downloads directly from the official CDN.
- **Offline:** Create your own reusable offline installer folder for air-gapped PCs.

### 🚀 3. Global RAM Optimizer (Embedded) **[NEW!]**
A powerful memory management engine **built directly into the script**.
- **Smart Mode:** Automatically detects apps using >100MB RAM and trims their working set.
- **Zero Dependencies:** No external files needed. It extracts itself from memory!
- **Safe:** Whitelists critical Windows system processes to prevent crashes.

### 🔑 4. System Activation
- Integrated activation toolkit for Windows and Office.
- Simple, one-click menu usage.

### 🧰 5. Advanced Toolkit (WinUtil)
- Integration with Chris Titus Tech's WinUtil for deep system debloating and tweaking.

---

## 🛡️ Secure Installation Method
If your antivirus or firewall blocks the "One-Line" command, use this safe 2-step method:

```powershell
# 1. Download the script securely to disk
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Priyanshu8494/my-choco-install-script/main/setup-new-pc.ps1" -OutFile "$env:TEMP\setup.ps1"

# 2. Run the file
& "$env:TEMP\setup.ps1"
```

---

## 📸 Screenshots
*(Add your screenshots here)*

---

## 👨‍💻 Created By
**Priyanshu Suryavanshi**  
*Building tools to make IT simple.*

---
*Disclaimer: This script is provided as-is. Always verify scripts before running them with administrative privileges.*
