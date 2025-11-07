# 🎯 Quick Start - Building Your Executable

## For Developers

### Build the Executable (3 Steps)

1. **Open PowerShell/CMD in project folder**

2. **Run the quick build script:**

   ```batch
   quick_build.bat
   ```

3. **Get your executable:**
   - Find it in: `release\LogLife_Operacao.exe`

### That's it! 🎉

---

## Update System Overview

### ✅ What's Already Configured

Your application now has:

1. **Automatic Update Checking**

   - Checks on startup (silent)
   - Manual check via "Sobre" tab
   - Compares versions with GitHub

2. **User-Friendly Dialogs**

   - Shows update notifications
   - Opens browser to download
   - Provides installation instructions

3. **Version Display**
   - Shown in window title
   - Shown in "Sobre" tab
   - Includes build date

### 🔄 How to Release a New Version

1. **Update version numbers:**

   - Edit `src\utils\version_manager.py` (lines 10-11)
   - Edit `version.json` (entire file)

2. **Build:**

   ```batch
   quick_build.bat
   ```

3. **Test the executable**

4. **Commit to Git:**

   ```bash
   git add .
   git commit -m "Release v1.1.0"
   git tag v1.1.0
   git push origin main --tags
   ```

5. **Create GitHub Release:**
   - Go to GitHub → Releases → New Release
   - Upload `release\LogLife_Operacao.exe`
   - The update system will now detect it!

---

## 📁 Important Files

| File                           | Purpose                  |
| ------------------------------ | ------------------------ |
| `main_ctk.spec`                | Build configuration      |
| `quick_build.bat`              | Easy build script        |
| `version.json`                 | Version info for updates |
| `src\utils\version_manager.py` | Update checker           |
| `BUILD_UPDATE_GUIDE.md`        | Detailed documentation   |
| `USER_GUIDE.md`                | For end users            |

---

## 🧪 Testing

### Test Version Manager:

```batch
python test_version_manager.py
```

### Test the Executable:

1. Build it: `quick_build.bat`
2. Run it: `release\LogLife_Operacao.exe`
3. Check "Sobre" tab
4. Click "Verificar Atualizações"

---

## 🎨 What Users See

### On Startup:

```
Window Title: "Operação LogLife - Versão 1.0.0 (Build: 2025-11-07)"
```

### If Update Available:

```
┌─────────────────────────────────────────┐
│     Nova versão disponível!             │
│                                         │
│  Versão atual: 1.0.0                    │
│  Nova versão: 1.1.0                     │
│                                         │
│  [Release notes here]                   │
│                                         │
│  Deseja baixar a atualização?           │
│                                         │
│         [Sim]    [Não]                  │
└─────────────────────────────────────────┘
```

---

## 🚨 Important Notes

### For the Update System to Work:

1. **`version.json` must be in GitHub repository root**
2. **Must be on `main` branch** (or update the URL)
3. **Must create GitHub Releases** for downloads
4. **Users need internet** to check for updates

### File Size:

- Executable will be ~200-300 MB
- This is normal for bundled Python apps
- Includes all dependencies

---

## 🆘 Quick Troubleshooting

**Build fails?**

- Run: `pip install -r requirements.txt`
- Run: `pip install pyinstaller`

**Icon missing?**

- Check: `src\assets\my_icon.ico` exists

**Update check fails?**

- Push `version.json` to GitHub
- Check internet connection

**Import errors in exe?**

- Add module to `hiddenimports` in `main_ctk.spec`

---

## 📞 Need Help?

1. Check `BUILD_UPDATE_GUIDE.md` for details
2. Run `test_version_manager.py` to diagnose
3. Check error logs in console/logs folder

---

**Happy Building! 🚀**
