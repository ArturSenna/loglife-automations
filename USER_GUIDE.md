# 🚀 LogLife Operations - User Guide

## What is this?

**LogLife Operations** is an automation tool for managing logistics operations including:

- ✅ Service updates
- ✅ Delivery notes (Minutas)
- ✅ Load tracking (Relação de cargas)
- ✅ File organization

## 📥 Installation

**No installation required!** This is a portable application.

### Quick Start:

1. **Download** `LogLife_Operacao.exe`
2. **Place it** in any folder you want
3. **Double-click** to run

That's it! No Python, no dependencies, no hassle.

## 💻 System Requirements

- Windows 7, 8, 10, or 11 (64-bit)
- 4GB RAM (minimum)
- 500MB free disk space
- Internet connection (for updates and API access)

## 🔄 Automatic Updates

The application checks for updates automatically when you start it.

### How to check for updates manually:

1. Open the application
2. Click on the **"Sobre"** tab
3. Click **"Verificar Atualizações"** button

If an update is available:

- You'll see a notification
- Click "Yes" to open download page
- Download the new version
- Close the old app
- Replace the .exe file with the new one

## 🎯 Features

### Tab 1: Serviços (Services)

- Update service data by date range
- Filter by regional office
- Clear existing data

### Tab 2: Minutas (Delivery Notes)

- Issue individual delivery notes
- Bulk issue for air deliveries
- Select material type and service

### Tab 3: Relação de Cargas (Load Tracking)

- Update load tracking data
- Update Fleury spreadsheet

### Tab 4: Arquivos (Files)

- Configure service spreadsheet
- Configure load tracking spreadsheet
- Set delivery notes folder
- Set Fleury spreadsheet

### Tab 5: Sobre (About)

- Check current version
- Check for updates
- View app information

## ⚙️ Configuration

### First Time Setup:

1. **Go to the "Arquivos" tab**
2. **Select required files:**

   - Service spreadsheet (Planilha de Serviços)
   - Load tracking file (Relação de Cargas)
   - Delivery notes folder (Pasta das minutas)
   - Fleury spreadsheet (if applicable)

3. **The app will remember these settings** for next time

## 🐛 Troubleshooting

### Application won't start

**Problem**: Windows says "Windows protected your PC"

**Solution**:

1. Click "More info"
2. Click "Run anyway"
3. This is a false positive - the app is safe

**Problem**: Antivirus blocking the application

**Solution**: Add the .exe to your antivirus whitelist/exceptions

### Features not working

**Problem**: "File not found" errors

**Solution**: Make sure you've configured all file paths in the "Arquivos" tab

**Problem**: Can't connect to API

**Solution**: Check your internet connection

**Problem**: Browser automation fails

**Solution**:

1. Make sure Firefox is not running
2. The app will download the necessary browser driver automatically

### Update issues

**Problem**: Update check fails

**Solution**:

1. Check your internet connection
2. Try again later
3. You can always download updates manually from GitHub

## 📁 Data Storage

The application stores temporary data in:

- `%TEMP%\loglife-automations\`

Your configured file paths are saved here for convenience.

## 🔒 Privacy & Security

- The app only accesses files and folders you explicitly select
- Update checks connect to GitHub (secure HTTPS)
- No personal data is collected or transmitted
- All automation runs locally on your machine

## 📞 Support

### Common Questions

**Q: Do I need Python installed?**  
A: No! The executable includes everything.

**Q: Can I run this on Mac or Linux?**  
A: No, this is Windows-only.

**Q: How do I uninstall?**  
A: Just delete the .exe file. That's it!

**Q: Where can I find logs?**  
A: Check the `logs/` folder in the same directory as the executable.

**Q: Can I run multiple instances?**  
A: It's not recommended - may cause conflicts.

### Need Help?

1. Check this README
2. Review error messages in the app
3. Contact your IT department or system administrator

## 📋 Version History

Check the "Sobre" tab in the application to see your current version.

For detailed changelog, visit:
https://github.com/ArturSenna/loglife-automations/releases

## ⚖️ License

© 2025 LogLife Automations - All rights reserved

---

**Developed for LogLife Logistics**  
**Last Updated**: November 2025

## 🎉 Tips for Best Experience

1. **Keep it updated** - Check for updates regularly
2. **Configure once** - Set up file paths in the "Arquivos" tab first
3. **Close browsers** - Make sure Firefox is closed before automation
4. **Stable internet** - Ensure good connection for API operations
5. **Check logs** - If something fails, check the logs folder

---

**Enjoy using LogLife Operations!** 🚀
