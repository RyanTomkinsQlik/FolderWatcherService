# FolderWatcherService
Folder watcher service prints new files


How to Install and Use:
Installation Steps:

Save all the files in your directory
Install node-windows:
powershellnpm install -g node-windows
cd your directory
npm link node-windows

Run as Administrator and install service:
powershell# Open PowerShell as Administrator
cd your directory
node install-service.js


Managing the Service:
Check service status:

Open services.msc
Look for "Folder Watcher Service"

Start/Stop service manually:
powershell# Start service
net start "Folder Watcher Service"

# Stop service  
net stop "Folder Watcher Service"
Uninstall service:
powershell# Run as Administrator
node uninstall-service.js
Service Features:
✅ Auto-starts when Windows boots
✅ Restarts automatically if it crashes
✅ Runs in background (no console window)
✅ Logs to Windows Event Viewer
✅ Can be managed through Windows Services
Configuration:
Edit service-config.js to change:

Watch folder path
Move destination
Printing settings
Auto-start behavior

The service will run continuously in the background, monitoring your folder and printing/moving files automatically!
