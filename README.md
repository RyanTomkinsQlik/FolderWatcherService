# Folder Watcher Service

A Windows service that monitors a specified folder for new files and automatically prints them to a configured printer. The service supports various file types including PDFs, text files, and Microsoft Office documents.

## Features

- **Real-time folder monitoring** - Detects new files immediately
- **Automatic printing** - Sends files to printer with queue management
- **File processing** - Moves processed files to a separate folder
- **Multiple file type support** - PDF, TXT, DOC, DOCX, XLS, XLSX, PPT, PPTX, and more
- **Print queue management** - Sequential processing to prevent printer conflicts
- **Error handling** - Robust error recovery and logging
- **Service integration** - Runs as Windows service for automatic startup

## System Requirements

- Windows 10/11 or Windows Server 2019/2022
- Node.js 14.x or later
- Brother MFC-J1205W Printer (or modify config for your printer)
- Administrator privileges for service installation

## Installation

### 1. Prerequisites

Install Node.js from [nodejs.org](https://nodejs.org) if not already installed.

### 2. Download Files

Place the following files in your service directory (e.g., `C:\FolderWatcher\`):
- `folder_watcher.js` - Main service code
- `service_config.js` - Configuration file

### 3. Install Dependencies

Open Command Prompt as Administrator and navigate to your service directory:

```bash
cd C:\FolderWatcher
npm init -y
npm install --save node-windows
```

### 4. Install as Windows Service

Create a service installer script (`install-service.js`):

```javascript
const Service = require('node-windows').Service;
const path = require('path');

// Create a new service object
const svc = new Service({
  name: 'FolderWatcherService',
  description: 'Monitors folder for new files, prints and moves them',
  script: path.join(__dirname, 'folder_watcher.js'),
  nodeOptions: [
    '--harmony',
    '--max_old_space_size=4096'
  ],
  workingDirectory: __dirname,
  allowServiceLogon: true
});

// Listen for the "install" event, which indicates the process is available as a service
svc.on('install', function() {
  console.log('Service installed successfully!');
  svc.start();
});

svc.on('alreadyinstalled', function() {
  console.log('Service is already installed.');
});

// Install the service
svc.install();
```

Run the installer:

```bash
node install-service.js
```

### 5. Verify Installation

Check that the service is running:
- Open Services.msc
- Look for "FolderWatcherService"
- Status should be "Running"

## Configuration

### Default Settings

Edit `service_config.js` to customize the service:

```javascript
module.exports = {
  serviceName: 'FolderWatcherService',
  serviceDescription: 'Monitors folder for new files, prints and moves them',
  
  // Folder paths
  watchPath: 'C:\\Users\\SDP\\Documents\\WatchedItems',
  moveToFolder: 'C:\\Users\\SDP\\Documents\\ProcessedFiles',
  
  // Printing settings
  enablePrinting: true,
  
  // Service behavior
  autoStart: true,
  restartOnFailure: true,
  maxRestarts: 3,
  
  // Logging
  logLevel: 'info',
  logToEventViewer: true
};
```

### Command Line Usage (Development)

For testing without installing as service:

```bash
# Basic usage
node folder_watcher.js

# Custom watch folder
node folder_watcher.js "C:\MyWatchFolder"

# Custom watch and destination folders
node folder_watcher.js "C:\MyWatchFolder" "C:\MyProcessedFolder"

# Disable printing
node folder_watcher.js "C:\MyWatchFolder" "C:\MyProcessedFolder" false

# Enable printing
node folder_watcher.js "C:\MyWatchFolder" "C:\MyProcessedFolder" true
```

## Supported File Types

### Text Files (Printed as formatted text)
- `.txt` - Plain text files
- `.log` - Log files
- `.json` - JSON files
- `.xml` - XML files
- `.csv` - Comma-separated values
- `.html` - HTML files
- `.css` - Stylesheet files
- `.js` - JavaScript files

### Document Files (Printed in original format)
- `.pdf` - PDF documents
- `.doc` / `.docx` - Microsoft Word documents
- `.xls` / `.xlsx` - Microsoft Excel spreadsheets
- `.ppt` / `.pptx` - Microsoft PowerPoint presentations

## Printer Configuration

### Default Printer Setup

The service is configured for "Brother MFC-J1205W Printer". To use a different printer:

1. Find your printer name:
   ```powershell
   Get-WmiObject -Class Win32_Printer | Select-Object Name
   ```

2. Edit the printer name in `folder_watcher.js` (search for "Brother MFC-J1205W Printer")

### PDF Printing Requirements

For optimal PDF printing, install one or more of these applications:

1. **SumatraPDF** (Recommended)
   - Download: [sumatrapdfreader.org](https://www.sumatrapdfreader.org)
   - Lightweight and excellent command-line support

2. **Adobe Acrobat Reader DC**
   - Download: [get.adobe.com/reader](https://get.adobe.com/reader)
   - Most compatible with complex PDFs

3. **PDFtoPrinter** (Command-line utility)
   - Download: [github.com/mhitza/PDFtoPrinter](https://github.com/mhitza/PDFtoPrinter)
   - Dedicated PDF printing utility

## Folder Structure

Create these folders or they will be created automatically:

```
C:\Users\SDP\Documents\
├── WatchedItems\          # Drop files here
└── ProcessedFiles\        # Files moved here after processing
```

## Service Management

### Start/Stop Service

```bash
# Start service
net start FolderWatcherService

# Stop service
net stop FolderWatcherService

# Restart service
net stop FolderWatcherService && net start FolderWatcherService
```

### Uninstall Service

Create an uninstaller script (`uninstall-service.js`):

```javascript
const Service = require('node-windows').Service;
const path = require('path');

const svc = new Service({
  name: 'FolderWatcherService',
  script: path.join(__dirname, 'folder_watcher.js')
});

svc.on('uninstall', function() {
  console.log('Service uninstalled successfully!');
});

svc.uninstall();
```

Run: `node uninstall-service.js`

## Monitoring and Logs

### Event Viewer Logs

1. Open Event Viewer (eventvwr.msc)
2. Navigate to: Applications and Services Logs > FolderWatcherService
3. View service logs and errors

### Service Status

Monitor service status:
```powershell
Get-Service FolderWatcherService
Get-WmiObject -Class Win32_PrintJob | Where-Object {$_.Name -like '*Brother*'}
```

## Troubleshooting

### Common Issues

#### Service Won't Start
- Check Event Viewer for error details
- Verify Node.js is installed correctly
- Ensure service has proper permissions
- Check that watch folders exist

#### Files Not Printing
1. Verify printer is online and has paper/toner
2. Check printer name matches configuration
3. Install PDF viewer applications (SumatraPDF/Adobe Reader)
4. Test printer with manual print job

#### Permission Errors
- Run Command Prompt as Administrator
- Ensure service account has folder access
- Check printer permissions

#### Files Not Moving
- Verify destination folder exists
- Check folder permissions
- Ensure files aren't locked by other applications

### Debug Mode

Run in development mode to see detailed logs:
```bash
node folder_watcher.js "C:\TestWatch" "C:\TestProcessed" true
```

### Print Queue Issues

Clear print queue if stuck:
```powershell
# Stop print spooler
Stop-Service Spooler

# Clear queue files
Remove-Item C:\Windows\System32\spool\PRINTERS\* -Force

# Start print spooler
Start-Service Spooler
```

## Performance Tuning

### Large File Handling
- Service includes delays for file write completion
- Print queue processes files sequentially
- Automatic retry logic for busy files

### Memory Usage
- Service includes garbage collection hints
- Temporary files are cleaned up automatically
- Service restarts automatically on memory issues

## Security Considerations

- Service runs with system privileges
- Monitor watch folder access carefully
- Consider virus scanning of incoming files
- Log all file processing activities

## Support

For issues or questions:
1. Check Event Viewer logs first
2. Verify printer connectivity
3. Test with simple text files
4. Check folder permissions
5. Review service configuration

## Version History

- **v1.0** - Initial release with basic file watching
- **v1.1** - Added print queue management
- **v1.2** - Enhanced PDF printing support
- **v1.3** - Added service recovery and error handling
- **v1.4** - Improved file type detection and processing

---

**Note**: This service is designed for trusted environments. Always verify file sources before processing.