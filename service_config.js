// service-config.js
// Configuration file for the Windows service

module.exports = {
  // Service settings
  serviceName: 'FolderWatcherService',
  serviceDescription: 'Monitors folder for new files, reads contents, prints and moves them',
  
  // Folder watcher settings
  watchPath: 'C:\\Users\\SDP\\Documents\\WatchedItems',
  moveToFolder: 'C:\\Users\\SDP\\Documents\\ProcessedFiles',
  enablePrinting: true,
  
  // Service behavior
  autoStart: true,
  restartOnFailure: true,
  maxRestarts: 3,
  
  // Logging
  logLevel: 'info',
  logToEventViewer: true
};