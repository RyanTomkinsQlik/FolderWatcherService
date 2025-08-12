const Service = require('node-windows').Service;
const path = require('path');

// Create a new service object (same config as installer)
const svc = new Service({
  name: 'Folder Watcher Service',
  script: path.join(__dirname, 'folder_watcher.js')
});

// Listen for the "uninstall" event
svc.on('uninstall', function(){
  console.log('✅ Folder Watcher Service has been uninstalled');
  console.log('🗑️  Service removed from Windows Services');
});

// Uninstall the service
console.log('🔧 Uninstalling Folder Watcher Service...');
svc.uninstall();