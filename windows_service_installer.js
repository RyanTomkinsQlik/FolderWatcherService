const Service = require('node-windows').Service;
const path = require('path');

// Create a new service object
const svc = new Service({
  name: 'Folder Watcher Service',
  description: 'Monitors folder for new files, reads contents, prints and moves them',
  script: path.join(__dirname, 'folder_watcher.js'),
  nodeOptions: [
    '--harmony',
    '--max_old_space_size=4096'
  ],
  //, workingDirectory: path.join(__dirname)
  //, allowServiceLogon: true
});

// Listen for the "install" event, which indicates the
// process is available as a service.
svc.on('install', function(){
  console.log('✅ Service installed successfully!');
  console.log('🚀 Starting service...');
  svc.start();
});

svc.on('start', function(){
  console.log('✅ Folder Watcher Service is now running!');
  console.log('📝 Check Windows Services (services.msc) to manage it');
  console.log('📋 Logs will be in the Windows Event Viewer');
});

svc.on('alreadyinstalled', function(){
  console.log('⚠️  Service is already installed');
  console.log('💡 Run uninstall-service.js first if you want to reinstall');
});

// Install the service
console.log('🔧 Installing Folder Watcher as Windows Service...');
svc.install();