const fs = require('fs');
const path = require('path');

class FolderWatcher {
  constructor(watchPath, moveToFolder = null, enablePrinting = false) {
    this.watchPath = watchPath;
    this.moveToFolder = moveToFolder;
    this.enablePrinting = enablePrinting;
    this.existingFiles = new Set();
    this.isInitialized = false;
    this.printQueue = [];
    this.isPrinting = false;
    this.watcher = null;
  }

  // Initialize by recording existing files
  async initialize() {
    try {
      // Ensure directory exists
      if (!fs.existsSync(this.watchPath)) {
        fs.mkdirSync(this.watchPath, { recursive: true });
        console.log(`Created directory: ${this.watchPath}`);
      }

      // Record existing files
      const files = fs.readdirSync(this.watchPath);
      files.forEach(file => {
        const filePath = path.join(this.watchPath, file);
        if (fs.statSync(filePath).isFile()) {
          this.existingFiles.add(file);
        }
      });

      console.log(`Watching folder: ${this.watchPath}`);
      console.log(`Initial files: ${Array.from(this.existingFiles).join(', ') || 'none'}`);
      console.log('Waiting for new files...\n');
      
      this.isInitialized = true;
      return true;
    } catch (error) {
      console.error('Error initializing watcher:', error.message);
      return false;
    }
  }

  // Read, display, and print file contents, then move file
  async printFileContents(filePath) {
    try {
      // Add a small delay to ensure file is fully written
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      const fileName = path.basename(filePath);
      const fileExtension = path.extname(fileName).toLowerCase();
      
      console.log('='.repeat(50));
      console.log(`📄 NEW FILE: ${fileName}`);
      console.log(`📍 Path: ${filePath}`);
      console.log('='.repeat(50));
      
      // Handle different file types
      let content = '';
      let fileSize = 0;
      let printType = 'text'; // 'text', 'original', or 'skip'
      
      // Determine print type based on file extension
      if (['.docx', '.doc', '.pdf', '.xls', '.xlsx', '.ppt', '.pptx'].includes(fileExtension)) {
        printType = 'original';
        const stats = fs.statSync(filePath);
        fileSize = stats.size;
        content = `[${fileExtension.toUpperCase().substring(1)} Document]\nFile: ${fileName}\nSize: ${fileSize} bytes\nThis document will be printed in its original format.`;
      } else if (['.txt', '.log', '.json', '.xml', '.csv', '.html', '.css', '.js'].includes(fileExtension)) {
        printType = 'text';
        try {
          content = fs.readFileSync(filePath, 'utf8');
          fileSize = content.length;
        } catch (error) {
          content = `[Error reading text file: ${error.message}]`;
          printType = 'skip';
        }
      } else {
        // Handle as potential text file first, fallback to binary
        try {
          content = fs.readFileSync(filePath, 'utf8');
          fileSize = content.length;
          // Check if it looks like binary data
          if (content.includes('\0') || /[\x00-\x08\x0E-\x1F\x7F-\xFF]{10,}/.test(content)) {
            throw new Error('Binary content detected');
          }
          printType = 'text';
        } catch (error) {
          const stats = fs.statSync(filePath);
          fileSize = stats.size;
          content = `[Binary/Unknown file - ${fileSize} bytes]\nFile type: ${fileExtension || 'unknown'}\nUse a specialized application to view this file.`;
          printType = 'skip';
        }
      }
      
      console.log(`📊 Size: ${fileSize} ${printType === 'text' ? 'characters' : 'bytes'}`);
      console.log(`🔖 Print Type: ${printType}`);
      console.log('='.repeat(50));
      console.log(content.length > 2000 ? content.substring(0, 2000) + '\n[... content truncated for display ...]' : content);
      console.log('='.repeat(50));
      console.log(''); // Empty line for spacing
      
      // Add to print queue if enabled (DO NOT move file yet)
      if (this.enablePrinting && printType !== 'skip') {
        console.log(`📋 Adding to print queue: ${fileName}`);
        await this.addToPrintQueueAndWait(filePath, fileName, content, printType);
      }
      
      // Move file AFTER printing is completely done
      await this.moveFileAfterProcessing(filePath, fileName);
      
    } catch (error) {
      console.error(`Error processing file ${filePath}:`, error.message);
      // Even if processing fails, try to move the file to prevent reprocessing
      try {
        await this.moveFileAfterProcessing(filePath, path.basename(filePath));
      } catch (moveError) {
        console.error(`Failed to move file after error:`, moveError.message);
      }
    }
  }

  // Add file to print queue and wait for completion
  async addToPrintQueueAndWait(filePath, fileName, content, printType) {
    return new Promise((resolve, reject) => {
      const printJob = { 
        filePath, 
        fileName, 
        content, 
        printType, 
        resolve, 
        reject,
        id: Date.now() + Math.random()
      };
      
      this.printQueue.push(printJob);
      console.log(`📋 Added to print queue: ${fileName} (Queue length: ${this.printQueue.length})`);
      
      // Process queue if not already processing
      if (!this.isPrinting) {
        this.processPrintQueue();
      }
    });
  }

  // Process print queue sequentially
  async processPrintQueue() {
    if (this.isPrinting || this.printQueue.length === 0) {
      return;
    }

    this.isPrinting = true;
    console.log(`🖨️  Starting print queue processing...`);

    while (this.printQueue.length > 0) {
      const printJob = this.printQueue.shift();
      console.log(`📄 Processing print job: ${printJob.fileName} (${this.printQueue.length} remaining)`);
      
      try {
        const success = await this.sendToPrinter(printJob.filePath, printJob.fileName, printJob.content, printJob.printType);
        
        if (printJob.resolve) {
          printJob.resolve(success);
        }
        
        // Add longer delay between print jobs to prevent conflicts
        console.log(`⏳ Waiting between print jobs...`);
        await new Promise(resolve => setTimeout(resolve, 8000)); // Increased delay
        
      } catch (error) {
        console.error(`❌ Print job failed for ${printJob.fileName}:`, error.message);
        
        if (printJob.reject) {
          printJob.reject(error);
        }
      }
    }

    this.isPrinting = false;
    console.log(`✅ Print queue processing completed`);
  }

  // Enhanced print function with improved PDF printing methods
  async sendToPrinter(filePath, fileName, content, printType) {
    try {
      console.log(`🖨️  Preparing to print: ${fileName} (${printType})`);
      console.log(`📍 Full path: ${filePath}`);
      
      // Verify file exists before attempting to print
      if (!fs.existsSync(filePath)) {
        throw new Error(`File not found: ${filePath}`);
      }
      
      const { exec, spawn } = require('child_process');
      const util = require('util');
      const execAsync = util.promisify(exec);
      
      if (printType === 'text') {
        // For text files, create a properly formatted document
        console.log(`🔄 Printing text content: ${fileName}`);
        
        const tempDir = require('os').tmpdir();
        const timestamp = Date.now();
        const tempFileName = `filewatch_${timestamp}.txt`;
        const tempFilePath = path.join(tempDir, tempFileName);
        
        // Create printable content with proper formatting
        const printableContent = [
          'FILE WATCHER PRINT JOB',
          '='.repeat(60),
          `File: ${fileName}`,
          `Original Path: ${filePath}`,
          `Processed: ${new Date().toLocaleString()}`,
          `Size: ${content.length} characters`,
          '='.repeat(60),
          '',
          content,
          '',
          '='.repeat(60),
          'End of Document',
          ''
        ].join('\n');
        
        // Write to temp file
        fs.writeFileSync(tempFilePath, printableContent, 'utf8');
        console.log(`📝 Created temp file: ${tempFilePath}`);
        
        try {
          // Use notepad to print (more reliable)
          await execAsync(`notepad.exe /p "${tempFilePath}"`, { timeout: 30000 });
          console.log(`✅ Text print job completed: ${fileName}`);
          
          // Clean up temp file after delay
          setTimeout(() => {
            try {
              if (fs.existsSync(tempFilePath)) {
                fs.unlinkSync(tempFilePath);
                console.log(`🗑️  Cleaned up temp file: ${tempFileName}`);
              }
            } catch (cleanupError) {
              console.log(`⚠️  Could not clean up temp file: ${cleanupError.message}`);
            }
          }, 10000);
          
          return true;
        } catch (error) {
          console.error(`❌ Notepad print failed: ${error.message}`);
          throw error;
        }
        
      } else if (printType === 'original') {
        const fileExtension = path.extname(fileName).toLowerCase();
        
        if (fileExtension === '.pdf') {
          console.log(`🔄 Printing PDF document: ${fileName}`);
          
          // First, check if the specific Brother printer is available
          try {
            const printerCheck = await execAsync('powershell -Command "Get-WmiObject -Class Win32_Printer | Where-Object {$_.Name -like \'*Brother MFC-J1205W*\'} | Select-Object Name, PrinterStatus, WorkOffline, Default"');
            console.log(`🖨️  Brother printer info:\n${printerCheck.stdout}`);
            
            if (!printerCheck.stdout.includes('Brother MFC-J1205W')) {
              console.log(`⚠️  Brother MFC-J1205W Printer not found! Available printers:`);
              const allPrinters = await execAsync('powershell -Command "Get-WmiObject -Class Win32_Printer | Select-Object Name | Format-Table -HideTableHeaders"');
              console.log(allPrinters.stdout);
            }
          } catch (e) {
            console.log(`⚠️  Could not check Brother printer: ${e.message}`);
          }
          
          // Try multiple PDF printing methods with improved error handling
          const pdfMethods = [
            {
              name: 'SumatraPDF with Brother Printer',
              command: async () => {
                const sumatraPaths = [
                  'C:\\Users\\SDP\\AppData\\Local\\SumatraPDF\\SumatraPDF.exe',
                  'C:\\Program Files\\SumatraPDF\\SumatraPDF.exe',
                  'C:\\Program Files (x86)\\SumatraPDF\\SumatraPDF.exe'
                ];
                
                for (const sumatraPath of sumatraPaths) {
                  if (fs.existsSync(sumatraPath)) {
                    console.log(`📄 Found SumatraPDF at: ${sumatraPath}`);
                    const cmd = `"${sumatraPath}" -print-to "Brother MFC-J1205W Printer" -silent "${filePath}"`;
                    console.log(`🔄 Running: ${cmd}`);
                    await execAsync(cmd, { timeout: 25000 });
                    console.log(`✅ SumatraPDF command executed successfully`);
                    return true;
                  }
                }
                throw new Error('SumatraPDF not found');
              }
            },
            {
              name: 'PowerShell Direct Printer Method',
              command: async () => {
                console.log(`🔄 Using PowerShell direct printing method`);
                const cmd = `powershell -ExecutionPolicy Bypass -Command "
                  try {
                    Write-Host 'Loading printing assemblies...';
                    Add-Type -AssemblyName System.Drawing;
                    Add-Type -AssemblyName System.Windows.Forms;
                    
                    Write-Host 'Finding Brother printer...';
                    $printerName = 'Brother MFC-J1205W Printer';
                    $printers = Get-WmiObject -Class Win32_Printer | Where-Object { $_.Name -eq $printerName };
                    
                    if (-not $printers) {
                      throw 'Brother MFC-J1205W Printer not found';
                    }
                    
                    Write-Host 'Setting up print job...';
                    $psi = New-Object System.Diagnostics.ProcessStartInfo;
                    $psi.FileName = 'cmd.exe';
                    $psi.Arguments = '/c print /d:\\\"$printerName\\\" \\\"${filePath}\\\"';
                    $psi.UseShellExecute = $false;
                    $psi.CreateNoWindow = $true;
                    $psi.RedirectStandardOutput = $true;
                    $psi.RedirectStandardError = $true;
                    
                    Write-Host 'Starting print process...';
                    $process = [System.Diagnostics.Process]::Start($psi);
                    $output = $process.StandardOutput.ReadToEnd();
                    $error = $process.StandardError.ReadToEnd();
                    $process.WaitForExit(30000);
                    
                    Write-Host \\\"Print command output: $output\\\";
                    if ($error) { Write-Host \\\"Print command error: $error\\\"; }
                    
                    if ($process.ExitCode -eq 0) {
                      Write-Host 'Print job completed successfully';
                    } else {
                      throw \\\"Print command failed with exit code: $($process.ExitCode)\\\";
                    }
                  } catch {
                    Write-Host \\\"Error: $($_.Exception.Message)\\\";
                    throw $_.Exception;
                  }"`;
                
                const result = await execAsync(cmd, { timeout: 45000 });
                console.log(`📄 PowerShell result: ${result.stdout}`);
                if (result.stderr) {
                  console.log(`⚠️  PowerShell stderr: ${result.stderr}`);
                }
                return result.stdout.includes('Print job completed successfully');
              }
            },
            {
              name: 'Adobe Reader Silent Print',
              command: async () => {
                const adobePaths = [
                  'C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe',
                  'C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe',
                  'C:\\Program Files\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe'
                ];
                
                for (const adobePath of adobePaths) {
                  if (fs.existsSync(adobePath)) {
                    console.log(`📄 Found Adobe Reader at: ${adobePath}`);
                    
                    // Try different Adobe command line options
                    const adobeCommands = [
                      `"${adobePath}" /s /o /h /t "${filePath}" "Brother MFC-J1205W Printer"`,
                      `"${adobePath}" /N /T "${filePath}" "Brother MFC-J1205W Printer"`,
                      `"${adobePath}" /p /h "${filePath}"`
                    ];
                    
                    for (let i = 0; i < adobeCommands.length; i++) {
                      try {
                        console.log(`🔄 Adobe attempt ${i + 1}: ${adobeCommands[i]}`);
                        await execAsync(adobeCommands[i], { timeout: 30000 });
                        console.log(`✅ Adobe command ${i + 1} executed successfully`);
                        
                        // Wait for Adobe to process
                        await new Promise(resolve => setTimeout(resolve, 5000));
                        
                        return true;
                      } catch (adobeError) {
                        console.log(`❌ Adobe command ${i + 1} failed: ${adobeError.message}`);
                        if (i === adobeCommands.length - 1) {
                          throw adobeError;
                        }
                      }
                    }
                  }
                }
                throw new Error('Adobe Reader not found');
              }
            },
            {
              name: 'Windows Shell Print Command',
              command: async () => {
                console.log(`🔄 Using Windows shell print command`);
                const cmd = `print /d:"Brother MFC-J1205W Printer" "${filePath}"`;
                console.log(`🔄 Running: ${cmd}`);
                const result = await execAsync(cmd, { timeout: 30000 });
                console.log(`📄 Shell print result: ${result.stdout}`);
                if (result.stderr) {
                  console.log(`⚠️  Shell print stderr: ${result.stderr}`);
                }
                return !result.stderr.includes('Error') && !result.stderr.includes('failed');
              }
            },
            {
              name: 'PDFtoPrinter Utility',
              command: async () => {
                // Check if PDFtoPrinter is available (free utility)
                const pdfToPrinterPaths = [
                  'C:\\Program Files\\PDFtoPrinter\\PDFtoPrinter.exe',
                  'C:\\Program Files (x86)\\PDFtoPrinter\\PDFtoPrinter.exe',
                  path.join(__dirname, 'PDFtoPrinter.exe'),
                  path.join(process.cwd(), 'PDFtoPrinter.exe')
                ];
                
                for (const pdfPath of pdfToPrinterPaths) {
                  if (fs.existsSync(pdfPath)) {
                    console.log(`📄 Found PDFtoPrinter at: ${pdfPath}`);
                    const cmd = `"${pdfPath}" "${filePath}" "Brother MFC-J1205W Printer"`;
                    console.log(`🔄 Running: ${cmd}`);
                    await execAsync(cmd, { timeout: 20000 });
                    return true;
                  }
                }
                throw new Error('PDFtoPrinter not found - download from https://github.com/mhitza/PDFtoPrinter');
              }
            },
            {
              name: 'PowerShell .NET PDF Printing',
              command: async () => {
                console.log(`🔄 Using .NET PDF printing method`);
                const cmd = `powershell -ExecutionPolicy Bypass -Command "
                  try {
                    Write-Host 'Loading .NET printing components...';
                    Add-Type -AssemblyName System.Drawing;
                    Add-Type -AssemblyName System.Drawing.Printing;
                    
                    Write-Host 'Creating print document...';
                    $printDoc = New-Object System.Drawing.Printing.PrintDocument;
                    $printDoc.PrinterSettings.PrinterName = 'Brother MFC-J1205W Printer';
                    
                    if (-not $printDoc.PrinterSettings.IsValid) {
                      throw 'Brother MFC-J1205W Printer is not valid or available';
                    }
                    
                    Write-Host 'Printer is valid, attempting to print...';
                    
                    # Use Start-Process with Shell Execute to open with default PDF viewer and print
                    $psi = New-Object System.Diagnostics.ProcessStartInfo;
                    $psi.FileName = '${filePath}';
                    $psi.Verb = 'print';
                    $psi.UseShellExecute = $true;
                    $psi.CreateNoWindow = $true;
                    $psi.WindowStyle = 'Hidden';
                    
                    $process = [System.Diagnostics.Process]::Start($psi);
                    if ($process) {
                      Write-Host 'Print process started successfully';
                      Start-Sleep 5;
                      Write-Host 'Print job queued successfully';
                    } else {
                      throw 'Failed to start print process';
                    }
                  } catch {
                    Write-Host \\\"Error in .NET printing: $($_.Exception.Message)\\\";
                    throw $_.Exception;
                  }"`;
                
                const result = await execAsync(cmd, { timeout: 45000 });
                console.log(`📄 .NET print result: ${result.stdout}`);
                if (result.stderr) {
                  console.log(`⚠️  .NET print stderr: ${result.stderr}`);
                }
                return result.stdout.includes('Print job queued successfully');
              }
            },
            {
              name: 'GSPrint with Brother Printer',
              command: async () => {
                // Try GSPrint if available (excellent for service contexts)
                const gsPrintPaths = [
                  'C:\\Program Files\\Ghostgum\\gsview\\gsprint.exe',
                  'C:\\Program Files (x86)\\Ghostgum\\gsview\\gsprint.exe',
                  'C:\\GSPrint\\gsprint.exe'
                ];
                
                for (const gsPath of gsPrintPaths) {
                  if (fs.existsSync(gsPath)) {
                    console.log(`📄 Found GSPrint at: ${gsPath}`);
                    const cmd = `"${gsPath}" -printer "Brother MFC-J1205W Printer" "${filePath}"`;
                    console.log(`🔄 Running: ${cmd}`);
                    await execAsync(cmd, { timeout: 20000 });
                    return true;
                  }
                }
                throw new Error('GSPrint not found - download from http://www.ghostgum.com.au/software/gsview.htm');
              }
            }
          ];
          
          let printSuccess = false;
          let lastError = null;
          
          for (let i = 0; i < pdfMethods.length && !printSuccess; i++) {
            try {
              console.log(`🔄 Trying PDF print method ${i + 1}: ${pdfMethods[i].name}`);
              await pdfMethods[i].command();
              printSuccess = true;
              console.log(`✅ PDF print method ${i + 1} (${pdfMethods[i].name}) succeeded!`);
              
              // Wait for print spooler
              console.log(`⏳ Waiting for print job to spool...`);
              await new Promise(resolve => setTimeout(resolve, 8000));
              
              // Check if print job was queued for Brother printer
              try {
                const queueCheck = await execAsync('powershell -Command "Get-WmiObject -Class Win32_PrintJob | Where-Object {$_.Name -like \'*Brother MFC-J1205W*\'} | Select-Object Name, Document, JobStatus | Format-List"', { timeout: 5000 });
                if (queueCheck.stdout.trim()) {
                  console.log(`📋 Brother printer queue contains:\n${queueCheck.stdout}`);
                } else {
                  console.log(`📋 Brother printer queue is empty (job may have completed quickly)`);
                  // Also check all print jobs
                  const allJobs = await execAsync('powershell -Command "Get-WmiObject -Class Win32_PrintJob | Select-Object Name, Document | Format-List"', { timeout: 5000 });
                  if (allJobs.stdout.trim()) {
                    console.log(`📋 All print jobs:\n${allJobs.stdout}`);
                  }
                }
              } catch (queueError) {
                console.log(`⚠️  Could not check Brother printer queue: ${queueError.message}`);
              }
              
              break;
              
            } catch (error) {
              lastError = error;
              console.log(`❌ PDF print method ${i + 1} (${pdfMethods[i].name}) failed: ${error.message}`);
              
              // Wait a bit before trying next method
              await new Promise(resolve => setTimeout(resolve, 2000));
            }
          }
          
          if (!printSuccess) {
            // Final fallback: Try to register PDF file association and retry
            console.log(`🔄 Attempting to fix PDF file associations and retry...`);
            try {
              await execAsync('powershell -Command "Start-Process -FilePath \'sfc\' -ArgumentList \'/scannow\' -Verb RunAs -WindowStyle Hidden"', { timeout: 5000 });
            } catch (e) {
              // Ignore SFC errors
            }
            
            throw new Error(`All PDF print methods failed. Last error: ${lastError?.message || 'Unknown error'}`);
          }
          
          console.log(`✅ PDF print job completed: ${fileName}`);
          return true;
          
        } else {
          // Handle other document types (Word, Excel, PowerPoint)
          console.log(`🔄 Printing ${fileExtension.toUpperCase()} document: ${fileName}`);
          
          try {
            const cmd = `powershell -Command "Start-Process -FilePath '${filePath}' -Verb Print -WindowStyle Hidden"`;
            console.log(`🔄 Running: ${cmd}`);
            await execAsync(cmd, { timeout: 30000 });
            console.log(`✅ Document print job completed: ${fileName}`);
            
            // Wait for print job to process
            await new Promise(resolve => setTimeout(resolve, 5000));
            return true;
          } catch (error) {
            console.error(`❌ Document print failed: ${error.message}`);
            throw error;
          }
        }
      }
      
      return false;
      
    } catch (error) {
      console.error(`❌ Print operation failed for ${fileName}:`, error.message);
      
      // Provide troubleshooting suggestions
      if (error.message.includes('Adobe Reader not found') || error.message.includes('SumatraPDF not found')) {
        console.log(`💡 Install Adobe Reader or SumatraPDF for better PDF printing support`);
        console.log(`💡 Download Adobe Reader: https://get.adobe.com/reader/`);
        console.log(`💡 Download SumatraPDF: https://www.sumatrapdfreader.org/download-free-pdf-viewer.html`);
        console.log(`💡 Download PDFtoPrinter: https://github.com/mhitza/PDFtoPrinter`);
      } else if (error.message.includes('timeout')) {
        console.log(`💡 Print operation timed out - printer might be slow or offline`);
      } else if (error.message.includes('Access is denied')) {
        console.log(`💡 Permission denied - try running service as administrator`);
      } else if (error.message.includes('No application is associated')) {
        console.log(`💡 PDF file association issue - try reinstalling Adobe Reader or setting default PDF viewer`);
        console.log(`💡 Run: assoc .pdf=AcroExch.Document.DC (in admin command prompt)`);
      }
      
      throw error;
    }
  }

  // Move file after processing with better error handling
  async moveFileAfterProcessing(filePath, fileName) {
    if (!this.moveToFolder) {
      console.log(`📍 No move folder specified, keeping file in place`);
      return;
    }
    
    try {
      // Check if source file still exists
      if (!fs.existsSync(filePath)) {
        console.log(`⚠️  Source file no longer exists: ${fileName}`);
        return;
      }
      
      // Ensure destination folder exists
      if (!fs.existsSync(this.moveToFolder)) {
        fs.mkdirSync(this.moveToFolder, { recursive: true });
        console.log(`📁 Created destination folder: ${this.moveToFolder}`);
      }
      
      const destinationPath = path.join(this.moveToFolder, fileName);
      
      // Handle duplicate filenames by adding timestamp
      let finalDestination = destinationPath;
      let counter = 1;
      
      while (fs.existsSync(finalDestination)) {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
        const ext = path.extname(fileName);
        const nameWithoutExt = path.basename(fileName, ext);
        finalDestination = path.join(this.moveToFolder, `${nameWithoutExt}_${timestamp}_${counter}${ext}`);
        counter++;
      }
      
      // Move the file with retry logic
      let retries = 3;
      while (retries > 0) {
        try {
          fs.renameSync(filePath, finalDestination);
          console.log(`📦 Moved file to: ${finalDestination}`);
          return;
        } catch (moveError) {
          retries--;
          if (moveError.code === 'EBUSY' || moveError.code === 'ENOENT') {
            console.log(`⏳ File busy, retrying move... (${3-retries}/3)`);
            await new Promise(resolve => setTimeout(resolve, 2000));
          } else {
            throw moveError;
          }
        }
      }
      
      throw new Error(`Failed to move file after 3 attempts`);
      
    } catch (error) {
      console.error(`❌ Error moving file ${fileName}:`, error.message);
      console.log(`💡 File remains in watch folder: ${filePath}`);
    }
  }

  // Handle file system events
  async handleFileEvent(eventType, filename) {
    if (!filename || !this.isInitialized) return;

    const filePath = path.join(this.watchPath, filename);

    try {
      // Check if file exists and is actually a file
      if (!fs.existsSync(filePath)) return;
      
      const stats = fs.statSync(filePath);
      if (!stats.isFile()) return;

      // Prevent duplicate processing - check if we're already processing this file
      const fileKey = `${filename}_${stats.mtime.getTime()}_${stats.size}`;
      if (this.existingFiles.has(fileKey)) {
        return; // Already processing or processed this exact file
      }

      console.log(`🔍 Detected new file: ${filename}`);
      this.existingFiles.add(fileKey);
      
      // Add extra delay for file write completion
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      // Double-check file still exists before processing
      if (fs.existsSync(filePath)) {
        await this.printFileContents(filePath);
      } else {
        console.log(`⚠️  File disappeared before processing: ${filename}`);
      }
    } catch (error) {
      // File might have been deleted quickly, ignore ENOENT
      if (error.code !== 'ENOENT') {
        console.error(`Error handling file event for ${filename}:`, error.message);
      }
    }
  }

  // Start watching the folder
  startWatching() {
    try {
      this.watcher = fs.watch(this.watchPath, { persistent: true }, (eventType, filename) => {
        // Wrap in try-catch to prevent uncaught exceptions
        try {
          this.handleFileEvent(eventType, filename);
        } catch (error) {
          console.error('Error in file event handler:', error.message);
        }
      });

      this.watcher.on('error', (error) => {
        console.error('Watcher error:', error.message);
        // Don't exit on watcher errors, try to restart
        this.restartWatcher();
      });

      console.log('✅ File system watcher started successfully');
      return this.watcher;
    } catch (error) {
      console.error('Failed to start file system watcher:', error.message);
      throw error;
    }
  }

  // Restart watcher on error
  async restartWatcher() {
    console.log('🔄 Attempting to restart file watcher...');
    try {
      if (this.watcher) {
        this.watcher.close();
      }
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait 5 seconds
      this.startWatching();
      console.log('✅ File watcher restarted successfully');
    } catch (error) {
      console.error('Failed to restart watcher:', error.message);
    }
  }

  // Clean shutdown
  shutdown() {
    console.log('🛑 Shutting down file watcher...');
    if (this.watcher) {
      this.watcher.close();
    }
  }
}

// Main execution with better error handling
async function main() {
  try {
    // Try to load service configuration
    let config = {};
    const configPath = path.join(__dirname, 'service-config.js');
    
    try {
      if (fs.existsSync(configPath)) {
        config = require(configPath);
        console.log('✅ Loaded configuration from service-config.js');
      }
    } catch (error) {
      console.log('⚠️  Could not load service-config.js, using defaults');
    }
    
    // Get watch path from config, command line, or use default
    const watchPath = process.argv[2] || config.watchPath || 'C:\\Users\\SDP\\Documents\\WatchedItems';
    
    // Get move destination folder
    const moveToFolder = process.argv[3] || config.moveToFolder || 'C:\\Users\\SDP\\Documents\\ProcessedFiles';
    
    // Check if printing is enabled
    let enablePrinting = false;
    if (process.argv[4] === 'print' || process.argv[4] === 'true') {
      enablePrinting = true;
    } else if (process.argv[4] === 'false' || process.argv[4] === 'disable') {
      enablePrinting = false;
    } else if (config.enablePrinting !== undefined) {
      enablePrinting = config.enablePrinting === true;
    } else {
      enablePrinting = true; // Default to enabled
    }
    
    console.log('🔍 File Watcher Service Starting...');
    console.log(`📅 Started at: ${new Date().toLocaleString()}`);
    console.log(`📂 Watching: ${watchPath}`);
    
    if (enablePrinting) {
      console.log('🖨️  Printer mode: ENABLED');
    } else {
      console.log('🖨️  Printer mode: DISABLED');
    }
    
    if (moveToFolder) {
      console.log(`📦 Move to: ${moveToFolder}`);
    }
    
    console.log('Press Ctrl+C to stop\n');

    const watcher = new FolderWatcher(watchPath, moveToFolder, enablePrinting);
    
    const initSuccess = await watcher.initialize();
    if (!initSuccess) {
      throw new Error('Failed to initialize folder watcher');
    }
    
    if (moveToFolder) {
      console.log(`📦 Files will be moved to: ${moveToFolder} after processing`);
    }
    
    if (enablePrinting) {
      console.log(`🖨️  Files will be automatically printed using queue system`);
      console.log(`💡 Make sure your default printer is available and has paper/toner`);
    }
    
    console.log(''); // Empty line
    
    // Start watching with error handling
    try {
      watcher.startWatching();
    } catch (error) {
      throw new Error(`Failed to start file watching: ${error.message}`);
    }
    
    // Keep the service running
    console.log('✅ Folder Watcher Service is running...');
    
    // Handle graceful shutdown
    const handleShutdown = (signal) => {
      console.log(`\n🛑 Received ${signal}, stopping service...`);
      watcher.shutdown();
      setTimeout(() => {
        console.log('✅ Service stopped');
        process.exit(0);
      }, 2000);
    };

    process.on('SIGINT', () => handleShutdown('SIGINT'));
    process.on('SIGTERM', () => handleShutdown('SIGTERM'));
    
    // Handle uncaught exceptions to prevent service crash
    process.on('uncaughtException', (error) => {
      console.error('❌ Uncaught Exception:', error.message);
      console.error('Stack:', error.stack);
      // Don't exit, try to continue
    });

    process.on('unhandledRejection', (reason, promise) => {
      console.error('❌ Unhandled Rejection at:', promise, 'reason:', reason);
      // Don't exit, try to continue
    });
    
    // Keep alive heartbeat
    const keepAliveInterval = setInterval(() => {
      console.log(`💓 Service heartbeat: ${new Date().toLocaleTimeString()}`);
    }, 300000); // Every 5 minutes
    
    // Clean up interval on shutdown
    process.on('exit', () => {
      clearInterval(keepAliveInterval);
    });
    
  } catch (error) {
    console.error('❌ Fatal Error starting Folder Watcher Service:', error.message);
    console.error('Stack:', error.stack);
    process.exit(1);
  }
}

// Export for use as module
module.exports = FolderWatcher;

// Run if this file is executed directly
if (require.main === module) {
  main().catch((error) => {
    console.error('❌ Fatal Error in main:', error.message);
    process.exit(1);
  });
}