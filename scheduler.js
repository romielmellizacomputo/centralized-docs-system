const fs = require('fs');
const { exec } = require('child_process');

const logFile = 'C:/Users/Admin/OneDrive/Documents/Web Project/Centralized Docs System/log.txt';

function log(message) {
  const timestamp = new Date().toISOString();
  fs.appendFileSync(logFile, `[${timestamp}] ${message}\n`);
}

log('Scheduler script started');

try {
  if (!fs.existsSync(logFile)) {
    fs.writeFileSync(logFile, 'Log file created\n');
    log('Log file did not exist. Created new log file.');
  }
} catch (error) {
  console.error("Error writing to log file:", error);
  log(`Error initializing log file: ${error.stack}`);
}

exec('node "C:/Users/Admin/OneDrive/Documents/Web Project/Centralized Docs System/fetchallmr.js"', (err, stdout, stderr) => {
  if (err) {
    log(`Error in fetchallmr.js: ${err.stack}`);
  } else {
    log(`fetchallmr.js stdout: ${stdout.trim()}`);
  }

  if (stderr) {
    log(`fetchallmr.js stderr: ${stderr.trim()}`);
  }
});

exec('node "C:/Users/Admin/OneDrive/Documents/Web Project/Centralized Docs System/fetchallissues.js"', (err, stdout, stderr) => {
  if (err) {
    log(`Error in fetchallissues.js: ${err.stack}`);
  } else {
    log(`fetchallissues.js stdout: ${stdout.trim()}`);
  }

  if (stderr) {
    log(`fetchallissues.js stderr: ${stderr.trim()}`);
  }
});
