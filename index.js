#!/usr/bin/env node

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

// Ensure Python and pip are installed
function checkPythonDependencies() {
  try {
    // Check if Python is available
    const pythonVersion = spawn('python3', ['--version']);
    pythonVersion.on('error', (err) => {
      console.error('Python3 is not installed or not in PATH. Please install Python 3.8+');
      process.exit(1);
    });

    // Check if pip is available
    const pipVersion = spawn('pip3', ['--version']);
    pipVersion.on('error', (err) => {
      console.error('pip3 is not installed or not in PATH. Please install pip');
      process.exit(1);
    });
  } catch (error) {
    console.error('Error checking Python dependencies:', error);
    process.exit(1);
  }
}

// Install required Python packages
function installPythonDependencies() {
  return new Promise((resolve, reject) => {
    console.log('Installing Python dependencies...');
    
    // Path to requirements.txt in the package
    const requirementsPath = path.join(__dirname, 'requirements.txt');
    
    // Install dependencies using pip
    const pip = spawn('pip3', ['install', '-r', requirementsPath]);
    
    pip.stdout.on('data', (data) => {
      console.log(`${data}`);
    });
    
    pip.stderr.on('data', (data) => {
      console.error(`${data}`);
    });
    
    pip.on('close', (code) => {
      if (code === 0) {
        console.log('Python dependencies installed successfully');
        resolve();
      } else {
        console.error(`pip exited with code ${code}`);
        reject(new Error(`Failed to install Python dependencies (exit code: ${code})`));
      }
    });
  });
}

// Run the MCP server
function runMCPServer() {
  console.log('Starting PowerPoint Translator MCP server...');
  
  // Path to the Python server script
  const serverPath = path.join(__dirname, 'server.py');
  
  // Make sure the script is executable
  fs.chmodSync(serverPath, '755');
  
  // Run the Python script in MCP mode
  const server = spawn('python3', [serverPath]);
  
  // Pipe stdin/stdout directly to the Python process for MCP communication
  process.stdin.pipe(server.stdin);
  server.stdout.pipe(process.stdout);
  server.stderr.on('data', (data) => {
    process.stderr.write(`${data}`);
  });
  
  server.on('close', (code) => {
    console.log(`MCP server exited with code ${code}`);
    process.exit(code);
  });
  
  // Handle process termination
  process.on('SIGINT', () => {
    console.log('Received SIGINT. Shutting down MCP server...');
    server.kill('SIGINT');
  });
  
  process.on('SIGTERM', () => {
    console.log('Received SIGTERM. Shutting down MCP server...');
    server.kill('SIGTERM');
  });
}

// Main function
async function main() {
  try {
    checkPythonDependencies();
    await installPythonDependencies();
    runMCPServer();
  } catch (error) {
    console.error('Error starting MCP server:', error);
    process.exit(1);
  }
}

// Run the main function
main();
