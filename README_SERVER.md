# How to Fix the CORS Error

## Understanding the Error

When you open `index.html` directly in your browser (using the `file://` protocol), the browser blocks JavaScript `fetch()` requests to local JSON files. This is a security feature called **CORS (Cross-Origin Resource Sharing)**.

**Error Message:**
```
Access to fetch at 'file:///C:/Users/DELL/Desktop/project%20mvp/data/clients.json' 
from origin 'null' has been blocked by CORS policy
```

## Solutions

### Solution 1: Use Python HTTP Server (Recommended - Easiest)

1. **Make sure Python is installed** (Python 3.x)
   - Check by running: `python --version` in Command Prompt

2. **Start the server:**
   - **Option A:** Double-click `server.bat`
   - **Option B:** Run in terminal:
     ```bash
     python server.py
     ```

3. **Open your browser** and go to:
   ```
   http://localhost:8000/index.html
   ```

4. **The CORS error will be gone!** ✅

### Solution 2: Use Node.js HTTP Server

If you have Node.js installed:

```bash
npx http-server -p 8000 -c-1
```

Then open: `http://localhost:8000/index.html`

### Solution 3: Use VS Code Live Server Extension

1. Install the "Live Server" extension in VS Code
2. Right-click on `index.html`
3. Select "Open with Live Server"

### Solution 4: Use Electron (If this is an Electron app)

If this project is meant to run as an Electron app, you should:
1. Set up Electron properly
2. The `loadJSONFile()` function already has Electron support built-in

## Why This Happens

- **File Protocol (`file://`)**: Browsers restrict access to local files for security
- **HTTP Protocol (`http://`)**: Local servers allow file access through HTTP
- **Solution**: Run a local web server to serve files over HTTP instead of `file://`

## Quick Start

**Easiest way:**
1. Double-click `server.bat`
2. Open `http://localhost:8000/index.html` in your browser
3. Done! No more CORS errors.

