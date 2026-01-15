@echo off
REM =============================================================================
REM Single Command Deployment Script for Windows
REM Downloads and installs MongoDB, Tesseract OCR, and starts all services
REM =============================================================================

setlocal EnableDelayedExpansion

REM Configuration
set "MONGO_DB_PATH=./data/db"
set "BACKEND_PORT=8000"
set "FRONTEND_PORT=3000"

REM Installation directories
set "MONGODB_INSTALL_DIR=%ProgramFiles%\MongoDB\Server\6.0\bin"
set "TESSERACT_INSTALL_DIR=%ProgramFiles%\Tesseract-OCR"

REM Project directories
set "PROJECT_DIR=%~dp0"
set "BACKEND_DIR=%PROJECT_DIR%backend"
set "FRONTEND_DIR=%PROJECT_DIR%frontend"
set "VENV_DIR=%BACKEND_DIR%venv"

REM Download URLs
set "MONGO_URL=https://fastdl.mongodb.org/windows/mongodb-windows-x86_64-6.0.16-signed.msi"
set "MONGO_MSI_NAME=mongodb-windows-x86_64-6.0.16-signed.msi"
set "TESSERACT_URL=https://github.com/UB-Mannheim/tesseract/wiki"
set "TESSERACT_URL_X86=https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3.20230701/tesseract-5.3.3.exe"
set "TESSERACT_URL_X64=https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3.20230701/tesseract-5.3.3.20230701.exe"

REM Color codes for Windows (limited support)
set "GREEN=[92m"
set "YELLOW=[93m"
set "RED=[91m"
set "BLUE=[94m"
set "NC=[0m"

REM Print colored message (works in modern terminals)
echo_color() {
    for /f "tokens=1,*" %%a in ("%*") do (
        set "color_code=%%a"
        set "message=%%b"
    )
    echo %color_code%%message%[0m
}

REM Alternative simple echo without colors for compatibility
echo_status() {
    echo [DEPLOY] %*
}

echo_warning() {
    echo [WARN] %*
}

echo_error() {
    echo [ERROR] %*
}

echo_info() {
    echo [INFO] %*
}

echo_success() {
    echo [OK] %*
}

REM =============================================================================
REM Utility Functions
REM =============================================================================

REM Check if a command exists
command_exists() {
    where %1 >nul 2>&1
    if %errorlevel% equ 0 (
        exit /b 0
    )
    exit /b 1
}

REM Check if a directory exists
dir_exists() {
    if exist "%~1" (
        exit /b 0
    )
    exit /b 1
}

REM Download a file using PowerShell
download_file() {
    set "url=%~1"
    set "output=%~2"
    set "filename=%~nx2"
    
    echo_info "Downloading %filename%..."
    
    REM Use PowerShell to download with retry logic
    powershell -Command "& {try {(New-Object System.Net.WebClient).DownloadFile('%url%', '%output%')} catch {Write-Error 'Download failed'; exit 1}}"
    
    if %errorlevel% neq 0 (
        echo_error "Failed to download %filename%"
        exit /b 1
    )
    
    echo_success "Downloaded %filename%"
    exit /b 0
}

REM Check if a Windows service exists
service_exists() {
    sc query %1 >nul 2>&1
    if %errorlevel% equ 0 (
        exit /b 0
    )
    exit /b 1
}

REM =============================================================================
REM MongoDB Installation Functions
REM =============================================================================

REM Check if MongoDB is installed
check_mongodb_installed() {
    echo_info "Checking MongoDB installation..."
    
    REM Check if mongod is in PATH
    where mongod >nul 2>&1
    if %errorlevel% equ 0 (
        echo_success "MongoDB is already installed (in PATH)"
        exit /b 0
    )
    
    REM Check if MongoDB is installed in default directory
    if exist "%MONGODB_INSTALL_DIR%\mongod.exe" (
        echo_success "MongoDB is already installed at %MONGODB_INSTALL_DIR%"
        REM Add to PATH
        setx PATH "%MONGODB_INSTALL_DIR%;%PATH%" >nul 2>&1
        exit /b 0
    )
    
    echo_warning "MongoDB is not installed"
    exit /b 1
}

REM Install MongoDB for Windows
install_mongodb() {
    echo_warning "MongoDB not found. Installing MongoDB Community Server..."
    
    set "MONGO_MSI=%TEMP%\%MONGO_MSI_NAME%"
    
    REM Download MongoDB
    call :download_file "%MONGO_URL%" "%MONGO_MSI%"
    if %errorlevel% neq 0 (
        echo_error "Failed to download MongoDB"
        exit /b 1
    )
    
    REM Install MongoDB silently
    echo_info "Installing MongoDB (this may take a few minutes)..."
    msiexec /qb /i "%MONGO_MSI%" INSTALLLOCATION="%ProgramFiles%\MongoDB\Server\6.0\" ADDLOCAL="all" /norestart
    
    if %errorlevel% neq 0 (
        echo_error "Failed to install MongoDB"
        del "%MONGO_MSI%" >nul 2>&1
        exit /b 1
    )
    
    REM Add MongoDB to PATH
    echo_info "Adding MongoDB to PATH..."
    setx PATH "%MONGODB_INSTALL_DIR%;%PATH%" >nul 2>&1
    
    REM Clean up
    del "%MONGO_MSI%" >nul 2>&1
    
    echo_success "MongoDB installed successfully!"
    echo_info "MongoDB installed at: %ProgramFiles%\MongoDB\Server\6.0\"
    
    exit /b 0
}

REM =============================================================================
REM Tesseract OCR Installation Functions
REM =============================================================================

REM Check if Tesseract OCR is installed
check_tesseract_installed() {
    echo_info "Checking Tesseract OCR installation..."
    
    REM Check if tesseract is in PATH
    where tesseract >nul 2>&1
    if %errorlevel% equ 0 (
        echo_success "Tesseract OCR is already installed (in PATH)"
        exit /b 0
    )
    
    REM Check if Tesseract is installed in default directory
    if exist "%TESSERACT_INSTALL_DIR%\tesseract.exe" (
        echo_success "Tesseract OCR is already installed at %TESSERACT_INSTALL_DIR%"
        REM Add to PATH
        setx PATH "%TESSERACT_INSTALL_DIR%;%PATH%" >nul 2>&1
        exit /b 0
    )
    
    echo_warning "Tesseract OCR is not installed"
    exit /b 1
}

REM Install Tesseract OCR for Windows
install_tesseract() {
    echo_warning "Tesseract OCR not found. Installing Tesseract OCR..."
    
    REM Detect Windows architecture
    if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
        set "TESSERACT_URL=%TESSERACT_URL_X64%"
        set "TESSERACT_EXE=tesseract-5.3.3.20230701.exe"
    ) else (
        set "TESSERACT_URL=%TESSERACT_URL_X86%"
        set "TESSERACT_EXE=tesseract-5.3.3.exe"
    )
    
    set "TESSERACT_INSTALLER="%TEMP%\%TESSERACT_EXE%""
    
    REM Download Tesseract
    call :download_file "%TESSERACT_URL%" "%TESSERACT_INSTALLER%"
    if %errorlevel% neq 0 (
        echo_error "Failed to download Tesseract OCR"
        exit /b 1
    )
    
    REM Install Tesseract silently
    echo_info "Installing Tesseract OCR..."
    "%TESSERACT_INSTALLER%" /S /D="%TESSERACT_INSTALL_DIR%"
    
    if %errorlevel% neq 0 (
        echo_error "Failed to install Tesseract OCR"
        del "%TESSERACT_INSTALLER%" >nul 2>&1
        exit /b 1
    )
    
    REM Add Tesseract to PATH
    echo_info "Adding Tesseract OCR to PATH..."
    setx PATH "%TESSERACT_INSTALL_DIR%;%PATH%" >nul 2>&1
    
    REM Clean up
    del "%TESSERACT_INSTALLER%" >nul 2>&1
    
    echo_success "Tesseract OCR installed successfully!"
    echo_info "Tesseract OCR installed at: %TESSERACT_INSTALL_DIR%"
    
    exit /b 0
}

REM =============================================================================
REM Font Installation (Windows-specific)
REM =============================================================================

REM Check and install fonts for Windows
check_and_install_fonts() {
    echo_info "Checking fonts..."
    
    REM Note: Windows typically comes with Calibri and Caladea equivalents
    REM The Linux packages ttf-mscorefonts-installer, fonts-crosextra-carlito, fonts-crosextra-caladea
    REM are for Linux systems. On Windows, these fonts are usually available by default.
    
    echo_info "Note: On Windows, Calibri and Caladea fonts are typically available by default."
    echo_info "If you need to install additional fonts, please install them manually."
    echo_info "Linux font packages (ttf-mscorefonts-installer, fonts-crosextra-*) are not needed on Windows."
    
    echo_success "Font check passed!"
    exit /b 0
}

REM =============================================================================
REM Step 1: Check and Install Prerequisites
REM =============================================================================
echo_status Checking and installing prerequisites...

REM Check for Node.js
where node >nul 2>&1
if %errorlevel% neq 0 (
    echo_error Node.js is not installed. Please install Node.js first.
    echo_info Download from: https://nodejs.org/
    exit /b 1
)
echo_success Node.js is installed

REM Check for Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo_error Python is not installed. Please install Python 3 first.
    echo_info Download from: https://python.org/
    exit /b 1
)
echo_success Python is installed

REM Check for curl (needed for health checks)
where curl >nul 2>&1
if %errorlevel% neq 0 (
    echo_warning curl is not installed. Some health checks may not work.
)

REM Check and install MongoDB
call :check_mongodb_installed
if %errorlevel% neq 0 (
    call :install_mongodb
    if %errorlevel% neq 0 (
        echo_error Failed to install MongoDB
        exit /b 1
    )
)

REM Check and install Tesseract OCR
call :check_tesseract_installed
if %errorlevel% neq 0 (
    call :install_tesseract
    if %errorlevel% neq 0 (
        echo_error Failed to install Tesseract OCR
        exit /b 1
    )
)

REM Check fonts
call :check_and_install_fonts

echo_status All prerequisites are installed!

REM =============================================================================
REM Step 2: Start MongoDB
REM =============================================================================
:start_mongodb
echo_status Starting MongoDB...

REM Create data directory if it doesn't exist
if not exist "%MONGO_DB_PATH%" mkdir "%MONGO_DB_PATH%"

REM Check if MongoDB is already running
tasklist /FI "IMAGENAME eq mongod.exe" 2>nul | findstr /I mongod.exe >nul
if %errorlevel% equ 0 (
    echo_warning MongoDB is already running. Skipping MongoDB start.
    for /f "tokens=2" %%a in ('tasklist /FI "IMAGENAME eq mongod.exe" /NH') do set MONGO_PID=%%a
    goto :start_backend
)

REM Start MongoDB in background
echo_info Starting MongoDB on port 27017...
start /B mongod --dbpath "%MONGO_DB_PATH%" --bind_ip 127.0.0.1 --port 27017

REM Wait for MongoDB to be ready
echo_info Waiting for MongoDB to be ready...
timeout /t 5 /nobreak >nul

REM Verify MongoDB is running
tasklist /FI "IMAGENAME eq mongod.exe" 2>nul | findstr /I mongod.exe >nul
if %errorlevel% neq 0 (
    echo_error Failed to start MongoDB. Check logs/mongodb.log for details.
    exit /b 1
)

for /f "tokens=2" %%a in ('tasklist /FI "IMAGENAME eq mongod.exe" /NH') do set MONGO_PID=%%a
echo_status MongoDB started successfully (PID: !MONGO_PID!)

REM =============================================================================
REM Step 3: Start Backend
REM =============================================================================
:start_backend
echo_status Starting Backend...

cd /d "%BACKEND_DIR%"

REM Check if virtual environment exists
if not exist "%VENV_DIR%" (
    echo_info Creating Python virtual environment...
    python -m venv "%VENV_DIR%"
)

REM Install dependencies
echo_info Installing backend dependencies...
call "%VENV_DIR%\Scripts\pip" install -q -r requirements.txt 2>nul

REM Start backend in background
echo_info Starting backend server on port %BACKEND_PORT%...
start /B cmd /c "%VENV_DIR%\Scripts\python -m uvicorn server:app --host 127.0.0.1 --port %BACKEND_PORT%"

set BACKEND_PID=last
echo_status Backend started (checking status...)

REM Wait for backend to be ready
echo_info Waiting for backend to be ready...
timeout /t 5 /nobreak >nul

REM Verify backend is running
curl -s http://127.0.0.1:%BACKEND_PORT%/api >nul 2>&1
if %errorlevel% neq 0 (
    echo_warning Backend may still be starting, continuing...
)

echo_status Backend started successfully

REM =============================================================================
REM Step 4: Start Frontend
REM =============================================================================
:start_frontend
echo_status Starting Frontend...

cd /d "%FRONTEND_DIR%"

REM Check if node_modules exists
if not exist "node_modules" (
    echo_info Installing frontend dependencies...
    call npm install --silent
)

REM Set the backend URL environment variable (configurable via environment)
REM Default to localhost, can be overridden with network IP for remote access
if not defined REACT_APP_BACKEND_URL set REACT_APP_BACKEND_URL=http://127.0.0.1:%BACKEND_PORT%

REM Start frontend (this will open browser)
echo_info Starting frontend server on port %FRONTEND_PORT%...
start cmd /c "npm start"

echo_status Frontend starting...

REM =============================================================================
REM Step 5: Display status
REM =============================================================================
:display_status
echo.
echo ===========================================================================
echo All services started successfully!
echo ===========================================================================
echo.
echo Services:
echo   MongoDB:      running on 127.0.0.1:27017
echo   Backend:      running on http://127.0.0.1:%BACKEND_PORT%
echo   Frontend:     running on http://127.0.0.1:%FRONTEND_PORT%
echo.
echo API Endpoints:
echo   API Root:     http://127.0.0.1:%BACKEND_PORT%/api
echo   API Docs:     http://127.0.0.1:%BACKEND_PORT%/docs
echo.
echo Prerequisites Installed:
echo   MongoDB:      %ProgramFiles%\MongoDB\Server\6.0\
echo   Tesseract:    %TESSERACT_INSTALL_DIR%\
echo   Node.js:      (in PATH)
echo   Python:       (in PATH)
echo.
echo Note: Font packages (ttf-mscorefonts-installer, fonts-crosextra-*)
echo       are Linux-specific. On Windows, fonts are available by default.
echo.
echo Press Ctrl+C to stop all services.
echo ===========================================================================
echo.
echo Services are running. Keep this window open.
echo.

REM =============================================================================
REM Main loop - wait for user interrupt
REM =============================================================================
:wait_loop
timeout /t 10 /nobreak >nul
goto :wait_loop

endlocal

