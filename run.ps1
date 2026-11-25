$pythonPath = "C:\Users\PC\AppData\Local\Programs\Python\Python312\python.exe"
if (Test-Path $pythonPath) {
    & $pythonPath (Join-Path $PSScriptRoot "main.py")
} else {
    python (Join-Path $PSScriptRoot "main.py")
}

