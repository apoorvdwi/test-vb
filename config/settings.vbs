' settings.vbs - Global settings and configuration
' ------------------------------------------------------------

Option Explicit

' BlueZone Configuration
Const BLUEZONE_SESSION_CONFIG = "mainframe.zmd"  ' Change to your BlueZone session file

' Mainframe Credentials
Const MAINFRAME_USERNAME = "username"  ' Replace with actual username
Const MAINFRAME_PASSWORD = "password"  ' Replace with actual password

' API Configuration
Const API_BASE_URL = "https://api.example.com"
Const API_KEY = "your-api-key-here"  ' Replace with your actual API key
Const AUTH_TOKEN = "your-auth-token-here"  ' Replace with your actual auth token
Const API_ENDPOINT = API_BASE_URL & "/data"
Const API_TIMEOUT = 30  ' Timeout in seconds
Const AUTH_TYPE = "Bearer"  ' Authentication type (Bearer, Basic, etc.)

' File System Paths
Const OUTPUT_DIRECTORY = "output"
Const LOG_DIRECTORY = "logs"

' Application Settings
Const DEBUG_MODE = True  ' Set to False in production
Const MAX_RETRY_ATTEMPTS = 3
Const RETRY_DELAY = 1000  ' Milliseconds between retry attempts

' Initialize global objects
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

' Ensure output directories exist
If Not FSO.FolderExists(OUTPUT_DIRECTORY) Then
    FSO.CreateFolder(OUTPUT_DIRECTORY)
End If

If Not FSO.FolderExists(LOG_DIRECTORY) Then
    FSO.CreateFolder(LOG_DIRECTORY)
End If
