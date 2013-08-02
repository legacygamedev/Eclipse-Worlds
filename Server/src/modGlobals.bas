Attribute VB_Name = "modGlobals"
Option Explicit

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back NPCs hp
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

' Text variables
Public vbQuote As String

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public IsShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Long

' High Indexing
Public Player_HighIndex As Long

' CPS Lock
Public CPSUnlock As Boolean
