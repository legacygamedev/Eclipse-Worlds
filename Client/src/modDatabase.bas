Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Private Const LOCALE_USER_DEFAULT& = &H400
    Private Const LOCALE_SDECIMAL& = &HE
    Private Const LOCALE_STHOUSAND& = &HF
    Private Declare Function GetLocaleInfo& Lib "kernel32" Alias _
        "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
        ByVal lpLCData As String, ByVal cchData As Long)
        
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
    ByVal lpVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long) As Long
    
Private Function DecimalSeparator() As String
      Dim R As Long, S As String
      S = String(10, "a")
      R = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, S, 10)
      DecimalSeparator = Left$(S, R)
End Function

Public Sub HandleError(ByVal ProcName As String, ByVal ContName As String, ByVal ErNumber, ByVal ErDesc, ByVal ErSource, ByVal ErHelpContext)
    Dim FileName As String, F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ChkDir(App.Path & "\logs\", Month(Now) & "-" & Day(Now) & "-" & year(Now))
    FileName = App.Path & "\logs\" & Month(Now) & "-" & Day(Now) & "-" & year(Now) & "\Errors.txt"
    
    If Not FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & ProcName & "' In '" & ContName & "'."
        Print #1, "Run-time error '" & ErNumber & "': " & ErDesc & "."
        Print #1, ""
    Close #1
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not RAW Then
        If Len(Dir(App.Path & FileName)) > 0 Then
            FileExist = True
        End If
    Else
        If Len(Dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If
    Exit Function
    
' Error handler
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function
Private Function InternationalizeDoubles(Value As String) As String
    InternationalizeDoubles = Value
    Dim i As Long, B() As Byte, dotsCounter As Long, commasCounter As Long, test As Double
    B = Value
    For i = 0 To UBound(B) Step 2
        If B(i) = 44 Then
            commasCounter = commasCounter + 1
            Mid(Value, i / 2 + 1, 1) = DecimalSeparator
        ElseIf B(i) = 46 Then
            dotsCounter = dotsCounter + 1
            Mid(Value, i / 2 + 1, 1) = DecimalSeparator
        ElseIf B(i) >= 48 And B(i) <= 57 Then
        
        Else
            Exit Function
        End If
    Next i
    If (commasCounter <> 0 And dotsCounter <> 0) Or (commasCounter > 1) Or (dotsCounter > 1) Then
        Exit Function
    End If
    test = CDbl(Value)
    InternationalizeDoubles = test
End Function
' Gets a string from a text File
Public Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default Value if not found
    Dim retrivedValue As String, test As Boolean
        ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    retrivedValue = Left$(GetVar, Len(GetVar) - 1)
    If InStr(retrivedValue, ",") <> 0 Or InStr(retrivedValue, ".") <> 0 Then
        retrivedValue = InternationalizeDoubles(retrivedValue)
    End If

    GetVar = retrivedValue
    Exit Function
    
' Error handler
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

' Writes a variable to a text File
Public Sub PutVar(file As String, Header As String, Var As String, Value As String)
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call WritePrivateProfileString$(Header, Var, Value, file)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub SaveOptions()
    Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\data files\config.ini"
    
    Call PutVar(FileName, "Options", "Username", Trim$(Options.UserName))
    Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
    Call PutVar(FileName, "Options", "SaveUsername", Trim$(Options.SaveUsername))
    Call PutVar(FileName, "Options", "SavePassword", Trim$(Options.SavePassword))
    Call PutVar(FileName, "Options", "Website", Trim$(Options.Website))
    Call PutVar(FileName, "Options", "IP", Trim$(Options.IP))
    Call PutVar(FileName, "Options", "Port", Trim$(Options.Port))
    Call PutVar(FileName, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(FileName, "Options", "Music", Trim$(Options.Music))
    Call PutVar(FileName, "Options", "Sound", Trim$(Options.Sound))
    Call PutVar(FileName, "Options", "WASD", Trim$(Options.WASD))
    Call PutVar(FileName, "Options", "Level", Trim$(Options.Levels))
    Call PutVar(FileName, "Options", "Guilds", Trim$(Options.Guilds))
    Call PutVar(FileName, "Options", "PlayerVitals", Trim$(Options.PlayerVitals))
    Call PutVar(FileName, "Options", "NpcVitals", Trim$(Options.NpcVitals))
    Call PutVar(FileName, "Options", "Titles", Trim$(Options.Titles))
    Call PutVar(FileName, "Options", "BattleMusic", Trim$(Options.BattleMusic))
    Call PutVar(FileName, "Options", "Mouse", Trim$(Options.Mouse))
    Call PutVar(FileName, "Options", "Debug", Trim$(Options.Debug))
    Call PutVar(FileName, "Options", "SwearFilter", Trim$(Options.SwearFilter))
    Call PutVar(FileName, "Options", "Weather", Trim$(Options.Weather))
    Call PutVar(FileName, "Options", "AutoTile", Trim$(Options.Autotile))
    Call PutVar(FileName, "Options", "Blood", Trim$(Options.Blood))
    Call PutVar(FileName, "Options", "MusicVolume", Trim$(Options.MusicVolume))
    Call PutVar(FileName, "Options", "SoundVolume", Trim$(Options.SoundVolume))
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub LoadOptionVariables()
    Dim FileName As String
    
    FileName = App.Path & "\data files\config.ini"

    ' Load options
    If GetVar(FileName, "Options", "Username") = "" Then
        Options.UserName = vbNullString
        Call PutVar(FileName, "Options", "Username", Trim$(Options.UserName))
    Else
        Options.UserName = GetVar(FileName, "Options", "Username")
    End If
    
    If GetVar(FileName, "Options", "Password") = "" Then
        Options.Password = vbNullString
        Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
    Else
        Options.Password = GetVar(FileName, "Options", "Password")
    End If
    
    If GetVar(FileName, "Options", "SaveUsername") = "" Then
        Options.SaveUsername = "1"
        Call PutVar(FileName, "Options", "SaveUsername", Trim$(Options.SaveUsername))
    Else
        Options.SaveUsername = GetVar(FileName, "Options", "SaveUsername")
    End If
    
    If GetVar(FileName, "Options", "SavePassword") = "" Then
        Options.SavePassword = "0"
        Call PutVar(FileName, "Options", "SavePassword", Trim$(Options.SavePassword))
    Else
        Options.SavePassword = GetVar(FileName, "Options", "SavePassword")
    End If
    
    If GetVar(FileName, "Options", "Website") = "" Then
        Options.Website = GAME_WEBSITE
        Call PutVar(FileName, "Options", "Website", Trim$(Options.Website))
    Else
        Options.Website = GetVar(FileName, "Options", "Website")
    End If
    
    If GetVar(FileName, "Options", "IP") = "" Then
        Options.IP = "127.0.0.1"
        Call PutVar(FileName, "Options", "IP", Trim$(Options.IP))
    Else
        Options.IP = GetVar(FileName, "Options", "IP")
    End If
    
    If GetVar(FileName, "Options", "Port") = "" Then
        Options.Port = "7001"
        Call PutVar(FileName, "Options", "Port", Trim$(Options.Port))
    Else
        Options.Port = GetVar(FileName, "Options", "Port")
    End If
    
    If GetVar(FileName, "Options", "MenuMusic") = "" Then
        Options.MenuMusic = "Victoriam Speramus"
        Call PutVar(FileName, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Else
        Options.MenuMusic = GetVar(FileName, "Options", "MenuMusic")
    End If
    
    If GetVar(FileName, "Options", "Music") = "" Then
        Options.Music = "1"
        Call PutVar(FileName, "Options", "Music", Trim$(Options.Music))
    Else
        Options.Music = GetVar(FileName, "Options", "Music")
    End If
    
    If GetVar(FileName, "Options", "Sound") = "" Then
        Options.Sound = "1"
        Call PutVar(FileName, "Options", "Sound", Trim$(Options.Sound))
    Else
        Options.Sound = GetVar(FileName, "Options", "Sound")
    End If

    If GetVar(FileName, "Options", "WASD") = "" Then
        Options.WASD = "0"
        Call PutVar(FileName, "Options", "WASD", Trim$(Options.WASD))
    Else
        Options.WASD = GetVar(FileName, "Options", "WASD")
    End If
    
    If GetVar(FileName, "Options", "Level") = "" Then
        Options.Levels = "1"
        Call PutVar(FileName, "Options", "Level", Trim$(Options.Levels))
    Else
        Options.Levels = GetVar(FileName, "Options", "Level")
    End If
    
    If GetVar(FileName, "Options", "Guilds") = "" Then
        Options.Guilds = "1"
        Call PutVar(FileName, "Options", "Guilds", Trim$(Options.Guilds))
    Else
        Options.Guilds = GetVar(FileName, "Options", "Guilds")
    End If
    
    If GetVar(FileName, "Options", "PlayerVitals") = "" Then
        Options.PlayerVitals = "1"
        Call PutVar(FileName, "Options", "PlayerVitals", Trim$(Options.PlayerVitals))
    Else
        Options.PlayerVitals = GetVar(FileName, "Options", "PlayerVitals")
    End If
    
    If GetVar(FileName, "Options", "NpcVitals") = "" Then
        Options.NpcVitals = "1"
        Call PutVar(FileName, "Options", "NpcVitals", Trim$(Options.NpcVitals))
    Else
        Options.NpcVitals = GetVar(FileName, "Options", "NpcVitals")
    End If
    
    If GetVar(FileName, "Options", "Titles") = "" Then
        Options.Titles = "1"
        Call PutVar(FileName, "Options", "Titles", Trim$(Options.Titles))
    Else
        Options.Titles = GetVar(FileName, "Options", "Titles")
    End If
    
    If GetVar(FileName, "Options", "BattleMusic") = "" Then
        Options.BattleMusic = "1"
        Call PutVar(FileName, "Options", "BattleMusic", Trim$(Options.BattleMusic))
    Else
        Options.BattleMusic = GetVar(FileName, "Options", "BattleMusic")
    End If
    
    If GetVar(FileName, "Options", "Mouse") = "" Then
        Options.Mouse = "0"
        Call PutVar(FileName, "Options", "Mouse", Trim$(Options.Mouse))
    Else
        Options.Mouse = GetVar(FileName, "Options", "Mouse")
    End If
    
    If GetVar(FileName, "Options", "Debug") = "" Then
        Options.Debug = "1"
        Call PutVar(FileName, "Options", "Debug", Trim$(Options.Debug))
    Else
        Options.Debug = GetVar(FileName, "Options", "Debug")
    End If
    
    If GetVar(FileName, "Options", "SwearFilter") = "" Then
        Options.SwearFilter = "1"
        Call PutVar(FileName, "Options", "SwearFilter", Trim$(Options.SwearFilter))
    Else
        Options.SwearFilter = GetVar(FileName, "Options", "SwearFilter")
    End If
    
    If GetVar(FileName, "Options", "Weather") = "" Then
        Options.Weather = "1"
        Call PutVar(FileName, "Options", "Weather", Trim$(Options.Weather))
    Else
        Options.Weather = GetVar(FileName, "Options", "Weather")
    End If
    
    If GetVar(FileName, "Options", "AutoTile") = "" Then
        Options.Autotile = "1"
        Call PutVar(FileName, "Options", "AutoTile", Trim$(Options.Autotile))
    Else
        Options.Autotile = GetVar(FileName, "Options", "AutoTile")
    End If
    
    If GetVar(FileName, "Options", "Blood") = "" Then
        Options.Blood = "1"
        Call PutVar(FileName, "Options", "Blood", Trim$(Options.Blood))
    Else
        Options.Blood = GetVar(FileName, "Options", "Blood")
    End If
    
    If GetVar(FileName, "Options", "MusicVolume") = "" Then
        Options.MusicVolume = InternationalizeDoubles("0.5")
        Call PutVar(FileName, "Options", "MusicVolume", Trim$(Options.MusicVolume))
    Else
        Options.MusicVolume = GetVar(FileName, "Options", "MusicVolume")
    End If
    
    If GetVar(FileName, "Options", "SoundVolume") = "" Then
        Options.SoundVolume = InternationalizeDoubles("0.8")
        Call PutVar(FileName, "Options", "SoundVolume", Trim$(Options.SoundVolume))
    Else
        Options.SoundVolume = GetVar(FileName, "Options", "SoundVolume")
    End If
End Sub

Public Sub LoadOptions()
    ' Load the variables in the options.ini
    Call LoadOptionVariables
    
    ' Set the form items based on what the options are
    ResetOptionButtons
End Sub

Public Function TimeStamp() As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TimeStamp = "[" & time & "]"
    Exit Function
    
' Error handler
errorhandler:
    HandleError "TimeStamp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub AddLog(ByVal text As String, ByVal LogFile As String)
    Dim FileName As String
    Dim F As Integer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ChkDir(App.Path & "\logs\", Month(Now) & "-" & Day(Now) & "-" & year(Now))
    FileName = App.Path & "\logs\" & Month(Now) & "-" & Day(Now) & "-" & year(Now) & "\" & LogFile & ".log"

    If Not FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    F = FreeFile
    
    Open FileName For Append As #F
        Print #F, TimeStamp & " - " & text
    Close #F
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "AddLog", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub LoadAnimatedSprites()
    Dim i As Integer, n As Integer
    Dim TmpArray() As String
    
    If AnimatedSpriteNumbers = vbNullString Then Exit Sub
    
    ' Split into an array of strings
    TmpArray() = Split(AnimatedSpriteNumbers, ",")

    ReDim AnimatedSprites(1 To NumCharacters)

    ' Loop through converting strings to values and store in the sprite array
    For i = 1 To NumCharacters
        For n = 0 To UBound(TmpArray)
            If i = Trim$(TmpArray(n)) Then
                AnimatedSprites(i) = 1
            End If
        Next
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "LoadAnimatedSprites", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckTilesets()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumTileSets = 1
    
    ReDim Tex_Tileset(1)

    While FileExist(GFX_PATH & "tilesets\" & i & GFX_EXT)
        ReDim Preserve Tex_Tileset(NumTileSets)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Tileset(NumTileSets).filepath = App.Path & GFX_PATH & "tilesets\" & i & GFX_EXT
        Tex_Tileset(NumTileSets).Texture = NumTextures
        NumTileSets = NumTileSets + 1
        i = i + 1
    Wend
    
    NumTileSets = NumTileSets - 1
    
    If NumTileSets < 1 Then Exit Sub
    
    For i = 1 To NumTileSets
        LoadTexture Tex_Tileset(i)
    Next
    Exit Sub

' Error handler
errorhandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckCharacters()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumCharacters = 1
    
    ReDim Tex_Character(1)
    Dim test As String
    test = Dir(GFX_PATH & "characters\" & "*" & GFX_EXT, vbNormal)
    
    While FileExist(GFX_PATH & "characters\" & i & GFX_EXT)
        ReDim Preserve Tex_Character(NumCharacters)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Character(NumCharacters).filepath = App.Path & GFX_PATH & "characters\" & i & GFX_EXT
        Tex_Character(NumCharacters).Texture = NumTextures
        NumCharacters = NumCharacters + 1
        i = i + 1
    Wend
    
    NumCharacters = NumCharacters - 1
    
    If NumCharacters < 1 Then Exit Sub
    
    For i = 1 To NumCharacters
        LoadTexture Tex_Character(i)
    Next
    
    ' Load the animated sprite numbers used in animating sprites
    Call LoadAnimatedSprites
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckPaperdolls()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumPaperdolls = 1
    
    ReDim Tex_Paperdoll(1)

    While FileExist(GFX_PATH & "paperdolls\" & i & GFX_EXT)
        ReDim Preserve Tex_Paperdoll(NumPaperdolls)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Paperdoll(NumPaperdolls).filepath = App.Path & GFX_PATH & "paperdolls\" & i & GFX_EXT
        Tex_Paperdoll(NumPaperdolls).Texture = NumTextures
        NumPaperdolls = NumPaperdolls + 1
        i = i + 1
    Wend
    
    NumPaperdolls = NumPaperdolls - 1
    
    If NumPaperdolls < 1 Then Exit Sub
    
    For i = 1 To NumPaperdolls
        LoadTexture Tex_Paperdoll(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckAnimations()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumAnimations = 1
    
    ReDim Tex_Animation(1)

    While FileExist(GFX_PATH & "animations\" & i & GFX_EXT)
        ReDim Preserve Tex_Animation(NumAnimations)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Animation(NumAnimations).Texture = NumTextures
        Tex_Animation(NumAnimations).filepath = App.Path & GFX_PATH & "animations\" & i & GFX_EXT
        NumAnimations = NumAnimations + 1
        i = i + 1
    Wend
    
    NumAnimations = NumAnimations - 1
    
    If NumAnimations < 1 Then Exit Sub

    For i = 1 To NumAnimations
        LoadTexture Tex_Animation(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckItems()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumItems = 1
    
    ReDim Tex_Item(1)

    While FileExist(GFX_PATH & "items\" & i & GFX_EXT)
        ReDim Preserve Tex_Item(NumItems)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Item(NumItems).filepath = App.Path & GFX_PATH & "items\" & i & GFX_EXT
        Tex_Item(NumItems).Texture = NumTextures
        NumItems = NumItems + 1
        i = i + 1
    Wend
    
    NumItems = NumItems - 1
    
    If NumItems < 1 Then Exit Sub
    
    For i = 1 To NumItems
        LoadTexture Tex_Item(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckResources()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumResources = 1
    
    ReDim Tex_Resource(1)

    While FileExist(GFX_PATH & "resources\" & i & GFX_EXT)
        ReDim Preserve Tex_Resource(NumResources)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Resource(NumResources).filepath = App.Path & GFX_PATH & "resources\" & i & GFX_EXT
        Tex_Resource(NumResources).Texture = NumTextures
        NumResources = NumResources + 1
        i = i + 1
    Wend
    
    NumResources = NumResources - 1
    
    If NumResources < 1 Then Exit Sub
    
    For i = 1 To NumResources
        LoadTexture Tex_Resource(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckSpellIcons()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumSpellIcons = 1
    
    ReDim Tex_SpellIcon(1)

    While FileExist(GFX_PATH & "spellicons\" & i & GFX_EXT)
        ReDim Preserve Tex_SpellIcon(NumSpellIcons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_SpellIcon(NumSpellIcons).filepath = App.Path & GFX_PATH & "spellicons\" & i & GFX_EXT
        Tex_SpellIcon(NumSpellIcons).Texture = NumTextures
        NumSpellIcons = NumSpellIcons + 1
        i = i + 1
    Wend

    NumSpellIcons = NumSpellIcons - 1
    
    If NumSpellIcons < 1 Then Exit Sub
    
    For i = 1 To NumSpellIcons
        LoadTexture Tex_SpellIcon(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckFaces()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumFaces = 1
    
    ReDim Tex_Face(1)

    While FileExist(GFX_PATH & "Faces\" & i & GFX_EXT)
        ReDim Preserve Tex_Face(NumFaces)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Face(NumFaces).filepath = App.Path & GFX_PATH & "faces\" & i & GFX_EXT
        Tex_Face(NumFaces).Texture = NumTextures
        NumFaces = NumFaces + 1
        i = i + 1
    Wend
    
    NumFaces = NumFaces - 1
     
    If NumFaces < 1 Then Exit Sub
    
    For i = 1 To NumFaces
        LoadTexture Tex_Face(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckFogs()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumFogs = 1
    
    ReDim Tex_Fog(1)
    
    While FileExist(GFX_PATH & "fogs\" & i & GFX_EXT)
        ReDim Preserve Tex_Fog(NumFogs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Fog(NumFogs).filepath = App.Path & GFX_PATH & "fogs\" & i & GFX_EXT
        Tex_Fog(NumFogs).Texture = NumTextures
        NumFogs = NumFogs + 1
        i = i + 1
    Wend
    
    NumFogs = NumFogs - 1
    
    If NumFogs < 1 Then Exit Sub
    
    For i = 1 To NumFogs
        LoadTexture Tex_Fog(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CheckFogs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckPanoramas()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumPanoramas = 1
    
    ReDim Tex_Panorama(1)
    While FileExist(GFX_PATH & "Panoramas\" & i & GFX_EXT)
        ReDim Preserve Tex_Panorama(NumPanoramas)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Panorama(NumPanoramas).filepath = App.Path & GFX_PATH & "Panoramas\" & i & GFX_EXT
        Tex_Panorama(NumPanoramas).Texture = NumTextures
        NumPanoramas = NumPanoramas + 1
        i = i + 1
    Wend
    
    NumPanoramas = NumPanoramas - 1
    
    If NumPanoramas < 1 Then Exit Sub
    
    For i = 1 To NumPanoramas
        LoadTexture Tex_Panorama(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CheckPanoramas", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckEmoticons()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    i = 1
    NumEmoticons = 1
    
    ReDim Tex_Emoticon(1)

    While FileExist(GFX_PATH & "Emoticons\" & i & GFX_EXT)
        ReDim Preserve Tex_Emoticon(NumEmoticons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Emoticon(NumEmoticons).filepath = App.Path & GFX_PATH & "Emoticons\" & i & GFX_EXT
        Tex_Emoticon(NumEmoticons).Texture = NumTextures
        NumEmoticons = NumEmoticons + 1
        i = i + 1
    Wend
    
    NumEmoticons = NumEmoticons - 1
    
    If NumEmoticons < 1 Then Exit Sub
    
    For i = 1 To NumEmoticons
        LoadTexture Tex_Emoticon(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CheckEmoticons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearPlayer(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    Player(Index).Name = vbNullString
    Player(Index).Status = vbNullString
    Player(Index).Class = 1
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = vbNullString
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearItems()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = vbNullString
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearAnimations()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearNPC(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).title = vbNullString
    NPC(Index).AttackSay = vbNullString
    NPC(Index).Music = vbNullString
    NPC(Index).Sound = vbNullString
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearNpcs()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearSpell(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = vbNullString
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearSpells()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearShop(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearShops()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearResource(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).FailMessage = vbNullString
    Resource(Index).Sound = vbNullString
    Exit Sub
    
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearResources()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMapItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
    MapItem(Index).PlayerName = vbNullString
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    Map.Music = vbNullString
    Map.BGS = vbNullString
    Map.Moral = 1
    Map.MaxX = MIN_MAPX
    Map.MaxY = MIN_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    InitAutotiles
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMapItems()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNPC(Index)), LenB(MapNPC(Index)))
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMapNpcs()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearBans()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To MAX_BANS
        Call ClearBan(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearBans", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearBan(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ZeroMemory(ByVal VarPtr(Ban(Index)), LenB(Ban(Index)))
    Ban(Index).PlayerLogin = vbNullString
    Ban(Index).PlayerName = vbNullString
    Ban(Index).Reason = vbNullString
    Ban(Index).IP = vbNullString
    Ban(Index).HDSerial = vbNullString
    Ban(Index).time = vbNullString
    Ban(Index).By = vbNullString
    Ban(Index).Date = vbNullString
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearBan", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearTitles()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To MAX_TITLES
        Call ClearTitle(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearTitles", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearTitle(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ZeroMemory(ByVal VarPtr(title(Index)), LenB(title(Index)))
    title(Index).Name = vbNullString
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearTitle", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMoral(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Moral(Index)), LenB(Moral(Index)))
    Moral(Index).Name = vbNullString
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearMoral", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMorals()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MORALS
        Call ClearMoral(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearMorals", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearClass(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ZeroMemory(ByVal VarPtr(Class(Index)), LenB(Class(Index)))
    Class(Index).Name = vbNullString
    Class(Index).CombatTree = 1
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearClasses()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_CLASSES
        Call ClearClass(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearClasses", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearEmoticon(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Emoticon(Index)), LenB(Emoticon(Index)))
    Emoticon(Index).Command = "/"
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearEmoticon", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearEmoticons()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_EMOTICONS
        Call ClearEmoticon(i)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearEmoticons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearEvents()
    Dim i As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To MAX_EVENTS
        Call ClearEvent(i)
    Next i
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearEvents", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearEvent(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Index <= 0 Or Index > MAX_EVENTS Then Exit Sub
    
    Call ZeroMemory(ByVal VarPtr(events(Index)), LenB(events(Index)))
    events(Index).Name = vbNullString
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearEvent", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

