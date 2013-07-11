Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

' For Clear functions
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> LCase$(tName) Then Call MkDir(LCase$(tDir & "\" & tName))
End Sub

' Outputs string to text file
Public Function TimeStamp() As String
    TimeStamp = "[" & Time & "]"
End Function

Public Sub AddLog(ByVal Text As String, ByVal LogFile As String)
    Dim FileName As String
    Dim F As Integer

    Call ChkDir(App.path & "\", "logs")
    Call ChkDir(App.path & "\logs\", Month(Now) & "-" & Day(Now) & "-" & Year(Now))
    FileName = App.path & "\logs\" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "\" & LogFile & ".log"

    If Not FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    F = FreeFile
    
    Open FileName For Append As #F
        Print #F, TimeStamp & " - " & Text
    Close #F
End Sub

' Gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' Writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    If Not RAW Then
        If Len(Dir(App.path & FileName)) > 0 Then
            FileExist = True
        End If
    Else
        If Len(Dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If
End Function

Public Sub InitOptions()
    Dim FileName As String
    
    ' File name used for options
    FileName = App.path & "\data\options.ini"
    
    ' Game Name
    If GetVar(FileName, "Options", "Name") = "" Then
        Options.Name = "Legends of Arteix"
        Call PutVar(FileName, "Options", "Name", Trim$(Options.Name))
    Else
        Options.Name = GetVar(FileName, "Options", "Name")
    End If
    
    ' Website
    If GetVar(FileName, "Options", "Website") = "" Then
        Options.Website = "http://www.arteixinc.com/"
        Call PutVar(FileName, "Options", "Website", Trim$(Options.Website))
    Else
        Options.Website = GetVar(FileName, "Options", "Website")
    End If
    
    ' Port
    If GetVar(FileName, "Options", "Port") = "" Then
        Options.Port = "7001"
        Call PutVar(FileName, "Options", "Port", Trim$(Options.Port))
    Else
        Options.Port = GetVar(FileName, "Options", "Port")
    End If
    
    ' Message of the Day
    If GetVar(FileName, "Options", "MOTD") = "" Then
        Options.MOTD = "Welcome to the Legends of Arteix!"
        Call PutVar(FileName, "Options", "MOTD", Trim$(Options.MOTD))
    Else
        Options.MOTD = GetVar(FileName, "Options", "MOTD")
    End If
    
    ' Staff Message of the Day
    If GetVar(FileName, "Options", "SMOTD") = "" Then
        Options.SMOTD = ""
        Call PutVar(FileName, "Options", "SMOTD", Trim$(Options.SMOTD))
    Else
        Options.SMOTD = GetVar(FileName, "Options", "SMOTD")
    End If

    ' Player Kill level
    If GetVar(FileName, "Options", "PKLevel") = "" Then
        Options.PKLevel = "10"
        Call PutVar(FileName, "Options", "PKLevel", Trim$(Options.PKLevel))
    Else
        Options.PKLevel = GetVar(FileName, "Options", "PKLevel")
    End If
    
    ' Same IP
    If GetVar(FileName, "Options", "MultipleIP") = "" Then
        Options.MultipleIP = "1"
        Call PutVar(FileName, "Options", "MultipleIP", Trim$(Options.MultipleIP))
    Else
        Options.MultipleIP = GetVar(FileName, "Options", "MultipleIP")
    End If
    
    ' Same Serial
    If GetVar(FileName, "Options", "MultipleSerial") = "" Then
        Options.MultipleSerial = "1"
        Call PutVar(FileName, "Options", "MultipleSerial", Trim$(Options.MultipleSerial))
    Else
        Options.MultipleSerial = GetVar(FileName, "Options", "MultipleSerial")
    End If
    
     ' Guild Cost
    If GetVar(FileName, "Options", "GuildCost") = "" Then
        Options.GuildCost = "5000"
        Call PutVar(FileName, "Options", "GuildCost", Trim$(Options.GuildCost))
    Else
        Options.GuildCost = GetVar(FileName, "Options", "GuildCost")
    End If
    
    ' News
    If GetVar(FileName, "Options", "News") = "" Then
        Options.News = "Welcome to the Legends of Arteix!"
        Call PutVar(FileName, "Options", "News", Trim$(Options.News))
    Else
        Options.News = GetVar(FileName, "Options", "News")
    End If
    
    ' Sound
    If GetVar(FileName, "Options", "MissSound") = "" Then
        Options.MissSound = "Miss2"
        Call PutVar(FileName, "Options", "MissSound", Trim$(Options.MissSound))
    Else
        Options.MissSound = GetVar(FileName, "Options", "MissSound")
    End If
    
    If GetVar(FileName, "Options", "DodgeSound") = "" Then
        Options.DodgeSound = "Dodge"
        Call PutVar(FileName, "Options", "DodgeSound", Trim$(Options.DodgeSound))
    Else
        Options.DodgeSound = GetVar(FileName, "Options", "DodgeSound")
    End If
    
    If GetVar(FileName, "Options", "DeflectSound") = "" Then
        Options.DeflectSound = "Saint3"
        Call PutVar(FileName, "Options", "DeflectSound", Trim$(Options.DeflectSound))
    Else
        Options.DeflectSound = GetVar(FileName, "Options", "DeflectSound")
    End If
    
    If GetVar(FileName, "Options", "BlockSound") = "" Then
        Options.BlockSound = "Block"
        Call PutVar(FileName, "Options", "BlockSound", Trim$(Options.BlockSound))
    Else
        Options.BlockSound = GetVar(FileName, "Options", "BlockSound")
    End If
    
    If GetVar(FileName, "Options", "CriticalSound") = "" Then
        Options.CriticalSound = "Critical"
        Call PutVar(FileName, "Options", "CriticalSound", Trim$(Options.CriticalSound))
    Else
        Options.CriticalSound = GetVar(FileName, "Options", "CriticalSound")
    End If
    
    If GetVar(FileName, "Options", "ResistSound") = "" Then
        Options.ResistSound = "Saint9"
        Call PutVar(FileName, "Options", "ResistSound", Trim$(Options.ResistSound))
    Else
        Options.ResistSound = GetVar(FileName, "Options", "ResistSound")
    End If
    
    If GetVar(FileName, "Options", "BuySound") = "" Then
        Options.BuySound = "Shop"
        Call PutVar(FileName, "Options", "BuySound", Trim$(Options.BuySound))
    Else
        Options.BuySound = GetVar(FileName, "Options", "BuySound")
    End If
    
    If GetVar(FileName, "Options", "SellSound") = "" Then
        Options.SellSound = "Sell"
        Call PutVar(FileName, "Options", "SellSound", Trim$(Options.SellSound))
    Else
        Options.SellSound = GetVar(FileName, "Options", "SellSound")
    End If
    
    ' Animations
    If GetVar(FileName, "Options", "DeflectAnimation") = "" Then
        Options.DeflectAnimation = 2
        Call PutVar(FileName, "Options", "DeflectAnimation", Trim$(Options.DeflectAnimation))
    Else
        Options.DeflectAnimation = GetVar(FileName, "Options", "DeflectAnimation")
    End If
    
    If GetVar(FileName, "Options", "CriticalAnimation") = "" Then
        Options.CriticalAnimation = 3
        Call PutVar(FileName, "Options", "CriticalAnimation", Trim$(Options.CriticalAnimation))
    Else
        Options.CriticalAnimation = GetVar(FileName, "Options", "CriticalAnimation")
    End If
    
    If GetVar(FileName, "Options", "DodgeAnimation") = "" Then
        Options.DodgeAnimation = 4
        Call PutVar(FileName, "Options", "DodgeAnimation", Trim$(Options.DodgeAnimation))
    Else
        Options.DodgeAnimation = GetVar(FileName, "Options", "DodgeAnimation")
    End If
End Sub

Public Sub SaveOptions()
    PutVar App.path & "\data\options.ini", "Options", "Name", Trim$(Options.Name)
    PutVar App.path & "\data\options.ini", "Options", "Port", Trim$(Options.Port)
    PutVar App.path & "\data\options.ini", "Options", "MOTD", Trim$(Options.MOTD)
    PutVar App.path & "\data\options.ini", "Options", "SMOTD", Trim$(Options.SMOTD)
    PutVar App.path & "\data\options.ini", "Options", "Website", Trim$(Options.Website)
    PutVar App.path & "\data\options.ini", "Options", "PKLevel", Trim$(Options.PKLevel)
    PutVar App.path & "\data\options.ini", "Options", "MultipleIP", Trim$(Options.MultipleIP)
    PutVar App.path & "\data\options.ini", "Options", "MultipleSerial", Trim$(Options.MultipleSerial)
    PutVar App.path & "\data\options.ini", "Options", "GuildCost", Trim$(Options.GuildCost)
    PutVar App.path & "\data\options.ini", "Options", "News", Trim$(Options.News)
    PutVar App.path & "\data\options.ini", "Options", "MissSound", Trim$(Options.MissSound)
    PutVar App.path & "\data\options.ini", "Options", "DodgeSound", Trim$(Options.DodgeSound)
    PutVar App.path & "\data\options.ini", "Options", "DeflectSound", Trim$(Options.DeflectSound)
    PutVar App.path & "\data\options.ini", "Options", "BlockSound", Trim$(Options.BlockSound)
    PutVar App.path & "\data\options.ini", "Options", "CriticalSound", Trim$(Options.CriticalSound)
    PutVar App.path & "\data\options.ini", "Options", "ResistSound", Trim$(Options.ResistSound)
    PutVar App.path & "\data\options.ini", "Options", "BuySound", Trim$(Options.BuySound)
    PutVar App.path & "\data\options.ini", "Options", "SellSound", Trim$(Options.SellSound)
    PutVar App.path & "\data\options.ini", "Options", "DeflectAnimation", Trim$(Options.DeflectAnimation)
    PutVar App.path & "\data\options.ini", "Options", "CriticalAnimation", Trim$(Options.CriticalAnimation)
    PutVar App.path & "\data\options.ini", "Options", "DodgeAnimation", Trim$(Options.DodgeAnimation)
End Sub

Public Sub LoadOptions()
    Options.Name = GetVar(App.path & "\data\options.ini", "Options", "Name")
    Options.Port = GetVar(App.path & "\data\options.ini", "Options", "Port")
    Options.MOTD = GetVar(App.path & "\data\options.ini", "Options", "MOTD")
    Options.Website = GetVar(App.path & "\data\options.ini", "Options", "Website")
    Options.PKLevel = GetVar(App.path & "\data\options.ini", "Options", "PKLevel")
    Options.MultipleIP = GetVar(App.path & "\data\options.ini", "Options", "MultipleIP")
    Options.MultipleSerial = GetVar(App.path & "\data\options.ini", "Options", "MultipleSerial")
    Options.GuildCost = GetVar(App.path & "\data\options.ini", "Options", "GuildCost")
    Options.News = GetVar(App.path & "\data\options.ini", "Options", "News")
    Options.MissSound = GetVar(App.path & "\data\options.ini", "Options", "MissSound")
    Options.DodgeSound = GetVar(App.path & "\data\options.ini", "Options", "DodgeSound")
    Options.DeflectSound = GetVar(App.path & "\data\options.ini", "Options", "DeflectSound")
    Options.BlockSound = GetVar(App.path & "\data\options.ini", "Options", "BlockSound")
    Options.CriticalSound = GetVar(App.path & "\data\options.ini", "Options", "CriticalSound")
    Options.ResistSound = GetVar(App.path & "\data\options.ini", "Options", "ResistSound")
    Options.BuySound = GetVar(App.path & "\data\options.ini", "Options", "BuySound")
    Options.SellSound = GetVar(App.path & "\data\options.ini", "Options", "SellSound")
    Options.DeflectAnimation = GetVar(App.path & "\data\options.ini", "Options", "DeflectAnimation")
    Options.CriticalAnimation = GetVar(App.path & "\data\options.ini", "Options", "CriticalAnimation")
    Options.DodgeAnimation = GetVar(App.path & "\data\options.ini", "Options", "DodgeAnimation")
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As String, ByVal Reason As String)
    Dim IP As String
    Dim i As Integer
    Dim n As Integer

    ' Cut off last portion of IP
    IP = GetPlayerIP(BanPlayerIndex)
    
    For i = Len(IP) To 1 Step -1
        If Mid$(IP, i, 1) = "." Then Exit For
    Next i

    IP = Mid$(IP, 1, i)

    For n = 1 To MAX_BANS
        If Not Len(Trim$(Ban(n).PlayerLogin)) > 0 And Not Len(Trim$(Ban(n).PlayerName)) > 0 Then
            With Ban(n)
                .Date = Date
                
                If BannedByIndex <> "server" Then
                    .By = GetPlayerName(BannedByIndex)
                Else
                    .By = "server"
                End If
                
                .Time = Time
                .HDSerial = GetPlayerHDSerial(BanPlayerIndex)
                .IP = IP
                .PlayerLogin = GetPlayerLogin(BanPlayerIndex)
                .PlayerName = GetPlayerName(BanPlayerIndex)
                .Reason = Reason
            End With
            Call SaveBan(n)
            Exit For
        End If
    Next n

    If Not BannedByIndex = "server" Then
        If Len(Reason) Then
            AdminMsg GetPlayerName(BanPlayerIndex) & " has been banned by " & GetPlayerName(BannedByIndex) & " for " & Reason & "!", BrightBlue
            AddLog GetPlayerName(BannedByIndex) & "/" & GetPlayerIP(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & "/" & GetPlayerIP(BanPlayerIndex) & " for " & Reason & ".", "Bans"
            AlertMsg BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & " for " & Reason & "!"
        Else
            AdminMsg GetPlayerName(BanPlayerIndex) & " has been banned by " & GetPlayerName(BannedByIndex) & "!", BrightBlue
            AddLog GetPlayerName(BannedByIndex) & "/" & GetPlayerIP(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & "/" & GetPlayerIP(BanPlayerIndex) & ".", "Admin"
            AlertMsg BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!"
        End If
    Else
        AdminMsg GetPlayerName(BanPlayerIndex) & " has been banned by the server!", BrightBlue
        AddLog GetPlayerName(BanPlayerIndex) & "/" & GetPlayerIP(BanPlayerIndex) & " was banned by the server!", "Admin"
        AlertMsg BanPlayerIndex, "You have been banned by the server!"
    End If
    Call LeftGame(BanPlayerIndex)
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim FileName As String
    
    Call ChkDir(App.path & "\data\accounts\", Trim(Name))
    FileName = "\data\accounts\" & Trim(Name) & "\data.bin"

    If FileExist(FileName) Then
        AccountExist = True
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim FileName As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    PasswordOK = False

    If AccountExist(Name) Then
        FileName = App.path & "\data\accounts\" & Trim$(Name) & "\data.bin"
        nFileNum = FreeFile
        Open FileName For Binary As #nFileNum
        Get #nFileNum, NAME_LENGTH, RightPassword
        Close #nFileNum
       
        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If
End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    
    ClearAccount Index
    
    Account(Index).Login = Name
    Account(Index).Password = Password
    
    Call SaveAccount(Index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    
    Call FileCopy(App.path & "\data\accounts\charlist.txt", App.path & "\data\accounts\chartemp.txt")
    
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.path & "\data\accounts\chartemp.txt" For Input As #f1
    
    f2 = FreeFile
    Open App.path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If
    Loop

    Close #f1
    Close #f2
    Call Kill(App.path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal Index As Long) As Boolean
    If Len(Trim$(Account(Index).Chars(GetPlayerChar(Index)).Name)) > 0 And Len(Trim$(Account(Index).Chars(GetPlayerChar(Index)).Name)) <= NAME_LENGTH Then
        CharExist = True
    End If
End Function

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Gender As Byte, ByVal ClassNum As Byte)
    Dim i As Long, F As Long

    With Account(Index).Chars(GetPlayerChar(Index))
        ' Basic things
        .Name = Name
        .Gender = Gender
        .Class = ClassNum
        
        ' Sprite and face
        If .Gender = GENDER_MALE Then
            .Sprite = Class(ClassNum).MaleSprite
            .Face = Class(ClassNum).MaleFace
        Else
            .Sprite = Class(ClassNum).FemaleSprite
            .Face = Class(ClassNum).FemaleFace
        End If
    
        ' Level
        .Level = 1
    
        ' Stats
        For i = 1 To Stats.Stat_count - 1
            .Stat(i) = Class(ClassNum).Stat(i)
        Next
        
        ' Skills
        For i = 1 To Skills.Skill_Count - 1
            Call SetPlayerSkillLevel(Index, 1, i)
        Next
        
        .CurrentCombatTree = 1
        
        ' Set the player's start values
        .Dir = DIR_DOWN
        .Map = START_MAP
        .X = START_X
        .Y = START_Y
        
        ' Vitals
        .Vital(Vitals.HP) = GetPlayerMaxVital(Index, Vitals.HP)
        .Vital(Vitals.MP) = GetPlayerMaxVital(Index, Vitals.MP)
        
        ' Set the checkpoint values
        .CheckPointMap = START_MAP
        .CheckPointX = START_X
        .CheckPointY = START_Y
        
        ' Set the trade status value
        .CanTrade = True
    
        ' Set the status to nothing
        .Status = vbNullString
        
        ' Check for new title
        Call CheckPlayerNewTitle(Index, False)
        
        ' Set starter equipment
        For i = 1 To MAX_INV
            If Class(ClassNum).StartItem(i) > 0 Then
                ' Item exist?
                If Len(Trim$(Item(Class(ClassNum).StartItem(i)).Name)) > 0 Then
                    .Inv(i).Num = Class(ClassNum).StartItem(i)
                    .Inv(i).Value = Class(ClassNum).StartItemValue(i)
                End If
            End If
        Next
        
        ' Set start spells
        For i = 1 To MAX_PLAYER_SPELLS
            If Class(ClassNum).StartSpell(i) > 0 Then
                ' Spell exist?
                If Len(Trim$(Spell(Class(ClassNum).StartItem(i)).Name)) > 0 Then
                    .Spell(i) = Class(ClassNum).StartSpell(i)
                End If
            End If
        Next
    End With
    
    ' Append name to file
    F = FreeFile
    
    Open App.path & "\data\accounts\charlist.txt" For Append As #F
        Print #F, Name
    Close #F
    
    Call SaveAccount(Index)
End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim s As String
    
    F = FreeFile
    
    Open App.path & "\data\accounts\charlist.txt" For Input As #F
        Do While Not EOF(F)
            Input #F, s
    
            If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
                FindChar = True
                Close #F
                Exit Function
            End If
        Loop
    Close #F
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            Call SaveAccount(i)
        End If
    Next
End Sub

Sub SaveAccount(ByVal Index As Long)
    Dim FileName As String
    Dim F As Long

    Call ChkDir(App.path & "\data\accounts\", GetPlayerLogin(Index))
    FileName = App.path & "\data\accounts\" & GetPlayerLogin(Index) & "\data.bin"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Account(Index)
    Close #F
End Sub

Sub LoadAccount(ByVal Index As Long, ByVal Name As String)
    Dim FileName As String
    Dim F As Long

    Call ClearAccount(Index)
    
    FileName = App.path & "\data\accounts\" & Name & "\data.bin"
    F = FreeFile
    
    Open FileName For Binary As #F
        Get #F, , Account(Index)
    Close #F
End Sub

Sub ClearAccount(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    TempPlayer(Index).HDSerial = vbNullString
    Set TempPlayer(Index).buffer = New clsBuffer
    
    ZeroMemory ByVal VarPtr(Account(Index)), LenB(Account(Index))
    Account(Index).Login = vbNullString
    Account(Index).Password = vbNullString
    Account(Index).CurrentChar = 1
    Account(Index).Chars(GetPlayerChar(Index)).Name = vbNullString
    Account(Index).Chars(GetPlayerChar(Index)).Status = vbNullString
    Account(Index).Chars(GetPlayerChar(Index)).Class = 1
    
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
End Sub

' ***********
' ** Classes **
' ***********
Sub SaveClasses()
    Dim i As Long

    For i = 1 To MAX_CLASSES
        Call SaveClass(i)
    Next
End Sub

Sub SaveClass(ByVal ClassNum As Long)
    Dim FileName As String
    Dim F  As Long
    
    FileName = App.path & "\data\classes\" & ClassNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Class(ClassNum)
    Close #F
End Sub

Sub LoadClasses()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckClasses

    For i = 1 To MAX_CLASSES
        FileName = App.path & "\data\classes\" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Class(i)
        Close #F
    Next
End Sub

Sub CheckClasses()
    Dim i As Long

    For i = 1 To MAX_CLASSES
        If Not FileExist("\data\classes\" & i & ".dat") Then
            Call ClearClass(i)
            Call SaveClass(i)
        End If
    Next
End Sub

Sub ClearClass(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Class(Index)), LenB(Class(Index)))
    Class(Index).Name = vbNullString
End Sub

Sub ClearClasses()
    Dim i As Long

    For i = 1 To MAX_CLASSES
        Call ClearClass(i)
    Next
End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next
End Sub

Sub SaveItem(ByVal ItemNum As Integer)
    Dim FileName As String
    Dim F  As Long
    
    FileName = App.path & "\data\items\" & ItemNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Item(ItemNum)
    Close #F
End Sub

Sub LoadItems()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckItems

    For i = 1 To MAX_ITEMS
        FileName = App.path & "\data\items\" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Item(i)
        Close #F
    Next
End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        If Not FileExist("\data\items\" & i & ".dat") Then
            Call ClearItem(i)
            Call SaveItem(i)
        End If
    Next
End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = vbNullString
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next
End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim FileName As String
    Dim F As Long
    
    FileName = App.path & "\data\shops\" & ShopNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Shop(ShopNum)
    Close #F
End Sub

Sub LoadShops()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckShops

    For i = 1 To MAX_SHOPS
        FileName = App.path & "\data\shops\" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Shop(i)
        Close #F
    Next
End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        If Not FileExist("\data\shops\" & i & ".dat") Then
            Call ClearShop(i)
            Call SaveShop(i)
        End If
    Next
End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next
End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal SpellNum As Long)
    Dim FileName As String
    Dim F As Long
    
    FileName = App.path & "\data\spells\spells" & SpellNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Spell(SpellNum)
    Close #F
End Sub

Sub SaveSpells()
    Dim i As Long
    
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next
End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        FileName = App.path & "\data\spells\" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Spell(i)
        Close #F
    Next
End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        If Not FileExist("\data\spells\" & i & ".dat") Then
            Call ClearSpell(i)
            Call SaveSpell(i)
        End If
    Next
End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).LevelReq = 1 ' Needs to be 1 for the spell editor
    Spell(Index).Sound = vbNullString
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next
End Sub

' **********
' ** Npcs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next
End Sub

Sub SaveNpc(ByVal npcnum As Long)
    Dim FileName As String
    Dim F As Long
    
    FileName = App.path & "\data\npcs\" & npcnum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , NPC(npcnum)
    Close #F
End Sub

Sub LoadNpcs()
    Dim i As Long

    Call CheckNpcs

    For i = 1 To MAX_NPCS
        Call LoadNpc(i)
    Next
End Sub

Sub LoadNpc(npcnum As Long)
    Dim F As Long
    Dim FileName As String
    
    FileName = App.path & "\data\npcs\" & npcnum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Get #F, , NPC(npcnum)
    Close #F
End Sub

Sub CheckNpcs()
    Dim i As Integer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    
    For i = 1 To MAX_NPCS
        If Not FileExist("\data\npcs\" & i & ".dat") Then
            Call ClearNpc(i)
            Call SaveNpc(i)
        End If
    Next
End Sub

Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).Title = vbNullString
    NPC(Index).AttackSay = vbNullString
    NPC(Index).Music = vbNullString
    NPC(Index).Sound = vbNullString
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next
End Sub

' ***************
' ** Resources **
' ***************
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next
End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim FileName As String
    Dim F As Long
    
    FileName = App.path & "\data\resources\" & ResourceNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
    Dim FileName As String
    Dim i As Integer
    Dim F As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        FileName = App.path & "\data\resources\" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Resource(i)
        Close #F
    Next
End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\data\resources\" & i & ".dat") Then
            Call ClearResource(i)
            Call SaveResource(i)
        End If
    Next
End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).FailMessage = vbNullString
    Resource(Index).Sound = vbNullString
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' ****************
' ** Animations **
' ****************
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next
End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim FileName As String
    Dim F As Long
    
    FileName = App.path & "\data\animations\" & AnimationNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F
End Sub

Sub LoadAnimations()
    Dim FileName As String
    Dim i As Integer
    Dim F As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        FileName = App.path & "\data\animations\" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Animation(i)
        Close #F
    Next
End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        If Not FileExist("\data\animations\" & i & ".dat") Then
            Call ClearAnimation(i)
            Call SaveAnimation(i)
        End If
    Next
End Sub

Sub ClearAnimation(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = vbNullString
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim F As Long
    Dim X As Long
    Dim Y As Long, i As Long, z As Long, w As Long
    
    FileName = App.path & "\data\maps\" & MapNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Map(MapNum).Name
        Put #F, , Map(MapNum).Music
        Put #F, , Map(MapNum).BGS
        Put #F, , Map(MapNum).Revision
        Put #F, , Map(MapNum).Moral
        Put #F, , Map(MapNum).Up
        Put #F, , Map(MapNum).Down
        Put #F, , Map(MapNum).Left
        Put #F, , Map(MapNum).Right
        Put #F, , Map(MapNum).BootMap
        Put #F, , Map(MapNum).BootX
        Put #F, , Map(MapNum).BootY
        
        Put #F, , Map(MapNum).Weather
        Put #F, , Map(MapNum).WeatherIntensity
        
        Put #F, , Map(MapNum).Fog
        Put #F, , Map(MapNum).FogSpeed
        Put #F, , Map(MapNum).FogOpacity
        
        Put #F, , Map(MapNum).Panorama
        
        Put #F, , Map(MapNum).Red
        Put #F, , Map(MapNum).Green
        Put #F, , Map(MapNum).Blue
        Put #F, , Map(MapNum).Alpha
        
        Put #F, , Map(MapNum).MaxX
        Put #F, , Map(MapNum).MaxY
        
        Put #F, , Map(MapNum).Npc_HighIndex
    
        For X = 0 To Map(MapNum).MaxX
            For Y = 0 To Map(MapNum).MaxY
                Put #F, , Map(MapNum).Tile(X, Y)
            Next
        Next
    
        For X = 1 To MAX_MAP_NPCS
            Put #F, , Map(MapNum).NPC(X)
            Put #F, , Map(MapNum).NpcSpawnType(X)
        Next
    Close #F
    
    ' This is for event saving, it is in .ini files becuase there are non-limited values (strings) that cannot easily be loaded/saved in the normal manner.
    FileName = App.path & "\data\maps\" & MapNum & "_eventdata.dat"
    PutVar FileName, "Events", "EventCount", Val(Map(MapNum).EventCount)
    
    If Map(MapNum).EventCount > 0 Then
        For i = 1 To Map(MapNum).EventCount
            With Map(MapNum).Events(i)
                PutVar FileName, "Event" & i, "Name", .Name
                PutVar FileName, "Event" & i, "Global", Val(.Global)
                PutVar FileName, "Event" & i, "x", Val(.X)
                PutVar FileName, "Event" & i, "y", Val(.Y)
                PutVar FileName, "Event" & i, "PageCount", Val(.PageCount)
            End With
            
            If Map(MapNum).Events(i).PageCount > 0 Then
                For X = 1 To Map(MapNum).Events(i).PageCount
                    With Map(MapNum).Events(i).Pages(X)
                        PutVar FileName, "Event" & i & "Page" & X, "chkVariable", Val(.chkVariable)
                        PutVar FileName, "Event" & i & "Page" & X, "VariableIndex", Val(.VariableIndex)
                        PutVar FileName, "Event" & i & "Page" & X, "VariableCondition", Val(.VariableCondition)
                        PutVar FileName, "Event" & i & "Page" & X, "VariableCompare", Val(.VariableCompare)
                        
                        PutVar FileName, "Event" & i & "Page" & X, "chkSwitch", Val(.chkSwitch)
                        PutVar FileName, "Event" & i & "Page" & X, "SwitchIndex", Val(.SwitchIndex)
                        PutVar FileName, "Event" & i & "Page" & X, "SwitchCompare", Val(.SwitchCompare)
                        
                        PutVar FileName, "Event" & i & "Page" & X, "chkHasItem", Val(.chkHasItem)
                        PutVar FileName, "Event" & i & "Page" & X, "HasItemIndex", Val(.HasItemIndex)
                        
                        PutVar FileName, "Event" & i & "Page" & X, "chkSelfSwitch", Val(.chkSelfSwitch)
                        PutVar FileName, "Event" & i & "Page" & X, "SelfSwitchIndex", Val(.SelfSwitchIndex)
                        PutVar FileName, "Event" & i & "Page" & X, "SelfSwitchCompare", Val(.SelfSwitchCompare)
                        
                        PutVar FileName, "Event" & i & "Page" & X, "GraphicType", Val(.GraphicType)
                        PutVar FileName, "Event" & i & "Page" & X, "Graphic", Val(.Graphic)
                        PutVar FileName, "Event" & i & "Page" & X, "GraphicX", Val(.GraphicX)
                        PutVar FileName, "Event" & i & "Page" & X, "GraphicY", Val(.GraphicY)
                        PutVar FileName, "Event" & i & "Page" & X, "GraphicX2", Val(.GraphicX2)
                        PutVar FileName, "Event" & i & "Page" & X, "GraphicY2", Val(.GraphicY2)
                        
                        PutVar FileName, "Event" & i & "Page" & X, "MoveType", Val(.MoveType)
                        PutVar FileName, "Event" & i & "Page" & X, "MoveSpeed", Val(.MoveSpeed)
                        PutVar FileName, "Event" & i & "Page" & X, "MoveFreq", Val(.MoveFreq)
                        
                        PutVar FileName, "Event" & i & "Page" & X, "IgnoreMoveRoute", Val(.IgnoreMoveRoute)
                        PutVar FileName, "Event" & i & "Page" & X, "RepeatMoveRoute", Val(.RepeatMoveRoute)
                        
                        PutVar FileName, "Event" & i & "Page" & X, "MoveRouteCount", Val(.MoveRouteCount)
                        
                        If .MoveRouteCount > 0 Then
                            For Y = 1 To .MoveRouteCount
                                PutVar FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Index", Val(.MoveRoute(Y).Index)
                                PutVar FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data1", Val(.MoveRoute(Y).Data1)
                                PutVar FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data2", Val(.MoveRoute(Y).Data2)
                                PutVar FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data3", Val(.MoveRoute(Y).Data3)
                                PutVar FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data4", Val(.MoveRoute(Y).Data4)
                                PutVar FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data5", Val(.MoveRoute(Y).data5)
                                PutVar FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data6", Val(.MoveRoute(Y).data6)
                            Next
                        End If
                        
                        PutVar FileName, "Event" & i & "Page" & X, "WalkAnim", Val(.WalkAnim)
                        PutVar FileName, "Event" & i & "Page" & X, "DirFix", Val(.DirFix)
                        PutVar FileName, "Event" & i & "Page" & X, "WalkThrough", Val(.WalkThrough)
                        PutVar FileName, "Event" & i & "Page" & X, "ShowName", Val(.ShowName)
                        PutVar FileName, "Event" & i & "Page" & X, "Trigger", Val(.Trigger)
                        PutVar FileName, "Event" & i & "Page" & X, "CommandListCount", Val(.CommandListCount)
                        
                        PutVar FileName, "Event" & i & "Page" & X, "Position", Val(.Position)
                    End With
                    
                    If Map(MapNum).Events(i).Pages(X).CommandListCount > 0 Then
                        For Y = 1 To Map(MapNum).Events(i).Pages(X).CommandListCount
                            PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "CommandCount", Val(Map(MapNum).Events(i).Pages(X).CommandList(Y).CommandCount)
                            PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "ParentList", Val(Map(MapNum).Events(i).Pages(X).CommandList(Y).ParentList)
                            If Map(MapNum).Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                For z = 1 To Map(MapNum).Events(i).Pages(X).CommandList(Y).CommandCount
                                    With Map(MapNum).Events(i).Pages(X).CommandList(Y).Commands(z)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Index", Val(.Index)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Text1", .Text1
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Text2", .Text2
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Text3", .Text3
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Text4", .Text4
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Text5", .Text5
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data1", Val(.Data1)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data2", Val(.Data2)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data3", Val(.Data3)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data4", Val(.Data4)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data5", Val(.data5)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data6", Val(.data6)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchCommandList", Val(.ConditionalBranch.CommandList)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchCondition", Val(.ConditionalBranch.Condition)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchData1", Val(.ConditionalBranch.Data1)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchData2", Val(.ConditionalBranch.Data2)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchData3", Val(.ConditionalBranch.Data3)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchElseCommandList", Val(.ConditionalBranch.ElseCommandList)
                                        PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRouteCount", Val(.MoveRouteCount)
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Index", Val(.MoveRoute(w).Index)
                                                PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data1", Val(.MoveRoute(w).Data1)
                                                PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data2", Val(.MoveRoute(w).Data2)
                                                PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data3", Val(.MoveRoute(w).Data3)
                                                PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data4", Val(.MoveRoute(w).Data4)
                                                PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data5", Val(.MoveRoute(w).data5)
                                                PutVar FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data6", Val(.MoveRoute(w).data6)
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
        
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next
End Sub

Sub LoadMaps()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Dim X As Long
    Dim Y As Long, z As Long, p As Long, w As Long
    Dim newtileset As Long, newtiley As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        FileName = App.path & "\data\maps\" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Map(i).Name
        Get #F, , Map(i).Music
        Get #F, , Map(i).BGS
        Get #F, , Map(i).Revision
        Get #F, , Map(i).Moral
        Get #F, , Map(i).Up
        Get #F, , Map(i).Down
        Get #F, , Map(i).Left
        Get #F, , Map(i).Right
        Get #F, , Map(i).BootMap
        Get #F, , Map(i).BootX
        Get #F, , Map(i).BootY
        
        Get #F, , Map(i).Weather
        Get #F, , Map(i).WeatherIntensity
        
        Get #F, , Map(i).Fog
        Get #F, , Map(i).FogSpeed
        Get #F, , Map(i).FogOpacity
        
        Get #F, , Map(i).Panorama
        
        Get #F, , Map(i).Red
        Get #F, , Map(i).Green
        Get #F, , Map(i).Blue
        Get #F, , Map(i).Alpha
        
        Get #F, , Map(i).MaxX
        Get #F, , Map(i).MaxY
        
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        Get #F, , Map(i).Npc_HighIndex
        
        For X = 0 To Map(i).MaxX
            For Y = 0 To Map(i).MaxY
                Get #F, , Map(i).Tile(X, Y)
            Next
        Next

        For X = 1 To MAX_MAP_NPCS
            Get #F, , Map(i).NPC(X)
            Get #F, , Map(i).NpcSpawnType(X)
            MapNpc(i).NPC(X).Num = Map(i).NPC(X)
        Next

        Close #F
        
        CacheResources i
        DoEvents
        CacheMapBlocks i
    Next
    
    For z = 1 To MAX_MAPS
        FileName = App.path & "\data\maps\" & z & "_eventdata.dat"
        Map(z).EventCount = Val(GetVar(FileName, "Events", "EventCount"))
        
        If Map(z).EventCount > 0 Then
            ReDim Map(z).Events(0 To Map(z).EventCount)
            For i = 1 To Map(z).EventCount
                With Map(z).Events(i)
                    .Name = GetVar(FileName, "Event" & i, "Name")
                    .Global = Val(GetVar(FileName, "Event" & i, "Global"))
                    .X = Val(GetVar(FileName, "Event" & i, "x"))
                    .Y = Val(GetVar(FileName, "Event" & i, "y"))
                    .PageCount = Val(GetVar(FileName, "Event" & i, "PageCount"))
                End With
                If Map(z).Events(i).PageCount > 0 Then
                    ReDim Map(z).Events(i).Pages(0 To Map(z).Events(i).PageCount)
                    For X = 1 To Map(z).Events(i).PageCount
                        With Map(z).Events(i).Pages(X)
                            .chkVariable = Val(GetVar(FileName, "Event" & i & "Page" & X, "chkVariable"))
                            .VariableIndex = Val(GetVar(FileName, "Event" & i & "Page" & X, "VariableIndex"))
                            .VariableCondition = Val(GetVar(FileName, "Event" & i & "Page" & X, "VariableCondition"))
                            .VariableCompare = Val(GetVar(FileName, "Event" & i & "Page" & X, "VariableCompare"))
                            
                            .chkSwitch = Val(GetVar(FileName, "Event" & i & "Page" & X, "chkSwitch"))
                            .SwitchIndex = Val(GetVar(FileName, "Event" & i & "Page" & X, "SwitchIndex"))
                            .SwitchCompare = Val(GetVar(FileName, "Event" & i & "Page" & X, "SwitchCompare"))
                            
                            .chkHasItem = Val(GetVar(FileName, "Event" & i & "Page" & X, "chkHasItem"))
                            .HasItemIndex = Val(GetVar(FileName, "Event" & i & "Page" & X, "HasItemIndex"))
                            
                            .chkSelfSwitch = Val(GetVar(FileName, "Event" & i & "Page" & X, "chkSelfSwitch"))
                            .SelfSwitchIndex = Val(GetVar(FileName, "Event" & i & "Page" & X, "SelfSwitchIndex"))
                            .SelfSwitchCompare = Val(GetVar(FileName, "Event" & i & "Page" & X, "SelfSwitchCompare"))
                            
                            .GraphicType = Val(GetVar(FileName, "Event" & i & "Page" & X, "GraphicType"))
                            .Graphic = Val(GetVar(FileName, "Event" & i & "Page" & X, "Graphic"))
                            .GraphicX = Val(GetVar(FileName, "Event" & i & "Page" & X, "GraphicX"))
                            .GraphicY = Val(GetVar(FileName, "Event" & i & "Page" & X, "GraphicY"))
                            .GraphicX2 = Val(GetVar(FileName, "Event" & i & "Page" & X, "GraphicX2"))
                            .GraphicY2 = Val(GetVar(FileName, "Event" & i & "Page" & X, "GraphicY2"))
                            
                            .MoveType = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveType"))
                            .MoveSpeed = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveSpeed"))
                            .MoveFreq = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveFreq"))
                            
                            .IgnoreMoveRoute = Val(GetVar(FileName, "Event" & i & "Page" & X, "IgnoreMoveRoute"))
                            .RepeatMoveRoute = Val(GetVar(FileName, "Event" & i & "Page" & X, "RepeatMoveRoute"))
                            
                            .MoveRouteCount = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveRouteCount"))
                            
                            If .MoveRouteCount > 0 Then
                                ReDim Map(z).Events(i).Pages(X).MoveRoute(0 To .MoveRouteCount)
                                For Y = 1 To .MoveRouteCount
                                    .MoveRoute(Y).Index = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Index"))
                                    .MoveRoute(Y).Data1 = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data1"))
                                    .MoveRoute(Y).Data2 = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data2"))
                                    .MoveRoute(Y).Data3 = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data3"))
                                    .MoveRoute(Y).Data4 = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data4"))
                                    .MoveRoute(Y).data5 = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data5"))
                                    .MoveRoute(Y).data6 = Val(GetVar(FileName, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data6"))
                                Next
                            End If
                            
                            .WalkAnim = Val(GetVar(FileName, "Event" & i & "Page" & X, "WalkAnim"))
                            .DirFix = Val(GetVar(FileName, "Event" & i & "Page" & X, "DirFix"))
                            .WalkThrough = Val(GetVar(FileName, "Event" & i & "Page" & X, "WalkThrough"))
                            .ShowName = Val(GetVar(FileName, "Event" & i & "Page" & X, "ShowName"))
                            .Trigger = Val(GetVar(FileName, "Event" & i & "Page" & X, "Trigger"))
                            .CommandListCount = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandListCount"))
                         
                            .Position = Val(GetVar(FileName, "Event" & i & "Page" & X, "Position"))
                        End With
                            
                        If Map(z).Events(i).Pages(X).CommandListCount > 0 Then
                            ReDim Map(z).Events(i).Pages(X).CommandList(0 To Map(z).Events(i).Pages(X).CommandListCount)
                            For Y = 1 To Map(z).Events(i).Pages(X).CommandListCount
                                Map(z).Events(i).Pages(X).CommandList(Y).CommandCount = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "CommandCount"))
                                Map(z).Events(i).Pages(X).CommandList(Y).ParentList = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "ParentList"))
                                If Map(z).Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                    ReDim Map(z).Events(i).Pages(X).CommandList(Y).Commands(Map(z).Events(i).Pages(X).CommandList(Y).CommandCount)
                                    For p = 1 To Map(z).Events(i).Pages(X).CommandList(Y).CommandCount
                                        With Map(z).Events(i).Pages(X).CommandList(Y).Commands(p)
                                            .Index = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Index"))
                                            .Text1 = GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Text1")
                                            .Text2 = GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Text2")
                                            .Text3 = GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Text3")
                                            .Text4 = GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Text4")
                                            .Text5 = GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Text5")
                                            .Data1 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data1"))
                                            .Data2 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data2"))
                                            .Data3 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data3"))
                                            .Data4 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data4"))
                                            .data5 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data5"))
                                            .data6 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data6"))
                                            .ConditionalBranch.CommandList = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchCommandList"))
                                            .ConditionalBranch.Condition = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchCondition"))
                                            .ConditionalBranch.Data1 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchData1"))
                                            .ConditionalBranch.Data2 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchData2"))
                                            .ConditionalBranch.Data3 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchData3"))
                                            .ConditionalBranch.ElseCommandList = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchElseCommandList"))
                                            .MoveRouteCount = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRouteCount"))
                                            If .MoveRouteCount > 0 Then
                                                ReDim .MoveRoute(1 To .MoveRouteCount)
                                                For w = 1 To .MoveRouteCount
                                                    .MoveRoute(w).Index = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Index"))
                                                    .MoveRoute(w).Data1 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data1"))
                                                    .MoveRoute(w).Data2 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data2"))
                                                    .MoveRoute(w).Data3 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data3"))
                                                    .MoveRoute(w).Data4 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data4"))
                                                    .MoveRoute(w).data5 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data5"))
                                                    .MoveRoute(w).data6 = Val(GetVar(FileName, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data6"))
                                                Next
                                            End If
                                        End With
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If
        DoEvents
    Next
End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        If Not FileExist("\data\maps\" & i & ".dat") Then
            Call ClearMap(i)
            Call SaveMap(i)
        End If
    Next
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Integer)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, Index)), LenB(MapItem(MapNum, Index)))
    MapItem(MapNum, Index).PlayerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
        Next
    Next
End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Integer)
    ReDim MapNpc(MapNum).NPC(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).NPC(Index)), LenB(MapNpc(MapNum).NPC(Index)))
End Sub

Sub ClearMapNpcs()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, Y)
        Next
    Next
End Sub

Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).Name = vbNullString
    Map(MapNum).Music = vbNullString
    Map(MapNum).BGS = vbNullString
    Map(MapNum).Moral = 1
    Map(MapNum).MaxX = MIN_MAPX
    Map(MapNum).MaxY = MIN_MAPY
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next
End Sub

' ************
' ** Guilds **
' ************
Sub SaveGuilds()
    Dim i As Long

    For i = 1 To MAX_GUILDS
        Call SaveGuild(i)
    Next
End Sub

Sub SaveGuild(ByVal GuildNum As Long)
    Dim FileName As String
    Dim F  As Long
    
    FileName = App.path & "\data\guilds\" & GuildNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Guild(GuildNum)
    Close #F
End Sub

Sub LoadGuilds()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckGuild

    For i = 1 To MAX_GUILDS
        FileName = App.path & "\data\guilds\" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Guild(i)
        Close #F
    Next
End Sub

Sub CheckGuild()
    Dim i As Long

    For i = 1 To MAX_GUILDS
        If Not FileExist("\data\guilds\" & i & ".dat") Then
            Call ClearGuild(i)
            Call SaveGuild(i)
        End If
    Next
End Sub

Sub ClearGuilds()
    Dim i As Long

    For i = 1 To MAX_GUILDS
        Call ClearGuild(i)
    Next
End Sub

Sub ClearGuild(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Guild(Index)), LenB(Guild(Index)))
    Guild(Index).Name = vbNullString
    Guild(Index).MOTD = vbNullString
End Sub

' ************
' ** Bans **
' ************
Sub SaveBan(ByVal BanNum As Long)
    Dim F As Long
    Dim FileName As String
    
    F = FreeFile
    FileName = App.path & "\data\bans\" & BanNum & ".dat"
    
    Open FileName For Binary As #F
        Put #F, , Ban(BanNum)
    Close #F
End Sub

Sub CheckBans()
    Dim i As Long

    For i = 1 To MAX_BANS
        If Not FileExist("\data\bans\" & i & ".dat") Then
            Call ClearBan(i)
            Call SaveBan(i)
        End If
    Next
End Sub

Sub ClearBans()
    Dim i As Long
    
    For i = 1 To MAX_BANS
        Call ClearBan(i)
    Next
End Sub

Sub ClearBan(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Ban(Index)), LenB(Ban(Index)))
    Ban(Index).PlayerLogin = vbNullString
    Ban(Index).PlayerName = vbNullString
    Ban(Index).Reason = vbNullString
    Ban(Index).IP = vbNullString
    Ban(Index).HDSerial = vbNullString
    Ban(Index).Time = vbNullString
    Ban(Index).By = vbNullString
    Ban(Index).Date = vbNullString
End Sub

' ************
' ** Titles **
' ************
Sub SaveTitle(ByVal TitleNum As Long)
    Dim F As Long
    Dim FileName As String

    F = FreeFile
    FileName = App.path & "\data\titles\" & TitleNum & ".dat"
    
    Open FileName For Binary As #F
        Put #F, , Title(TitleNum)
    Close #F
End Sub

Sub LoadTitles()
    Dim i As Long

    CheckTitles
    
    For i = 1 To MAX_TITLES
        Call LoadTitle(i)
    Next
End Sub

Sub LoadTitle(Index As Long)
    Dim F As Long
    Dim FileName  As String

    F = FreeFile
    FileName = App.path & "\data\titles\" & Index & ".dat"
    
    Open FileName For Binary As #F
        Get #F, , Title(Index)
    Close #F
End Sub

Sub CheckTitles()
    Dim i As Long

    For i = 1 To MAX_TITLES
        If Not FileExist("\data\titles\" & i & ".dat") Then
            Call ClearTitle(i)
            Call SaveTitle(i)
        End If
    Next
End Sub

Sub ClearTitles()
    Dim i As Long
    
    For i = 1 To MAX_TITLES
        Call ClearTitle(i)
    Next
End Sub

Sub ClearTitle(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Title(Index)), LenB(Title(Index)))
    Title(Index).Name = vbNullString
End Sub

' ************
' ** Morals **
' ************
Sub SaveMorals()
    Dim i As Long

    For i = 1 To MAX_MORALS
        Call SaveMoral(i)
    Next
End Sub

Sub SaveMoral(ByVal MoralNum As Long)
    Dim FileName As String
    Dim F  As Long
    
    FileName = App.path & "\data\morals\" & MoralNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Moral(MoralNum)
    Close #F
End Sub

Sub LoadMorals()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckMorals

    For i = 1 To MAX_MORALS
        FileName = App.path & "\data\morals\" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Moral(i)
        Close #F
    Next
End Sub

Sub CheckMorals()
    Dim i As Long

    For i = 1 To MAX_MORALS
        If Not FileExist("\data\morals\" & i & ".dat") Then
            Call ClearMoral(i)
            Call SaveMoral(i)
        End If
    Next
End Sub

Sub ClearMoral(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Moral(Index)), LenB(Moral(Index)))
    Moral(Index).Name = vbNullString
End Sub

Sub ClearMorals()
    Dim i As Long

    For i = 1 To MAX_MORALS
        Call ClearMoral(i)
    Next
End Sub

' **************
' ** Emoticons **
' **************
Sub SaveEmoticons()
    Dim i As Long

    For i = 1 To MAX_EMOTICONS
        Call SaveEmoticon(i)
    Next
End Sub

Sub SaveEmoticon(ByVal EmoticonNum As Long)
    Dim FileName As String
    Dim F  As Long
    
    FileName = App.path & "\data\emoticons\" & EmoticonNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Emoticon(EmoticonNum)
    Close #F
End Sub

Sub LoadEmoticons()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckEmoticons

    For i = 1 To MAX_EMOTICONS
        FileName = App.path & "\data\emoticons\" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Emoticon(i)
        Close #F
    Next
End Sub

Sub CheckEmoticons()
    Dim i As Long

    For i = 1 To MAX_EMOTICONS
        If Not FileExist("\data\emoticons\" & i & ".dat") Then
            Call ClearEmoticon(i)
            Call SaveEmoticon(i)
        End If
    Next
End Sub

Sub ClearEmoticons()
    Dim i As Long

    For i = 1 To MAX_EMOTICONS
        Call ClearEmoticon(i)
    Next
End Sub

Sub ClearEmoticon(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Emoticon(Index)), LenB(Emoticon(Index)))
    Emoticon(Index).Command = "/"
End Sub

' ***********
' ** Party **
' ***********
Sub ClearParty(ByVal PartyNum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(PartyNum)), LenB(Party(PartyNum)))
End Sub

Sub SaveTempChar(ByVal Index As Long, ByVal Login As String)
    Dim FileName As String
    Dim F As Long

    FileName = App.path & "\data\accounts\" & Trim$(Login) & "\data.bin"
    
    F = FreeFile
    
    Open FileName For Binary As #F
    Put #F, , TempChar(Index)
    Close #F
End Sub

Sub LoadTempChar(ByVal Index As Long, ByVal Login As String)
    Dim FileName As String
    Dim F As Long
    
    Call ClearTempChar(Index)
    FileName = App.path & "\data\Accounts\" & Trim$(Login) & "\data.bin"
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , TempChar(Index)
    Close #F
End Sub

Sub ClearTempChar(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(TempChar(Index)), LenB(TempChar(Index)))
End Sub

Sub SaveSwitches()
    Dim i As Long, FileName As String
    
    FileName = App.path & "\data\switches.ini"
    
    For i = 1 To MAX_SWITCHES
        Call PutVar(FileName, "Switches", "Switch" & CStr(i) & "Name", Switches(i))
    Next
End Sub

Sub SaveVariables()
    Dim i As Long, FileName As String
    
    FileName = App.path & "\data\variables.ini"
    
    For i = 1 To MAX_VARIABLES
        Call PutVar(FileName, "Variables", "Variable" & CStr(i) & "Name", Variables(i))
    Next
End Sub

Sub LoadSwitches()
    Dim i As Long, FileName As String
    
    FileName = App.path & "\data\switches.ini"
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = GetVar(FileName, "Switches", "Switch" & CStr(i) & "Name")
    Next
End Sub

Sub LoadVariables()
    Dim i As Long, FileName As String
    
    FileName = App.path & "\data\variables.ini"
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = GetVar(FileName, "Variables", "Variable" & CStr(i) & "Name")
    Next
End Sub
