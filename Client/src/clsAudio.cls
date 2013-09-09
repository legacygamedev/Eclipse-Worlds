VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Music and Sound Constants
Private Music As Long
Private SoundIndex As Byte
Const BASSFALSE = 0
Private Sub Class_Initialize()
    ' Change and set the current path, to prevent from VB not finding Bass.dll
    ChDrive App.Path
    ChDir App.Path

    ' Don't run while in IDE
    'If App.LogMode = 0 Then Exit Sub

    ' Check the correct BASS was loaded
    If Not HiWord(BASS_GetVersion) = BASSVERSION Then
        Call AlertMsg("An incorrect version of bass.dll was loaded!")
        DestroyGame
    End If
    
    ' Initialize output - default device, 44100hz, stereo, 16 bits
    If BASS_Init(-1, 44100, 0, frmMain.hWnd, 0) = BASSFALSE Then
        'Call Error_("Can't initialize digital sound system!")
        'DestroyGame
    End If
End Sub

Public Sub PlayMusic(FileName As String)
    If Options.Music = 0 Or FileName = vbNullString Or CurrentMusic = FileName Then Exit Sub
    
    Call StopMusic
    
    If FileExist(App.Path & MUSIC_PATH & FileName, True) Then
        If Left$(FileName, 7) = "http://" Then
            If Right$(FileName, 4) = "Mid" Then
                Music = BASS_MidI_StreamCreateURL(FileName, 0, BASS_SAMPLE_LOOP, 0, 0, 0)
            Else
                Music = BASS_StreamCreateURL(FileName, 0, BASS_SAMPLE_LOOP, 0, 0)
            End If
        Else
            Music = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & MUSIC_PATH & FileName), 0, 0, BASS_SAMPLE_LOOP)
        End If
        
        Call BASS_ChannelSetAttribute(Music, BASS_ATTRIB_VOL, Options.MusicVolume)
        Call BASS_ChannelPlay(Music, BASSFALSE)
        CurrentMusic = FileName
    Else
        If FormVisible("frmMain") Then
            Call AddText(FileName & " does not exist!", BrightRed)
        End If
    End If
End Sub

Public Sub PlaySound(FileName As String, Optional ByVal X As Long, Optional ByVal Y As Long)
    Dim i As Long
    
    If Options.Sound = 0 Or FileName = vbNullString Then Exit Sub
    
    If FileExist(App.Path & SOUND_PATH & FileName, True) Then
        ' Find the index
        For i = 1 To MAX_SOUNDS
            If Sounds(i) = 0 Then
                 SoundIndex = i
                 Exit For
            ElseIf BASS_ChannelIsActive(Sounds(i)) Then
                SoundIndex = i
                Exit For
            End If
            
            ' Set it to the beginning and stop the sound if all sounds are being used
            If i = MAX_SOUNDS Then
                StopSound 1
                SoundIndex = 1
            End If
        Next
        
        ' Create the sound
        Sounds(SoundIndex) = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & SOUND_PATH & FileName), 0, 0, 0)
        Call BASS_ChannelSetAttribute(Sounds(SoundIndex), BASS_ATTRIB_VOL, Options.SoundVolume)
        
        ' Play the sound
        Call BASS_ChannelPlay(Sounds(SoundIndex), BASSFALSE)
    Else
        If FormVisible("frmMain") Then
            Call AddText(FileName & " does not exist!", BrightRed)
        End If
    End If
End Sub

Public Sub StopMusic()
    If Not CurrentMusic = vbNullString Then
        Call BASS_ChannelStop(Music)
        Call ZeroMemory(ByVal VarPtr(Music), LenB(Music))
        CurrentMusic = vbNullString
    End If
End Sub

Public Sub StopSound(ByVal Index As Byte)
    If Sounds(Index) = 0 Then
        Call BASS_ChannelStop(Sounds(Index))
        Call ZeroMemory(ByVal VarPtr(Sounds(Index)), LenB(Sounds(Index)))
    End If
End Sub

Public Sub StopSounds()
    Dim i As Long
    
    For i = 1 To MAX_SOUNDS
        Call StopSound(i)
    Next
End Sub

Public Sub StopMapSound(ByVal Index As Byte)
    If Not MapSounds(Index).handle = 0 Then
        Call BASS_ChannelStop(MapSounds(Index).handle)
        Call ZeroMemory(ByVal VarPtr(MapSounds(Index)), LenB(MapSounds(Index).handle))
    End If
End Sub

Sub StopMapSounds()
    Dim i As Long
    
    If MapSoundCount > 0 Then
        For i = 1 To MapSoundCount
            Call StopMapSound(i)
        Next
        
        MapSoundCount = 0
        ReDim MapSounds(0)
    End If
End Sub

Function CalculateSoundVolume(X As Long, Y As Long) As Double
    Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, Distance As Double, volume As Long
    
    If InGame = False Then CalculateSoundVolume = Options.SoundVolume: Exit Function
    
    If X > -1 Or Y > -1 Then
        X1 = (Player(MyIndex).X * 32) + TempPlayer(MyIndex).xOffset
        Y1 = (Player(MyIndex).Y * 32) + TempPlayer(MyIndex).yOffset
        X2 = (X * 32) + 16
        Y2 = (Y * 32) + 16
        
        If ((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2) < 0 Then
            Distance = Sqr(((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2) * -1)
        Else
            Distance = Sqr(((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2))
        End If
        
        If Distance >= 512 Then
            CalculateSoundVolume = 0
        Else
            CalculateSoundVolume = 1 - Distance / 512
        End If
    Else
        CalculateSoundVolume = Options.SoundVolume
    End If
End Function

Sub UpdateMapSounds()
    Dim i As Long, X As Long
    
    If MapSoundCount > 0 Then
        For i = 1 To MapSoundCount
            BASS_ChannelSetAttribute MapSounds(i).handle, BASS_ATTRIB_VOL, CalculateSoundVolume(MapSounds(i).X, MapSounds(MapSoundCount).Y)
        Next
    End If
End Sub

Sub CacheNewMapSounds()
    Dim i As Long, X As Long, Y As Long
    
    StopMapSounds
    
    If Trim$(Map.BGS) <> vbNullString Then
        MapSoundCount = MapSoundCount + 1
        ReDim Preserve MapSounds(MapSoundCount)
        MapSounds(MapSoundCount).handle = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & SOUND_PATH & Trim$(Map.BGS)), 0, 0, BASS_SAMPLE_LOOP)
        MapSounds(MapSoundCount).X = -1
        MapSounds(MapSoundCount).Y = -1
        Call BASS_ChannelSetAttribute(MapSounds(MapSoundCount).handle, BASS_ATTRIB_VOL, CalculateSoundVolume(MapSounds(MapSoundCount).X, MapSounds(MapSoundCount).Y))
        
        ' Play the sound
        Call BASS_ChannelPlay(MapSounds(MapSoundCount).handle, BASSFALSE)
    End If
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If Map.Tile(X, Y).Type = TILE_TYPE_SOUND Then
                MapSoundCount = MapSoundCount + 1
                ReDim Preserve MapSounds(MapSoundCount)
                MapSounds(MapSoundCount).handle = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & SOUND_PATH & Trim$(Map.Tile(X, Y).Data4)), 0, 0, BASS_SAMPLE_LOOP)
                MapSounds(MapSoundCount).X = X
                MapSounds(MapSoundCount).Y = Y
                Call BASS_ChannelSetAttribute(MapSounds(MapSoundCount).handle, BASS_ATTRIB_VOL, CalculateSoundVolume(MapSounds(MapSoundCount).X, MapSounds(MapSoundCount).Y))
                
                ' Play the sound
                Call BASS_ChannelPlay(MapSounds(MapSoundCount).handle, BASSFALSE)
            End If
        Next
    Next
End Sub

Public Sub Error_(ByVal es As String)
    If Options.Debug = 0 Then
        Call AlertMsg(es & vbCrLf & "(error code: " & BASS_ErrorGetCode() & ")")
    End If
End Sub
