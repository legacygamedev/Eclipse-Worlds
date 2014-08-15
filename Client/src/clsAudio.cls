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
Public IsMuted As Boolean

' Music engine for MIDI
Public Performance As DirectMusicPerformance
Public Segment As DirectMusicSegment
Public Loader As DirectMusicLoader

Private Sub Class_Initialize()
    ' Don't run while in IDE
    'If App.LogMode = 0 Then Exit Sub

    ' Check the correct BASS was loaded
    If Not HiWord(BASS_GetVersion) = BASSVERSION Then
        Call AlertMsg("An incorrect version of bass.dll was loaded!")
        DestroyGame
    End If
    
    ' Initialize output - default device, 44100hz, stereo, 16 bits
    Call BASS_Init(-1, 44100, 0, frmMain.hWnd, 0)
    
    Set Loader = DX7.DirectMusicLoaderCreate
    Set Performance = DX7.DirectMusicPerformanceCreate
   
    Performance.Init Nothing, frmMain.hWnd
    Performance.SetPort -1, 80
   
    ' Adjust volume 0-100
    Performance.SetMasterVolume (Options.MusicVolume * 100) * 42 - 3000
    Performance.SetMasterAutoDownload True
End Sub

Public Sub PlayMusic(FileName As String, Optional ByVal Ignore As Boolean = False)
    Dim SplitMusic() As String
    
    If Ignore = False And (Options.Music = 0 Or Trim$(FileName) = vbNullString Or Trim$(FileName) = "None" Or CurrentMusic = Trim$(FileName)) Then Exit Sub
    
    Call StopMusic
    
    ' File doesn't exist
    If Not FileExist(App.Path & MUSIC_PATH & Trim$(FileName), True) Then
        If FormVisible("frmMain") Then
            Call AddText(Trim$(FileName) & " does not exist!", BrightRed)
        End If
        Exit Sub
    End If
    
    If Right$(Trim$(FileName), 4) = ".mid" Then
        SplitMusic = Split(Trim$(FileName), ".", , vbTextCompare)
        
        If Performance Is Nothing Then Exit Sub
        If LenB(Trim$(FileName)) < 1 Then Exit Sub
        If UBound(SplitMusic) <> 1 Then Exit Sub
        If SplitMusic(1) <> "mid" Then Exit Sub
        If IsMuted Then Exit Sub
        
        Set Segment = Nothing
        Set Segment = Loader.LoadSegment(App.Path & MUSIC_PATH & Trim$(FileName))
        
        ' Repeat midi file
        Segment.SetLoopPoints 0, 0
        Segment.SetRepeats 100
        Segment.SetStandardMidiFile
        
        Performance.PlaySegment Segment, 0, 0
        
        CurrentMusic = Trim$(FileName)
    Else
        If Left$(Trim$(FileName), 7) = "http://" Then
            Music = BASS_StreamCreateURL(Trim$(FileName), 0, BASS_SAMPLE_LOOP, 0, 0)
        Else
            Music = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & MUSIC_PATH & Trim$(FileName)), 0, 0, BASS_SAMPLE_LOOP)
        End If

        If Not IsMuted Then
                Call BASS_ChannelSetAttribute(Music, BASS_ATTRIB_VOL, Options.MusicVolume)
                Call BASS_ChannelPlay(Music, BASSFALSE)
        End If
        CurrentMusic = Trim$(FileName)
        CurrentMusicIndex = Music
    End If
End Sub

Public Sub PlaySound(FileName As String, _
                     Optional ByVal X As Long, _
                     Optional ByVal Y As Long, _
                     Optional ByVal Ignore As Boolean = False)
   Dim i As Long

   '
   If Ignore = False And (Options.Sound = 0 Or Trim$(FileName) = vbNullString Or Trim$(FileName) = "None") Then Exit Sub
    
   If FileExist(App.Path & SOUND_PATH & Trim$(FileName), True) Then

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
      Sounds(SoundIndex) = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & SOUND_PATH & Trim$(FileName)), 0, 0, 0)
      
      If Not IsMuted Then
         If X > 0 Or Y > 0 Then
            Call BASS_ChannelSetAttribute(Sounds(SoundIndex), BASS_ATTRIB_VOL, CalculateSoundVolume(X, Y))
         Else
            Call BASS_ChannelSetAttribute(Sounds(SoundIndex), BASS_ATTRIB_VOL, Options.SoundVolume)
         End If
           
         ' Play the sound
         Call BASS_ChannelPlay(Sounds(SoundIndex), BASSFALSE)
      End If
   End If

End Sub

Public Sub StopMusic()
    If Not CurrentMusic = vbNullString Then
        If Not (Performance Is Nothing) Then Performance.Stop Segment, Nothing, 0, 0
        Call BASS_ChannelStop(Music)
        Call ZeroMemory(ByVal VarPtr(Music), LenB(Music))
        CurrentMusic = vbNullString
        CurrentMusicIndex = 0
    End If
End Sub

Public Sub StopSound(ByVal Index As Byte, Optional ByVal Clear As Boolean = True)
    Call BASS_ChannelStop(Sounds(Index))
    If Clear Then Call ZeroMemory(ByVal VarPtr(Sounds(Index)), LenB(Sounds(Index)))
End Sub

Public Sub StopMapSound(ByVal Index As Byte, Optional ByVal Clear As Boolean = True)
    Call BASS_ChannelStop(MapSounds(Index).handle)
    If Clear Then Call ZeroMemory(ByVal VarPtr(MapSounds(Index)), LenB(MapSounds(Index).handle))
End Sub

Public Sub StopSounds()
    Dim i As Long
    
    For i = 1 To MAX_SOUNDS
        Call StopSound(i)
    Next
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

    If InGame = False Then CalculateSoundVolume = Options.Sound: Exit Function
    If InGame And X = GetPlayerX(MyIndex) And Y = GetPlayerY(MyIndex) Then CalculateSoundVolume = Options.Sound: Exit Function

    If X > -1 Or Y > -1 Then
        If X = -1 Then X = 0
        If Y = -1 Then Y = 0
        X1 = (Player(MyIndex).X * 32) + TempPlayer(MyIndex).xOffset
        Y1 = (Player(MyIndex).Y * 32) + TempPlayer(MyIndex).yOffset
        X2 = (X * 32)
        Y2 = (Y * 32)

        If ((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2) < 0 Then
            Distance = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2) * -1)
        Else
            Distance = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
        End If
        
        ' If the range is greater than 32 (1024), do not send a sound
        If (Distance / 32) > 32 Then
            CalculateSoundVolume = 0
        Else
            CalculateSoundVolume = Distance / 1024

            If CalculateSoundVolume > Options.Sound Then
                CalculateSoundVolume = ((1 - CalculateSoundVolume) * (Options.Sound))
            ElseIf Options.Sound > CalculateSoundVolume Then
                CalculateSoundVolume = (Options.Sound) * (1 - CalculateSoundVolume)
            Else
                CalculateSoundVolume = Options.Sound
            End If
        End If

    Else
        CalculateSoundVolume = Options.Sound
    End If
End Function

Sub UpdateSounds()
    Dim i As Long, X As Long

    If MapSoundCount > 0 Then
        For i = 1 To MapSoundCount
            If BASS_ChannelIsActive(MapSounds(i).handle) Then
                BASS_ChannelSetAttribute MapSounds(i).handle, BASS_ATTRIB_VOL, CalculateSoundVolume(MapSounds(i).X, MapSounds(MapSoundCount).Y)
            Else
                Call StopMapSound(i)
            End If
        Next
    End If
End Sub

Sub UpdateMapSounds()
    Dim i As Long, X As Long
    
    If MapSoundCount > 0 Then
        For i = 1 To MapSoundCount
            If BASS_ChannelIsActive(MapSounds(i).handle) Then
                BASS_ChannelSetAttribute MapSounds(i).handle, BASS_ATTRIB_VOL, CalculateSoundVolume(MapSounds(i).X, MapSounds(MapSoundCount).Y)
            Else
                Call StopMapSound(i)
            End If
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

Public Sub MuteVolume()

    Dim i As Long

    ' Adjust volume
    If CurrentMusicIndex <> 0 Then
        Call BASS_ChannelSetAttribute(CurrentMusicIndex, BASS_ATTRIB_VOL, 0)
    End If

    Call BASS_ChannelPlay(CurrentMusicIndex, BASSFALSE)

    StopSounds
    StopMapSounds

    IsMuted = True
End Sub

Public Sub UpdateVolume()

    Dim i As Long

    If CurrentMusicIndex <> 0 Then
        Call BASS_ChannelSetAttribute(CurrentMusicIndex, BASS_ATTRIB_VOL, Options.Music)
    End If

    Call BASS_ChannelPlay(CurrentMusicIndex, BASSFALSE)
    
    ' Adjust sound volume
    For i = 1 To MAX_SOUNDS

        If BASS_ChannelIsActive(Sounds(i)) Then
            Call BASS_ChannelSetAttribute(Sounds(i), BASS_ATTRIB_VOL, Options.Sound)
            Call BASS_ChannelPlay(Sounds(i), BASSFALSE)
        Else
            Call StopSound(i)
        End If

    Next

    If MapSoundCount > 0 Then

        For i = 1 To MapSoundCount
            
            If BASS_ChannelIsActive(MapSounds(i).handle) Then
                Call BASS_ChannelSetAttribute(MapSounds(i).handle, BASS_ATTRIB_VOL, Options.Sound)
            Else
                Call StopMapSound(i)
            End If

        Next

    End If

    IsMuted = False
End Sub

Public Sub Error_(ByVal es As String)
    If Options.Debug = 0 Then
        Call AlertMsg(es & vbCrLf & "(error code: " & BASS_ErrorGetCode() & ")")
    End If
End Sub
