Attribute VB_Name = "modBass"
' BASS 2.4 Visual Basic module
' Copyright (c) 1999-2013 Un4seen Developments Ltd.
'
' See the BASS.CHM file for more detailed documentation

' NOTE: VB does not support 64-bit integers, so VB users only have access
'       to the low 32-bits of 64-bit return values. 64-bit parameters can
'       be specified though, using the "64" version of the function.

' NOTE: Use "StrPtr(filename)" to pass a filename to the BASS_MusicLoad,
'       BASS_SampleLoad and BASS_StreamCreateFile functions.

' NOTE: Use the VBStrFromAnsiPtr function to convert "char *" to VB "String".

Global Const BASSVERSION = &H204    'API version

' Device info structure
Type BASS_DEVICEINFO
    name As Long          ' description
    driver As Long        ' driver
    flags As Long
End Type


Type BASS_INFO
    flags As Long         ' device capabilities (DSCAPS_xxx flags)
    hwsize As Long        ' size of total device hardware memory
    hwfree As Long        ' size of free device hardware memory
    freesam As Long       ' number of free sample slots in the hardware
    free3d As Long        ' number of free 3D sample slots in the hardware
    minrate As Long       ' min sample rate supported by the hardware
    maxrate As Long       ' max sample rate supported by the hardware
    eax As Long           ' device supports EAX? (always BASSFALSE if BASS_DEVICE_3D was not used)
    minbuf As Long        ' recommended minimum buffer length in ms (requires BASS_DEVICE_LATENCY)
    dsver As Long         ' DirectSound version
    latency As Long       ' delay (in ms) before start of playback (requires BASS_DEVICE_LATENCY)
    initflags As Long     ' BASS_Init "flags" parameter
    speakers As Long      ' number of speakers available
    freq As Long          ' current output rate
End Type

' Recording device info structure
Type BASS_RECORDINFO
    flags As Long         ' device capabilities (DSCCAPS_xxx flags)
    formats As Long       ' supported standard formats (WAVE_FORMAT_xxx flags)
    inputs As Long        ' number of inputs
    singlein As Long      ' BASSTRUE = only 1 input can be set at a time
    freq As Long          ' current input rate
End Type

' Sample info structure
Type BASS_SAMPLE
    freq As Long          ' default playback rate
    volume As Single      ' default volume (0-100)
    pan As Single         ' default pan (-100=left, 0=Middle, 100=right)
    flags As Long         ' BASS_SAMPLE_xxx flags
    Length As Long        ' length (in samples, not bytes)
    max As Long           ' maximum simultaneous playbacks
    origres As Long       ' original resolution
    chans As Long         ' number of channels
    mingap As Long        ' minimum gap (ms) between creating channels
    mode3d As Long        ' BASS_3DMODE_xxx mode
    mindist As Single     ' minimum distance
    MAXDIST As Single     ' maximum distance
    iangle As Long        ' angle of inside projection cone
    oangle As Long        ' angle of outside projection cone
    outvol As Single      ' delta-volume outside the projection cone
    vam As Long           ' voice allocation/management flags (BASS_VAM_xxx)
    priority As Long      ' priority (0=lowest, &Hffffffff=highest)
End Type
Global Const BASS_SAMPLE_LOOP = 4           ' looped

Global Const BASS_STREAM_PRESCAN = &H20000   ' Enable pin-point seeking/length (MP3/MP2/MP1)

Global Const BASS_UNICODE = &H80000000


' Channel info structure
Type BASS_CHANNELINFO
    freq As Long          ' default playback rate
    chans As Long         ' channels
    flags As Long         ' BASS_SAMPLE/STREAM/MUSIC/SPEAKER flags
    ctype As Long         ' type of channel
    origres As Long       ' original resolution
    plugin As Long        ' plugin
    sample As Long        ' sample
    FileName As Long      ' Filename
End Type

Type BASS_PLUGINFORM
    ctype As Long         ' channel type
    name As Long          ' Format description
    exts As Long          ' File extension filter (*.ext1;*.ext2;etc...)
End Type

Type BASS_PLUGININFO
    Version As Long       ' version (same form as BASS_GetVersion)
    formatc As Long       ' number of formats
    formats As Long       ' the array of formats
End Type

' 3D vector (for 3D positions/velocities/orientations)
Type BASS_3DVECTOR
    X As Single           ' +=right, -=left
    Y As Single           ' +=up, -=down
    Z As Single           ' +=front, -=behind
End Type

' EAX environments, use with BASS_SetEAXParameters
Global Const EAX_ENVIRONMENT_GENERIC = 0
Global Const EAX_ENVIRONMENT_PADDEDCELL = 1
Global Const EAX_ENVIRONMENT_ROOM = 2
Global Const EAX_ENVIRONMENT_BATHROOM = 3
Global Const EAX_ENVIRONMENT_LIVINGROOM = 4
Global Const EAX_ENVIRONMENT_STONEROOM = 5
Global Const EAX_ENVIRONMENT_AUDITORIUM = 6
Global Const EAX_ENVIRONMENT_CONCERTHALL = 7
Global Const EAX_ENVIRONMENT_CAVE = 8
Global Const EAX_ENVIRONMENT_ARENA = 9
Global Const EAX_ENVIRONMENT_HANGAR = 10
Global Const EAX_ENVIRONMENT_CARPETEDHALLWAY = 11
Global Const EAX_ENVIRONMENT_HALLWAY = 12
Global Const EAX_ENVIRONMENT_STONECORRIDOR = 13
Global Const EAX_ENVIRONMENT_ALLEY = 14
Global Const EAX_ENVIRONMENT_FOREST = 15
Global Const EAX_ENVIRONMENT_CITY = 16
Global Const EAX_ENVIRONMENT_MOUNTAINS = 17
Global Const EAX_ENVIRONMENT_QUARRY = 18
Global Const EAX_ENVIRONMENT_PLAIN = 19
Global Const EAX_ENVIRONMENT_PARKINGLOT = 20
Global Const EAX_ENVIRONMENT_SEWERPIPE = 21
Global Const EAX_ENVIRONMENT_UNDERWATER = 22
Global Const EAX_ENVIRONMENT_DRUGGED = 23
Global Const EAX_ENVIRONMENT_DIZZY = 24
Global Const EAX_ENVIRONMENT_PSYCHOTIC = 25

Type BASS_FILEPROCS
    close As Long
    Length As Long
    read As Long
    seek As Long
End Type

' Channel attributes
Global Const BASS_ATTRIB_VOL = 2

' ID3v1 tag structure
Type TAG_ID3
    id As String * 3
    title As String * 30
    artist As String * 30
    album As String * 30
    year As String * 4
    comment As String * 30
    genre As Byte
End Type

' Binary APEv2 tag structure
Type TAG_APE_BINARY
    key As Long
    data As Long
    Length As Long
End Type

' BWF "bext" tag structure
Type TAG_BEXT
    Description As String * 256         ' description
    Originator As String * 32           ' name of the originator
    OriginatorReference As String * 32  ' reference of the originator
    OriginationDate As String * 10      ' date of creation (yyyy-mm-dd)
    OriginationTime As String * 8       ' time of creation (hh-mm-ss)
    TimeReferenceLo As Long             ' low 32 bits of first sample count since Midnight (little-endian)
    TimeReferenceHi As Long             ' high 32 bits of first sample count since Midnight (little-endian)
    Version As Integer                  ' BWF version (little-endian)
    UMid(0 To 63) As Byte               ' SMPTE UMid
    Reserved(0 To 189) As Byte
    CodingHistory() As String           ' history
End Type

Type BASS_DX8_CHORUS
    fWetDryMix As Single
    fDepth As Single
    fFeedback As Single
    fFrequency As Single
    lWaveform As Long   ' 0=triangle, 1=sine
    fDelay As Single
    lPhase As Long              ' BASS_DX8_PHASE_xxx
End Type

Type BASS_DX8_COMPRESSOR
    fGain As Single
    fAttack As Single
    fRelease As Single
    fThreshold As Single
    fRatio As Single
    fPredelay As Single
End Type

Type BASS_DX8_DISTORTION
    fGain As Single
    fEdge As Single
    fPostEQCenterFrequency As Single
    fPostEQBandwidth As Single
    fPreLowpassCutoff As Single
End Type

Type BASS_DX8_ECHO
    fWetDryMix As Single
    fFeedback As Single
    fLeftDelay As Single
    fRightDelay As Single
    lPanDelay As Long
End Type

Type BASS_DX8_FLANGER
    fWetDryMix As Single
    fDepth As Single
    fFeedback As Single
    fFrequency As Single
    lWaveform As Long   ' 0=triangle, 1=sine
    fDelay As Single
    lPhase As Long              ' BASS_DX8_PHASE_xxx
End Type

Type BASS_DX8_GARGLE
    dwRateHz As Long               ' Rate of modulation in hz
    dwWaveShape As Long            ' 0=triangle, 1=square
End Type

Type BASS_DX8_I3DL2REVERB
    lRoom As Long                    ' [-10000, 0]      default: -1000 mB
    lRoomHF As Long                  ' [-10000, 0]      default: 0 mB
    flRoomRolloffFactor As Single    ' [0.0, 10.0]      default: 0.0
    flDecayTime As Single            ' [0.1, 20.0]      default: 1.49s
    flDecayHFRatio As Single         ' [0.1, 2.0]       default: 0.83
    lReflections As Long             ' [-10000, 1000]   default: -2602 mB
    flReflectionsDelay As Single     ' [0.0, 0.3]       default: 0.007 s
    lReverb As Long                  ' [-10000, 2000]   default: 200 mB
    flReverbDelay As Single          ' [0.0, 0.1]       default: 0.011 s
    flDiffusion As Single            ' [0.0, 100.0]     default: 100.0 %
    flDensity As Single              ' [0.0, 100.0]     default: 100.0 %
    flHFReference As Single          ' [20.0, 20000.0]  default: 5000.0 Hz
End Type

Type BASS_DX8_PARAMEQ
    fCenter As Single
    fBandwidth As Single
    fGain As Single
End Type

Type BASS_DX8_REVERB
    fInGain As Single                ' [-96.0,0.0]            default: 0.0 dB
    fReverbMix As Single             ' [-96.0,0.0]            default: 0.0 db
    fReverbTime As Single            ' [0.001,3000.0]         default: 1000.0 ms
    fHighFreqRTRatio As Single       ' [0.001,0.999]          default: 0.001
End Type

Type Guid       ' used with BASS_Init - use VarPtr(guid) in clsid parameter
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type



Declare Function BASS_GetVersion Lib "bass.dll" () As Long
Declare Function BASS_ErrorGetCode Lib "bass.dll" () As Long
Declare Function BASS_Init Lib "bass.dll" (ByVal device As Long, ByVal freq As Long, ByVal flags As Long, ByVal win As Long, ByVal clsid As Long) As Long
Declare Function BASS_Free Lib "bass.dll" () As Long

Declare Function BASS_PluginGetInfo_ Lib "bass.dll" Alias "BASS_PluginGetInfo" (ByVal handle As Long) As Long

Declare Function BASS_SetEAXParameters Lib "bass.dll" (ByVal env As Long, ByVal vol As Single, ByVal decay As Single, ByVal damp As Single) As Long

Declare Function BASS_MusicLoad64 Lib "bass.dll" Alias "BASS_MusicLoad" (ByVal mem As Long, ByVal file As Any, ByVal offset As Long, ByVal offsethigh As Long, ByVal Length As Long, ByVal flags As Long, ByVal freq As Long) As Long

Declare Function BASS_SampleLoad64 Lib "bass.dll" Alias "BASS_SampleLoad" (ByVal mem As Long, ByVal file As Any, ByVal offset As Long, ByVal offsethigh As Long, ByVal Length As Long, ByVal max As Long, ByVal flags As Long) As Long

Declare Function BASS_StreamCreateFile64 Lib "bass.dll" Alias "BASS_StreamCreateFile" (ByVal mem As Long, ByVal file As Any, ByVal offset As Long, ByVal offsethigh As Long, ByVal Length As Long, ByVal lengthhigh As Long, ByVal flags As Long) As Long
Declare Function BASS_StreamCreateURL Lib "bass.dll" (ByVal url As String, ByVal offset As Long, ByVal flags As Long, ByVal proc As Long, ByVal User As Long) As Long

Declare Function BASS_ChannelBytes2Seconds64 Lib "bass.dll" Alias "BASS_ChannelBytes2Seconds" (ByVal handle As Long, ByVal pos As Long, ByVal poshigh As Long) As Double
Declare Function BASS_ChannelIsActive Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelPlay Lib "bass.dll" (ByVal handle As Long, ByVal restart As Long) As Long
Declare Function BASS_ChannelStop Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelSetAttribute Lib "bass.dll" (ByVal handle As Long, ByVal attrib As Long, ByVal Value As Single) As Long
Declare Function BASS_ChannelSetPosition64 Lib "bass.dll" Alias "BASS_ChannelSetPosition" (ByVal handle As Long, ByVal pos As Long, ByVal poshigh As Long, ByVal Mode As Long) As Long
Declare Function BASS_ChannelSetSync64 Lib "bass.dll" Alias "BASS_ChannelSetSync" (ByVal handle As Long, ByVal type_ As Long, ByVal param As Long, ByVal paramhigh As Long, ByVal proc As Long, ByVal User As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Public Function BASS_SPEAKER_N(ByVal n As Long) As Long
BASS_SPEAKER_N = n * (2 ^ 24)
End Function

' 32-bit wrappers for 64-bit BASS functions
Function BASS_MusicLoad(ByVal mem As Long, ByVal file As Long, ByVal offset As Long, ByVal Length As Long, ByVal flags As Long, ByVal freq As Long) As Long
BASS_MusicLoad = BASS_MusicLoad64(mem, file, offset, 0, Length, flags Or BASS_UNICODE, freq)
End Function

Function BASS_SampleLoad(ByVal mem As Long, ByVal file As Long, ByVal offset As Long, ByVal Length As Long, ByVal max As Long, ByVal flags As Long) As Long
BASS_SampleLoad = BASS_SampleLoad64(mem, file, offset, 0, Length, max, flags Or BASS_UNICODE)
End Function

Function BASS_StreamCreateFile(ByVal mem As Long, ByVal file As Long, ByVal offset As Long, ByVal Length As Long, ByVal flags As Long) As Long
BASS_StreamCreateFile = BASS_StreamCreateFile64(mem, file, offset, 0, Length, 0, flags Or BASS_UNICODE)
End Function

Function BASS_ChannelBytes2Seconds(ByVal handle As Long, ByVal pos As Long) As Double
BASS_ChannelBytes2Seconds = BASS_ChannelBytes2Seconds64(handle, pos, 0)
End Function

Function BASS_ChannelSetPosition(ByVal handle As Long, ByVal pos As Long, ByVal Mode As Long) As Long
BASS_ChannelSetPosition = BASS_ChannelSetPosition64(handle, pos, 0, Mode)
End Function

Function BASS_ChannelSetSync(ByVal handle As Long, ByVal type_ As Long, ByVal param As Long, ByVal proc As Long, ByVal User As Long) As Long
BASS_ChannelSetSync = BASS_ChannelSetSync64(handle, type_, param, 0, proc, User)
End Function

' BASS_PluginGetInfo wrappers
Function BASS_PluginGetInfo(ByVal handle As Long) As BASS_PLUGININFO
Dim pinfo As BASS_PLUGININFO, plug As Long
plug = BASS_PluginGetInfo_(handle)
If plug Then
    Call CopyMemory(pinfo, ByVal plug, LenB(pinfo))
End If
BASS_PluginGetInfo = pinfo
End Function

Function BASS_PluginGetInfoFormat(ByVal handle As Long, ByVal Index As Long) As BASS_PLUGINFORM
Dim pform As BASS_PLUGINFORM, plug As Long
plug = BASS_PluginGetInfo(handle).formats
If plug Then
    plug = plug + (Index * LenB(pform))
    Call CopyMemory(pform, ByVal plug, LenB(pform))
End If
BASS_PluginGetInfoFormat = pform
End Function

' callback functions
Function STREAMPROC(ByVal handle As Long, ByVal buffer As Long, ByVal Length As Long, ByVal User As Long) As Long
    
    'CALLBACK FUNCTION !!!
    
    ' User stream callback function
    ' NOTE: A stream function should obviously be as quick
    ' as possible, other streams (and MOD musics) can't be mixed until it's finished.
    ' handle : The stream that needs writing
    ' buffer : Buffer to write the samples in
    ' length : Number of bytes to write
    ' user   : The 'user' parameter value given when calling BASS_StreamCreate
    ' RETURN : Number of bytes written. Set the BASS_STREAMPROC_END flag to end
    '          the stream.
    
End Function

Sub DOWNLOADPROC(ByVal buffer As Long, ByVal Length As Long, ByVal User As Long)
    
    'CALLBACK FUNCTION !!!

    ' Internet stream download callback function.
    ' buffer : Buffer containing the downloaded data... NULL=end of download
    ' length : Number of bytes in the buffer
    ' user   : The 'user' parameter given when calling BASS_StreamCreateURL
    
End Sub

Sub SYNCPROC(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal User As Long)
    
    'CALLBACK FUNCTION !!!
    
    'Similarly in here, write what to do when sync function
    'is called, i.e screen flash etc.
    
    ' NOTE: a sync callback function should be very quick as other
    ' syncs cannot be processed until it has finished.
    ' handle : The sync that has occured
    ' channel: Channel that the sync occured in
    ' data   : Additional data associated with the sync's occurance
    ' user   : The 'user' parameter given when calling BASS_ChannelSetSync */
    
End Sub

Sub DSPPROC(ByVal handle As Long, ByVal channel As Long, ByVal buffer As Long, ByVal Length As Long, ByVal User As Long)

    'CALLBACK FUNCTION !!!

    ' VB doesn't support pointers, so you should copy the buffer into an array,
    ' process it, and then copy it back into the buffer.

    ' DSP callback function. NOTE: A DSP function should obviously be as quick as
    ' possible... other DSP functions, streams and MOD musics can not be processed
    ' until it's finished.
    ' handle : The DSP handle
    ' channel: Channel that the DSP is being applied to
    ' buffer : Buffer to apply the DSP to
    ' length : Number of bytes in the buffer
    ' user   : The 'user' parameter given when calling BASS_ChannelSetDSP
    
End Sub

Function RECORDPROC(ByVal handle As Long, ByVal buffer As Long, ByVal Length As Long, ByVal User As Long) As Long

    'CALLBACK FUNCTION !!!

    ' Recording callback function.
    ' handle : The recording handle
    ' buffer : Buffer containing the recorded samples
    ' length : Number of bytes
    ' user   : The 'user' parameter value given when calling BASS_RecordStart
    ' RETURN : BASSTRUE = continue recording, BASSFALSE = stop

End Function

' User file stream callback functions (BASS_FILEPROCS)
Sub FILECLOSEPROC(ByVal User As Long)

End Sub

Function FILELENPROC(ByVal User As Long) As Currency ' ???

End Function

Function FILEREADPROC(ByVal buffer As Long, ByVal Length As Long, ByVal User As Long) As Long

End Function

Function FILESEEKPROC(ByVal offset As Long, ByVal offsethigh As Long, ByVal User As Long) As Long

End Function


Function BASS_SetEAXPreset(preset) As Long
' This function is a workaround, because VB doesn't support multiple comma seperated
' paramaters for each Global Const, simply pass the EAX_ENVIRONMENT_xxx value to this function
' instead of BASS_SetEAXParameters as you would do in C/C++
Select Case preset
    Case EAX_ENVIRONMENT_GENERIC
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_GENERIC, 0.5, 1.493, 0.5)
    Case EAX_ENVIRONMENT_PADDEDCELL
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_PADDEDCELL, 0.25, 0.1, 0)
    Case EAX_ENVIRONMENT_ROOM
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_ROOM, 0.417, 0.4, 0.666)
    Case EAX_ENVIRONMENT_BATHROOM
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_BATHROOM, 0.653, 1.499, 0.166)
    Case EAX_ENVIRONMENT_LIVINGROOM
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_LIVINGROOM, 0.208, 0.478, 0)
    Case EAX_ENVIRONMENT_STONEROOM
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_STONEROOM, 0.5, 2.309, 0.888)
    Case EAX_ENVIRONMENT_AUDITORIUM
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_AUDITORIUM, 0.403, 4.279, 0.5)
    Case EAX_ENVIRONMENT_CONCERTHALL
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_CONCERTHALL, 0.5, 3.961, 0.5)
    Case EAX_ENVIRONMENT_CAVE
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_CAVE, 0.5, 2.886, 1.304)
    Case EAX_ENVIRONMENT_ARENA
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_ARENA, 0.361, 7.284, 0.332)
    Case EAX_ENVIRONMENT_HANGAR
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_HANGAR, 0.5, 10, 0.3)
    Case EAX_ENVIRONMENT_CARPETEDHALLWAY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_CARPETEDHALLWAY, 0.153, 0.259, 2)
    Case EAX_ENVIRONMENT_HALLWAY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_HALLWAY, 0.361, 1.493, 0)
    Case EAX_ENVIRONMENT_STONECORRIDOR
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_STONECORRIDOR, 0.444, 2.697, 0.638)
    Case EAX_ENVIRONMENT_ALLEY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_ALLEY, 0.25, 1.752, 0.776)
    Case EAX_ENVIRONMENT_FOREST
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_FOREST, 0.111, 3.145, 0.472)
    Case EAX_ENVIRONMENT_CITY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_CITY, 0.111, 2.767, 0.224)
    Case EAX_ENVIRONMENT_MOUNTAINS
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_MOUNTAINS, 0.194, 7.841, 0.472)
    Case EAX_ENVIRONMENT_QUARRY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_QUARRY, 1, 1.499, 0.5)
    Case EAX_ENVIRONMENT_PLAIN
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_PLAIN, 0.097, 2.767, 0.224)
    Case EAX_ENVIRONMENT_PARKINGLOT
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_PARKINGLOT, 0.208, 1.652, 1.5)
    Case EAX_ENVIRONMENT_SEWERPIPE
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_SEWERPIPE, 0.652, 2.886, 0.25)
    Case EAX_ENVIRONMENT_UNDERWATER
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_UNDERWATER, 1, 1.499, 0)
    Case EAX_ENVIRONMENT_DRUGGED
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_DRUGGED, 0.875, 8.392, 1.388)
    Case EAX_ENVIRONMENT_DIZZY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_DIZZY, 0.139, 17.234, 0.666)
    Case EAX_ENVIRONMENT_PSYCHOTIC
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_PSYCHOTIC, 0.486, 7.563, 0.806)
End Select
End Function

Public Function LoByte(ByVal lParam As Long) As Long
LoByte = lParam And &HFF&
End Function
Public Function HiByte(ByVal lParam As Long) As Long
HiByte = (lParam And &HFF00&) / &H100&
End Function
'Public Function LoWord(ByVal lParam As Long) As Long
'LoWord = lParam And &HFFFF&
'End Function
'Public Function HiWord(ByVal lParam As Long) As Integer
'If lParam < 0 Then
'    HiWord = CInt((lParam \ &H10000 - 1) And &HFFFF&)
'Else
'    HiWord = lParam \ &H10000
'End If
'End Function
Public Function HiWord(dwDWord As Long) As Integer
    ' REQUIRES: None
    ' MODIFIES: None
    '  EFFECTS: This function returns the high-order word (as
    '           signed integer) of the specified double-word
    '           (passed in as a signed long)
    Dim dW& ' To handle the sign bit
    dW& = IIf(dwDWord >= 0&, dwDWord \ &H10000, _
          &HFFFF& + dwDWord \ &H10001) And &HFFFF&
    If (dW& >= &H8000&) Then dW& = dW& - &H10000
    HiWord = CInt(dW&)
End Function ' HiWord()
Public Function LoWord(dwDWord As Long) As Integer
    ' REQUIRES: None
    ' MODIFIES: None
    '  EFFECTS: This function returns the low-order word (as
    '           signed integer) of the specified double-word
    '           (passed in as a signed long)
    Dim dW& ' To handle the sign bit
    dW& = dwDWord And &HFFFF&
    If (dW& >= &H8000&) Then _
           dW& = dW& - &H10000
    LoWord = CInt(dW&)
End Function ' LoWord()
Function MakeWord(ByVal LoByte As Long, ByVal HiByte As Long) As Long
MakeWord = (LoByte And &HFF&) Or ((HiByte And &HFF&) * &H100&)
End Function
Function MakeLong(ByVal LoWord As Long, ByVal HiWord As Long) As Long
MakeLong = LoWord And &HFFFF&
HiWord = HiWord And &HFFFF&
If HiWord And &H8000& Then
    MakeLong = MakeLong Or (((HiWord And &H7FFF&) * &H10000) Or &H80000000)
Else
    MakeLong = MakeLong Or (HiWord * &H10000)
End If
End Function

Public Function VBStrFromAnsiPtr(ByVal lpStr As Long) As String
Dim bStr() As Byte
Dim cChars As Long
On Error Resume Next
' Get the number of characters in the buffer
cChars = lstrlen(lpStr)
If cChars Then
    ' Resize the byte array
    ReDim bStr(0 To cChars - 1) As Byte
    ' Grab the ANSI buffer
    Call CopyMemory(bStr(0), ByVal lpStr, cChars)
End If
' Now convert to a VB Unicode string
VBStrFromAnsiPtr = StrConv(bStr, vbUnicode)
End Function
