Attribute VB_Name = "modBassMidi"
' BASSMidI 2.4 Visual Basic module
' Copyright (c) 2006-2013 Un4seen Developments Ltd.
'
' See the BASSMidI.CHM file for more detailed documentation

Type BASS_MidI_FONT
    font As Long            ' soundfont
    preset As Long          ' preset number (-1=all)
    bank As Long
End Type

Type BASS_MidI_FONTEX
    font As Long            ' soundfont
    spreset As Long         ' source preset number
    sbank As Long           ' source bank number
    dpreset As Long         ' destination preset/program number
    dbank As Long           ' destination bank number
    dbanklsb As Long        ' destination bank number LSB
End Type

Type BASS_MidI_FONTINFO
    name As Long
    copyright As Long
    comment As Long
    presets As Long         ' number of presets/instruments
    samsize As Long         ' total size (in bytes) of the sample data
    samload As Long         ' amount of sample data currently loaded
    samtype As Long         ' sample format (CTYPE) if packed
End Type

Type BASS_MidI_MARK
    track As Long           ' track containing marker
    pos As Long             ' marker position
    text As Long            ' marker text
End Type


Type BASS_MidI_EVENT
        event_ As Long          ' MidI_EVENT_xxx
        param As Long
        chan As Long
        tick As Long            ' Event position (ticks)
        pos As Long             ' Event position (bytes)
End Type




Type BASS_MidI_DEVICEINFO
        name As Long    ' description
        id As Long
        flags As Long
End Type

Declare Function BASS_MidI_StreamCreateFile64 Lib "bassMidi.dll" Alias "BASS_MidI_StreamCreateFile" (ByVal mem As Long, ByVal file As Any, ByVal offset As Long, ByVal offsethi As Long, ByVal Length As Long, ByVal lengthhi As Long, ByVal flags As Long, ByVal freq As Long) As Long
Declare Function BASS_MidI_StreamCreateURL Lib "bassMidi.dll" (ByVal url As String, ByVal offset As Long, ByVal flags As Long, ByVal proc As Long, ByVal User As Long, ByVal freq As Long) As Long


' 32-bit wrappers for 64-bit BASS functions
Function BASS_MidI_StreamCreateFile(ByVal mem As Long, ByVal file As Long, ByVal offset As Long, ByVal Length As Long, ByVal flags As Long, ByVal freq As Long) As Long
BASS_MidI_StreamCreateFile = BASS_MidI_StreamCreateFile64(mem, file, offset, 0, Length, 0, flags Or BASS_UNICODE, freq)
End Function

' callback functions
Sub MidIINPROC(ByVal device As Long, ByVal time As Double, ByVal buffer As Long, ByVal Length As Long, ByVal User As Long)
    
    'CALLBACK FUNCTION !!!
    
    ' User MidI input callback function
    ' device : MIDI input device
    ' time   : Timestamp
    ' buffer : Buffer containing MIDI data
    ' length : Number of bytes of data
    ' user   : The 'user' parameter value given when calling BASS_MIDI_InInit
    
End Sub
