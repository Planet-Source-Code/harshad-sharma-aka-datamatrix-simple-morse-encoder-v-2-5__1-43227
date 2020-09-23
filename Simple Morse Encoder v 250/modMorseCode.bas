Attribute VB_Name = "modMorseCode"
Option Explicit
' This API call is used to make the program "sleep" for some time...
' i.e. to pause execution of the program for the given milliseconds
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' This API is used to make the sound from the PC Speaker
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Enum OpDevice
    PCSpeaker
    SoundCard
End Enum

Public Sub PlayMorse(aCode As String, aSpeed As Integer, aFreq As Integer, aDevice As OpDevice)
' If the Code variable is not filled, then fill it with default value
' Use the trim function to first remove extra spaces
' and then see if there is anything else
Dim char As String * 1
If Trim(aCode) = "" Then aCode = " "
' Use the Len function to check the size of the string
' A valid string contains more than one charachter...
If Len(aCode) = 0 Then aCode = " "
' OK Now we have at least ONE charachter to transmit...

    Do While (Len(aCode) > 0)
    'check if we have been paused...
    If frmMain.aPause = False Then
        ' if we are not supposed to pause, then just proceed...
        ' take the left-most single charachter from the given string
        char = Left(aCode, 1)
        
        ' remove the left-most charachter from the code
        aCode = Right(aCode, Len(aCode) - 1)
        
        '>>>---------------------<W><A><R><N><I><N><G>------------------------
        ' Application Specific Code Here.... if you want to use the code
        ' csomewhere else, be sure to remove this line...
        frmMain.txtMorse.Text = aCode
        aSpeed = frmMain.sldSpeed.Max - frmMain.sldSpeed.Value
        '<<<---------------------<W><A><R><N><I><N><G>------------------------
        
        ' Observation:
        ' when speed = 64, we get 13 wpm
        DoEvents
        
        ' pause between bits...
        Sleep (aSpeed)
        
        Select Case char
        Case " "
            ' pause for 3 tu (time units)
            Sleep (3 * aSpeed)
            
        Case "."
            ' beep for one tu
            If aDevice = PCSpeaker Then
                Beep aFreq, aSpeed
            Else
                dBeep aFreq, aSpeed, 100
            End If
            
        Case "-"
            ' beep for 3 tu
            If aDevice = PCSpeaker Then
                Beep aFreq, 3 * aSpeed
            Else
                dBeep aFreq, 3 * aSpeed, 100
            End If
            
        Case "/"
            ' pause for 7 tu
            Sleep (7 * aSpeed)
            
        End Select
    End If ' //If frmMain.aPause = False Then// ends here
    ' if we don't allow the OS to perform other tasks (with our app)
    ' then when we pause the app, it wil HANG! (Find out why)
    ' To prevent that...
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    Loop ' //Do While (Len(aCode) > 0)// Ends Here
End Sub

Public Function EncodeToMorse(TextToEncode As String) As String
Dim x As Integer
Dim y As Integer
Dim char As String
Dim EncodedMorse As String
If Len(TextToEncode) > 1 Then
    For x = 0 To Len(TextToEncode) - 1
        char = Left(TextToEncode, 1)
        TextToEncode = Right(TextToEncode, Len(TextToEncode) - 1)
        EncodedMorse = EncodedMorse & " " & EncodeCharachter(char)
    Next
    EncodeToMorse = EncodedMorse
Else
    EncodeToMorse = ""
End If
End Function

Public Function DecodeToEnglish(CodeToDecode As String) As String
On Error Resume Next
Dim x As Integer
Dim y As Integer
Dim char As String
Dim DecodedEnglish As String
CodeToDecode = CodeToDecode & " "
If Len(CodeToDecode) > 1 Then
    For x = 0 To Len(CodeToDecode) '- 1
        char = char & Left(CodeToDecode, 1)
        If Len(CodeToDecode) > 0 Then
            CodeToDecode = Right(CodeToDecode, Len(CodeToDecode) - 1)
        End If
        ' when we encounter a space, the charachter is complete...
        If Left(CodeToDecode, 1) = " " Then
            DecodedEnglish = DecodedEnglish & DecodeCharachter(char)
            CodeToDecode = Right(CodeToDecode, Len(CodeToDecode) - 1)
            char = ""
        End If
    Next
    DecodeToEnglish = DecodedEnglish
Else
    DecodeToEnglish = ""
End If
End Function


Private Function EncodeCharachter(aCharachter As String) As String
aCharachter = LCase(aCharachter)
Select Case aCharachter
    Case " "
        EncodeCharachter = "/"     ' This may look odd, but it's only
                                    ' to keep a track of the spaces between
                                    ' words... as we have single spaces between
                                    ' each charachter, we need to explicitly
                                    ' manage the spaces between words.
    Case "a"
        EncodeCharachter = ".-"
    Case "b"
        EncodeCharachter = "-..."
    Case "c"
        EncodeCharachter = "-.-."
    Case "d"
        EncodeCharachter = "-.."
    Case "e"
        EncodeCharachter = "."
    Case "f"
        EncodeCharachter = "..-."
    Case "g"
        EncodeCharachter = "--."
    Case "h"
        EncodeCharachter = "...."
    Case "i"
        EncodeCharachter = ".."
    Case "j"
        EncodeCharachter = ".---"
    Case "k"
        EncodeCharachter = "-.-"
    Case "l"
        EncodeCharachter = ".-.."
    Case "m"
        EncodeCharachter = "--"
    Case "n"
        EncodeCharachter = "-."
    Case "o"
        EncodeCharachter = "---"
    Case "p"
        EncodeCharachter = ".--."
    Case "q"
        EncodeCharachter = "--.-"
    Case "r"
        EncodeCharachter = ".-."
    Case "s"
        EncodeCharachter = "..."
    Case "t"
        EncodeCharachter = "-"
    Case "u"
        EncodeCharachter = "..-"
    Case "v"
        EncodeCharachter = "...-"
    Case "w"
        EncodeCharachter = ".--"
    Case "x"
        EncodeCharachter = "-..-"
    Case "y"
        EncodeCharachter = "-.--"
    Case "z"
        EncodeCharachter = "--.."
    Case "1"
        EncodeCharachter = ".----"
    Case "2"
        EncodeCharachter = "..---"
    Case "3"
        EncodeCharachter = "...--"
    Case "4"
        EncodeCharachter = "....-"
    Case "5"
        EncodeCharachter = "....."
    Case "6"
        EncodeCharachter = "-...."
    Case "7"
        EncodeCharachter = "--..."
    Case "8"
        EncodeCharachter = "---.."
    Case "9"
        EncodeCharachter = "----."
    Case "0"
        EncodeCharachter = "-----"
    Case "."
        EncodeCharachter = ".-.-.-"
    Case "?"
        EncodeCharachter = "..--.."
    Case ","
        EncodeCharachter = "--..--"
    Case "'"
        EncodeCharachter = ".----."
    ' The other charachters which are not listed here
    ' are (as per my knowledge) NOT supported by the MORSE Code.
    '  If you know any which are mistakenly left out, then just
    ' copy some code from above lines and add the new charachters.
    ' And please let me know about it so that I can update this
    ' project accordingly.
    ' My email is : harshad.sharma@bigfoot.com
End Select
End Function

Private Function DecodeCharachter(aCharachter As String) As String
aCharachter = LCase(aCharachter)
Select Case aCharachter
    Case "/"
        DecodeCharachter = " "
    Case ".-"
        DecodeCharachter = "a"
    Case "-..."
        DecodeCharachter = "b"
    Case "-.-."
        DecodeCharachter = "c"
    Case "-.."
        DecodeCharachter = "d"
    Case "."
        DecodeCharachter = "e"
    Case "..-."
        DecodeCharachter = "f"
    Case "--."
        DecodeCharachter = "g"
    Case "...."
        DecodeCharachter = "h"
    Case ".."
        DecodeCharachter = "i"
    Case ".---"
        DecodeCharachter = "j"
    Case "-.-"
        DecodeCharachter = "k"
    Case ".-.."
        DecodeCharachter = "l"
    Case "--"
        DecodeCharachter = "m"
    Case "-."
        DecodeCharachter = "n"
    Case "---"
        DecodeCharachter = "o"
    Case ".--."
        DecodeCharachter = "p"
    Case "--.-"
        DecodeCharachter = "q"
    Case ".-."
        DecodeCharachter = "r"
    Case "..."
        DecodeCharachter = "s"
    Case "-"
        DecodeCharachter = "t"
    Case "..-"
        DecodeCharachter = "u"
    Case "...-"
        DecodeCharachter = "v"
    Case ".--"
        DecodeCharachter = "w"
    Case "-..-"
        DecodeCharachter = "x"
    Case "-.--"
        DecodeCharachter = "y"
    Case "--.."
        DecodeCharachter = "z"
    Case ".----"
        DecodeCharachter = "1"
    Case "..---"
        DecodeCharachter = "2"
    Case "...--"
        DecodeCharachter = "3"
    Case "....-"
        DecodeCharachter = "4"
    Case "....."
        DecodeCharachter = "5"
    Case "-...."
        DecodeCharachter = "6"
    Case "--..."
        DecodeCharachter = "7"
    Case "---.."
        DecodeCharachter = "8"
    Case "----."
        DecodeCharachter = "9"
    Case "-----"
        DecodeCharachter = "0"
    Case ".-.-.-"
        DecodeCharachter = "."
    Case "..--.."
        DecodeCharachter = "?"
    Case "--..--"
        DecodeCharachter = ","
    Case ".----."
        DecodeCharachter = "'"
    ' The other charachters which are not listed here
    ' are (as per my knowledge) NOT supported by the MORSE Code.
    '  If you know any which are mistakenly left out, then just
    ' copy some code from above lines and add the new charachters.
    ' And please let me know about it so that I can update this
    ' project accordingly.
    ' My email is : harshad.sharma@bigfoot.com
End Select
End Function

