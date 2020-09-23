Attribute VB_Name = "mdlSound"
Private Declare Function PlaySound& Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long)
    Global Const SND_SYNC = &H0 'play sound synchronously (return control after sound finishes)
    Public Const SND_ASYNC = &H1 'play sound asynchronously  (return control after sound begins)
    Global Const SND_NODEFAULT = &H2 'if sound/device unavailable--fail attempt
    Global Const SND_LOOP = &H8 'repeat the sound until the function is called again
    Global Const SND_NOSTOP = &H10 'if currently a sound is played the
    Global Const SND_NOWAIT& = &H2000 ' Do not use with SND_ALIAS or SND_FILENAME
    Global Const SND_RESOURCE& = &H40004
    Global Const SND_ALIAS& = &H10000
    Global Const SND_FILENAME& = &H20000
    Public Const enumSourceFile = SND_FILENAME&
    Public Const enumSourceRegistry = SND_ALIAS&
Public Function Play_Sound(strSoundName As String, intParamaters As Long) As Boolean
    If (PlaySound(strSoundName, 0&, intParamaters + SND_ASYNC + SND_NODEFAULT)) Then
        Play_Sound = True
    Else
        Play_Sound = False
    End If
End Function


