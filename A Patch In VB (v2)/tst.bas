Attribute VB_Name = "Tst"
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' Returns the track's title
Function Title() As String
        Title = Space$(20)
        CopyMem ByVal Title, ByVal uFMOD_GetTitle, 20
        If Len(Trim$(Title)) = 0 Then Title = "{anonymous track}"
End Function

Sub Main()

        ' Start playback
        '    1 is the ID of the XM resource
        If uFMOD_PlaySong(1, 0, XM_RESOURCE) <> 0 Then
Form1.Show
        End If
End Sub
