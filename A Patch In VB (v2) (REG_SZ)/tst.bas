Attribute VB_Name = "Tst"
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' Returns the track's title
Function Title() As String
        Title = Space$(20)
        CopyMem ByVal Title, ByVal uFMOD_GetTitle, 20
        If Len(Trim$(Title)) = 0 Then Title = "{anonymous track}"
End Function


