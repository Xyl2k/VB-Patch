Attribute VB_Name = "bassmod"
Option Explicit
Dim TempPfad As String
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function BASSMOD_Init Lib "bassmod.dll" (ByVal device As Long, ByVal freq As Long, ByVal flags As Long) As Integer
Declare Function BASSMOD_MusicLoad Lib "bassmod.dll" (ByVal mem As Integer, ByVal pfile As String, ByVal offset As Long, ByVal Length As Long, ByVal flags As Long) As Integer
Declare Function BASSMOD_MusicPlay Lib "bassmod.dll" () As Integer
Declare Function BASSMOD_MusicStop Lib "bassmod.dll" () As Integer
Declare Function BASSMOD_MusicPause Lib "bassmod.dll" () As Integer
Declare Sub BASSMOD_Free Lib "bassmod.dll" ()
Const MAX_PATH = 260
Public Function GetTmpPath()
    Dim sFolder As String
    Dim lRet As Long
    sFolder = String(MAX_PATH, 0)
    lRet = GetTempPath(MAX_PATH, sFolder)
    If lRet <> 0 Then
        GetTmpPath = Left(sFolder, InStr(sFolder, Chr(0)) - 1)
    Else
        GetTmpPath = vbNullString
    End If
End Function
Public Function GetSysPath()
    GetSysPath = Trim(Environ$("systemroot") & "\system32\")
End Function

Public Function CreateFileFromRessource(ByVal ID As Integer, FileName As String)
    Dim DataArray() As Byte
    DataArray = LoadResData(ID, "CUSTOM")
    Dim Handle As Integer
    Handle = FreeFile
    Open FileName For Binary As #Handle
        Put #Handle, , DataArray
    Close #Handle
    Erase DataArray
End Function
Public Function playXM(XMFileName As String, ByVal RES_ID As Integer)
    TempPfad = GetTmpPath
    If Dir(GetSysPath & "bassmod.dll") = vbNullString Then
        CreateFileFromRessource 101, GetSysPath & "bassmod.dll"
    End If
    CreateFileFromRessource RES_ID, TempPfad & XMFileName
    BASSMOD_Init -1, 44100, 0
   BASSMOD_MusicLoad 0, TempPfad & XMFileName, 0, 0, 4 Or 512
    BASSMOD_MusicPlay
End Function
Public Function stopXM(XMFileName As String)
    BASSMOD_MusicStop
    BASSMOD_Free
    If Dir(TempPfad & XMFileName) <> vbNullString Then
        Kill TempPfad & XMFileName
    End If
End Function



