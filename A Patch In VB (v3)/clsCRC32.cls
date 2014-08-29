VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Fletcher checksum (CRC32)
' ---------------------------------------------------------------
' Based on the implementation by Mark Adler
' 2002 Ed Preston - epreston@selectedsystems.com
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=38270&lngWId=1
'
' Interface
'
'   CRC32(ByVal lngCRC32 As Long, ByRef bArrayIn() As Byte, _
'   ByVal dblLength As Double) As Long
'
' Notes
'
'   Used for detection of data corruption. Provides an algorithm for
'   computing a unique (sort of) numeric value that represents the
'   composition of a file. For a faster/better method see Adler32
'   checksum.

'   Fletcher check insensitive to some single byte changes 0 <-> 255
'
'   When passing byte arrays be carefull of the file size.  It is best
'   to break the file into chunks and call CRC32 multiple times.
'   Byte array is limited to the amount of addressable memory in the
'   process space.  The outcome is, pass a 712 meg array and watch your
'   machine grind to a halt while the hard disk tries to page things
'   in and out of memory.
'
' Limitations
'
'   Files size assumed to be less than 2 gig.
' ---------------------------------------------------------------

' Variables to hold our lookup table and its status.
Private bolTableReady As Boolean
Private lngArrayCRCByte(0 To 255) As Long

' ----------------------------
' Public Methods
' ----------------------------

Public Function CRC32(ByVal lngCRC32 As Long, ByRef bArrayIn() As Byte, ByVal dblLength As Double) As Long
    Dim dblCurPos As Double
    Dim lngTemp As Long
    
    ' Make sure the table is ready, just in case.
    If Not bolTableReady Then bolTableReady = BuildTable
    
    ' Array could be empty
    If UBound(bArrayIn) < LBound(bArrayIn) Then
        ' Files with no data have a CRC of 0
        CRC32 = 0
    Else
        lngTemp = lngCRC32 Xor &HFFFFFFFF
        
        For dblCurPos = 0 To dblLength
            lngTemp = (((lngTemp And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (lngArrayCRCByte((lngTemp And &HFF) Xor bArrayIn(dblCurPos)))
        Next dblCurPos
        
        ' Return the value
        CRC32 = lngTemp Xor &HFFFFFFFF
    End If
End Function

' ----------------------------
' Support Routines
' ----------------------------

' Create a lookup table of the CRC's of all possible eight
' bit values so we can do a byte-wise 32-bit CRC calculation.
' This will speed things up because we work at the byte
' level, else, we have to go bit by bit. Not good.
Private Function BuildTable() As Boolean
    Dim intBytePos As Integer
    Dim bBitPos As Byte
    
    ' Definition of polynomial for this type of CRC
    Const lngLimit = &HEDB88320
    
    Dim lngCRC As Long
    
    For intBytePos = 0 To 255 Step 1
        lngCRC = intBytePos
        
        For bBitPos = 0 To 7 Step 1
            If lngCRC And 1 Then
              lngCRC = (((lngCRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor lngLimit
            Else
              lngCRC = ((lngCRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next bBitPos
        
        ' Add the checksum of the current byte to the lookup table
        lngArrayCRCByte(intBytePos) = lngCRC
    Next intBytePos
    
    ' Table is ready.  Error handling routine would return false.
    BuildTable = True
End Function

Private Sub Class_Initialize()
    ' Build the lookup table when the class is created, this saves time
    ' if the methods of this class are not called immidiately
    bolTableReady = BuildTable
End Sub


