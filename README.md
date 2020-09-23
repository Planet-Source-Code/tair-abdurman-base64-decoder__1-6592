<div align="center">

## base64 decoder


</div>

### Description

Decode base64 encoded Input file into Output file.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tair Abdurman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tair-abdurman.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tair-abdurman-base64-decoder__1-6592/archive/master.zip)





### Source Code

```
'(C) 2000 by Tair Abdurman
'WWW: www.tair.freeservers.com
'e-mail: broadcast_line@usa.net
'this version to decode Outlook encrypted
'attachments
'Base64 decode routines
' based on RFC 1421
'----------------------------------------------------------------------------------------------------
' Quantum of decoded content
'----------------------------------------------------------------------------------------------------
'    3       2       1       0
' 00XXXXXX 00XXXXXX 00XXXXXX 00XXXXXX
'   |    |   | | |  |   |  | | |  |    |
'    A1    A2 B1    B2  C1    C2
'----------------------------------------------------------------------------------------------------
' Bit positions:
'----------------------------------------------------------------------------------------------------
'      AND     SHIFT RIGHT   SHIFT LEFT     BYTE NUMB
'  A1  3FH         01H         08H          3
'  A2  30H         10H         01H          2
'
'  B1   0FH         01H        10H          2
'  B2   3CH         08H        01H          1
'
'  C1   03H         01H        40H          1
'  C2   3FH         01H        01H          0
'----------------------------------------------------------------------------------------------------
' Decoded Triple
'   DA      DB     DC
' XXXXXXXX XXXXXXXX XXXXXXXX
'----------------------------------------------------------------------------------------------------
'  VB Formula:
'  Ydecoded(DZ)=(Xencoded(Z1bytenum) AND Z1and)*Z1shiftright +
'          (Xencoded(Z2bytenum) AND Z2and)/Z2shiftleft
'----------------------------------------------------------------------------------------------------
Option Explicit
Private Type b64encoded
   Byte1 As Byte
   Byte2 As Byte
   Byte3 As Byte
   Byte4 As Byte
End Type
Private Type b64decoded
   Byte1 As Byte
   Byte2 As Byte
   Byte3 As Byte
End Type
Private Type codecodeBytes
   Byte1 As Byte
   Byte2 As Byte
   Byte3 As Byte
   Byte4 As Byte
End Type
Dim keyByteA As codecodeBytes
Dim keyByteB As codecodeBytes
Dim keyByteC As codecodeBytes
Private Sub InitDecodeEncodeMachine()
'-------------------------------
keyByteA.Byte1 = &H3F
keyByteA.Byte2 = &H4
keyByteA.Byte3 = &H30
keyByteA.Byte4 = &H10
'-------------------------------
'-------------------------------
keyByteB.Byte1 = &HF
keyByteB.Byte2 = &H10
keyByteB.Byte3 = &H3C
keyByteB.Byte4 = &H4
'-------------------------------
'-------------------------------
keyByteC.Byte1 = &H3
keyByteC.Byte2 = &H40
keyByteC.Byte3 = &H3F
keyByteC.Byte4 = &H1
'-------------------------------
End Sub
'Decode source file encoded by base64 into destination
Public Sub DecodeFile(ByVal srcFile As String, ByVal dstFile As String)
  Dim tempBuffer As String * 78
  Dim tempBufferNC As String * 74
  Dim tempEncoded As b64encoded
  Dim tempDecoded As b64decoded
  Dim bResult As Byte
  Dim iCntr As Long
  Dim btResult As Byte
  Call InitDecodeEncodeMachine
btResult = 0
iCntr = 0
  Open srcFile For Random As #1 Len = 78
  Open dstFile For Random As #2 Len = 1
   Do While Not (EOF(1))
    Get #1, , tempBuffer
    iCntr = 0
    Do While iCntr < Len(tempBuffer)
      If Mid(tempBuffer, (iCntr + 1), 2) = vbCrLf Then Exit Do
      tempEncoded.Byte1 = DeMapCode(Mid(tempBuffer, (iCntr + 1), 1))
      tempEncoded.Byte2 = DeMapCode(Mid(tempBuffer, (iCntr + 2), 1))
      tempEncoded.Byte3 = DeMapCode(Mid(tempBuffer, (iCntr + 3), 1))
      tempEncoded.Byte4 = DeMapCode(Mid(tempBuffer, (iCntr + 4), 1))
      bResult = 0
      bResult = Base64Decode(tempEncoded, tempDecoded)
      Select Case bResult
      Case 1
        Put #2, , tempDecoded.Byte1
      Case 2
        Put #2, , tempDecoded.Byte1
        Put #2, , tempDecoded.Byte2
      Case 3
        Put #2, , tempDecoded.Byte1
        Put #2, , tempDecoded.Byte2
        Put #2, , tempDecoded.Byte3
      End Select
      'EOF encoded part
      If (bResult = 0) Then Exit Do
      'FOUR bytes as step
      iCntr = iCntr + 4
    Loop
    'if end of encoded text
    If (bResult = 0) Then Exit Do
   Loop
  Close #2
  Close #1
End Sub
Private Function Base64Decode(srcBase64Encoded As b64encoded, dstBase64Decoded As b64decoded) As Byte
'return amoun of decoded bytes
If (srcBase64Encoded.Byte1 > 64) Then
 Base64Decode = 0
 Exit Function
End If
If ((srcBase64Encoded.Byte3 = 64) And (srcBase64Encoded.Byte4 = 64)) Then
 dstBase64Decoded.Byte1 = (srcBase64Encoded.Byte1 And keyByteA.Byte1) * keyByteA.Byte2 + _
                     (srcBase64Encoded.Byte2 And keyByteA.Byte3) / keyByteA.Byte4
 dstBase64Decoded.Byte2 = 0
 dstBase64Decoded.Byte3 = 0
 Base64Decode = 1
 Exit Function
End If
If (srcBase64Encoded.Byte4 = 64) Then
 dstBase64Decoded.Byte1 = (srcBase64Encoded.Byte1 And keyByteA.Byte1) * keyByteA.Byte2 + _
                    (srcBase64Encoded.Byte2 And keyByteA.Byte3) / keyByteA.Byte4
 dstBase64Decoded.Byte2 = (srcBase64Encoded.Byte2 And keyByteB.Byte1) * keyByteB.Byte2 + _
                    (srcBase64Encoded.Byte3 And keyByteB.Byte3) / keyByteB.Byte4
 dstBase64Decoded.Byte3 = 0
 Base64Decode = 2
 Exit Function
End If
dstBase64Decoded.Byte1 = (srcBase64Encoded.Byte1 And keyByteA.Byte1) * keyByteA.Byte2 + _
                    (srcBase64Encoded.Byte2 And keyByteA.Byte3) / keyByteA.Byte4
dstBase64Decoded.Byte2 = (srcBase64Encoded.Byte2 And keyByteB.Byte1) * keyByteB.Byte2 + _
                    (srcBase64Encoded.Byte3 And keyByteB.Byte3) / keyByteB.Byte4
dstBase64Decoded.Byte3 = (srcBase64Encoded.Byte3 And keyByteC.Byte1) * keyByteC.Byte2 + _
                    (srcBase64Encoded.Byte4 And keyByteC.Byte3) / keyByteC.Byte4
Base64Decode = 3
End Function
Private Function DeMapCode(srcChar As String) As Byte
  If Len(srcChar) <> 1 Then
    DeMapCode = 0
    Exit Function
  End If
  Select Case srcChar
    Case "A" To "Z"
        DeMapCode = Asc(srcChar) - 65
    Case "a" To "z"
        DeMapCode = Asc(srcChar) - 97 + 26
    Case "0" To "9"
        DeMapCode = Asc(srcChar) - 48 + 52
    Case "+"
        DeMapCode = 62
    Case "/"
        DeMapCode = 63
    Case "="
        DeMapCode = 64
    Case Else
        DeMapCode = 65
  End Select
End Function
```

