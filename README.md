<div align="center">

## Base64 Encode


</div>

### Description

Converts a string of data to Base64. Base64 is an encryption algorithm used to encode binary data that is being sent through the internet.
 
### More Info
 
"sData" is the string that you want to encode.

It returns the encoded Base64 string.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andrew](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrew.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andrew-base64-encode__1-13478/archive/master.zip)





### Source Code

```
Public Function sBase64Enc(sData As String) As String
  'Base64 Conversion
  'Example:
  '  Dim sMyConv As String
  '  sMyConv = sBase64Enc("Hello =)")
  On Error Resume Next
  Dim x   As Long
  Dim nByte As Long
  Dim nAsc As Long
  Dim sBin As String
  Dim sRet As String
  Dim sByte As String
  Dim nIncr As Integer
  'Convert the data to standard
  'base-2 binary.
  For x = 1 To Len(sData)
    DoEvents
    nByte = CLng(Asc(Mid(sData, x, 1)))
    For y = 1 To 8
      nIncr = CInt(2 ^ (8 - y))
      If CLng(nByte) - CLng(nIncr) >= 0 Then
        nByte = nByte - CLng(nIncr)
        sBin = sBin & "1"
      Else: sBin = sBin & "0"
      End If
    Next y
  Next x
  'Check to see if the conversion was completed
  'and if so, encode the data using the Base64
  'algorithm.
  If CLng(Len(sBin) Mod 8) = 0 Then
    'Binary conversion ok!, parse
    'every 6 bits of data.
    For x = 1 To Len(sBin) Step 6
      DoEvents
      sByte = Mid(sBin, x, 6)
      For y = 1 To Len(sByte)
        DoEvents
        nByte = Val(Mid(sByte, y, 1))
        If Not nByte = 0 Then
          nAsc = nAsc + CInt(2 ^ (6 - (y)))
        End If
      Next y
      'Base64 Conversion:
      Select Case (nAsc + 65)
      Case Is > 90 'Either lowercase or numeric
        If (nAsc + 71) > 122 Then
          sByte = Chr(nAsc - 4)
        Else
          sByte = Chr(nAsc + 71)
        End If
      Case Is < 90 'Uppercase
        sByte = Chr(nAsc + 65)
      End Select
      'Append new characters to the final
      'string and reset temporary variables.
      sRet = sRet & sByte
      nAsc = 0
    Next x
  End If
  'Finished, output the data to the
  'function variable.
  sBase64Enc = sRet
End Function
```

