<div align="center">

## Convert to from Binary in very little code


</div>

### Description

Here is are two nice functions that will convert Decimal values to binary and binary to decimal in a surprisingly short amount of code.<BR><BR>

Comments welcome. Please
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MrEnigma](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mrenigma.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mrenigma-convert-to-from-binary-in-very-little-code__1-9692/archive/master.zip)





### Source Code

```
Public Function DecimalToBinary(sValue As String) As String
Dim i As Integer
Const sTable As String = "0000,0001,0010,0011,0100,0101,0110,0111,1000,1001,1010,1011,1100,1101,1110,1111"
Dim asBinTable() As String
Dim sHexValue As String
   If Len(sValue) > 9 Then
     ' the HEX Function cannot handle larger numbers
     Exit Function
   End If
   DecimalToBinary = ""
   ' Set up the Binary Table
   asBinTable = Split(sTable, ",")
   sHexValue = Hex(Val(sValue))
   For i = 1 To Len(sHexValue)
     DecimalToBinary = DecimalToBinary & asBinTable(Val("&H" & Mid$(sHexValue, i, 1)))
   Next
End Function
Public Function BinaryToDecimal(sBinary As String) As String
Dim i As Integer
   BinaryToDecimal = 0
   If Len(sBinary) > 49 Then
     ' Binary numbers larger than 49 bits
     ' Will return an Error E+
     Exit Function
   End If
   For i = 0 To Len(sBinary) - 1
     If Mid$(sBinary, Len(sBinary) - i, 1) Then
      BinaryToDecimal = BinaryToDecimal + 2 ^ i
     End If
   Next
End Function
```

