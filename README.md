<div align="center">

## DecToBin


</div>

### Description

Converting numbers to binary

algie@tcp.co.uk (Alan Davis)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-dectobin__1-558/archive/master.zip)





### Source Code

```
Public Function DecToBin(ByVal DecNumber As Currency) As String
On Error GoTo DecToBin_Finally
Dim BinNumber As String
Dim i%
  For i = 64 To 0 Step -1
    If Int(DecNumber / (2 ^ i)) = 1 Then
      BinNumber = BinNumber & "1"
      DecNumber = DecNumber - (2 ^ i)
    Else
      If BinNumber <> "" Then
        BinNumber = BinNumber & "0"
      End If
    End If
  Next
  DecToBin = BinNumber
DecToBin_Finally:
  If Err <> 0 Or BinNumber = "" Then DecToBin = "-E-"
  Exit Function
End Function
```

