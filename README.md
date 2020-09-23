<div align="center">

## Credit Card Identification


</div>

### Description

Determines type of Credit Card by it's number.
 
### More Info
 
Card Number as String

This is based on documents from CyberCash's home page.

Card Type as String

Is not Year 2061 Compliant


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Anderson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-anderson.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-anderson-credit-card-identification__1-903/archive/master.zip)





### Source Code

```
Public Function CardType(CCNum As String) As String
Dim Header As String
  Select Case Left$(CCNum, 1)
    Case "5"
      Header = Left$(CCNum, 2)
      If Header >= 51 And Header <= 55 And Len(CCNum) = 16 Then
        CardType = "MasterCard"
      End If
    Case "4"
      If Len(CCNum) = 13 Or Len(CCNum) = 16 Then
        CardType = "Visa"
      End If
    Case "3"
      Header = Left$(CCNum, 3)
      If Header >= 340 And Header <= 379 And Len(CCNum) = 15 Then
        CardType = "AMEX"
      End If
      If Header >= 300 And Header <= 305 And Len(CCNum) = 14 Then
        CardType = "Diners Club"
      End If
      If Header >= 360 And Header <= 369 And Len(CCNum) = 14 Then
        CardType = "Diners Club"
      End If
      If Header >= 380 And Header <= 389 And Len(CCNum) = 14 Then
        CardType = "Diners Club"
      End If
      If Header >= 300 And Header <= 399 And Len(CCNum) = 16 Then
        CardType = "JCB"
      End If
    Case "6"
      Header = Left$(CCNum, 4)
      If Header = "6011" And Len(CCNum) = 16 Then
        CardType = "Discover"
      End If
    Case "2"
      Header = Left$(CCNum, 4)
      If (Header = "2014" Or Header = "2149") And Len(CCNum) = 15 Then
        CardType = "enRoute"
      End If
      If Header = "2131" And Len(CCNum) = 15 Then
        CardType = "JCB"
      End If
    Case "1"
      Header = Left$(CCNum, 4)
      If Header = "1800" And Len(CCNum) = 15 Then
        CardType = "JCB"
      End If
  End Select
  If CardType = "" Then CardType = "Unknown"
End Function
```

