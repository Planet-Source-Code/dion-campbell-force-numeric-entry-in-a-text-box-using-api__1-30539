<div align="center">

## Force Numeric Entry in a Text Box Using API


</div>

### Description

This routine will cause a textbox control to accept only numeric input. Any other input (including the '-' and '.' keys) will be ignored. Note that it is still possible to paste non-numeric data into the textbox. There have been plenty of examples of this in the last couple of days - in the comments section of one of these someone suggested using the API. So, for anyone that is interested, this is one way of doing it...

I've put this in the intermediate category only because it uses API functions - otherwise it's straightforward code. I've only tested this in VB6.0 on Win2k, but it should work on any Windows platform from Win95 up.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dion Campbell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dion-campbell.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dion-campbell-force-numeric-entry-in-a-text-box-using-api__1-30539/archive/master.zip)

### API Declarations

```
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const ES_NUMBER = &H2000
Private Const GWL_STYLE = -16
```


### Source Code

```
Private Sub ForceNumeric(Box As TextBox)
  On Error GoTo Catch
  Dim nStyle As Long
  nStyle = GetWindowLong(Box.hWnd, GWL_STYLE)
  Call SetWindowLong(Box.hWnd, GWL_STYLE, nStyle Or ES_NUMBER)
  GoTo Finally
Catch:
  Call MsgBox(Err.Description, vbCritical Or vbOKOnly, "Error")
Finally:
End Sub
```

