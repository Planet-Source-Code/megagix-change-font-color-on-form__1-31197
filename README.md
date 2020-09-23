<div align="center">

## Change Font Color On Form


</div>

### Description

This code will change the forecolor for Labels and Textboxes (can be edited for Other Controls).
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Megagix](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/megagix.md)
**Level**          |Beginner
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/megagix-change-font-color-on-form__1-31197/archive/master.zip)





### Source Code

```
Sub ChangeFont(frm As Form, color As String)
On Error GoTo errhand
Dim Control
For Each Control In frm.Controls
If TypeOf Control Is Label Then Control.ForeColor = color
If TypeOf Control Is TextBox Then Control.ForeColor = color
Next Control
Exit Sub
errhand:
MsgBox "Error Changing Font Color!", vbCritical
End Sub
```

