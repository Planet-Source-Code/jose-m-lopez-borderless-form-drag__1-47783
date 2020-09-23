<div align="center">

## Borderless Form Drag


</div>

### Description

To drag a borderless form.

Very Easy and Simple.

Just a few lines of code.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jose M\. Lopez](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jose-m-lopez.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jose-m-lopez-borderless-form-drag__1-47783/archive/master.zip)





### Source Code

```
Dim DragX As Long, DragY As Long
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragX = X: DragY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Form1.Move Form1.Left + X - DragX, Form1.Top + Y - DragY
End If
End Sub
```

