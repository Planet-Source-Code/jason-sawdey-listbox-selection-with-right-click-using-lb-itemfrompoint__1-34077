<div align="center">

## Listbox selection with Right Click \(using LB\_ITEMFROMPOINT\)


</div>

### Description

Allows selection of listbox items with right-click. *Not trying to get any votes, just sharing help I've provided in VB Discussion forum to everyone. Enjoy.*
 
### More Info
 
Just add a listbox with the default name of List1 to the form, and paste the code into the form.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jason Sawdey](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-sawdey.md)
**Level**          |Beginner
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jason-sawdey-listbox-selection-with-right-click-using-lb-itemfrompoint__1-34077/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
Private Const LB_ITEMFROMPOINT = &H1A9
```


### Source Code

```
Private Sub Form_Load()
  Dim i As Integer
  'Fill the listbox
  For i = 1 To 5
    List1.AddItem "Item " & i
  Next
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lRet As Long
  Dim lXPos As Long, lYPos As Long
  'Convert the cursor position into pixels, because that is what is needed
  lXPos = CLng(X / Screen.TwipsPerPixelX)
  lYPos = CLng(Y / Screen.TwipsPerPixelY)
  'If the right mouse button is clicked...
  If Button = 2 Then
    'Get the listitem closest to the cursor
    'NOTE: Since the X and Y values have to be in the form of high and low
    'order words, send the values as ((lYPos * 65536) + lXPos)
    lRet = SendMessage(List1.hWnd, LB_ITEMFROMPOINT, 0, ByVal _
      ((lYPos * 65536) + lXPos))
    'If the returned value is a valid index, then set that item as the selected
    'item
    If lRet < List1.ListCount Then
      List1.ListIndex = lRet
    End If
  End If
End Sub
```

