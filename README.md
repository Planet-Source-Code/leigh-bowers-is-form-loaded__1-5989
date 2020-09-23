<div align="center">

## Is Form Loaded?


</div>

### Description

This function will tell you if your form is loaded (instantiated) or not. If you use the form's "Visible" property to determine this, the form is instantiated if it's not already loaded. This results in the form's "Load" event being executed unnecessarily. This function has none of the above overheads...
 
### More Info
 
Syntax: If IsLoaded("frmTest") Then...

True (loaded) or False (not loaded)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Leigh Bowers](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/leigh-bowers.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/leigh-bowers-is-form-loaded__1-5989/archive/master.zip)





### Source Code

```
Public Function IsLoaded(sForm As String) as Boolean
Dim Frm As Form
' Loop through the Forms collection looking
' for the form of interest...
 For Each Frm In Forms
 If Frm.Name = sForm Then
  ' Found form in the collection
  IsLoaded = True
  Exit For
 End If
 Next
End Function
```

