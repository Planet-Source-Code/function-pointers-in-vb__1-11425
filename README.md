<div align="center">

## Function Pointers in VB


</div>

### Description

This code makes it possible to use function pointers (no, not OBJPtr) basic funciton pointers, in VB.

It uses two API calls. Basically, it uses a window callback as the entry point for your new function, redirecting the window call to the function supplied.

A little out of the norm, but a handy thing to have when callbacks to non-objects is needed.

If you are attempting to call objects back, I recommend either the undocumented ObjPtr() or CallByName (when you know the name of the object function to call)
 
### More Info
 
No inputs

Nothing, absolutely nothing.

No Returns, just reap the benefits.

None that I know of, been using this for two years.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/function-pointers-in-vb__1-11425/archive/master.zip)

### API Declarations

```
Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal
lpPrevWndFunc&, ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any,
lpvSource As Any, ByVal cbCopy As Long)
```


### Source Code

```
Option Explicit
Private Function StripStringFromPointer$(ByVal lpString&, ByVal nStrLen&)
  Dim Info$
  Info = String$(nStrLen, vbNullChar)
  CopyMemory ByVal StrPtr(Info), ByVal lpString, nStrLen * 2
  StripStringFromPointer = Info
End Function
Private Function GetAddress(Addr&)
  GetAddress = Addr
End Function
Private Function MyFunction&(ByVal lpString&, ByVal nStrLen&, ByVal param3&,
ByVal param4&)
  Debug.Print StripStringFromPointer(lpString, nStrLen)
End Function
Public Sub Main()
  Dim FunctAddr&, Info$
  Info = "Holy Smoke"
  FunctAddr = GetAddress(AddressOf MyFunction)
  CallWindowProc FunctAddr, StrPtr(Info), CLng(Len(Info)), 0&, 0&
  End
End Sub
```

