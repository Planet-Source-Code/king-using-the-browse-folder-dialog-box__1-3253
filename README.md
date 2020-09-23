<div align="center">

## Using the Browse Folder Dialog Box


</div>

### Description

You may have noticed that in Windows the Browse Folder dialog is used in may programs, even the shell if you have used the find program you can choose browse and the folder below appears.WITH NO NEED FOR MODULES!!! You can implement this dialog bow into your applications very easily by using the following API calls.

SHBrowseForFolder

SHGetPathFromIDList

lstrcat
 
### More Info
 
a command button


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[King](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/king.md)
**Level**          |Unknown
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/king-using-the-browse-folder-dialog-box__1-3253/archive/master.zip)





### Source Code

```
Option Explicit
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Declare Function SHBrowseForFolder Lib _
	"shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib _
	"shell32" (ByVal pidList As Long, ByVal lpBuffer _
	As String) As Long
Private Declare Function lstrcat Lib "kernel32" _
	Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
	lpString2 As String) As Long
Private Type BrowseInfo
	hWndOwner As Long
	pIDLRoot As Long
	pszDisplayName As Long
	lpszTitle As Long
	ulFlags As Long
	lpfnCallback As Long
	lParam As Long
	iImage As Long
End Type
Private Sub Command1_Click()
'Opens a Browse Folders Dialog Box that displays the
'directories in your computer
Dim lpIDList As Long ' Declare Varibles
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo
szTitle = "Hello World. Click on a directory and " & _
	"it's path will be displayed in a message box"
' Text to appear in the the gray area under the title bar
' telling you what to do
With tBrowseInfo
	.hWndOwner = Me.hWnd ' Owner Form
	.lpszTitle = lstrcat(szTitle, "")
	.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
	sBuffer = Space(MAX_PATH)
	SHGetPathFromIDList lpIDList, sBuffer
	sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
	MsgBox sBuffer
End If
End Sub
```

