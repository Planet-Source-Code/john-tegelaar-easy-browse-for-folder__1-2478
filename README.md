<div align="center">

## Easy Browse For Folder


</div>

### Description

Let the user select a directory (as workingdirectory, or for save/load data to/from, etc.), which is also known as a "Browse for Folder" function. This is the most easy way to provide this feature, just using the standard Common Dialog control.
 
### More Info
 
Start a new VB5/6 project, and put a CommandButton and a CommonDialog control on

the form. Paste in this code and you're ready to go.

(c)1999 John Tegelaar, The Netherlands

Selected directory path


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Tegelaar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-tegelaar.md)
**Level**          |Unknown
**User Rating**    |4.9 (49 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-tegelaar-easy-browse-for-folder__1-2478/archive/master.zip)





### Source Code

```
'Browse For Folder - the easy way
'
'Providing your users an elegant method of selecting a folder(as working
'directory, or to save/load files to/from, or....) is often desired, but
'hard to implement. Cumbersome SHFileOpen routines, or complicated
'hand-made alternatives are needed. Pardon... were needed.
'Although it is said at several places, including Microsoft (!), you can't
'do "Browse for Folder" with a Common Dialog control, I'll show you it can be
'done. Quick and Easy. And with very familiar interface to the users, including
'all standard options for navigating and browsing - even creation of a new folder
'and use of network paths.
'
'Start a new VB6 project, and put a CommandButton and a CommonDialog control on
'the form. Paste in this code and you're ready to go.
'(c)1999 John Tegelaar, The Netherlands
Option Explicit
Dim sTempDir As String
Dim sMyNewDirectory As String
Private Sub Command1_Click()
'Set up the CommonDialog control
On Local Error Resume Next     'Don't break on errors here
sTempDir = CurDir          'Store the current active directory
CommonDialog1.DialogTitle = "Select a directory" 'Titlebar caption
CommonDialog1.InitDir = App.Path  'Folder to start with, might be "C:\" or so also
CommonDialog1.FileName = "Select a Directory" 'Put something in filenamebox
CommonDialog1.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly 'Set CD Flags
'Here comes the big trick
CommonDialog1.Filter = "Folders|*.~#!"
'This reads as "show the user 'Folders' as filetype", while the files-filter
'is specified as being an impossible filetype. This causes the dialog to show
'folders only (as there's no matching file found).
CommonDialog1.CancelError = True  'allow escape key/cancel
CommonDialog1.ShowSave       'show the dialog.
'Note: ShowSave has more approperiate button captions then ShowOpen in this case.
If Err <> 32755 Then        'User didn't chose Cancel.
  sMyNewDirectory = CurDir    'CurDir has been changed to the selected one
  MsgBox ("Directory selected: " & sMyNewDirectory) 'Show the result
End If
ChDir sTempDir           'restore path to what it was at entering
End Sub
```

