Attribute VB_Name = "Module1"
Option Explicit
Public UserSettings As New clsSettings 'make visible to all forms
'granted, the UserSettings object could be just as easily implemented in this
'bas module, but the code is easier to read and self documenting using a class module


Public Declare Function GetOpenFileName Lib "comdlg32.dll" _
Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Enum FlagConstants
  cdlOFNAllowMultiselect = &H200 'The user can select more than one file atrun time by pressing the SHIFT key and using the UP ARROW and DOWN ARROW keys to select the desired files. When this is done, the FileName property returns a string containing the names of all selected files. The names in the string are delimited by spaces.
  cdlOFNCreatePrompt = &H2000 ' Specifies that the dialog box prompts the user to create a file that doesn't currently exist. This flag automatically sets the cdlOFNPathMustExist and cdlOFNFileMustExist flags.
  cdlOFNExplorer = &H80000 ' Use the Explorer-like Open A File dialog box template. Works with Windows 95 and Windows NT 4.0.
  CdlOFNExtensionDifferent = &H400 ' Indicates that the extension of the returned filename is different from the extension specified by the DefaultExt property. This flag isn't set if the DefaultExt property is Null, if the extensions match, or if the file has no extension. This flag value can be checked upon closing the dialog box.
  cdlOFNFileMustExist = &H1000 ' Specifies that the user can enter only names of existing files in the File Name text box. If this flag is set and the user enters an invalid filename, a warning is displayed. This flag automatically sets the cdlOFNPathMustExist flag.
  cdlOFNHelpButton = &H10 ' Causes the dialog box to display the Help button.
  cdlOFNHideReadOnly = &H4 'Hides the Read Onlycheck box.
  cdlOFNLongNames = &H200000 ' Use long filenames.
  cdlOFNNoChangeDir = &H8 'Forces the dialog box to set the current directory to what it was when the dialog box was opened.
  CdlOFNNoDereferenceLinks = &H100000 ' Do not dereference shell links (also known as shortcuts). By default, choosing a shell link causes it to be dereferenced by the shell.
  cdlOFNNoLongNames = &H40000 ' No long file names.
  CdlOFNNoReadOnlyReturn = &H8000 ' Specifies that the returned file won't have the Read Only attribute set and won't be in a write-protected directory.
  cdlOFNNoValidate = &H100 ' Specifies that the common dialog box allows invalid characters in the returned filename.
  cdlOFNOverwritePrompt = &H2 'Causes the Save As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.
  cdlOFNPathMustExist = &H800 ' Specifies that the user can enter only valid paths. If this flag is set and the user enters an invalid path, a warning message is displayed.
  cdlOFNReadOnly = &H1 'Causes the Read Only check box to be initially checked when the dialog box is created. This flag also indicates the state of the Read Only check box when the dialog box is closed.
  cdlOFNShareAware = &H4000 ' Specifies that sharing violation errors will be ignored.
End Enum
 
Public Type OPENFILENAME
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     lpstrFilter As String
     lpstrCustomFilter As String
     nMaxCustFilter As Long
     nFilterIndex As Long
     FileName As String
     nMaxFile As Long
     lpstrFileTitle As String
     nMaxFileTitle As Long
     lpstrInitialDir As String
     lpstrTitle As String
     flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     lpstrDefExt As String
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type


Public Function BackSlash(ByVal sPath As String) As String
  If Right(sPath, 1) = "\" Then
    BackSlash = sPath
  Else
    BackSlash = sPath & "\"
  End If
End Function

