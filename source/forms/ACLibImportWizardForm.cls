Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: ACLibImportWizardForm
'---------------------------------------------------------------------------------------
'
' Wizard form for import of CodeLib elements
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

' verwendete Erweiterungen
Private Const EXTENSION_KEY_APPFILE As String = "AppFile"
Private Const APPFILE_PROPNAME_APPICON As String = "AppIcon"
Private Const RepositorySource_GitHub As Long = 1
Private Const RepositorySource_LocalRepository As Long = 2

Private Const TEMPDB_TABNAME As String = "tRepositoryFiles"
Private Const TEMPDB_TABDDL As String = "create table " & TEMPDB_TABNAME & " (LocalRepositoryPath varchar(255) primary key, ObjectName varchar(255), Description memo)"
Private m_TempDb As TempDbHandler

Private m_LastSelectionID As Variant

Private WithEvents m_L10nDict As L10nDict
Attribute m_L10nDict.VB_VarHelpID = -1
Private WithEvents m_RepositoryTree As Form_ACLibRepositoryTreeForm
Attribute m_RepositoryTree.VB_VarHelpID = -1
Private WithEvents m_ACLibFileManager As ACLibFileManager
Attribute m_ACLibFileManager.VB_VarHelpID = -1

Private Sub cbxLangCode_AfterUpdate()
   With Me.cbxLangCode
      If IsNull(.Value) Then
         .Value = "EN"
      End If
      L10n.LangCode = .Value
   End With
End Sub

Private Sub Form_Open(Cancel As Integer)
   Me.cbxLangCode.Value = L10n.LangCode
   L10n.TranslateControls Me.Controls
   Set m_L10nDict = L10n
   SetFormCaption
End Sub

Private Sub m_ACLibFileManager_MissingLocalRepositoryFile(ByVal ACLibPath As String, ByVal FullFilePath As String)
   With ACLibGitHubImporter
      CreateDirectoryIfMissing FileTools.PathFromFullFileName(FullFilePath)
      .DownloadACLibFileFromWeb ACLibPath, FullFilePath
   End With
End Sub

Private Sub m_ACLibFileManager_PropertyMissingLocalRepositoryRootDirectory(NewValue As String)
   NewValue = CurrentLocalRepositoryPath & "\"
End Sub

Private Sub m_L10nDict_LanguageChanged()
   If Nz(Me.cbxLangCode.Value, vbNullString) <> L10n.LangCode Then
      Me.cbxLangCode.Value = L10n.LangCode
   End If
   SetFormCaption
   L10n.TranslateControls Me.Controls
End Sub

Public Sub SetFormCaption()

   Dim FormCaption As String

   FormCaption = "Access Code Library - Import Wizard  " & ChrW(&H25AA) & "  (Version " & CurrentApplication.Version & ")"

   Me.Caption = FormCaption

End Sub

Private Sub BindTextbox(ByRef TextBoxRef As Textbox, Optional ByVal BaseFolderPath As String = vbNullString)

   'Late binding so that ApplicationHandler_DirTextbox class does not have to exist.
   Dim DirTextbox As Object ' = ApplicationHandler_DirTextbox

On Error GoTo HandleErr

   'Use default instance:
   Set DirTextbox = CurrentApplication.Extensions("DirTextbox")

   'bind textbox
   If Not (DirTextbox Is Nothing) Then
      Set DirTextbox.Textbox = TextBoxRef
      DirTextbox.BaseFolderPath = BaseFolderPath
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "BindTextbox", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub chkImportTests_AfterUpdate()
   CurrentACLibConfiguration.ImportTestsDefaultValue = Nz(Me.chkImportTests.Value, False)
End Sub

Private Sub cmdAddFile_Click()

On Error GoTo HandleErr

   AddFileFromFileName Me.txtFileString & ""

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdAddFile_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub cmdClose_Click()
   DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdImportFiles_Click()

   Dim FileNameArray() As String
   Dim ArraySize As Long, i As Long
   Dim lb As ListBox
   Dim ACLibFileMngr As ACLibFileManager

On Error GoTo HandleErr

   Set lb = Me.lstImportFiles

   ArraySize = lb.ListCount

   If ArraySize <= 0 Then
      MsgBox "No files selected", vbInformation
      Exit Sub
   End If

   ReDim FileNameArray(ArraySize)

   For i = 0 To ArraySize - 1
      FileNameArray(i) = lb.ItemData(i)
   Next

   Me.labInfo.Caption = L10n.Text("Import process running ...")
   Me.labInfo.Visible = True
   Me.Repaint

   Set ACLibFileMngr = CurrentACLibFileManager
   If Me.ogRepositorySource = RepositorySource_GitHub Then
      Set m_ACLibFileManager = ACLibFileMngr
   Else '
      Set m_ACLibFileManager = Nothing
   End If
   ACLibFileMngr.ImportRepositoryFiles FileNameArray, Nz(Me.ogImportMode.Value, 0), Nz(Me.chkImportTests.Value, False)

   Me.labInfo.Caption = L10n.Text("Files have been imported")
   Me.Repaint

   TempDb.Execute "delete from " & TEMPDB_TABNAME
   lb.Requery

   Me.SetFocus

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdImportFiles_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub cmdImportFiles_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)

On Error GoTo HandleErr

   If Button = 2 Then
      OpenImportFileShortcutMenu
      Button = 0
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdImportFiles_MouseDown", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Function OpenImportFileShortcutMenu() As Long

   Dim mnu As WinApiShortcutMenu
On Error GoTo HandleErr

   Set mnu = New WinApiShortcutMenu

   With mnu
      Set .MenuControl = Me.cmdOpenMenu
      Set .AccessForm = Me
      .ControlSection = acDetail

      .AddMenuItem 21, L10n.Text("Update all existing objects")
      .AddMenuItem 22, L10n.Text("Update all existing CodeModules")

      If Me.ogRepositorySource.Value = RepositorySource_LocalRepository Then

         .AddMenuItem -2, "", MF_SEPARATOR

         .AddMenuItem 31, L10n.Text("Export all existing objects")
         .AddMenuItem 32, L10n.Text("Export all existing CodeModules")
         .AddMenuItem 41, L10n.Text("Export all existing objects to the application folder")

      End If

   End With

   With CurrentACLibFileManager
      Select Case mnu.OpenMenu
         Case 21
            .RefreshAll clim_ImportAllUsedItems, True
         Case 22
            .RefreshAllModules clim_ImportAllUsedItems, True
         Case 31
            .ExportAll
         Case 32
            .ExportAllModules
         Case 41
            .ExportAllToApplicationSourceFolder = True
            .ExportAll
            .ExportAllToApplicationSourceFolder = False
         Case Else
            '
      End Select
   End With

   Set mnu = Nothing

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ShowImportFileShortcutMenu", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Private Sub cmdOpenMenu_Click()

On Error GoTo HandleErr

   OpenImportFileShortcutMenu

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdOpenMenu_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Sub

Private Sub cmdSelectFile_Click()

On Error GoTo HandleErr

   If Me.ogRepositorySource.Value = 2 Then
      SelectFileFromLocalRepository
   Else
      SelectFileFromGitHubRepositoryApi
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdSelectFile_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub SelectFileFromGitHubRepositoryApi()
   Set m_RepositoryTree = New Form_ACLibRepositoryTreeForm
   m_RepositoryTree.Visible = True
End Sub

Private Sub m_RepositoryTree_FilesSelected(ByRef FileArray() As String)
   Set m_RepositoryTree = Nothing
   DownLoadFiles FileArray
   AddFiles FileArray
End Sub

Private Sub DownLoadFiles(ByRef FileArray() As String)

   Dim RepoPath As String
   Dim i As Long
   Dim TargetFilePath As String

   RepoPath = CurrentLocalRepositoryPath & "\"

   With ACLibGitHubImporter
      For i = LBound(FileArray) To UBound(FileArray)
         TargetFilePath = RepoPath & Replace(FileArray(i), "/", "\")
         CreateDirectoryIfMissing FileTools.PathFromFullFileName(TargetFilePath)
         .DownloadACLibFileFromWeb FileArray(i), TargetFilePath
         FileArray(i) = TargetFilePath
      Next

   End With

End Sub

Private Function DownLoadFile(ByRef ACLibPath As String) As String

   Dim RepoPath As String
   Dim TargetFilePath As String

   RepoPath = CurrentLocalRepositoryPath & "\"

   With ACLibGitHubImporter
      TargetFilePath = RepoPath & Replace(ACLibPath, "/", "\")
      CreateDirectoryIfMissing FileTools.PathFromFullFileName(TargetFilePath)
      .DownloadACLibFileFromWeb ACLibPath, TargetFilePath
   End With

   DownLoadFile = TargetFilePath

End Function

Private Sub SelectFileFromLocalRepository()

Static LastStartFolder As String

   Dim StartFolder As String
   Dim SelectedFiles As String
   Dim FileArray() As String
   Dim Pos As Long

On Error GoTo HandleErr

   StartFolder = Replace(GetDirFromFullFileName(Me.txtFileString.Value & ""), "/", "\")
   If Len(StartFolder) = 0 Then
      StartFolder = LastStartFolder
   End If

   If Len(StartFolder) > 0 Then
      Do While Left$(StartFolder, 1) = "\"
         StartFolder = Mid$(StartFolder, 1)
         If Len(StartFolder) = 0 Then Exit Do
      Loop
   End If

   If Mid(StartFolder, 2, 1) <> ":" Then
      StartFolder = CurrentLocalRepositoryPath & StartFolder
   End If
   Do While Not DirExists(StartFolder)
      Pos = InStrRev(StartFolder, "\")
      If Pos = 0 Then Exit Do
      StartFolder = Left$(StartFolder, Pos - 1)
   Loop

   Me.sysFirst.SetFocus
   Me.cmdSelectFile.SetFocus

   SelectedFiles = SelectFile(StartFolder, , , True)
   If Len(SelectedFiles) = 0 Then
      Exit Sub
   End If

   FileArray = Split(SelectedFiles, "|")
   AddFiles FileArray

   LastStartFolder = Replace(GetDirFromFullFileName(FileArray(0)), CurrentLocalRepositoryPath, vbNullString)

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "SelectFileFromLocalRepository", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub


Private Sub AddFiles(ByRef FileArray() As String)

   Dim lb As ListBox
   Dim i As Long
   Dim ArraySize As Long

   Dim cli As CodeLibInfo

On Error GoTo HandleErr

   ArraySize = UBound(FileArray)

   Set lb = Me.lstImportFiles
   For i = 0 To ArraySize
      cli = CurrentACLibFileManager.GetCodeLibInfoFromFilePath(FileArray(i))
      TempDb.Execute "insert into " & TEMPDB_TABNAME & " (ObjectName, LocalRepositoryPath, Description) VALUES (" & _
                           SqlTools.TextToSqlText(cli.Name) & ", " & SqlTools.TextToSqlText(GetLocalRepositoryPath(FileArray(i))) & _
                           ", " & SqlTools.TextToSqlText(cli.Description) & ")", dbFailOnError
   Next

   lb.Requery
   Me.labInfo.Visible = (lb.ListCount = 0)

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdSelectFile_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Function GetLocalRepositoryPath(ByRef FullPath As String) As String

On Error GoTo HandleErr

   GetLocalRepositoryPath = Replace(GetRelativPathFromFullPath(Replace(FullPath, "/", "\"), CurrentLocalRepositoryPath, False), "\", "/")

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "getLocalRepositoryPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Private Property Get CurrentLocalRepositoryPath() As String

On Error GoTo HandleErr

   If Me.ogRepositorySource.Value = RepositorySource_LocalRepository Then
      CurrentLocalRepositoryPath = Me.txtLocalRepositoryPath.Value
   Else
      CurrentLocalRepositoryPath = GitHubTempRepositoryPath
   End If

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CurrentLocalRepositoryPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

Private Property Get GitHubTempRepositoryPath() As String

   'Dim TempRoot As String

   'TempRoot = FileTools.TempPath & "ACLibTempRepo"
   'FileTools.CreateDirectoryIfMissing TempRoot

   GitHubTempRepositoryPath = FileTools.TempPath & "ACLibTempRepo"

End Property

Private Sub RemoveTempFiles()

   If FileTools.DirExists(GitHubTempRepositoryPath) Then
      CreateObject("Scripting.FileSystemObject").DeleteFolder GitHubTempRepositoryPath, True
   End If

End Sub

Private Sub cmdSelectLocalRepository_Click()

   Dim SelectedRepositoryPath As String

On Error GoTo HandleErr

   SelectedRepositoryPath = SelectFolder(Nz(Me.txtLocalRepositoryPath, vbNullString), "Lokalen Repository-Ordner auswählen", , False, 1)

   If Len(SelectedRepositoryPath) > 0 Then
      If Right$(SelectedRepositoryPath, 1) = "\" Then
         SelectedRepositoryPath = Left$(SelectedRepositoryPath, Len(SelectedRepositoryPath) - 1)
      End If

      SetLocalRepositoryPath SelectedRepositoryPath

   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdSelectLocalRepository_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub SetEnableMode()

On Error GoTo HandleErr

   Me.cmdImportFiles.Enabled = Len(Me.txtLocalRepositoryPath.Value & vbNullString) > 0

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "setEnableMode", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub Form_Load()

On Error GoTo HandleErr

   With CurrentApplication
      Me.Caption = .ApplicationTitle & "  " & ChrW(&H25AA) & "  Version " & .Version
   End With

   LoadIconFromAppFiles

   With CurrentACLibConfiguration
      Me.txtLocalRepositoryPath.Value = .LocalRepositoryPath
      Me.txtGitHubAuthPersonalAccessToken.Value = .GitHubAuthPersonalAccessToken
      Me.chkImportTests.Value = .ImportTestsDefaultValue
   End With

   ConfigReproSourceMode 1 + Abs(Len(Me.txtLocalRepositoryPath.Value & vbNullString) > 0)
   EnableCodeModuleDescription Me.tbViewCodeModuleDescription.Value

   SetEnableMode

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "Form_Load", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub ConfigReproSourceMode(Optional NewMode As Long = 0)

   Dim UseLocalRepo As Boolean

   With Me.ogRepositorySource
      If NewMode > 0 Then
         .Value = NewMode
      End If
      UseLocalRepo = (.Value = RepositorySource_LocalRepository)
   End With

   Me.txtLocalRepositoryPath.Visible = UseLocalRepo
   Me.cmdSelectLocalRepository.Visible = UseLocalRepo

   Me.txtGitHubAuthPersonalAccessToken.Visible = Not UseLocalRepo

   Me.labRepositoryData.Visible = Not UseLocalRepo
   Me.txtRepoOwner.Visible = Not UseLocalRepo
   Me.txtRepoName.Visible = Not UseLocalRepo
   Me.txtRepoBranch.Visible = Not UseLocalRepo

   SetACLibGitHubImporterConfig

End Sub

Private Sub SetACLibGitHubImporterConfig()
   If Me.ogRepositorySource = RepositorySource_GitHub Then
   If Len(Me.txtGitHubAuthPersonalAccessToken.Value) > 0 Then
      ACLibGitHubImporter.GitHubApiAuthorizationToken = Me.txtGitHubAuthPersonalAccessToken.Value
   End If
   End If
End Sub

Private Sub Form_Unload(ByRef Cancel As Integer)
On Error Resume Next
   If Not (m_TempDb Is Nothing) Then
      Me.lstImportFiles.RowSource = vbNullString
      DBEngine.Idle dbRefreshCache
      m_TempDb.Dispose
   End If
   RemoveTempFiles
   DisposeCurrentApplicationHandler
End Sub

Private Sub lstImportFiles_AfterUpdate()
   RefreshCodeModuleDescription
End Sub

Private Sub lstImportFiles_DblClick(ByRef Cancel As Integer)
   OpenSelectItemFormImportFilesListboxInTextViewer
End Sub

Private Sub lstImportFiles_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)

On Error GoTo HandleErr

   If KeyCode = vbKeyDelete Then
      RemoveSelectedItemsFromListbox
   ElseIf KeyCode = vbKeyF2 Then
      OpenSelectItemFormImportFilesListboxInTextViewer
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "lstImportFiles_KeyDown", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub RemoveSelectedItemsFromListbox()

   Dim lb As ListBox
   Dim SelectedItem As Variant
   Dim ItemFilter As String

On Error GoTo HandleErr

   Set lb = Me.lstImportFiles

   For Each SelectedItem In lb.ItemsSelected
      ItemFilter = ", " & SqlTools.TextToSqlText(lb.Column(1, SelectedItem))
   Next

   If Len(ItemFilter) <= 2 Then
      Exit Sub
   End If

   ItemFilter = Mid$(ItemFilter, 3)
   TempDb.Execute "delete from " & TEMPDB_TABNAME & " where LocalRepositoryPath IN (" & ItemFilter & ")"

   lb.Requery

   RefreshCodeModuleDescription

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "removeSelectedItemsFormListbox", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub lstImportFiles_LostFocus()
On Error Resume Next
   m_LastSelectionID = Me.lstImportFiles.Column(1)
   Me.lstImportFiles = Null
End Sub

Private Sub m_RepositoryTree_Unload()
   Set m_RepositoryTree = Nothing
End Sub

Private Sub ogRepositorySource_AfterUpdate()
   ConfigReproSourceMode
End Sub

Private Sub tbViewCodeModuleDescription_AfterUpdate()
   EnableCodeModuleDescription Me.tbViewCodeModuleDescription.Value
   If Len(m_LastSelectionID) > 0 Then
      SelectListItem m_LastSelectionID
   End If
End Sub

Private Sub SelectListItem(ByVal ItemID As String)
   Dim i As Long
   Dim lb As ListBox
   Set lb = Me.lstImportFiles
   For i = 0 To (lb.ListCount - 1)
      If lb.Column(1, i) = ItemID Then
         lb.SetFocus
         lb.Selected(i) = True
         RefreshCodeModuleDescriptionFromId ItemID, lb.Column(0, i)
         Exit Sub
      End If
   Next
End Sub

Private Sub txtFileString_GotFocus()
On Error Resume Next
   If Me.ogRepositorySource = RepositorySource_LocalRepository Then
      BindTextbox Me.txtFileString, CurrentLocalRepositoryPath
   End If
End Sub

Private Sub txtFileString_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)

On Error GoTo HandleErr

   If KeyCode = vbKeyReturn Then
      If Me.txtFileString.Text = ".." Then
         Exit Sub
      ElseIf Replace(Right$(Me.txtFileString.Text, 3), "/", "\") = "\.." Then
         Exit Sub
      Else
         AddFileFromFileName Me.txtFileString.Text
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyF2 Then
      OpenRepositoryFileInTextViewer Me.txtFileString.Text
      KeyCode = 0
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "txtFileString_KeyDown", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub AddFileFromFileName(ByVal ACLibFileString As String)

   Dim FileString As String
   Dim FileArray(0) As String

On Error GoTo HandleErr

   If Me.ogRepositorySource = RepositorySource_GitHub Then

      FileArray(0) = ACLibFileString
      DownLoadFiles FileArray
      AddFiles FileArray

   Else

      FileString = Trim$(Replace(ACLibFileString, "/", "\"))
      If Len(FileString) > 0 Then
         Do While Left$(FileString, 1) = "\"
            FileString = Trim$(Mid$(FileString, 2))
            If Len(FileString) = 0 Then Exit Sub
         Loop
         AddFile CurrentLocalRepositoryPath & FileString
      End If

   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "addFileFromFileName", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub


Private Function AddFile(ByRef NewFileName As String) As Boolean

   Dim FileArray(0) As String

On Error GoTo HandleErr

   If Len(NewFileName) = 0 Then
      Exit Function
   End If

   NewFileName = Replace(NewFileName, "\\", "\")

   If Not FileExists(NewFileName) Then
      MsgBox "This file does not exist", vbInformation
      AddFile = False
      Exit Function
   End If

   FileArray(0) = NewFileName
   AddFiles FileArray

   AddFile = True

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "addFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Private Sub txtGitHubAuthPersonalAccessToken_AfterUpdate()

On Error GoTo HandleErr

   SetGitHubAuthPersonalAccessToken Me.txtGitHubAuthPersonalAccessToken & vbNullString

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "txtLocalRepositoryPath_AfterUpdate", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub SetGitHubAuthPersonalAccessToken(ByVal NewAuthToken As String)

On Error GoTo HandleErr

   CurrentACLibConfiguration.GitHubAuthPersonalAccessToken = NewAuthToken

   'so that possible modifications from CurrentACLibConfiguration are taken over:
   Me.txtGitHubAuthPersonalAccessToken.Value = CurrentACLibConfiguration.GitHubAuthPersonalAccessToken
   ACLibGitHubImporter.GitHubApiAuthorizationToken = Me.txtGitHubAuthPersonalAccessToken.Value

   SetEnableMode

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "SetGitHubAuthPersonalAccessToken", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub txtLocalRepositoryPath_AfterUpdate()

On Error GoTo HandleErr

   SetLocalRepositoryPath Me.txtLocalRepositoryPath & vbNullString

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "txtLocalRepositoryPath_AfterUpdate", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub SetLocalRepositoryPath(ByVal NewRoot As String)

On Error GoTo HandleErr

   CurrentACLibConfiguration.LocalRepositoryPath = NewRoot

   'damit mögliche Modifikationen aus CurrentACLibConfiguration übernommen werden:
   Me.txtLocalRepositoryPath.Value = CurrentACLibConfiguration.LocalRepositoryPath

   SetEnableMode

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "SetLocalRepositoryPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub txtLocalRepositoryPath_BeforeUpdate(ByRef Cancel As Integer)

   Dim NewPath As String
On Error GoTo HandleErr

   NewPath = Me.txtLocalRepositoryPath & ""

   If Len(NewPath) > 0 Then
      Cancel = Not DirExists(NewPath)
      If Cancel Then
         MsgBox "Verzeichnis ist nicht vorhanden", vbInformation
      End If
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "txtLocalRepositoryPath_BeforeUpdate", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub LoadIconFromAppFiles()

   Dim IconFilePath As String
   Dim IconFileName As String

   'Latebindung, damit ApplicationHandler_AppFile-Klasse nicht vorhanden sein muss
   Dim AppFile As Object ' ... ApplicationHandler_AppFile

On Error GoTo HandleErr

   If Val(SysCmd(acSysCmdAccessVer)) <= 9 Then 'Abbruch, da Ac00 sonst abstürzt
      Exit Sub
   End If

   Set AppFile = CurrentApplication.Extensions(EXTENSION_KEY_APPFILE)

   'Textbox binden
   If Not (AppFile Is Nothing) Then
      IconFileName = ACLibIconFileName
      IconFilePath = CurrentACLibConfiguration.ACLibConfigDirectory

      If Len(ACLibIconFileName) = 0 Then 'nur Temp-Datei erzeugen
         IconFileName = Me.Name & ".ico"
         IconFilePath = TempPath
      End If

      IconFilePath = IconFilePath & IconFileName

      If Len(Dir$(IconFilePath)) = 0 Then
         If Not AppFile.CreateAppFile(APPFILE_PROPNAME_APPICON, IconFilePath) Then
            Exit Sub
         End If
      End If

      WinAPI.Image.SetFormIconFromFile Me, IconFilePath

   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "loadIconFromAppFiles", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Property Get TempDb() As TempDbHandler

On Error GoTo HandleErr

   If m_TempDb Is Nothing Then
      Set m_TempDb = New TempDbHandler
      m_TempDb.CreateDatabase True, True
      m_TempDb.CreateTable TEMPDB_TABNAME, TEMPDB_TABDDL
      Me.lstImportFiles.RowSource = "select ObjectName, LocalRepositoryPath FROM [" & m_TempDb.CurrentDatabase.Name & "]." & TEMPDB_TABNAME
   End If
   Set TempDb = m_TempDb

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "TempDb", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

Private Sub OpenSelectItemFormImportFilesListboxInTextViewer()
   OpenRepositoryFileInTextViewer Me.lstImportFiles.Column(1)
End Sub

Private Sub OpenRepositoryFileInTextViewer(ByVal ACLibPath As String)

   Dim FullPath As String

   If Me.ogRepositorySource = RepositorySource_GitHub Then
      FullPath = DownLoadFile(ACLibPath)
   Else
      FullPath = CurrentACLibFileManager.GetRepositoryFullPath(ACLibPath)
   End If

   WinAPI.Shell.Execute FullPath, "open"

End Sub

Private Sub EnableCodeModuleDescription(ByVal ViewDescription As Boolean)

   With Me.lstImportFiles
      If ViewDescription Then
         .Width = Me.lblImportMode.Left + Me.lblImportMode.Width - .Left
         RefreshCodeModuleDescription
      Else
         .Width = Me.ogImportMode.Width
      End If
   End With

   Me.txtCodeModuleDescription.Visible = ViewDescription
   Me.txtCodeModuleName.Visible = ViewDescription

End Sub

Private Sub RefreshCodeModuleDescription()
   RefreshCodeModuleDescriptionFromId Nz(Me.lstImportFiles.Column(1), vbNullString), Nz(Me.lstImportFiles.Column(0), vbNullString)
End Sub

Private Sub RefreshCodeModuleDescriptionFromId(ByVal LocalRepositoryPath As String, ByVal ModuleName As String)

   Dim Description As String
   If Len(LocalRepositoryPath) > 0 Then
      Description = Nz(m_TempDb.LookupSql("select Description from " & TEMPDB_TABNAME & " where LocalRepositoryPath = " & SqlTools.TextToSqlText(LocalRepositoryPath)), vbNullString)
   End If
   Me.txtCodeModuleName.Value = ModuleName
   Me.txtCodeModuleDescription.Value = Description

End Sub

Private Sub txtRepoBranch_AfterUpdate()
   ACLibGitHubImporter.BranchName = Nz(Me.txtRepoBranch.Value, vbNullString)
End Sub

Private Sub txtRepoName_AfterUpdate()
   ACLibGitHubImporter.RepositoryName = Nz(Me.txtRepoName.Value, vbNullString)
End Sub

Private Sub txtRepoOwner_AfterUpdate()

   Dim OwnerNameSplitPos As String

   With Me.txtRepoOwner

      OwnerNameSplitPos = InStr(1, .Value, "/")
      If OwnerNameSplitPos > 0 Then
         Me.txtRepoName.Value = Mid(.Value, OwnerNameSplitPos + 1)
         Call txtRepoName_AfterUpdate
         .Value = Left(.Value, OwnerNameSplitPos - 1)
      End If

      ACLibGitHubImporter.RepositoryOwner = Nz(.Value, vbNullString)

   End With

End Sub
