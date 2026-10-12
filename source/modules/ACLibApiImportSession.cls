VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACLibApiImportSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: ACLibApiImportSession
'---------------------------------------------------------------------------------------
'
' Headless import session for the automation API (see modApi.API)
'
' Author:
'     Josef Poetzl
'
' Remarks:
'---------------------------------------------------------------------------------------
'| Uses the ACLibFileManager instance of the wizard (SilentMode during the session).
'| Public functions return the complete JSON result ({"ok":true,...}),
'| errors are raised to the caller (modApi converts them to JSON).
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const ERR_INVALID_ARGUMENT As Long = vbObjectError + 1001
Private Const ERR_FILE_NOT_FOUND As Long = vbObjectError + 1002

Private Const SOURCE_CONFIG As String = "config"
Private Const SOURCE_LOCAL As String = "local"
Private Const SOURCE_GITHUB As String = "github"

Private Const DEFAULT_GITHUB_OWNER As String = "AccessCodeLib"
Private Const DEFAULT_GITHUB_REPOSITORY As String = "AccessCodeLib"
Private Const DEFAULT_GITHUB_BRANCH As String = "master"

Private Const MODULNAME_CONFIG_APPLICATION As String = "_config_Application"
Private Const GITHUB_NOT_FOUND_CONTENT As String = "404: Not Found"

#Const EARLYBINDING = 0

Private WithEvents m_FileManager As ACLibFileManager
Attribute m_FileManager.VB_VarHelpID = -1

Private m_Source As String
Private m_ImportMode As CodeLibImportMode
Private m_ImportTests As Boolean
Private m_ImportExamples As Boolean

' Settings of the predeclared ACLibGitHubImporter before the session (restored in Class_Terminate)
Private m_GitHubSettingsSaved As Boolean
Private m_SavedRepositoryOwner As String
Private m_SavedRepositoryName As String
Private m_SavedBranchName As String
Private m_SavedAuthorizationToken As String

' Resolve/import state
Private m_PlannedModes As Collection      ' key: full file path, item: CodeLibImportMode
Private m_PlannedExecutes As Collection   ' key/item: execute string
Private m_ImportLog As Collection         ' items: JSON objects
Private m_MissingSelectedFiles As String  ' selected files (not dependencies) that are not available

Private Sub Class_Initialize()
   Set m_FileManager = CurrentACLibFileManager
   m_FileManager.ResetRepositoryRootDirectory
   m_FileManager.ClearMessages
   m_FileManager.SilentMode = True
   m_ImportMode = clim_ImportMissingItems
End Sub

Private Sub Class_Terminate()
On Error Resume Next ' cleanup only, must not raise from Terminate
   If Not m_FileManager Is Nothing Then
      m_FileManager.ResetRepositoryRootDirectory
      m_FileManager.SilentMode = False
      m_FileManager.ClearMessages
   End If
   RestoreGitHubImporterSettings
   Set m_FileManager = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Sub: Configure
'---------------------------------------------------------------------------------------
'
' Read options: "mode=missing|all|selected;tests=0|1;examples=0|1;source=config|local|github;repo=owner/name@branch"
'
'---------------------------------------------------------------------------------------
Public Sub Configure(ByVal Options As String)

   Dim OptionItems() As String
   Dim OptionItem As String
   Dim OptionName As String
   Dim OptionValue As String
   Dim SeparatorPos As Long
   Dim i As Long
   Dim Source As String
   Dim RepositoryString As String

   Source = SOURCE_CONFIG
   m_ImportMode = clim_ImportMissingItems
   m_ImportTests = False
   m_ImportExamples = False

   OptionItems = Split(Options, ";")
   For i = LBound(OptionItems) To UBound(OptionItems)
      OptionItem = Trim$(OptionItems(i))
      If Len(OptionItem) > 0 Then
         SeparatorPos = InStr(1, OptionItem, "=")
         If SeparatorPos = 0 Then
            Err.Raise ERR_INVALID_ARGUMENT, "ACLibApiImportSession.Configure", "Invalid option '" & OptionItem & "' (expected name=value)"
         End If
         OptionName = LCase$(Trim$(Left$(OptionItem, SeparatorPos - 1)))
         OptionValue = Trim$(Mid$(OptionItem, SeparatorPos + 1))
         Select Case OptionName
            Case "mode"
               m_ImportMode = ParseImportMode(OptionValue)
            Case "tests"
               m_ImportTests = ParseBooleanOption(OptionName, OptionValue)
            Case "examples"
               m_ImportExamples = ParseBooleanOption(OptionName, OptionValue)
            Case "source"
               Select Case OptionValue
                  Case SOURCE_CONFIG, SOURCE_LOCAL, SOURCE_GITHUB
                     Source = LCase$(OptionValue)
                  Case Else
                     Err.Raise ERR_INVALID_ARGUMENT, "ACLibApiImportSession.Configure", "Invalid value for source: '" & OptionValue & "' (config|local|github)"
               End Select
            Case "repo"
               RepositoryString = OptionValue
            Case Else
               Err.Raise ERR_INVALID_ARGUMENT, "ACLibApiImportSession.Configure", "Unknown option '" & OptionName & "'"
         End Select
      End If
   Next

   If Len(RepositoryString) > 0 Then
      If Source = SOURCE_LOCAL Then
         Err.Raise ERR_INVALID_ARGUMENT, "ACLibApiImportSession.Configure", "Option repo can not be combined with source=local"
      End If
      Source = SOURCE_GITHUB
   End If

   Select Case Source
      Case SOURCE_CONFIG
         If Len(CurrentACLibConfiguration.LocalRepositoryPath) > 0 Then
            Source = SOURCE_LOCAL
         Else
            Source = SOURCE_GITHUB
         End If
      Case SOURCE_LOCAL
         If Len(CurrentACLibConfiguration.LocalRepositoryPath) = 0 Then
            Err.Raise ERR_INVALID_ARGUMENT, "ACLibApiImportSession.Configure", "No local repository path configured"
         End If
   End Select

   m_Source = Source
   If m_Source = SOURCE_GITHUB Then
      ConfigureGitHubImporter RepositoryString
   End If

   m_FileManager.ResetRepositoryRootDirectory

End Sub

Private Function ParseImportMode(ByVal ModeString As String) As CodeLibImportMode
   Select Case ModeString
      Case "missing"
         ParseImportMode = clim_ImportMissingItems
      Case "all"
         ParseImportMode = clim_ImportAllUsedItems
      Case "selected"
         ParseImportMode = clim_ImportSelectedOnly
      Case Else
         Err.Raise ERR_INVALID_ARGUMENT, "ACLibApiImportSession.ParseImportMode", "Invalid value for mode: '" & ModeString & "' (missing|all|selected)"
   End Select
End Function

Private Function ParseBooleanOption(ByVal OptionName As String, ByVal OptionValue As String) As Boolean
   Select Case OptionValue
      Case "1", "true"
         ParseBooleanOption = True
      Case "0", "false"
         ParseBooleanOption = False
      Case Else
         Err.Raise ERR_INVALID_ARGUMENT, "ACLibApiImportSession.ParseBooleanOption", "Invalid value for " & OptionName & ": '" & OptionValue & "' (0|1)"
   End Select
End Function

Private Sub ConfigureGitHubImporter(ByVal RepositoryString As String)
' RepositoryString: owner/name[@branch]

   Dim RepositoryOwner As String
   Dim RepositoryName As String
   Dim BranchName As String
   Dim SeparatorPos As Long
   Dim AuthorizationToken As String

   RepositoryOwner = DEFAULT_GITHUB_OWNER
   RepositoryName = DEFAULT_GITHUB_REPOSITORY
   BranchName = DEFAULT_GITHUB_BRANCH

   If Len(RepositoryString) > 0 Then
      SeparatorPos = InStr(1, RepositoryString, "@")
      If SeparatorPos > 0 Then
         BranchName = Trim$(Mid$(RepositoryString, SeparatorPos + 1))
         RepositoryString = Left$(RepositoryString, SeparatorPos - 1)
      End If
      SeparatorPos = InStr(1, RepositoryString, "/")
      If SeparatorPos = 0 Then
         Err.Raise ERR_INVALID_ARGUMENT, "ACLibApiImportSession.ConfigureGitHubImporter", "Invalid value for repo: '" & RepositoryString & "' (owner/name@branch)"
      End If
      RepositoryOwner = Trim$(Left$(RepositoryString, SeparatorPos - 1))
      RepositoryName = Trim$(Mid$(RepositoryString, SeparatorPos + 1))
      If Len(RepositoryOwner) = 0 Or Len(RepositoryName) = 0 Or Len(BranchName) = 0 Then
         Err.Raise ERR_INVALID_ARGUMENT, "ACLibApiImportSession.ConfigureGitHubImporter", "Invalid value for repo: '" & RepositoryString & "' (owner/name@branch)"
      End If
   End If

   AuthorizationToken = CurrentACLibConfiguration.GitHubAuthPersonalAccessToken

   SaveGitHubImporterSettings
   With ACLibGitHubImporter
      .RepositoryOwner = RepositoryOwner
      .RepositoryName = RepositoryName
      .BranchName = BranchName
      If Len(AuthorizationToken) > 0 Then
         .GitHubApiAuthorizationToken = AuthorizationToken
      End If
   End With

End Sub

Private Sub SaveGitHubImporterSettings()
   If m_GitHubSettingsSaved Then Exit Sub
   With ACLibGitHubImporter
      m_SavedRepositoryOwner = .RepositoryOwner
      m_SavedRepositoryName = .RepositoryName
      m_SavedBranchName = .BranchName
      m_SavedAuthorizationToken = .GitHubApiAuthorizationToken
   End With
   m_GitHubSettingsSaved = True
End Sub

Private Sub RestoreGitHubImporterSettings()
   If Not m_GitHubSettingsSaved Then Exit Sub
   With ACLibGitHubImporter
      .RepositoryOwner = m_SavedRepositoryOwner
      .RepositoryName = m_SavedRepositoryName
      .BranchName = m_SavedBranchName
      .GitHubApiAuthorizationToken = m_SavedAuthorizationToken
   End With
   m_GitHubSettingsSaved = False
End Sub

Private Property Get GitHubTempRepositoryPath() As String
   GitHubTempRepositoryPath = FileTools.TempPath & "ACLibTempRepo"
End Property

Private Property Get RepositoryDisplayName() As String
   If m_Source = SOURCE_GITHUB Then
      With ACLibGitHubImporter
         RepositoryDisplayName = .RepositoryOwner & "/" & .RepositoryName & "@" & .BranchName
      End With
   Else
      RepositoryDisplayName = CurrentACLibConfiguration.LocalRepositoryPath
   End If
End Property

'---------------------------------------------------------------------------------------
' Events of ACLibFileManager
'---------------------------------------------------------------------------------------

Private Sub m_FileManager_PropertyMissingLocalRepositoryRootDirectory(ByRef NewValue As String)
' same behaviour as the wizard form: only answer in GitHub mode,
' otherwise ACLibFileManager uses the configured local repository
   If m_Source = SOURCE_GITHUB Then
      NewValue = GitHubTempRepositoryPath & "\"
   End If
End Sub

Private Sub m_FileManager_MissingLocalRepositoryFile(ByVal ACLibPath As String, ByVal FullFilePath As String)

   If m_Source <> SOURCE_GITHUB Then
      Exit Sub
   End If

   FileTools.CreateDirectoryIfMissing FileTools.PathFromFullFileName(FullFilePath)
   ACLibGitHubImporter.DownloadACLibFileFromWeb Replace(ACLibPath, "\", "/"), FullFilePath

   If IsGitHubNotFoundFile(FullFilePath) Then
      Kill FullFilePath
      Err.Raise ERR_FILE_NOT_FOUND, "ACLibApiImportSession.MissingLocalRepositoryFile", "File not found on GitHub: " & Replace(ACLibPath, "\", "/")
   End If

End Sub

Private Sub m_FileManager_ImportRepositoryFile(ByVal ObjectName As String, ByVal RepositoryFile As String, _
                                               ByVal ElementType As CodeLibElementType, ByRef Dependency As Variant, _
                                               ByRef ImportFile As Object, ByRef Cancel As Integer)

   Dim ItemMode As Variant
   Dim ObjectExists As Boolean
   Dim Action As String

   If m_ImportLog Is Nothing Then Exit Sub

   If Not TryGetCollectionItem(m_PlannedModes, ImportFile.Path, ItemMode) Then
      ItemMode = m_ImportMode
   End If
   Action = DetermineAction(ObjectName, ElementType, ItemMode, ObjectExists)

   m_ImportLog.Add "{" & _
         JsonPair("name", JsonString(ObjectName)) & "," & _
         JsonPair("file", JsonString(RepositoryFile)) & "," & _
         JsonPair("type", JsonString(ElementTypeName(ElementType))) & "," & _
         JsonPair("action", JsonString(Action)) & "}"

End Sub

Private Function IsGitHubNotFoundFile(ByVal FilePath As String) As Boolean

   Dim FileNumber As Integer
   Dim FileContent As String

   If Not FileTools.FileExists(FilePath) Then Exit Function
   If FileLen(FilePath) > 64 Then Exit Function

   FileNumber = FreeFile
   Open FilePath For Binary Access Read As #FileNumber
   FileContent = String$(LOF(FileNumber), 0)
   Get #FileNumber, , FileContent
   Close #FileNumber

   IsGitHubNotFoundFile = (Trim$(Replace(Replace(FileContent, vbCr, vbNullString), vbLf, vbNullString)) = GITHUB_NOT_FOUND_CONTENT)

End Function

'---------------------------------------------------------------------------------------
' Commands
'---------------------------------------------------------------------------------------

Public Function GetConfig() As String

   Dim ImporterJson As String

   With ACLibGitHubImporter
      ImporterJson = "{" & _
            JsonPair("owner", JsonString(.RepositoryOwner)) & "," & _
            JsonPair("repo", JsonString(.RepositoryName)) & "," & _
            JsonPair("branch", JsonString(.BranchName)) & "}"
   End With

   With CurrentACLibConfiguration
      GetConfig = OkResult( _
            JsonPair("source", JsonString(m_Source)) & "," & _
            JsonPair("repository", JsonString(RepositoryDisplayName)) & "," & _
            JsonPair("localRepositoryPath", JsonString(.LocalRepositoryPath)) & "," & _
            JsonPair("privateRepositoryPath", JsonString(.PrivateRepositoryPath)) & "," & _
            JsonPair("github", ImporterJson) & "," & _
            JsonPair("hasGitHubToken", JsonBool(Len(.GitHubAuthPersonalAccessToken) > 0)) & "," & _
            JsonPair("importTestsDefault", JsonBool(.ImportTestsDefaultValue)))
   End With

End Function

Public Function ListInstalled() As String

   Dim ModuleItems As Collection
   Dim cli As CodeLibInfo
#If EARLYBINDING = 1 Then
   Dim vbc As VBIDE.VBComponent
#Else
   Dim vbc As Object
#End If

   Set ModuleItems = New Collection

   For Each vbc In m_FileManager.CurrentVbProject.VBComponents
      cli = m_FileManager.GetCodeLibInfoFromCodeModule(vbc.Name)
      If Len(cli.RepositoryFile) > 0 Then
         ModuleItems.Add "{" & _
               JsonPair("name", JsonString(vbc.Name)) & "," & _
               JsonPair("type", JsonString(ComponentTypeName(vbc.Type, vbc.Name))) & "," & _
               JsonPair("file", JsonString(cli.RepositoryFile)) & "}"
      End If
   Next

   ListInstalled = OkResult( _
         JsonPair("database", JsonString(CurrentDb.Name)) & "," & _
         JsonPair("modules", JsonArray(ModuleItems)))

End Function

Public Function Resolve(ByVal Files As String) As String

   Dim FileItems As Collection
   Dim MissingItems As Collection

   ResolveFiles Files, FileItems, MissingItems

   Resolve = OkResult( _
         SessionInfoJson & "," & _
         JsonPair("files", JsonArray(FileItems)) & "," & _
         JsonPair("missing", JsonArray(MissingItems)) & "," & _
         JsonPair("executes", JsonArray(m_PlannedExecutes, True)))

End Function

Public Function Import(ByVal Files As String) As String

   Dim FileItems As Collection
   Dim MissingItems As Collection
   Dim PathArray() As String
   Dim ReferencesBefore As Collection
   Dim ReferencesAfter As Collection
   Dim MessageItems As Collection
   Dim MessageItem As Variant

   ResolveFiles Files, FileItems, MissingItems

   ' selected files must exist, missing dependencies are reported (wizard skips them)
   If Len(m_MissingSelectedFiles) > 0 Then
      Err.Raise ERR_FILE_NOT_FOUND, "ACLibApiImportSession.Import", "Selected file(s) not available: " & Mid$(m_MissingSelectedFiles, 3)
   End If

   PathArray = SplitFiles(Files)
   ' ACLibFileManager.ImportRepositoryFiles only reads up to UBound - 1 (array from wizard form)
   ReDim Preserve PathArray(UBound(PathArray) + 1)

   Set ReferencesBefore = GetReferenceItems
   Set m_ImportLog = New Collection

   m_FileManager.ImportRepositoryFiles PathArray, m_ImportMode, m_ImportTests, m_ImportExamples

   Set ReferencesAfter = GetReferenceItems

   Set MessageItems = New Collection
   For Each MessageItem In m_FileManager.Messages
      MessageItems.Add MessageItem
   Next

   Import = OkResult( _
         SessionInfoJson & "," & _
         JsonPair("files", JsonArray(m_ImportLog)) & "," & _
         JsonPair("missing", JsonArray(MissingItems)) & "," & _
         JsonPair("references", "{" & _
               JsonPair("added", JsonArray(CollectionDifference(ReferencesAfter, ReferencesBefore))) & "," & _
               JsonPair("removed", JsonArray(CollectionDifference(ReferencesBefore, ReferencesAfter))) & "}") & "," & _
         JsonPair("executes", JsonArray(m_PlannedExecutes, True)) & "," & _
         JsonPair("messages", JsonArray(MessageItems, True)))

End Function

Private Property Get SessionInfoJson() As String
   SessionInfoJson = _
         JsonPair("source", JsonString(m_Source)) & "," & _
         JsonPair("repository", JsonString(RepositoryDisplayName)) & "," & _
         JsonPair("mode", JsonString(ImportModeName(m_ImportMode))) & "," & _
         JsonPair("tests", JsonBool(m_ImportTests)) & "," & _
         JsonPair("examples", JsonBool(m_ImportExamples))
End Property

'---------------------------------------------------------------------------------------
' Resolve (read only, same rules as ACLibFileManager.ImportFile)
'---------------------------------------------------------------------------------------

Private Sub ResolveFiles(ByVal Files As String, ByRef FileItems As Collection, ByRef MissingItems As Collection)

   Dim SelectedFiles() As String
   Dim Queue As Collection
   Dim Records As Collection
   Dim ReplacedFiles As Collection
   Dim Record As Variant
   Dim Action As String
   Dim i As Long

   SelectedFiles = SplitFiles(Files)

   Set Queue = New Collection
   Set Records = New Collection
   Set ReplacedFiles = New Collection
   Set MissingItems = New Collection
   Set m_PlannedModes = New Collection
   Set m_PlannedExecutes = New Collection
   m_MissingSelectedFiles = vbNullString

   For i = LBound(SelectedFiles) To UBound(SelectedFiles)
      Queue.Add Array(SelectedFiles(i), vbNullString, m_ImportMode, vbNullString)
   Next

   Do While Queue.Count > 0
      ResolveQueueItem Queue(1), Queue, Records, ReplacedFiles, MissingItems
      Queue.Remove 1
   Loop

   ' Record: Array(FullPath, JsonBeforeAction, Action, JsonAfterAction)
   Set FileItems = New Collection
   For Each Record In Records
      Action = Record(2)
      If CollectionContains(ReplacedFiles, Record(0)) Then
         Action = "replaced"
      End If
      FileItems.Add Record(1) & JsonString(Action) & Record(3)
   Next

End Sub

Private Sub ResolveQueueItem(ByVal QueueItem As Variant, ByVal Queue As Collection, ByVal Records As Collection, _
                             ByVal ReplacedFiles As Collection, ByVal MissingItems As Collection)
' QueueItem: Array(RepositoryPath, DownloadBase, ImportMode, RequiredBy)

   Dim RepositoryPath As String
   Dim DownloadBase As String
   Dim ItemMode As CodeLibImportMode
   Dim RequiredBy As String
   Dim FullPath As String
   Dim ErrorMessage As String
   Dim cli As CodeLibInfo
   Dim ReplacedFullPath As String
   Dim ReplacedCli As CodeLibInfo
   Dim ObjectExists As Boolean
   Dim Action As String
   Dim i As Long

   RepositoryPath = QueueItem(0)
   DownloadBase = QueueItem(1)
   ItemMode = QueueItem(2)
   RequiredBy = QueueItem(3)

   If Not TryReadCodeLibInfo(RepositoryPath, DownloadBase, (ItemMode <> clim_ImportSelectedOnly), FullPath, cli, ErrorMessage) Then
      If Not CollectionContains(MissingItems, RepositoryPath) Then
         MissingItems.Add "{" & _
               JsonPair("file", JsonString(RepositoryPath)) & "," & _
               JsonPair("requiredBy", JsonStringOrNull(RequiredBy)) & "," & _
               JsonPair("message", JsonString(ErrorMessage)) & "}", RepositoryPath
      End If
      If Len(RequiredBy) = 0 Then
         m_MissingSelectedFiles = m_MissingSelectedFiles & ", " & RepositoryPath & " (" & ErrorMessage & ")"
      End If
      Exit Sub
   End If

   If CollectionContains(m_PlannedModes, FullPath) Then ' already resolved
      Exit Sub
   End If
   m_PlannedModes.Add ItemMode, FullPath

   If Len(cli.RepositoryFile) > 0 Then
      RequiredBy = cli.RepositoryFile
   Else
      RequiredBy = RepositoryPath
   End If

   ' <replace>
   If Len(cli.RepositoryFileReplacement) > 0 Then
      If TryReadCodeLibInfo(cli.RepositoryFileReplacement, vbNullString, False, ReplacedFullPath, ReplacedCli, ErrorMessage) Then
         If Not CollectionContains(ReplacedFiles, ReplacedFullPath) Then
            ReplacedFiles.Add ReplacedFullPath, ReplacedFullPath
         End If
      End If
   End If

   ' <license>
   If Len(cli.LicenseFile) > 0 Then
      Queue.Add Array(cli.LicenseFile, DownloadBase, clim_ImportMissingItems, RequiredBy)
   End If

   ' <use>
   If ItemMode <> clim_ImportSelectedOnly Then
      If (0 / 1) + (Not Not cli.Dependency) Then
         For i = 0 To UBound(cli.Dependency)
            Queue.Add Array(cli.Dependency(i), DownloadBase, ItemMode, RequiredBy)
         Next
      End If
   End If

   ' <test>
   If m_ImportTests Then
      If (0 / 1) + (Not Not cli.TestFiles) Then
         For i = 0 To UBound(cli.TestFiles)
            Queue.Add Array(cli.TestFiles(i), DownloadBase, ItemMode, RequiredBy)
         Next
      End If
   End If

   ' <example>
   If m_ImportExamples Then
      If (0 / 1) + (Not Not cli.ExampleFiles) Then
         For i = 0 To UBound(cli.ExampleFiles)
            Queue.Add Array(cli.ExampleFiles(i), DownloadBase, ItemMode, RequiredBy)
         Next
      End If
   End If

   ' <execute>
   If (0 / 1) + (Not Not cli.ExecuteList) Then
      For i = 0 To UBound(cli.ExecuteList)
         If Not CollectionContains(m_PlannedExecutes, cli.ExecuteList(i)) Then
            m_PlannedExecutes.Add cli.ExecuteList(i), cli.ExecuteList(i)
         End If
      Next
   End If

   Action = DetermineAction(cli.Name, cli.Type, ItemMode, ObjectExists)

   Records.Add Array(FullPath, _
         "{" & _
         JsonPair("name", JsonString(cli.Name)) & "," & _
         JsonPair("file", JsonString(cli.RepositoryFile)) & "," & _
         JsonPair("requestedPath", JsonString(RepositoryPath)) & "," & _
         JsonPair("type", JsonString(ElementTypeName(cli.Type))) & "," & _
         JsonPair("exists", JsonBool(ObjectExists)) & "," & _
         """action"":", _
         Action, _
         "," & _
         JsonPair("description", JsonString(cli.Description)) & "," & _
         JsonPair("license", JsonString(cli.LicenseFile)) & "," & _
         JsonPair("replaces", JsonString(cli.RepositoryFileReplacement)) & "," & _
         JsonPair("dependencies", JsonStringArray(cli.Dependency)) & "," & _
         JsonPair("references", ReferencesJson(cli.References)) & "," & _
         JsonPair("tests", JsonStringArray(cli.TestFiles)) & "," & _
         JsonPair("examples", JsonStringArray(cli.ExampleFiles)) & "," & _
         JsonPair("executes", JsonStringArray(cli.ExecuteList)) & "}")

End Sub

Private Function TryReadCodeLibInfo(ByVal RepositoryPath As String, ByRef DownloadBase As String, ByVal FindDependency As Boolean, _
                                    ByRef FullPath As String, ByRef CodeLibInf As CodeLibInfo, ByRef ErrorMessage As String) As Boolean
' Errors are not swallowed: the error message is returned and reported as missing file

On Error GoTo HandleErr

   FullPath = m_FileManager.GetRepositoryFullPath(RepositoryPath, DownloadBase)
   If Not FileTools.FileExists(FullPath) Then
      ErrorMessage = "File not found: " & FullPath
      Exit Function
   End If

   CodeLibInf = m_FileManager.GetCodeLibInfoFromFilePath(FullPath, FindDependency)
   TryReadCodeLibInfo = True

   Exit Function

HandleErr:
   ErrorMessage = Err.Description & " (" & Err.Number & ")"

End Function

Private Function DetermineAction(ByVal ObjectName As String, ByVal ElementType As CodeLibElementType, _
                                 ByVal ItemMode As CodeLibImportMode, ByRef ObjectExists As Boolean) As String
' same rules as ACLibFileManager.ImportVbComponent / ImportAccessObject

#If EARLYBINDING = 1 Then
   Dim vbc As VBIDE.VBComponent
#Else
   Dim vbc As Object
#End If

   ObjectExists = False

   Select Case ElementType
      Case CodeLibElementType.clet_StdModule, CodeLibElementType.clet_ClassModule
         Set vbc = FindVbComponent(ObjectName)
         If vbc Is Nothing Then
            DetermineAction = "import"
         Else
            ObjectExists = True
            If vbc.Type <> ElementType Then
               DetermineAction = "overwrite"
            ElseIf ItemMode = clim_ImportMissingItems Then
               DetermineAction = "skip"
            ElseIf vbc.Name = MODULNAME_CONFIG_APPLICATION Then
               DetermineAction = "skip"
            Else
               DetermineAction = "overwrite"
            End If
         End If

      Case CodeLibElementType.clet_Form, CodeLibElementType.clet_Report
         If ElementType = CodeLibElementType.clet_Form Then
            ObjectExists = AccessObjectExists(CurrentProject.AllForms, ObjectName)
         Else
            ObjectExists = AccessObjectExists(CurrentProject.AllReports, ObjectName)
         End If
         If Not ObjectExists Then
            DetermineAction = "import"
         ElseIf ItemMode = clim_ImportMissingItems Then
            DetermineAction = "skip"
         Else
            DetermineAction = "overwrite"
         End If

      Case Else ' clet_Package: codelib block only
         DetermineAction = "none"

   End Select

End Function

Private Function FindVbComponent(ByVal ComponentName As String) As Object

#If EARLYBINDING = 1 Then
   Dim vbc As VBIDE.VBComponent
#Else
   Dim vbc As Object
#End If

   For Each vbc In m_FileManager.CurrentVbProject.VBComponents
      If vbc.Name = ComponentName Then
         Set FindVbComponent = vbc
         Exit Function
      End If
   Next

End Function

Private Function AccessObjectExists(ByVal AccessObjects As Object, ByVal ObjectName As String) As Boolean

   Dim AcObj As AccessObject

   For Each AcObj In AccessObjects
      If AcObj.Name = ObjectName Then
         AccessObjectExists = True
         Exit Function
      End If
   Next

End Function

Private Function ReferencesJson(ByRef CodeLibRefs() As CodeLibInfoReference) As String

   Dim RefItems As Collection
   Dim i As Long

   Set RefItems = New Collection

   If (0 / 1) + (Not Not CodeLibRefs) Then
      For i = 0 To UBound(CodeLibRefs)
         RefItems.Add "{" & _
               JsonPair("name", JsonString(CodeLibRefs(i).Name)) & "," & _
               JsonPair("guid", JsonString(CodeLibRefs(i).GUID)) & "," & _
               JsonPair("major", CStr(CodeLibRefs(i).Major)) & "," & _
               JsonPair("minor", CStr(CodeLibRefs(i).Minor)) & "," & _
               JsonPair("status", JsonString(ReferenceStatus(CodeLibRefs(i)))) & "}"
      Next
   End If

   ReferencesJson = JsonArray(RefItems)

End Function

Private Function ReferenceStatus(ByRef CodeLibRef As CodeLibInfoReference) As String
' same rules as ACLibFileManager.AddMissingReference

   Dim ref As Access.Reference

   For Each ref In Access.References
      If ref.Name = CodeLibRef.Name Then
         If ref.Major < CodeLibRef.Major Then
            ReferenceStatus = "update"
         ElseIf ref.Major = CodeLibRef.Major And ref.Minor < CodeLibRef.Minor Then
            ReferenceStatus = "update"
         Else
            ReferenceStatus = "present"
         End If
         Exit Function
      End If
   Next

   ReferenceStatus = "add"

End Function

Private Function GetReferenceItems() As Collection
' key: JSON object of the reference

   Dim RefItems As Collection
   Dim ref As Access.Reference
   Dim RefJson As String

   Set RefItems = New Collection

   For Each ref In Access.References
      RefJson = "{" & _
            JsonPair("name", JsonString(ref.Name)) & "," & _
            JsonPair("guid", JsonString(ref.GUID)) & "," & _
            JsonPair("major", CStr(ref.Major)) & "," & _
            JsonPair("minor", CStr(ref.Minor)) & "}"
      If Not CollectionContains(RefItems, RefJson) Then
         RefItems.Add RefJson, RefJson
      End If
   Next

   Set GetReferenceItems = RefItems

End Function

'---------------------------------------------------------------------------------------
' Helper
'---------------------------------------------------------------------------------------

Private Function SplitFiles(ByVal Files As String) As String()

   Dim FileItems() As String
   Dim Result() As String
   Dim FileItem As String
   Dim ResultIdx As Long
   Dim i As Long

   FileItems = Split(Files, ";")
   ResultIdx = -1

   For i = LBound(FileItems) To UBound(FileItems)
      FileItem = Replace(Trim$(FileItems(i)), "\", "/")
      If Len(FileItem) > 0 Then
         ResultIdx = ResultIdx + 1
         ReDim Preserve Result(ResultIdx)
         Result(ResultIdx) = FileItem
      End If
   Next

   If ResultIdx < 0 Then
      Err.Raise ERR_INVALID_ARGUMENT, "ACLibApiImportSession.SplitFiles", "No files specified"
   End If

   SplitFiles = Result

End Function

Private Function CollectionDifference(ByVal SourceCollection As Collection, ByVal ExcludeCollection As Collection) As Collection
' items of SourceCollection which are not a key in ExcludeCollection

   Dim Result As Collection
   Dim Item As Variant

   Set Result = New Collection
   For Each Item In SourceCollection
      If Not CollectionContains(ExcludeCollection, Item) Then
         Result.Add Item
      End If
   Next

   Set CollectionDifference = Result

End Function

Private Function CollectionContains(ByVal col As Collection, ByVal Key As String) As Boolean
   Dim Item As Variant
   CollectionContains = TryGetCollectionItem(col, Key, Item)
End Function

Private Function TryGetCollectionItem(ByVal col As Collection, ByVal Key As String, ByRef Item As Variant) As Boolean
' Collection has no Exists method => check via error (not found is a valid result)

On Error GoTo HandleErr

   If col Is Nothing Then Exit Function
   Item = col.Item(Key)
   TryGetCollectionItem = True

   Exit Function

HandleErr:
   Err.Clear

End Function

Private Function ElementTypeName(ByVal ElementType As CodeLibElementType) As String
   Select Case ElementType
      Case CodeLibElementType.clet_StdModule
         ElementTypeName = "module"
      Case CodeLibElementType.clet_ClassModule
         ElementTypeName = "class"
      Case CodeLibElementType.clet_Form
         ElementTypeName = "form"
      Case CodeLibElementType.clet_Report
         ElementTypeName = "report"
      Case CodeLibElementType.clet_Package
         ElementTypeName = "package"
      Case Else
         ElementTypeName = "unknown"
   End Select
End Function

Private Function ComponentTypeName(ByVal ComponentType As Long, ByVal ComponentName As String) As String
   Select Case ComponentType
      Case vbext_ct_StdModule
         ComponentTypeName = "module"
      Case vbext_ct_ClassModule
         ComponentTypeName = "class"
      Case vbext_ct_Document
         If Left$(ComponentName, 5) = "Form_" Then
            ComponentTypeName = "form"
         ElseIf Left$(ComponentName, 7) = "Report_" Then
            ComponentTypeName = "report"
         Else
            ComponentTypeName = "document"
         End If
      Case Else
         ComponentTypeName = "unknown"
   End Select
End Function

Private Function ImportModeName(ByVal ImportMode As CodeLibImportMode) As String
   Select Case ImportMode
      Case clim_ImportMissingItems
         ImportModeName = "missing"
      Case clim_ImportAllUsedItems
         ImportModeName = "all"
      Case clim_ImportSelectedOnly
         ImportModeName = "selected"
   End Select
End Function

'---------------------------------------------------------------------------------------
' JSON writer
'---------------------------------------------------------------------------------------

Private Function OkResult(ByVal BodyJson As String) As String
   OkResult = "{""ok"":true," & BodyJson & "}"
End Function

Private Function JsonPair(ByVal Name As String, ByVal ValueJson As String) As String
   JsonPair = JsonString(Name) & ":" & ValueJson
End Function

Private Function JsonBool(ByVal Value As Boolean) As String
   If Value Then
      JsonBool = "true"
   Else
      JsonBool = "false"
   End If
End Function

Private Function JsonStringOrNull(ByVal Value As String) As String
   If Len(Value) = 0 Then
      JsonStringOrNull = "null"
   Else
      JsonStringOrNull = JsonString(Value)
   End If
End Function

Private Function JsonStringArray(ByRef Values() As String) As String

   Dim Result As String
   Dim i As Long

   If (0 / 1) + (Not Not Values) Then
      For i = LBound(Values) To UBound(Values)
         Result = Result & "," & JsonString(Values(i))
      Next
      Result = Mid$(Result, 2)
   End If

   JsonStringArray = "[" & Result & "]"

End Function

Private Function JsonArray(ByVal Items As Collection, Optional ByVal QuoteItems As Boolean = False) As String
' Items: JSON values (QuoteItems = False) or plain strings (QuoteItems = True)

   Dim Result As String
   Dim Item As Variant

   If Not Items Is Nothing Then
      For Each Item In Items
         If QuoteItems Then
            Result = Result & "," & JsonString(Item)
         Else
            Result = Result & "," & Item
         End If
      Next
      Result = Mid$(Result, 2)
   End If

   JsonArray = "[" & Result & "]"

End Function

Private Function JsonString(ByVal Value As String) As String

   Dim Result As String
   Dim CharCode As Long
   Dim i As Long

   For i = 1 To Len(Value)
      CharCode = AscW(Mid$(Value, i, 1)) And &HFFFF&
      Select Case CharCode
         Case 34 ' "
            Result = Result & "\"""
         Case 92 ' \
            Result = Result & "\\"
         Case 8
            Result = Result & "\b"
         Case 9
            Result = Result & "\t"
         Case 10
            Result = Result & "\n"
         Case 12
            Result = Result & "\f"
         Case 13
            Result = Result & "\r"
         Case Is < 32
            Result = Result & "\u" & Right$("000" & Hex$(CharCode), 4)
         Case Else
            Result = Result & Mid$(Value, i, 1)
      End Select
   Next

   JsonString = """" & Result & """"

End Function
