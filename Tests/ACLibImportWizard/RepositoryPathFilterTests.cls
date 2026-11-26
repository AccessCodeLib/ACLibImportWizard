VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryPathFilterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

'TestSuite.AddByClassName("RepositoryPathFilterTests").Run

Private Const DefaultExtensions As String = "bas;cls;frm;rep;package"

Private m_Filter As RepositoryPathFilter

Public Sub Setup()
   Set m_Filter = New RepositoryPathFilter
End Sub

Public Sub Teardown()
   Set m_Filter = Nothing
End Sub

'--- extension list

'AccUnit:Row("", "").Name("empty = no additions")
'AccUnit:Row("   ", "").Name("blanks = no additions")
'AccUnit:Row(" ;, ; ", "").Name("only separators = no additions")
'AccUnit:Row(".txt;md", ";txt;md").Name("with and without dot, semicolon")
'AccUnit:Row("txt, md ", ";txt;md").Name("comma and blanks")
'AccUnit:Row("txt,md;json", ";txt;md;json").Name("mixed separators")
'AccUnit:Row(" .TXT ; Md", ";txt;md").Name("case and blanks")
'AccUnit:Row("txt;;md;", ";txt;md").Name("empty parts")
'AccUnit:Row("txt;TXT", ";txt").Name("duplicates in additions")
'AccUnit:Row(".CLS;bas", "").Name("duplicates of the default list")
'AccUnit:Row("md;.Package;txt", ";md;txt").Name("default entry between additions")
Public Sub Initialize_AdditionalExtensions_AddsToDefaultList(ByVal Additions As String, ByVal ExpectedAdded As String)

   m_Filter.Initialize Additions

   Assert.That m_Filter.ExtensionList, Iz.EqualTo(DefaultExtensions & ExpectedAdded)

End Sub

Public Sub Initialize_CalledTwice_DoesNotKeepEarlierAdditions()

   m_Filter.Initialize "md"
   m_Filter.Initialize "txt"

   Assert.That m_Filter.ExtensionList, Iz.EqualTo(DefaultExtensions & ";txt")

End Sub
Public Sub ExtensionList_NewInstance_IsDefaultList()
   Assert.That m_Filter.ExtensionList, Iz.EqualTo(DefaultExtensions)
End Sub

'--- single files

'AccUnit:Row("data/SqlTools.cls", True).Name("cls")
'AccUnit:Row("data/modTools.bas", True).Name("bas")
'AccUnit:Row("forms/MyForm.frm", True).Name("frm")
'AccUnit:Row("reports/MyReport.rep", True).Name("rep")
'AccUnit:Row("_packages/Access-Add-In-Template.package", True).Name("package")
'AccUnit:Row("data/SQLTOOLS.CLS", True).Name("upper case extension")
'AccUnit:Row("data/Tools.Bas", True).Name("mixed case extension")
'AccUnit:Row("a.b/c.cls", True).Name("dot in folder name")
'AccUnit:Row("README.md", False).Name("md")
'AccUnit:Row("notes.txt", False).Name("txt")
'AccUnit:Row("LICENSE", False).Name("no extension")
'AccUnit:Row("a.cls/Makefile", False).Name("dot only in folder name")
'AccUnit:Row("x.", False).Name("empty extension")
'AccUnit:Row(".gitignore", False).Name("gitignore")
'AccUnit:Row(".gitattributes", False).Name("gitattributes")
'AccUnit:Row(".editorconfig", False).Name("editorconfig")
'AccUnit:Row("x.github", False).Name("extension github")
'AccUnit:Row("data/SqlTools.cls.txt", False).Name("double extension")
'AccUnit:Row("_dotnetlib/lib/x.dll", False).Name("dll")
'AccUnit:Row(".github/x.cls", False).Name("cls below github folder")
'AccUnit:Row(".git/x.bas", False).Name("bas below git folder")
'AccUnit:Row("sub/.GitHub/y.cls", False).Name("nested github folder, mixed case")
'AccUnit:Row(".github", False).Name("file named github")
Public Sub IsVisibleFile_Path_ReturnsExpected(ByVal EntryPath As String, ByVal Expected As Boolean)
   Assert.That m_Filter.IsVisibleFile(EntryPath), Iz.EqualTo(Expected)
End Sub

Public Sub IsVisibleFile_AdditionalExtensions_AreShownBesideDefault()

   m_Filter.Initialize "md, .TXT"

   Assert.That m_Filter.IsVisibleFile("README.md"), Iz.EqualTo(True)
   Assert.That m_Filter.IsVisibleFile("notes.txt"), Iz.EqualTo(True)
   Assert.That m_Filter.IsVisibleFile("data/SqlTools.cls"), Iz.EqualTo(True)
   Assert.That m_Filter.IsVisibleFile(".github/README.md"), Iz.EqualTo(False)

End Sub

'--- entries

Public Sub RemoveHiddenEntries_MixedEntries_KeepsVisibleFilesAndTheirFolders()

   Dim Entries As Scripting.Dictionary
   Set Entries = New Scripting.Dictionary
   Entries.Add ".github", "tree"
   Entries.Add ".github/workflows", "tree"
   Entries.Add ".github/workflows/x.yml", "blob"
   Entries.Add ".github/y.cls", "blob"
   Entries.Add ".gitignore", "blob"
   Entries.Add "README.md", "blob"
   Entries.Add "data", "tree"
   Entries.Add "data/.gitkeep", "blob"
   Entries.Add "data/SqlTools.cls", "blob"
   Entries.Add "data/Readme.txt", "blob"
   Entries.Add "_packages", "tree"
   Entries.Add "_packages/Template.package", "blob"

   Dim Result As Scripting.Dictionary
   Set Result = m_Filter.RemoveHiddenEntries(Entries)

   Assert.That KeyList(Result), Iz.EqualTo("data|data/SqlTools.cls|_packages|_packages/Template.package")
   Assert.That Result("data"), Iz.EqualTo("tree")
   Assert.That Result("data/SqlTools.cls"), Iz.EqualTo("blob")
   Assert.That Entries.Count, Iz.EqualTo(12)

End Sub

Public Sub RemoveHiddenEntries_FolderWithoutVisibleFile_IsRemoved()

   Dim Entries As Scripting.Dictionary
   Set Entries = New Scripting.Dictionary
   Entries.Add "empty", "tree"
   Entries.Add "docs", "tree"
   Entries.Add "docs/README.md", "blob"
   Entries.Add "docs/img", "tree"
   Entries.Add "docs/img/logo.gif", "blob"
   Entries.Add "src", "tree"
   Entries.Add "src/a.bas", "blob"

   Assert.That KeyList(m_Filter.RemoveHiddenEntries(Entries)), Iz.EqualTo("src|src/a.bas")

End Sub

Public Sub RemoveHiddenEntries_VisibleFileDeepInside_KeepsAllParentFolders()

   Dim Entries As Scripting.Dictionary
   Set Entries = New Scripting.Dictionary
   Entries.Add "a", "tree"
   Entries.Add "a/b", "tree"
   Entries.Add "a/b/c", "tree"
   Entries.Add "a/b/c/Deep.cls", "blob"
   Entries.Add "a/b/other", "tree"
   Entries.Add "a/b/other/x.txt", "blob"
   Entries.Add "a/x.md", "blob"

   Assert.That KeyList(m_Filter.RemoveHiddenEntries(Entries)), Iz.EqualTo("a|a/b|a/b/c|a/b/c/Deep.cls")

End Sub

Public Sub RemoveHiddenEntries_FolderPathInOtherCase_IsStillKept()

   Dim Entries As Scripting.Dictionary
   Set Entries = New Scripting.Dictionary
   Entries.Add "Data", "tree"
   Entries.Add "data/A.cls", "blob"

   Assert.That KeyList(m_Filter.RemoveHiddenEntries(Entries)), Iz.EqualTo("Data|data/A.cls")

End Sub

Public Sub RemoveHiddenEntries_GithubIsFile_IsRemoved()

   Dim Entries As Scripting.Dictionary
   Set Entries = New Scripting.Dictionary
   Entries.Add ".github", "blob"
   Entries.Add "src", "tree"
   Entries.Add "src/a.cls", "blob"

   Assert.That KeyList(m_Filter.RemoveHiddenEntries(Entries)), Iz.EqualTo("src|src/a.cls")

End Sub

Public Sub RemoveHiddenEntries_FolderNamedLikeCodeFile_IsRemovedWithoutVisibleFile()

   Dim Entries As Scripting.Dictionary
   Set Entries = New Scripting.Dictionary
   Entries.Add "x.cls", "tree"
   Entries.Add "x.cls/readme.md", "blob"

   Assert.That m_Filter.RemoveHiddenEntries(Entries).Count, Iz.EqualTo(0)

End Sub

Public Sub RemoveHiddenEntries_NoEntries_ReturnsEmpty()
   Assert.That m_Filter.RemoveHiddenEntries(New Scripting.Dictionary).Count, Iz.EqualTo(0)
End Sub

Public Sub RemoveHiddenEntries_AdditionalExtensions_ShowsAddedFilesAndFolders()

   m_Filter.Initialize "md"

   Dim Entries As Scripting.Dictionary
   Set Entries = New Scripting.Dictionary
   Entries.Add "docs", "tree"
   Entries.Add "docs/README.md", "blob"
   Entries.Add "src", "tree"
   Entries.Add "src/a.cls", "blob"

   Assert.That KeyList(m_Filter.RemoveHiddenEntries(Entries)), Iz.EqualTo("docs|docs/README.md|src|src/a.cls")

End Sub

Private Function KeyList(ByVal Entries As Scripting.Dictionary) As String

   Dim Result As String
   Dim Key As Variant

   For Each Key In Entries.Keys
      If Len(Result) > 0 Then Result = Result & "|"
      Result = Result & CStr(Key)
   Next

   KeyList = Result

End Function
