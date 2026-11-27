VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryPathFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: RepositoryPathFilter
'---------------------------------------------------------------------------------------
'
' Reduces the entries of the repository tree to the files that can be imported
'
' Remarks:
'---------------------------------------------------------------------------------------
'| The tree is used to import code modules, so only files with a listed extension are shown
'| (default: .bas, .cls, .frm, .rep and .package; the wizard shows package files, but does not import them).
'| Package files are recognized by the extension, because the tree does not contain file contents.
'| Folders without a visible file (recursively) are removed.
'| The folders .git and .github are always hidden, including everything below them.
'| The default list is part of the code, so later extensions of it reach every installation.
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const DefaultExtensionList As String = ".bas;.cls;.frm;.rep;.package"
Private Const ExcludedFolderNames As String = ".git|.github"

Private Const PathSeparator As String = "/"
Private Const NameListSeparator As String = "|"
Private Const ExtensionSeparator As String = ";"
Private Const ExtensionDot As String = "."
Private Const EntryTypeFolder As String = "tree"

Private m_Extensions As Scripting.Dictionary

Private Sub Class_Initialize()
   Set m_Extensions = ParseExtensionList(DefaultExtensionList)
End Sub

' True if the path is a file with a listed extension outside of the excluded folders
Public Function IsVisibleFile(ByVal EntryPath As String) As Boolean

   If IsInExcludedFolder(EntryPath) Then Exit Function
   IsVisibleFile = m_Extensions.Exists(GetExtension(EntryPath))

End Function

' Returns a new Scripting.Dictionary (same format as GitHubTreeEntryParser.ParseRecursiveTree) with the
' visible files and the folders that contain at least one visible file; the order is kept.
Public Function RemoveHiddenEntries(ByVal Entries As Scripting.Dictionary) As Scripting.Dictionary

   Dim VisibleFolders As Scripting.Dictionary
   Dim Result As Scripting.Dictionary
   Dim EntryPath As Variant

   Set VisibleFolders = CollectFoldersOfVisibleFiles(Entries)

   Set Result = New Scripting.Dictionary
   Result.CompareMode = Entries.CompareMode

   For Each EntryPath In Entries.Keys
      If IsVisibleEntry(CStr(EntryPath), CStr(Entries.Item(EntryPath)), VisibleFolders) Then
         Result.Add EntryPath, Entries.Item(EntryPath)
      End If
   Next

   Set RemoveHiddenEntries = Result

End Function

Private Function IsVisibleEntry(ByVal EntryPath As String, ByVal EntryType As String, _
                                ByVal VisibleFolders As Scripting.Dictionary) As Boolean

   If IsFolderType(EntryType) Then
      IsVisibleEntry = VisibleFolders.Exists(EntryPath)
   Else
      IsVisibleEntry = IsVisibleFile(EntryPath)
   End If

End Function

Private Function IsFolderType(ByVal EntryType As String) As Boolean
   IsFolderType = (StrComp(EntryType, EntryTypeFolder, vbTextCompare) = 0)
End Function

Private Function CollectFoldersOfVisibleFiles(ByVal Entries As Scripting.Dictionary) As Scripting.Dictionary

   Dim Folders As Scripting.Dictionary
   Dim EntryPath As Variant

   Set Folders = New Scripting.Dictionary
   Folders.CompareMode = TextCompare

   For Each EntryPath In Entries.Keys
      If Not IsFolderType(CStr(Entries.Item(EntryPath))) Then
         If IsVisibleFile(CStr(EntryPath)) Then AddParentFolders CStr(EntryPath), Folders
      End If
   Next

   Set CollectFoldersOfVisibleFiles = Folders

End Function

Private Sub AddParentFolders(ByVal FilePath As String, ByVal Folders As Scripting.Dictionary)

   Dim SeparatorPos As Long

   SeparatorPos = InStr(1, FilePath, PathSeparator, vbBinaryCompare)

   Do While SeparatorPos > 0
      Folders.Item(Left$(FilePath, SeparatorPos - 1)) = True
      SeparatorPos = InStr(SeparatorPos + 1, FilePath, PathSeparator, vbBinaryCompare)
   Loop

End Sub

Private Function IsInExcludedFolder(ByVal EntryPath As String) As Boolean

   Dim Segments() As String
   Dim Index As Long

   Segments = Split(EntryPath, PathSeparator)

   For Index = 0 To UBound(Segments)
      If IsNameInList(Segments(Index), ExcludedFolderNames) Then
         IsInExcludedFolder = True
         Exit Function
      End If
   Next

End Function

Private Function IsNameInList(ByVal SegmentName As String, ByVal NameList As String) As Boolean

   Dim Names() As String
   Dim Index As Long

   Names = Split(NameList, NameListSeparator)

   For Index = LBound(Names) To UBound(Names)
      If StrComp(SegmentName, Names(Index), vbTextCompare) = 0 Then
         IsNameInList = True
         Exit Function
      End If
   Next

End Function

' Extension of the last path segment without dot, lower case; empty if the name has no dot
Private Function GetExtension(ByVal EntryPath As String) As String

   Dim DotPos As Long

   DotPos = InStrRev(EntryPath, ExtensionDot, -1, vbBinaryCompare)
   If DotPos > InStrRev(EntryPath, PathSeparator, -1, vbBinaryCompare) Then
      GetExtension = LCase$(Mid$(EntryPath, DotPos + 1))
   End If

End Function

Private Function ParseExtensionList(ByVal ExtensionList As String) As Scripting.Dictionary

   Dim Extensions As Scripting.Dictionary
   Dim Parts() As String
   Dim Index As Long

   Set Extensions = New Scripting.Dictionary
   Extensions.CompareMode = TextCompare

   Parts = Split(ExtensionList, ExtensionSeparator)

   For Index = 0 To UBound(Parts)
      Extensions.Item(LCase$(Mid$(Parts(Index), Len(ExtensionDot) + 1))) = True
   Next

   Set ParseExtensionList = Extensions

End Function
