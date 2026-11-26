VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryTreeSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: RepositoryTreeSource
'---------------------------------------------------------------------------------------
'
' Delivers all nodes of the GitHub repository tree to RepositoryTreeModel
'
' Remarks:
'---------------------------------------------------------------------------------------
'| Repository, branch and token come from the predeclared ACLibGitHubImporter.
'| No filtering here: hiding items is model logic.
'| Test doubles implement this class (Implements RepositoryTreeSource).
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const BranchPlaceholder As String = "{branch}"
Private Const EntryTypeFolder As String = "tree"

' Returns a Collection of RepositoryTreeItem in source order: the whole tree, one API call
Public Function ReadAllItems() As Collection
   Set ReadAllItems = CreateItems(ReadTreeEntries())
End Function

Private Function ReadTreeEntries() As Scripting.Dictionary

   Dim JsonString As String

   With New GitHubTreeNode
      JsonString = .ReadRecursiveTreeJson(BranchPlaceholder)
   End With

   With New GitHubTreeEntryParser
      Set ReadTreeEntries = .ParseRecursiveTree(JsonString)
   End With

End Function

Private Function CreateItems(ByVal Entries As Scripting.Dictionary) As Collection

   Dim Items As Collection
   Dim Item As RepositoryTreeItem
   Dim EntryPath As Variant

   Set Items = New Collection

   For Each EntryPath In Entries.Keys
      Set Item = New RepositoryTreeItem
      Item.Init EntryPath, (StrComp(Entries.Item(EntryPath), EntryTypeFolder, vbBinaryCompare) = 0)
      Items.Add Item
   Next

   Set CreateItems = Items

End Function
