VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryTreeSourceFake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestRelated

Implements RepositoryTreeSource

Private Const FakeErrorNumber As Long = vbObjectError + 42

Private m_ReadAllCallCount As Long
Private m_RaiseOnReadAll As Boolean
Private m_UseOrphanItems As Boolean

' True: the list contains items without a parent folder (parent missing or parent is a file)
Public Property Let UseOrphanItems(ByVal NewValue As Boolean)
   m_UseOrphanItems = NewValue
End Property

Public Property Get ReadAllCallCount() As Long
   ReadAllCallCount = m_ReadAllCallCount
End Property

Public Property Get RaiseOnReadAll() As Boolean
   RaiseOnReadAll = m_RaiseOnReadAll
End Property

Public Property Let RaiseOnReadAll(ByVal NewValue As Boolean)
   m_RaiseOnReadAll = NewValue
End Property

Private Function RepositoryTreeSource_ReadAllItems() As Collection

   m_ReadAllCallCount = m_ReadAllCallCount + 1

   If m_RaiseOnReadAll Then
      Err.Raise FakeErrorNumber, "RepositoryTreeSourceFake", "Fake source error"
   End If

   Dim Items As Collection
   Set Items = New Collection

   If m_UseOrphanItems Then
      AddItem Items, "README.md", False
      AddItem Items, "README.md/child.txt", False
      AddItem Items, "missing/file.txt", False
      AddItem Items, "ok", True
      AddItem Items, "ok/a.txt", False
      Set RepositoryTreeSource_ReadAllItems = Items
      Exit Function
   End If

   AddItem Items, "README.md", False
   AddItem Items, "_info.txt", False
   AddItem Items, "_tools", True
   AddItem Items, "_tools/x.cls", False
   AddItem Items, "base", True
   AddItem Items, "base/ApplicationHandler.cls", False
   AddItem Items, "base/modErrorHandler.bas", False
   AddItem Items, "data", True
   AddItem Items, "data/_info.txt", False
   AddItem Items, "data/SqlTools.cls", False
   AddItem Items, "data/dao", True
   AddItem Items, "data/dao/DaoHandler.cls", False
   AddItem Items, "data/dao/empty", True
   AddItem Items, "text", True
   AddItem Items, "text/_INFO.TXT", False
   AddItem Items, "text/StringTools.bas", False

   Set RepositoryTreeSource_ReadAllItems = Items

End Function

Private Sub AddItem(ByVal Items As Collection, ByVal Key As String, ByVal IsFolder As Boolean)

   Dim Item As RepositoryTreeItem
   Set Item = New RepositoryTreeItem
   Item.Init Key, IsFolder

   Items.Add Item

End Sub
