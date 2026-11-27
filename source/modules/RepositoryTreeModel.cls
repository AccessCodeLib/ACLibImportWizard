VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryTreeModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: RepositoryTreeModel
'---------------------------------------------------------------------------------------
'
' State of the repository tree: expand/collapse, check state (tri-state), filter, result
'
' Remarks:
'---------------------------------------------------------------------------------------
'| Independent of the display: no form references, no message boxes, no CurrentDb.
'| Initialize reads the whole tree with one source call; afterwards the source is never used again.
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const ERR_UNKNOWN_ITEM_KEY As Long = vbObjectError + 6001
Private Const ERR_NOT_INITIALIZED As Long = vbObjectError + 6002

Private Const ErrorSourcePrefix As String = "RepositoryTreeModel."
Private Const UnknownItemKeyMessage As String = "Unknown tree item: "
Private Const NotInitializedMessage As String = "Repository tree is not initialized"

' Items with one of these path parts are not shown (compared case insensitive)
Private Const HiddenPathPartList As String = "_info.txt|_tools"
Private Const HiddenPathPartListSeparator As String = "|"
Private Const PathSeparator As String = "/"

Private m_HiddenPathParts() As String

Private m_IsInitialized As Boolean
Private m_Items As Scripting.Dictionary   ' Key => RepositoryTreeItem
Private m_RootKeys As Collection          ' keys of the root items in source order
Private m_FilterText As String

Private Sub Class_Initialize()
   m_HiddenPathParts = Split(HiddenPathPartList, HiddenPathPartListSeparator)
End Sub

' Reads the whole tree (one source call). On an error the previous state is kept and the error is passed on.
Public Sub Initialize(ByVal Source As RepositoryTreeSource)

   Dim SourceItems As Collection
   Dim NewItems As Scripting.Dictionary
   Dim NewRootKeys As Collection

   Set SourceItems = Source.ReadAllItems()

   Set NewItems = New Scripting.Dictionary
   Set NewRootKeys = New Collection
   BuildTree SourceItems, NewItems, NewRootKeys

   Set m_Items = NewItems
   Set m_RootKeys = NewRootKeys
   m_FilterText = vbNullString
   m_IsInitialized = True

End Sub

Private Sub BuildTree(ByVal SourceItems As Collection, _
                      ByVal Items As Scripting.Dictionary, ByVal RootKeys As Collection)
' The source delivers parents before their children (GitHub tree order).
' Own item instances: a second Initialize always starts with a fresh state.

   Dim SourceItem As RepositoryTreeItem
   Dim Item As RepositoryTreeItem
   Dim ParentItem As RepositoryTreeItem

   For Each SourceItem In SourceItems
      If Not IsHiddenPath(SourceItem.Key) Then

         Set Item = New RepositoryTreeItem
         Item.Init SourceItem.Key, SourceItem.IsFolder

         If Len(Item.ParentKey) = 0 Then
            Items.Add Item.Key, Item
            RootKeys.Add Item.Key
         ElseIf Items.Exists(Item.ParentKey) Then
            Set ParentItem = Items.Item(Item.ParentKey)
            If ParentItem.IsFolder Then
               Items.Add Item.Key, Item
               ParentItem.ChildKeys.Add Item.Key
            End If
         End If
         ' Items without a known parent folder are ignored

      End If
   Next

End Sub

Private Function IsHiddenPath(ByVal Key As String) As Boolean

   Dim PathPart As Variant
   Dim HiddenPathPart As Variant

   For Each PathPart In Split(Key, PathSeparator)
      For Each HiddenPathPart In m_HiddenPathParts
         If StrComp(PathPart, HiddenPathPart, vbTextCompare) = 0 Then
            IsHiddenPath = True
            Exit Function
         End If
      Next
   Next

End Function

Public Property Get FilterText() As String
   CheckInitialized "FilterText"
   FilterText = m_FilterText
End Property

Public Property Let FilterText(ByVal NewValue As String)
   CheckInitialized "FilterText"
   m_FilterText = Trim$(NewValue)
End Property

Private Property Get IsFilterActive() As Boolean
   IsFilterActive = Len(m_FilterText) > 0
End Property

' Files match if their name contains the filter text; folder names are no matches
Private Function IsFilterMatch(ByVal Item As RepositoryTreeItem) As Boolean
   If Not Item.IsFolder Then
      IsFilterMatch = InStr(1, Item.Name, m_FilterText, vbTextCompare) > 0
   End If
End Function

Public Sub ToggleExpanded(ByVal Key As String)

   Dim Item As RepositoryTreeItem

   Set Item = FindItem(Key, "ToggleExpanded")

   ' With an active filter all visible folders are shown expanded
   If IsFilterActive Then
      Exit Sub
   End If

   If Item.HasChildren Then
      Item.IsExpanded = Not Item.IsExpanded
   End If

End Sub

Public Sub ToggleChecked(ByVal Key As String)

   Dim Item As RepositoryTreeItem

   Set Item = FindItem(Key, "ToggleChecked")

   If Not Item.IsFolder Then
      ToggleFileChecked Item
   ElseIf IsFilterActive Then
      ToggleMatchingFilesChecked Item
   Else
      ToggleFolderChecked Item
   End If

   DeriveFolderCheckStates

End Sub

Private Sub ToggleFileChecked(ByVal Item As RepositoryTreeItem)
   If Item.CheckState = CheckStateChecked Then
      Item.CheckState = CheckStateUnchecked
   Else
      Item.CheckState = CheckStateChecked
   End If
End Sub

Private Sub ToggleFolderChecked(ByVal Folder As RepositoryTreeItem)

   Dim NewState As RepositoryTreeCheckState
   Dim Item As RepositoryTreeItem

   If Folder.CheckState = CheckStateChecked Then
      NewState = CheckStateUnchecked
   Else
      NewState = CheckStateChecked
   End If

   Folder.CheckState = NewState
   For Each Item In GetDescendants(Folder)
      Item.CheckState = NewState
   Next

End Sub

Private Sub ToggleMatchingFilesChecked(ByVal Folder As RepositoryTreeItem)
' Only the files below the folder that match the filter are changed:
' all checked => uncheck them, otherwise check them all.
' The folder state itself does not count because it includes hidden files.

   Dim MatchingFiles As Collection
   Dim Item As RepositoryTreeItem
   Dim NewState As RepositoryTreeCheckState

   Set MatchingFiles = New Collection
   For Each Item In GetDescendants(Folder)
      If IsFilterMatch(Item) Then
         MatchingFiles.Add Item
      End If
   Next

   NewState = CheckStateUnchecked
   For Each Item In MatchingFiles
      If Item.CheckState <> CheckStateChecked Then
         NewState = CheckStateChecked
         Exit For
      End If
   Next

   For Each Item In MatchingFiles
      Item.CheckState = NewState
   Next

End Sub

Private Sub DeriveFolderCheckStates()

   Dim RootKey As Variant
   Dim FileCount As Long
   Dim CheckedFileCount As Long

   For Each RootKey In m_RootKeys
      DeriveCheckState m_Items.Item(RootKey), FileCount, CheckedFileCount
   Next

End Sub

Private Sub DeriveCheckState(ByVal Item As RepositoryTreeItem, _
                             ByRef FileCount As Long, ByRef CheckedFileCount As Long)
' Adds the files of the subtree to the counters and sets the state of every folder from the files below it.
' A folder without files in its subtree keeps its own state and does not affect its ancestors.

   Dim ChildKey As Variant
   Dim SubtreeFileCount As Long
   Dim SubtreeCheckedFileCount As Long

   If Not Item.IsFolder Then
      FileCount = FileCount + 1
      If Item.CheckState = CheckStateChecked Then
         CheckedFileCount = CheckedFileCount + 1
      End If
      Exit Sub
   End If

   For Each ChildKey In Item.ChildKeys
      DeriveCheckState m_Items.Item(ChildKey), SubtreeFileCount, SubtreeCheckedFileCount
   Next

   If SubtreeFileCount > 0 Then
      Item.CheckState = GetCheckStateFromFileCount(SubtreeFileCount, SubtreeCheckedFileCount)
   End If

   FileCount = FileCount + SubtreeFileCount
   CheckedFileCount = CheckedFileCount + SubtreeCheckedFileCount

End Sub

Private Function GetCheckStateFromFileCount(ByVal FileCount As Long, _
                                            ByVal CheckedFileCount As Long) As RepositoryTreeCheckState
   If CheckedFileCount = 0 Then
      GetCheckStateFromFileCount = CheckStateUnchecked
   ElseIf CheckedFileCount = FileCount Then
      GetCheckStateFromFileCount = CheckStateChecked
   Else
      GetCheckStateFromFileCount = CheckStatePartiallyChecked
   End If
End Function

Public Function GetItem(ByVal Key As String) As RepositoryTreeItem
   Set GetItem = FindItem(Key, "GetItem")
End Function

' Returns a Collection of RepositoryTreeItem in display order and sets IsShownExpanded of each returned item
Public Function GetVisibleItems() As Collection

   Dim VisibleItems As Collection
   Dim RootKey As Variant

   CheckInitialized "GetVisibleItems"

   Set VisibleItems = New Collection
   For Each RootKey In m_RootKeys
      If IsFilterActive Then
         AddFilteredItems m_Items.Item(RootKey), VisibleItems
      Else
         AddExpandedItems m_Items.Item(RootKey), VisibleItems
      End If
   Next

   Set GetVisibleItems = VisibleItems

End Function

Private Sub AddExpandedItems(ByVal Item As RepositoryTreeItem, ByVal VisibleItems As Collection)

   Dim ChildKey As Variant

   Item.IsShownExpanded = Item.IsExpanded
   VisibleItems.Add Item

   If Item.IsExpanded Then
      For Each ChildKey In Item.ChildKeys
         AddExpandedItems m_Items.Item(ChildKey), VisibleItems
      Next
   End If

End Sub

Private Sub AddFilteredItems(ByVal Item As RepositoryTreeItem, ByVal VisibleItems As Collection)
' Matching files and all their ancestors; every shown folder is shown expanded

   Dim SubtreeItems As Collection
   Dim ChildKey As Variant
   Dim SubtreeItem As RepositoryTreeItem

   If Not Item.IsFolder Then
      If IsFilterMatch(Item) Then
         Item.IsShownExpanded = False
         VisibleItems.Add Item
      End If
      Exit Sub
   End If

   Set SubtreeItems = New Collection
   For Each ChildKey In Item.ChildKeys
      AddFilteredItems m_Items.Item(ChildKey), SubtreeItems
   Next

   If SubtreeItems.Count > 0 Then
      Item.IsShownExpanded = True
      VisibleItems.Add Item
      For Each SubtreeItem In SubtreeItems
         VisibleItems.Add SubtreeItem
      Next
   End If

End Sub

' Returns a Collection of String: all checked file keys in tree order, independent of the filter, never folders
Public Function GetCheckedFileKeys() As Collection

   Dim CheckedFileKeys As Collection
   Dim RootKey As Variant
   Dim Item As RepositoryTreeItem

   CheckInitialized "GetCheckedFileKeys"

   Set CheckedFileKeys = New Collection
   For Each RootKey In m_RootKeys
      Set Item = m_Items.Item(RootKey)
      AddCheckedFileKey Item, CheckedFileKeys
      For Each Item In GetDescendants(Item)
         AddCheckedFileKey Item, CheckedFileKeys
      Next
   Next

   Set GetCheckedFileKeys = CheckedFileKeys

End Function

Private Sub AddCheckedFileKey(ByVal Item As RepositoryTreeItem, ByVal CheckedFileKeys As Collection)
   If Not Item.IsFolder Then
      If Item.CheckState = CheckStateChecked Then
         CheckedFileKeys.Add Item.Key
      End If
   End If
End Sub

Public Property Get CheckedFileCount() As Long
   CheckInitialized "CheckedFileCount"
   CheckedFileCount = GetCheckedFileKeys().Count
End Property

' Returns all items below the item in tree order (depth first, source order per level)
Private Function GetDescendants(ByVal Item As RepositoryTreeItem) As Collection

   Dim Descendants As Collection

   Set Descendants = New Collection
   AddDescendants Item, Descendants

   Set GetDescendants = Descendants

End Function

Private Sub AddDescendants(ByVal Item As RepositoryTreeItem, ByVal Descendants As Collection)

   Dim ChildKey As Variant
   Dim ChildItem As RepositoryTreeItem

   For Each ChildKey In Item.ChildKeys
      Set ChildItem = m_Items.Item(ChildKey)
      Descendants.Add ChildItem
      AddDescendants ChildItem, Descendants
   Next

End Sub

Private Function FindItem(ByVal Key As String, ByVal ProcedureName As String) As RepositoryTreeItem

   CheckInitialized ProcedureName

   If Not m_Items.Exists(Key) Then
      Err.Raise ERR_UNKNOWN_ITEM_KEY, ErrorSourcePrefix & ProcedureName, UnknownItemKeyMessage & Key
   End If

   Set FindItem = m_Items.Item(Key)

End Function

Private Sub CheckInitialized(ByVal ProcedureName As String)
   If Not m_IsInitialized Then
      Err.Raise ERR_NOT_INITIALIZED, ErrorSourcePrefix & ProcedureName, NotInitializedMessage
   End If
End Sub
