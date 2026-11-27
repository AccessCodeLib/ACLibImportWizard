VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryTreeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: RepositoryTreeItem
'---------------------------------------------------------------------------------------
'
' One node (file or folder) of the repository tree: data and display state
'
' Remarks:
'---------------------------------------------------------------------------------------
'| The state is changed only by RepositoryTreeModel (Friend setters).
'| The display reads HasChildren and IsShownExpanded for the expand symbol.
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

' declared here because the item holds the state
Public Enum RepositoryTreeCheckState
   CheckStateUnchecked = 0
   CheckStateChecked = 1
   CheckStatePartiallyChecked = 2
End Enum

Private Const PathSeparator As String = "/"

Private m_Key As String
Private m_Name As String
Private m_ParentKey As String
Private m_Level As Long
Private m_IsFolder As Boolean
Private m_IsExpanded As Boolean
Private m_IsShownExpanded As Boolean
Private m_CheckState As RepositoryTreeCheckState
Private m_ChildKeys As Collection

Private Sub Class_Initialize()
   Set m_ChildKeys = New Collection
End Sub

Friend Sub Init(ByVal Key As String, ByVal IsFolder As Boolean)

   Dim SeparatorPosition As Long

   m_Key = Key
   m_IsFolder = IsFolder

   SeparatorPosition = InStrRev(Key, PathSeparator, , vbBinaryCompare)
   m_Name = Mid$(Key, SeparatorPosition + 1)
   If SeparatorPosition > 0 Then
      m_ParentKey = Left$(Key, SeparatorPosition - 1)
   Else
      m_ParentKey = vbNullString
   End If

   m_Level = Len(Key) - Len(Replace(Key, PathSeparator, vbNullString, , , vbBinaryCompare))

End Sub

' full repository path with "/", e.g. data/dao/DaoHandler.cls
Public Property Get Key() As String
   Key = m_Key
End Property

' last part of the path
Public Property Get Name() As String
   Name = m_Name
End Property

' path without the last part; empty on root level
Public Property Get ParentKey() As String
   ParentKey = m_ParentKey
End Property

' 0 on root level
Public Property Get Level() As Long
   Level = m_Level
End Property

Public Property Get IsFolder() As Boolean
   IsFolder = m_IsFolder
End Property

' stored state (folders only); the filter does not change it
Public Property Get IsExpanded() As Boolean
   IsExpanded = m_IsExpanded
End Property

Friend Property Let IsExpanded(ByVal NewValue As Boolean)
   m_IsExpanded = NewValue
End Property

' display in the current view, set by RepositoryTreeModel.GetVisibleItems
Public Property Get IsShownExpanded() As Boolean
   IsShownExpanded = m_IsShownExpanded
End Property

Friend Property Let IsShownExpanded(ByVal NewValue As Boolean)
   m_IsShownExpanded = NewValue
End Property

Public Property Get CheckState() As RepositoryTreeCheckState
   CheckState = m_CheckState
End Property

Friend Property Let CheckState(ByVal NewValue As RepositoryTreeCheckState)
   m_CheckState = NewValue
End Property

' keys of the child items in source order
Friend Property Get ChildKeys() As Collection
   Set ChildKeys = m_ChildKeys
End Property

Public Property Get HasChildren() As Boolean
   HasChildren = m_IsFolder And m_ChildKeys.Count > 0
End Property
