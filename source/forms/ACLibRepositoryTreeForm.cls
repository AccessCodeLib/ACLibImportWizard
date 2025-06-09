Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: ACLibRepositoryTreeForm
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

#Const TreeViewEarlyBinding = 0

' MSComctlLib GUID = {831FDD16-0C5C-11D2-A9FC-0000F8754DA1}
#If TreeViewEarlyBinding = 1 Then
Private m_TreeView As MSComctlLib.TreeView
#Else
Private m_TreeView As Object
Const tvwChild As Long = 4
#End If

Private m_RepoRootNode As GitHubTreeNode
Private m_TreeviewEventHandlerOn As Boolean

Public Event FilesSelected(ByRef FileArray() As String)
Public Event Unload()

Private Sub Form_Load()
   Me.TimerInterval = 1
End Sub

Private Sub Form_Timer()
   Me.TimerInterval = 0
   InitTreeView
   Me.labStartInfo.Visible = False
   Me.tvRepositoryTree.Visible = True
End Sub

Private Sub cmdCommit_Click()
   SendSelectedFiles
End Sub

Private Sub cmdCancel_Click()
   DoCmd.Close
End Sub

Private Sub SendSelectedFiles()

#If TreeViewEarlyBinding = 1 Then
   Dim TvNode As Node
#Else
   Dim TvNode As Object
#End If

   Dim Files() As String
   Dim SelectedIndex As Long

   ReDim Files(m_TreeView.Nodes.Count) As String

   SelectedIndex = -1
   For Each TvNode In m_TreeView.Nodes
      If TvNode.Checked Then
      If TvNode.Children = 0 Then
         SelectedIndex = SelectedIndex + 1
         Files(SelectedIndex) = TvNode.Key
      End If
      End If
   Next

   If SelectedIndex = -1 Then
      MsgBox "Nothing has been selected.", vbInformation, "ACLib Import Wizard"
      Exit Sub
   End If

   ReDim Preserve Files(SelectedIndex)

   RaiseEvent FilesSelected(Files)

End Sub

Private Sub InitTreeView()
   Set m_TreeView = Me.tvRepositoryTree.Object
   FillRepositoryTree
End Sub

Private Sub FillRepositoryTree()

   m_TreeView.Nodes.Clear
   Set m_RepoRootNode = GetRootTreeNode()
   FillSubTree Nothing, m_RepoRootNode

End Sub

#If TreeViewEarlyBinding = 1 Then
Private Sub CheckSubTree(ByVal ParentNode As Node)
#Else
Private Sub CheckSubTree(ByVal ParentNode As Object)
#End If

   Dim ReproNode As GitHubTreeNode

#If TreeViewEarlyBinding = 1 Then
   Dim TvNode As Node
#Else
   Dim TvNode As Object
#End If

   If ParentNode.Children = 0 Then
      FillTreeViewSubNode ParentNode
      Exit Sub
   End If

   Set ReproNode = GetReproNode(ParentNode.Key)

   Set TvNode = ParentNode.Child
   Do While Not (TvNode Is Nothing)
      If TvNode.Children = 0 Then
         FillTreeViewSubNode TvNode
      End If
      Set TvNode = TvNode.Next
   Loop

End Sub

#If TreeViewEarlyBinding = 1 Then
Private Sub FillSubTree(ByVal ParentNode As Node, ByVal ReproParentNode As GitHubTreeNode)
#Else
Private Sub FillSubTree(ByVal ParentNode As Object, ByVal ReproParentNode As GitHubTreeNode)
#End If

   Dim NodeKey As Variant
   Dim ReproNode As GitHubTreeNode

#If TreeViewEarlyBinding = 1 Then
   Dim TvNode As Node
#Else
   Dim TvNode As Object
#End If


   For Each NodeKey In ReproParentNode.Nodes.Keys
      Set ReproNode = ReproParentNode.Nodes.Item(NodeKey)

      If CheckShowGitHubNode(ReproNode) Then

         If ParentNode Is Nothing Then
            Set TvNode = m_TreeView.Nodes.Add(, , ReproNode.Path, ReproNode.Path)
         Else
            Set TvNode = m_TreeView.Nodes.Add(ParentNode.Key, tvwChild, ParentNode.Key & "/" & ReproNode.Path, ReproNode.Path)
         End If

         If ReproNode.IsFolder Then
            TvNode.Tag = "FLDR"
            FillSubTree TvNode, ReproNode
         End If

      End If
   Next

End Sub

Private Function CheckShowGitHubNode(ByVal ReproNode As GitHubTreeNode) As Boolean

   Select Case ReproNode.Path
      Case "_info.txt", "_tools"
         CheckShowGitHubNode = False
      Case Else
         CheckShowGitHubNode = True
   End Select

End Function

Private Function GetRootTreeNode() As GitHubTreeNode

   Dim Node As GitHubTreeNode

   Set Node = New GitHubTreeNode
   With New ACLibConfiguration
      Node.BuildFromSHA "{branch}", 1
   End With
   Set GetRootTreeNode = Node

End Function

Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload
End Sub

Private Sub tvRepositoryTree_Expand(ByVal Node As Object)
   If m_TreeviewEventHandlerOn Then
      CheckSubTree Node
   End If
End Sub

Private Sub tvRepositoryTree_GotFocus()
  m_TreeviewEventHandlerOn = True
End Sub

Private Sub tvRepositoryTree_LostFocus()
   m_TreeviewEventHandlerOn = False
End Sub

Private Sub tvRepositoryTree_NodeCheck(ByVal Node As Object)

   If Node.Children > 0 Then
      CheckSubBlobNodes Node
   End If

End Sub

Private Sub tvRepositoryTree_NodeClick(ByVal Node As Object)
   'FillTreeViewSubNode Node
   CheckSubTree Node
End Sub

Private Sub CheckSubBlobNodes(ByVal Node As Object)

#If TreeViewEarlyBinding = 1 Then
   Dim ChildNode As MSComctlLib.Node
#Else
   Dim ChildNode As Object
#End If

   Set ChildNode = Node.Child
   Do While Not (ChildNode Is Nothing)
      If ChildNode.Tag = "FLDR" Then
         ' do not check sub nodes
      Else
         ChildNode.Checked = Node.Checked
      End If
      Set ChildNode = ChildNode.Next
   Loop

End Sub

#If TreeViewEarlyBinding = 1 Then
Private Sub FillTreeViewSubNode(ByVal Node As MSComctlLib.Node)
#Else
Private Sub FillTreeViewSubNode(ByVal Node As Object)
#End If

   Dim ReproNode As GitHubTreeNode

   If Node.Children > 0 Then
      ' Childs exists + no refresh => exit
      If Node.Child.Children > 0 Then
         Exit Sub
      End If

   End If

   Set ReproNode = GetReproNode(Node.Key)
   If ReproNode.NodeType = "tree" Then
      FillSubTree Node, ReproNode
   End If

End Sub

Private Function GetReproNode(ByVal NodeKey As String, Optional ByVal ReadSubTreeLevel As Long = 1) As GitHubTreeNode

   Dim ReproNode As GitHubTreeNode
   Dim Keys() As String
   Dim i As Long

   NodeKey = Replace(NodeKey, ACLibGitHubImporter.BranchName & "/", vbNullString)
   Keys = Split(NodeKey, "/")

   Set ReproNode = m_RepoRootNode
   For i = LBound(Keys) To UBound(Keys)
      Set ReproNode = ReproNode.Nodes(Keys(i))
   Next

   If ReproNode.Nodes.Count = 0 Then
      ReproNode.FillNextTreeLevel ReadSubTreeLevel
   End If

   Set GetReproNode = ReproNode

End Function
