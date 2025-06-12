VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitHubTreeNodeTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'AccUnit:TestClass

'TestSuite.AddByClassName("GitHubTreeNodeTests").Run

Private m_TreeNode As GitHubTreeNode

Public Sub Setup()
   Set m_TreeNode = New GitHubTreeNode
End Sub

Public Sub Teardown()
   Set m_TreeNode = Nothing
End Sub

Public Sub JsonString_Without_Nodes()

Const TestJson As String = _
"{" & vbNewLine & _
"  ""sha"": ""f369b695f119081de974a5c9a0357bf563e7ec94""," & vbNewLine & _
"  ""url"": ""https://api.github.com/repos/AccessCodeLib/AccessCodeLib/git/trees/f369b695f119081de974a5c9a0357bf563e7ec94"" " & vbNewLine & _
"}"

   m_TreeNode.BuildFromJson TestJson

   Assert.That m_TreeNode.Nodes.Count, Iz.EqualTo(0)
   Assert.That m_TreeNode.Properties.Count, Iz.EqualTo(2)

End Sub

Public Sub JsonString_With_Nodes()

Const TestJson As String = _
"{" & vbNewLine & _
"  ""sha"": ""f369b695f119081de974a5c9a0357bf563e7ec94""," & vbNewLine & _
"  ""url"": ""https://api.github.com/repos/AccessCodeLib/AccessCodeLib/git/trees/f369b695f119081de974a5c9a0357bf563e7ec94""," & vbNewLine & _
"  ""tree"": [" & vbNewLine & _
"    {" & vbNewLine & _
"      ""path"": ""LICENSE""," & vbNewLine & _
"      ""mode"": ""100644""," & vbNewLine & _
"      ""type"": ""blob""," & vbNewLine & _
"      ""sha"": ""4bd7d68f9f4597e859a7667b0a1fe1576cff05d3""," & vbNewLine & _
"      ""size"": 1505," & vbNewLine & _
"      ""url"": ""https://api.github.com/repos/AccessCodeLib/AccessCodeLib/git/blobs/4bd7d68f9f4597e859a7667b0a1fe1576cff05d3""" & vbNewLine & _
"    }," & vbNewLine & _
"    {" & vbNewLine & _
"      ""path"": ""README.md""," & vbNewLine & _
"      ""mode"": ""100644""," & vbNewLine & _
"      ""type"": ""blob""," & vbNewLine & _
"      ""sha"": ""ad54412d9b3f1efc4cb5c008d699fddde24bfc16""," & vbNewLine & _
"      ""size"": 225," & vbNewLine & _
"      ""url"": ""https://api.github.com/repos/AccessCodeLib/AccessCodeLib/git/blobs/ad54412d9b3f1efc4cb5c008d699fddde24bfc16""" & vbNewLine & _
"    }" & vbNewLine & _
"    ]," & vbNewLine & _
"}"

   m_TreeNode.BuildFromJson TestJson

   Assert.That m_TreeNode.Properties.Count, Iz.EqualTo(2)
   Assert.That m_TreeNode.Nodes.Count, Iz.EqualTo(2)

   Dim NodeKey As Variant
   Dim Node As GitHubTreeNode
   For Each NodeKey In m_TreeNode.Nodes.Keys
      Set Node = m_TreeNode.Nodes(NodeKey)
      Assert.That Node.Nodes.Count, Iz.EqualTo(0)
   Next

End Sub

Public Sub BuildFromSHA_master_CheckCount()

   m_TreeNode.BuildFromSHA "master", 0

   Assert.That m_TreeNode.Properties.Count, Iz.EqualTo(3)
   Assert.That m_TreeNode.Nodes.Count, Iz.EqualTo(28)

   Dim NodeKey As Variant
   Dim Node As GitHubTreeNode
   For Each NodeKey In m_TreeNode.Nodes.Keys
      Set Node = m_TreeNode.Nodes(NodeKey)
      Assert.That Node.Nodes.Count, Iz.EqualTo(0)
   Next

End Sub

Public Sub BuildFromSHA_masterWithSubTrees_CheckCount()

   m_TreeNode.BuildFromSHA "master", 1

   Assert.That m_TreeNode.Properties.Count, Iz.EqualTo(3)
   Assert.That m_TreeNode.Nodes.Count, Iz.EqualTo(28)

   Dim Node As GitHubTreeNode
   Set Node = m_TreeNode.Nodes.Item("data")
   Assert.That Node.Nodes.Count, Iz.GreaterThan(0)

End Sub

Public Sub BuildFromSHA_master_CheckSubTreeDaoFromData()

   m_TreeNode.BuildFromSHA "master", 1

   Assert.That m_TreeNode.Properties.Count, Iz.EqualTo(3)
   Assert.That m_TreeNode.Nodes.Count, Iz.EqualTo(28)

   Dim DataNode As GitHubTreeNode
   Dim DaoNode As GitHubTreeNode

   Set DataNode = m_TreeNode.Nodes.Item("data")
   Assert.That DataNode.Nodes.Count, Iz.GreaterThan(0)
   Set DaoNode = DataNode.Nodes("dao")
   Assert.That DataNode.NodeType, Iz.EqualTo("tree")

End Sub
