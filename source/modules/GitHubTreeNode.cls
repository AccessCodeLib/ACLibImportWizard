VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitHubTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const GitHubApiTreesUrl As String = "https://api.github.com/repos/{owner}/{repo}/git/trees/{SHA}"

Private m_Properties As Dictionary
Private m_Nodes As Dictionary

Private Sub Class_Initialize()
   Set m_Properties = New Dictionary
   Set m_Nodes = New Dictionary
End Sub

Public Property Get Properties() As Dictionary
   Set Properties = m_Properties
End Property

Public Property Get Nodes() As Dictionary
   Set Nodes = m_Nodes
End Property

Public Property Get IsFolder() As Boolean
   If NodeType = "tree" Then
      IsFolder = True
   ElseIf Len(NodeType) = 0 Then
      IsFolder = True
   End If
End Property

Public Property Get NodeType() As String
   If m_Properties.Exists("type") Then
      NodeType = m_Properties.Item("type")
   End If
End Property

Public Property Get Path() As String
   If m_Properties.Exists("path") Then
      Path = m_Properties.Item("path")
   End If
End Property

Public Property Get Sha() As String
   If m_Properties.Exists("sha") Then
      Sha = m_Properties.Item("sha")
   End If
End Property

Public Sub BuildFromSHA(ByVal Sha As String, _
               Optional ByVal ReadSubTreeLevel As Long = 0)

   Dim JsonString As String

   With ACLibGitHubImporter

      JsonString = .GetJsonString(Replace(GitHubApiTreesUrl, "{SHA}", Sha))

      If JsonString Like "{""message"":*" Then
         Err.Raise vbObjectError, "GitHubTreeNode.BuildFromSHA", BuildErrMsgFromGitHubApiMessage(JsonString)
      End If

      BuildFromJson JsonString, ReadSubTreeLevel

   End With

End Sub

Private Function BuildErrMsgFromGitHubApiMessage(ByVal MsgJson As String) As String
' {"message":"API rate limit exceeded for 84.112.118.32. (But here's the good news: Authenticated requests get a higher rate limit. Check out the documentation for more details.)","documentation_url":"https://docs.github.com/rest/overview/resources-in-the-rest-api#rate-limiting"}

   Dim Msg As String
   Msg = Replace(MsgJson, "{""message"":""", vbNullString)
   Msg = Replace(Msg, """,""documentation_url"":""", vbNewLine & "URL: ")
   Msg = Replace(Msg, """}", vbNullString)

   BuildErrMsgFromGitHubApiMessage = Msg

End Function

Friend Sub BuildFromJson(ByVal JsonString As String, Optional ByVal ReadSubTreeLevel As Long = 0)

   Dim Props As JsonProperties

   With New GitHubTreeJsonConverter
      Set Props = .Json2TreeNode(JsonString)
   End With

   InitFromJsonProperties Props, ReadSubTreeLevel

End Sub

Friend Sub InitFromJsonProperties(ByVal Props As JsonProperties, Optional ByVal ReadSubTreeLevel As Long = 0)

   Dim Prop As JsonProperty
   Dim i As Long

   For i = 1 To Props.Count
      Set Prop = Props.Item(i)
      If Prop.Name = "Tree" Then
         AddNodesfromJsonCollection Prop.Value, ReadSubTreeLevel
      Else
         m_Properties(Prop.Name) = Prop.Value
      End If
   Next

End Sub

Private Sub AddNodesfromJsonCollection(ByVal JsonNodeCollection As Collection, Optional ByVal ReadSubTreeLevel As Long = 0)

   Dim NodeProps As JsonProperties
   Dim Node As GitHubTreeNode

   For Each NodeProps In JsonNodeCollection
      Set Node = GetNodeFromNodeProps(NodeProps)
      Set m_Nodes(Node.Path) = Node
   Next

   If ReadSubTreeLevel > 0 Then
      FillNextTreeLevel ReadSubTreeLevel
   End If

End Sub

Private Function GetNodeFromNodeProps(ByVal NodeProps As JsonProperties) As GitHubTreeNode

   Dim Node As GitHubTreeNode

   Set Node = New GitHubTreeNode
   Node.InitFromJsonProperties NodeProps

   Set GetNodeFromNodeProps = Node

End Function

Friend Sub FillNextTreeLevel(Optional ByVal ReadSubTreeLevel As Long = 0)

   Dim Node As GitHubTreeNode
   Dim NodeKey As Variant

   If Not IsFolder Then
      Exit Sub
   End If

   If m_Nodes.Count = 0 Then
      Me.BuildFromSHA Me.Sha, ReadSubTreeLevel
      Exit Sub
   End If

   If ReadSubTreeLevel > 0 Then
      For Each NodeKey In m_Nodes.Keys
         Set Node = m_Nodes(NodeKey)
         If Node.IsFolder Then
            Node.FillNextTreeLevel ReadSubTreeLevel - 1
         End If
      Next
   End If

End Sub
