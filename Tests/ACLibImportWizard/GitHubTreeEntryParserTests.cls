VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitHubTreeEntryParserTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

'TestSuite.AddByClassName("GitHubTreeEntryParserTests").Run

Private Const ErrInvalidTreeJson As Long = vbObjectError + 6101
Private Const ErrTreeTruncated As Long = vbObjectError + 6102

Private Const EntrySha As String = "4bd7d68f9f4597e859a7667b0a1fe1576cff05d3"

Private m_Parser As GitHubTreeEntryParser

Public Sub Setup()
   Set m_Parser = New GitHubTreeEntryParser
End Sub

Public Sub Teardown()
   Set m_Parser = Nothing
End Sub

'--- P01

Public Sub ParseRecursiveTree_CompactResponse_ReturnsEntriesInOrder()

   Dim Json As String
   Json = Response(Entry("data", "tree") & "," & _
                   Entry("data/SqlTools.cls", "blob") & "," & _
                   Entry("README.md", "blob"))

   Dim Result As Scripting.Dictionary
   Set Result = m_Parser.ParseRecursiveTree(Json)

   Assert.That Result.Count, Iz.EqualTo(3)
   Assert.That KeyList(Result), Iz.EqualTo("data|data/SqlTools.cls|README.md")
   Assert.That Result("data"), Iz.EqualTo("tree")
   Assert.That Result("data/SqlTools.cls"), Iz.EqualTo("blob")
   Assert.That Result("README.md"), Iz.EqualTo("blob")

End Sub

'--- P02

Public Sub ParseRecursiveTree_FormattedResponse_ReturnsEntriesInOrder()

   Dim Json As String
   Json = "{" & vbCrLf & _
          "  ""sha"": """ & EntrySha & """," & vbCrLf & _
          "  ""url"": ""https://api.github.com/repos/x/y/git/trees/abc""," & vbCrLf & _
          "  ""tree"": [" & vbCrLf & _
          "    {" & vbCrLf & _
          "      ""path"": ""data""," & vbCrLf & _
          "      ""mode"": ""040000""," & vbCrLf & _
          "      ""type"": ""tree""," & vbCrLf & _
          "      ""sha"": """ & EntrySha & """," & vbCrLf & _
          "      ""url"": ""https://api.github.com/x""" & vbCrLf & _
          "    }," & vbCrLf & _
          "    {" & vbCrLf & _
          "      ""path"": ""data/SqlTools.cls""," & vbCrLf & _
          "      ""mode"": ""100644""," & vbCrLf & _
          "      ""type"": ""blob""," & vbCrLf & _
          "      ""sha"": """ & EntrySha & """," & vbCrLf & _
          "      ""size"": 1505," & vbCrLf & _
          "      ""url"": ""https://api.github.com/y""" & vbCrLf & _
          "    },"  & vbCrLf
   Json = Json & _
          "    {" & vbCrLf & _
          "      ""path"": ""README.md""," & vbCrLf & _
          "      ""mode"": ""100644""," & vbCrLf & _
          "      ""type"": ""blob""," & vbCrLf & _
          "      ""sha"": """ & EntrySha & """," & vbCrLf & _
          "      ""size"": 225," & vbCrLf & _
          "      ""url"": ""https://api.github.com/z""" & vbCrLf & _
          "    }" & vbCrLf & _
          "  ]," & vbCrLf & _
          "  ""truncated"": false" & vbCrLf & _
          "}"

   Dim Result As Scripting.Dictionary
   Set Result = m_Parser.ParseRecursiveTree(Json)

   Assert.That Result.Count, Iz.EqualTo(3)
   Assert.That KeyList(Result), Iz.EqualTo("data|data/SqlTools.cls|README.md")
   Assert.That Result("data"), Iz.EqualTo("tree")
   Assert.That Result("data/SqlTools.cls"), Iz.EqualTo("blob")
   Assert.That Result("README.md"), Iz.EqualTo("blob")

End Sub

'--- P03

Public Sub ParseRecursiveTree_TruncatedTrue_RaisesTreeTruncated()

   Dim Json As String
   Json = Response(Entry("README.md", "blob"), "true")

   Assert.Throws ErrTreeTruncated
   m_Parser.ParseRecursiveTree Json

End Sub

'--- P04

Public Sub ParseRecursiveTree_GitHubMessage_RaisesInvalidTreeJson()

   Assert.Throws ErrInvalidTreeJson
   m_Parser.ParseRecursiveTree "{""message"":""Not Found""}"

End Sub

Public Sub ParseRecursiveTree_EmptyString_RaisesInvalidTreeJson()

   Assert.Throws ErrInvalidTreeJson
   m_Parser.ParseRecursiveTree vbNullString

End Sub

'--- P05

Public Sub ParseRecursiveTree_EmptyTreeArray_ReturnsEmptyDictionary()

   Dim Result As Scripting.Dictionary
   Set Result = m_Parser.ParseRecursiveTree(Response(vbNullString))

   Assert.That Result.Count, Iz.EqualTo(0)

End Sub

'--- P06

Public Sub ParseRecursiveTree_CommitEntry_IsSkipped()

   Dim Json As String
   Json = Response(Entry("data", "tree") & "," & _
                   Entry("submodule", "commit") & "," & _
                   Entry("README.md", "blob"))

   Dim Result As Scripting.Dictionary
   Set Result = m_Parser.ParseRecursiveTree(Json)

   Assert.That KeyList(Result), Iz.EqualTo("data|README.md")

End Sub

'--- P07

Public Sub ParseRecursiveTree_EscapedPath_IsDecoded()

   ' JSON text: a + raw a-umlaut + backslash-u-00e4 + backslash-slash + b + backslash-quote + c
   Dim JsonPath As String
   JsonPath = "a" & ChrW(&HE4) & "\" & "u00e4\/b\""c"

   Dim ExpectedKey As String
   ExpectedKey = "a" & ChrW(&HE4) & ChrW(&HE4) & "/b""c"

   Dim Result As Scripting.Dictionary
   Set Result = m_Parser.ParseRecursiveTree(Response(Entry(JsonPath, "blob")))

   Assert.That Result.Count, Iz.EqualTo(1)
   Assert.That Result.Keys()(0), Iz.EqualTo(ExpectedKey)

End Sub

'--- P08

Public Sub ParseRecursiveTree_1000Entries_ParsesInUnderOneSecond()

   Dim Entries As String
   Dim Index As Long
   For Index = 1 To 1000
      If Index > 1 Then Entries = Entries & ","
      Entries = Entries & Entry("folder/file" & Index & ".txt", "blob")
   Next
   Dim Json As String
   Json = Response(Entries)

   Dim StartTime As Single
   StartTime = Timer

   Dim Result As Scripting.Dictionary
   Set Result = m_Parser.ParseRecursiveTree(Json)

   Dim ElapsedSeconds As Single
   ElapsedSeconds = Timer - StartTime

   Assert.That Result.Count, Iz.EqualTo(1000)
   Assert.That ElapsedSeconds, Iz.LessThan(1)

End Sub

'--- P09

Public Sub ParseRecursiveTree_TypeBeforePath_IsAssignedCorrectly()

   Dim Json As String
   Json = Response("{""type"":""tree"",""path"":""data"",""sha"":""" & EntrySha & """}," & _
                   "{""sha"":""" & EntrySha & """,""type"":""blob"",""path"":""data/x.cls""}")

   Dim Result As Scripting.Dictionary
   Set Result = m_Parser.ParseRecursiveTree(Json)

   Assert.That Result.Count, Iz.EqualTo(2)
   Assert.That Result("data"), Iz.EqualTo("tree")
   Assert.That Result("data/x.cls"), Iz.EqualTo("blob")

End Sub

'--- P10

Public Sub ParseRecursiveTree_ValuesEqualToKeywords_AreNotMistakenForKeys()

   Dim Json As String
   Json = Response(Entry("tree", "tree") & "," & _
                   Entry("tree/type", "blob") & "," & _
                   Entry("tree/path.txt", "blob") & "," & _
                   Entry("a\\b", "blob"))

   Dim Result As Scripting.Dictionary
   Set Result = m_Parser.ParseRecursiveTree(Json)

   Assert.That Result.Count, Iz.EqualTo(4)
   Assert.That KeyList(Result), Iz.EqualTo("tree|tree/type|tree/path.txt|a\b")
   Assert.That Result("tree"), Iz.EqualTo("tree")
   Assert.That Result("tree/type"), Iz.EqualTo("blob")
   Assert.That Result("tree/path.txt"), Iz.EqualTo("blob")
   Assert.That Result("a\b"), Iz.EqualTo("blob")

End Sub

'--- additional cases from review S3

Public Sub ParseRecursiveTree_FormattedTruncatedTrueAfterArray_RaisesTreeTruncated()

   Dim Json As String
   Json = "{" & vbCrLf & _
          "  ""sha"": """ & EntrySha & """," & vbCrLf & _
          "  ""tree"": [" & vbCrLf & _
          "    " & Entry("README.md", "blob") & vbCrLf & _
          "  ]," & vbCrLf & _
          "  ""truncated"": true" & vbCrLf & _
          "}"

   Assert.Throws ErrTreeTruncated
   m_Parser.ParseRecursiveTree Json

End Sub

Public Sub ParseRecursiveTree_TruncatedTrueWithoutTree_RaisesInvalidTreeJson()

   Assert.Throws ErrInvalidTreeJson
   m_Parser.ParseRecursiveTree "{""sha"":""abc"",""truncated"": true}"

End Sub

Public Sub ParseRecursiveTree_UnknownNestedFields_AreSkipped()

   Dim Json As String
   Json = Response("{""path"":""data"",""extra"":{""a"":[1,2,{""b"":""x]}""}],""c"":""y""}," & _
                   """list"":[""p"",""q""],""type"":""tree"",""n"":null,""f"":false}," & _
                   Entry("README.md", "blob"))

   Dim Result As Scripting.Dictionary
   Set Result = m_Parser.ParseRecursiveTree(Json)

   Assert.That KeyList(Result), Iz.EqualTo("data|README.md")
   Assert.That Result("data"), Iz.EqualTo("tree")

End Sub

Public Sub ParseRecursiveTree_UnterminatedString_RaisesInvalidTreeJson()

   Assert.Throws ErrInvalidTreeJson
   m_Parser.ParseRecursiveTree "{""tree"":[{""path"":""abc"

End Sub

Public Sub ParseRecursiveTree_InvalidEscape_RaisesInvalidTreeJson()

   Assert.Throws ErrInvalidTreeJson
   m_Parser.ParseRecursiveTree Response(Entry("a\xb", "blob"))

End Sub

Public Sub ParseRecursiveTree_ShortUnicodeEscape_RaisesInvalidTreeJson()

   Assert.Throws ErrInvalidTreeJson
   m_Parser.ParseRecursiveTree Response(Entry("a\u12", "blob"))

End Sub

Public Sub ParseRecursiveTree_ResponseCutOffInArray_RaisesInvalidTreeJson()

   Dim Json As String
   Json = "{""sha"":""abc"",""tree"":[" & Entry("README.md", "blob") & ","

   Assert.Throws ErrInvalidTreeJson
   m_Parser.ParseRecursiveTree Json

End Sub

Public Sub ParseRecursiveTree_UnicodeEscapeUpperCaseHex_IsDecoded()

   Dim JsonPath As String
   JsonPath = "\" & "u00C4"

   Dim Result As Scripting.Dictionary
   Set Result = m_Parser.ParseRecursiveTree(Response(Entry(JsonPath, "blob")))

   Assert.That Result.Count, Iz.EqualTo(1)
   Assert.That Result.Keys()(0), Iz.EqualTo(ChrW(&HC4))

End Sub

'--- helpers (private: not tests)

' One tree entry in the field order of the real GitHub response; Path must already be JSON-escaped.
Private Function Entry(ByVal JsonPath As String, ByVal EntryType As String) As String
   Entry = "{""path"":""" & JsonPath & """,""mode"":""100644"",""type"":""" & EntryType & _
           """,""sha"":""" & EntrySha & """,""size"":10,""url"":""https://api.github.com/x""}"
End Function

Private Function Response(ByVal Entries As String, Optional ByVal Truncated As String = "false") As String
   Response = "{""sha"":""" & EntrySha & """,""url"":""https://api.github.com/t""," & _
              """tree"":[" & Entries & "],""truncated"":" & Truncated & "}"
End Function

Private Function KeyList(ByVal Entries As Scripting.Dictionary) As String

   Dim Result As String
   Dim Key As Variant

   For Each Key In Entries.Keys
      If Len(Result) > 0 Then Result = Result & "|"
      Result = Result & CStr(Key)
   Next

   KeyList = Result

End Function
