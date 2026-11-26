VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitHubTreeEntryParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: GitHubTreeEntryParser
'---------------------------------------------------------------------------------------
'
' Parses the flat response of the GitHub trees API (recursive=1)
'
' Remarks:
'---------------------------------------------------------------------------------------
'| Reads only "path" and "type" of the entries in the "tree" array and the flag "truncated".
'| The response is read in one pass (InStr from the current position, no substrings
'| of the whole remaining text), so the run time grows linearly with the response size.
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const ERR_INVALID_TREE_JSON As Long = vbObjectError + 6101
Private Const ERR_TREE_TRUNCATED As Long = vbObjectError + 6102

Private Const ErrorSource As String = "GitHubTreeEntryParser.ParseRecursiveTree"
Private Const InvalidTreeJsonMessage As String = "Invalid GitHub tree response"
Private Const TreeTruncatedMessage As String = _
      "The repository tree is too large to be loaded completely (GitHub response truncated)."

Private Const JsonNameTree As String = "tree"
Private Const JsonNameTruncated As String = "truncated"
Private Const JsonNamePath As String = "path"
Private Const JsonNameType As String = "type"

Private Const EntryTypeFolder As String = "tree"
Private Const EntryTypeFile As String = "blob"

Private Const JsonTrue As String = "true"
Private Const JsonFalse As String = "false"
Private Const JsonWhitespaceChars As String = " " & vbTab & vbCr & vbLf
Private Const JsonLiteralEndChars As String = ",}]" & JsonWhitespaceChars
Private Const UnicodeEscapeHexLength As Long = 4

Private m_Json As String
Private m_JsonLength As Long
Private m_Position As Long

' Returns Scripting.Dictionary: key = full path ("data/SqlTools.cls"), value = "tree" or "blob", in response order.
Public Function ParseRecursiveTree(ByVal JsonString As String) As Scripting.Dictionary

   Dim Entries As Scripting.Dictionary
   Dim IsTruncated As Boolean
   Dim MemberName As String

   BeginParse JsonString

   If Not TryReadChar("{") Then
      RaiseInvalidTreeJson
   End If

   If Not TryReadChar("}") Then
      Do
         MemberName = ReadString()
         ReadExpectedChar ":"
         If IsJsonName(MemberName, JsonNameTree) Then
            Set Entries = ReadTreeEntries()
         ElseIf IsJsonName(MemberName, JsonNameTruncated) Then
            IsTruncated = ReadBoolean()
         Else
            SkipValue
         End If
      Loop While TryReadChar(",")
      ReadExpectedChar "}"
   End If

   If Entries Is Nothing Then
      RaiseInvalidTreeJson
   End If

   If IsTruncated Then
      Err.Raise ERR_TREE_TRUNCATED, ErrorSource, TreeTruncatedMessage
   End If

   Set ParseRecursiveTree = Entries

End Function

Private Sub BeginParse(ByVal JsonString As String)
   m_Json = JsonString
   m_JsonLength = Len(JsonString)
   m_Position = 1
End Sub

Private Function ReadTreeEntries() As Scripting.Dictionary

   Dim Entries As Scripting.Dictionary

   Set Entries = New Scripting.Dictionary

   ReadExpectedChar "["
   If Not TryReadChar("]") Then
      Do
         ReadTreeEntry Entries
      Loop While TryReadChar(",")
      ReadExpectedChar "]"
   End If

   Set ReadTreeEntries = Entries

End Function

Private Sub ReadTreeEntry(ByVal Entries As Scripting.Dictionary)

   Dim MemberName As String
   Dim EntryPath As String
   Dim EntryType As String

   ReadExpectedChar "{"
   If Not TryReadChar("}") Then
      Do
         MemberName = ReadString()
         ReadExpectedChar ":"
         If IsJsonName(MemberName, JsonNamePath) Then
            EntryPath = ReadString()
         ElseIf IsJsonName(MemberName, JsonNameType) Then
            EntryType = ReadString()
         Else
            SkipValue
         End If
      Loop While TryReadChar(",")
      ReadExpectedChar "}"
   End If

   ' Other types (e.g. "commit" for submodules) are no files or folders of the repository
   If IsJsonName(EntryType, EntryTypeFolder) Or IsJsonName(EntryType, EntryTypeFile) Then
      If Len(EntryPath) > 0 Then
         Entries.Item(EntryPath) = EntryType
      End If
   End If

End Sub

Private Function IsJsonName(ByVal Value As String, ByVal ExpectedName As String) As Boolean
   ' JSON names and values are case sensitive
   IsJsonName = (StrComp(Value, ExpectedName, vbBinaryCompare) = 0)
End Function

Private Sub SkipValue()

   Select Case PeekChar()
      Case """"
         ReadString
      Case "{"
         SkipObject
      Case "["
         SkipArray
      Case Else
         SkipLiteral
   End Select

End Sub

Private Sub SkipObject()

   ReadExpectedChar "{"
   If Not TryReadChar("}") Then
      Do
         ReadString
         ReadExpectedChar ":"
         SkipValue
      Loop While TryReadChar(",")
      ReadExpectedChar "}"
   End If

End Sub

Private Sub SkipArray()

   ReadExpectedChar "["
   If Not TryReadChar("]") Then
      Do
         SkipValue
      Loop While TryReadChar(",")
      ReadExpectedChar "]"
   End If

End Sub

Private Sub SkipLiteral()
' numbers, true, false, null

   Dim StartPosition As Long

   SkipWhitespace
   StartPosition = m_Position
   Do While m_Position <= m_JsonLength
      If InStr(1, JsonLiteralEndChars, Mid$(m_Json, m_Position, 1), vbBinaryCompare) > 0 Then
         Exit Do
      End If
      m_Position = m_Position + 1
   Loop

   If m_Position = StartPosition Then
      RaiseInvalidTreeJson
   End If

End Sub

Private Function ReadBoolean() As Boolean

   SkipWhitespace

   If IsJsonName(Mid$(m_Json, m_Position, Len(JsonTrue)), JsonTrue) Then
      m_Position = m_Position + Len(JsonTrue)
      ReadBoolean = True
   ElseIf IsJsonName(Mid$(m_Json, m_Position, Len(JsonFalse)), JsonFalse) Then
      m_Position = m_Position + Len(JsonFalse)
      ReadBoolean = False
   Else
      RaiseInvalidTreeJson
   End If

End Function

Private Function ReadString() As String

   Dim Result As String
   Dim QuotePosition As Long
   Dim BackslashOffset As Long
   Dim Segment As String

   ReadExpectedChar """"

   Do
      QuotePosition = InStr(m_Position, m_Json, """", vbBinaryCompare)
      If QuotePosition = 0 Then
         RaiseInvalidTreeJson
      End If

      Segment = Mid$(m_Json, m_Position, QuotePosition - m_Position)
      BackslashOffset = InStr(1, Segment, "\", vbBinaryCompare)

      If BackslashOffset = 0 Then
         m_Position = QuotePosition + 1
         ReadString = Result & Segment
         Exit Function
      End If

      Result = Result & Left$(Segment, BackslashOffset - 1)
      m_Position = m_Position + BackslashOffset - 1
      Result = Result & ReadEscapedChar()
   Loop

End Function

Private Function ReadEscapedChar() As String
' m_Position is at the backslash; moves behind the escape sequence

   Dim EscapeChar As String
   Dim HexDigits As String

   If m_Position >= m_JsonLength Then
      RaiseInvalidTreeJson
   End If

   EscapeChar = Mid$(m_Json, m_Position + 1, 1)

   ' AscW: escape letters are case sensitive (Option Compare Text would also accept upper case letters)
   Select Case AscW(EscapeChar)
      Case AscW(""""), AscW("\"), AscW("/")
         ReadEscapedChar = EscapeChar
      Case AscW("b")
         ReadEscapedChar = vbBack
      Case AscW("f")
         ReadEscapedChar = vbFormFeed
      Case AscW("n")
         ReadEscapedChar = vbLf
      Case AscW("r")
         ReadEscapedChar = vbCr
      Case AscW("t")
         ReadEscapedChar = vbTab
      Case AscW("u")
         HexDigits = Mid$(m_Json, m_Position + 2, UnicodeEscapeHexLength)
         If Not HexDigits Like "[0-9A-F][0-9A-F][0-9A-F][0-9A-F]" Then
            RaiseInvalidTreeJson
         End If
         ReadEscapedChar = ChrW$(CLng("&H" & HexDigits))
         m_Position = m_Position + UnicodeEscapeHexLength
      Case Else
         RaiseInvalidTreeJson
   End Select

   m_Position = m_Position + 2

End Function

Private Function PeekChar() As String
   SkipWhitespace
   PeekChar = Mid$(m_Json, m_Position, 1)
End Function

Private Function TryReadChar(ByVal ExpectedChar As String) As Boolean

   If StrComp(PeekChar(), ExpectedChar, vbBinaryCompare) = 0 Then
      m_Position = m_Position + 1
      TryReadChar = True
   End If

End Function

Private Sub ReadExpectedChar(ByVal ExpectedChar As String)
   If Not TryReadChar(ExpectedChar) Then
      RaiseInvalidTreeJson
   End If
End Sub

Private Sub SkipWhitespace()
   Do While m_Position <= m_JsonLength
      If InStr(1, JsonWhitespaceChars, Mid$(m_Json, m_Position, 1), vbBinaryCompare) = 0 Then
         Exit Do
      End If
      m_Position = m_Position + 1
   Loop
End Sub

Private Sub RaiseInvalidTreeJson()
   Err.Raise ERR_INVALID_TREE_JSON, ErrorSource, InvalidTreeJsonMessage
End Sub
