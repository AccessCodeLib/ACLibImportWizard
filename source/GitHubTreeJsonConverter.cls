VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitHubTreeJsonConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const SubPropIndentLen As Long = 2

Public Function Json2TreeNode(ByVal JsonString As String) As JsonProperties

    JsonString = OptimizeJsonString(JsonString)
    
    If Left(JsonString, 1) <> "{" Then
        Err.Raise vbObjectError, "JsonConverter.Parse", "Missing JSON structure"
    End If
    
    If Right(JsonString, 1) <> "}" Then
        Err.Raise vbObjectError, "JsonConverter.Parse", "Missing JSON structure"
    End If
    
    Set Json2TreeNode = ParseProperties(JsonString)
    
End Function

'#############################################
' Parse JSON

Private Function OptimizeJsonString(ByVal JsonString As String) As String

    JsonString = Replace(JsonString, vbNewLine, vbNullString)
    JsonString = Replace(JsonString, vbTab, vbNullString)
    JsonString = Replace(JsonString, vbCr, vbNullString)
    JsonString = Replace(JsonString, vbLf, vbNullString)
    
    JsonString = Trim(JsonString)
    
    OptimizeJsonString = JsonString

End Function

Private Function ParseProperties(ByVal JsonString As String) As JsonProperties

   Dim Props As JsonProperties
   Set Props = New JsonProperties
   
   Dim PropStringArray() As String
   Dim i As Long
   Dim PropString As String
   
   JsonString = Trim(JsonString)
   If Left(JsonString, 1) = "{" And Right(JsonString, 1) = "}" Then ' if
      JsonString = Mid(JsonString, 2, Len(JsonString) - 2)
   Else
      Err.Raise vbObjectError, "JsonConverter.ParseProperties", "Missing JSON structure"
   End If
   
   
   PropStringArray = Split(SetSplitMarker(JsonString), ",|")
   
   For i = LBound(PropStringArray) To UBound(PropStringArray)
      PropString = Trim(PropStringArray(i))
      If Len(PropString) > 0 Then
         Props.Append ConvertString2Property(PropString)
      End If
   Next
   
   Set ParseProperties = Props

End Function

Private Function ConvertString2Property(ByVal JsonString As String) As JsonProperty
'   "Name" : "abc" => Value
'   "Name" : 123   => Value
'   "Name" : { ... => Properties

    Dim Prop As JsonProperty
    
    Dim PropName As String
    Dim PropValueString As Variant
    
    Dim CutPos As Long
    
    Set Prop = New JsonProperty
    
    CutPos = InStr(1, JsonString, ":")
    
    PropName = Trim(Left(JsonString, CutPos - 1))
    If Left(PropName, 1) = """" And Right(PropName, 1) = """" Then
        PropName = Mid(PropName, 2, Len(PropName) - 2)
    End If
    Prop.Name = PropName
    
    PropValueString = Trim(Mid(JsonString, CutPos + 1))
    
    If Left(PropValueString, 1) = "{" Then ' new properties
        Set Prop.Properties = ParseProperties(PropValueString)
    ElseIf Left(PropValueString, 1) = "[" Then ' new properties
        Set Prop.ObjectValue = ConvertString2JsonItem(PropValueString)
    Else
        Prop.Value = ConvertString2JsonItem(PropValueString)
    End If
    
    Set ConvertString2Property = Prop
    
End Function

Private Function ConvertString2JsonItem(ByVal PropValueString As String) As Variant

   If Left(PropValueString, 1) = "{" Then ' new properties
        Set ConvertString2JsonItem = ParseProperties(PropValueString)
    ElseIf Left(PropValueString, 1) = "[" Then ' new nodes
        Set ConvertString2JsonItem = ParseJsonArray(PropValueString)
    ElseIf PropValueString = "null" Then
        ConvertString2JsonItem = Null
    ElseIf PropValueString = "true" Then
        ConvertString2JsonItem = True
    ElseIf PropValueString = "false" Then
        ConvertString2JsonItem = False
    ElseIf Left(PropValueString, 1) = """" Then 'string
        ConvertString2JsonItem = Mid(PropValueString, 2, Len(PropValueString) - 2)
    Else ' Variant
        ConvertString2JsonItem = CVar(PropValueString)
    End If

End Function

Private Function ParseJsonArray(ByVal JsonString As String) As Collection

   Dim Items As Collection
   Dim PropStringArray() As String
   Dim i As Long
   Dim Props As Variant
   
   If Left(JsonString, 1) = "[" And Right(JsonString, 1) = "]" Then
       JsonString = Trim(Mid(JsonString, 2, Len(JsonString) - 2))
   End If
   
   PropStringArray = Split(SetSplitMarker(JsonString), ",|")
   
   Set Items = New Collection
   For i = LBound(PropStringArray) To UBound(PropStringArray)
      Items.Add ConvertString2JsonItem(Trim(PropStringArray(i)))
   Next
   
   Set ParseJsonArray = Items

End Function

Private Function ParseValueFromString(ByVal JsonString As String) As Variant
    
    JsonString = Trim(JsonString)
    
    If JsonString = "null" Then
        ParseValueFromString = Null
    ElseIf JsonString = "true" Then
        ParseValueFromString = True
    ElseIf JsonString = "false" Then
        ParseValueFromString = False
    ElseIf Left(JsonString, 1) = """" And Right(JsonString, 1) = """" Then 'string
        ParseValueFromString = Mid(JsonString, 2, Len(JsonString) - 2)
    Else
        ParseValueFromString = Val(JsonString)
    End If
    
End Function

Private Function SetSplitMarker(ByVal JsonString As String) As String

    Dim Pos As Long
    Dim LeftPart As String
    Dim CheckString As String
    
    JsonString = Replace(JsonString, ",", ", ") ' place to insert marker

    Pos = InStr(1, JsonString, ",")
    
    Do While Pos > 0
        CheckString = Left(JsonString, Pos - 1)
        If Not PosIsInSubObject(CheckString) Then
            If Not PosIsInArray(CheckString) Then
                Mid(JsonString, Pos + 1, 1) = "|"
            End If
        End If
        Pos = InStr(Pos + 1, JsonString, ",")
    Loop
    
    SetSplitMarker = Replace(JsonString, ", ", ",")

End Function

Private Function PosIsInSubObject(ByVal StringToCheck As String) As Boolean

    PosIsInSubObject = PosIsInOpenStruct(StringToCheck, "{", "}")
    
End Function

Private Function PosIsInArray(ByVal StringToCheck As String) As Boolean

    PosIsInArray = PosIsInOpenStruct(StringToCheck, "[", "]")
    
End Function

Private Function PosIsInOpenStruct(ByVal StringToCheck As String, ByVal OpeningChar As String, ByVal ClosingChar As String) As Boolean

    Dim CntOpeningBrackets As Long
    Dim CntClosingBrackets As Long
    
    CntOpeningBrackets = Len(StringToCheck) - Len(Replace(StringToCheck, OpeningChar, vbNullString))
    CntClosingBrackets = Len(StringToCheck) - Len(Replace(StringToCheck, ClosingChar, vbNullString))
    
    If CntOpeningBrackets > CntClosingBrackets Then
        PosIsInOpenStruct = True
    Else
        PosIsInOpenStruct = False
    End If
    
End Function
