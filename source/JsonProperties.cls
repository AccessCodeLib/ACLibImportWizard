VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_Properties As Collection

Private Sub Class_Initialize()
   Set m_Properties = New Collection
End Sub

Public Function NewProperty(ByVal Name As String) As JsonProperty
   With New JsonProperty
      .Name = Name
      Append .Self
      Set NewProperty = .Self
   End With
End Function

Public Sub Append(ByVal Property As JsonProperty)
   m_Properties.Add Property, Property.Name
End Sub

Public Property Get Item(ByVal Index As Variant) As JsonProperty
Attribute Item.VB_UserMemId = 0
   Set Item = m_Properties.Item(Index)
End Property

Public Property Get Count() As Long
   Count = m_Properties.Count
End Property

Public Sub SetProperties(ByVal NewProperties As JsonProperties)
   
   Dim CurrentProp As JsonProperty
   Dim NewProp As JsonProperty
   Dim i As Long
   
   For i = 1 To NewProperties.Count
      Set NewProp = NewProperties.Item(i)
      Set CurrentProp = m_Properties.Item(NewProp.Name)
      CurrentProp.Value = NewProp.Value
      If NewProp.Properties.Count > 0 Then
         CurrentProp.Properties.SetProperties NewProp.Properties
      End If
   Next
   
End Sub
