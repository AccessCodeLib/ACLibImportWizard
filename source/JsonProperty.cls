VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_Name As String
Private m_Value As Variant
Private m_Properties As JsonProperties

Public Property Get Self() As JsonProperty
   Set Self = Me
End Property

Public Property Get Name() As String
   Name = m_Name
End Property

Friend Property Let Name(ByVal NewValue As String)
   m_Name = NewValue
End Property

Public Property Get Value() As Variant
''Attribute Value.VB_UserMemId = 0
   If IsObject(m_Value) Then
      Set Value = m_Value
   Else
      Value = m_Value
   End If
End Property

Public Property Let Value(ByVal NewValue As Variant)
   m_Value = NewValue
End Property

Public Property Set ObjectValue(ByVal NewValue As Object)
   Set m_Value = NewValue
End Property

Public Property Get Properties() As JsonProperties
   If m_Properties Is Nothing Then
      Set m_Properties = New JsonProperties
   End If
   Set Properties = m_Properties
End Property

Public Property Set Properties(ByVal NewRef As JsonProperties)
   Set m_Properties = NewRef
End Property
