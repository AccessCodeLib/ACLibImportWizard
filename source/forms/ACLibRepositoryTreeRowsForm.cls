Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: ACLibRepositoryTreeRowsForm
'---------------------------------------------------------------------------------------
'
' Continuous sub form: shows the visible rows of the repository tree and reports clicks
'
' Remarks:
'---------------------------------------------------------------------------------------
'| The rows are temporary data: they are held in a disconnected ADODB recordset (late bound,
'| built in memory) and bound via Set Me.Recordset; nothing is written to a table or file.
'| The form holds no tree logic: it shows the items it gets and raises events for clicks.
'---------------------------------------------------------------------------------------
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

Option Compare Text
Option Explicit

' ADO constants (late binding, the project has no ADODB reference)
Private Const AdoProgId As String = "ADODB.Recordset"
Private Const AdoUseClient As Long = 3
Private Const AdoOpenStatic As Long = 3
Private Const AdoLockBatchOptimistic As Long = 4
Private Const AdoStateClosed As Long = 0
' adLongVarWChar: the form cannot read the other text field types (adVarWChar: "Reserved Error")
Private Const AdoTypeText As Long = 203
Private Const AdoTextFieldSize As Long = 4096
' IsFolder is stored as text as well: a bound boolean field cannot be read from its text box ("Reserved Error")
Private Const FolderFlagYes As String = "1"
Private Const FolderFlagNo As String = "0"

' Code points of the font Segoe MDL2 Assets (Windows 10 and 11; also in Segoe Fluent Icons)
Private Const GlyphFontName As String = "Segoe MDL2 Assets"
Private Const GlyphChevronDown As Long = &HE70D&
Private Const GlyphChevronRight As Long = &HE76C&
' OpenFolderHorizontal: reads as a folder at small sizes (the upright glyph Folder, E8B7, looks like a page)
Private Const GlyphFolder As Long = &HED25&
Private Const GlyphPage As Long = &HE7C3&
Private Const GlyphCheckbox As Long = &HE739&
Private Const GlyphCheckboxComposite As Long = &HE73A&
Private Const GlyphCheckboxIndeterminate As Long = &HE73C&

' The icon font has no blank glyph of full width: an item without children gets a chevron in the
' background color of the label, so that the icons of files and folders stay aligned.
' A system color (high bit set) cannot be converted without an API call: white is used then.
Private Const FallbackHiddenGlyphColor As String = "#FFFFFF"
Private Const SystemColorFlag As Long = &H80000000
Private Const IndentSpacesPerLevel As Long = 4
Private Const HtmlSpace As String = "&nbsp;"
' Equal gaps between check box, chevron, icon and name (the chevron glyph touches the icon otherwise)
Private Const GlyphGapHtml As String = "&nbsp;&nbsp;"

Private m_RowsRecordset As Object
Private m_HiddenGlyphColor As String

Public Event CheckClicked(ByVal Key As String)
Public Event ExpandClicked(ByVal Key As String)

Private Sub Form_Load()
   InitRowsForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ReleaseRows
End Sub

Private Sub txtCheckGlyph_Click()
   RaiseCheckClicked
End Sub

Private Sub txtTreeLabel_Click()
   RaiseTreeLabelClicked
End Sub

'---------------------------------------------------------------------------------------
' Sub: ShowItems
'---------------------------------------------------------------------------------------
'
' Replaces all rows by the given items (display order) and moves to the row of CurrentKey
' and keeps the position of that row in the list (it does not jump to the top)
'
' Parameters:
'     Items       - Collection of RepositoryTreeItem in display order
'     CurrentKey  - (optional) key of the row that becomes the current record
'
'---------------------------------------------------------------------------------------
Public Sub ShowItems(ByVal Items As Collection, Optional ByVal CurrentKey As String = vbNullString)

   Dim NewRows As Object
   Dim OldRows As Object
   Dim CurrentRowIndex As Long
   Dim RowOffset As Long
   Dim ErrNumber As Long
   Dim ErrSource As String
   Dim ErrDescription As String

On Error GoTo HandleErr

   RowOffset = GetCurrentRowOffset

   Me.Painting = False

   ' The recordset is built offline: changes of a bound recordset update the form row by row (slow).
   Set NewRows = BuildRows(Items, CurrentKey, CurrentRowIndex)
   Set OldRows = m_RowsRecordset
   Set Me.Recordset = NewRows
   ' Take over the bound recordset before closing the old one, so that a failing Close
   ' can neither leave the form bound to a closed recordset nor let HandleErr close the bound one.
   Set m_RowsRecordset = NewRows
   Set NewRows = Nothing
   CloseRows OldRows

   MoveToRow CurrentRowIndex, RowOffset

   Me.Painting = True
   Exit Sub

HandleErr:
   ErrNumber = Err.Number
   ErrSource = Err.Source
   ErrDescription = Err.Description

   Me.Painting = True
   If Not (NewRows Is Nothing) Then
      CloseRows NewRows
   End If

   Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub

Private Sub InitRowsForm()
   m_HiddenGlyphColor = GetHtmlColor(Me.txtTreeLabel.BackColor)
End Sub

Private Function GetHtmlColor(ByVal AccessColor As Long) As String

   If (AccessColor And SystemColorFlag) <> 0 Then
      GetHtmlColor = FallbackHiddenGlyphColor
      Exit Function
   End If

   ' Access stores colors as &HBBGGRR
   GetHtmlColor = "#" & Right$("0" & Hex$(AccessColor And &HFF&), 2) & _
                        Right$("0" & Hex$((AccessColor \ &H100&) And &HFF&), 2) & _
                        Right$("0" & Hex$((AccessColor \ &H10000) And &HFF&), 2)

End Function

Private Sub ReleaseRows()

   Set Me.Recordset = Nothing
   CloseRows m_RowsRecordset
   Set m_RowsRecordset = Nothing

End Sub

Private Sub CloseRows(ByVal Rows As Object)

   If Not (Rows Is Nothing) Then
      If Rows.State <> AdoStateClosed Then
         Rows.Close
      End If
   End If

End Sub

' Builds the disconnected recordset in display order and returns the 1-based index of CurrentKey
' (0 if there is no such row)
Private Function BuildRows(ByVal Items As Collection, ByVal CurrentKey As String, _
                           ByRef CurrentRowIndex As Long) As Object

   Dim Rows As Object
   Dim Item As RepositoryTreeItem
   Dim RowIndex As Long

   Set Rows = CreateObject(AdoProgId)
   With Rows
      .CursorLocation = AdoUseClient
      .Fields.Append "ItemKey", AdoTypeText, AdoTextFieldSize
      .Fields.Append "IsFolder", AdoTypeText, AdoTextFieldSize
      .Fields.Append "CheckGlyph", AdoTypeText, AdoTextFieldSize
      .Fields.Append "TreeLabelHtml", AdoTypeText, AdoTextFieldSize
      .Open , , AdoOpenStatic, AdoLockBatchOptimistic

      For Each Item In Items
         RowIndex = RowIndex + 1
         .AddNew
         .Fields("ItemKey").Value = Item.Key
         .Fields("IsFolder").Value = IIf(Item.IsFolder, FolderFlagYes, FolderFlagNo)
         .Fields("CheckGlyph").Value = GetCheckGlyph(Item.CheckState)
         .Fields("TreeLabelHtml").Value = GetTreeLabelHtml(Item)
         .Update
         If StrComp(Item.Key, CurrentKey, vbBinaryCompare) = 0 Then
            CurrentRowIndex = RowIndex
         End If
      Next

      If .RecordCount > 0 Then
         .MoveFirst
      End If
   End With

   Set BuildRows = Rows

End Function

' Number of rows between the top of the list and the current row (0 if there is no current row)
Private Function GetCurrentRowOffset() As Long

   Dim RowHeight As Long

   If m_RowsRecordset Is Nothing Then
      Exit Function
   End If
   If m_RowsRecordset.RecordCount = 0 Then
      Exit Function
   End If

   RowHeight = Me.Section(acDetail).Height
   If RowHeight > 0 Then
      GetCurrentRowOffset = Me.CurrentSectionTop \ RowHeight
   End If

End Function

Private Sub MoveToRow(ByVal RowIndex As Long, ByVal RowOffset As Long)

   ' The key is always a visible item; if it is not found, the current record simply stays.
   If RowIndex < 1 Then
      Exit Sub
   End If

   Me.Recordset.MoveFirst
   Me.Recordset.Move RowIndex - 1
   KeepRowOffset RowOffset

End Sub

' Binding a new recordset scrolls the current row to the top of the list: scroll to the row that
' was at the top before, then back to the current row, so that it keeps its place.
' Two steps are needed because a single SelTop = CurrentRow puts the row at the top;
' the first SelTop makes the former top row visible, the second one moves the current
' row back to its old position (measured: row offset stays unchanged).
Private Sub KeepRowOffset(ByVal RowOffset As Long)

   Dim CurrentRow As Long

   CurrentRow = Me.CurrentRecord
   If CurrentRow > RowOffset Then
      Me.SelTop = CurrentRow - RowOffset
      Me.SelTop = CurrentRow
   End If

End Sub

Private Sub RaiseCheckClicked()
   RaiseEvent CheckClicked(Me.txtItemKey.Value)
End Sub

Private Sub RaiseTreeLabelClicked()

   If Me.txtIsFolder.Value = FolderFlagYes Then
      RaiseEvent ExpandClicked(Me.txtItemKey.Value)
   Else
      RaiseEvent CheckClicked(Me.txtItemKey.Value)
   End If

End Sub

Private Function GetCheckGlyph(ByVal CheckState As RepositoryTreeCheckState) As String

   Select Case CheckState
      Case RepositoryTreeCheckState.CheckStateChecked
         GetCheckGlyph = ChrW$(GlyphCheckboxComposite)
      Case RepositoryTreeCheckState.CheckStatePartiallyChecked
         GetCheckGlyph = ChrW$(GlyphCheckboxIndeterminate)
      Case Else
         GetCheckGlyph = ChrW$(GlyphCheckbox)
   End Select

End Function

Private Function GetTreeLabelHtml(ByVal Item As RepositoryTreeItem) As String

   GetTreeLabelHtml = GlyphGapHtml & GetIndentHtml(Item.Level) & _
                      GetExpanderHtml(Item) & GlyphGapHtml & _
                      GetGlyphHtml(GetItemIconGlyph(Item)) & _
                      GlyphGapHtml & HtmlEncode(Item.Name)

End Function

Private Function GetIndentHtml(ByVal Level As Long) As String

   Dim SpaceCount As Long

   SpaceCount = Level * IndentSpacesPerLevel
   If SpaceCount > 0 Then
      GetIndentHtml = Replace$(Space$(SpaceCount), " ", HtmlSpace)
   End If

End Function

Private Function GetExpanderHtml(ByVal Item As RepositoryTreeItem) As String

   If Not Item.HasChildren Then
      GetExpanderHtml = "<font face=""" & GlyphFontName & """ color=""" & m_HiddenGlyphColor & """>" & _
                        ChrW$(GlyphChevronRight) & "</font>"
   ElseIf Item.IsShownExpanded Then
      GetExpanderHtml = GetGlyphHtml(GlyphChevronDown)
   Else
      GetExpanderHtml = GetGlyphHtml(GlyphChevronRight)
   End If

End Function

Private Function GetItemIconGlyph(ByVal Item As RepositoryTreeItem) As Long

   If Item.IsFolder Then
      GetItemIconGlyph = GlyphFolder
   Else
      GetItemIconGlyph = GlyphPage
   End If

End Function

Private Function GetGlyphHtml(ByVal CodePoint As Long) As String
   GetGlyphHtml = "<font face=""" & GlyphFontName & """>" & ChrW$(CodePoint) & "</font>"
End Function

Private Function HtmlEncode(ByVal Text As String) As String

   Dim Result As String

   Result = Replace$(Text, "&", "&amp;")
   Result = Replace$(Result, "<", "&lt;")
   Result = Replace$(Result, ">", "&gt;")
   Result = Replace$(Result, """", "&quot;")

   HtmlEncode = Result

End Function
