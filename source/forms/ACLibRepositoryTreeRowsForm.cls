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
'| The rows are stored in a private temp database (TempDbHandler) and bound via Set Me.Recordset;
'| no table of the add-in or of the host database is used.
'| The form holds no tree logic: it shows the items it gets and raises events for clicks.
'---------------------------------------------------------------------------------------
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

Option Compare Text
Option Explicit

Private Const RowsTableName As String = "RepositoryTreeRows"
Private Const RowsTableDdl As String = "CREATE TABLE " & RowsTableName & " (" & _
      "ItemKey TEXT(255) CONSTRAINT PK_" & RowsTableName & " PRIMARY KEY, " & _
      "SortIndex LONG, IsFolder BIT, CheckGlyph TEXT(1), TreeLabelHtml MEMO)"
Private Const RowsSelectSql As String = "SELECT ItemKey, SortIndex, IsFolder, CheckGlyph, TreeLabelHtml" & _
      " FROM " & RowsTableName & " ORDER BY SortIndex"

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

Private m_TempDb As TempDbHandler
Private m_RowsRecordset As DAO.Recordset
Private m_HiddenGlyphColor As String

Public Event CheckClicked(ByVal Key As String)
Public Event ExpandClicked(ByVal Key As String)

Private Sub Form_Load()
   InitRowsForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DisposeRowsDatabase
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
'
' Parameters:
'     Items       - Collection of RepositoryTreeItem in display order
'     CurrentKey  - (optional) key of the row that becomes the current record
'
'---------------------------------------------------------------------------------------
Public Sub ShowItems(ByVal Items As Collection, Optional ByVal CurrentKey As String = vbNullString)

   Dim TransactionStarted As Boolean
   Dim ErrNumber As Long
   Dim ErrSource As String
   Dim ErrDescription As String

On Error GoTo HandleErr

   Me.Painting = False

   DBEngine.Workspaces(0).BeginTrans
   TransactionStarted = True
   WriteRows Items
   DBEngine.Workspaces(0).CommitTrans
   TransactionStarted = False

   RefreshRows
   MoveToItem CurrentKey

   Me.Painting = True
   Exit Sub

HandleErr:
   ErrNumber = Err.Number
   ErrSource = Err.Source
   ErrDescription = Err.Description

   Me.Painting = True
   If TransactionStarted Then
      DBEngine.Workspaces(0).Rollback
   End If

   Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub

Private Sub InitRowsForm()
   m_HiddenGlyphColor = GetHtmlColor(Me.txtTreeLabel.BackColor)
   InitRowsDatabase
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

Private Sub InitRowsDatabase()

   Set m_TempDb = New TempDbHandler
   m_TempDb.CreateDatabase
   m_TempDb.CreateTable RowsTableName, RowsTableDdl

   Set m_RowsRecordset = m_TempDb.CurrentDatabase.OpenRecordset(RowsSelectSql, dbOpenDynaset)
   Set Me.Recordset = m_RowsRecordset

End Sub

Private Sub DisposeRowsDatabase()

   ' The form and the recordset must release the file first: Dispose deletes it under On Error Resume Next.
   Set Me.Recordset = Nothing

   If Not (m_RowsRecordset Is Nothing) Then
      m_RowsRecordset.Close
      Set m_RowsRecordset = Nothing
   End If

   If Not (m_TempDb Is Nothing) Then
      m_TempDb.Dispose
      Set m_TempDb = Nothing
   End If

End Sub

Private Sub WriteRows(ByVal Items As Collection)

   Dim Rows As DAO.Recordset
   Dim Item As RepositoryTreeItem
   Dim SortIndex As Long

   m_TempDb.CurrentDatabase.Execute "DELETE FROM " & RowsTableName, dbFailOnError

   Set Rows = m_TempDb.CurrentDatabase.OpenRecordset(RowsTableName, dbOpenDynaset, dbAppendOnly)
   For Each Item In Items
      SortIndex = SortIndex + 1
      Rows.AddNew
      Rows.Fields("ItemKey").Value = Item.Key
      Rows.Fields("SortIndex").Value = SortIndex
      Rows.Fields("IsFolder").Value = Item.IsFolder
      Rows.Fields("CheckGlyph").Value = GetCheckGlyph(Item.CheckState)
      Rows.Fields("TreeLabelHtml").Value = GetTreeLabelHtml(Item)
      Rows.Update
   Next
   Rows.Close
   Set Rows = Nothing

End Sub

Private Sub RefreshRows()
   Me.Requery
End Sub

Private Sub MoveToItem(ByVal Key As String)

   If Len(Key) = 0 Then
      Exit Sub
   End If

   ' The key is always a visible item; if it is not found, the current record simply stays.
   Me.Recordset.FindFirst "ItemKey = " & SqlTools.TextToSqlText(Key)

End Sub

Private Sub RaiseCheckClicked()
   RaiseEvent CheckClicked(Me!ItemKey)
End Sub

Private Sub RaiseTreeLabelClicked()

   If Me!IsFolder Then
      RaiseEvent ExpandClicked(Me!ItemKey)
   Else
      RaiseEvent CheckClicked(Me!ItemKey)
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

   GetTreeLabelHtml = GetIndentHtml(Item.Level) & _
                      GetExpanderHtml(Item) & _
                      GetGlyphHtml(GetItemIconGlyph(Item)) & _
                      HtmlSpace & HtmlEncode(Item.Name)

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
