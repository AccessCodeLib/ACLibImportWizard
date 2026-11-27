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

Option Compare Text
Option Explicit

Private Const FilterButtonCaption As String = "Filter"
Private Const FilterButtonPendingCaption As String = "Filter *"
Private Const NothingSelectedMessage As String = "Nothing has been selected."
Private Const MessageTitle As String = "ACLib Import Wizard"

Private m_TreeModel As RepositoryTreeModel
Private WithEvents m_TreeRows As Form_ACLibRepositoryTreeRowsForm
Attribute m_TreeRows.VB_VarHelpID = -1
Private m_AppliedFilterText As String
Private m_IsTreeShown As Boolean
Private m_IsFilterFieldFocused As Boolean

Public Event FilesSelected(ByRef FileArray() As String)
Public Event Unload()

Private Sub Form_Load()
   InitTreeForm
End Sub

Private Sub Form_Timer()

   Dim CloseAfterError As Boolean

On Error GoTo HandleErr

   Me.TimerInterval = 0
   LoadRepositoryTree

ExitHere:
   If CloseAfterError Then
      DoCmd.Close acForm, Me.Name
   End If
   Exit Sub

HandleErr:
   DoCmd.Hourglass False
   ' an error before the tree is shown for the first time closes the form (the wizard stays usable)
   CloseAfterError = Not m_IsTreeShown
   Select Case HandleError(Err.Number, "Form_Timer", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ReleaseTreeAndRaiseUnload
End Sub

Private Sub txtFilter_Change()
   UpdateFilterButtonState
End Sub

Private Sub txtFilter_GotFocus()
   TrackFilterFieldFocus True
End Sub

Private Sub txtFilter_LostFocus()
   TrackFilterFieldFocus False
End Sub

Private Sub txtFilter_AfterUpdate()

On Error GoTo HandleErr

   ApplyFilter

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "txtFilter_AfterUpdate", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub cmdFilter_Click()

On Error GoTo HandleErr

   ApplyFilter

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdFilter_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub cmdCommit_Click()
   SendSelectedFiles
End Sub

Private Sub cmdCancel_Click()
   DoCmd.Close
End Sub

Private Sub m_TreeRows_ExpandClicked(ByVal Key As String)

On Error GoTo HandleErr

   ToggleItemExpanded Key

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "m_TreeRows_ExpandClicked", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub m_TreeRows_CheckClicked(ByVal Key As String)

On Error GoTo HandleErr

   ToggleItemChecked Key

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "m_TreeRows_CheckClicked", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub InitTreeForm()
   Set m_TreeRows = Me.subRepositoryTree.Form
   ' show the form before the first (slow) GitHub call
   Me.TimerInterval = 1
End Sub

Private Sub ReleaseTreeAndRaiseUnload()
   Set m_TreeRows = Nothing
   Set m_TreeModel = Nothing
   RaiseEvent Unload
End Sub

Private Sub LoadRepositoryTree()

   Dim TreeModel As RepositoryTreeModel

   DoCmd.Hourglass True

   Set TreeModel = New RepositoryTreeModel
   TreeModel.Initialize New RepositoryTreeSource
   TreeModel.FilterText = GetFilterText
   m_AppliedFilterText = TreeModel.FilterText
   Set m_TreeModel = TreeModel

   ShowTree
   KeepCaretInFilterField
   UpdateFilterButtonState
   Me.subRepositoryTree.Visible = True
   Me.labStartInfo.Visible = False
   m_IsTreeShown = True

   DoCmd.Hourglass False

End Sub

' The text of the field is only valid while the field has the focus (.Text), afterwards .Value is used.
' Only reads: the field is never assigned, focused or moved here.
Private Function GetFilterText() As String
   If m_IsFilterFieldFocused Then
      GetFilterText = Trim$(Me.txtFilter.Text)
   Else
      GetFilterText = Trim$(Nz(Me.txtFilter.Value, vbNullString))
   End If
End Function

' Binding the rows to the subform for the first time moves the caret of the text box that has the focus
' (measured: Access still reports txtFilter as active control, the caret appears at the start of the first row).
' Setting SelStart puts it back; the text is not changed.
Private Sub KeepCaretInFilterField()
   If m_IsFilterFieldFocused Then
      Me.txtFilter.SelStart = Len(Me.txtFilter.Text)
   End If
End Sub

' Enter in the filter field presses the Filter button; elsewhere the form has no default button
Private Sub TrackFilterFieldFocus(ByVal HasFocus As Boolean)
   m_IsFilterFieldFocused = HasFocus
   Me.cmdFilter.Default = HasFocus
End Sub

' The button shows that the entered text is not applied yet (bold, marked caption)
Private Sub UpdateFilterButtonState()

   Dim IsFilterPending As Boolean

   IsFilterPending = StrComp(GetFilterText, m_AppliedFilterText, vbBinaryCompare) <> 0

   Me.cmdFilter.FontBold = IsFilterPending
   If IsFilterPending Then
      Me.cmdFilter.Caption = FilterButtonPendingCaption
   Else
      Me.cmdFilter.Caption = FilterButtonCaption
   End If

End Sub

' Enter or leaving the field (AfterUpdate) and the Filter button both end up here:
' the tree is rebuilt only if the text differs from the text that is already applied.
' Before the tree is shown nothing happens: loading the tree applies the text of the field.
Private Sub ApplyFilter()

   Dim FilterText As String

   If Not m_IsTreeShown Then
      Exit Sub
   End If

   FilterText = GetFilterText
   If StrComp(FilterText, m_AppliedFilterText, vbBinaryCompare) = 0 Then
      Exit Sub
   End If

   m_TreeModel.FilterText = FilterText
   m_AppliedFilterText = FilterText
   ShowTree
   UpdateFilterButtonState

End Sub

Private Sub ToggleItemExpanded(ByVal Key As String)
   m_TreeModel.ToggleExpanded Key
   ShowTree Key
End Sub

Private Sub ToggleItemChecked(ByVal Key As String)
   m_TreeModel.ToggleChecked Key
   ShowTree Key
End Sub

Private Sub ShowTree(Optional ByVal CurrentKey As String = vbNullString)
   m_TreeRows.ShowItems m_TreeModel.GetVisibleItems, CurrentKey
   ShowSelectionInfo m_TreeModel.CheckedFileCount
End Sub

Private Sub ShowSelectionInfo(ByVal SelectedFileCount As Long)
   If SelectedFileCount = 1 Then
      Me.labSelectionInfo.Caption = "1 file selected"
   Else
      Me.labSelectionInfo.Caption = SelectedFileCount & " files selected"
   End If
End Sub

Private Sub SendSelectedFiles()

   Dim Files() As String

   If GetCheckedFileCount() = 0 Then
      MsgBox NothingSelectedMessage, vbInformation, MessageTitle
      Exit Sub
   End If

   Files = GetCheckedFileArray()

   ' last statement: the wizard releases this form instance in its event handler
   RaiseEvent FilesSelected(Files)

End Sub

Private Function GetCheckedFileCount() As Long
   If Not (m_TreeModel Is Nothing) Then
      GetCheckedFileCount = m_TreeModel.CheckedFileCount
   End If
End Function

Private Function GetCheckedFileArray() As String()

   Dim CheckedKeys As Collection
   Dim Files() As String
   Dim Index As Long

   Set CheckedKeys = m_TreeModel.GetCheckedFileKeys
   ReDim Files(0 To CheckedKeys.Count - 1)
   For Index = 1 To CheckedKeys.Count
      Files(Index - 1) = CheckedKeys(Index)
   Next

   GetCheckedFileArray = Files

End Function
