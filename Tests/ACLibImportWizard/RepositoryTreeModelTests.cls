VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryTreeModelTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

'TestSuite.AddByClassName("RepositoryTreeModelTests").Run

Private Const ErrFakeSource As Long = vbObjectError + 42
Private Const ErrUnknownItemKey As Long = vbObjectError + 6001
Private Const ErrNotInitialized As Long = vbObjectError + 6002

Private Const InitialVisibleKeys As String = "README.md|base|data|text"

Private m_Source As RepositoryTreeSourceFake
Private m_Model As RepositoryTreeModel

Public Sub Setup()
   Set m_Source = New RepositoryTreeSourceFake
   Set m_Model = New RepositoryTreeModel
End Sub

Public Sub Teardown()
   Set m_Model = Nothing
   Set m_Source = Nothing
End Sub

'--- M01

Public Sub Initialize_FakeTree_RootItemsVisibleAndOneSourceCall()

   m_Model.Initialize m_Source

   Assert.That VisibleKeys, Iz.EqualTo(InitialVisibleKeys)
   Assert.That m_Source.ReadAllCallCount, Iz.EqualTo(1)

End Sub

'--- M02

Public Sub ToggleExpanded_Folder_ExpandsAndCollapsesWithoutSourceCall()

   InitModel

   m_Model.ToggleExpanded "data"
   Assert.That VisibleKeys, Iz.EqualTo("README.md|base|data|data/SqlTools.cls|data/dao|text")

   m_Model.ToggleExpanded "data"
   Assert.That VisibleKeys, Iz.EqualTo(InitialVisibleKeys)

   m_Model.ToggleExpanded "data"
   Assert.That VisibleKeys, Iz.EqualTo("README.md|base|data|data/SqlTools.cls|data/dao|text")

   Assert.That m_Source.ReadAllCallCount, Iz.EqualTo(1)

End Sub

'--- M04

Public Sub ToggleExpanded_File_NothingChanges()

   InitModel

   m_Model.ToggleExpanded "README.md"

   Assert.That VisibleKeys, Iz.EqualTo(InitialVisibleKeys)
   Assert.That m_Model.GetItem("README.md").IsExpanded, Iz.EqualTo(False)

End Sub

'--- M05

Public Sub ToggleExpanded_UnknownKey_RaisesUnknownItemKey()

   InitModel

   Assert.Throws ErrUnknownItemKey
   m_Model.ToggleExpanded "nope"

End Sub

Public Sub ToggleChecked_UnknownKey_RaisesUnknownItemKey()

   InitModel

   Assert.Throws ErrUnknownItemKey
   m_Model.ToggleChecked "nope"

End Sub

Public Sub GetItem_UnknownKey_RaisesUnknownItemKey()

   InitModel

   Assert.Throws ErrUnknownItemKey
   m_Model.GetItem "nope"

End Sub

'--- M06

Public Sub GetVisibleItems_NotInitialized_RaisesNotInitialized()

   Assert.Throws ErrNotInitialized
   m_Model.GetVisibleItems

End Sub

'--- M07

Public Sub ToggleChecked_FilesOfFolder_FolderStateFollowsFiles()

   InitModel
   m_Model.ToggleExpanded "base"

   m_Model.ToggleChecked "base/ApplicationHandler.cls"
   Assert.That CheckState("base"), Iz.EqualTo(CheckStatePartiallyChecked)

   m_Model.ToggleChecked "base/modErrorHandler.bas"
   Assert.That CheckState("base"), Iz.EqualTo(CheckStateChecked)

   m_Model.ToggleChecked "base/ApplicationHandler.cls"
   Assert.That CheckState("base"), Iz.EqualTo(CheckStatePartiallyChecked)

   m_Model.ToggleChecked "base/modErrorHandler.bas"
   Assert.That CheckState("base"), Iz.EqualTo(CheckStateUnchecked)

End Sub

'--- M08, M24

Public Sub ToggleChecked_Folder_ChecksWholeSubtree()

   InitModel

   m_Model.ToggleChecked "data"

   Assert.That CheckState("data"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckState("data/dao"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckedKeys, Iz.EqualTo("data/SqlTools.cls|data/dao/DaoHandler.cls")
   Assert.That m_Source.ReadAllCallCount, Iz.EqualTo(1)

End Sub

Public Sub CheckedFileCount_AfterFolderToggle_CountsFilesOnly()

   InitModel

   m_Model.ToggleChecked "data"

   Assert.That m_Model.CheckedFileCount, Iz.EqualTo(2)

End Sub

'--- M09

Public Sub ToggleChecked_CheckedFolderAgain_UnchecksSubtree()

   InitModel
   m_Model.ToggleChecked "data"

   m_Model.ToggleChecked "data"

   Assert.That CheckState("data"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckState("data/dao"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckState("data/SqlTools.cls"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckState("data/dao/DaoHandler.cls"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckedKeys, Iz.EqualTo(vbNullString)

End Sub

'--- M10

Public Sub ToggleChecked_EmptyFolder_KeepsOwnStateAndReturnsNoKey()

   InitModel

   m_Model.ToggleChecked "data/dao/empty"

   Assert.That CheckState("data/dao/empty"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckedKeys, Iz.EqualTo(vbNullString)
   Assert.That CheckState("data/dao"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckState("data"), Iz.EqualTo(CheckStateUnchecked)

   m_Model.ToggleChecked "data/dao/empty"

   Assert.That CheckState("data/dao/empty"), Iz.EqualTo(CheckStateUnchecked)

End Sub

'--- M11

Public Sub ToggleChecked_PartiallyCheckedFolder_ChecksAllFiles()

   InitModel
   m_Model.ToggleChecked "base/ApplicationHandler.cls"

   m_Model.ToggleChecked "base"

   Assert.That CheckState("base/ApplicationHandler.cls"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckState("base/modErrorHandler.bas"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckState("base"), Iz.EqualTo(CheckStateChecked)

End Sub

'--- M13

'AccUnit:Row("sql", "sql").Name("lower case")
'AccUnit:Row("SQL", "SQL").Name("upper case")
'AccUnit:Row("  sql ", "sql").Name("trimmed")
Public Sub FilterText_Value_ShowsMatchingFileAndAncestors(ByVal FilterInput As String, ByVal ExpectedFilterText As String)

   InitModel

   m_Model.FilterText = FilterInput

   Assert.That VisibleKeys, Iz.EqualTo("data|data/SqlTools.cls")
   Assert.That m_Model.GetItem("data").IsShownExpanded, Iz.EqualTo(True)
   Assert.That m_Model.GetItem("data").IsExpanded, Iz.EqualTo(False)
   Assert.That m_Model.FilterText, Iz.EqualTo(ExpectedFilterText)
   Assert.That m_Source.ReadAllCallCount, Iz.EqualTo(1)

End Sub

'--- M14

Public Sub FilterText_Cleared_RestoresPreviousView()

   InitModel
   m_Model.FilterText = "sql"

   m_Model.FilterText = vbNullString

   Assert.That VisibleKeys, Iz.EqualTo(InitialVisibleKeys)
   Assert.That m_Model.GetItem("data").IsShownExpanded, Iz.EqualTo(False)

End Sub

'--- M15

Public Sub ToggleChecked_FolderWithFilter_ChecksMatchingFilesOnly()

   InitModel
   m_Model.FilterText = "sql"

   m_Model.ToggleChecked "data"

   Assert.That CheckState("data/SqlTools.cls"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckState("data/dao/DaoHandler.cls"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckState("data"), Iz.EqualTo(CheckStatePartiallyChecked)

End Sub

'--- M15b

Public Sub ToggleChecked_FolderWithFilterTwice_UnchecksMatchingFiles()

   InitModel
   m_Model.FilterText = "sql"
   m_Model.ToggleChecked "data"

   m_Model.ToggleChecked "data"

   Assert.That CheckState("data/SqlTools.cls"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckState("data"), Iz.EqualTo(CheckStateUnchecked)

End Sub

'--- M15c

Public Sub ToggleChecked_FolderWithFilterAndHiddenCheckedFile_KeepsHiddenFile()

   InitModel
   m_Model.ToggleChecked "data/dao/DaoHandler.cls"
   m_Model.FilterText = "sql"

   m_Model.ToggleChecked "data"
   Assert.That CheckState("data/SqlTools.cls"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckState("data"), Iz.EqualTo(CheckStateChecked)

   m_Model.ToggleChecked "data"
   Assert.That CheckState("data/SqlTools.cls"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckState("data/dao/DaoHandler.cls"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckState("data"), Iz.EqualTo(CheckStatePartiallyChecked)

End Sub

'--- M16

Public Sub GetCheckedFileKeys_FileHiddenByFilter_StaysInResult()

   InitModel
   m_Model.ToggleChecked "README.md"

   m_Model.FilterText = "sql"

   Assert.That CheckedKeys, Iz.EqualTo("README.md")

End Sub

'--- M17

Public Sub FilterText_NoMatch_NothingVisible()

   InitModel

   m_Model.FilterText = "zzz"

   Assert.That VisibleKeys, Iz.EqualTo(vbNullString)

End Sub

'--- M18

Public Sub ToggleExpanded_WithFilter_NothingChanges()

   InitModel
   m_Model.FilterText = "sql"

   m_Model.ToggleExpanded "data"

   Assert.That VisibleKeys, Iz.EqualTo("data|data/SqlTools.cls")
   Assert.That m_Model.GetItem("data").IsExpanded, Iz.EqualTo(False)

End Sub

'--- M19

Public Sub Initialize_SourceRaisesError_PassesErrorOn()

   m_Source.RaiseOnReadAll = True

   Assert.Throws ErrFakeSource
   m_Model.Initialize m_Source

End Sub

Public Sub GetVisibleItems_AfterFailedInitialize_RaisesNotInitialized()

   m_Source.RaiseOnReadAll = True
   Assert.That InitializeErrorNumber, Iz.EqualTo(ErrFakeSource)

   Assert.Throws ErrNotInitialized
   m_Model.GetVisibleItems

End Sub

'--- M23

Public Sub GetItem_NestedFile_HasLevelAndParentKey()

   InitModel

   With m_Model.GetItem("data/dao/DaoHandler.cls")
      Assert.That .Level, Iz.EqualTo(2)
      Assert.That .ParentKey, Iz.EqualTo("data/dao")
   End With

End Sub

Public Sub GetItem_HiddenFolderChild_RaisesUnknownItemKey()

   InitModel

   Assert.Throws ErrUnknownItemKey
   m_Model.GetItem "_tools/x.cls"

End Sub

'--- M25

Public Sub FilterText_FileInSubFolder_ShowsAllAncestors()

   InitModel

   m_Model.FilterText = "dao"

   Assert.That VisibleKeys, Iz.EqualTo("data|data/dao|data/dao/DaoHandler.cls")

End Sub

Public Sub FilterText_FolderNameOnly_NothingVisible()

   InitModel

   m_Model.FilterText = "empty"

   Assert.That VisibleKeys, Iz.EqualTo(vbNullString)

End Sub

'--- M26

Public Sub HasChildren_FileEmptyFolderAndFolder_FalseFalseTrue()

   InitModel

   Assert.That m_Model.GetItem("README.md").HasChildren, Iz.EqualTo(False)
   Assert.That m_Model.GetItem("data/dao/empty").HasChildren, Iz.EqualTo(False)
   Assert.That m_Model.GetItem("data").HasChildren, Iz.EqualTo(True)

End Sub

'--- M27

Public Sub Initialize_SecondCallFails_KeepsPreviousState()

   InitModel
   m_Model.ToggleChecked "README.md"
   m_Source.RaiseOnReadAll = True

   Assert.That InitializeErrorNumber, Iz.EqualTo(ErrFakeSource)

   Assert.That CheckState("README.md"), Iz.EqualTo(CheckStateChecked)
   Assert.That VisibleKeys, Iz.EqualTo(InitialVisibleKeys)

End Sub

Public Sub Initialize_SecondCallSucceeds_ResetsState()

   InitModel
   m_Model.ToggleChecked "README.md"
   m_Model.ToggleExpanded "data"
   m_Model.FilterText = "sql"

   m_Model.Initialize m_Source

   Assert.That CheckState("README.md"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That m_Model.GetItem("data").IsExpanded, Iz.EqualTo(False)
   Assert.That m_Model.FilterText, Iz.EqualTo(vbNullString)
   Assert.That VisibleKeys, Iz.EqualTo(InitialVisibleKeys)
   Assert.That m_Source.ReadAllCallCount, Iz.EqualTo(2)

End Sub

'--- M28

Public Sub ToggleChecked_FilesInSubFolders_EmptyFolderDoesNotCount()

   InitModel

   m_Model.ToggleChecked "data/dao/DaoHandler.cls"
   Assert.That CheckState("data/dao"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckState("data"), Iz.EqualTo(CheckStatePartiallyChecked)

   m_Model.ToggleChecked "data/SqlTools.cls"
   Assert.That CheckState("data"), Iz.EqualTo(CheckStateChecked)

End Sub

'--- M29

Public Sub GetItem_HiddenInfoFileInOtherCase_RaisesUnknownItemKey()

   InitModel

   Assert.Throws ErrUnknownItemKey
   m_Model.GetItem "text/_INFO.TXT"

End Sub

Public Sub ToggleExpanded_FolderWithHiddenInfoFile_ShowsOtherFileOnly()

   InitModel

   m_Model.ToggleExpanded "text"

   Assert.That VisibleKeys, Iz.EqualTo("README.md|base|data|text|text/StringTools.bas")

End Sub

'--- M30

Public Sub ToggleChecked_PartiallyCheckedFolderTwice_ChecksThenUnchecksAllFiles()

   InitModel
   m_Model.ToggleChecked "data/SqlTools.cls"
   Assert.That CheckState("data"), Iz.EqualTo(CheckStatePartiallyChecked)

   m_Model.ToggleChecked "data"
   Assert.That CheckState("data/SqlTools.cls"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckState("data/dao/DaoHandler.cls"), Iz.EqualTo(CheckStateChecked)
   Assert.That CheckState("data"), Iz.EqualTo(CheckStateChecked)

   m_Model.ToggleChecked "data"
   Assert.That CheckState("data/SqlTools.cls"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckState("data/dao/DaoHandler.cls"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckState("data"), Iz.EqualTo(CheckStateUnchecked)

End Sub

'--- M31

Public Sub ReadAllCallCount_AfterAllOperations_StaysOne()

   InitModel

   m_Model.ToggleExpanded "data"
   m_Model.ToggleChecked "data"
   m_Model.FilterText = "sql"
   m_Model.ToggleChecked "data"
   m_Model.FilterText = vbNullString
   m_Model.GetVisibleItems
   m_Model.GetCheckedFileKeys
   Assert.That m_Model.CheckedFileCount, Iz.GreaterThanOrEqualTo(0)

   Assert.That m_Source.ReadAllCallCount, Iz.EqualTo(1)

End Sub

'--- additional cases from review S2

Public Sub Initialize_ItemsWithoutParentFolder_AreIgnored()

   m_Source.UseOrphanItems = True

   m_Model.Initialize m_Source

   Assert.That VisibleKeys, Iz.EqualTo("README.md|ok")

End Sub

Public Sub GetItem_ItemWithMissingParent_RaisesUnknownItemKey()

   m_Source.UseOrphanItems = True
   m_Model.Initialize m_Source

   Assert.Throws ErrUnknownItemKey
   m_Model.GetItem "missing/file.txt"

End Sub

Public Sub GetItem_ItemWithFileAsParent_RaisesUnknownItemKey()

   m_Source.UseOrphanItems = True
   m_Model.Initialize m_Source

   Assert.Throws ErrUnknownItemKey
   m_Model.GetItem "README.md/child.txt"

End Sub

Public Sub ToggleExpanded_NotInitialized_RaisesNotInitialized()

   Assert.Throws ErrNotInitialized
   m_Model.ToggleExpanded "data"

End Sub

Public Sub ToggleChecked_NotInitialized_RaisesNotInitialized()

   Assert.Throws ErrNotInitialized
   m_Model.ToggleChecked "data"

End Sub

Public Sub GetItem_NotInitialized_RaisesNotInitialized()

   Assert.Throws ErrNotInitialized
   m_Model.GetItem "data"

End Sub

Public Sub GetCheckedFileKeys_NotInitialized_RaisesNotInitialized()

   Assert.Throws ErrNotInitialized
   m_Model.GetCheckedFileKeys

End Sub

Public Sub CheckedFileCount_NotInitialized_RaisesNotInitialized()

   Dim FileCount As Long

   Assert.Throws ErrNotInitialized
   FileCount = m_Model.CheckedFileCount

End Sub

Public Sub FilterTextGet_NotInitialized_RaisesNotInitialized()

   Dim CurrentFilter As String

   Assert.Throws ErrNotInitialized
   CurrentFilter = m_Model.FilterText

End Sub

Public Sub FilterTextLet_NotInitialized_RaisesNotInitialized()

   Assert.Throws ErrNotInitialized
   m_Model.FilterText = "sql"

End Sub

Public Sub ToggleChecked_FolderWithoutMatchingFileUnderFilter_HasNoEffect()

   InitModel
   m_Model.FilterText = "sql"

   m_Model.ToggleChecked "base"

   Assert.That CheckState("base"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckState("base/ApplicationHandler.cls"), Iz.EqualTo(CheckStateUnchecked)
   Assert.That CheckedKeys, Iz.EqualTo(vbNullString)

End Sub

Public Sub IsShownExpanded_FileUnderFilter_IsFalse()

   InitModel
   m_Model.FilterText = "sql"

   m_Model.GetVisibleItems

   Assert.That m_Model.GetItem("data/SqlTools.cls").IsShownExpanded, Iz.EqualTo(False)

End Sub

'--- helpers (private: not tests)

Private Sub InitModel()
   m_Model.Initialize m_Source
End Sub

' Returns the error number of Initialize; the number is checked by the caller right away.
Private Function InitializeErrorNumber() As Long

   On Error Resume Next
   m_Model.Initialize m_Source
   InitializeErrorNumber = Err.Number
   On Error GoTo 0

End Function

Private Function CheckState(ByVal Key As String) As Long
   CheckState = m_Model.GetItem(Key).CheckState
End Function

Private Function VisibleKeys() As String
   VisibleKeys = JoinKeys(m_Model.GetVisibleItems)
End Function

Private Function CheckedKeys() As String
   CheckedKeys = JoinKeys(m_Model.GetCheckedFileKeys)
End Function

' Joins the keys of a collection of RepositoryTreeItem or of String with "|".
Private Function JoinKeys(ByVal Keys As Collection) As String

   Dim Result As String
   Dim Index As Long
   Dim Element As Variant
   Dim Key As String

   For Index = 1 To Keys.Count
      If IsObject(Keys(Index)) Then
         Set Element = Keys(Index)
         Key = Element.Key
      Else
         Key = Keys(Index)
      End If
      If Index > 1 Then Result = Result & "|"
      Result = Result & Key
   Next

   JoinKeys = Result

End Function
