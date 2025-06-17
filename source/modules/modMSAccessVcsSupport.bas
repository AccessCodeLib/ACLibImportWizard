Attribute VB_Name = "modMSAccessVcsSupport"
Option Compare Database
Option Explicit

Public Sub VcsRunBeforeExport()

   Const ACLib_ConfigTableName As String = "ACLib_ConfigTable"

   If TableDefExists(ACLib_ConfigTableName) Then
      CurrentDb.TableDefs.Delete ACLib_ConfigTableName
      DBEngine.Idle dbRefreshCache
   End If

   CurrentApplication.SetApplicationProperty "StartUpForm", vbNullString

End Sub
