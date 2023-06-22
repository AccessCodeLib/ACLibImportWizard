Attribute VB_Name = "_config_Application"
'---------------------------------------------------------------------------------------
' Modul: _initApplication
'---------------------------------------------------------------------------------------
'
' Application configuration
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/_config_Application.bas</file>
'  <replace>base/_config_Application.bas</replace> 'dieses Modul ersetzt base/_config_Application.bas
'  <license>_codelib/license.bas</license>
'  <use>%AppFolder%/source/defGlobal_ACLibImportWizard.bas</use>
'  <use>base/modApplication.bas</use>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/ApplicationHandler_AppFile.cls</use>
'  <use>base/modErrorHandler.bas</use>
'  <use>_codelib/addins/shared/ACLibConfiguration.cls</use>
'  <use>%AppFolder%/source/ACLibFileManager.cls</use>
'  <use>%AppFolder%/source/ACLibImportWizardForm.frm</use>
'  <use>usability/ApplicationHandler_DirTextbox.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'Versionsnummer
Private Const APPLICATION_VERSION As String = "1.3.0"

#Const USE_CLASS_ApplicationHandler_AppFile = 1
#Const USE_CLASS_ApplicationHandler_DirTextbox = 1

Public Const APPLICATION_NAME As String = "ACLib Import Wizard"
Private Const APPLICATION_FULLNAME As String = "Access Code Library - Import Wizard"
Private Const APPLICATION_TITLE As String = APPLICATION_FULLNAME
Private Const APPLICATION_ICONFILE As String = "ACLib.ico"
Private Const APPLICATION_STARTFORMNAME As String = "ACLibImportWizardForm"

Private Const DEFAULT_ERRORHANDLERMODE As Long = ACLibErrorHandlerMode.aclibErrMsgBox

Private m_Extensions As ApplicationHandler_ExtensionCollection

'---------------------------------------------------------------------------------------
' Sub: InitConfig
'---------------------------------------------------------------------------------------
'
' Init application configuration
'
' Parameters:
'     CurrentAppHandlerRef - Possibility of a reference transfer so that CurrentApplication does not have to be used</param>
'
'---------------------------------------------------------------------------------------
Public Sub InitConfig(Optional ByRef CurrentAppHandlerRef As ApplicationHandler = Nothing)

On Error GoTo HandleErr

'----------------------------------------------------------------------------
' Error handler
'

   modErrorHandler.DefaultErrorHandlerMode = DEFAULT_ERRORHANDLERMODE

   
'----------------------------------------------------------------------------
' Set global variables
'
   defGlobal_ACLibImportWizard.ACLibIconFileName = APPLICATION_ICONFILE

'----------------------------------------------------------------------------
' Application instance
'
   If CurrentAppHandlerRef Is Nothing Then
      Set CurrentAppHandlerRef = CurrentApplication
   End If

   With CurrentAppHandlerRef
   
      'To be on the safe side, set AccDb
      Set .AppDb = CodeDb 'must point to CodeDb,
                          'as this application is used as an add-in
   
      ''Application name
      .ApplicationName = APPLICATION_NAME
      .ApplicationFullName = APPLICATION_FULLNAME
      .ApplicationTitle = APPLICATION_TITLE
      
      'Version
      .Version = APPLICATION_VERSION
      
      'Form called at the end of CurrentApplication.Start
      .ApplicationStartFormName = APPLICATION_STARTFORMNAME
   
   End With

'----------------------------------------------------------------------------
' Extensions:
'
   Set m_Extensions = New ApplicationHandler_ExtensionCollection
   With m_Extensions
      Set .ApplicationHandler = CurrentAppHandlerRef
      
#If USE_CLASS_ApplicationHandler_AppFile = 1 Then
      .Add New ApplicationHandler_AppFile
#End If

#If USE_CLASS_ApplicationHandler_DirTextbox = 1 Then
      .Add New ApplicationHandler_DirTextbox
#End If
      
      .Add New ACLibConfiguration
      .Add New ACLibFileManager
   End With

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "InitConfig", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

'############################################################################
'
' Functions for application maintenance
' (only needed in the application design)
'
'----------------------------------------------------------------------------
' Auxiliary function for saving files to the local AppFile table
'----------------------------------------------------------------------------
Private Sub SetAppFiles()

   Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & APPLICATION_ICONFILE)

End Sub
