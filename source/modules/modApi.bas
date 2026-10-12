Attribute VB_Name = "modApi"
'---------------------------------------------------------------------------------------
' Package: modApi
'---------------------------------------------------------------------------------------
'
' Automation API of the wizard (without UI, result as JSON)
'
' Author:
'     Josef Poetzl
'
' Remarks:
'---------------------------------------------------------------------------------------
'| Call (e.g. via Application.Run or msaccess-vcs-mcp vcs_call_vba):
'|     Application.Run "<add-in path without .accda>.API", "Resolve", "data/SqlTools.cls", "mode=missing"
'|
'| Commands:
'|     GetVersion
'|     GetConfig      [Options]
'|     ListInstalled
'|     Resolve        Files, [Options]   (read only)
'|     Import         Files, [Options]
'|
'| Files:   "data/SqlTools.cls;form/filter/FilterControlManager.cls"
'| Options: "mode=missing|all|selected;tests=0|1;examples=0|1;source=config|local|github;repo=owner/name@branch"
'|
'| Result:  {"ok":true,...} or {"ok":false,"error":{"number":..,"source":..,"description":..}}
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const API_VERSION As Long = 1
Private Const ERR_INVALID_ARGUMENT As Long = vbObjectError + 1001
Private Const ERR_UNKNOWN_COMMAND As Long = vbObjectError + 1003

'---------------------------------------------------------------------------------------
' Function: API
'---------------------------------------------------------------------------------------
'
' Dispatcher for automation commands
'
' Parameters:
'     Command - GetVersion, GetConfig, ListInstalled, Resolve, Import
'     Args    - command arguments (strings)
'
' Returns:
'     JSON string
'
' Remarks:
'     Errors are returned as JSON (ok:false) at this API boundary, no MsgBox.
'
'---------------------------------------------------------------------------------------
Public Function API(ByVal Command As String, ParamArray Args() As Variant) As String

   Dim ArgList As Variant
   Dim SavedErrorHandlerMode As ACLibErrorHandlerMode
   Dim ErrorHandlerModeChanged As Boolean
   Dim Session As ACLibApiImportSession
   Dim AppVersion As String

On Error GoTo HandleErr

   ArgList = Args

   ' First access initializes the application (InitConfig sets the default error handler mode)
   AppVersion = CurrentApplication.Version

   SavedErrorHandlerMode = modErrorHandler.DefaultErrorHandlerMode
   modErrorHandler.DefaultErrorHandlerMode = ACLibErrorHandlerMode.aclibErrRaise
   ErrorHandlerModeChanged = True

   Select Case Trim$(Command)
      Case "GetVersion"
         API = "{""ok"":true,""version"":" & JsonString(AppVersion) & ",""api"":" & API_VERSION & "}"

      Case "GetConfig"
         Set Session = New ACLibApiImportSession
         Session.Configure GetArg(ArgList, 0)
         API = Session.GetConfig

      Case "ListInstalled"
         Set Session = New ACLibApiImportSession
         API = Session.ListInstalled

      Case "Resolve"
         Set Session = New ACLibApiImportSession
         Session.Configure GetArg(ArgList, 1)
         API = Session.Resolve(GetRequiredArg(ArgList, 0, "Files"))

      Case "Import"
         Set Session = New ACLibApiImportSession
         Session.Configure GetArg(ArgList, 1)
         API = Session.Import(GetRequiredArg(ArgList, 0, "Files"))

      Case Else
         Err.Raise ERR_UNKNOWN_COMMAND, "modApi.API", "Unknown command '" & Command & "' (GetVersion, GetConfig, ListInstalled, Resolve, Import)"

   End Select

ExitHere:
   Set Session = Nothing ' Class_Terminate restores wizard state
   If ErrorHandlerModeChanged Then
      modErrorHandler.DefaultErrorHandlerMode = SavedErrorHandlerMode
   End If
   Exit Function

HandleErr:
   API = ErrorResult(Err.Number, Err.Source, Err.Description)
   Resume ExitHere

End Function

Private Function GetArg(ByRef ArgList As Variant, ByVal Index As Long) As String
   If UBound(ArgList) >= Index Then
      If Not IsMissing(ArgList(Index)) Then
         GetArg = Nz(ArgList(Index), vbNullString)
      End If
   End If
End Function

Private Function GetRequiredArg(ByRef ArgList As Variant, ByVal Index As Long, ByVal ArgName As String) As String
   GetRequiredArg = GetArg(ArgList, Index)
   If Len(Trim$(GetRequiredArg)) = 0 Then
      Err.Raise ERR_INVALID_ARGUMENT, "modApi.API", "Argument '" & ArgName & "' is missing"
   End If
End Function

Private Function ErrorResult(ByVal ErrNumber As Long, ByVal ErrSource As String, ByVal ErrDescription As String) As String
   ErrorResult = "{""ok"":false,""error"":{" & _
                 """number"":" & ErrNumber & "," & _
                 """source"":" & JsonString(ErrSource) & "," & _
                 """description"":" & JsonString(ErrDescription) & "}}"
End Function

Private Function JsonString(ByVal Value As String) As String
' minimal escaping for the result envelope (full JSON writer: ACLibApiImportSession)

   Dim Result As String
   Dim CharCode As Long
   Dim i As Long

   For i = 1 To Len(Value)
      CharCode = AscW(Mid$(Value, i, 1)) And &HFFFF&
      Select Case CharCode
         Case 34, 92 ' " \
            Result = Result & "\" & Mid$(Value, i, 1)
         Case Is < 32
            Result = Result & "\u" & Right$("000" & Hex$(CharCode), 4)
         Case Else
            Result = Result & Mid$(Value, i, 1)
      End Select
   Next

   JsonString = """" & Result & """"

End Function
