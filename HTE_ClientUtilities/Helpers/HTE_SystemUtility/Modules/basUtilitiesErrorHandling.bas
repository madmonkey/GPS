Attribute VB_Name = "basUtilitiesErrorHandling"
Option Explicit

Public Enum UEH_ErrorRecovery
    erAbort = 1
    erRetry
    erIgnore
End Enum

Public Enum UEH_LogLevel
    logVerbose = 8
    logInformation = 4
    logWarning = 2
    logError = 1
End Enum

Public Enum UEH_LogSourceDataType
    LogString
    logBinary
    logXML
End Enum

Private mvarUEH_DateFormat As String
Private mvarUEH_Delimiter As String
Private mvarUEH_OutputLogLevelImediate As UEH_LogLevel
Private mvarProjectName As String
Private LevelMask As Long
Private Logger As Object

Private TopCaller As String
Private ErrNumber As Long
Private ErrSource As String
Private ErrDescription As String
Private ErrHelpFile As String
Private ErrHelpContext As String
Private ErrLastDllError As Long

Public Property Let UEH_DateFormat(vData As String)
    mvarUEH_DateFormat = vData
    If Not Logger Is Nothing Then
        Logger.DateFormat = vData
    End If
End Property

Public Property Get UEH_DateFormat() As String
    UEH_DateFormat = mvarUEH_DateFormat
End Property

Public Property Let UEH_Delimiter(vData As String)
    mvarUEH_Delimiter = vData
    If Not Logger Is Nothing Then
        Logger.Delimiter = vData
    End If
End Property

Public Property Get UEH_Delimiter() As String
    UEH_Delimiter = mvarUEH_Delimiter
End Property

'*******************************************************************
' Name:     UEH_ImmediateLogging
' Purpose:
'       Sets the level for debug.print statements.   Log functions will
'       perform debug.print for each level that is passed.  For example:
'       UEH_ImmediateLogging(logVerbose or logError) will debug.print error
'       and verbose level statements.
'
' Who       When                What
' ------------------------------------------------------------------
' NLJ       03/03/1999      Initial verion
'*******************************************************************
Public Sub UEH_ImmediateLogging(LogLevel As UEH_LogLevel)
    LevelMask = LogLevel
End Sub

'*******************************************************************
' Name:     UEH_GetParameterString
' Purpose:
'       Used to build a formated string of parameters that are used in a
'       method call.  Allows the logging of parameters passed to methods
'       that experienced errors
'
' Who       When                What
' ------------------------------------------------------------------
' NLJ       03/03/1999      Initial verion
'*******************************************************************
Public Function UEH_GetParameterString(ParamArray Strings()) As String
    Dim i As Long
    
    SaveErr "UEH_GetParameterString"
    
    If IsObject(Strings(0)) Then
        If Strings(0) Is Nothing Then
            UEH_GetParameterString = "<nothing>"
        Else
            UEH_GetParameterString = "<object>"
        End If
    Else
        If VarType(Strings(0)) = vbString Then
            UEH_GetParameterString = Chr(34) & CStr(Strings(0)) & Chr(34)
        Else
            UEH_GetParameterString = CStr(Strings(0))
        End If
    End If
    For i = 1 To (UBound(Strings))
        Select Case True
            Case IsObject(Strings(i))
                If Strings(i) Is Nothing Then
                    UEH_GetParameterString = UEH_GetParameterString & ", <object>"
                Else
                    UEH_GetParameterString = UEH_GetParameterString & ", <nothing>"
                End If
            Case VarType(Strings(i)) = vbString
                UEH_GetParameterString = UEH_GetParameterString & ", " & Chr(34) & CStr(Strings(i)) & Chr(34)
            Case Else
                UEH_GetParameterString = UEH_GetParameterString & ", " & CStr(Strings(i))
        End Select
    Next [i]
    
    RestoreErr "UEH_GetParameterString"
End Function

'*******************************************************************
' Name: UEH_BeginLogging
' Purpose:
'       Global function to start error logging.  Profile is key in
'       registry that contains initialization information.
'
' Who       When                What
' ------------------------------------------------------------------
' NLJ       02/11/1999      Initial verion
'*******************************************************************
Public Function UEH_BeginLogging(Optional Profile As String, Optional ProjectName As String) As String
    
    On Error GoTo errHandler
    If Logger Is Nothing Then Set Logger = CreateObject("HTE_Logger.Log")
    
    mvarProjectName = ProjectName
    Logger.DateFormat = UEH_DateFormat
    Logger.Delimiter = UEH_Delimiter
    Logger.InitializeProfile Profile
    Exit Function
    
errHandler:
    ' if err.number = -1 then we were unable to open log file
    If Err.Number <> -1 Then
        UEH_WriteLine "c:\htelog.log", "Failed to create initialze logging: " + Err.Description + _
            vbNewLine + "Make sure HTE_LOGGER is registered"
    Else
        Debug.Print "Logger failed to open any log file"
    End If
End Function

'*******************************************************************
' Name: UEH_EndLogging
' Purpose:
'       Global function to end error logging.
'
' Who       When                What
' ------------------------------------------------------------------
' NLJ       02/11/1999      Initial verion
'*******************************************************************
Public Function UEH_EndLogging() As String
    Set Logger = Nothing
End Function

'*******************************************************************
' Name: UEH_GetErrorString
' Purpose:
'       Global function to retrieve an error string from a resource file
'          and replace tagged values passed in.
'
' Example: UEH_GetErrorString(112, "Foo", "1", "Foo2", "2")
'
' Who       When                What
' ------------------------------------------------------------------
' NLJ       02/11/1999      Initial verion
'*******************************************************************
Public Function UEH_GetErrorString(ErrorID As Long, ParamArray Strings()) As String
    Dim i As Long
    
    SaveErr "UEH_GetErrorString"
    UEH_GetErrorString = LoadResString(ErrorID)
    For i = 1 To (UBound(Strings) + 1) Step 2
        UEH_GetErrorString = Replace(UEH_GetErrorString, "<" & Strings(i - 1) & ">", Strings(i))
    Next [i]
    RestoreErr "UEH_GetErrorString"
End Function

'*******************************************************************
' Name:     UEH_Log
' Purpose:
'       Log a message
'
' Who       When                What
' ------------------------------------------------------------------
' ADM       02/11/1999      Initial verion
'*******************************************************************
Public Sub UEH_Log(Object As String, _
                   method As String, _
                   Optional Message As String, _
                   Optional LogLevel As UEH_LogLevel = logVerbose, _
                   Optional ErrorID As Long, _
                   Optional ByVal Source As String, _
                   Optional SourceDataType As UEH_LogSourceDataType = LogString)
        
    SaveErr "UEH_Log"
    
    If LogLevel = 0 Then LogLevel = logError
    
    If LogLevel And LevelMask Then
        Debug.Print Object + "." + method + " : " + Message
    End If
    
    If Not Logger Is Nothing Then
        Message = Replace(Message, vbNewLine, " ")
        Message = Trim(Message)
        
        If (SourceDataType <> logBinary) Then
            Source = Replace(Source, vbNewLine, " ")
        End If
        
        Logger.Log LogLevel, mvarProjectName, Object, method, Message, ErrorID, Source, SourceDataType
    Else
        UEH_WriteLine "c:\htelog.log", Object + vbTab + method + vbTab + Message
    End If
    
    RestoreErr "UEH_Log"
End Sub

'*******************************************************************
' Name:     LogError
' Purpose:
'       Formats error message from err object
'
' Who       When                What
' ------------------------------------------------------------------
' ADM       02/11/1999      Initial verion
'*******************************************************************
Public Sub UEH_LogError(Object As String, _
                        method As String, _
                        e As ErrObject, _
                        Optional Source As String, _
                        Optional SourceDataType As UEH_LogSourceDataType)
    Dim lsMsg As String
    
    SaveErr "UEH_LogError"
    
    If e.Number = 0 Then Exit Sub
    
    lsMsg = vbCrLf + "An Exception Occured: " + vbCrLf
    lsMsg = lsMsg + "Source: " + e.Source + vbCrLf
    lsMsg = lsMsg + "Number: " + CStr(e.Number) + vbCrLf
    lsMsg = lsMsg + "Description: " + e.Description
    
    UEH_Log Object, method, lsMsg, logError, , Source, SourceDataType
    RestoreErr "UEH_LogError"
End Sub
'*******************************************************************
' Name:     UEH_RaiseError
' Purpose:
'       Helper function to re-raise an error that occured
'
' Who       When                What
' ------------------------------------------------------------------
' ADM       03/03/1999      Initial verion
'*******************************************************************
Public Sub UEH_RaiseError(e As ErrObject)
    If e.Number = 0 Then Exit Sub
    e.Raise e.Number, e.Source, e.Description, e.HelpFile, e.HelpContext
End Sub

'*******************************************************************
' Name:
' Purpose:
'       Global functions used to display error messages
'
' Who       When                What
' ------------------------------------------------------------------
' ADM       02/11/1999      Initial verion
'*******************************************************************
Public Function UEH_ReportError(e As ErrObject, Optional Title As String, Optional Module As Boolean, Optional Recoverable As Boolean) As UEH_ErrorRecovery
    Dim lsErr As String
    Dim llResponse As Long
    Dim llButtons As Long
    
    SaveErr "UEH_ReportError"
    
    lsErr = "An Error occured " & vbNewLine & vbNewLine
    If Module Then
        lsErr = lsErr & "Module: " & e.Source & vbNewLine
    End If
    lsErr = lsErr & "Description: " & e.Description & vbNewLine
             
    If Title = "" Then Title = App.Title
    If e.HelpFile = "" Then e.HelpFile = App.HelpFile
    
    If e.HelpFile <> "" And e.HelpContext <> 0 Then
        llButtons = llButtons + vbMsgBoxHelpButton
    End If
    
    If Recoverable Then
        llButtons = llButtons + vbAbortRetryIgnore
    Else
        llButtons = llButtons + vbOKOnly
    End If
    
    llResponse = MsgBox(lsErr, vbCritical + llButtons, Title, e.HelpFile, e.HelpContext)
    
    Select Case llResponse
        Case vbAbort
            UEH_ReportError = erAbort
        Case vbRetry
            UEH_ReportError = erRetry
        Case vbIgnore
            UEH_ReportError = erIgnore
    End Select
    RestoreErr "UEH_ReportError"
End Function


'*******************************************************************
' Name:     UEH_WriteLine
' Purpose:
'       General purpose function to append a string to a file.  Used when
'       creation and initialization of logger object fails
'
' Who       When                What
' ------------------------------------------------------------------
' ADM       03/05/1999      Initial verion
'*******************************************************************
Public Sub UEH_WriteLine(FileName As String, LogString As String)
        
    SaveErr "UEH_WriteLine"
    Debug.Print LogString
    
    On Error Resume Next
    Open FileName For Append As #1
    Print #1, LogString
    Close #1
    
    RestoreErr "UEH_WriteLine"
End Sub

Private Sub SaveErr(Caller As String)
    If TopCaller = "" Then
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        ErrHelpContext = Err.HelpContext
        ErrHelpFile = Err.HelpFile
        TopCaller = Caller
    End If
End Sub

Private Sub RestoreErr(Caller As String)
    If TopCaller = Caller Then
        Err.Number = ErrNumber
        Err.Source = ErrSource
        Err.Description = ErrDescription
        Err.HelpContext = ErrHelpContext
        Err.HelpFile = ErrHelpFile
        TopCaller = ""
    End If
End Sub
