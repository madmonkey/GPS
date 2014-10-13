Attribute VB_Name = "basUtilitiesErrorHandling"
Option Explicit

'* Note: Every public error handling/logging method saves the state of the global
'* Err object at the beginning of the method and restores the state at the end
'* of the method.
'* Exceptions: UEH_RaiseError, UEH_EndLogging

'* BAM 7/14-16/03 Major Changes:
'* 1) Err object is global, no need to pass it via parameters to ANY method.  Changed
'*    any such methods to accept the ErrObject parameter optionally, and not use it at
'*    all so one day it may be removed from the parameter list.
'* 2) Added structured error handling to all public methods.  Error handling will write
'*    any internal errors to a text file.
'* 3) Made sure that the state of the global Err object will be the same entering any
'*    public method as when it exits the method
'* 4) Added Specific method wrappers for the UEH_Log method for:
'*      A) Starting of processors
'*      B) Receiving messsages
'*      C) Dumping message contents.

Private Type ErrObject_Holder
    Caller As String
    
    Description As String
    HelpContext As String
    HelpFile As String
    LastDllError As Long
    Number As Long
    Source As String
End Type

Private Const sCLASS_Name As String = "basUtilitiesErrorHandling"

Private Const sDEFAULT_LOG_FILE_PATH As String = "C:\htelog.log"

Private Const sDEFAULT_PROFILE As String = "MSW"

'* Property instance variables
Private m_sDateFormat As String
Private m_sLogFileDelimiter As String
Private m_eLevelMask As HTE_LOGGGER_TLB.HTE_LogLevel

Private m_oLogger As Object

'* Used to Keep Err state between beginning and end of methods
Private m_udLastError As ErrObject_Holder

Public Property Let UEH_DateFormat(vData As String)
'!BC********************************************************************************
' Description:
' ##BD  Format to store date fields in log file.
'
' Parameters:
' ##PD vData Date format desired.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'07/15/2003     NLJ     Initial Version
'!EC********************************************************************************
    m_sDateFormat = vData
    If Not m_oLogger Is Nothing Then m_oLogger.DateFormat = vData
End Property

Public Property Get UEH_DateFormat() As String
Attribute UEH_DateFormat.VB_Description = "Format to store date fields in log file."
'!BC********************************************************************************
' Description:
' ##BD  Format to store date fields in log file.
'
' Returns:
' ##RD Date format desired.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'07/15/2003     NLJ     Initial Version
'!EC********************************************************************************
    UEH_DateFormat = m_sDateFormat
End Property

Public Property Let UEH_Delimiter(vData As String)
'!BC********************************************************************************
' Description:
' ##BD  Log file field delimiter.
'
' Parameters:
' ##PD vData Delimiter to use for log file fields.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'07/15/2003     NLJ     Initial Version
'!EC********************************************************************************
    m_sLogFileDelimiter = vData
    If Not m_oLogger Is Nothing Then m_oLogger.Delimiter = vData
End Property

Public Property Get UEH_Delimiter() As String
Attribute UEH_Delimiter.VB_Description = "Log file field delimiter."
'!BC********************************************************************************
' Description:
' ##BD  Log file field delimiter.
'
' Returns:
' ##RD Delimiter to use for log file fields.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'07/15/2003     NLJ     Initial Version
'!EC********************************************************************************
    UEH_Delimiter = m_sLogFileDelimiter
End Property

Public Sub UEH_ImmediateLogging(LogLevel As HTE_LOGGGER_TLB.HTE_LogLevel)
Attribute UEH_ImmediateLogging.VB_Description = " Sets the level for debug.print statements.   Log functions will perform debug.print for each level that is passed.  For example: UEH_ImmediateLogging(logVerbose or logError) will debug.print error and verbose level statements."
'!BC********************************************************************************
' Description:
' ##BD  Sets the level for debug.print statements.   Log functions will perform _
debug.print for each level that is passed.  For example: _
UEH_ImmediateLogging(logVerbose or logError) will debug.print error and _
verbose level statements.
'
' Parameters:
' ##PD LogLevel Level bit mask for log levels you wish to display to the intermediate window.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'03/03/1999     NLJ     Initial verion
'06/30/2003     BAM     Changed LogLevel type to TLB enumeration
'!EC********************************************************************************
    m_eLevelMask = LogLevel
End Sub

Public Function UEH_GetVariableLogText(var As Variant) As String

    Select Case VarType(var)

        Case Is >= vbArray

            UEH_GetVariableLogText = "<array(" & UBound(var) & ")>"

        Case vbObject

            If var Is Nothing Then

                UEH_GetVariableLogText = "<nothing>"

            Else 'Not var Is Nothing

                UEH_GetVariableLogText = "<object>"

            End If 'var Is Nothing

        Case vbString, vbDate

            UEH_GetVariableLogText = """" & CStr(var) & """"

        Case vbUserDefinedType

            UEH_GetVariableLogText = "UDT"

        Case Else 'Any other variable type

            UEH_GetVariableLogText = CStr(var) & ""

    End Select

End Function

Public Function UEH_GetParameterString(ParamArray vParameters()) As String
Attribute UEH_GetParameterString.VB_Description = "Used to build a formated string of parameters that are used in a method call.  Allows the logging of parameters passed to methods that experienced errors"
'!BC********************************************************************************
' Description:
' ##BD  Used to build a formated string of parameters that are used in a method _
call.  Allows the logging of parameters passed to methods that experienced _
errors
'
' Parameters:
' ##PD vParameters List of variables you wish to concatenate values into one string.
'
' Returns:
' ##RD All values passed in through the parameter list concatenated into one string.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'03/03/1999     NLJ     Initial Version
'07/15/2003     BAM     Added error handling and supporting of more types
'!EC********************************************************************************

    Const sMETHOD_Name As String = "UEH_GetParameterString"
    
    SaveErr sMETHOD_Name

    On Error GoTo ErrorHandler
    
    Dim sReturnValue As String
    sReturnValue = ""

    Dim l_lParametersUpperBound As Long
    l_lParametersUpperBound = UBound(vParameters)

    If l_lParametersUpperBound = -1 Then GoTo Cleanup

    Dim lParamCounter As Long

    '* Add the value of each parameter to a comma-delimeted list
    For lParamCounter = 1 To l_lParametersUpperBound Step 2

        sReturnValue = sReturnValue & vParameters(lParamCounter - 1) & " = " & _
            UEH_GetVariableLogText(vParameters(lParamCounter)) & "; "

    Next lParamCounter

    sReturnValue = Left$(sReturnValue, Len(sReturnValue) - 2)

Cleanup:

    UEH_GetParameterString = sReturnValue
    
    RestoreErr sMETHOD_Name
    
    Exit Function

ErrorHandler:

    LogInternalError sMETHOD_Name

    On Error Resume Next

    Resume Cleanup

End Function 'UEH_GetParameterString(...)

Public Function UEH_BeginLogging(Optional sProfile As String, Optional sProjectName As String) As String
Attribute UEH_BeginLogging.VB_Description = "Global function to start error logging.  Profile is key in registry that contains initialization information."
'!BC********************************************************************************
' Description:
' ##BD  Global function to start error logging.  Profile is key in registry that _
contains initialization information.
'
' Parameters:
' ##PD sProfile Registry profile to initialize Logger DLL from
' ##PD sProjectName Current project Name (Deprecated, not needed.)
'
'  When         Who     What
'--------------------------------------------------------------------------------
'02/11/1999     NLJ     Initial Version
'!EC********************************************************************************

    Const sMETHOD_Name As String = "UEH_BeginLogging"
    
    SaveErr sMETHOD_Name
    
    On Error GoTo HTE_Logger_InstantiateError

    If m_oLogger Is Nothing Then Set m_oLogger = CreateObject("HTE_Logger.Log")
    
    On Error GoTo ErrorHandler
    
    If LenB(sProfile) = 0 Then sProfile = sDEFAULT_PROFILE
    
    m_oLogger.DateFormat = UEH_DateFormat
    m_oLogger.Delimiter = UEH_Delimiter
    m_oLogger.InitializeProfile sProfile

Cleanup:
    
    RestoreErr sMETHOD_Name
    
    Exit Function

HTE_Logger_InstantiateError:

    '* Modify global Err object's source and description
    Err.Description = "Failed to intantiate HTE_Logger.Log object [" & _
        Err.Description & "], make sure HTE_Logger.DLL is registered!"

ErrorHandler:
    
    '* If we were unable to log the internal error, print to the immediate window
    '* that we were unable to.
    If Not LogInternalError(sMETHOD_Name) Then Debug.Print "Logger failed to open any log file"

    On Error Resume Next

    Resume Cleanup

End Function 'UEH_BeginLogging(...)

Public Function UEH_EndLogging() As String
Attribute UEH_EndLogging.VB_Description = "Global function to end error logging."
'!BC********************************************************************************
' Description:
' ##BD  Global function to end error logging.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'02/11/1999     NLJ     Initial Version
'!EC********************************************************************************

    Const sMETHOD_Name As String = "UEH_EndLogging"
    
    SaveErr sMETHOD_Name
    
    On Error GoTo ErrorHandler
    
    Set m_oLogger = Nothing

Cleanup:

    RestoreErr sMETHOD_Name

    Exit Function

ErrorHandler:

    LogInternalError sMETHOD_Name

    On Error Resume Next

    Resume Cleanup

End Function 'UEH_EndLogging()

Public Function UEH_GetErrorString(lErrorID As Long, ParamArray asStrings()) As String
Attribute UEH_GetErrorString.VB_Description = "Retrieves an error string from a resource file and replaces tagged values passed in."
'!BC********************************************************************************
' Description:
' ##BD  Retrieves an error string from a resource file and replaces tagged values _
passed in.  Example: UEH_GetErrorString(112, "Foo", "1", "Foo2", "2")
'
' Parameters:
' ##PD lErrorID Err.Number of error to look up.
' ##PD asStrings Parameter list of tokens and replacement values.
'
' Returns:
' ##RD String from resource file with values replaced from parameter list.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'02/11/1999     NLJ     Initial Version
'!EC********************************************************************************

    Const sMETHOD_Name As String = "UEH_GetErrorString"

    SaveErr sMETHOD_Name

    On Error GoTo ErrorHandler

    Dim l_sReturnValue As String
    l_sReturnValue = LoadResString(lErrorID)

    Dim l_i As Long

    For l_i = 1 To (UBound(asStrings) + 1) Step 2

        l_sReturnValue = Replace(l_sReturnValue, "<" & asStrings(l_i - 1) & ">", asStrings(l_i))

    Next l_i

Cleanup:

    UEH_GetErrorString = l_sReturnValue

    RestoreErr sMETHOD_Name

    Exit Function

ErrorHandler:

    LogInternalError sMETHOD_Name

    On Error Resume Next

    Resume Cleanup

End Function 'UEH_GetErrorString(...)

Public Sub UEH_Log(sClass As String, _
                   sMethod As String, _
                   Optional ByVal sMessage As String, _
                   Optional ByVal eLogLevel As HTE_LOGGGER_TLB.HTE_LogLevel = logVerbose, _
                   Optional ByVal lErrorID As Long, _
                   Optional ByVal sSource As String, _
                   Optional ByVal eSourceDataType As HTE_LOGGGER_TLB.HTE_LogSourceDataType = logString)
Attribute UEH_Log.VB_Description = "Insert an entry into the log file."
'!BC********************************************************************************
' Description:
' ##BD  Insert an entry into the log file.
'
' Parameters:
' ##PD sClass In what Class/Module/Form the entry originated.
' ##PD sMethod In what Sub/Function/Property the entry originated.
' ##PD sMessage Short text for the entry.
' ##PD eLogLevel At what level to log the entry.
' ##PD lErrorID Err.Number of any error associated with the entry.
' ##PD sSource Detailed text of the entry.
' ##PD eSourceDataType What format the detailed text for the entry is (text, binary, XML)
'
'  When         Who     What
'--------------------------------------------------------------------------------
'02/11/1999     NLJ     Initial Version
'06/30/2003     BAM     Changed LogLevel and SourceDataType to TLB types.  Changed
'                       hard-coded path to module-level constant.
'!EC********************************************************************************

    Const sMETHOD_Name As String = "UEH_Log"

    SaveErr sMETHOD_Name

    On Error GoTo ErrorHandler

    If eLogLevel = 0 Then eLogLevel = HTE_LOGGGER_TLB.logVerbose

    '* Output log entry to immediate window if we have the log level mask set
    If eLogLevel And m_eLevelMask Then Debug.Print sClass & "." & sMethod & " : " & sMessage

    '* Output log entry to m_oLogger DLL if available or emergency plain text file
    If Not m_oLogger Is Nothing Then

        Dim l_sFormatedMessage As String '* Remove extra white space from message
        l_sFormatedMessage = Trim$(Replace$(sMessage, vbNewLine, " "))

        Dim l_sFormattedSource As String '* Replace vbCrLf's in text/XML source
        l_sFormattedSource = IIf(eSourceDataType = logBinary, sSource, Replace$(sSource, vbNewLine, " "))

        m_oLogger.Log eLogLevel, App.Title, sClass, sMethod, l_sFormatedMessage, lErrorID, l_sFormattedSource, eSourceDataType
    
    Else
        
        App.LogEvent sClass & vbTab & sMethod & vbTab & sMessage, vbLogEventTypeWarning
    
    End If

Cleanup:

    RestoreErr sMETHOD_Name

    Exit Sub

ErrorHandler:

    LogInternalError sMETHOD_Name

    On Error Resume Next

    Resume Cleanup

End Sub 'UEH_Log(...)

Public Sub UEH_LogPropertyChange( _
                ByVal sClass As String, _
                ByVal sProperty As String, _
                ByVal sOldValue As String, _
                ByVal sNewValue As String)
Attribute UEH_LogPropertyChange.VB_Description = "Records an objects property change to the log."
'!BC********************************************************************************
' Description:
' ##BD  Records an objects property change to the log.
'
' Parameters:
' ##PD sClass Class the property is in.
' ##PD sMethod Property Name.
' ##PD sOldValue Old member value.
' ##PD sNewValue New member value.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'07/22/2003     BAM     Initial Version
'!EC********************************************************************************

    Const sMETHOD_Name As String = "UEH_LogPropertyChange"

    On Error GoTo ErrorHandler
    
    UEH_Log sClass, sProperty, "Member variable change through property", logDebug, , _
        "Old Value [" & sOldValue & "] => [" & sNewValue & "]", logString

Cleanup:

    Exit Sub

ErrorHandler:

    LogInternalError sMETHOD_Name

    On Error Resume Next

    Resume Cleanup

End Sub

Public Sub UEH_LogProcessorStartup( _
                ByVal sClass As String, _
                ByVal sMethod As String, _
                ByVal sProcessorShortName As String, _
                ByVal sAppExtension As String, _
                ParamArray aProperties())
'!BC********************************************************************************
' Description:
' ##BD  Enters log data for the startup of a processor.
'
' Parameters:
' ##PD sClass Class Name of the processor.
' ##PD sMethod Method the startup occurred in.
' ##PD sProcessorShortName Short english Name of the processor (e.g., "TFF")
' ##PD sAppExtension Extension of the processor application, usually "OCX"
' ##PD aProperties Name/Value pairs of properties and values to log (e.g., "Segment Size", 500)
'
'  When         Who     What
'--------------------------------------------------------------------------------
'07/16/2003     BAM     Initial Version
'!EC********************************************************************************

    Const sMETHOD_Name As String = "UEH_LogProcessorStartup"

    SaveErr sMETHOD_Name

    On Error GoTo ErrorHandler

    UEH_Log sClass, sMethod, sProcessorShortName & " Processor successfully initialized.", logInformation
    UEH_Log sClass, sMethod, "Version = " & App.Major & "." & App.Minor & "." & App.Revision, logInformation
    UEH_Log sClass, sMethod, "Path = " & App.Path & "\" & App.EXEName & "." & sAppExtension, logInformation
    
    Dim iPropertyCounter As Long
    
    '* Log each property Name/value tuple
    For iPropertyCounter = 0 To UBound(aProperties) Step 2
        
        UEH_Log sClass, sMethod, "Processor Paramter [" & _
            aProperties(iPropertyCounter) & "] = " & UEH_GetVariableLogText(aProperties(iPropertyCounter + 1)), logVerbose
    
    Next iPropertyCounter
    
Cleanup:

    RestoreErr sMETHOD_Name

    Exit Sub

ErrorHandler:

    UEH_LogError sCLASS_Name, sMETHOD_Name

    On Error Resume Next

    Resume Cleanup

End Sub 'UEH_LogProcessorStartup(...)

Public Sub UEH_LogError(sObject As String, _
                        sMethod As String, _
                        Optional oE As ErrObject, _
                        Optional sSource As String, _
                        Optional eSourceDataType As HTE_LOGGGER_TLB.HTE_LogSourceDataType)
Attribute UEH_LogError.VB_Description = "Formats a log entry based on the Err object and logError log level."
'!BC********************************************************************************
' Description:
' ##BD  Formats a log entry based on the Err object and logError log level.
'
' Parameters:
' ##PD sObject Class the error occurred in.
' ##PD sMethod Method the error occurred in.
' ##PD oE ErrObject to pull error from (deprecated, just uses global Err object)
' ##PD sSource Extra information you wish to log with the error.
' ##PD eSourceDataType What format the extra information is in (text, binary, XML)
'
'  When         Who     What
'--------------------------------------------------------------------------------
'02/11/1999     ADM     Initial Version
'06/30/2003     BAM     Changed SourceDataType to TLB type.
'!EC********************************************************************************

    Const sMETHOD_Name As String = "UEH_LogError"

    SaveErr sMETHOD_Name

    On Error GoTo ErrorHandler

    If m_udLastError.Number = 0 Then Exit Sub
    
    Dim l_sFormattedNumber As String, l_sMsg As String
    l_sFormattedNumber = IIf(m_udLastError.Number < 0, "0x" & Hex$(m_udLastError.Number), m_udLastError.Number)

    l_sMsg = vbCrLf & "An Exception Occured: " & vbCrLf & _
        "Number: " & l_sFormattedNumber & vbCrLf & _
        "Source: " & m_udLastError.Source & vbCrLf & _
        "Description: " & m_udLastError.Description

    UEH_Log sObject, sMethod, l_sMsg, HTE_LogLevel.logError, m_udLastError.Number, sSource, eSourceDataType

Cleanup:

    RestoreErr sMETHOD_Name

    Exit Sub

ErrorHandler:

    LogInternalError sMETHOD_Name

    On Error Resume Next

    Resume Cleanup

End Sub 'UEH_LogError(...)

Public Sub UEH_RaiseError(Optional oE As ErrObject)
Attribute UEH_RaiseError.VB_Description = "Helper function to re-raise an error that occured"
'!BC********************************************************************************
' Description:
' ##BD  Helper function to re-raise an error that occurred
'
' Parameters:
' ##PD oE Err object to raise an error from. (Deprecated, do not need to send)
'
'  When         Who     What
'--------------------------------------------------------------------------------
'03/03/1999     ADM     Initial Version
'!EC********************************************************************************

    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub 'UEH_RaiseError(...)

Public Function UEH_ReportError(Optional oE As ErrObject, _
                                Optional sTitle As String, _
                                Optional bModule As Boolean = False, _
                                Optional bRecoverable As Boolean = False) As VbMsgBoxResult
Attribute UEH_ReportError.VB_Description = "Global functions used to display error messages"
'!BC********************************************************************************
' Description:
' ##BD  Global functions used to display error messages
'
' Parameters:
' ##PD oE Err object to display message from (deprecated, do not have to pass)
' ##PD sTitle Title for message box (Defaults to App.Title)
' ##PD bModule If the messagebox's caller is a method in a module.
' ##PD bRecoverable If the error may be ignored.
'
' Returns:
' ##RD Message box result from user input.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'02/11/1999     ADM     Initial Version
'06/30/2003     BAM     Added line continuations to function declaration for
'                       easier reading.  Added local variables so parameter
'                       variables were not modified.
'!EC********************************************************************************

    Const sMETHOD_Name As String = "UEH_ReportError"
    
    SaveErr sMETHOD_Name
    
    On Error GoTo ErrorHandler

    Dim l_sMessageBoxTitle As String
    l_sMessageBoxTitle = IIf(LenB(sTitle) = 0, App.Title, sTitle)

    Dim l_sHelpFile As String
    l_sHelpFile = IIf(LenB(Err.HelpFile) = 0, App.HelpFile, Err.HelpFile)

    Dim l_sErrorMessage As String
    l_sErrorMessage = "An Error occured " & vbNewLine & vbNewLine & _
        "Number: " & Err.Number & vbNewLine & _
        "Source: " & Err.Source & vbNewLine & _
        "Description: " & Err.Description & vbNewLine

    Dim l_lButtons As VbMsgBoxStyle
    l_lButtons = vbCritical + IIf(bRecoverable, vbAbortRetryIgnore, vbOKOnly)

    '* If we have a help file and context, put help button on message box
    If LenB(l_sHelpFile) > 0 And Err.HelpContext <> 0 Then l_lButtons = l_lButtons + vbMsgBoxHelpButton

Cleanup:

    UEH_ReportError = MsgBox(l_sErrorMessage, l_lButtons, l_sMessageBoxTitle, l_sHelpFile, Err.HelpContext)
    
    RestoreErr sMETHOD_Name

    Exit Function

ErrorHandler:

    LogInternalError sMETHOD_Name

    On Error Resume Next

    Resume Cleanup

End Function 'UEH_RaiseError(...)

Public Function UEH_WriteLine(sFileName As String, sLogString As String) As Boolean
Attribute UEH_WriteLine.VB_Description = "General purpose function to append a string to a file.  Used when creation and initialization of logger object fails"
'!BC********************************************************************************
' Description:
' ##BD  General purpose function to append a string to a file.  Used when creation _
and initialization of m_oLogger object fails, or for logging internal module errors.
'
' Parameters:
' ##PD sFileName File path to write to.
' ##PD sLogString Text to write to file.
'
' Returns:
' ##RD True if data was written to file, false otherwise.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'03/05/1999     ADM     Initial Version
'!EC********************************************************************************
    
    Const sMETHOD_Name As String = "UEH_WriteLine"
    
    SaveErr sMETHOD_Name
    
    '* Ignore any errors in this method and at the end return true/false
    '* if the log file was written to.
    On Error Resume Next
    
    Debug.Print sLogString
    
    Err.Clear
    
    Open sFileName For Append As #1
    Print #1, sLogString
    Close #1
    
    '* Return true if write to file was successful
    Dim bGoodWrite As Boolean
    bGoodWrite = (Err.Number = 0)
    
    RestoreErr sMETHOD_Name
    
    UEH_WriteLine = bGoodWrite
    
End Function 'UEH_WriteLine(...)

Public Function LogInternalError(ByVal sMethod As String) As Boolean

    Const sMETHOD_Name As String = "LogInternalError"
    
    SaveErr sMETHOD_Name
    
    '* Ignore any errors in this method and at the end return true/false
    '* if the log file was written to.
    On Error Resume Next
    
    RestoreErr sMETHOD_Name
    
    '* Format:
    '* Error in [EXE/DLL/OCX.CLASS.METHOD] ERR_NUM, ERR_SRC: ERR_DESC
    Dim l_sErrorEntry As String
    l_sErrorEntry = "Error in [" & App.Title & "." & sCLASS_Name & "." & sMethod & "] " & _
        Hex$(m_udLastError.Number) & ", " & m_udLastError.Source & ": " & m_udLastError.Description

    '* Example:
    '* Error in [HTE_MSW_PS_TFF.basUtilitiesErrorHandling.UEH_BeginLogging] 6DCAB85, ADO: Invalid SQL Query

    App.LogEvent l_sErrorEntry, vbLogEventTypeWarning
    
    RestoreErr sMETHOD_Name
    
    LogInternalError = True

End Function 'LogInternalError(...)

Public Sub SaveErr(ByVal sCaller As String)
Attribute SaveErr.VB_Description = "Saves the state of the global Err object to a user-defined type variable local to the Utilities Error Handling module."
'!BC********************************************************************************
' Description:
' ##BD  Saves the state of the global Err object to a user-defined type variable _
local to the Utilities Error Handling module.
'
' Parameters:
' ##PD sCaller Method Name to store.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'07/16/2003     BAM     Updated version
'!EC********************************************************************************

    '* Only save error if we are not currenly saving one
    If LenB(m_udLastError.Caller) > 0 Then Exit Sub
        
    m_udLastError.Caller = sCaller
    
    m_udLastError.Description = Err.Description
    m_udLastError.HelpContext = Err.HelpContext
    m_udLastError.HelpFile = Err.HelpFile
    m_udLastError.LastDllError = Err.LastDllError
    m_udLastError.Number = Err.Number
    m_udLastError.Source = Err.Source
    
End Sub 'SaveErr(...)

Public Sub RestoreErr(ByVal sCaller As String)
Attribute RestoreErr.VB_Description = "Restores the state of the global Err object from a user-defined type variable local to the Utilities Error Handling module."
'!BC********************************************************************************
' Description:
' ##BD  Restores the state of the global Err object from a user-defined type _
variable local to the Utilities Error Handling module.
'
' Parameters:
' ##PD sCaller Method Name to restore from.
'
'  When         Who     What
'--------------------------------------------------------------------------------
'07/16/2003     BAM     Updated version
'!EC********************************************************************************
   
    '* Only restore error if it is from the correct caller
    If m_udLastError.Caller <> sCaller Then Exit Sub
        
    m_udLastError.Caller = ""
    
    Err.Description = m_udLastError.Description
    Err.HelpContext = m_udLastError.HelpContext
    Err.HelpFile = m_udLastError.HelpFile
    '* LastDLLError is read-only property :( can't assign
    'Err.LastDllError = m_udLastError.LastDllError
    Err.Number = m_udLastError.Number
    Err.Source = m_udLastError.Source
    
End Sub 'RestoreErr(...)
