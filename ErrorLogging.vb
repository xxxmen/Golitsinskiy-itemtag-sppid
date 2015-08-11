Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module ErrorLogging
	
	'***********************************************************************
	'Copyright (c) 1998-2003, Intergraph Corporation
	'
	'Module:
	'   ErrorLogging
	'
	'
	'Abstract:
	'   Error Logging routines
	'
	'Description:
	'   Call LogAndRaiseError, LogError, and DisplayError to provide error logging.
	'
	'***********************************************************************
	
	Private m_strDescription As String
	
	
	Public Const ERROR_PROCESS_ABORTED As Short = 1067
	
	' error code from RegQueryValueEx
	Public Const ERROR_NONE As Short = 0
	Public Const ERROR_BADDB As Short = 1
	Public Const ERROR_BADKEY As Short = 2
	Public Const ERROR_CANTOPEN As Short = 3
	Public Const ERROR_CANTREAD As Short = 4
	Public Const ERROR_CANTWRITE As Short = 5
	Public Const ERROR_OUTOFMEMORY As Short = 6
	Public Const ERROR_ARENA_TRASHED As Short = 7
	Public Const ERROR_ACCESS_DENIED As Short = 8
	Public Const ERROR_INVALID_PARAMETERS As Short = 87
	Public Const ERROR_NO_MORE_ITEMS As Short = 259
	
	Private Enum ReportingLevelConstants
		igReportingLevelOff = 0
		igReportingLevelLog = 1
		igReportingLevelDisplay = 2
		igReportingLevelLogAndDisplay = 3
	End Enum
	' ------------- Error Logging Public Subroutines -------------
	
	Public Sub LogError(ByRef ErrorTag As String, Optional ByRef ErrorNumber As Object = Nothing)
		Dim eNumber As Integer
        If IsNothing(ErrorNumber) Then
            eNumber = Err.Number
        Else
            eNumber = ErrorNumber
        End If
		Dim objErrorLog As Object
		objErrorLog = CreateObject("igrErrorLogging412.ErrorLog")
		
        objErrorLog.ReportingLevel = 1 'ReportingLevelConstants.igReportingLevelLog
        objErrorLog.LogError("", ErrorTag, , , , , , eNumber)
        objErrorLog.ReportingLevel = 0 'ReportingLevelConstants.igReportingLevelOff
		
        objErrorLog = Nothing
	End Sub
	
	Public Function DisplayError(ByRef ErrorTag As String, ByRef ErrorNumber As Object, Optional ByRef Buttons As MsgBoxStyle = MsgBoxStyle.OKOnly) As MsgBoxResult
		Dim eNumber As Integer
        If IsNothing(ErrorNumber) Then
            eNumber = Err.Number
        Else
            eNumber = ErrorNumber
        End If
		
		Dim objErrorLog As Object
		objErrorLog = CreateObject("igrErrorLogging412.ErrorLog")
		
        objErrorLog.ReportingLevel = 2 'ReportingLevelConstants.igReportingLevelDisplay
        DisplayError = objErrorLog.LogError("", ErrorTag, , , , , Buttons, eNumber)
        objErrorLog.ReportingLevel = 0 'ReportingLevelConstants.igReportingLevelOff
		
        objErrorLog = Nothing
	End Function
	
	
	Public Sub LogAndRaiseError(ByRef ErrorTag As String, Optional ByRef ErrorNumber As Object = Nothing)
		Dim eNumber As Integer
		
		'   the following was added so that the original error and description are preserved / propagated
		'   to the top of the stack
        If IsNothing(ErrorNumber) Then
            If Err.Number = 0 Then 'first call on stack, app wants to raise an error
                Err.Number = vbObjectError ' SOME number, a system error has not occurred
                m_strDescription = ErrorTag
            Else 'err.number has already been set , preserve the description
                m_strDescription = Err.Description
            End If
        Else ' if app is setting the error number, I assume its the first call on stack
            If ErrorNumber > 512 Then
                Err.Number = vbObjectError + ErrorNumber
            Else
                Err.Number = vbObjectError
            End If
            m_strDescription = ErrorTag
        End If
		
		eNumber = Err.Number
		
		
		Dim objErrorLog As Object
		objErrorLog = CreateObject("igrErrorLogging412.ErrorLog")
		
        objErrorLog.ReportingLevel = 1 'ReportingLevelConstants.igReportingLevelLog
        objErrorLog.LogError(ErrorTag, m_strDescription, nErrorNumberForLogFile:=eNumber, bRaise:=True)
        objErrorLog.ReportingLevel = 0 'ReportingLevelConstants.igReportingLevelOff
		
        objErrorLog = Nothing
	End Sub
	
	
	Public Sub WriteToErrorLog(ByRef strMsg As String, Optional ByRef bPrintTime As Boolean = True, Optional ByRef ShowMsgBox As Boolean = False)
		
		Dim objErrorLog As Object
		objErrorLog = CreateObject("igrErrorLogging412.ErrorLog")
		If bPrintTime = True Then
            objErrorLog.LogMsgWithTimeStamp(strMsg)
		Else
            objErrorLog.LogMsg(strMsg)
		End If
		
	End Sub
	
	
	Public Function DisplayTime(ByRef strQualifier As String, Optional ByRef fPreviousTime As Single = 0#) As Single
		Dim fDelta As Single
		Dim fCurrentTime As Single
		
		fCurrentTime = VB.Timer()
		If fPreviousTime <> 0# Then
			fDelta = fCurrentTime - fPreviousTime
			WriteToErrorLog(strQualifier & vbTab & ";Time: " & CStr(fCurrentTime) & vbTab & ";Delta: " & CStr(fDelta), False)
		Else
			WriteToErrorLog(strQualifier & vbTab & ";Time: " & CStr(fCurrentTime), False)
		End If
		DisplayTime = fCurrentTime
	End Function
End Module