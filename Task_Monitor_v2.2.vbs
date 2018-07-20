
' Task Monitor v.2.2
'
' by Alex Dragokas

Const AddTimeToLogName = true

Dim bWin2000: bWin2000 = isWin2000()
Dim oDict: set oDict = CreateObject("Scripting.Dictionary")

sCurFolder = left(WScript.ScriptFullname, instrrev(WScript.ScriptFullname, "\"))

if AddTimeToLogName then
	sTime = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "_" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now)
	
	sLogFile1 = sCurFolder & "Processes_" & sTime & ".log"
	sLogFile2 = sCurFolder & "Processes_" & sTime & ".csv"
else
	sLogFile1 = sCurFolder & "Processes.log"
	sLogFile2 = sCurFolder & "Processes.csv"
end if
sMarker = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%temp%") & "\Marker_Process_Watcher_Dragokas"
sComputer = "."

' �������� �� �������� ������ Win32_Process ������ 1 �������
Set objWMIService = GetObject("winmgmts:\\" & sComputer & "\root\cimv2")
Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
  ("SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'")

set oFSO = CreateObject("Scripting.FileSystemObject")

' ���� ��� ����������� �������, ������� ������ � ��������� ������� ��������
if CheckMarker("Exist") then CheckMarker("Delete"): WScript.Echo "���������� ��������.": WScript.Quit

sMessage = "������ �������� ���������� ��������/���������� ���������." & vbLf & "�������� ����������� � ����: " & sLogFile1 & " (.csv)" &_
  vbLf & vbLf & "����� ��������� ����������, ��������� ������ ��� ��� :)"

' ���� ������ ��� ���� �� ������������
if WScript.Arguments.Count = 0 then
    ' ���� ������ ������� ��� �����. �������
    if isAdminRights() then
        WScript.Echo sMessage
    else
        if Msgbox(sMessage & vbLf & vbLF & "��������� � ������� ��������������?" & vbCrLf & vbCrLf & "��� - ��������� � ������������� �������.", vbYesNo + VbInformation, "Task Monitor by Dragokas") = vbYes then
             ' ��������� ���� ��������������
             Elevate()
             WScript.Quit
        end if
    end if
end if

' ������� ����� ���������
set oTS1 = oFSO.CreateTextFile(sLogFile1, true) ' true - overwrite
set oTS2 = oFSO.CreateTextFile(sLogFile2, true) ' true - overwrite
oTS1.WriteLine "Event		Date-Time	PID	Process Name	Par.PID	Parent Proc.Name	Domain\Username		Command Line Arguments"
oTS1.WriteLine "-----		---------	----	------------	-------	----------------	---------------		----------------------"
oTS2.WriteLine "Event;Date-Time;Process ID;Process Name;Parent PID;Parent Process Name;User;Command Line Arguments"

CheckMarker("Create")

Delim_Pl = vbTab
Delim_CS = ";"

Do
    ' ������� �������
    Set objLatestEvent = colMonitoredEvents.NextEvent()
 
    set oProc = objLatestEvent.TargetInstance

    ' ��������� ��� �������
    Select Case objLatestEvent.Path_.Class
      Case "__InstanceCreationEvent"
        if not CheckMarker("Exist") then Exit Do
        ProcInfo oProc, "Created", txtPlain, txtCSV
    	oTS1.WriteLine "Created" & txtPlain
    	oTS2.WriteLine "Created" & txtCSV
      Case "__InstanceDeletionEvent"
        if not CheckMarker("Exist") then Exit Do
        ProcInfo oProc, "Deleted", txtPlain, txtCSV
    	oTS1.WriteLine "Deleted" & txtPlain
    	oTS2.WriteLine "Deleted" & txtCSV
    End Select
Loop

' ��������� ����-��������
oTS1.Close()
oTS2.Close()

set oProc = Nothing: set objLatestEvent = Nothing: set oFSO = Nothing: set oTS = Nothing: Set colMonitoredEvents = Nothing: Set objWMIService = Nothing: set oDict = Nothing

Sub ProcInfo(objProcess, EventType, txtPlain, txtCSV) ' ��������� ������� ������� ��������, ����������� ������������� ��������
  with objProcess
    ParentPID = .ParentProcessId
    if ParentPID <> 0 then
	    set oParentProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = '" & ParentPID & "'")
		if not oParentProcesses is Nothing then
			For each oParentProc in oParentProcesses
				ParentName = oParentProc.Name
			next
			set oParentProcesses = Nothing: set oParentProc = Nothing
		end if
		set oParentProc = Nothing
	end if

	txtPlain = Delim_Pl & Now & Delim_Pl & .ProcessId & Delim_Pl & .Name:      if len(.Name) < 8      then txtPlain = txtPlain & Delim_Pl
	txtPlain = txtPlain & Delim_Pl & .ParentProcessId & Delim_Pl & ParentName: if len(ParentName) < 8 then txtPlain = txtPlain & Delim_Pl
	
	txtCSV   = Delim_CS & Now & Delim_CS & .ProcessId & Delim_CS & .Name & Delim_CS & .ParentProcessId & Delim_CS & ParentName
	
	' ���� ������������ / ����� ���������� ��� ������� "������� ��������".
	if EventType <> "Deleted" then
		On Error Resume next
		objProcess.GetOwner strUserName, strUserDomain
		On Error Goto 0
		
		txtPlain = txtPlain & Delim_Pl & Delim_Pl & strUserDomain & "\" & strUserName
		txtCSV   = txtCSV   & Delim_CS            & strUserDomain & "\" & strUserName
		
		if len(strUserDomain & "\" & strUserName) < 16 then txtPlain = txtPlain & Delim_Pl
		
		if not oDict.Exists(.ProcessId) then
			oDict.Add .ProcessId, strUserDomain & "\" & strUserName
		else
			oDict(.ProcessId) = strUserDomain & "\" & strUserName
		end if
	else
		if oDict.Exists(.ProcessId) then
			txtPlain = txtPlain & Delim_Pl & Delim_Pl & oDict(.ProcessId)
			txtCSV   = txtCSV   & Delim_CS            & oDict(.ProcessId)
			
			if len(oDict(.ProcessId)) < 16 then txtPlain = txtPlain & Delim_Pl
		else
			txtPlain = txtPlain & Delim_Pl & Delim_Pl & Delim_Pl & Delim_Pl
			txtCSV   = txtCSV   & Delim_CS
		end if
	end if
	if bWin2000 then
		txtPlain = txtPlain & Delim_Pl & "unknown"
		txtCSV   = txtCSV   & Delim_CS & "unknown"
	else
		txtPlain = txtPlain & Delim_Pl & .CommandLine
		txtCSV   = txtCSV   & Delim_CS & .CommandLine
	end if
	
  end with
End sub

Function CheckMarker(Action) ' �������� �� �������� �������/���������� ��������
  Select case Action
    Case "Create"
      set oMarker = oFSO.CreateTextFile(sMarker, true)
      oMarker.Close()
      set oMarker = Nothing
    Case "Delete"
      On error resume next
      For N = 1 to 10
        oFSO.DeleteFile sMarker, true 'true - Read Only Force
        if Err = 0 then Exit For
        Err.Clear
        WScript.Sleep(100) '���� ������ �������� ������ �������, ���� 100 ��. � ��������
      next
      On error Goto 0
    Case "Exist"
      CheckMarker = oFSO.FileExists(sMarker)
  End Select
End Function

Sub Elevate()
  Const DQ = """"
  Set colOS = GetObject("winmgmts:\root\cimv2").ExecQuery("Select * from Win32_OperatingSystem")
  For Each oOS In colOS
    strOSLong = oOS.Version
  Next
  If instr(strOSLong, "6.") = 1 or instr(strOSLong, "10.") = 1 Then
    If Not isAdminRights Then
        Set oShellApp = CreateObject("Shell.Application")
        oShellApp.ShellExecute WScript.FullName, DQ & WScript.ScriptFullName & DQ & " " & DQ & "Twice" & DQ, "", "runas", 1
        WScript.Quit
    End If
  End If
  set oOS = Nothing: set colOS = Nothing: set oShellApp = Nothing
End Sub

Function isAdminRights()
    Const KQV = &H1, KSV = &H2, HKCU = &H80000001, HKLM = &H80000002
    Set oReg = GetObject("winmgmts:root\default:StdRegProv")
    strKey = "System\CurrentControlSet\Control\Session Manager"
    intErrNum = oReg.CheckAccess(HKLM, strKey, KQV + KSV, flagAccess)
    isAdminRights = flagAccess
    Set oReg = Nothing
End Function

Function isWin2000()
    Set colOS = GetObject("winmgmts:\root\cimv2").ExecQuery("Select * from Win32_OperatingSystem")
    For Each oOS In colOS
      if instr(oOS.Version, "5.0.") = 1 then isWin2000 = true
    Next
End Function