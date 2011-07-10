' User in Groups? 
' "So, [SysAd], what users aren't a member of these [arbitrary number of] groups?" 
' This script answers that previously tedious-if-not-impossible-to-answer question. 
' 
' Copyright 2010 Harold "Waldo" Grunenwald (Harold.Grunenwald@gmail.com) 
' 
' While I retain copyright of this work, I do encourage it's use for internal 
' systems.  (I want to help my fellow geeks get their work done.  I do insist 
' that my copyright and attribution be preserved.  If this code is to be 
' incorporated as part of a commercial product, I require written approval to be 
' granted, and other conditions may apply.  
' 
' Identify all users in Domain.  If they are not in any of the groups presumably  
' read in from command line arguments 
 
' Groups are specified as command-line arguments by their Common Name.  Active 
' Directory is then searched for those names.  Search AD for those Groups.  The 
' script quits if any of the groups can't be found. 
' 
' Gathers all users from AD and then checks whether they are in at least one of 
' the groups specified.  If the particular user is not a member of any of the 
' groups, report in the Logfile 
 
On Error Resume Next            ' Sadly this line is required 
 
' ===Establish target group(s)=== 
If Wscript.Arguments.Count = 0 Then 
    Wscript.Echo vbCrLf & "Please provide a group name as an argument" 
    WScript.Quit 
Else 
    Dim arrGroups() 
    For i = 0 to (WScript.Arguments.Count - 1) 
        Redim Preserve arrGroups(i) 
        arrGroups(i) = WScript.Arguments(i) 
        'WScript.Echo arrGroups(i) 
        'fnGetGroupDN(arrGroups(i)) 
        arrGroups(i) = fnGetGroupDN(arrGroups(i)) 
        'WScript.Echo arrGroups(i) 
    Next 
End If 
 
 
dim datetime 
datetime = fnFixDateTime(now()) 
 
dim logfile 
logfile = ".\Group_Membership-" & datetime & ".log" 
 
'---Create & Open Logfile for Writing (Appending, really)...--- 
const ForAppending = 8 
Set objFSO = CreateObject("Scripting.FilesystemObject") 
Set objLogFile = objFSO.OpenTextFile(logfile, ForAppending, True) 
 
'===Boilerplate LDAP Search Code=== 
Set objRootDSE                            = GetObject("LDAP://rootDSE") 
strDomain                                = "LDAP://" & objRootDSE.Get("defaultNamingContext") 
Const ADS_SCOPE_SUBTREE                    = 2 
Set objConnection                        = CreateObject("ADODB.Connection") 
Set objCommand                            = CreateObject("ADODB.Command") 
objConnection.Provider                    = "ADsDSOObject" 
objConnection.Open "Active Directory Provider" 
Set objCommand.ActiveConnection            = objConnection 
objCommand.Properties("Page Size")        = 1000 
objCommand.Properties("Searchscope")    = ADS_SCOPE_SUBTREE 
 
strProperties = "Name,memberOf,distinguishedName" 
 
'---Get Users--- 
objCommand.CommandText = _ 
    "<"& strDomain &">;(objectCategory=User);" _ 
        & strProperties & ";Subtree" 
Set objRecordSet = objCommand.Execute 
objRecordSet.MoveFirst 
 
'---Counters--- 
userTotalCount    = 0 
userFailCount    = 0 
 
'---GO FORTH!--- 
Do Until objRecordSet.EOF 
    strUser        = objRecordSet.Fields("Name").Value 
    strUserDN    = objRecordSet.Fields("distinguishedName").Value 
    groupCount    = 0 
    userTotalCount = userTotalCount + 1 
     
    Set objUser = GetObject _ 
        ("LDAP://" & strUserDN) 
    For each objGroupDN in arrGroups 
        ' Bind to the group object. 
        Set objGroup = GetObject("LDAP://" & objGroupDN) 
        strGroup    = objGroup.name 
        strGroupDN    = objGroup.distinguishedName 
         
        If (objGroup.IsMember("LDAP://" & strUserDN) = True) Then 
            'Wscript.Echo "User " & strUser & " is a member of " & strGroup 
            groupCount = groupCount + 1 
        Else 
            'Wscript.Echo "User " & strUser & " is NOT a member of " & strGroup 
        End If 
         
    Next 
     
    if (groupCount = 0) then 
        'wscript.echo strUserDN 
        objLogFile.WriteLine(strUserDN) 
        userFailCount = userFailCount + 1 
    end if 
     
    objRecordSet.MoveNext        ' next user 
Loop 
 
WScript.Echo 
WScript.Echo "Total Users = " & VbTab & VbTab & VbTab & VbTab & userTotalCount 
WScript.Echo "Users not in any specified Group = " & VbTab & userFailCount 
objLogFile.WriteLine(VbCrLf & VbCrLf) 
objLogFile.WriteLine("Session Began at " & datetime) 
objLogFile.WriteLine("Total Users =" & VbTab & VbTab & VbTab & VbTab & userTotalCount) 
objLogFile.WriteLine("Users not in any specified Group =" & VbTab & userFailCount) 
For Each strGroup in arrGroups 
    strGroups = strGroups & VbCrLf & VbTab & strGroup 
Next 
objLogFile.WriteLine("Groups Specified: " & strGroups) 
 
 
'====================================== 
'====================================== 
'===Putting the "Fun" in "Functions"=== 
'====================================== 
'====================================== 
 
 
Function fnFixDateTime(dumbformat) 
    dim dDate,dMonth,dYear,dDay 
    dDate = dumbformat 
    dYear = year(dDate) 
    dMonth = Month(dDate) 
    dDay = day(dDate) 
     
    dHour = hour(dDate) 
    dMinute = minute(dDate) 
    dSecond = Second(dDate) 
     
    if len(dMonth) < 2 then  
        dMonth = "0"& dMonth  
    end if 
     
    if len(dDay) < 2 then 
        dDay = "0"& dDay 
    end if 
     
    if len(dHour) < 2 then 
        dHour = "0" & dHour 
    End If 
     
    if len(dMinute) < 2 then 
        dMinute = "0" & dMinute 
    End If 
     
    if len(dSecond) < 2 then 
        dSecond = "0" & dSecond 
    End If 
     
    dDate = dYear & dMonth & dDay & "_" & dHour & dMinute & dSecond 
    fnFixDateTime = dDate 
End Function 
 
 
'====================================== 
'====================================== 
 
 
Function fnGetGroupDN(strGroup) 
    '===Boilerplate LDAP Search Code=== 
    ' Yes, you do need this boilerplate code again.  The variables from the main 
    ' body of the script don't carry over to the function's scope. 
    ' Either that, or I'm an idiot... 
    Set objRootDSE                            = GetObject("LDAP://rootDSE") 
    strDomain                                = "LDAP://" & objRootDSE.Get("defaultNamingContext") 
    Const ADS_SCOPE_SUBTREE                    = 2 
    Set objConnection                        = CreateObject("ADODB.Connection") 
    Set objCommand                            = CreateObject("ADODB.Command") 
    objConnection.Provider                    = "ADsDSOObject" 
    objConnection.Open "Active Directory Provider" 
    Set objCommand.ActiveConnection            = objConnection 
    objCommand.Properties("Page Size")        = 1000 
    objCommand.Properties("Searchscope")    = ADS_SCOPE_SUBTREE 
     
    objCommand.CommandText = _ 
        "SELECT Name, distinguishedName FROM '"& strDomain &"' WHERE objectCategory='group' " & _ 
            "AND Name='" & strGroup & "'" 
 
    Set objRecordSet = objCommand.Execute 
 
    if objRecordSet.RecordCount > 0 then 
        objRecordSet.MoveFirst 
        Do Until objRecordSet.EOF 
            'Wscript.Echo objRecordSet.Fields("distinguishedName").Value 
            fnGetGroupDN = objRecordSet.Fields("distinguishedName").Value 
            objRecordSet.MoveNext 
        Loop 
    else 
        'WScript.Echo "No time for love, Dr. Jones!" 
        WScript.Echo strGroup & " not found." 
        WScript.Echo "If the group name has a space, please enclose the name in double-quotes." 
        WScript.Quit 
    end if 
End Function
