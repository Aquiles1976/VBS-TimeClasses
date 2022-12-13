Option Explicit

Class TimeInterval

    Private strTIInitialYear
    Private strTIInitialMonth
    Private strTIInitialDay
    Private strTIInitialHour
    Private strTIInitialMinute
    Private strTIInitialSecond

    Private strTIFinalYear
    Private strTIFinalMonth
    Private strTIFinalDay
    Private strTIFinalHour
    Private strTIFinalMinute
    Private strTIFinalSecond

    Private Sub Class_Initialize ()
        SetInitialNow
    End Sub

    Public Function GetFixedDigits (intDigits)
        Dim objRegExp
        Set objRegExp = New RegExp
            objRegExp.Pattern = "^[0-9]$" 
        If  objRegExp.Test(intDigits) Then 
            GetFixedDigits = "0" & CStr (intDigits) 
        Else 
            GetFixedDigits = CStr (intDigits) 
        End If
    End Function

    Public Sub SetInitialNow
        strTIInitialYear   = Year ( Now )
        strTIInitialMonth  = GetFixedDigits ( Month (  Now ) ) 
        strTIInitialDay    = GetFixedDigits ( Day (    Now ) ) 
        strTIInitialHour   = GetFixedDigits ( Hour (   Now ) ) 
        strTIInitialMinute = GetFixedDigits ( Minute ( Now ) ) 
        strTIInitialSecond = GetFixedDigits ( Second ( Now ) ) 
        SetFinalNow
    End Sub

    Public Sub SetFinalNow
        strTIFinalYear   = Year ( Now )
        strTIFinalMonth  = GetFixedDigits ( Month (  Now ) ) 
        strTIFinalDay    = GetFixedDigits ( Day (    Now ) ) 
        strTIFinalHour   = GetFixedDigits ( Hour (   Now ) ) 
        strTIFinalMinute = GetFixedDigits ( Minute ( Now ) ) 
        strTIFinalSecond = GetFixedDigits ( Second ( Now ) ) 
    End Sub

    Public Function GetFormatedTime (intTimeDifference)
        Const SecondsPerMinute = 60
        Const SecondsPerHour   = 3600  '    60*60 = 3600
        Const SecondsPerDay    = 86400 ' 24*60*60 = 86400

        Dim intDays
            intDays = Int ( intTimeDifference / SecondsPerDay ) 
        intTimeDifference = intTimeDifference - ( SecondsPerDay * intDays )
        
        Dim intHours
            intHours = Int ( intTimeDifference / SecondsPerHour ) 
        intTimeDifference = intTimeDifference - ( SecondsPerHour * intHours )
        
        Dim intMinutes
            intMinutes = Int ( intTimeDifference / SecondsPerMinute ) 

        Dim intSeconds
            intSeconds = intTimeDifference - ( SecondsPerMinute * intMinutes )
                
        If intDays > 0 Then
            GetFormatedTime = intDays & "d " &_
                              GetFixedDigits ( intHours ) & ":" &_
                              GetFixedDigits (intMinutes) & ":" &_
                              GetFixedDigits (intSeconds) 
        Else
            GetFormatedTime = GetFixedDigits ( intHours ) & ":" &_
                              GetFixedDigits (intMinutes) & ":" &_
                              GetFixedDigits (intSeconds) 
        End If
    End Function

    Public Function GetInitial
        If SameDay Then
            GetInitial = strTIInitialHour   & ":" &_
                         strTIInitialMinute & ":" &_
                         strTIInitialSecond 
        Else
            GetInitial = strTIInitialYear   & "-" &_
                         strTIInitialMonth  & "-" &_
                         strTIInitialDay    & " " &_
                         strTIInitialHour   & ":" &_
                         strTIInitialMinute & ":" &_
                         strTIInitialSecond
        End If
    End Function
    
    Public Function GetFinal
        If SameDay Then
            GetFinal = strTIFinalHour   & ":" &_
                       strTIFinalMinute & ":" &_
                       strTIFinalSecond
        Else
            GetFinal = strTIFinalYear   & "-" &_
                       strTIFinalMonth  & "-" &_
                       strTIFinalDay    & " " &_
                       strTIFinalHour   & ":" &_
                       strTIFinalMinute & ":" &_
                       strTIFinalSecond
        End If
    End Function

    Public Function SameDay
        SameDay = (strTIInitialYear  = strTIFinalYear)  AND _
                  (strTIInitialMonth = strTIFinalMonth) AND _
                  (strTIInitialDay   = strTIFinalDay) 
    End Function

    Public Function GetDuration
        Dim FirstDate, LastDate, TimeIntervalInSeconds
        FirstDate = FormatDateTime (GetInitial)
        LastDate  = FormatDateTime (GetFinal)
        TimeIntervalInSeconds = DateDiff ("s",FirstDate,LastDate)
        GetDuration = GetFormatedTime (TimeIntervalInSeconds)
    End Function

End Class ' TimeInterval
    
Dim RunningPeriod
Set RunningPeriod = New TimeInterval

With RunningPeriod
    WScript.Echo "Default Initial value:" & vbTab & .GetInitial
    WScript.Echo "Default Final value:" & vbTab & .GetFinal
    WScript.Sleep 2000
    .SetInitialNow
    WScript.Sleep 20000
    .SetFinalNow
    If .SameDay Then WScript.Echo "Same Day!"
    WScript.Echo "Duration: " & .GetDuration & " from " & .GetInitial & " to " & .GetFinal
End With