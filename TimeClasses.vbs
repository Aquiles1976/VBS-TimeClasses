Option Explicit

Class TimeInstant

    Private strTimeInstantYear      ' from 1000 to 9999
    Private strTimeInstantMonth     ' from 01 to 12
    Private strTimeInstantDay       ' from 01 to 31
    Private strTimeInstantHour      ' from 00 to 23
    Private strTimeInstantMinute    ' from 00 to 59
    Private strTimeInstantSecond    ' from 00 to 59
    Private blnUpdated              ' True when object updated, false by default

    Private Sub Class_Initialize()
        'This event is called when an instance of the class is instantiated
        'Initialize properties here and perform other start-up tasks
        strTimeInstantYear   = "1000"   ' from 1000 to 9999
        strTimeInstantMonth  = "01"     ' from 01 to 12
        strTimeInstantDay    = "01"     ' from 01 to 31
        strTimeInstantHour   = "00"     ' from 00 to 23
        strTimeInstantMinute = "00"     ' from 00 to 59
        strTimeInstantSecond = "00"     ' from 00 to 59
        blnUpdated = False
    End Sub

    Private Sub Class_Terminate()
        'This event is called when a class instance is destroyed
        'either explicitly (Set objClassInstance = Nothing) or
        'implicitly (it goes out of scope)
    End Sub

    Public Property Get Updated
        Updated = blnUpdated
    End Property

    Public Property Let Updated (inputUpdated)
        If TypeName(inputUpdated) = "Boolean" Then
            blnUpdated = inputUpdated
        End If
    End Property
    '*************************************************************************** YEAR
    Public Property Get Year 
        Year = strTimeInstantYear
    End Property

    Public Property Let Year(inputYear)
        If IsObject(inputYear) Or IsNull(inputYear) Or IsEmpty(inputYear) then
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid input type for YEAR: " & TypeName(inputYear) & " (Must be numerical)"
            Exit property
        End If
        Dim objRegExp
        Set objRegExp = New RegExp
            objRegExp.Pattern = "^[1-9]\d{3}$" ' The year in #### format, from 1000 to 9999
        If objRegExp.Test(inputYear) Then 
            strTimeInstantYear = CStr(inputYear) 
        Else
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid value for YEAR, must be between 1000 and 9999."
        End If
        Set objRegExp = Nothing
    End Property

    '*************************************************************************** MONTH
    Public Property Get Month
        Month = strTimeInstantMonth
    End Property

    Public Property Let Month(inputMonth)
        If IsObject(inputMonth) Or IsNull(inputMonth) Or IsEmpty(inputMonth) Then 
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid input type for MONTH: " & TypeName(inputMonth) & " (Must be numerical)"
            Exit property
        End If
        Dim objRegExp
        Set objRegExp = New RegExp
            objRegExp.Pattern = "^(0[1-9])$|^(1[0-2])$" ' The month in ## format (01-12)
        If objRegExp.Test(inputMonth) Then 
            strTimeInstantMonth = CStr(inputMonth) 
        Else 
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid value for MONTH, must be in ## format, between 01 and 12, but you entered: " & inputMonth
        End If
        Set objRegExp = Nothing
    End Property

    '*************************************************************************** DAY
    Public Property Get Day
        Day = strTimeInstantDay
    End Property

    Public Property Let Day(inputDay)
        If IsObject(inputDay) Or IsNull(inputDay) Or IsEmpty(inputDay) Then 
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid input type for DAY: " & TypeName(inputDay) & " (Must be numerical)"
            Exit property
        End If
        Dim objRegExp
        Set objRegExp = New RegExp
            objRegExp.Pattern = "^(0[1-9])$|^([1-2][0-9])$|^(3[0-1])$" ' The day in ## format (01-31)
        If objRegExp.Test(inputDay) Then 
            strTimeInstantDay = CStr(inputDay) 
        Else 
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid value for DAY, must be in ## format, between 01 and 31, but you entered: " & inputDay
        End If
        Set objRegExp = Nothing
    End Property

    '*************************************************************************** HOUR
    Public Property Get Hour
        Hour = strTimeInstantHour
    End Property

    Public Property Let Hour(inputHour)
        If IsObject(inputHour) Or IsNull(inputHour) Or IsEmpty(inputHour) Then 
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid input type for HOUR: " & TypeName(inputHour) & " (Must be numerical)"
            Exit property
        End If
        Dim objRegExp
        Set objRegExp = New RegExp
            objRegExp.Pattern = "^([0-1][0-9])$|^(2[0-3])$" ' The hour in ## format (00-23)
        If objRegExp.Test(inputHour) Then 
            strTimeInstantHour = CStr(inputHour) 
        Else 
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid value for HOUR, must be in ## format, between 00 and 23, but you entered: " & inputHour
        End If
        Set objRegExp = Nothing
    End Property

    '*************************************************************************** MINUTE
    Public Property Get Minute
        Minute = strTimeInstantMinute
    End Property

    Public Property Let Minute(inputMinute)
        If IsObject(inputMinute) Or IsNull(inputMinute) Or IsEmpty(inputMinute) Then 
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid input type for MINUTE: " & TypeName(inputMinute) & " (Must be numerical)"
            Exit property
        End If
        Dim objRegExp
        Set objRegExp = New RegExp
            objRegExp.Pattern = "^[0-5][0-9]$" ' The minute in ## format, from 00 to 59
        If objRegExp.Test(inputMinute) Then 
            strTimeInstantMinute = CStr(inputMinute) 
        Else 
            objRegExp.Pattern = "^[0-9]$" ' The minute in # format
            If objRegExp.Test(inputMinute) Then 
                strTimeInstantMinute = "0" & CStr(inputMinute) 
            Else
                Err.Raise vbObjectError + 1000, "TimeInstant Class", _
                "Invalid value for MINUTE, must be in ## format, between 00 and 59, but you entered: " & inputMinute
            End If
        End If
        Set objRegExp = Nothing
    End Property

    '*************************************************************************** SECOND
    Public Property Get Second
        Second = strTimeInstantSecond
    End Property

    Public Property Let Second(inputSecond)
        If IsObject(inputSecond) Or IsNull(inputSecond) Or IsEmpty(inputSecond) Then 
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid input type for SECOND: " & TypeName(inputSecond) & " (Must be numerical)"
            Exit property
        End If
        Dim objRegExp
        Set objRegExp = New RegExp
            objRegExp.Pattern = "^[0-5][0-9]$" ' The second in ## format
        If objRegExp.Test(inputSecond) Then 
            strTimeInstantSecond = CStr(inputSecond) 
        Else 
            Err.Raise vbObjectError + 1000, "TimeInstant Class", _
            "Invalid value for SECOND, must be in ## format, between 00 and 59, but you entered: " & inputSecond
        End If
        Set objRegExp = Nothing
    End Property    

End Class ' TimeInstant












Class TimePeriod

    Private objStartInstant
    Private objEndInstant

    Private Sub Class_Initialize()
        'This event is called when an instance of the class is instantiated
        'Initialize properties here and perform other start-up tasks
        Set objStartInstant = New TimeInstant
        Set objEndInstant   = New TimeInstant    
    End Sub

    Private Sub Class_Terminate()
        'This event is called when a class instance is destroyed
        'either explicitly (Set objClassInstance = Nothing) or
        'implicitly (it goes out of scope)
        Set objStartInstant = Nothing
        Set objEndInstant   = Nothing
    End Sub

    Public Property Get FirstMoment
        Set FirstMoment = objStartInstant
    End Property

    Public Property Get LastMoment
        Set LastMoment = objEndInstant
    End Property
    
    Public Sub SetStartNow
        objStartInstant.Year    = Year(Now)
        objStartInstant.Month   = GetFixedDigits(Month(Now))
        objStartInstant.Day     = GetFixedDigits(Day(Now))
        objStartInstant.Hour    = GetFixedDigits(Hour(Now))
        objStartInstant.Minute  = GetFixedDigits(Minute(Now))
        objStartInstant.Second  = GetFixedDigits(Second(Now))
        objStartInstant.Updated = True
    End Sub

    Public Sub SetEndNow
        objEndInstant.Year    = Year(Now)
        objEndInstant.Month   = GetFixedDigits(Month(Now))
        objEndInstant.Day     = GetFixedDigits(Day(Now))
        objEndInstant.Hour    = GetFixedDigits(Hour(Now))
        objEndInstant.Minute  = GetFixedDigits(Minute(Now))
        objEndInstant.Second  = GetFixedDigits(Second(Now))
        objEndInstant.Updated = True
    End Sub

    Public Function GetFirstMoment
        If objStartInstant.Updated Then
            GetFirstMoment = objStartInstant.Year   & "-" &_
                             objStartInstant.Month  & "-" &_
                             objStartInstant.Day    & " " &_
                             objStartInstant.Hour   & ":" &_
                             objStartInstant.Minute & ":" &_
                             objStartInstant.Second
        Else
            GetFirstMoment = "Undefined"
        End If
    End Function

    Public Function GetShortedFirstMoment
        If objStartInstant.Updated Then
            If SameDay Then
                GetShortedFirstMoment = objStartInstant.Hour   & ":" &_
                                        objStartInstant.Minute & ":" &_
                                        objStartInstant.Second
            Else
                GetShortedFirstMoment = objStartInstant.Year   & "-" &_
                                        objStartInstant.Month  & "-" &_
                                        objStartInstant.Day    & " " &_
                                        objStartInstant.Hour   & ":" &_
                                        objStartInstant.Minute & ":" &_
                                        objStartInstant.Second
            End If
        Else
            GetShortedFirstMoment = "Undefined"
        End If
    End Function
    

    Public Function GetLastMoment
        If objEndInstant.Updated Then
            GetLastMoment = objEndInstant.Year    & "-" &_
                             objEndInstant.Month  & "-" &_
                             objEndInstant.Day    & " " &_
                             objEndInstant.Hour   & ":" &_
                             objEndInstant.Minute & ":" &_
                             objEndInstant.Second
        Else
            GetLastMoment = "Undefined"
        End If
    End Function

    Public Function GetShortedLastMoment
        If objEndInstant.Updated Then
            If SameDay Then
                GetShortedLastMoment =  objEndInstant.Hour   & ":" &_
                                        objEndInstant.Minute & ":" &_
                                        objEndInstant.Second
            Else
                GetShortedLastMoment =  objEndInstant.Year   & "-" &_
                                        objEndInstant.Month  & "-" &_
                                        objEndInstant.Day    & " " &_
                                        objEndInstant.Hour   & ":" &_
                                        objEndInstant.Minute & ":" &_
                                        objEndInstant.Second
            End If
        Else
            GetShortedLastMoment = "Undefined"
        End If
    End Function

    Public Function SameDay
        SameDay = (objStartInstant.Year = objEndInstant.Year)   AND _
                  (objStartInstant.Month = objEndInstant.Month) AND _
                  (objStartInstant.Day = objEndInstant.Day)
    End Function

    Public Function GetDuration
        Dim FirstDate, LastDate, TimeIntervalInSeconds
        FirstDate = FormatDateTime(GetFirstMoment)
        LastDate  = FormatDateTime(GetLastMoment)
        TimeIntervalInSeconds = DateDiff("s",FirstDate,LastDate)
        GetDuration = GetFormatedTime(TimeIntervalInSeconds)
    End Function

    Public Function GetFormatedTime(intElapsedSeconds)
        Const SecondsPerMinute = 60
        Const SecondsPerHour   = 3600  ' 60*60 = 3600
        Const SecondsPerDay    = 86400 ' 60*60*24 = 86400

        Dim intTimeDifference
            intTimeDifference = intElapsedSeconds

        Dim intDays
            intDays = Int( intTimeDifference / SecondsPerDay ) 
        
        intTimeDifference = intTimeDifference - ( SecondsPerDay * intDays )
        
        Dim intHours
            intHours = Int( intTimeDifference / SecondsPerHour ) 

        intTimeDifference = intTimeDifference - ( SecondsPerHour * intHours )
        
        Dim intMinutes
            intMinutes = Int( intTimeDifference / SecondsPerMinute ) 

        Dim intSeconds
            intSeconds = intTimeDifference - ( SecondsPerMinute * intMinutes )
                
        If intDays > 0 Then
            GetFormatedTime = intDays & ":" & intHours & ":" & intMinutes & ":" & intSeconds
        Else
            GetFormatedTime = intHours & ":" & intMinutes & ":" & intSeconds
        End If

    End Function

    Public Function GetFixedDigits(intDigits)
        Dim objRegExp
        Set objRegExp = New RegExp
            objRegExp.Pattern = "^[0-9]$" 
        If  objRegExp.Test(intDigits) Then 
            GetFixedDigits = "0" & CStr(intDigits) 
        Else 
            objRegExp.Pattern = "^[0-9][0-9]$" 
            If  objRegExp.Test(intDigits) Then 
                GetFixedDigits = CStr(intDigits) 
            Else
                GetFixedDigits = "00"
            End If
        End If
    End Function

End Class ' TimePeriod

Dim RunningPeriod
Set RunningPeriod = New TimePeriod

With RunningPeriod
    WScript.Echo "Default YEAR value:" & .FirstMoment.Year
    .SetStartNow
    WScript.Echo "First moment: " & .GetFirstMoment
    WScript.Sleep 70000
    .SetEndNow
    WScript.Echo "Last moment: " & .GetLastMoment
    If .SameDay Then WScript.Echo "Same Day!"
    WScript.Echo "Duration: " & .GetDuration & " from " & .GetShortedFirstMoment & " to " & .GetShortedLastMoment
End With

'WScript.Echo "Default YEAR value:" & RunningPeriod.FirstMoment.Year
'RunningPeriod.SetStartNow
'WScript.Echo "First moment: " & RunningPeriod.GetFirstMoment
'WScript.Sleep 70000
'RunningPeriod.SetEndNow
'WScript.Echo "Last moment: " RunningPeriod.GetLastMoment
'If RunningPeriod.SameDay Then WScript.Echo "Same Day!"
'WScript.Echo "Duration: " & RunningPeriod.GetDuration & " from " & RunningPeriod.GetShortedFirstMoment
