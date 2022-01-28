# VBS-TimeClasses
Classes for Time Functions with Visual Basic Scripting

Operations with time are required to be done precisely. This code has removed the complexity inherent to the native Visual Basic functions, allowing access to detailed time information in an easier and smarter way.

Classes:
--------------------------------------
- TimeInstant
- TimePeriod

TimeInstant Class:
--------------------------------------

- Properties:

Name   | Availability | Type  | Description | Valid values
---    | ---          | ---   | ---         | ---
Updated|Read          |Boolean|Indicates that time data is valid.| True or False.
Year   |Read/Write    |String |Contains the year in #### format.| Valid values are between 1000 and 9999.
Month  |Read/Write    |String |Contains the month number in ## format.| Valid values are between 01 and 12.
Day    |Read/Write    |String |Contains the day number in ## format.| Valid values are between 01 and 31.
Hour   |Read/Write    |String |Contains the hour number in ## format.| Valid values are between 00 and 23.
Minute |Read/Write    |String |Contains the minute number in ## format.| Valid values are between 00 and 59.
Second |Read/Write    |String |Contains the second number in ## format.| Valid values are between 00 and 59.

- Methods:

None.

TimePeriod Class:
-------------------------------------
- Properties:

Name   | Availability | Type  | Description 
---    | ---          | ---   | ---         
FirstMoment|Read|TimeInstant|Indicates that time data for the first moment in the period.
LastMoment |Read|TimeInstant|Indicates that time data for the last moment in the period.

- Methods:

Name   | Availability | Type  | Description 
---    | ---          | ---   | ---         
SameDay|Read|Boolean|Returns true only if first and last moments are in the same day of the year, and both have been updated.
GetFirstMoment|Read|String|Returns the time data for the first moment in the period.
GetLastMoment |Read|String|Returns the time data for the last moment in the period.
GetShortedFirstMoment|Read|String|Returns the time data for the first moment in the period, if both moments are in the same day returns only the time part.
GetShortedLastMoment |Read|String|Returns the time data for the last moment in the period, if both moments are in the same day returns only the time part.
GetDuration|Read|String|Returns the time difference between the first and the last moment.
GetFormatedTime|Read|String|Returns the time in ##:##:## format. The input has to be in seconds.
GetFixedDigits|Read|String|Returns a two char string representing a numeric value always in ## format.
SetStartNow|Write|None|Gets the actual local time and writes it to the first moment object.
SetEndNow|Write|None|Gets the actual local time and writes it to the last moment object.


