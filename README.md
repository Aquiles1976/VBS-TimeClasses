# VBS-TimeClasses
Classes for Time Functions with Visual Basic Scripting

Dealing with time often require to be done precisely.
The goal of this code is, to remove the complexity inherente to the native functions of VBS and allow to easyly access to detailed time information.

Classes:
--------------------------------------
1. TimeInstant
2. TimePeriod

TimeInstant Class:
--------------------------------------
Properties:
    Name     Availability   Type    Description
    -------  ------------   ------- ---------------------------------------------------------------------------
  - Updated  Read           Boolean Indicates that instant info is valid.
  - Year     Read/Write     String  Contains the year in #### format. Valid values are between 1000 and 9999.
  - Month    Read/Write     String  Contains the month number in ## format. Valid values are between 01 and 12.
  - Day      Read/Write     String  Contains the day number in ## format. Valid values are between 01 and 31.
  - Hour     Read/Write     String  Contains the hour number in ## format. Valid values are between 00 and 23.
  - Minute   Read/Write     String  Contains the minute number in ## format. Valid values are between 00 and 59.
  - Second   Read/Write     String  Contains the second number in ## format. Valid values are between 00 and 59.

Methods:

TimePeriod Class:
-------------------------------------
- Properties:
- Methods:
