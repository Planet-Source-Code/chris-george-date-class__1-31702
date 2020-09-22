VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'eDate Class
'By: Chris George
'February 6,2002
'=============================================================
'This is a date class which simplifies mathamatical operations
'using dates.  Use this class instead of adding heavyweight
'calander controls to do your math on dates.

'PLEASE include my name along with this code if you use it. :)
'=============================================================

Private pValue As Date          'Current Value

'This enum is used to enumerate the different
'units of dates that can be added or subtracted
Public Enum dEnum
    Days = 0
    Weeks = 1
    Months = 2
    Years = 3
End Enum

'This enumeration is used to help the user
'specify how they want the day of the week
'returned when using the Get Day property
Public Enum DateTypeEnum
    Numeric = 0
    ShortDate = 1
    LongDate = 2
End Enum

Private Sub Class_Initialize()
    'initialize the value of pvalue to now
    pValue = Format(Now, "MM/DD/YYYY")
End Sub

Public Sub Add(HowMany As Integer, DateUnits As dEnum)
    'This method is used to add a unit of dates to the
    'current value
    Dim TempDate As Integer
    Dim TempYear As Integer
    
    'error handle
    On Error GoTo Err_Handle
    
    Select Case DateUnits
        Case Days
            'add days to the current value. To increase
            'the day value of any date, just add an integer to it
            pValue = pValue + HowMany
        Case Weeks
            'add 7 days to the current value. To increase
            'the day value of any date, just add an integer to it
            'then multiply it times 7
            pValue = pValue + (7 * HowMany)
        Case Months
            'add months to the current value. To increase months
            'you must store the month value into an integer value
            'do your math on the integer value then concatinate it
            'back into the date
            
            'make sure the user isn't adding more than 12 months
            If HowMany > 12 Then
                'if they are then divide it by twelve to get the number of years
                'and get the remander to be the number of months
                TempYear = Int(HowMany / 12)
                TempDate = HowMany Mod 12
                'now add to the current month and year
                TempDate = Val(Format(pValue, "MM")) + TempDate
                'make sure the month value isn't greater than 12
                If TempDate > 12 Then TempDate = TempDate - 12: TempYear = TempYear + 1
                'make sure the month value isn't less than 1 (incase the user passed a negative value)
                If TempDate < 1 Then TempDate = TempDate + 12: TempYear = TempYear - 1
                TempYear = Val(Format(pValue, "YYYY")) + TempYear
            Else
                'add to the current month and store the year
                TempDate = Val(Format(pValue, "MM")) + HowMany
                TempYear = Val(Format(pValue, "YYYY"))
                'make sure the month value isn't greater than 12
                If TempDate > 12 Then TempDate = TempDate - 12: TempYear = TempYear + 1
                'make sure the month value isn't less than 1 (incase the user passed a negative value)
                If TempDate < 1 Then TempDate = TempDate + 12: TempYear = TempYear - 1
            End If
            
            'concatinate the values with the days
            pValue = TempDate & Format(pValue, "/dd/") & TempYear
        Case Years
            'add years to the current value. To increase years
            'you must store the year value into an integer value
            'do your math on the integer value then concatinate it
            'back into the date
            TempDate = Val(Format(pValue, "YYYY")) + HowMany
            pValue = Format(pValue, "MM/DD/") & TempDate
        Case Else
            'if the user passed something unknown then send an error message
            MsgBox "Invalid DateUnits Format: " & DateUnits, vbExclamation + vbOKOnly
    End Select
    
    Exit Sub
    
Err_Handle:
    MsgBox "Error adding value to date: " & Error, vbExclamation + vbOKOnly
End Sub

Public Function GetDay(Optional DateType As DateTypeEnum, Optional OfDate As String) As String
    'This function is used to get a day value of a date string
    'if the user passed a date then get the day value of
    'the date passed, otherwise pass the day value of the current value
    'DayOfWeek specifies the format to return the value of the day
    Select Case DateType
        Case Numereic '(1-31)
            If OfDate = "" Then
                GetDay = Format(pValue, "DD")
            Else
                GetDay = Format(OfDate, "DD")
            End If
        Case ShortDate '(Sun-Sat)
            If OfDate = "" Then
                GetDay = Format(pValue, "DDD")
            Else
                GetDay = Format(OfDate, "DDD")
            End If
        Case LongDate '(Sunday-Saturday)
            If OfDate = "" Then
                GetDay = Format(pValue, "DDDD")
            Else
                GetDay = Format(OfDate, "DDDD")
            End If
    End Select
End Function

Public Function GetMonth(Optional DateType As DateTypeEnum, Optional OfDate As String) As String
    'This function is used to get a day value of a date string
    'if the user passed a date then get the day value of
    'the date passed, otherwise pass the day value of the current value
    'DayOfWeek specifies the format to return the value of the day
    Select Case DateType
        Case Numereic '(1-31)
            If OfDate = "" Then
                GetMonth = Format(pValue, "MM")
            Else
                GetMonth = Format(OfDate, "MM")
            End If
        Case ShortDate '(Jan-Dec)
            If OfDate = "" Then
                GetMonth = Format(pValue, "MMM")
            Else
                GetMonth = Format(OfDate, "MMM")
            End If
        Case LongDate '(Januraru-December)
            If OfDate = "" Then
                GetMonth = Format(pValue, "MMMM")
            Else
                GetMonth = Format(OfDate, "MMMM")
            End If
    End Select
End Function

Public Function GetYear(Optional DateType As DateTypeEnum, Optional OfDate As String) As String
    'This function is used to get a year value of a date string
    'if the user passed a date then get the year value of
    'the date passed, otherwise pass the year value of the current value
    'DayOfWeek specifies the format to return the value of the year
    
    Select Case DateType
        Case Numereic '(0-99)
            If OfDate = "" Then
                GetYear = Format(pValue, "YY")
            Else
                GetYear = Format(OfDate, "YY")
            End If
        Case ShortDate '(00-99)
            If OfDate = "" Then
                GetYear = Format(pValue, "YY")
            Else
                GetYear = Format(OfDate, "YY")
            End If
        Case LongDate '(0000-9999)
            If OfDate = "" Then
                GetYear = Format(pValue, "YYYY")
            Else
                GetYear = Format(OfDate, "YYYY")
            End If
    End Select
End Function

Public Sub Subtract(HowMany As Integer, DateUnits As dEnum)
    'This method is used to subtract a unit of dates from the
    'current value
    Dim TempDate As Integer
    Dim TempYear As Integer
    
    'error handle
    On Error GoTo Err_Handle
    
    Select Case DateUnits
        Case Days
            'add days to the current value. To decrease
            'the day value of any date, just subtract an integer to it
            pValue = pValue - HowMany
        Case Weeks
            'add 7 days to the current value. To decrease
            'the day value of any date, just subtract an integer to it
            'then multiply it times 7
            pValue = pValue - (7 * HowMany)
        Case Months
            'subtract months from the current value. To decrease months
            'you must store the month value into an integer value
            'do your math on the integer value then concatinate it
            'back into the date
            'make sure the user isn't subtacting more than 12 months
            If HowMany > 12 Then
                'if they are then divide it by twelve to get the number of years
                'and get the remander to be the number of months
                TempYear = Int(HowMany / 12)
                TempDate = HowMany Mod 12
                'now subract from the current month and year
                TempDate = Val(Format(pValue, "MM")) - TempDate
                'make sure the month value isn't less than 1
                If TempDate < 1 Then TempDate = TempDate + 12: TempYear = TempYear - 1
                'make sure the month value isn't greater than 12 (incase the user passed a negative value)
                If TempDate > 12 Then TempDate = TempDate - 12: TempYear = TempYear + 1
                TempYear = Val(Format(pValue, "YYYY")) - TempYear
            Else
                'subtract from the current month and store the year
                TempDate = Val(Format(pValue, "MM")) - HowMany
                TempYear = Val(Format(pValue, "YYYY"))
                'make sure the month value isn't less than 1
                If TempDate < 1 Then TempDate = TempDate + 12: TempYear = TempYear - 1
                'make sure the month value isn't greater than 12 (incase the user passed a negative value)
                If TempDate > 12 Then TempDate = TempDate - 12: TempYear = TempYear + 1
            End If
            'concatinate the values with the days
            pValue = TempDate & Format(pValue, "/dd/") & TempYear
        Case Years
            'add years to the current value. To decrease years
            'you must store the year value into an integer value
            'do your math on the integer value then concatinate it
            'back into the date
            TempDate = Val(Format(pValue, "YYYY")) - HowMany
            pValue = Format(pValue, "MM/DD/") & TempDate
        Case Else
            'if the user passed something unknown then send an error message
            MsgBox "Invalid DateUnits Format: " & DateUnits, vbExclamation + vbOKOnly
    End Select
    
    Exit Sub
    
Err_Handle:
    MsgBox "Error subtracting value from date: " & Error, vbExclamation + vbOKOnly
End Sub

'The following methods make use of the add and subtract methods above.
'I added these just to increase functionality of the class.  All of these
'functions could be performed using the add and subtract methods.
'=====================================================================
Public Function NextDay() As Date
    'This method just simply adds a day to the current value
    Me.Add 1, Days
    NextDay = pValue
End Function

Public Function NextWeek() As Date
    'This method adds a week to the current value
    Me.Add 1, Weeks
    NextWeek = pValue
End Function

Public Function NextMonth() As Date
    'This method adds a month to the current value
    Me.Add 1, Months
    NextMonth = pValue
End Function

Public Function NextYear() As Date
    'This method adds a year to the current value
    Me.Add 1, Years
    NextYear = pValue
End Function

Public Function PreviousDay() As Date
    'This method subtracts a day from the current value
    Me.Subtract 1, Days
    PreviousDay = pValue
End Function

Public Function PreviousWeek() As Date
    'This method subtracts a week from the current value
    Me.Subtract 1, Weeks
    PreviousWeek = pValue
End Function

Public Function PreviousMonth() As Date
    'This method subtracts a month from the current value
    Me.Subtract 1, Months
    PreviousMonth = pValue
End Function

Public Function PreviousYear() As Date
    'This method subtracts a year from the current value
    Me.Subtract 1, Years
    PreviousYear = pValue
End Function

Public Property Get Day() As Integer
    'return the current day of pvalue
    Day = Val(Format(pValue, "DD"))
End Property

Public Property Let Day(ByVal vNewDay As Integer)
    'make sure the new day is a valid day
    Dim TempDate As Date
    Dim TempStr As String
    'make sure vnewday is a value between 1 and 31
    If vNewDay < 1 Or vNewDay > 31 Then
        MsgBox "Day value cannot be greater than 31 or less than 1.", vbInformation + vbOKOnly
        Exit Property
    End If
    
    'make sure the new date value is a real date
    TempStr = Format(Format(pValue, "MM/") & vNewDay & Format(pValue, "/YYYY"), "MM/DD/YYYY")
    
    If IsDate(TempStr) = True Then
        pValue = TempStr
    Else
        'send a message to the user indicating that it is not a valid date
        MsgBox vNewDay & " is not a valid day value for the month of " & Format(pValue, "MMMM"), vbExclamation + vbOKOnly
    End If
End Property

Public Property Get Month() As Integer
    'return the current day of pvalue
    Month = Val(Format(pValue, "MM"))
End Property

Public Property Let Month(ByVal vNewMonth As Integer)
    'make sure the new month is a valid month
    Dim TempDate As Date
    Dim TempStr As String
    'make sure vnewmonth is a value between 1 and 12
    If vNewMonth < 1 Or vNewMonth > 12 Then
        MsgBox "Month value cannot be greater than 12 or less than 1.", vbInformation + vbOKOnly
        Exit Property
    End If
    
    'make sure the new date value is a real date
    TempStr = Format(vNewMonth & Format(pValue, "/DD/YYYY"), "MM/DD/YYYY")
    
    If IsDate(TempStr) = True Then
        pValue = TempStr
    Else
        'send a message to the user indicating that it is not a valid date
        MsgBox vNewMonth & " is not a valid month value.", vbExclamation + vbOKOnly
    End If
End Property


Public Property Get Year() As Integer
    'return the current year of pvalue
    Year = Val(Format(pValue, "YYYY"))
End Property

Public Property Let Year(ByVal vNewYear As Integer)
    'make sure the new year is a valid year
    Dim TempDate As Date
    Dim TempStr As String
    'make sure vnewyear is a value between 1 and 9999
    If vNewYear < 0 Or vNewYear > 9999 Then
        MsgBox "year value cannot be greater than 9999 or less than 1.", vbInformation + vbOKOnly
        Exit Property
    End If
    
    'make sure the new date value is a real date
    TempStr = Format(Format(pValue, "MM/DD/") & vNewYear, "MM/DD/YYYY")
    
    If IsDate(TempStr) = True Then
        pValue = TempStr
    Else
        'send a message to the user indicating that it is not a valid date
        MsgBox vNewYear & " is not a valid year value.", vbExclamation + vbOKOnly
    End If
End Property


Public Property Get Value() As Date
    'return the current value
    Value = pValue
End Property

Public Property Let Value(ByVal vNewDate As Date)
    'set it to the new value
    pValue = Format(vNewDate, "MM/DD/YYYY")
End Property
