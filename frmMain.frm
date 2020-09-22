VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Date Class Example"
   ClientHeight    =   2685
   ClientLeft      =   1575
   ClientTop       =   1935
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   3615
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'create an instance of the date class
    Dim MyDate As New eDate

    Me.FontName = "Courier New"
    Me.FontSize = 10
    
    'set the value of the date class to today
    'this can be two different ways
    '1)
    MyDate.Value = Now
        'you can set value = to now and the class
        'will automatically extract the date from
        'the now expression
    '2)
    MyDate.Month = Format(Now, "MM")
    MyDate.Day = Format(Now, "DD")
    MyDate.Year = Format(Now, "YYYY")
        'you can set the value of each property
        'of the date individually
        
        'NOTE: the date class automatically initalizes
        '      the value to today's date
    
    Me.AddToText "Value = " & MyDate.Value
    Me.AddToText "------------------------"
    Me.AddToText "Month = " & MyDate.Month
    Me.AddToText "Day   = " & MyDate.Day
    Me.AddToText "Year  = " & MyDate.Year
    Me.AddToText ""
    
    'now you can do math on the date
    'you can easy move in single increments of dates
    Me.AddToText "Next and Previous functions:"
    Me.AddToText "Next Day:   " & MyDate.NextDay
    Me.AddToText "Next Week:  " & MyDate.NextWeek
    Me.AddToText "Next Month: " & MyDate.NextMonth
    Me.AddToText "Next Year:  " & MyDate.NextYear
    Me.AddToText "Previous Day:   " & MyDate.PreviousDay
    Me.AddToText "Previous Week:  " & MyDate.PreviousWeek
    Me.AddToText "Previous Month: " & MyDate.PreviousMonth
    Me.AddToText "Previous Year:  " & MyDate.PreviousYear
    Me.AddToText ""
    
    'NOTE: when you use the next or previous functions,
    '      it changes the value stored in the date class
    
    'you can also do addition and subtraction of larger increments
    Me.AddToText "Addition and Subtraction:"
    MyDate.Add 3, Days
    Me.AddToText "Value + 3 days:   " & MyDate.Value
    MyDate.Subtract 3, Days
    Me.AddToText "Value - 3 days:   " & MyDate.Value
    MyDate.Add 4, Weeks
    Me.AddToText "Value + 4 weeks:  " & MyDate.Value
    MyDate.Subtract 4, Weeks
    Me.AddToText "Value - 4 weeks:  " & MyDate.Value
    MyDate.Add 5, Months
    Me.AddToText "Value + 5 months: " & MyDate.Value
    MyDate.Subtract 5, Months
    Me.AddToText "Value - 5 months: " & MyDate.Value
    MyDate.Add 6, Years
    Me.AddToText "Value + 6 years:  " & MyDate.Value
    MyDate.Subtract 6, Years
    Me.AddToText "Value - 6 years:  " & MyDate.Value
    Me.AddToText ""
    
    'you can also use the get functions to get values from a date
    Me.AddToText "Get functions:"
    Me.AddToText "Get Month: " & MyDate.GetMonth
    Me.AddToText "Get Day: " & MyDate.GetDay
    Me.AddToText "Get Year: " & MyDate.GetYear
    
    'NOTE: you can also specify a different date value to extract the
    '      month, day, and year from
    
    Me.AddToText "Get Month 8/14/2002: (numeric) " & MyDate.GetMonth(Numeric, "8/14/2002")
    Me.AddToText "Get Month 8/14/2002: (short date) " & MyDate.GetMonth(ShortDate, "8/14/2002")
    Me.AddToText "Get Month 8/14/2002: (long date) " & MyDate.GetMonth(LongDate, "8/14/2002")
    Me.AddToText "Get Day 8/14/2002: (numeric) " & MyDate.GetDay(Numeric, "8/14/2002")
    Me.AddToText "Get Day 8/14/2002: (short date) " & MyDate.GetDay(ShortDate, "8/14/2002")
    Me.AddToText "Get Day 8/14/2002: (long date) " & MyDate.GetDay(LongDate, "8/14/2002")
    Me.AddToText "Get Year 8/14/2002: (numeric) " & MyDate.GetYear(Numeric, "8/14/2002")
    Me.AddToText "Get Year 8/14/2002: (short date) " & MyDate.GetYear(ShortDate, "8/14/2002")
    Me.AddToText "Get Year 8/14/2002: (long date) " & MyDate.GetYear(LongDate, "8/14/2002")
    
    'clean up
    Set MyDate = Nothing
End Sub

Private Sub Form_Resize()
    'resize the textbox to fit on the form
    Me.txtInfo.Width = Me.Width - Me.txtInfo.Left - 100
    Me.txtInfo.Height = Me.Height - 400 - Me.txtInfo.Top
End Sub

Public Sub AddToText(TextToAdd As String)
    'this method just adds a strin to the textbox
    Me.txtInfo.Text = Me.txtInfo.Text & TextToAdd & vbCrLf
End Sub
