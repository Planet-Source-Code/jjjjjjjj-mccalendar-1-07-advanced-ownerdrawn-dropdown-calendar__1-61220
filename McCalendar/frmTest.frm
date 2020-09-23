VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "McCalendar - 1.6  [ Test Form ]        -       'Jim Jose'"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10140
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSpecial 
      Height          =   360
      Left            =   120
      TabIndex        =   52
      Text            =   "21-1,26-1,19-2,8-3,24-3,25-3,14-4,21-4,13-8,15-8,20-8,26-8,14-9,15-9,16-9,17-9,21-9,11-10,12-10,1-11,3-11,10-12,25,12"
      Top             =   4560
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   5040
      TabIndex        =   0
      Top             =   840
      Width           =   2415
      Begin VB.ListBox lstMode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         ItemData        =   "frmTest.frx":0000
         Left            =   120
         List            =   "frmTest.frx":0013
         TabIndex        =   50
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Text            =   "195"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtHeader 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1680
         TabIndex        =   39
         Text            =   "18"
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkSkip 
         Caption         =   "Skip Enabled"
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox chkSensitive 
         Caption         =   "Sensitive"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox cmbFormat 
         Height          =   360
         ItemData        =   "frmTest.frx":0080
         Left            =   120
         List            =   "frmTest.frx":008D
         TabIndex        =   30
         Text            =   "[dd-mm-yyyy]"
         Top             =   4080
         Width           =   2175
      End
      Begin VB.CheckBox chkAnimate 
         Caption         =   "Animate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txtCurvature 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1680
         TabIndex        =   7
         Text            =   "0"
         Top             =   1560
         Width           =   495
      End
      Begin VB.CheckBox chkAppearance 
         Caption         =   "3D Appearance"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox chkBorder 
         Caption         =   "Border"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Modes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   51
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "Header Height"
         Height          =   240
         Left            =   120
         TabIndex        =   40
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date Format"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   31
         Top             =   3840
         Width           =   1290
      End
      Begin VB.Label Label4 
         Caption         =   "Calendar Height"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Curvature"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Date Picker (Up)"
      Height          =   735
      Left            =   5040
      TabIndex        =   47
      Top             =   5400
      Width           =   2415
      Begin VB.TextBox txtDatePick 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   48
         Text            =   "DatePicker"
         Top             =   240
         Width           =   1695
      End
      Begin AdvancedCalendar.McCalendar McCalendar3 
         Height          =   345
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Theme           =   4
         Animate         =   -1  'True
         CalendarHeight  =   150
         Mode            =   4
         CalendarBackCol =   13820669
         MonthGradient   =   0   'False
         MonthBackCol    =   13820669
         HeaderBackCol   =   13820669
         WeekDayCol      =   14934998
         DayCol          =   15726078
         DaySelCol       =   11651986
         WeekDaySelCol   =   15857131
         DaySunCol       =   14934998
         WeekDaySunCol   =   14004415
         YearBackCol     =   13820669
         HeaderHeight    =   23
         SpecialDays     =   "21-1,26-1,19-2,8-3,24-3,25-3,14-4,21-4,13-8,15-8,20-8,26-8,14-9,15-9,16-9,17-9,21-9,11-10,12-10,1-11,3-11,10-12,25-12"
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Date Picker (Down)"
      Height          =   735
      Left            =   5040
      TabIndex        =   44
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtDtaePick2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   45
         Text            =   "DatePicker"
         Top             =   240
         Width           =   1695
      End
      Begin AdvancedCalendar.McCalendar McCalendar4 
         Height          =   345
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Theme           =   3
         Animate         =   -1  'True
         CalendarHeight  =   150
         Mode            =   3
         CalendarBackCol =   11651986
         MonthGradient   =   0   'False
         MonthBackCol    =   11651986
         HeaderBackCol   =   11651986
         WeekDayCol      =   16443612
         DayCol          =   14805973
         DaySelCol       =   14334632
         WeekDaySelCol   =   15857131
         DaySunCol       =   11446008
         YearBackCol     =   11651986
         HeaderHeight    =   23
         SpecialDays     =   "21-1,26-1,19-2,8-3,24-3,25-3,14-4,21-4,13-8,15-8,20-8,26-8,14-9,15-9,16-9,17-9,21-9,11-10,12-10,1-11,3-11,10-12,25-12"
      End
   End
   Begin AdvancedCalendar.McCalendar McCalendar1 
      Height          =   3135
      Left            =   240
      TabIndex        =   32
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Animate         =   -1  'True
      CalendarHeight  =   191
      CalendarBackCol =   16743805
      SpecialDays     =   "21-1,26-1,19-2,8-3,24-3,25-3,14-4,21-4,13-8,15-8,20-8,26-8,14-9,15-9,16-9,17-9,21-9,11-10,12-10,1-11,3-11,10-12,25-12"
   End
   Begin VB.ListBox lstTheme 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      ItemData        =   "frmTest.frx":00BB
      Left            =   120
      List            =   "frmTest.frx":00D7
      TabIndex        =   18
      Top             =   5160
      Width           =   4815
   End
   Begin VB.Frame Frame3 
      Height          =   6015
      Left            =   7560
      TabIndex        =   12
      Top             =   120
      Width           =   2415
      Begin VB.CheckBox chkGradient 
         Caption         =   "HeaderGradient"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   4920
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkGradient 
         Caption         =   "YearGradient"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   5400
         Width           =   2175
      End
      Begin VB.CheckBox chkGradient 
         Caption         =   "MonthGradient"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   5160
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkGradient 
         Caption         =   "CalendarGradient"
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   5640
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.OptionButton optCalendarBackCol 
         Caption         =   "CalendarBackCol"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   2175
      End
      Begin VB.OptionButton optCalendarGradientCol 
         Caption         =   "CalendarGradientCol"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2280
         Width           =   2295
      End
      Begin VB.OptionButton optDayCol 
         Caption         =   "DayCol"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2760
         Width           =   2175
      End
      Begin VB.OptionButton optWeekDayCol 
         Caption         =   "WeekDayCol"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3600
         Width           =   2175
      End
      Begin VB.OptionButton optWeekDaySelCol 
         Caption         =   "WeekDaySelCol"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3840
         Width           =   2175
      End
      Begin VB.OptionButton optWeekDaySunCol 
         Caption         =   "WeekDaySunCol"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4080
         Width           =   2175
      End
      Begin VB.OptionButton optDaySunCol 
         Caption         =   "DaySunCol"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3240
         Width           =   2175
      End
      Begin VB.OptionButton optDaySelCol 
         Caption         =   "DaySelCol"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optYearGradientCol 
         Caption         =   "YearGradientCol"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton optYearBackCol 
         Caption         =   "YearBackCol"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton optMonthGradientCol 
         Caption         =   "MonthGradientCol"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton optMonthBackCol 
         Caption         =   "MonthBackCol"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton optHeaderGradientCol 
         Caption         =   "HeaderGradientCol"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optHeaderBackCol 
         Caption         =   "HeaderBackCol"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         Picture         =   "frmTest.frx":0169
         ScaleHeight     =   405
         ScaleWidth      =   2145
         TabIndex        =   13
         Top             =   4440
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Can Show outside it's container on dropdown mode"
      Height          =   615
      Left            =   120
      TabIndex        =   41
      Top             =   6120
      Width           =   9975
      Begin AdvancedCalendar.McCalendar McCalendar2 
         Height          =   270
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Theme           =   6
         Animate         =   -1  'True
         CalendarHeight  =   420
         Mode            =   2
         CalendarBackCol =   10985207
         MonthBackCol    =   10985207
         HeaderBackCol   =   10985207
         WeekDayCol      =   11446008
         DayCol          =   16777215
         DaySelCol       =   14934998
         WeekDaySelCol   =   15857131
         DaySunCol       =   14078715
         WeekDaySunCol   =   14078715
         YearBackCol     =   10985207
         SpecialDays     =   "21-1,26-1,19-2,8-3,24-3,25-3,14-4,21-4,13-8,15-8,20-8,26-8,14-9,15-9,16-9,17-9,21-9,11-10,12-10,1-11,3-11,10-12,25-12"
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Special Dates (String)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   53
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lbCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Themes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lbDate 
      AutoSize        =   -1  'True
      Caption         =   "N/a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   450
   End
   Begin VB.Label lbYear 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "N/a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4320
      TabIndex        =   5
      Top             =   3720
      Width           =   450
   End
   Begin VB.Label lbMonth 
      AutoSize        =   -1  'True
      Caption         =   "N/a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   450
   End
   Begin VB.Label lbDay 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "N/a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4320
      TabIndex        =   3
      Top             =   3480
      Width           =   450
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAnimate_Click()
    McCalendar1.Animate = chkAnimate
End Sub

Private Sub chkAppearance_Click()
    McCalendar1.Appearance = chkAppearance
End Sub

Private Sub chkBorder_Click()
    McCalendar1.BorderStyle = chkBorder
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub chkGradient_Click(Index As Integer)
Select Case Index
    Case 0
        McCalendar1.HeaderGradient = chkGradient(0)
    Case 1
        McCalendar1.MonthGradient = chkGradient(1)
    Case 2
        McCalendar1.YearGradient = chkGradient(2)
    Case 3
        McCalendar1.CalendarGradient = chkGradient(3)
End Select
End Sub

Private Sub chkSensitive_Click()
    McCalendar1.Sensitive = chkSensitive
End Sub

Private Sub chkSkip_Click()
    McCalendar1.SkipEnabled = chkSkip
End Sub

Private Sub cmbFormat_Click()
    McCalendar1.DateFormat = cmbFormat.ListIndex
End Sub

Private Sub lstMode_Click()
    McCalendar1.Mode = lstMode.ListIndex
End Sub

Private Sub lstTheme_Click()
    McCalendar1.Theme = lstTheme.ListIndex + 1
End Sub

Private Sub McCalendar1_DateChanged()
    lbDate = "Date " & McCalendar1
    lbDay = "Day " & McCalendar1.DayX
    lbMonth = "Month " & McCalendar1.MonthX
    lbYear = "Year " & McCalendar1.YearX
    lbCaption = McCalendar1.Caption(True)
End Sub

Private Sub McCalendar3_DateChanged()
    txtDatePick = McCalendar3.DateX
End Sub

Private Sub McCalendar4_DateChanged()
    txtDtaePick2 = McCalendar4.DateX
End Sub

Private Sub picSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If optHeaderBackCol Then McCalendar1.HeaderBackCol = picSel.Point(X, Y)
    If optHeaderGradientCol Then McCalendar1.HeaderGradientCol = picSel.Point(X, Y)
    
    If optMonthBackCol Then McCalendar1.MonthBackCol = picSel.Point(X, Y)
    If optMonthGradientCol Then McCalendar1.MonthGradientCol = picSel.Point(X, Y)
    
    If optYearBackCol Then McCalendar1.YearBackCol = picSel.Point(X, Y)
    If optYearGradientCol Then McCalendar1.YearGradientCol = picSel.Point(X, Y)
    
    If optCalendarBackCol Then McCalendar1.CalendarBackCol = picSel.Point(X, Y)
    If optCalendarGradientCol Then McCalendar1.CalendarGradientCol = picSel.Point(X, Y)
    
    If optDayCol Then McCalendar1.DayCol = picSel.Point(X, Y)
    If optDaySelCol Then McCalendar1.DaySelCol = picSel.Point(X, Y)
    If optDaySunCol Then McCalendar1.DaySunCol = picSel.Point(X, Y)
    
    If optWeekDayCol Then McCalendar1.WeekDayCol = picSel.Point(X, Y)
    If optWeekDaySelCol Then McCalendar1.WeekDaySelCol = picSel.Point(X, Y)
    If optWeekDaySunCol Then McCalendar1.WeekDaySunCol = picSel.Point(X, Y)

End Sub

Private Sub txtCurvature_Change()
    McCalendar1.Curvature = Val(txtCurvature)
End Sub

Private Sub txtHeader_Change()
    McCalendar1.HeaderHeight = Val(txtHeader)
End Sub

Private Sub txtHeight_Change()
    McCalendar1.CalendarHeight = Val(txtHeight)
End Sub

Private Sub txtSpecial_Change()
    McCalendar1.SpecialDays = txtSpecial
End Sub
