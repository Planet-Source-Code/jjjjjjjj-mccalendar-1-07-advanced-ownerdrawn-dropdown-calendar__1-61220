VERSION 5.00
Begin VB.UserControl McCalendar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   FillColor       =   &H00257A4B&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MouseIcon       =   "McCalendar.ctx":0000
   ScaleHeight     =   69
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   59
   ToolboxBitmap   =   "McCalendar.ctx":030A
   Begin VB.PictureBox picCalendar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00257A4B&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   30
      MouseIcon       =   "McCalendar.ctx":061C
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "McCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^Gtech Creations >^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^¶¶^^^^^^¶¶^^^^^^^^¶¶¶¶¶^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^¶¶^^^¶¶^^^^¶¶^^^^^^$
'$^^^¶¶¶^^^^¶¶¶^^^^^^^¶¶^^^^¶^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^¶¶^^^¶¶^^¶¶¶¶^^^^^^$
'$^^^¶¶¶¶^^¶¶¶¶^^¶¶¶¶^¶¶^^^^¶^^^¶¶¶^^¶¶^^¶¶¶¶^^¶¶^¶¶¶^^^¶¶¶¶¶¶^^¶¶¶¶^^¶¶^¶^^^^^¶¶^^^¶¶^^^^¶¶^^^^^^$
'$^^^¶^¶¶¶¶¶^¶¶^¶¶^^^^¶¶^^^^^^^¶^^¶¶^¶¶^¶¶^^¶¶^¶¶¶^^¶¶^¶¶^^^¶¶^¶¶^^¶¶^¶¶¶¶^^^^^^¶¶^¶¶^^^^^¶¶^^^^^^$
'$^^^¶^^¶¶¶^^¶¶^¶¶^^^^¶¶^^^^^^^^^^¶¶^¶¶^¶¶^^¶¶^¶¶^^^¶¶^¶¶^^^¶¶^¶¶^^¶¶^¶¶^^^^^^^^¶¶^¶¶^^^^^¶¶^^^^^^$
'$^^^¶^^^¶^^^¶¶^¶¶^^^^¶¶^^^^^^^¶¶¶¶¶^¶¶^¶¶¶¶¶¶^¶¶^^^¶¶^¶¶^^^¶¶^¶¶¶¶¶¶^¶¶^^^^^^^^¶¶^¶¶^^^^^¶¶^^^^^^$
'$^^^¶^^^^^^^¶¶^¶¶^^^^¶¶^^^^¶^¶¶^^¶¶^¶¶^¶¶^^^^^¶¶^^^¶¶^¶¶^^^¶¶^¶¶^^^^^¶¶^^^^^^^^^¶¶¶^^^^^^¶¶^^^^^^$
'$^^^¶^^^^^^^¶¶^¶¶^^^^¶¶^^^^¶^¶¶^^¶¶^¶¶^¶¶^^^¶^¶¶^^^¶¶^¶¶^^¶¶¶^¶¶^^^¶^¶¶^^^^^^^^^¶¶¶^^^^^^¶¶^^^^^^$
'$^^^¶^^^^^^^¶¶^^¶¶¶¶^^¶¶¶¶¶^^^¶¶¶¶¶^¶¶^^¶¶¶¶^^¶¶^^^¶¶^^¶¶¶^¶¶^^¶¶¶¶^^¶¶^^^^^^^^^¶¶¶^^^^¶¶¶¶¶¶^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^ By 'Jim Jose' ^^^^^^^^ Email - jimjosev33@yahoo.com ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


'--------------------------------------------------------------------------------------------------
' Source Code   : McCalendar ActiveX Control
' Auther        : Jim Jose
' eMail         : jimjosev33@yahoo.com
' Purpose       : An Ownerdrawn, Sizable, Stylish , List/DropDown Calendar
' BasicWorking  : The DropDown mode 'Exports' the'picCalendar' to the parent form
'               : using 'SetParent' API call
'---------------------------------------------------------------------------------------------------
'
'About The Control
'-----------------
'       I coded this control as a complete solution for fully customizable
' and resizable calendar control. Normally all the Calendar controls we can
' found in PSC are a huge stack of controls and control arrays, and thus the
' control size will be extreamly high. I used only one picturebox for building
' this control. I think it could be justified because we need an hwnd container
' to export the calendar to the parent form.
'
'       The calendar is Owner draw which enables it to draw calenders of any size
' and it also offers high performance and speed.
'
'       The calender is provided with two MODES. You can use this control as a
' 'dropdown' calendar and as a 'Listed' calendar. The calendar colors are fully
' customizable and also provided with eight different themes.
'
'       A unique animation technique is used in this control(i already posted that
' a an another submission and the response was inspiring).
'
'Upgrade Destinations
'--------------------
'       I think the control is error free now. If there is any hidden bugs or
' any operatingtime errors please inform me, so that I can correct it soon.
'
'       If you need any aditional properties/options feel free to inform me.
' You can also use the email, if you have any doubts(even if the code is fully
' commented)
'
'       I have a plan to upgrade the control by removing the Picturebox and
' use CreateWindowEx Api for that. I am not familier with this method. So I
' welcomes comments from experienced hands on this matter.

'Thanks
'------
'   Carles P.V. of his fast gradient routine
'   Ken Foster and Ben Vonk for the orginal idea (not rented a bit of code)
'   PSC, I learned fully from there
'
'History:
'--------
'
' Vesion 1.0
' Submitted to PSC 15-6-2005
'
' Update Vesion 1.1
' Resubmitted the control with full range of color options.
' Improved themes.
'
' Updated Version 1.2
' Added additional features sugested by Ken Forster. This verion have
' three modes 1.List 2.PopUp 3.PopDown. Also clicking on Today Region
' will set calenar back to current day.
'
' Updated Version 1.3
' Added additional features sugested by Ruturag. Added two more
' properties, 'Sensitive'(for popup/down mode) and 'SkipEnabled'.
' The calendar will close when clicking on cross-filled dates if
' 'Sensitive' is true. The popdown/up arrow will reverse (as Guturaj
' suggested) only if 'Sensitive' is false.
' If 'SkipEnabled' is true, then the calendar will skip into next or
' last month (according to the back or end of days clicked)
'
' Updated Version 1.4
' Added DatePicker Compatable mode as suggested by Dennis.
' This version have two more modes. 1.Datepicker PopUp
' 2. DatePicker PopDown. These two modes only show the
' left popDown buttons. You can place it near the TextBox
' to which the dates must send (see the sample).
' I used this method to get the functionality, bcose otherwise
' we needs to add a additional textbox into the control only
' for this purpose. The usercontrol is not resized to the
' popDown button size. This is bcose u can use the Usercontrol
' width to adjust the calendar width, otherwise  u may need to use a
' property, the earlier is more sensible.
' This Version also contains one more property 'Header Height', which is
' needed to adjust the header height to the DatePicker's TextBox Height
'
' Updated Version 1.5
' Added one more event function. DbClick on Days will close the Calendar
' (except ListMode)
'
' Updated Version 1.6
' Added Property special days. You can add days to this property as shon bellow
' 21-1,26-1,19-2,8-3,24-3,25-3,14-4,21-4,13-8,15-8,20-8,26-8,14-9,15-9,16-9,17-9,21-9,11-10,12-10,1-11,3-11,10-12,25-12
' This is the complete list of holydays for my country(India)(some of them are only for my state)
' The days are added as 'day-month' and seperated by ",". These days are applicable to all
' years. The Special days will be indicated by a Rectangle halfly-filled on it's cell.
' See JAN - 26 (republic day)
'
' Updated Version 1.7
' Language problem solved. Now the calendar will load the MonthNames and the WeekDayNames
' according to the corrent language selection of the user. Thanks to Cote for his attension on this part.
'
'---------------------------------------------------------------------------------------------------

Option Explicit

'[Apis]
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function BringWindowToTop Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function Ellipse Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'[Enums]
Public Enum GradientDirectionCts
    [Fill_None] = 0
    [Fill_Horizontal] = 1
    [Fill_Vertical] = 2
    [Fill_DownwardDiagonal] = 3
    [Fill_UpwardDiagonal] = 4
End Enum

Public Enum AppearanceConstants
    [Flat] = 0
    [3D] = 1
End Enum

Public Enum BorderStyleEnum
    [None] = 0
    [Fixel Single] = 1
End Enum

Public Enum CalendarMode
    [List Mode] = 0
    [PopDown Mode] = 1
    [PopUp Mode] = 2
    [DatePicker PopDown] = 3
    [DatePicker Popup] = 4
End Enum

Public Enum CalendarTheme
    [Cal_Attraction] = 1
    [Cal Blue] = 2
    [Cal Green] = 3
    [Cal Orange] = 4
    [Cal Purple] = 5
    [Cal Red] = 6
    [Cal Silver] = 7
    [Cal Yellow] = 8
End Enum

Private Enum ArrowDir
    [Arw_Left] = 0
    [Arw_Right] = 1
    [Arw_Up] = 2
    [Arw_Down] = 3
End Enum

Private Enum AnimeEventEnum
    aUnload = 0
    aload = 1
End Enum

Private Enum AnimeEffectEnum
    eAppearFromLeft = 0
    eAppearFromRight = 1
    eAppearFromTop = 2
    eAppearFromBottom = 3
End Enum

Public Enum DateFormatEnum
    [dd-mm-yyyy] = 0
    [mm-dd-yyyy] = 1
    [yyyy-mm-dd] = 2
End Enum

'[Types]
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

'[Property Variables:]
Private m_SelYear As Long
Private m_SelMonth As Long
Private m_SelDay As Long

Private m_Appearance As Integer
Private m_BorderStyle As Integer
Private m_Enabled As Boolean
Private m_Font As Font
Private m_Theme As CalendarTheme
Private m_Animate As Boolean
Private m_Mode As CalendarMode
Private m_CalendarHeight As Long
Private m_Curvature  As Long
Private m_CalendarGradient As Boolean
Private m_CalendarBackCol As OLE_COLOR
Private m_SkipEnabled As Boolean
Private m_Sensitive As Boolean

Private m_TrackWidth As Long
Private m_iHeight   As Double
Private m_iWidth    As Double
Private m_HeaderHeight As Long
Private m_WeekDaysHeight As Long
Private m_Right2Left As Boolean
Private m_MouseX     As Single
Private m_MouseY     As Single
Private m_FirstDay   As Long
Private m_MonthDays   As Long
Private m_MonthPopupMode  As Boolean
Private m_MonthPopWidth As Double
Private m_Poped     As Boolean
Private m_SpecialDayStack() As String
Private m_SpecialDays As String

Private m_MonthBackCol As OLE_COLOR
Private m_DayCol As OLE_COLOR
Private m_DaySunCol As OLE_COLOR
Private m_WeekDaySunCol As OLE_COLOR
Private m_WeekDayCol As OLE_COLOR
Private m_ArrowCol As OLE_COLOR
Private m_WeekDaySelCol As OLE_COLOR
Private m_DaySelCol As OLE_COLOR
Private m_DateFormat As DateFormatEnum
Private m_YearBackCol As OLE_COLOR
Private m_YearGradient As Boolean
Private m_YearGradientCol As OLE_COLOR
Private m_HeaderGradientCol As OLE_COLOR
Private m_MonthGradientCol As OLE_COLOR
Private m_CalendarGradientCol As OLE_COLOR
Private m_MonthGradient As Boolean
Private m_HeaderGradient As Boolean
Private m_HeaderBackCol As OLE_COLOR

'[Default Property Values:]
Private Const m_def_Appearance = [Flat]
Private Const m_def_BorderStyle = [Fixel Single]
Private Const m_def_Enabled = True
Private Const m_def_Theme = Cal_Attraction
Private Const m_def_Animate = False
Private Const m_def_Mode = [List Mode]
Private Const m_def_CalendarHeight = 125
Private Const DIB_RGB_ColS As Long = 0
Private Const m_def_Curvature = 0
Private Const m_Def_CalendarGradient = True
Private Const m_def_CalendarBackCol = &HFFFFFF
Private Const m_def_SkipEnabled = False
Private Const m_def_Sensitive = True
Private Const m_def_HeaderHeight = 18

Private Const m_def_DateFormat = 0
Private Const m_def_WeekDayCol = &HFF9A35
Private Const m_def_DayCol = &HFDDBAC
Private Const m_def_DaySelCol = &HC4F9F9
Private Const m_def_WeekDaySelCol = &H59B4CA
Private Const m_def_DaySunCol = &HCAB7FD
Private Const m_def_WeekDaySunCol = &H8080FF
Private Const m_def_MonthGradient = True
Private Const m_def_HeaderGradient = True
Private Const m_def_MonthBackCol = &HFF7D7D
Private Const m_def_HeaderBackCol = &HFF7D7D
Private Const m_def_HeaderGradientCol = &HFFFFFF
Private Const m_def_MonthGradientCol = &HFFFFFF
Private Const m_def_CalendarGradientCol = &HFFFFFF
Private Const m_def_YearBackCol = &HFF7D7D
Private Const m_def_YearGradient = False
Private Const m_def_YearGradientCol = &HFFFFFF
Private Const m_def_SpecialDays = ""

Private Const m_Months = 12
Private Const m_HeaderDays = 7
Private Const m_RowDays = 5

Private Const RGN_AND As Long = 1
Private Const RGN_OR As Long = 2
Private Const RGN_XOR As Long = 3
Private Const RGN_COPY As Long = 5
Private Const RGN_DIFF As Long = 4

'[Event Declarations:]
Public Event DateChanged()



Private Sub ApplyTheme(ByVal ThemeIndex As CalendarTheme)

Debug.Print "Applying new theme "
Select Case ThemeIndex
    Case [Cal_Attraction]
        m_HeaderBackCol = &HFF7D7D
        m_ArrowCol = &H257A4B
        m_MonthBackCol = &HFF7D7D
        m_DayCol = &HFDDBAC
        m_DaySunCol = &HCAB7FD
        m_WeekDayCol = &HFF9A35
        m_WeekDaySunCol = &H8080FF
        m_WeekDaySelCol = &H59B4CA
        m_DaySelCol = &HC4F9F9
        picCalendar.ForeColor = 0
        
    Case [Cal Blue]
        m_HeaderBackCol = &HDABAA8
        m_ArrowCol = &HDCC1AD
        m_MonthBackCol = &HEDC5A7
        m_DayCol = &HFCF4EF
        m_DaySunCol = &HAEA6F8
        m_WeekDayCol = &HFAE8DC
        m_WeekDaySunCol = &H8080FF
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HD8E5C8
        picCalendar.ForeColor = 0 ' &H864E02
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Green]
        m_HeaderBackCol = &HB1CB92
        m_ArrowCol = &H213B00
        m_MonthBackCol = &HB1CB92
        m_DayCol = &HE1EBD5
        m_DaySunCol = &HAEA6F8
        m_WeekDaySunCol = &H8080FF
        m_WeekDayCol = &HFAE8DC
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HDABAA8
        picCalendar.ForeColor = 0 ' &H185232
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Orange]
        m_HeaderBackCol = &HD2E2FD
        m_ArrowCol = &H16366D
        m_MonthBackCol = &HD2E2FD
        m_DayCol = &HEFF5FE
        m_DaySunCol = &HE3E3D6
        m_WeekDaySunCol = &HD5B0BF
        m_WeekDayCol = &HE3E3D6
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HB1CB92
        picCalendar.ForeColor = 0 '&H80FF&
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Purple]
        m_HeaderBackCol = &HD5B0BF
        m_ArrowCol = &H46202F
        m_MonthBackCol = &HD5B0BF
        m_DayCol = &HF7F1F3
        m_DaySunCol = &HB1CB92
        m_WeekDaySunCol = &HD5B0BF
        m_WeekDayCol = &HD1A9B9
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HE3E3D6
        picCalendar.ForeColor = 0 '&HE4616D
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Red]
        m_HeaderBackCol = &HAEA6F8
        m_ArrowCol = &H1D156A
        m_MonthBackCol = &HA79EF7
        m_DayCol = &HFFFFFF
        m_DaySunCol = &HD6D2FB
        m_WeekDaySunCol = &HD6D2FB
        m_WeekDayCol = &HAEA6F8
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HE3E3D6
        picCalendar.ForeColor = 0 ' &HC0&
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Silver]
        m_HeaderBackCol = &HD9D6D3
        m_ArrowCol = &H4A4744
        m_MonthBackCol = &HD9D6D3
        m_DayCol = &HFFFFFF
        m_DaySunCol = &HD6D2FB
        m_WeekDaySunCol = &HD6D2FB
        m_WeekDayCol = &HD9D6D3
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HE3E3D6
        picCalendar.ForeColor = 0 '&H808080
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Yellow]
        m_HeaderBackCol = &HB9EEF4
        m_ArrowCol = &H66D5E1
        m_MonthBackCol = &HB9EEF4
        m_DayCol = &HFFFFFF
        m_DaySunCol = &HD6D2FB
        m_WeekDaySunCol = &HD6D2FB
        m_WeekDayCol = &HB9EEF4
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HE3E3D6
        picCalendar.ForeColor = 0 ' &H57C9E
        m_HeaderBackCol = m_MonthBackCol
        
End Select

m_MonthGradientCol = m_HeaderGradientCol
m_CalendarGradientCol = m_HeaderGradientCol
m_CalendarBackCol = m_MonthBackCol
m_YearBackCol = m_MonthBackCol
m_CalendarGradientCol = m_MonthGradientCol

End Sub


Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    m_Appearance = New_Appearance
    UserControl.Appearance = New_Appearance
    PropertyChanged "Appearance"
    RedrawControl
End Property


Public Property Get BorderStyle() As BorderStyleEnum
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleEnum)
    m_BorderStyle = New_BorderStyle
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    RedrawControl
End Property


Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl.Enabled = New_Enabled
End Property


Public Property Get CalendarGradient() As Boolean
    CalendarGradient = m_CalendarGradient
End Property

Public Property Let CalendarGradient(ByVal vNewValue As Boolean)
    m_CalendarGradient = vNewValue
    PropertyChanged "CalendarGradient"
    RedrawControl
End Property


Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    Set UserControl.Font = New_Font
    Set picCalendar.Font = New_Font
    PropertyChanged "Font"
    UserControl_Resize
End Property


Public Property Get SpecialDays() As String
    SpecialDays = m_SpecialDays
End Property

Public Property Let SpecialDays(ByVal New_SpecialDays As String)
    
    ' This property is read only at runtime
    If Ambient.UserMode Then Err.Raise 382
    m_SpecialDays = New_SpecialDays
    PropertyChanged "SpecialDays"
    If Not m_SpecialDays = vbNullString Then m_SpecialDayStack = Split(m_SpecialDays, ",")
    RedrawControl

End Property

Public Property Get Sensitive() As Boolean
    Sensitive = m_Sensitive
End Property

Public Property Let Sensitive(ByVal New_Sensitive As Boolean)
    m_Sensitive = New_Sensitive
    PropertyChanged "Sensitive"
End Property


Public Property Get SkipEnabled() As Boolean
    SkipEnabled = m_SkipEnabled
End Property

Public Property Let SkipEnabled(ByVal New_SkipEnabled As Boolean)
    m_SkipEnabled = New_SkipEnabled
    PropertyChanged "SkipEnabled"
End Property


Public Property Get Theme() As CalendarTheme
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As CalendarTheme)
    m_Theme = New_Theme
    PropertyChanged "Theme"
    ApplyTheme New_Theme
    RedrawControl
End Property


Public Property Get Animate() As Boolean
    Animate = m_Animate
End Property

Public Property Let Animate(ByVal New_Animate As Boolean)
    m_Animate = New_Animate
    PropertyChanged "Animate"
End Property


Public Property Get CalendarHeight() As Long
    CalendarHeight = m_CalendarHeight
End Property

Public Property Let CalendarHeight(ByVal vNewValue As Long)
    m_CalendarHeight = vNewValue
    If m_Mode = [List Mode] Then Height = (m_CalendarHeight + m_HeaderHeight) * Screen.TwipsPerPixelY
    PropertyChanged "CalendarHeight"
    RedrawControl
End Property


Public Property Get Caption(ByVal LongDate As Boolean) As String
Dim vMonth As String
vMonth = MonthName(m_SelMonth)
    
    If LongDate Then
        Caption = m_SelDay & " " & vMonth & " " & m_SelYear & "  '" & WeekdayName(Weekday(DateSerial(m_SelYear, m_SelMonth, m_SelDay))) & "'   "
    Else
        Caption = m_SelDay & " " & vMonth & " " & m_SelYear & "  '" & UCase(Left$(WeekdayName(Weekday(DateSerial(m_SelYear, m_SelMonth, m_SelDay))), 3)) & "'   "
    End If
    
End Property


Public Property Get Curvature() As Long
    Curvature = m_Curvature
End Property

Public Property Let Curvature(ByVal vNewValue As Long)
    m_Curvature = vNewValue
    PropertyChanged "Curvature"
    RedrawControl
End Property


Public Property Get DateX() As String
Attribute DateX.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute DateX.VB_UserMemId = 0
Attribute DateX.VB_MemberFlags = "200"
    DateX = Format(DateSerial(m_SelYear, m_SelMonth, m_SelDay), GetFormat)
End Property

Public Property Let YearX(ByVal vNewValue As Long)
    m_SelYear = vNewValue
    PropertyChanged "YearX"
    LoadDay (999)
    RedrawControl
End Property


Public Property Get YearX() As Long
    YearX = m_SelYear
End Property

Public Property Get MonthX() As Long
    MonthX = m_SelMonth
End Property

Public Property Let MonthX(ByVal vNewValue As Long)
    m_SelMonth = vNewValue
    PropertyChanged "MonthX"
    LoadDay (999)
    RedrawControl
End Property


Public Property Get DayX() As Long
    DayX = m_SelDay
End Property

Public Property Let DayX(ByVal vNewValue As Long)
    m_SelDay = vNewValue
    PropertyChanged "DayX"
    LoadDay vNewValue
    RedrawControl
End Property


Public Property Get Mode() As CalendarMode
    Mode = m_Mode
End Property

Public Property Let Mode(ByVal vNewValue As CalendarMode)
    m_Mode = vNewValue
    PropertyChanged "Mode"
    If Not m_Mode = [List Mode] Then
        Height = m_HeaderHeight * Screen.TwipsPerPixelX
        picCalendar.Visible = False
    Else
        Height = (m_HeaderHeight + m_CalendarHeight) * Screen.TwipsPerPixelX
        ImportCalendar
    End If
    UserControl_Resize
End Property


Public Property Get CalendarBackCol() As OLE_COLOR
    CalendarBackCol = m_CalendarBackCol
End Property

Public Property Let CalendarBackCol(ByVal vNewValue As OLE_COLOR)
    m_CalendarBackCol = vNewValue
    PropertyChanged "CalendarBackCol"
    RedrawControl
End Property


Public Property Get MonthGradient() As Boolean
    MonthGradient = m_MonthGradient
End Property

Public Property Let MonthGradient(ByVal New_MonthGradient As Boolean)
    m_MonthGradient = New_MonthGradient
    PropertyChanged "MonthGradient"
    RedrawControl
End Property


Public Property Get HeaderGradient() As Boolean
    HeaderGradient = m_HeaderGradient
End Property

Public Property Let HeaderGradient(ByVal New_HeaderGradient As Boolean)
    m_HeaderGradient = New_HeaderGradient
    PropertyChanged "HeaderGradient"
    RedrawControl
End Property


Public Property Get MonthBackCol() As OLE_COLOR
    MonthBackCol = m_MonthBackCol
End Property

Public Property Let MonthBackCol(ByVal New_MonthBackCol As OLE_COLOR)
    m_MonthBackCol = New_MonthBackCol
    PropertyChanged "MonthBackCol"
    RedrawControl
End Property


Public Property Get HeaderBackCol() As OLE_COLOR
    HeaderBackCol = m_HeaderBackCol
End Property

Public Property Let HeaderBackCol(ByVal New_HeaderBackCol As OLE_COLOR)
    m_HeaderBackCol = New_HeaderBackCol
    PropertyChanged "HeaderBackCol"
    RedrawControl
End Property


Public Property Get WeekDayCol() As OLE_COLOR
    WeekDayCol = m_WeekDayCol
End Property

Public Property Let WeekDayCol(ByVal New_WeekDayCol As OLE_COLOR)
    m_WeekDayCol = New_WeekDayCol
    PropertyChanged "WeekDayCol"
    RedrawControl
End Property


Public Property Get DayCol() As OLE_COLOR
    DayCol = m_DayCol
End Property

Public Property Let DayCol(ByVal New_DayCol As OLE_COLOR)
    m_DayCol = New_DayCol
    PropertyChanged "DayCol"
    RedrawControl
End Property


Public Property Get DaySelCol() As OLE_COLOR
    DaySelCol = m_DaySelCol
End Property

Public Property Let DaySelCol(ByVal New_DaySelCol As OLE_COLOR)
    m_DaySelCol = New_DaySelCol
    PropertyChanged "DaySelCol"
    RedrawControl
End Property


Public Property Get WeekDaySelCol() As OLE_COLOR
    WeekDaySelCol = m_WeekDaySelCol
End Property

Public Property Let WeekDaySelCol(ByVal New_WeekDaySelCol As OLE_COLOR)
    m_WeekDaySelCol = New_WeekDaySelCol
    PropertyChanged "WeekDaySelCol"
    RedrawControl
End Property


Public Property Get HeaderHeight() As Long
    HeaderHeight = m_HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal New_HeaderHeight As Long)
    m_HeaderHeight = New_HeaderHeight
    PropertyChanged "HeaderHeight"
    UserControl_Resize
End Property


Public Property Get DaySunCol() As OLE_COLOR
    DaySunCol = m_DaySunCol
End Property

Public Property Let DaySunCol(ByVal New_DaySunCol As OLE_COLOR)
    m_DaySunCol = New_DaySunCol
    PropertyChanged "DaySunCol"
    RedrawControl
End Property


Public Property Get WeekDaySunCol() As OLE_COLOR
    WeekDaySunCol = m_WeekDaySunCol
End Property

Public Property Let WeekDaySunCol(ByVal New_WeekDaySunCol As OLE_COLOR)
    m_WeekDaySunCol = New_WeekDaySunCol
    PropertyChanged "WeekDaySunCol"
    RedrawControl
End Property


Public Property Get HeaderGradientCol() As OLE_COLOR
    HeaderGradientCol = m_HeaderGradientCol
End Property

Public Property Let HeaderGradientCol(ByVal New_HeaderGradientCol As OLE_COLOR)
    m_HeaderGradientCol = New_HeaderGradientCol
    PropertyChanged "HeaderGradientCol"
    RedrawControl
End Property


Public Property Get MonthGradientCol() As OLE_COLOR
    MonthGradientCol = m_MonthGradientCol
End Property

Public Property Let MonthGradientCol(ByVal New_MonthGradientCol As OLE_COLOR)
    m_MonthGradientCol = New_MonthGradientCol
    PropertyChanged "MonthGradientCol"
    RedrawControl
End Property


Public Property Get CalendarGradientCol() As OLE_COLOR
    CalendarGradientCol = m_CalendarGradientCol
End Property

Public Property Let CalendarGradientCol(ByVal New_CalendarGradientCol As OLE_COLOR)
    m_CalendarGradientCol = New_CalendarGradientCol
    PropertyChanged "CalendarGradientCol"
    RedrawControl
End Property


Public Property Get YearBackCol() As OLE_COLOR
    YearBackCol = m_YearBackCol
End Property

Public Property Let YearBackCol(ByVal New_YearBackCol As OLE_COLOR)
    m_YearBackCol = New_YearBackCol
    PropertyChanged "YearBackCol"
    RedrawControl
End Property


Public Property Get YearGradient() As Boolean
    YearGradient = m_YearGradient
End Property

Public Property Let YearGradient(ByVal New_YearGradient As Boolean)
    m_YearGradient = New_YearGradient
    PropertyChanged "YearGradient"
    RedrawControl
End Property


Public Property Get YearGradientCol() As OLE_COLOR
    YearGradientCol = m_YearGradientCol
End Property

Public Property Let YearGradientCol(ByVal New_YearGradientCol As OLE_COLOR)
    m_YearGradientCol = New_YearGradientCol
    PropertyChanged "YearGradientCol"
    RedrawControl
End Property


Public Property Get DateFormat() As DateFormatEnum
    DateFormat = m_DateFormat
End Property

Public Property Let DateFormat(ByVal New_DateFormat As DateFormatEnum)
    m_DateFormat = New_DateFormat
    PropertyChanged "DateFormat"
    RedrawControl
End Property


Private Sub picCalendar_DblClick()
    picCalendar_MouseDown vbLeftButton, 111, m_MouseX, m_MouseY
End Sub

Private Sub picCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nDay As Long

' This is the case when the user try to select a month from the
' popuped month list.
If m_MonthPopupMode Then

    If X < m_TrackWidth Then
        ' Not moving over the list. Nothing is to do here
    Else
        ' Calculate the Month that the User selected
        m_SelMonth = Int((X - m_TrackWidth) / m_MonthPopWidth + 1)
    End If
    
    ' Disable popup mode and draw new month
    m_MonthPopupMode = False
    LoadDay (999)
    RedrawControl
    Exit Sub
    
End If

' Moving over the Calendarn whaever may happen

' Moving through the left Month Display region
If X < m_TrackWidth Then
    
    ' Clecking over the Month Up button
    If Y < m_TrackWidth Then
        
        ' Select the last Month
        If m_SelMonth = 1 Then      'Back of year
            m_SelMonth = 12
            m_SelYear = m_SelYear - 1
        Else                         'No problem preceed
            m_SelMonth = m_SelMonth - 1
        End If
        
        ' Load the day, 999 checks the new Month posses Days more
        ' than selected date
        LoadDay (999)
        
    ' Clicking over Month Down Button
    ElseIf Y > picCalendar.ScaleHeight - m_TrackWidth Then
    
        ' Select the last Month
        If m_SelMonth = 12 Then     ' End of year
            m_SelMonth = 1
            m_SelYear = m_SelYear + 1
        Else                        ' No problem proceed
            m_SelMonth = m_SelMonth + 1
        End If
        
        ' Load the day, 999 checks the new Month posses Days more
        ' than selected date
        LoadDay (999)
        
    'Moving throgh left TrackRegion( Month Show).
    Else
    
        ' Clicking for popuping Month list
        If X < m_TrackWidth And X > 0.75 * m_TrackWidth Then
            PopupMonthList
            DrawBody
            Exit Sub
        End If
        
    End If

' Not throgh month display region.
' To the Right of that
Else
    
    ' Moving through Header ( Week days display )
    If Y < m_WeekDaysHeight Then
        ' No events added till this version
    
    ' Moving through "DAYS'
    Else
        
        ' Trying to select a Day
        If Y < picCalendar.ScaleHeight - m_iHeight Then
        
            ' Calculate Day\Load it
            If Shift = 111 Then ' Event was send from DbClick
                If Not m_Mode = [List Mode] Then CollapseCalendar
            Else
                nDay = Int((Y - m_WeekDaysHeight) / m_iHeight) * m_HeaderDays + (Int((X - m_TrackWidth) / m_iWidth)) + 2 - m_FirstDay
                LoadDay nDay
            End If
        'Year selection region
        Else
            
            'Next Year Selecting Button
            If X > picCalendar.ScaleWidth - m_iWidth + 10 Then
                
                ' Load Next year . RightClick will jump FIVE
                If Button = vbLeftButton Then m_SelYear = m_SelYear + 1 Else m_SelYear = m_SelYear + 5
                
                ' Load the day, 999 checks the new Month in new year posses Days more
                ' than selected date
                LoadDay (999)
                
            'Last Year Selecting Button
            ElseIf X > m_iWidth * 4 + m_TrackWidth And X < m_iWidth * 4 + m_TrackWidth + m_iWidth - 10 Then
                
                ' Load Lastyear
                If Button = vbLeftButton Then m_SelYear = m_SelYear - 1 Else m_SelYear = m_SelYear - 5
                
                ' Load the day, 999 checks the new Month in new year posses Days more
                ' than selected date
                LoadDay (999)
                
            Else
                ' Over the Today Region
                m_SelDay = Day(Date)
                m_SelMonth = Month(Date)
                m_SelYear = Year(Date)
                LoadDay m_SelDay
            End If
            
        End If
    End If
End If

' Store x,y to send from double click event/Redraw control
m_MouseX = X
m_MouseY = Y
RedrawControl

End Sub

Private Sub picCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nDay As String

' Refer MouseClick Event for more info

' Moving through popuped month list
If m_MonthPopupMode Then

    If X < m_TrackWidth Then
        picCalendar.ToolTipText = MonthName(m_SelMonth)
        picCalendar.MousePointer = vbNormal
    Else
        picCalendar.ToolTipText = MonthName(Int((X - m_TrackWidth) / m_MonthPopWidth) + 1)
        picCalendar.MousePointer = vbCustom
    End If
    
Exit Sub
End If


If X <= m_TrackWidth Then

    ' Month Selection Up
    If Y < m_TrackWidth Then
    
        picCalendar.ToolTipText = "Last Month"
        picCalendar.MousePointer = vbCustom

    ' Month Selection Down
    ElseIf Y > picCalendar.ScaleHeight - m_TrackWidth Then
    
        picCalendar.ToolTipText = "Next Month"
        picCalendar.MousePointer = vbCustom
        
    ' b/w both
    Else
    
        ' Popup button
        If X < m_TrackWidth And X > 0.75 * m_TrackWidth Then
        
            picCalendar.MousePointer = vbCustom
            picCalendar.ToolTipText = "Popup Month List >"
            
        Else
        
            picCalendar.MousePointer = vbNormal
            picCalendar.ToolTipText = MonthName(m_SelMonth)
            
        End If
        
    End If
    
Else

    ' week days
    If Y < m_WeekDaysHeight Then
    
        picCalendar.ToolTipText = WeekdayName(Int((X - 1 - m_TrackWidth) / m_iWidth) + 1)
        picCalendar.MousePointer = vbNormal
    
    ' Bellow
    Else
    
        'Through days
        If Y < picCalendar.ScaleHeight - m_iHeight Then
            
            ' Calculate day
            nDay = Int((Y - m_WeekDaysHeight) / m_iHeight) * m_HeaderDays + (Int((X - m_TrackWidth) / m_iWidth)) + 2 - m_FirstDay
            If nDay < 0 Then
                If m_MonthDays > (36 - m_FirstDay) Then
                    nDay = (35 - m_FirstDay) + (m_FirstDay + nDay)
                End If
            End If
            
            ' Clicking on Diagonal croseed days will unload Calendar if Sensitive=True
            If nDay > m_MonthDays Then
                 If m_Sensitive And Not m_Mode = [List Mode] Then picCalendar.ToolTipText = "Close": picCalendar.MousePointer = vbCustom: Exit Sub
                 If m_SkipEnabled Then picCalendar.ToolTipText = "Skip Next Month" Else picCalendar.ToolTipText = ""
            
            ElseIf nDay <= 0 Then
                 If m_Sensitive And Not m_Mode = [List Mode] Then picCalendar.ToolTipText = "Close": picCalendar.MousePointer = vbCustom: Exit Sub
                 If m_SkipEnabled Then picCalendar.ToolTipText = "Skip Last Month" Else picCalendar.ToolTipText = ""
            Else
                picCalendar.ToolTipText = "Day " & nDay
            End If
            picCalendar.MousePointer = vbCustom
            
        Else    ' footer
        
            ' Year selecting region
            If X > m_iWidth * 4 + m_TrackWidth Then
            
                'Last Year Button
                If X < m_iWidth * 4 + m_TrackWidth + m_iWidth - 10 Then
                
                    picCalendar.ToolTipText = "Last Year"
                    picCalendar.MousePointer = vbCustom
                    
                ' Next Year Button
                ElseIf X > picCalendar.ScaleWidth - m_iWidth + 10 Then
                
                    picCalendar.ToolTipText = "Next Year"
                    picCalendar.MousePointer = vbCustom
                    
                Else 'Middle
                
                    picCalendar.MousePointer = vbNormal
                    picCalendar.ToolTipText = "Year " & m_SelYear
                    
                End If
                
            Else 'Through Date Display
            
                picCalendar.MousePointer = vbCustom
                picCalendar.ToolTipText = "Today " & Format$(Date$, GetFormat)
            
            End If
        End If
    End If

End If

End Sub

Private Sub UserControl_DblClick()
    m_Right2Left = False: DrawBody
End Sub

Private Sub UserControl_Initialize()
    
    Debug.Print vbCrLf & "--------------------------------------" & vbCrLf & "New Compile" & vbCrLf & "--------------------------------------"
    m_SelDay = Day(Date)
    m_SelMonth = Month(Date)
    m_SelYear = Year(Date)
    m_Right2Left = True
    LoadDay m_SelDay
    
End Sub

Private Sub UserControl_InitProperties()

    Me.Appearance = m_def_Appearance
    Me.BorderStyle = m_def_BorderStyle
    Me.Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_Theme = m_def_Theme
    m_Animate = m_def_Animate
    m_CalendarHeight = m_def_CalendarHeight
    m_Curvature = m_def_Curvature
    m_Mode = m_def_Mode
    m_SelDay = Day(Date)
    m_SelMonth = Month(Date)
    m_SelYear = Year(Date)
    m_CalendarBackCol = m_def_CalendarBackCol
    m_MonthGradient = m_def_MonthGradient
    m_MonthBackCol = m_def_MonthBackCol
    m_HeaderBackCol = m_def_HeaderBackCol
    m_WeekDayCol = m_def_WeekDayCol
    m_DayCol = m_def_DayCol
    m_DaySelCol = m_def_DaySelCol
    m_WeekDaySelCol = m_def_WeekDaySelCol
    m_DaySunCol = m_def_DaySunCol
    m_WeekDaySunCol = m_def_WeekDaySunCol
    m_HeaderGradientCol = m_def_HeaderGradientCol
    m_CalendarGradientCol = m_def_CalendarGradientCol
    m_YearBackCol = m_def_YearBackCol
    m_YearGradient = m_def_YearGradient
    m_YearGradientCol = m_def_YearGradientCol
    
    ApplyTheme Cal_Attraction
    m_MonthGradientCol = m_def_MonthGradientCol
    m_HeaderGradient = m_def_HeaderGradient
    m_CalendarGradient = m_Def_CalendarGradient
    m_DateFormat = m_def_DateFormat
    
    m_Sensitive = m_def_Sensitive
    m_SkipEnabled = m_def_SkipEnabled
    m_HeaderHeight = m_def_HeaderHeight
    
    m_SpecialDays = m_def_SpecialDays
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Clicking on PopDown button
    If X > ScaleWidth - (m_HeaderHeight + 10) And m_Mode = [PopDown Mode] Then

        If Not m_Poped Then
            m_Right2Left = False
            DrawBody
            
            ' Export Calendar up/down
            ExportCalendar True
            DrawCalendar
        Else
            CollapseCalendar
            DrawBody
        End If
        
    ' Clicking on PopUp button
    ElseIf X < (m_HeaderHeight + 10) And (m_Mode = [PopUp Mode] Or m_Mode = [DatePicker PopDown] Or m_Mode = [DatePicker Popup]) Then
        
        If Not m_Poped Then
            m_Right2Left = False
            DrawBody
            
            ' Export Calendar up/down
            If m_Mode = [DatePicker PopDown] Then
                ExportCalendar True
            Else
                ExportCalendar False
            End If
            DrawCalendar
        Else
            CollapseCalendar
            DrawBody
        End If
        
    End If

Exit Sub
End Sub


Public Sub CollapseCalendar()

    If m_Animate Then
        If m_Mode = [PopDown Mode] Or m_Mode = [DatePicker PopDown] Then
            AnimateForm picCalendar, aUnload, eAppearFromBottom, 10, 22
        Else
            AnimateForm picCalendar, aUnload, eAppearFromTop, 10, 22
        End If
    End If
    picCalendar.Visible = False
    
    ' This is necessary to redefine the Calendar region to full size
    AnimateForm picCalendar, aload, eAppearFromLeft, 0, 1
    m_Poped = False
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Over the button
    If X > ScaleWidth - (m_HeaderHeight + 10) And m_Mode = [PopDown Mode] Then
        UserControl.MousePointer = vbCustom
    ElseIf X < (m_HeaderHeight + 10) And (m_Mode = [PopUp Mode] Or m_Mode = [DatePicker PopDown] Or m_Mode = [DatePicker Popup]) Then
        UserControl.MousePointer = vbCustom
    Else
        UserControl.MousePointer = vbNormal
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If X > ScaleWidth - (m_HeaderHeight + 10) Or X < (m_HeaderHeight + 10) Then
        
        m_Right2Left = True
        DrawBody
        
    End If
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Debug.Print "Reading Properties "
    
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
    m_Animate = PropBag.ReadProperty("Animate", m_def_Animate)
    m_CalendarHeight = PropBag.ReadProperty("CalendarHeight", m_def_CalendarHeight)
    m_Curvature = PropBag.ReadProperty("Curvature", m_def_Curvature)
    m_Mode = PropBag.ReadProperty("Mode", m_def_Mode)
    
    m_CalendarGradient = PropBag.ReadProperty("CalendarGradient", m_Def_CalendarGradient)
    m_CalendarBackCol = PropBag.ReadProperty("CalendarBackCol", m_def_CalendarBackCol)
    m_MonthGradient = PropBag.ReadProperty("MonthGradient", m_def_MonthGradient)
    m_HeaderGradient = PropBag.ReadProperty("HeaderGradient", m_def_HeaderGradient)
    m_MonthBackCol = PropBag.ReadProperty("MonthBackCol", m_def_MonthBackCol)
    m_HeaderBackCol = PropBag.ReadProperty("HeaderBackCol", m_def_HeaderBackCol)
    m_WeekDayCol = PropBag.ReadProperty("WeekDayCol", m_def_WeekDayCol)
    m_DayCol = PropBag.ReadProperty("DayCol", m_def_DayCol)
    m_DaySelCol = PropBag.ReadProperty("DaySelCol", m_def_DaySelCol)
    m_WeekDaySelCol = PropBag.ReadProperty("WeekDaySelCol", m_def_WeekDaySelCol)
    m_DaySunCol = PropBag.ReadProperty("DaySunCol", m_def_DaySunCol)
    m_WeekDaySunCol = PropBag.ReadProperty("WeekDaySunCol", m_def_WeekDaySunCol)
    m_HeaderGradientCol = PropBag.ReadProperty("HeaderGradientCol", m_def_HeaderGradientCol)
    m_MonthGradientCol = PropBag.ReadProperty("MonthGradientCol", m_def_MonthGradientCol)
    m_CalendarGradientCol = PropBag.ReadProperty("CalendarGradientCol", m_def_CalendarGradientCol)
    m_YearBackCol = PropBag.ReadProperty("YearBackCol", m_def_YearBackCol)
    m_YearGradient = PropBag.ReadProperty("YearGradient", m_def_YearGradient)
    m_YearGradientCol = PropBag.ReadProperty("YearGradientCol", m_def_YearGradientCol)
    m_DateFormat = PropBag.ReadProperty("DateFormat", m_def_DateFormat)
    m_Sensitive = PropBag.ReadProperty("Sensitive", m_def_Sensitive)
    m_SkipEnabled = PropBag.ReadProperty("SkipEnabled", m_def_SkipEnabled)
    m_HeaderHeight = PropBag.ReadProperty("HeaderHeight", m_def_HeaderHeight)
    m_SpecialDays = PropBag.ReadProperty("SpecialDays", m_def_SpecialDays)

    UserControl.Appearance = m_Appearance
    UserControl.BorderStyle = m_BorderStyle
    UserControl.Enabled = m_Enabled
    Set m_Font = m_Font
    ImportCalendar
    If Not m_SpecialDays = vbNullString Then m_SpecialDayStack = Split(m_SpecialDays, ",")
    
End Sub

Private Sub UserControl_Resize()
On Error GoTo Handle
    
    ' set the height
    If Not m_Mode = [List Mode] Then
        Height = m_HeaderHeight * Screen.TwipsPerPixelX
    Else
        m_CalendarHeight = Height / Screen.TwipsPerPixelY - m_HeaderHeight
    End If
    picCalendar.Width = Width / Screen.TwipsPerPixelX
    picCalendar.Height = m_CalendarHeight
    
    RedrawControl
    
Handle:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Debug.Print "Writing Properties "
    
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
    Call PropBag.WriteProperty("Animate", m_Animate, m_def_Animate)
    Call PropBag.WriteProperty("CalendarHeight", m_CalendarHeight, m_def_CalendarHeight)
    Call PropBag.WriteProperty("Curvature", m_Curvature, m_def_Curvature)
    Call PropBag.WriteProperty("Mode", m_Mode, m_def_Mode)
    Call PropBag.WriteProperty("CalendarGradient", m_CalendarGradient, m_Def_CalendarGradient)
    Call PropBag.WriteProperty("CalendarBackCol", m_CalendarBackCol, m_def_CalendarBackCol)
    
    Call PropBag.WriteProperty("MonthGradient", m_MonthGradient, m_def_MonthGradient)
    Call PropBag.WriteProperty("HeaderGradient", m_HeaderGradient, m_def_HeaderGradient)
    Call PropBag.WriteProperty("MonthBackCol", m_MonthBackCol, m_def_MonthBackCol)
    Call PropBag.WriteProperty("HeaderBackCol", m_HeaderBackCol, m_def_HeaderBackCol)
    Call PropBag.WriteProperty("WeekDayCol", m_WeekDayCol, m_def_WeekDayCol)
    Call PropBag.WriteProperty("DayCol", m_DayCol, m_def_DayCol)
    Call PropBag.WriteProperty("DaySelCol", m_DaySelCol, m_def_DaySelCol)
    Call PropBag.WriteProperty("WeekDaySelCol", m_WeekDaySelCol, m_def_WeekDaySelCol)
    Call PropBag.WriteProperty("DaySunCol", m_DaySunCol, m_def_DaySunCol)
    Call PropBag.WriteProperty("WeekDaySunCol", m_WeekDaySunCol, m_def_WeekDaySunCol)
    Call PropBag.WriteProperty("MonthGradientCol", m_MonthGradientCol, m_def_MonthGradientCol)
    Call PropBag.WriteProperty("CalendarGradientCol", m_CalendarGradientCol, m_def_CalendarGradientCol)

    Call PropBag.WriteProperty("YearBackCol", m_YearBackCol, m_def_YearBackCol)
    Call PropBag.WriteProperty("YearGradient", m_YearGradient, m_def_YearGradient)
    Call PropBag.WriteProperty("YearGradientCol", m_YearGradientCol, m_def_YearGradientCol)
    Call PropBag.WriteProperty("DateFormat", m_DateFormat, m_def_DateFormat)
    Call PropBag.WriteProperty("Sensitive", m_Sensitive, m_def_Sensitive)
    Call PropBag.WriteProperty("SkipEnabled", m_SkipEnabled, m_def_SkipEnabled)
    Call PropBag.WriteProperty("HeaderHeight", m_HeaderHeight, m_def_HeaderHeight)
    Call PropBag.WriteProperty("SpecialDays", m_SpecialDays, m_def_SpecialDays)
End Sub

Private Sub RedrawControl()

    ' This is needed to redefine the Usercontrol region
    ' after we changed from the DatePickerMode
    If Not (m_Mode = [DatePicker PopDown] Or m_Mode = [DatePicker Popup]) Then
    Dim hrgn As Long
        hrgn = CreateRectRgn(0, 0, ScaleWidth + 2, ScaleHeight + 2)
        SetWindowRgn UserControl.hwnd, hrgn, True
    End If
    
    If m_Mode = [List Mode] Then ImportCalendar
    DrawCalendar
    DrawBody

End Sub

Private Sub DrawBody()
Dim Rct As RECT
Dim vStrDate As String
    
    Debug.Print "Drawing the body "
    
    If m_Mode = [DatePicker PopDown] Or m_Mode = [DatePicker Popup] Then GoTo DrawDatePicker

    ' Define the RECT
    Rct.Left = 0
    Rct.Right = ScaleWidth
    Rct.Top = (m_HeaderHeight - TextHeight("A")) / 2
    Rct.Bottom = m_HeaderHeight
    
    ' Get the Selectecd Day string
    vStrDate = Me.Caption(True)
    
    ' Resize to fit of needed
    If TextWidth("A") * Len(vStrDate) > ScaleWidth - m_WeekDaysHeight Then vStrDate = Me.Caption(False)
    If m_MonthPopupMode Then vStrDate = "Select Month  "
    
    ' Darw Gradients and Caption
    UserControl.Cls
    UserControl.FontBold = True
    UserControl.BackColor = m_HeaderBackCol
    If m_HeaderGradient Then PaintGradient UserControl.hdc, 0, 0, ScaleWidth, m_HeaderHeight, m_HeaderBackCol, m_HeaderGradientCol, Fill_Vertical, m_Right2Left
    DrawText UserControl.hdc, vStrDate, -1, Rct, 1
    
    ' Draw the arrow buttons
    FillStyle = vbSolid: FillColor = m_ArrowCol
    
    If m_Mode = [PopDown Mode] Then
        If m_Poped And Not m_Sensitive Then
            DrawArrow hdc, ScaleWidth - (m_HeaderHeight + 10) / 2, m_HeaderHeight * 0.6, m_HeaderHeight * 0.7, Arw_Up
        Else
            DrawArrow hdc, ScaleWidth - (m_HeaderHeight + 10) / 2, m_HeaderHeight * 0.2, m_HeaderHeight * 0.7, Arw_Down
        End If
    End If
    
    If m_Mode = [PopUp Mode] Then
        If m_Poped And Not m_Sensitive Then
            DrawArrow hdc, (m_HeaderHeight + 10) / 2, m_HeaderHeight * 0.2, m_HeaderHeight * 0.7, Arw_Down
        Else
            DrawArrow hdc, (m_HeaderHeight + 10) / 2, m_HeaderHeight * 0.6, m_HeaderHeight * 0.7, Arw_Up
        End If
    End If
    UserControl.Refresh
    
Exit Sub
DrawDatePicker:
    
    ' The DatePicker Mode is selected.
    ' We are only showing the Extreame right part of the Header.
    Dim hrgn As Long
    ' Draw The Gradient
    If m_HeaderGradient Then PaintGradient UserControl.hdc, 0, 0, m_HeaderHeight, m_HeaderHeight, m_HeaderBackCol, m_HeaderGradientCol, Fill_Vertical, m_Right2Left
    
    hrgn = CreateRectRgn(0, 0, m_HeaderHeight + 1, m_HeaderHeight)
    SetWindowRgn UserControl.hwnd, hrgn, True
    
    If m_Mode = [DatePicker PopDown] Then
        DrawArrow hdc, m_HeaderHeight / 2, m_HeaderHeight * 0.3, m_HeaderHeight * 0.5, Arw_Down
    Else
        DrawArrow hdc, m_HeaderHeight / 2, m_HeaderHeight * 0.5, m_HeaderHeight * 0.5, Arw_Up
    End If
    UserControl.Refresh
    
End Sub

Private Sub DrawCalendar()
Dim X As Long
Dim Y As Long
Dim Rct As RECT
Dim vMonth As String
Dim vStrDate As String
Dim tmpValue As Double
Dim vSelWeekDay As Long

    Debug.Print "Drawing the Calendar"
    picCalendar.Cls
    
    ' Load some the neccessary values
    m_WeekDaysHeight = m_CalendarHeight / 10

    m_TrackWidth = 0.1 * ScaleWidth
    m_iHeight = (picCalendar.ScaleHeight - m_WeekDaysHeight) / (m_RowDays + 1)
    m_iWidth = (picCalendar.ScaleWidth - m_TrackWidth - 1) / m_HeaderDays
    vSelWeekDay = Weekday(DateSerial(m_SelYear, m_SelMonth, m_SelDay))
    
    ' Fill the background Gradient
    picCalendar.BackColor = m_CalendarBackCol
    If m_CalendarGradient Then PaintGradient picCalendar.hdc, 0, 0, ScaleWidth, picCalendar.ScaleHeight, m_CalendarBackCol, m_CalendarGradientCol, Fill_UpwardDiagonal, True
    
    '-------------------------------
    '| Draw Month Selection Region |
    '-------------------------------
    picCalendar.FontBold = True
    picCalendar.FillColor = m_MonthBackCol
    
    ' draw the gradient
    RoundRect picCalendar.hdc, 0, 0, m_TrackWidth, picCalendar.ScaleHeight, 0, 0
    If m_MonthGradient Then PaintGradient picCalendar.hdc, 1, 1, m_TrackWidth - 2, picCalendar.ScaleHeight - 2, m_MonthBackCol, m_MonthGradientCol, Fill_Horizontal, False
    
    ' Get the Month Name
    vMonth = MonthName(m_SelMonth)
    tmpValue = picCalendar.TextHeight("A") * Len(vMonth)
    If tmpValue > picCalendar.ScaleHeight - m_TrackWidth Then vMonth = Left$(vMonth, 3): tmpValue = picCalendar.TextHeight("A") * 3
    Rct.Top = (picCalendar.ScaleHeight - tmpValue) / 2
    Rct.Bottom = picCalendar.ScaleHeight: Rct.Right = m_TrackWidth
    
    ' Sort Downward ( downward month print )
    For X = 1 To Len(vMonth)
        vStrDate = vStrDate & Mid$(vMonth, X, 1) & vbCrLf
    Next X
    
    ' Draw it
    DrawText picCalendar.hdc, UCase$(vStrDate), -1, Rct, 1
    
    ' Draw Month Selecting Arrows
    DrawArrow picCalendar.hdc, m_TrackWidth / 2, m_TrackWidth * 0.55, m_TrackWidth * 0.5, Arw_Up
    DrawArrow picCalendar.hdc, m_TrackWidth / 2, picCalendar.ScaleHeight - m_TrackWidth * 0.55, m_TrackWidth * 0.5, Arw_Down
    
    ' Draw the month popuping arrow
    DrawArrow picCalendar.hdc, m_TrackWidth * 0.8, picCalendar.ScaleHeight / 2, m_TrackWidth, Arw_Right, m_TrackWidth * 0.08
    
    '-----------------------------------
    '|Draw Horizontal Header Week Days |
    '-----------------------------------
    picCalendar.FontBold = False
    tmpValue = m_TrackWidth + 1
    Rct.Top = (m_WeekDaysHeight - TextHeight("A")) / 2: Rct.Bottom = m_WeekDaysHeight
    
    ' Through weekdays
    For X = 1 To m_HeaderDays
        
        ' Get weekday name
        vStrDate = Mid$(WeekdayName(X), 1, 3)
        Rct.Left = Int(tmpValue): Rct.Right = Int(tmpValue + m_iWidth - 1)
        If X = m_HeaderDays Then Rct.Right = picCalendar.ScaleWidth
        
        
        Select Case X
            Case vSelWeekDay ' Put selected weekday Col
                picCalendar.FillColor = m_WeekDaySelCol
                
            Case vbSunday ' Put sunday header Col
                picCalendar.FillColor = m_WeekDaySunCol
                

            Case Else   ' Put normal week-day Col
                picCalendar.FillColor = m_WeekDayCol
                
        End Select
        
        ' Draw the Week days name
        RoundRect picCalendar.hdc, tmpValue, 0, Rct.Right, m_WeekDaysHeight, 0, 0
        DrawText picCalendar.hdc, vStrDate, -1, Rct, 1
        tmpValue = tmpValue + m_iWidth
        
    Next X
    
    
    '------------------------------
    '|Draw Each Days in the Month |
    '------------------------------
    Dim vDayCell As Long
    X = 1: Y = 0
    
    ' Through days
    For vDayCell = 1 To 34 + m_FirstDay

        ' Some ordering
        If X > m_HeaderDays Then X = 1: Y = Y + 1
        If vDayCell = 36 Then X = 1: Y = 0

        ' Define rect
        Rct.Left = Int(m_TrackWidth + (X - 1) * m_iWidth + 1)
        Rct.Top = m_WeekDaysHeight + Y * m_iHeight + 1
        Rct.Bottom = Rct.Top + m_iHeight - 1
        Rct.Right = Int(Rct.Left + m_iWidth - 1)
        If X = m_HeaderDays Then Rct.Right = picCalendar.ScaleWidth
        
        ' Set day Cols
        If vDayCell - m_FirstDay + 1 = m_SelDay Then
            
            ' Put selected day Col
            picCalendar.FillColor = m_DaySelCol
        Else
            
            ' Check if sunday then put DaySunCol
            If X = 1 Then
                picCalendar.FillColor = m_DaySunCol
            Else    ' Normal day Col
                picCalendar.FillColor = m_DayCol
            End If
        End If
        
        ' Draw back
        RoundRect picCalendar.hdc, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom, m_Curvature, m_Curvature
        If IsSpecial(vDayCell - m_FirstDay + 1, m_SelMonth) = True Then picCalendar.FillColor = m_DaySunCol: RoundRect picCalendar.hdc, Rct.Left, Rct.Top, Rct.Right, Rct.Top + m_iHeight * 0.2, 0, 0


        ' Days are not of this month
        If Not (vDayCell - m_FirstDay + 1) <= m_MonthDays Then
            
            ' Fill cross
            picCalendar.FillStyle = vbDiagonalCross
            If X = 1 Then picCalendar.FillColor = m_DayCol Else picCalendar.FillColor = m_DaySunCol
            RoundRect picCalendar.hdc, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom, m_Curvature, m_Curvature
            picCalendar.FillStyle = vbSolid
            
        End If
        
        ' Days are of this month
        If Month(DateSerial(m_SelYear, m_SelMonth, vDayCell - m_FirstDay + 1)) = m_SelMonth Then
            
            ' Darw it
            Rct.Top = m_WeekDaysHeight + (Y) * m_iHeight + (m_iHeight - TextHeight("A")) / 2
            DrawText picCalendar.hdc, vDayCell - m_FirstDay + 1, -1, Rct, 1
        
        End If
        X = X + 1

    Next vDayCell

    '----------------------------------
    '| Draw The Year Selection Region |
    '----------------------------------
    Rct.Top = picCalendar.ScaleHeight - m_iHeight + 1
    Rct.Bottom = picCalendar.ScaleHeight
    Rct.Left = picCalendar.ScaleWidth - 3 * m_iWidth + 1
    picCalendar.FillColor = m_YearBackCol
    RoundRect picCalendar.hdc, Rct.Left, Rct.Top, picCalendar.ScaleWidth, Rct.Bottom, 0, 0
    If m_YearGradient Then PaintGradient picCalendar.hdc, Rct.Left + 1, Rct.Top + 1, 3 * m_iWidth - 3, (Rct.Bottom - Rct.Top) - 2, m_YearBackCol, m_YearGradientCol, Fill_Vertical, True

    ' Define Rect
    picCalendar.FontBold = True
    Rct.Left = m_TrackWidth + 1
    Rct.Right = m_TrackWidth + 4 * m_iWidth
    Rct.Top = picCalendar.ScaleHeight - m_iHeight + (m_iHeight - TextHeight("A")) / 2
    
    ' Draw Today
    DrawText picCalendar.hdc, "Today " & Format$(Date$, GetFormat), -1, Rct, 1
    
    ' Draw year
    Rct.Left = picCalendar.ScaleWidth - 3 * m_iWidth + 1
    Rct.Right = picCalendar.ScaleWidth
    DrawText picCalendar.hdc, Format$(m_SelYear, "0000"), -1, Rct, 1
    
    ' Draw year selecting arrows
    DrawArrow picCalendar.hdc, Rct.Right - m_iWidth / 2, Rct.Bottom - m_iHeight / 2, m_iHeight * 0.5, Arw_Right
    DrawArrow picCalendar.hdc, Rct.Left + m_iWidth / 2, Rct.Bottom - m_iHeight / 2, m_iHeight * 0.5, Arw_Left

End Sub

Private Sub ExportCalendar(ByVal vDown As Boolean)
Dim Rct1 As RECT
Dim Rct2 As RECT
Dim hWidth As Long
Dim ScrX As Long
Dim ScrY As Long
Dim fLeft As Long
Dim fTop As Long
Dim fOffset As Long

    Debug.Print "Exporting calendar "
    
    'Check Mode
    If m_Mode = [List Mode] Then ImportCalendar: Exit Sub
    m_Poped = True
    
    ' Get screen parameters
    ScrX = Screen.TwipsPerPixelX
    ScrY = Screen.TwipsPerPixelY
    
    With UserControl.Parent
    
        ' Some tricks are needed to locate the exact
        ' screen pos of the control
        
        ' Get the border width of tye parent form
        fLeft = (.Width - .ScaleWidth)
        fTop = (.Height - .ScaleHeight)
        
        ' There is a slight variation in Pos for different
        'borderstyles even if I used the GetWindowRect method
        'Anyone having a better idea please inform me
        Select Case .BorderStyle
            Case 0
                fOffset = 0
            Case 2, 5
                fOffset = 5
            Case Else
                fOffset = 3
        End Select
        
        ' Get parent rect\usercontrol Rect
        GetWindowRect .hwnd, Rct2
        GetWindowRect UserControl.hwnd, Rct1
        
        ' Set the new parent and Bring on top
        SetParent picCalendar.hwnd, .hwnd
        BringWindowToTop picCalendar.hwnd
        
        ' Calculate the form position
        hWidth = Rct1.Right - Rct1.Left
        Rct1.Left = Rct1.Left - Rct2.Left - fLeft / ScrX + fOffset
        Rct1.Top = Rct1.Top - Rct2.Top - fTop / ScrY + fOffset
        

        ' Place the claender
        If vDown Then
            ' Down
            picCalendar.Move (Rct1.Left) * ScrX, (Rct1.Top + m_HeaderHeight) * ScrY, hWidth * ScrX, m_CalendarHeight * ScrY
        Else
            ' Up
            picCalendar.Move (Rct1.Left) * ScrX, (Rct1.Top - m_CalendarHeight) * ScrY, hWidth * ScrX, m_CalendarHeight * ScrY
        End If
        
        picCalendar.Visible = True
        If m_Animate Then
            If m_Mode = [PopDown Mode] Then
                AnimateForm picCalendar, aload, eAppearFromRight, 10, 22
            Else
                AnimateForm picCalendar, aload, eAppearFromLeft, 10, 22
            End If
        Else
            ' This is necessary to redefine the Calendar region to full size
            AnimateForm picCalendar, aload, eAppearFromLeft, 0, 1
        End If
        
    End With

    
End Sub

Private Sub ImportCalendar()

    Debug.Print "Importing calendar "
    
    ' Set the new parent and Bring on top
    SetParent picCalendar.hwnd, UserControl.hwnd
    
    picCalendar.Move -1, m_HeaderHeight, ScaleWidth + 2, m_CalendarHeight
    
    picCalendar.Visible = True
    
End Sub

Private Sub DrawArrow(hdc As Long, _
                        ByVal X As Long, _
                        ByVal Y As Long, _
                        ByVal vSize As Long, _
                        ByVal vArrow As ArrowDir, _
                        Optional vThickness As Long = -1)
Dim Pnts(2) As POINTAPI
    
    ' Nothing here, define a point arry, fill it
    picCalendar.FillColor = m_ArrowCol
    If vThickness = -1 Then
        vThickness = vSize / 2
    Else
        ' Special case of popup month button
        picCalendar.FillColor = m_HeaderBackCol:
    End If
    
    ' Self explonatory
    If vArrow = Arw_Left Or vArrow = Arw_Right Then
        Pnts(0).X = X: Pnts(0).Y = Y - vSize / 2
        Pnts(1).X = X: Pnts(1).Y = Y + vSize / 2
        Pnts(2).Y = Y
        If vArrow = Arw_Left Then Pnts(2).X = X - vThickness Else Pnts(2).X = X + vThickness
    Else
        Pnts(0).X = X - vSize / 2: Pnts(0).Y = Y
        Pnts(1).X = X + vSize / 2: Pnts(1).Y = Y
        Pnts(2).X = X
        If vArrow = Arw_Down Then Pnts(2).Y = Y + vThickness Else Pnts(2).Y = Y - vThickness
    End If
    
    ' draw it
    Polygon hdc, Pnts(0), 3
    
End Sub


Private Function IsSpecial(ByVal vDay As Long, ByVal vMonth As Long) As Boolean
Dim X As Long
Dim xMax As Long
Dim vDayID As String
    
    On Error GoTo Handle
    
    ' This function is used to check whether the give day is
    ' loaded as a special day. The special days are already loaded to m_SpecialdaysStack
    If m_SpecialDays = vbNullString Or vDay = -1 Then IsSpecial = False: Exit Function
    vDayID = vDay & "-" & vMonth
    xMax = UBound(m_SpecialDayStack)
    
    ' Loop through all the days in specialday stack
    For X = 0 To xMax
        If m_SpecialDayStack(X) = vDayID Then IsSpecial = True: Exit Function
    Next X
    
Handle:
    IsSpecial = False
    
End Function


Private Function GetFormat() As String
Select Case m_DateFormat
    Case 0
        GetFormat = "dd-mm-yyyy"
    Case 1
        GetFormat = "mm-dd-yyyy"
    Case 2
        GetFormat = "yyyy-mm-dd"
End Select
End Function


Private Sub LoadDay(ByVal nDay As Long)
Dim dDate As String

    If m_SelMonth < 1 Then
        m_SelMonth = 12
        m_SelYear = m_SelYear - 1
    ElseIf m_SelMonth > 12 Then
        m_SelMonth = 1
        m_SelYear = m_SelYear + 1
    End If
    
    ' calculate Days in the month
    dDate = DateSerial(m_SelYear, m_SelMonth, 1)
    m_MonthDays = DateDiff("d", dDate, DateAdd("m", 1, dDate))
    m_FirstDay = Weekday(DateSerial(m_SelYear, m_SelMonth, 1))
    
    ' Special case Currently selected date is over the daycount
    If nDay = 999 Then
        If m_SelDay >= m_MonthDays Then m_SelDay = m_MonthDays: Exit Sub Else nDay = m_SelDay
    End If
    
    ' Splecial case moving last day up'
    ' Take July of 2005 and look 31 is on top
    If nDay < 0 Then
        If m_MonthDays > (36 - m_FirstDay) Then
            nDay = (35 - m_FirstDay) + (m_FirstDay + nDay)
        End If
    End If
    
    ' Cross filled cell was selected. Collapse if Sensitive else Skip to another Month
    If (nDay <= 0 Or nDay > m_MonthDays) Then
        If m_Sensitive And Not m_Mode = [List Mode] Then CollapseCalendar: Exit Sub
        
        If m_SkipEnabled Then
            If nDay <= 0 Then m_SelMonth = m_SelMonth - 1: LoadDay (999): Exit Sub
            If nDay > m_MonthDays Then m_SelMonth = m_SelMonth + 1: LoadDay (999): Exit Sub
        Else
            Exit Sub
        End If
    
    End If
    

    ' nO PROBLEM Load the day
    m_SelDay = nDay
    RaiseEvent DateChanged
End Sub


Private Sub PopupMonthList()
Dim X As Long
Dim Rct As RECT
Dim mName As String
Dim Y As Long
Dim vText As String
    
    ' Enable Popup mode\Define rect
    m_MonthPopupMode = True
    m_MonthPopWidth = (picCalendar.ScaleWidth - m_TrackWidth) / 12
    Rct.Bottom = picCalendar.ScaleHeight: Rct.Top = 0
    picCalendar.FontBold = False
    picCalendar.FillStyle = vbSolid
    picCalendar.FillColor = m_HeaderBackCol
    
    ' through months
    For X = 1 To 12
    
        ' Get month name
        vText = vbNullString
        mName = UCase$(MonthName(X))
        
        ' Sort it Downward
        For Y = 1 To Len(mName)
            vText = vText & Mid$(mName, Y, 1) & vbCrLf
        Next Y
        
        ' Draw it
        Rct.Left = m_TrackWidth + Int((X - 1) * m_MonthPopWidth) - 1
        Rct.Right = Int(Rct.Left + m_MonthPopWidth) + 2
        RoundRect picCalendar.hdc, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom, 0, 0
        DrawText picCalendar.hdc, vText, -1, Rct, 1
        
    Next X
    picCalendar.Refresh
    
End Sub


Private Sub PaintGradient(ByVal hdc As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As GradientDirectionCts, _
                         Optional Right2Left As Boolean = True)

  Dim uBIH    As BITMAPINFOHEADER
  Dim lBits() As Long
  Dim lGrad() As Long
  
  Dim R1      As Long
  Dim G1      As Long
  Dim B1      As Long
  Dim R2      As Long
  Dim G2      As Long
  Dim B2      As Long
  Dim dR      As Long
  Dim dG      As Long
  Dim dB      As Long
  
  Dim Scan    As Long
  Dim i       As Long
  Dim iEnd    As Long
  Dim iOffset As Long
  Dim j       As Long
  Dim jEnd    As Long
  Dim iGrad   As Long
  Dim tmpCol  As Long
  
  
    '-- A minor check
    If GradientDirection = Fill_None Then Exit Sub
    If (Width < 1 Or Height < 1) Then Exit Sub
    
    If Right2Left Then
        tmpCol = Col1
        Col1 = Col2
        Col2 = tmpCol
    End If
    
    '-- Decompose Cols
    Col1 = Col1 And &HFFFFFF
    R1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    G1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    B1 = Col1 Mod &H100&
    Col2 = Col2 And &HFFFFFF
    R2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    G2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    B2 = Col2 Mod &H100&
    
    '-- Get Col distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1
    
    '-- Size gradient-Cols array
    Select Case GradientDirection
        Case [Fill_Horizontal]
            ReDim lGrad(0 To Width - 1)
        Case [Fill_Vertical]
            ReDim lGrad(0 To Height - 1)
        Case Else
            ReDim lGrad(0 To Width + Height - 2)
    End Select
    
    '-- Calculate gradient-Cols
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If
    
    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width
    
    '-- Render gradient DIB
    Select Case GradientDirection
        
        Case [Fill_Horizontal]
        
            For j = 0 To jEnd
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(i - iOffset)
                Next i
                iOffset = iOffset + Scan
            Next j
        
        Case [Fill_Vertical]
        
            For j = jEnd To 0 Step -1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(j)
                Next i
                iOffset = iOffset + Scan
            Next j
            
        Case [Fill_DownwardDiagonal]
            
            iOffset = jEnd * Scan
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset - Scan
                iGrad = j
            Next j
            
        Case [Fill_UpwardDiagonal]
            
            iOffset = 0
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset + Scan
                iGrad = j
            Next j
    End Select
    
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With
    
    '-- Paint it!
    Call StretchDIBits(hdc, X, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_ColS, vbSrcCopy)
End Sub


Private Function AnimateForm(hwndObject As Object, ByVal aEvent As AnimeEventEnum, _
                            Optional ByVal aEffect As AnimeEffectEnum = 11, _
                            Optional ByVal FrameTime As Long = 1, _
                            Optional ByVal FrameCount As Long = 33) As Boolean
On Error GoTo Handle
Dim X1 As Long, Y1 As Long
Dim hrgn As Long, tmpRgn As Long
Dim XValue As Long, YValue As Long
Dim XIncr As Double, YIncr As Double
Dim wHeight As Long, wWidth As Long

    wWidth = hwndObject.Width / Screen.TwipsPerPixelX
    wHeight = hwndObject.Height / Screen.TwipsPerPixelY
'    hwndObject.Visible = True
    
    Select Case aEffect
    
        Case eAppearFromLeft
        
            XIncr = wWidth / FrameCount
            For X1 = 0 To FrameCount
            
                ' Define the size of current frame/Create it
                XValue = X1 * XIncr
                hrgn = CreateRectRgn(0, 0, XValue, wHeight)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hrgn, hrgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hrgn, True: DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eAppearFromRight
        
            XIncr = wWidth / FrameCount
            For X1 = 0 To FrameCount
                
                ' Define the size of current frame/Create it
                XValue = wWidth - X1 * XIncr
                hrgn = CreateRectRgn(XValue, 0, wWidth, wHeight)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hrgn, hrgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hrgn, True:  DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eAppearFromTop
        
            YIncr = wHeight / FrameCount
            For Y1 = 0 To FrameCount
            
                ' Define the size of current frame/Create it
                YValue = Y1 * YIncr
                hrgn = CreateRectRgn(0, 0, wWidth, YValue)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hrgn, hrgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hrgn, True:   DoEvents
                Sleep FrameTime
                
            Next Y1
            
        Case eAppearFromBottom
        
            YIncr = wHeight / FrameCount
            For Y1 = 0 To FrameCount
            
                ' Define the size of current frame/Create it
                YValue = wHeight - Y1 * YIncr
                hrgn = CreateRectRgn(0, YValue, wWidth, wHeight)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hrgn, hrgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hrgn, True: DoEvents
                Sleep FrameTime
                
            Next Y1
    End Select

    AnimateForm = True
    
Exit Function
Handle:
    AnimateForm = False
End Function

