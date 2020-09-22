VERSION 5.00
Begin VB.Form Statistics 
   Caption         =   "Statistical Analysis"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9480
   Icon            =   "Statistics.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Statistics"
   ScaleHeight     =   4905
   ScaleWidth      =   9480
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Statistics.frx":0442
      Left            =   1605
      List            =   "Statistics.frx":0464
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   -15
      Width           =   615
   End
   Begin VB.CheckBox Option2 
      Caption         =   "Show all decimals"
      Height          =   225
      Left            =   2355
      TabIndex        =   48
      Top             =   45
      Width           =   1635
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   11
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3555
      Width           =   1710
   End
   Begin VB.CommandButton Button 
      Caption         =   "&SORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   18
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   3615
      Width           =   2340
   End
   Begin VB.CommandButton Button 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   2535
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1965
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1410
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   840
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   14
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   3105
      Width           =   2340
   End
   Begin VB.CommandButton Button 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   17
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   4125
      Width           =   2340
   End
   Begin VB.PictureBox Display 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000015&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   6135
      ScaleHeight     =   375
      ScaleWidth      =   3300
      TabIndex        =   0
      Top             =   270
      Width           =   3330
   End
   Begin VB.Frame Frame1 
      Caption         =   "Method"
      Height          =   945
      Index           =   0
      Left            =   4260
      TabIndex        =   36
      Top             =   3795
      Width           =   1740
      Begin VB.OptionButton Option1 
         Caption         =   "Sample"
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   38
         Top             =   225
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Population"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   37
         Top             =   555
         Width           =   1185
      End
   End
   Begin VB.CommandButton Button 
      Caption         =   "C&P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   16
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "48"
      ToolTipText     =   "Clear Previous Entry"
      Top             =   2535
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   15
      Left            =   7875
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   2535
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "C&D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   13
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "48"
      ToolTipText     =   "Clear Digit"
      Top             =   1965
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "C&E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   12
      Left            =   8475
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "48"
      ToolTipText     =   "Cear Entry"
      Top             =   1410
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "&C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   11
      Left            =   8475
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "48"
      ToolTipText     =   "Clear All Data"
      Top             =   840
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   10
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   2535
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   9
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   840
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   8
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   840
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1410
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1410
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1965
      Width           =   435
   End
   Begin VB.CommandButton Button 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1965
      Width           =   435
   End
   Begin VB.ListBox ArrayList 
      Height          =   3570
      ItemData        =   "Statistics.frx":0487
      Left            =   4245
      List            =   "Statistics.frx":0489
      TabIndex        =   2
      Top             =   180
      Width           =   1785
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   10
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1755
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   9
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   315
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   0
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   675
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   1
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   8
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3195
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   7
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4290
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   6
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3900
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   5
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   4
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2460
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   3
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2115
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   2
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1395
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "Decimal places"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   420
      TabIndex        =   50
      Top             =   30
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "The skew is:"
      Height          =   195
      Index           =   12
      Left            =   90
      TabIndex        =   46
      Top             =   3255
      Width           =   2025
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   11
      Left            =   9030
      Picture         =   "Statistics.frx":048B
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   0
      Left            =   6285
      Picture         =   "Statistics.frx":0AA9
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   1
      Left            =   6540
      Picture         =   "Statistics.frx":0F63
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   2
      Left            =   6765
      Picture         =   "Statistics.frx":13C5
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   3
      Left            =   7020
      Picture         =   "Statistics.frx":187F
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   4
      Left            =   7275
      Picture         =   "Statistics.frx":1D39
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   5
      Left            =   7515
      Picture         =   "Statistics.frx":21F3
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   6
      Left            =   7770
      Picture         =   "Statistics.frx":26AD
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   7
      Left            =   8025
      Picture         =   "Statistics.frx":2B67
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   8
      Left            =   8280
      Picture         =   "Statistics.frx":3021
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   9
      Left            =   8535
      Picture         =   "Statistics.frx":34DB
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   10
      Left            =   8790
      Picture         =   "Statistics.frx":3995
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "The mode is:"
      Height          =   255
      Index           =   10
      Left            =   90
      TabIndex        =   22
      Top             =   1785
      Width           =   1920
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   75
      X2              =   3945
      Y1              =   2055
      Y2              =   2055
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   75
      X2              =   3975
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label1 
      Caption         =   "The number of entries is:"
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   18
      Top             =   390
      Width           =   2025
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   75
      X2              =   3960
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   75
      X2              =   3960
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   75
      X2              =   3945
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   75
      X2              =   3930
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   75
      X2              =   3930
      Y1              =   2775
      Y2              =   2775
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   75
      X2              =   3975
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   75
      X2              =   3945
      Y1              =   3495
      Y2              =   3495
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   75
      X2              =   3930
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   75
      X2              =   3945
      Y1              =   4230
      Y2              =   4230
   End
   Begin VB.Label Label1 
      Caption         =   "The average deviation is:"
      Height          =   255
      Index           =   9
      Left            =   105
      TabIndex        =   10
      Top             =   3600
      Width           =   2130
   End
   Begin VB.Label Label1 
      Caption         =   "The coefficient of variation is:"
      Height          =   255
      Index           =   8
      Left            =   90
      TabIndex        =   9
      Top             =   4335
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "The standard deviation is:"
      Height          =   255
      Index           =   7
      Left            =   90
      TabIndex        =   8
      Top             =   3975
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "The variance is:"
      Height          =   255
      Index           =   6
      Left            =   90
      TabIndex        =   7
      Top             =   2895
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "The mean is:"
      Height          =   255
      Index           =   4
      Left            =   90
      TabIndex        =   6
      Top             =   2535
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   "The median is:"
      Height          =   255
      Index           =   5
      Left            =   90
      TabIndex        =   5
      Top             =   2175
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   "The sum of squares is:"
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   4
      Top             =   1455
      Width           =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "The sum squared is:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   1095
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "The sum of the numbers is:"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   750
      Width           =   1965
   End
End
Attribute VB_Name = "Statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheArray() As Double
Dim SortedArray() As Double
Dim DigitArray() As Integer
Dim Counter As Double
Dim Digits As Integer
Dim Sample As Boolean
Dim Calculator As Boolean
Dim DecimalPlaced As Boolean
Dim Sorted As Boolean
Dim Rounding As Boolean
Dim Places As Integer
Private Sub AddToArray()
Dim i As Integer
Dim Temp As String
If Digits = 0 Then Exit Sub
For i = 1 To Digits
Select Case DigitArray(i)
Case 0 To 9
Temp = Temp & CStr(DigitArray(i))
Case 10
Temp = Temp & "."
Case 11
Temp = "-"
Case Else
Exit Sub
End Select
Next
Counter = Counter + 1
ReDim Preserve TheArray(Counter)
TheArray(Counter) = Val(Temp)
ShowResults
 If Sorted Then
 SortListbox
 Else
 ListUnsorted
 End If
End Sub

Private Function AveDev()
Dim i As Long
Dim Temp As Double
For i = 1 To Counter
Temp = Temp + Abs((TheArray(i) - Mean))
Next
AveDev = Temp / Counter
End Function


Private Function CoefDev() As Variant
On Error GoTo Err1
'coefficient of deviation in percent form
CoefDev = (StdDev / Mean * 100)
Exit Function
Err1:
CoefDev = "#Div/0!"
End Function

Private Function Mode() As Variant
'the most repeated value
On Error GoTo Err1:
Dim i As Long
Dim j As Long
Dim Temp As Variant
Dim Element As Long
ReDim n(Counter) As Long
'load number of repetitions into an array n()
For i = 1 To Counter
For j = 1 To Counter
If TheArray(i) = TheArray(j) Then
n(i) = n(i) + 1
End If
Next
Next
'compare elements of the repetition counting array
j = n(1)
Element = 1
For i = 2 To Counter
If n(i) > j Then
Element = i 'this element has higher value
j = n(i) 'update for next rep
End If
Next
'get results
If Element = 1 And n(1) = 1 Then 'no repetitions
Mode = "None"
Exit Function
End If
'look for the highest tying values
For i = 1 To Counter
If Element <> i Then 'skip same one
If n(Element) = n(i) Then 'if it is a match
If TheArray(Element) <> TheArray(i) Then 'but not same value
If InStr(1, Temp, TheArray(i)) = 0 Then 'if not already listed
If Temp = "" Then 'put in the first one
Temp = TheArray(Element)
End If
Temp = Temp & " or " & TheArray(i) 'add the matching reps
End If
End If
End If
End If
Next

'if no ties found show the highest repeated value
If Temp = "" Then
Temp = TheArray(Element)
End If

Mode = Temp
Exit Function
Err1:
Mode = "Error"
End Function

Private Sub ShowResults()
Dim Temp As String
Dim Number As Double
SortArray
txtOutput(0) = Str(Total)
txtOutput(1) = Str(Total * Total)
txtOutput(2) = Str(SqrTotal)
txtOutput(3) = Str(Median)

Temp = Str(Mean)
Temp = Format(Temp, "####.0000000000")
 If Not Rounding Then
 Number = CDbl(Temp)
 Else
 Number = Round(CDbl(Temp), Places)
 End If
txtOutput(4) = Str(Number)

Temp = CStr(Variance)
Temp = Format(Temp, "####.0000000000")
If IsNumeric(Temp) Then
 If Not Rounding Then
 Number = CDbl(Temp)
 Else
 Number = Round(CDbl(Temp), Places)
 End If
txtOutput(5) = Str(Number)
Else
txtOutput(5) = Temp
End If

Temp = CStr(StdDev) 'call function
Temp = Format(Temp, "####.0000000000") 'remove any scientific notation
If IsNumeric(Temp) Then 'see if it's a number or error message
 If Not Rounding Then
 Number = CDbl(Temp) ' convert back to double to drop trailing zeros
 Else
 Number = Round(CDbl(Temp), Places)
 End If
txtOutput(6) = Str(Number) 'display as string
Else
txtOutput(6) = Temp 'display error message
End If

Temp = CStr(CoefDev)
Temp = Format(Temp, "####.0000000000")
If IsNumeric(Temp) Then
 If Not Rounding Then
 Number = CDbl(Temp)
 Else
 Number = Round(CDbl(Temp), Places)
 End If
txtOutput(7) = Str(Number) & " %"
Else
txtOutput(7) = Temp
End If

Temp = CStr(Skew)
Temp = Format(Temp, "####.0000000000")
If IsNumeric(Temp) Then
 If Not Rounding Then
 Number = CDbl(Temp)
 Else
 Number = Round(CDbl(Temp), Places)
 End If
txtOutput(8) = Str(Number)
Else
txtOutput(8) = Temp
End If

txtOutput(9) = Str(Counter)
txtOutput(10) = CStr(Mode)

Temp = Str(AveDev)
Temp = Format(Temp, "####.0000000000")
If IsNumeric(Temp) Then
If Not Rounding Then
Number = CDbl(Temp)
Else
Number = Round(CDbl(Temp), Places)
End If
txtOutput(11) = Str(Number)
Else
txtOutput(11) = Temp
End If

End Sub
Private Function Mean() As Double
'average of all values
Mean = Total / Counter
End Function


Private Function Median() As Double
'the middle value of array or average of two if even number
Select Case Counter Mod 2
Case 0
Median = Str((SortedArray(Counter / 2) + SortedArray((Counter / 2) + 1)) / 2)
Case 1
Median = Str(SortedArray((Counter + 1) / 2))
End Select
End Function


Private Function Skew() As Variant
On Error GoTo Err1:
Dim i As Long
Dim Temp As Double
For i = 1 To Counter
Temp = Temp + ((TheArray(i) - Mean) / StdDev) ^ 3
Next
Skew = (Counter / ((Counter - 1) * (Counter - 2))) * Temp
Exit Function
Err1:
Skew = "#Div/0!"
End Function

Private Sub SortArray() 'only to get median value
Dim Temp As Double
Dim j As Long
Dim i As Long
'first copy the array so we can still remove the
'last element of original array in the undo feature
ReDim SortedArray(Counter)
For i = 1 To Counter
SortedArray(i) = TheArray(i)
Next
'then loop through swapping values through temp
For i = 1 To Counter
For j = 1 To Counter
Temp = SortedArray(i)
If Temp < SortedArray(j) Then
SortedArray(i) = SortedArray(j)
SortedArray(j) = Temp
End If
Next
Next
End Sub



Private Sub ListUnsorted()
Dim i As Integer
ArrayList.Clear
For i = 1 To Counter
ArrayList.AddItem Str(TheArray(i))
Next
Button(18).Caption = "&SORT"
Sorted = False
End Sub
Private Sub SortListbox()
Dim i As Integer
ArrayList.Clear
For i = 1 To Counter
ArrayList.AddItem Str(SortedArray(i))
Next
Button(18).Caption = "UN&SORT"
Sorted = True
End Sub

Private Function SqrTotal() As Double
Dim i As Long
For i = 1 To Counter
SqrTotal = (TheArray(i) * TheArray(i)) + SqrTotal
Next
End Function
Private Function StdDev() As Variant
On Error GoTo Err1: 'standard deviation
StdDev = Sqr(Abs(Variance))
Exit Function
Err1:
StdDev = "#Div/0!"
End Function

Private Function Total() As Double
Dim i As Long 'total of values
For i = 1 To Counter
Total = TheArray(i) + Total
Next
End Function


Private Function Variance() As Variant
On Error GoTo Err1
Dim Sum As Double
Dim i As Long
For i = 1 To Counter 'summation of squares for think formula
Sum = Sum + ((TheArray(i) - Mean) * (TheArray(i) - Mean))
Next
If Sample Then 'sample method
   If Calculator Then 'for using hand calculator
   Variance = ((SqrTotal - ((Total * Total) / Counter))) / (Counter - 1)
   Else 'using "think" formula...better
   Variance = Sum / (Counter - 1)
   End If
Else 'population method
   If Calculator Then
   Variance = ((SqrTotal - ((Total * Total) / Counter))) / Counter
   Else
   Variance = Sum / Counter
   End If
End If
Exit Function
Err1:
Variance = "#Div/0!"
End Function

Private Sub ArrayList_Click()
SortListbox
End Sub

Private Sub Button_Click(Index As Integer)
Dim i As Integer
Select Case Index
Case 0 To 9 'numbers 0 through 9
Form_KeyPress (Index + 48)

Case 10 'decimal point
Form_KeyPress 46

Case 11 'clear all data
Digits = 0
Counter = 0
Erase DigitArray
Erase TheArray
Display.Cls
DecimalPlaced = False
For i = 0 To 11
txtOutput(i) = ""
Next
ArrayList.Clear

Case 12 'clear entry
Digits = 0
Erase DigitArray
Display.Cls
DecimalPlaced = False

Case 13 'backspace key
Form_KeyPress 8

Case 14 'enter button
Form_KeyPress 13

Case 15 'minus sign
Form_KeyPress 45

Case 16 'clear previous value (and current entry)
If Counter > 0 Then
Digits = 0
Erase DigitArray
Display.Cls
Counter = Counter - 1
DecimalPlaced = False
 If Counter > 0 Then
 ShowResults
 Else
 Button_Click 11 'clear all data
 End If
 
 If Sorted Then 'refresh listbox
 SortListbox
 Else
 ListUnsorted
 End If
End If

Case 17
End

Case 18 'show list box (un)sorted
If Sorted Then 'refresh listbox
ListUnsorted
Else
SortListbox
End If

End Select
SetFocus
End Sub

Private Sub Combo1_Click()
Places = CInt(Combo1.Text)
If Counter > 0 Then
ShowResults
End If
End Sub


Private Sub Digit_Click(Index As Integer)
'image controls containing number bitmaps
End Sub

Private Sub Display_Click()
' calculator picture box
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Dim TheDigit As Integer
Dim i As Integer
Select Case KeyAscii
Case 8 'backspace
If Digits > 0 Then
If DigitArray(Digits) = 10 Then
DecimalPlaced = False
End If
Digits = Digits - 1
End If

Case 13 'enter key
AddToArray
Digits = 0
Erase DigitArray
DecimalPlaced = False

Case 45 'minus sign
If Digits > 0 Then Exit Sub
Digits = Digits + 1
ReDim Preserve DigitArray(Digits) As Integer
DigitArray(Digits) = 11

Case 46 'decimal point
If DecimalPlaced Then Exit Sub
Digits = Digits + 1
ReDim Preserve DigitArray(Digits) As Integer
DigitArray(Digits) = 10
DecimalPlaced = True

Case 48 To 57 'numbers 0 through 9
Digits = Digits + 1
ReDim Preserve DigitArray(Digits) As Integer
DigitArray(Digits) = CInt(Chr(KeyAscii))

Case 27
Button_Click 12

Case 101
Button_Click 12

Case 99
Button_Click 11

Case 100
Button_Click 13

Case 112
Button_Click 16

Case 115
Button_Click 18

Case 120
End

Case Else

Exit Sub
End Select

Display.Cls

For i = 1 To Digits
TheDigit = DigitArray(Digits - i + 1)
Display.PaintPicture Digit(TheDigit).Picture, (Display.Width - 100) - Digit(TheDigit).Width * i, 0
Next

Display.Refresh
End Sub

Private Sub Form_Load()
Sample = True
Rounding = True
Places = 3
Combo1.Text = "3"
End Sub


Private Sub Frame1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
' the two frames to separate option button pairs
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
Sample = False
Case 1
Sample = True
End Select
If Counter > 0 Then ShowResults
End Sub


Private Sub Option2_Click()
If Option2 Then
Rounding = False
Combo1.Enabled = False
Combo1.Text = 10
Else
Rounding = True
Combo1.Enabled = True
Combo1.Text = 3
End If
If Counter > 0 Then ShowResults
End Sub


