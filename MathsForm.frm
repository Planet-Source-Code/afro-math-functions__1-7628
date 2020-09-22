VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form MathsForm 
   Caption         =   "Maths Functions"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   3000
      Picture         =   "MathsForm.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   62
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   3480
      Picture         =   "MathsForm.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   61
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   3960
      Picture         =   "MathsForm.frx":074C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   60
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   4440
      Picture         =   "MathsForm.frx":0B8E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   59
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   4920
      Picture         =   "MathsForm.frx":0FD0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   58
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5520
      Top             =   6720
   End
   Begin TabDlg.SSTab Option 
      Height          =   6180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   10901
      _Version        =   327680
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Calculating equations"
      TabPicture(0)   =   "MathsForm.frx":12DA
      Tab(0).ControlCount=   14
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label19"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label20"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command7"
      Tab(0).Control(6).Enabled=   -1  'True
      Tab(0).Control(7)=   "Command8"
      Tab(0).Control(7).Enabled=   -1  'True
      Tab(0).Control(8)=   "Command9"
      Tab(0).Control(8).Enabled=   -1  'True
      Tab(0).Control(9)=   "Command10"
      Tab(0).Control(9).Enabled=   -1  'True
      Tab(0).Control(10)=   "Round"
      Tab(0).Control(10).Enabled=   -1  'True
      Tab(0).Control(11)=   "Answer"
      Tab(0).Control(11).Enabled=   -1  'True
      Tab(0).Control(12)=   "Command11"
      Tab(0).Control(12).Enabled=   -1  'True
      Tab(0).Control(13)=   "Calc"
      Tab(0).Control(13).Enabled=   -1  'True
      TabCaption(1)   =   "Area"
      TabPicture(1)   =   "MathsForm.frx":12F6
      Tab(1).ControlCount=   4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(3).Enabled=   0   'False
      TabCaption(2)   =   "Volume"
      TabPicture(2)   =   "MathsForm.frx":1312
      Tab(2).ControlCount=   3
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame7"
      Tab(2).Control(2).Enabled=   0   'False
      TabCaption(3)   =   "Converter"
      TabPicture(3)   =   "MathsForm.frx":132E
      Tab(3).ControlCount=   2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame8"
      Tab(3).Control(1).Enabled=   0   'False
      TabCaption(4)   =   "Fuel Expenses"
      TabPicture(4)   =   "MathsForm.frx":134A
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame10"
      Tab(4).Control(0).Enabled=   0   'False
      TabCaption(5)   =   "Misc."
      TabPicture(5)   =   "MathsForm.frx":1366
      Tab(5).ControlCount=   4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame11"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame12"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame13"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Frame14"
      Tab(5).Control(3).Enabled=   0   'False
      Begin VB.Frame Frame14 
         Caption         =   "Pythagorus Thereom"
         Height          =   2295
         Left            =   -70440
         TabIndex        =   181
         Top             =   3600
         Width           =   3735
         Begin VB.TextBox a 
            Height          =   375
            Left            =   600
            TabIndex        =   187
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox b 
            Height          =   375
            Left            =   600
            TabIndex        =   186
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox c 
            Height          =   375
            Left            =   600
            TabIndex        =   185
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Solve"
            Height          =   375
            Left            =   600
            TabIndex        =   184
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox PythagorumTheorum 
            Height          =   375
            Left            =   2400
            TabIndex        =   183
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Quit"
            Height          =   375
            Left            =   600
            TabIndex        =   182
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label66 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a ="
            BeginProperty Font 
               Name            =   "Notehand"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   191
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label65 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "b ="
            BeginProperty Font 
               Name            =   "Notehand"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   190
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label64 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "c ="
            BeginProperty Font 
               Name            =   "Notehand"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   189
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label63 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "c ="
            BeginProperty Font 
               Name            =   "Notehand"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   188
            Top             =   1800
            Width           =   375
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Calculating Averages"
         Height          =   2175
         Left            =   -74760
         TabIndex        =   174
         Top             =   3720
         Width           =   3495
         Begin VB.CommandButton Command18 
            Caption         =   "Average"
            Height          =   375
            Left            =   1080
            TabIndex        =   178
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Add Number"
            Height          =   375
            Left            =   1080
            TabIndex        =   177
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Average 
            Height          =   375
            Left            =   1800
            TabIndex        =   176
            Text            =   "0"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Number 
            Height          =   375
            Left            =   480
            TabIndex        =   175
            Text            =   "0"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Sum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   600
            TabIndex        =   180
            Top             =   1680
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label Total 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   360
            TabIndex        =   179
            Top             =   1680
            Visible         =   0   'False
            Width           =   90
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Age Calculator"
         Height          =   2895
         Left            =   -70440
         TabIndex        =   168
         Top             =   600
         Width           =   3735
         Begin VB.CommandButton Command19 
            Caption         =   "Calculate my age"
            Height          =   375
            Left            =   840
            TabIndex        =   171
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1560
            TabIndex        =   170
            Text            =   "12/04/1984"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Help"
            Height          =   375
            Left            =   2280
            TabIndex        =   169
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Age 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Needlepoint"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1440
            TabIndex        =   173
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label62 
            Alignment       =   2  'Center
            Caption         =   "Enter your Date Of Birth in the box below and I will calculate your exact age"
            BeginProperty Font 
               Name            =   "Notehand"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   915
            Left            =   720
            TabIndex        =   172
            Top             =   240
            Width           =   2595
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Calculating Sales Tax"
         Height          =   2895
         Left            =   -74760
         TabIndex        =   152
         Top             =   600
         Width           =   3495
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   360
            TabIndex        =   160
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   360
            TabIndex        =   159
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   1920
            TabIndex        =   158
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Calculate Tax Paid"
            Height          =   255
            Left            =   240
            TabIndex        =   157
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Calculate Tax to pay"
            Height          =   255
            Left            =   240
            TabIndex        =   156
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox Tax 
            Height          =   375
            Left            =   120
            TabIndex        =   155
            Text            =   "17.5"
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   1920
            TabIndex        =   154
            Top             =   1320
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1560
            TabIndex        =   153
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Money Spent"
            Height          =   195
            Left            =   1800
            TabIndex        =   167
            Top             =   1080
            Width           =   945
         End
         Begin VB.Label Label60 
            Caption         =   "Tax Rate"
            Height          =   255
            Left            =   240
            TabIndex        =   166
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label59 
            Caption         =   "%"
            Height          =   255
            Left            =   960
            TabIndex        =   165
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label Label58 
            Caption         =   "Tax Paid"
            Height          =   255
            Left            =   1680
            TabIndex        =   164
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label57 
            Caption         =   "£"
            Height          =   255
            Left            =   1440
            TabIndex        =   163
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label Label56 
            Caption         =   "£"
            Height          =   255
            Left            =   240
            TabIndex        =   162
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label53 
            Caption         =   "£"
            Height          =   255
            Left            =   240
            TabIndex        =   161
            Top             =   1440
            Width           =   255
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Calculating fuel expenses"
         Height          =   5295
         Left            =   -73920
         TabIndex        =   132
         Top             =   900
         Width           =   5895
         Begin VB.CommandButton Command1 
            Caption         =   "Help"
            Height          =   375
            Left            =   3840
            TabIndex        =   149
            Top             =   4440
            Width           =   1095
         End
         Begin VB.TextBox txtMileage 
            Height          =   375
            Left            =   1920
            TabIndex        =   137
            Text            =   "100"
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtMpG 
            Height          =   375
            Left            =   1920
            TabIndex        =   136
            Text            =   "12.4"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtModel 
            Height          =   375
            Left            =   1920
            TabIndex        =   135
            Text            =   "Chrysler Viper 8.0 GTS"
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox txtCost 
            Height          =   375
            Left            =   1920
            TabIndex        =   134
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox txtLitreCost 
            Height          =   375
            Left            =   1920
            TabIndex        =   133
            Text            =   "45"
            Top             =   2280
            Width           =   975
         End
         Begin Threed.SSPanel pnlMessage 
            Height          =   375
            Left            =   120
            TabIndex        =   138
            Top             =   4800
            Width           =   7770
            _Version        =   65536
            _ExtentX        =   13705
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "ffvfv"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Alignment       =   1
         End
         Begin Threed.SSCommand cmdCalc 
            Default         =   -1  'True
            Height          =   375
            Left            =   120
            TabIndex        =   150
            Top             =   240
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "+"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            Font3D          =   1
            RoundedCorners  =   0   'False
            AutoSize        =   1
            MousePointer    =   99
            MouseIcon       =   "MathsForm.frx":1382
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "@"
            Height          =   195
            Left            =   2640
            TabIndex        =   151
            Top             =   3360
            Width           =   165
         End
         Begin VB.Label lblMileage 
            Caption         =   "Mileage"
            Height          =   495
            Left            =   240
            TabIndex        =   148
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblMpG 
            Caption         =   "Miles per Gallon (urban)"
            Height          =   375
            Left            =   240
            TabIndex        =   147
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label lblPPL 
            Caption         =   "Price per Litre"
            Height          =   375
            Left            =   240
            TabIndex        =   146
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label lblCostLabel 
            Caption         =   "Cost"
            Height          =   375
            Left            =   240
            TabIndex        =   145
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lblVATLabel 
            Caption         =   "Cost - VAT "
            Height          =   375
            Left            =   240
            TabIndex        =   144
            Top             =   3960
            Width           =   1215
         End
         Begin VB.Label lblVAT 
            Caption         =   "VAT"
            Height          =   375
            Left            =   1800
            TabIndex        =   143
            Top             =   3960
            Width           =   750
         End
         Begin VB.Label lblVATRate 
            AutoSize        =   -1  'True
            Caption         =   "17.5%"
            Height          =   195
            Left            =   2880
            TabIndex        =   142
            Top             =   3360
            Width           =   435
         End
         Begin VB.Label Label54 
            Caption         =   "Model"
            Height          =   375
            Left            =   240
            TabIndex        =   141
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblVATTitle 
            Caption         =   "VAT paid = "
            Height          =   375
            Left            =   240
            TabIndex        =   140
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label lblCostpreVAT 
            Alignment       =   1  'Right Justify
            Caption         =   """VAT"""
            Height          =   375
            Left            =   1800
            TabIndex        =   139
            Top             =   3360
            Width           =   750
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Conversions"
         Height          =   4695
         Left            =   -73680
         TabIndex        =   89
         Top             =   1620
         Width           =   5535
         Begin VB.TextBox Pints 
            Height          =   375
            Left            =   480
            TabIndex        =   129
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton PL 
            Caption         =   "Convert"
            Height          =   375
            Left            =   2400
            TabIndex        =   128
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox Litres2 
            Height          =   375
            Left            =   3480
            TabIndex        =   127
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox Miles 
            Height          =   375
            Left            =   480
            TabIndex        =   124
            Top             =   4200
            Width           =   1215
         End
         Begin VB.CommandButton MK 
            Caption         =   "Convert"
            Height          =   375
            Left            =   2400
            TabIndex        =   123
            Top             =   4200
            Width           =   975
         End
         Begin VB.TextBox KM 
            Height          =   375
            Left            =   3480
            TabIndex        =   122
            Top             =   4200
            Width           =   1215
         End
         Begin VB.TextBox Metres2 
            Height          =   375
            Left            =   3480
            TabIndex        =   119
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton YM 
            Caption         =   "Convert"
            Height          =   375
            Left            =   2400
            TabIndex        =   118
            Top             =   3720
            Width           =   975
         End
         Begin VB.TextBox Yards 
            Height          =   375
            Left            =   480
            TabIndex        =   117
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox Kw 
            Height          =   375
            Left            =   3480
            TabIndex        =   112
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton BK 
            Caption         =   "Convert"
            Height          =   375
            Left            =   2400
            TabIndex        =   111
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox Bhp 
            Height          =   375
            Left            =   480
            TabIndex        =   110
            Top             =   3240
            Width           =   1215
         End
         Begin VB.TextBox Lbs 
            Height          =   375
            Left            =   480
            TabIndex        =   101
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton KP 
            Caption         =   "Convert"
            Height          =   375
            Left            =   2400
            TabIndex        =   100
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox Kg 
            Height          =   375
            Left            =   3480
            TabIndex        =   99
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox Litres 
            Height          =   375
            Left            =   3480
            TabIndex        =   98
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton GL 
            Caption         =   "Convert"
            Height          =   375
            Left            =   2400
            TabIndex        =   97
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox Gallons 
            Height          =   375
            Left            =   480
            TabIndex        =   96
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox Metres 
            Height          =   375
            Left            =   3480
            TabIndex        =   95
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Inches 
            Height          =   375
            Left            =   480
            TabIndex        =   94
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton FtM 
            Caption         =   "Convert"
            Height          =   375
            Left            =   2400
            TabIndex        =   93
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Feet 
            Height          =   375
            Left            =   480
            TabIndex        =   92
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton ICM 
            Caption         =   "Convert"
            Height          =   375
            Left            =   2400
            TabIndex        =   91
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Cm 
            Height          =   375
            Left            =   3480
            TabIndex        =   90
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pints"
            Height          =   195
            Left            =   1680
            TabIndex        =   131
            Top             =   1920
            Width           =   345
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Litres"
            Height          =   195
            Left            =   4680
            TabIndex        =   130
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Miles"
            Height          =   195
            Left            =   1680
            TabIndex        =   126
            Top             =   4320
            Width           =   360
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kilometres"
            Height          =   195
            Left            =   4680
            TabIndex        =   125
            Top             =   4320
            Width           =   720
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Metres"
            Height          =   195
            Left            =   4680
            TabIndex        =   121
            Top             =   3840
            Width           =   480
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Yards"
            Height          =   195
            Left            =   1680
            TabIndex        =   120
            Top             =   3840
            Width           =   405
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Imperial"
            Height          =   195
            Left            =   720
            TabIndex        =   116
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Metric"
            Height          =   195
            Left            =   3840
            TabIndex        =   115
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kilowatts"
            Height          =   195
            Left            =   4680
            TabIndex        =   114
            Top             =   3360
            Width           =   630
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bhp"
            Height          =   195
            Left            =   1680
            TabIndex        =   113
            Top             =   3360
            Width           =   285
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pounds"
            Height          =   195
            Left            =   1680
            TabIndex        =   109
            Top             =   2880
            Width           =   540
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Litres"
            Height          =   195
            Left            =   4680
            TabIndex        =   108
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cm."
            Height          =   195
            Left            =   4680
            TabIndex        =   107
            Top             =   1440
            Width           =   270
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kilograms"
            Height          =   195
            Left            =   4680
            TabIndex        =   106
            Top             =   2880
            Width           =   675
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gallons"
            Height          =   195
            Left            =   1680
            TabIndex        =   105
            Top             =   2400
            Width           =   525
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inches"
            Height          =   195
            Left            =   1680
            TabIndex        =   104
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Metres"
            Height          =   195
            Left            =   4680
            TabIndex        =   103
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Feet"
            Height          =   195
            Left            =   1680
            TabIndex        =   102
            Top             =   960
            Width           =   315
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Conversion method"
         Height          =   855
         Left            =   -74640
         TabIndex        =   86
         Top             =   780
         Width           =   1815
         Begin VB.OptionButton Option1 
            Caption         =   "Imperial - Metric"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Metric - Imperial"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Volume of a Sphere"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   240
         TabIndex        =   77
         Top             =   3420
         Width           =   3255
         Begin VB.CommandButton Command13 
            Caption         =   "Round"
            Height          =   375
            Left            =   240
            TabIndex        =   85
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox SphereAns 
            Height          =   315
            Left            =   720
            TabIndex        =   80
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   240
            TabIndex        =   79
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox SphereRadius 
            Height          =   285
            Left            =   720
            TabIndex        =   78
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Answer"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   960
            Width           =   525
         End
         Begin VB.Label Label24 
            Caption         =   "Formula: ¾ * Pi * Radius³"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Radius"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.TextBox Calc 
         Height          =   375
         Left            =   -72720
         TabIndex        =   70
         Top             =   1860
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -72720
         TabIndex        =   69
         Top             =   2340
         Width           =   1215
      End
      Begin VB.TextBox Answer 
         Height          =   375
         Left            =   -72720
         TabIndex        =   68
         Top             =   4620
         Width           =   1215
      End
      Begin VB.CommandButton Round 
         Caption         =   "Round"
         Height          =   375
         Left            =   -72720
         TabIndex        =   67
         Top             =   2700
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Round (2 d.p.)"
         Height          =   375
         Left            =   -72720
         TabIndex        =   66
         Top             =   3060
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Add Pi"
         Height          =   375
         Left            =   -72720
         TabIndex        =   65
         Top             =   3420
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add Pi (3 d.p.)"
         Height          =   375
         Left            =   -72720
         TabIndex        =   64
         Top             =   3780
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear Equation"
         Height          =   375
         Left            =   -72720
         TabIndex        =   63
         Top             =   4140
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Area of a Triangle"
         Height          =   2415
         Left            =   -71040
         TabIndex        =   49
         Top             =   780
         Width           =   2535
         Begin VB.TextBox TriHeight 
            Height          =   375
            Left            =   720
            TabIndex        =   56
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   720
            TabIndex        =   52
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TriAnswer 
            Height          =   375
            Left            =   720
            TabIndex        =   51
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox TriBase 
            Height          =   375
            Left            =   720
            TabIndex        =   50
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Height 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formula: 0.5 * base * height"
            Height          =   195
            Left            =   240
            TabIndex        =   55
            Top             =   2160
            Width           =   1950
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Answer"
            Height          =   195
            Left            =   150
            TabIndex        =   54
            Top             =   1800
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base"
            Height          =   195
            Left            =   240
            TabIndex        =   53
            Top             =   480
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Area of a Circle"
         Height          =   2415
         Left            =   -74520
         TabIndex        =   42
         Top             =   3420
         Width           =   2535
         Begin VB.TextBox CirRadius 
            Height          =   375
            Left            =   720
            TabIndex        =   45
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox CirAnswer 
            Height          =   375
            Left            =   720
            TabIndex        =   44
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   720
            TabIndex        =   43
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Radius"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Answer"
            Height          =   195
            Left            =   150
            TabIndex        =   47
            Top             =   1680
            Width           =   555
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formula: R² * Pi"
            Height          =   195
            Left            =   660
            TabIndex        =   46
            Top             =   2040
            Width           =   1110
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Volume of a Cylinder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   240
         TabIndex        =   31
         Top             =   660
         Width           =   3255
         Begin VB.TextBox CylRadius 
            Height          =   285
            Left            =   840
            TabIndex        =   35
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox CylHeight 
            Height          =   285
            Left            =   840
            TabIndex        =   34
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox CylAnswer 
            Height          =   285
            Left            =   840
            TabIndex        =   33
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label31 
            Caption         =   "Radius"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label32 
            Caption         =   "Height"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   735
            Width           =   615
         End
         Begin VB.Label Label33 
            Caption         =   "or Diameter ÷ 2"
            Height          =   255
            Left            =   1800
            TabIndex        =   39
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label34 
            Caption         =   "Height of cylinder"
            Height          =   255
            Left            =   1800
            TabIndex        =   38
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label35 
            Caption         =   "Formula: (R² x Height) x Pi"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label36 
            Caption         =   "Answer"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Volume of a Cube"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   3720
         TabIndex        =   20
         Top             =   660
         Width           =   3135
         Begin VB.TextBox CubeLength 
            Height          =   285
            Left            =   840
            TabIndex        =   25
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox CubeWidth 
            Height          =   285
            Left            =   840
            TabIndex        =   24
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox CubeHeight 
            Height          =   285
            Left            =   840
            TabIndex        =   23
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox CubeAnswer 
            Height          =   285
            Left            =   840
            TabIndex        =   21
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label37 
            Caption         =   "Length"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label38 
            Caption         =   "Width"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label39 
            Caption         =   "Height"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label40 
            Caption         =   "Formula: (Length x Width) x height"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label Label41 
            Caption         =   "Answer"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Trapazoid / Parallellogram"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -71280
         TabIndex        =   9
         Top             =   3420
         Width           =   3255
         Begin VB.TextBox TrapBase1 
            Height          =   285
            Left            =   840
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox TrapBase2 
            Height          =   285
            Left            =   840
            TabIndex        =   13
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TrapAltitude 
            Height          =   285
            Left            =   840
            TabIndex        =   12
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox TrapAnswer 
            Height          =   405
            Left            =   840
            TabIndex        =   11
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Side 1"
            Height          =   195
            Left            =   360
            TabIndex        =   19
            Top             =   375
            Width           =   450
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Side 2"
            Height          =   195
            Left            =   360
            TabIndex        =   18
            Top             =   735
            Width           =   450
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Altitude"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   1095
            Width           =   525
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formula: (Side1 + Side2) x height x ½"
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   2400
            Width           =   2610
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Answer"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   1560
            Width           =   525
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Area of a Square/Rectangle"
         Height          =   2415
         Left            =   -74520
         TabIndex        =   1
         Top             =   780
         Width           =   2535
         Begin VB.CommandButton CmdRect 
            Caption         =   "Calculate"
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   84
            Top             =   1150
            Width           =   975
         End
         Begin VB.TextBox RecAnswer 
            Height          =   375
            Left            =   720
            TabIndex        =   6
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox RecSide2 
            Height          =   375
            Left            =   720
            TabIndex        =   3
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox RecSide1 
            Height          =   375
            Left            =   720
            TabIndex        =   2
            Top             =   360
            Width           =   975
         End
         Begin VB.Label RecFormula 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formula: Side1 * Side2"
            Height          =   195
            Left            =   405
            TabIndex        =   8
            Top             =   2040
            Width           =   1620
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Answer"
            Height          =   195
            Left            =   150
            TabIndex        =   7
            Top             =   1680
            Width           =   555
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Side2"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Side1"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   480
            Width           =   405
         End
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "^ = ""To the Power of"". For example: 2^3 = 8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   795
         Left            =   -70680
         TabIndex        =   76
         Top             =   1620
         Width           =   1815
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use * (SHIFT + 8) to multiply"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   555
         Left            =   -70680
         TabIndex        =   75
         Top             =   2460
         Width           =   1815
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use / to divide"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   -70680
         TabIndex        =   74
         Top             =   3060
         Width           =   1815
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use \ to divide and produce an integer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   735
         Left            =   -70680
         TabIndex        =   73
         Top             =   3420
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use + to add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   -70680
         TabIndex        =   72
         Top             =   4260
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use - to subtract"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   -70680
         TabIndex        =   71
         Top             =   4620
         Width           =   1815
      End
   End
End
Attribute VB_Name = "MathsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngReturn As Long
Dim X As Integer

Dim sModel As String
Dim sVATRate As String
Dim sUrban As String
Dim sLitreCost As String

Dim dLitreCost As Double
Dim dGallonCost As Double
Dim dVATRate As Double
Dim dVATCost As Double
Dim dCost As Double
Dim dCostPreVat As Double
Function Agecal(mydate As Variant) As Integer
Dim Step1
Dim Step2
Dim Step3
    mydate = CDate(Text4)
    Dim Totdays As Long
    Totdays = DateDiff("y", mydate, Date) ' Total Number of days
    Step1 = Abs(Totdays / 365.25) ' Number of years
    Step2 = (Step1 - Int(Step1)) * 365.25 / 30.435 'Number of months
    Step3 = CInt((Step2 - Int(Step2)) * 30.435) ' Number of days


    If mydate < Date Then
        Age = "You are exactly " & Int(Step1) & " Year(s) " & Int(Step2) & " Month(s) and " & Int(Step3) & " Day(s)" & " old"
    Else
        Age = " There are " & Int(Step1) & " Year(s) " & Int(Step2) & " Month(s) and " & Int(Step3) & " Day(s)" & " To this date"
    End If
End Function


Private Sub spbLitreCost_SpinUp()

txtLitreCost.Text = CStr(CDbl(txtLitreCost.Text + 0.1))
Call CalcGallonCost

End Sub
Private Sub spbLitreCost_SpinDown()

txtLitreCost.Text = CStr(CDbl(txtLitreCost.Text - 0.1))
Call CalcGallonCost

End Sub


Public Sub LoadLitreCosts()

Dim Y As Integer

sLitreCost = Space(7)
lngReturn = _
    GetPrivateProfileString("Current", "LitreCost", "******", sLitreCost, 6, "SBDFUEL.INI")

sLitreCost = Trim$(Left$(sLitreCost, lngReturn))
txtLitreCost.Text = sLitreCost

'litre costs are in the format nn.n, eg. 62.9

'display gallon cost
dGallonCost = CDbl(sLitreCost) * 0.04546
lblGallonCost.Caption = dGallonCost
lblGallonCost.Caption = Format(dGallonCost, "£0.00")

End Sub

Public Sub CalcGallonCost()

dGallonCost = CDbl(txtLitreCost.Text) * 0.04546
lblGallonCost.Caption = dGallonCost
lblGallonCost = Format(lblGallonCost, "£0.00")

End Sub

Function EvaluateString(ByVal sStrCalc As String) As Variant

Dim vntTotal As Variant ' Total variable
    Dim vntNumber As Variant
    Dim sOperator As String
    Dim sCalc As String ' Current parameter in calculation str
    Dim sCurChar As String * 1 ' String comparisons
    Dim nCounter As Integer
    
    
    
    '
    ' Reads a string and evaluates the numer
    '     ic result. If there is a syntax error
    ' the function returns "Error".
    '
    On Error GoTo Error_EvaluateString
    
    sCalc = sStrCalc
    vntTotal = 0
    sOperator = "+"
    
    


    Do While sOperator <> "" And sCalc <> ""
        
        vntNumber = Val(sCalc)
        


        If IsNumeric(Left$(sCalc, 1)) Then
            
            sCalc = Mid$(sCalc, Len(Trim$(Str(vntNumber))) + 1)
            


            Select Case sOperator
                
                Case "+": vntTotal = vntTotal + vntNumber
                Case "-": vntTotal = vntTotal - vntNumber
                Case "*": vntTotal = vntTotal * vntNumber
                Case "/": vntTotal = vntTotal / vntNumber
                Case "^": vntTotal = vntTotal ^ vntNumber
                Case "\": vntTotal = vntTotal \ vntNumber
                End Select
        
    Else
        
        sOperator = Left$(sCalc, 1)
        sCalc = Mid$(sCalc, 2)
        
    End If
    
Loop


EvaluateString = vntTotal

Exit Function


Error_EvaluateString:

EvaluateString = "Error"

Exit Function


End Function




Private Sub BK_Click()
If Kw.Enabled = False Then
    Kw = Bhp * 1.36
    Kw = Format(Kw, "##,###.00")
ElseIf Bhp.Enabled = False Then
    Bhp = Kw / 1.36
    Bhp = Format(Bhp, "##,###.00")
End If
End Sub

Private Sub cmdCalc_Click()
    dVATRate = 17.5
    pnlMessage.Caption = ""
    
    If Trim(txtMileage.Text) = "" Then
        pnlMessage.Caption = "Enter Mileage"
        'Beep
        txtMileage.SetFocus
        Exit Sub
    End If
       
    'mileage/urban mpg * 4.546 (litres to gallons)
    ' * price per litre
    
    'calculate gallons used
    dCost = CDbl(txtMileage.Text) / CDbl(txtMpG.Text)
    pnlMessage.Caption = "Gallons used = " & Format(dCost, "0.00")
    
    'convert gallons to litres and multiply by litre price
    dCost = (dCost * 4.546 * CDbl(txtLitreCost.Text)) / 100
    txtCost.Text = Format(dCost, "£0.00")
    
     'calculate VAT
    dCostPreVat = dCost / dVATRate
    dVATCost = dCost - dCostPreVat
    lblVAT.Caption = Format(dVATCost, "£0.00")
    
    lblCostpreVAT.Caption = Format(dCostPreVat, "£0.00")
    

End Sub

Private Sub CmdRect_Click(Index As Integer)
RecAnswer = RecSide1 * RecSide2
End Sub


Private Sub Command1_Click()
Dim sAbout As String

sAbout = "Fuel Costs" & vbCrLf & _
    "Author :  Paul Davies" & vbCrLf & vbCrLf & _
    "This application calculates fuel costs for " _
    & vbCrLf & " expense claims.  The formula is as follows ..." _
    & vbCrLf & vbCrLf & _
    "(mileage / urban mpg) * 4.546 * price per litre"

MsgBox sAbout, vbInformation
End Sub

Private Sub Command10_Click()
Answer = EvaluateString(Calc)
Answer = Format(Answer, "##,###.00")
End Sub

Private Sub Command11_Click()
Answer = EvaluateString(Calc)
End Sub


Private Sub Command12_Click()
SphereAns = 3 / 4 * 3.14159265359 * SphereRadius ^ 3
End Sub

Private Sub Command13_Click()
SphereAns = 3 / 4 * 3.14159265359 * SphereRadius ^ 3
SphereAns = Format(SphereAns, "##,###.00")
End Sub



Private Sub Command14_Click()
Total = Total + Val(Number)
Sum = Sum + 1
End Sub

Private Sub Command15_Click()
MsgBox "Just enter any date in the format 'DD/MM/YYYY' and click the button below"
End Sub

Private Sub Command16_Click()
On Error GoTo err
Text2 = Text1.Text / (Tax.Text + 100) * 100
Text3.Text = Text1.Text - Text2.Text
Text2.Text = Format(Text2, "0.00")
Text3.Text = Format(Text3, "0.00")
err:
Exit Sub
End Sub

Private Sub Command17_Click()
On Error GoTo err
Dim Ans2
Text3 = Text1.Text / 100 * Tax.Text
Ans2 = Text1.Text / 100 * Tax.Text + Text1.Text
Text2 = Format(Ans2, "0.00")
err:
Exit Sub
End Sub

Private Sub Command18_Click()
Average = Total / Sum
Average = Format(Average, "##,###.##")
End Sub

Private Sub Command19_Click()
Dim mydate

If IsDate(Text4) = False Then
        MsgBox " Please enter a valid date. "
        Text4.SetFocus
        Text4.SelStart = 0: Text4.SelLength = Len(Text4)
        Exit Sub
    End If
    Agecal (mydate)
End Sub




Private Sub Command2_Click()
TriAnswer = 0.5 * TriBase * TriHeight
TriAnswer = Format(TriAnswer, "##,###.00")
End Sub


Private Sub Command20_Click()
End
End Sub

Private Sub Command21_Click()

   If a.Text = "" Then
      PythagorumTheorum = Sqr(c ^ 2 - b ^ 2)
      Label63.Visible = True
      Label63.Caption = "a ="
   ElseIf b.Text = "" Then
      PythagorumTheorum = Sqr(c ^ 2 - a ^ 2)
      Label63.Visible = True
      Label63.Caption = "b ="
   ElseIf c.Text = "" Then
      PythagorumTheorum = Sqr(a ^ 2 + Val(b ^ 2))
      Label63.Visible = True
      Label63.Caption = "c ="
   Else: PythagorumTheorum = "?"
   ' returns a "?" if no "?" was entered for one of the values
   End If

End Sub

Private Sub Command3_Click()
CirAnswer = CirRadius ^ 2 * 3.142
CirAnswer = Format(CirAnswer, "##,###.00")
End Sub












Private Sub Command4_Click()
TrapAnswer = (TrapBase1 + TrapBase2) * TrapAltitude * 0.5
TrapAnswer = Format(TrapAnswer, "##,###.00")
End Sub


Private Sub Command5_Click()
CylAnswer = CylRadius ^ 2 * CylHeight * 3.142
CylAnswer = Format(CylAnswer, "##,###.00")
End Sub

Private Sub Command6_Click()
CubeAnswer = (CubeLength * CubeWidth) * CubeHeight
CubeAnswer = Format(CubeAnswer, "##,###.00")
End Sub


Private Sub Command7_Click()
Calc = ""
Calc.SetFocus ' Set the focus to the input box
End Sub

Private Sub Command8_Click()
Calc = Calc + "3.142" ' This is Pi to 3 decimal places
End Sub

Private Sub Command9_Click()
Calc = Calc + "3.14159265359" ' This is Pi to 11 decimal places
End Sub

Private Sub Form_Load()
Me.Icon = Picture1
Option1 = True
End Sub

Private Sub FtM_Click()
If Metres.Enabled = False Then
    Metres = Feet * 100 / 2.54 / 12
    Metres = Format(Metres, "##,###.00")
ElseIf Feet.Enabled = False Then
    Feet = Metres / 100 * 2.54 * 12
    Feet = Format(Feet, "##,###.00")
End If
End Sub

Private Sub GL_Click()
If Litres.Enabled = False Then
    Litres = Gallons * 4.55
    Litres = Format(Litres, "##,###.00")
ElseIf Gallons.Enabled = False Then
    Gallons = Litres / 4.55
    Gallons = Format(Gallons, "##,###.00")
End If
End Sub

Private Sub ICM_Click()
If Cm.Enabled = False Then
    Cm = Inches / 2.54
    Cm = Format(Cm, "##,###.00")
ElseIf Inches.Enabled = False Then
    Inches = Cm * 2.54
    Inches = Format(Inches, "##,###.00")
End If

End Sub

Private Sub KP_Click()
If Kg.Enabled = False Then
    Kg = Lbs / 2.25
    Kg = Format(Kg, "##,###.00")
ElseIf Lbs.Enabled = False Then
    Lbs = Kg * 2.25
    Lbs = Format(Lbs, "##,###.00")
End If
End Sub

Private Sub MK_Click()
If KM.Enabled = False Then
    KM = Miles * 1.6
    KM = Format(KM, "##,###.00")
ElseIf Miles.Enabled = False Then
    Miles = KM / 1.6
    Miles = Format(Miles, "##,###.00")
End If

End Sub

Private Sub Option1_Click()

Label45.FontBold = False
Label46.FontBold = True

Metres.Enabled = False
Feet.Enabled = True

Cm.Enabled = False
Inches.Enabled = True

Litres.Enabled = False
Gallons.Enabled = True

Kg.Enabled = False
Lbs.Enabled = True

Kw.Enabled = False
Bhp.Enabled = True

Metres2.Enabled = False
Yards.Enabled = True

KM.Enabled = False
Miles.Enabled = True

Litres2.Enabled = False
Pints.Enabled = True

End Sub

Private Sub Option2_Click()

Label45.FontBold = True
Label46.FontBold = False

Metres.Enabled = True
Feet.Enabled = False

Cm.Enabled = True
Inches.Enabled = False

Litres.Enabled = True
Gallons.Enabled = False

Kg.Enabled = True
Lbs.Enabled = False

Kw.Enabled = True
Bhp.Enabled = False

Metres2.Enabled = True
Yards.Enabled = False

KM.Enabled = True
Miles.Enabled = False

Litres2.Enabled = True
Pints.Enabled = False

End Sub


Private Sub Option3_Click()
Command16.Visible = False
Command17.Visible = True
Command16.Enabled = False
Command17.Enabled = True
End Sub

Private Sub Option4_Click()
Command16.Visible = True
Command17.Visible = False
Command16.Enabled = True
Command17.Enabled = False
End Sub

Private Sub PL_Click()
If Litres2.Enabled = False Then
    Litres2 = Pints * 0.57
    Litres2 = Format(Litres2, "##,###.00")
ElseIf Pints.Enabled = False Then
    Pints = Litres2 / 0.57
    Pints = Format(Pints, "##,###.00")
End If
End Sub

Private Sub Round_Click()
Answer = EvaluateString(Calc)
Answer = Format(Answer, "##,###")
End Sub


Private Sub Timer1_Timer()
If Me.Icon = Picture1 Then
    Me.Icon = Picture2
Exit Sub
End If

If Me.Icon = Picture2 Then
    Me.Icon = Picture3
Exit Sub
End If

If Me.Icon = Picture3 Then
    Me.Icon = Picture4
Exit Sub
End If

If Me.Icon = Picture4 Then
    Me.Icon = Picture5
Exit Sub
End If

If Me.Icon = Picture5 Then
    Me.Icon = Picture1
Exit Sub
End If

End Sub


Private Sub YM_Click()
If Metres2.Enabled = False Then
    Metres2 = Yards - (Yards / 10)
    Metres2 = Format(Metres2, "##,###.00")
ElseIf Yards.Enabled = False Then
    Yards = Metres2 + (Metres2 / 10)
    Yards = Format(Yards, "##,###.00")
End If
End Sub


