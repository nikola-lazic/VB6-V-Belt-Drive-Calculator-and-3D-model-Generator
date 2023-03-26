VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remeni kaišni prenosnici v2"
   ClientHeight    =   8700
   ClientLeft      =   5070
   ClientTop       =   4935
   ClientWidth     =   16695
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   16695
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16725
      _ExtentX        =   29501
      _ExtentY        =   15478
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Proraèun"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "CAD"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame8(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame8(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame8 
         Caption         =   "Dimenzije remena"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Index           =   1
         Left            =   -65280
         TabIndex        =   138
         Top             =   4200
         Width           =   6855
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Index           =   9
            Left            =   1320
            TabIndex        =   182
            Top             =   2280
            Width           =   240
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Index           =   8
            Left            =   1320
            TabIndex        =   181
            Top             =   1920
            Width           =   240
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Index           =   7
            Left            =   1320
            TabIndex        =   180
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Index           =   6
            Left            =   1320
            TabIndex        =   179
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            Height          =   240
            Index           =   15
            Left            =   720
            TabIndex        =   153
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   240
            Index           =   14
            Left            =   480
            TabIndex        =   152
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "h"
            Height          =   240
            Index           =   12
            Left            =   240
            TabIndex        =   151
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            Height          =   240
            Index           =   11
            Left            =   720
            TabIndex        =   150
            Top             =   1920
            Width           =   480
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   240
            Index           =   10
            Left            =   480
            TabIndex        =   149
            Top             =   1920
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "d"
            Height          =   240
            Index           =   9
            Left            =   360
            TabIndex        =   148
            Top             =   2400
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "h"
            Height          =   240
            Index           =   8
            Left            =   240
            TabIndex        =   147
            Top             =   1920
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            Height          =   240
            Index           =   7
            Left            =   720
            TabIndex        =   146
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   240
            Index           =   6
            Left            =   480
            TabIndex        =   145
            Top             =   1560
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "d"
            Height          =   240
            Index           =   5
            Left            =   360
            TabIndex        =   144
            Top             =   1680
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "b"
            Height          =   240
            Index           =   4
            Left            =   240
            TabIndex        =   143
            Top             =   1560
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            Height          =   240
            Index           =   3
            Left            =   720
            TabIndex        =   142
            Top             =   1200
            Width           =   480
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   240
            Index           =   2
            Left            =   480
            TabIndex        =   141
            Top             =   1200
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   240
            Index           =   1
            Left            =   360
            TabIndex        =   140
            Top             =   1320
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "b"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   139
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image Image5 
            Height          =   2115
            Left            =   2040
            Picture         =   "Form1.frx":0038
            Top             =   840
            Width           =   2400
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Dimenzije remenice"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Index           =   0
         Left            =   -74880
         TabIndex        =   137
         Top             =   4200
         Width           =   9495
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "o"
            Height          =   240
            Index           =   5
            Left            =   1080
            TabIndex        =   178
            Top             =   2400
            Width           =   120
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Index           =   4
            Left            =   1320
            TabIndex        =   177
            Top             =   2160
            Width           =   240
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Index           =   3
            Left            =   1320
            TabIndex        =   176
            Top             =   1800
            Width           =   240
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Index           =   2
            Left            =   1320
            TabIndex        =   175
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Index           =   1
            Left            =   1320
            TabIndex        =   174
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Index           =   0
            Left            =   1320
            TabIndex        =   173
            Top             =   720
            Width           =   240
         End
         Begin VB.Image Image6 
            Height          =   3465
            Left            =   2040
            Picture         =   "Form1.frx":84D8
            Top             =   240
            Width           =   7215
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0.0"
            Height          =   240
            Index           =   37
            Left            =   720
            TabIndex        =   172
            Top             =   2520
            Width           =   360
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   240
            Index           =   36
            Left            =   480
            TabIndex        =   171
            Top             =   2520
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "a"
            BeginProperty Font 
               Name            =   "GreekC"
               Size            =   9.75
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   34
            Left            =   240
            TabIndex        =   170
            Top             =   2520
            Width           =   165
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            Height          =   240
            Index           =   33
            Left            =   720
            TabIndex        =   169
            Top             =   2160
            Width           =   480
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   240
            Index           =   32
            Left            =   480
            TabIndex        =   168
            Top             =   2160
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "t"
            Height          =   240
            Index           =   30
            Left            =   240
            TabIndex        =   167
            Top             =   2160
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            Height          =   240
            Index           =   29
            Left            =   720
            TabIndex        =   166
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   240
            Index           =   28
            Left            =   480
            TabIndex        =   165
            Top             =   1800
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "f"
            Height          =   240
            Index           =   27
            Left            =   240
            TabIndex        =   164
            Top             =   1800
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            Height          =   240
            Index           =   26
            Left            =   720
            TabIndex        =   163
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   240
            Index           =   25
            Left            =   480
            TabIndex        =   162
            Top             =   1440
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "e"
            Height          =   240
            Index           =   23
            Left            =   240
            TabIndex        =   161
            Top             =   1440
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            Height          =   240
            Index           =   22
            Left            =   720
            TabIndex        =   160
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   240
            Index           =   21
            Left            =   480
            TabIndex        =   159
            Top             =   1080
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "c"
            Height          =   240
            Index           =   19
            Left            =   240
            TabIndex        =   158
            Top             =   1080
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            Height          =   240
            Index           =   18
            Left            =   720
            TabIndex        =   157
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   240
            Index           =   17
            Left            =   480
            TabIndex        =   156
            Top             =   720
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   240
            Index           =   16
            Left            =   360
            TabIndex        =   155
            Top             =   840
            Width           =   120
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "b"
            Height          =   240
            Index           =   13
            Left            =   240
            TabIndex        =   154
            Top             =   720
            Width           =   120
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Izrada 3D modela"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74880
         TabIndex        =   118
         Top             =   480
         Width           =   16455
         Begin VB.OptionButton Option2 
            Caption         =   "Nezavisno modeliranje remenice"
            Height          =   375
            Left            =   240
            TabIndex        =   130
            Top             =   720
            Width           =   3975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Predhodno proracunata remenica"
            Height          =   375
            Left            =   240
            TabIndex        =   129
            Top             =   360
            Width           =   4455
         End
         Begin VB.Frame Frame5 
            Caption         =   "Geometrija"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   120
            TabIndex        =   119
            Top             =   1200
            Width           =   9375
            Begin VB.TextBox Text9 
               Height          =   375
               Left            =   3960
               TabIndex        =   185
               Text            =   "0"
               Top             =   1920
               Width           =   855
            End
            Begin VB.TextBox Text8 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   3960
               TabIndex        =   184
               Text            =   "0.000"
               Top             =   1920
               Width           =   1455
            End
            Begin VB.OptionButton Option4 
               Height          =   255
               Left            =   5880
               TabIndex        =   136
               Top             =   1920
               Width           =   255
            End
            Begin VB.OptionButton Option3 
               Height          =   255
               Left            =   5880
               TabIndex        =   135
               Top             =   1440
               Width           =   255
            End
            Begin VB.ComboBox Combo7 
               Height          =   360
               Left            =   3960
               Style           =   2  'Dropdown List
               TabIndex        =   122
               Top             =   480
               Width           =   1500
            End
            Begin VB.ComboBox Combo8 
               Height          =   360
               Left            =   3960
               Style           =   2  'Dropdown List
               TabIndex        =   121
               Top             =   960
               Width           =   5175
            End
            Begin VB.TextBox Text7 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   3960
               TabIndex        =   120
               Text            =   "0.000"
               Top             =   1440
               Width           =   1455
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               Caption         =   "Broj žlebova:               z"
               Height          =   240
               Index           =   1
               Left            =   240
               TabIndex        =   183
               Top             =   1920
               Width           =   3480
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Preènik gonjene remenice:"
               Height          =   240
               Left            =   240
               TabIndex        =   134
               Top             =   1920
               Width           =   3000
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3720
               TabIndex        =   133
               Top             =   2040
               Width           =   105
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "d"
               Height          =   240
               Left            =   3600
               TabIndex        =   132
               Top             =   1920
               Width           =   120
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "mm"
               Height          =   240
               Left            =   5520
               TabIndex        =   131
               Top             =   1920
               Width           =   240
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Profil kaiša:"
               Height          =   240
               Left            =   240
               TabIndex        =   128
               Top             =   480
               Width           =   1560
            End
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               Caption         =   "mm"
               Height          =   240
               Left            =   5520
               TabIndex        =   127
               Top             =   1440
               Width           =   240
            End
            Begin VB.Label Label69 
               AutoSize        =   -1  'True
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3720
               TabIndex        =   126
               Top             =   1560
               Width           =   105
            End
            Begin VB.Label Label76 
               AutoSize        =   -1  'True
               Caption         =   "d"
               Height          =   240
               Left            =   3600
               TabIndex        =   125
               Top             =   1440
               Width           =   120
            End
            Begin VB.Label Label84 
               AutoSize        =   -1  'True
               Caption         =   "Preènik pogonske remenice:"
               Height          =   240
               Left            =   240
               TabIndex        =   124
               Top             =   1440
               Width           =   3120
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "Tip kaiša:"
               Height          =   240
               Left            =   240
               TabIndex        =   123
               Top             =   960
               Width           =   1200
            End
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Napravi &3D model"
         Height          =   495
         Left            =   -60960
         TabIndex        =   117
         Top             =   8040
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "I&zlaz"
         Height          =   495
         Left            =   -59640
         TabIndex        =   116
         Top             =   8040
         Width           =   1200
      End
      Begin VB.CommandButton Command7 
         Caption         =   "I&zlaz"
         Height          =   495
         Left            =   15360
         TabIndex        =   37
         Top             =   8040
         Width           =   1200
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Izveštaj"
         Height          =   495
         Left            =   14040
         TabIndex        =   9
         Top             =   8040
         Width           =   1200
      End
      Begin VB.Frame Frame4 
         Caption         =   "Rezultati"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7335
         Left            =   9600
         TabIndex        =   36
         Top             =   480
         Width           =   6975
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   10
            Left            =   4920
            TabIndex        =   115
            Top             =   4080
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   9
            Left            =   4920
            TabIndex        =   114
            Top             =   4440
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   8
            Left            =   4920
            TabIndex        =   113
            Top             =   4800
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   7
            Left            =   4920
            TabIndex        =   112
            Top             =   1320
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   6
            Left            =   4920
            TabIndex        =   111
            Top             =   1680
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   4920
            TabIndex        =   110
            Top             =   2040
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   4920
            TabIndex        =   109
            Top             =   2400
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   4920
            TabIndex        =   108
            Top             =   2760
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   4920
            TabIndex        =   107
            Top             =   3720
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   4920
            TabIndex        =   106
            Top             =   960
            Width           =   150
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Kontrola:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label92 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6480
            TabIndex        =   91
            Top             =   1680
            Width           =   300
         End
         Begin VB.Label Label91 
            AutoSize        =   -1  'True
            Caption         =   "d"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4680
            TabIndex        =   90
            Top             =   1800
            Width           =   105
         End
         Begin VB.Label Label90 
            AutoSize        =   -1  'True
            Caption         =   "L"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4440
            TabIndex        =   89
            Top             =   1680
            Width           =   150
         End
         Begin VB.Label Label89 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   88
            Top             =   1680
            Width           =   810
         End
         Begin VB.Label Label85 
            AutoSize        =   -1  'True
            Caption         =   "Standardna vrednost dužine kaiša:"
            Height          =   240
            Left            =   120
            TabIndex        =   87
            Top             =   1755
            Width           =   3960
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6480
            TabIndex        =   86
            Top             =   1320
            Width           =   300
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            Caption         =   "dr"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4680
            TabIndex        =   85
            Top             =   1440
            Width           =   210
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "L"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4440
            TabIndex        =   84
            Top             =   1320
            Width           =   150
         End
         Begin VB.Label Label79 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   83
            Top             =   1320
            Width           =   810
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            Caption         =   "Raèunska vrednost dužine kaiša:"
            Height          =   240
            Left            =   120
            TabIndex        =   82
            Top             =   1395
            Width           =   3720
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "s"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4560
            TabIndex        =   80
            Top             =   1080
            Width           =   105
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            Caption         =   "z"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   4440
            TabIndex        =   79
            Top             =   2760
            Width           =   150
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   78
            Top             =   2760
            Width           =   810
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "Broj žlebova:"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   77
            Top             =   2835
            Width           =   1560
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "z"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4440
            TabIndex        =   76
            Top             =   2400
            Width           =   150
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   75
            Top             =   2400
            Width           =   810
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            Caption         =   "Raèunska vrednost broja žlebova:"
            Height          =   240
            Left            =   120
            TabIndex        =   74
            Top             =   2475
            Width           =   3840
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "r"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4560
            TabIndex        =   73
            Top             =   2520
            Width           =   105
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6480
            TabIndex        =   72
            Top             =   2040
            Width           =   300
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "s"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4560
            TabIndex        =   71
            Top             =   2160
            Width           =   105
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "Stvarna vrednost osnog rastojanja:"
            Height          =   240
            Left            =   120
            TabIndex        =   70
            Top             =   2115
            Width           =   4080
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   69
            Top             =   2040
            Width           =   810
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "a"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4440
            TabIndex        =   68
            Top             =   2040
            Width           =   150
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Geometrija:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   1485
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "i"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4440
            TabIndex        =   66
            Top             =   960
            Width           =   150
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   65
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Stvarna vrednost prenosnog odnosa:"
            Height          =   240
            Left            =   120
            TabIndex        =   64
            Top             =   1035
            Width           =   4080
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "doz"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4560
            TabIndex        =   63
            Top             =   4920
            Width           =   315
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "fs"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   62
            Top             =   4800
            Width           =   270
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "s"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6360
            TabIndex        =   61
            Top             =   4725
            Width           =   150
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "-1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6480
            TabIndex        =   60
            Top             =   4680
            Width           =   210
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   59
            Top             =   4800
            Width           =   810
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Dozvoljena uèestanost savijanja:"
            Height          =   240
            Left            =   120
            TabIndex        =   58
            Top             =   4770
            Width           =   3840
         End
         Begin VB.Label Label78 
            AutoSize        =   -1  'True
            Caption         =   "fs"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   57
            Top             =   4440
            Width           =   270
         End
         Begin VB.Label Label86 
            AutoSize        =   -1  'True
            Caption         =   "s"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6360
            TabIndex        =   56
            Top             =   4365
            Width           =   150
         End
         Begin VB.Label Label88 
            AutoSize        =   -1  'True
            Caption         =   "-1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6480
            TabIndex        =   55
            Top             =   4320
            Width           =   210
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   54
            Top             =   4440
            Width           =   810
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Uèestanost savijanja:"
            Height          =   240
            Left            =   120
            TabIndex        =   53
            Top             =   4410
            Width           =   2520
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "max"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4560
            TabIndex        =   52
            Top             =   4080
            Width           =   315
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "u"
            BeginProperty Font 
               Name            =   "GreekC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4320
            TabIndex        =   51
            Top             =   3960
            Width           =   165
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "m/s"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6360
            TabIndex        =   50
            Top             =   4020
            Width           =   450
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   49
            Top             =   4080
            Width           =   810
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Maksimalna obimna brzina:"
            Height          =   240
            Left            =   120
            TabIndex        =   48
            Top             =   4050
            Width           =   3000
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "u"
            BeginProperty Font 
               Name            =   "GreekC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4320
            TabIndex        =   47
            Top             =   3600
            Width           =   165
         End
         Begin VB.Label Label87 
            AutoSize        =   -1  'True
            Caption         =   "m/s"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6360
            TabIndex        =   46
            Top             =   3660
            Width           =   450
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   45
            Top             =   3720
            Width           =   810
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Obimna brzina:"
            Height          =   240
            Left            =   120
            TabIndex        =   44
            Top             =   3690
            Width           =   1680
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   4920
            TabIndex        =   43
            Top             =   600
            Width           =   150
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Raèunska vrednost gonjene remenice:"
            Height          =   240
            Left            =   120
            TabIndex        =   42
            Top             =   675
            Width           =   4200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   41
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "d"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4440
            TabIndex        =   40
            Top             =   600
            Width           =   150
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "2r"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4560
            TabIndex        =   39
            Top             =   720
            Width           =   210
         End
         Begin VB.Label Label81 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6480
            TabIndex        =   38
            Top             =   600
            Width           =   300
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Proracun"
         Height          =   495
         Left            =   12720
         TabIndex        =   8
         Top             =   8040
         Width           =   1200
      End
      Begin VB.Frame Frame3 
         Caption         =   "Geometrija"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   22
         Top             =   4800
         Width           =   9375
         Begin VB.ComboBox Combo4 
            Height          =   360
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   480
            Width           =   1500
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Izbor jaceg profila"
            Height          =   375
            Left            =   6600
            TabIndex        =   4
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text6 
            Height          =   360
            Left            =   3960
            TabIndex        =   7
            Top             =   2400
            Width           =   1500
         End
         Begin VB.TextBox Text5 
            Height          =   360
            Left            =   3960
            TabIndex        =   23
            Top             =   1920
            Width           =   1500
         End
         Begin VB.ComboBox Combo6 
            Height          =   360
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1440
            Width           =   1500
         End
         Begin VB.ComboBox Combo5 
            Height          =   360
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   960
            Width           =   5175
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Profil kaiša:"
            Height          =   240
            Left            =   240
            TabIndex        =   105
            Top             =   480
            Width           =   1560
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            Caption         =   "Label75"
            Height          =   240
            Left            =   6360
            TabIndex        =   81
            Top             =   2400
            Width           =   840
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Left            =   5520
            TabIndex        =   35
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Left            =   5520
            TabIndex        =   34
            Top             =   1920
            Width           =   240
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Osno rastojanje:"
            Height          =   240
            Left            =   240
            TabIndex        =   33
            Top             =   2400
            Width           =   1920
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   240
            Left            =   3720
            TabIndex        =   32
            Top             =   2400
            Width           =   120
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "d"
            Height          =   240
            Left            =   3600
            TabIndex        =   31
            Top             =   1920
            Width           =   120
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3720
            TabIndex        =   30
            Top             =   2040
            Width           =   105
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   240
            Left            =   5520
            TabIndex        =   29
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Preènik gonjene remenice:"
            Height          =   240
            Left            =   240
            TabIndex        =   28
            Top             =   1920
            Width           =   3000
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3720
            TabIndex        =   27
            Top             =   1560
            Width           =   105
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "d"
            Height          =   240
            Left            =   3600
            TabIndex        =   26
            Top             =   1440
            Width           =   120
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Preènik pogonske remenice:"
            Height          =   240
            Left            =   240
            TabIndex        =   25
            Top             =   1440
            Width           =   3120
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Tip kaiša:"
            Height          =   240
            Left            =   240
            TabIndex        =   24
            Top             =   960
            Width           =   1200
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ulazni podaci"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   9375
         Begin VB.Frame Frame2 
            Caption         =   "Faktor radnih uslova"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   120
            TabIndex        =   92
            Top             =   1800
            Width           =   9135
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3840
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   480
               Width           =   5175
            End
            Begin VB.ComboBox Combo2 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3840
               Style           =   2  'Dropdown List
               TabIndex        =   95
               Top             =   960
               Width           =   5175
            End
            Begin VB.ComboBox Combo3 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3840
               Style           =   2  'Dropdown List
               TabIndex        =   94
               Top             =   1440
               Width           =   5175
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   3840
               TabIndex        =   93
               Top             =   1920
               Width           =   1500
            End
            Begin VB.Label Label93 
               AutoSize        =   -1  'True
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3600
               TabIndex        =   102
               Top             =   2040
               Width           =   105
            End
            Begin VB.Label Label94 
               AutoSize        =   -1  'True
               Caption         =   "K"
               Height          =   240
               Left            =   3480
               TabIndex        =   101
               Top             =   1920
               Width           =   120
            End
            Begin VB.Label Label95 
               AutoSize        =   -1  'True
               Caption         =   "Pogonska mašina:"
               Height          =   240
               Left            =   120
               TabIndex        =   100
               Top             =   480
               Width           =   1920
            End
            Begin VB.Label Label96 
               AutoSize        =   -1  'True
               Caption         =   "Dnevni rad:"
               Height          =   240
               Left            =   120
               TabIndex        =   99
               Top             =   960
               Width           =   1320
            End
            Begin VB.Label Label97 
               AutoSize        =   -1  'True
               Caption         =   "Radna mašina:"
               Height          =   240
               Left            =   120
               TabIndex        =   98
               Top             =   1440
               Width           =   1560
            End
            Begin VB.Label Label98 
               AutoSize        =   -1  'True
               Caption         =   "Faktor radnih uslova:"
               Height          =   240
               Left            =   120
               TabIndex        =   97
               Top             =   1920
               Width           =   2520
            End
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   3960
            TabIndex        =   3
            Top             =   1320
            Width           =   1500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   3960
            TabIndex        =   2
            Top             =   840
            Width           =   1500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   3960
            TabIndex        =   0
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "i"
            Height          =   240
            Left            =   3720
            TabIndex        =   21
            Top             =   1320
            Width           =   120
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Prenosni odnos:"
            Height          =   240
            Left            =   240
            TabIndex        =   20
            Top             =   1440
            Width           =   1800
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "n"
            Height          =   240
            Left            =   3600
            TabIndex        =   19
            Top             =   840
            Width           =   120
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3720
            TabIndex        =   18
            Top             =   960
            Width           =   105
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "P"
            Height          =   240
            Left            =   3600
            TabIndex        =   17
            Top             =   360
            Width           =   120
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "-1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5880
            TabIndex        =   16
            Top             =   840
            Width           =   210
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "min"
            Height          =   240
            Left            =   5520
            TabIndex        =   15
            Top             =   960
            Width           =   360
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Broj obrtaja na ulazu:"
            Height          =   240
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   2640
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "kW"
            Height          =   240
            Left            =   5520
            TabIndex        =   13
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3720
            TabIndex        =   12
            Top             =   480
            Width           =   105
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Snaga na ulazu:"
            Height          =   240
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   1800
         End
      End
      Begin VB.Image Image3 
         Height          =   405
         Left            =   -74760
         Picture         =   "Form1.frx":1987C
         Top             =   8160
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ka As Variant 'Faktor radnih uslova
Public P1 As Variant 'Snaga
Public n1 As Variant 'Ulazni broj obrtaja
Public i As Variant 'Prenos odnos
Public fsdoz As Double 'Dozvoljenja ucestanost savijanja
Public fs As Double 'Ucestanost savijanja
Public d1 As Double 'Precnik pogonske remenice
Public d2 As Double 'Precnik gonjene remenice
Public iss As Double 'Stvarna vrednost prenosnog odnosa
Public amin As Double 'Minimalna vrednost osnog rastojanja
Public amax As Double 'Maksimalna vrednost osnog rastojanja
Public az As Double 'Zeljena vrednost osnog rastojanja
Public Pi  As Double '3.1415
Public alfa As Double  'Ugao nagiba ogranka
Public beta As Double  'Obvojni ugao
Public Ldr As Double  'Racunska vrednost duzine kaisa
Public Ld As Double  'Standardna vrednost duzine kaisa
Public deltaL As Double  'Korekcija duzine kaisa
Public a As Double 'Stvarno osno rastojanje
Public Cbeta As Double 'Faktor obvojnog ugla
Public CL As Double 'Faktor duzine uskih remena
Public Pm As Double 'Merodavna snaga
Public Pn As Double 'Nominalna snaga po remenu
Public deltaP As Double 'Dodatna snaga po remenu
Public v As Double 'Obimna brzina
Public vmax As Double 'Maksimalna obimna brzina
Public zr As Double 'Racunska vrednost broja zlebova
Public z As Integer 'Celobrojna vrednost broja zlebova
Public ime_prezime As String 'Promenljiva za unos imena i prezimena
Public b1_CAD As Double 'CAD geometrijska velicina
Public c_CAD As Double 'CAD geometrijska velicina
Public E_CAD As Double 'CAD geometrijska velicina
Public f_CAD As Double 'CAD geometrijska velicina
Public t_CAD As Double 'CAD geometrijska velicina
Public alfa_CAD As Double 'CAD geometrijska velicina
Public b0_CAD As Double 'CAD geometrijska velicina
Public bd_CAD As Double 'CAD geometrijska velicina
Public h_CAD As Double 'CAD geometrijska velicina
Public hd_CAD As Double 'CAD geometrijska velicina
Public d_CAD As Double 'CAD geometrijska velicina
Public t2_CAD As Double 'CAD geometrijska velicina

Private Sub Combo4_Change()

P1 = Val(Text2.Text)
n1 = Val(Text3.Text)
i = Val(Text4.Text)
Ka = Val(Text1.Text)
Pm = Val(P1) * Val(Ka) 'Merodavna snaga

'----------------------------------------
'Resetovanje vrednosti (Frame: Rezultati)
'----------------------------------------
Label1.Caption = "0.000" 'Racunska vrednost gonjene remenice
Label52.Caption = "0.000" 'Stvarna vrednost prenosnog odnosa
Label36.Caption = "0.000" 'Maksimalna obimna brzina
Label79.Caption = "0.000" 'Racunska vrednost duzine kaisa
Label89.Caption = "0.000" 'Standardna vrednost duzine kaisa
Label58.Caption = "0.000" 'Stvarna vrednost osnog rastojanja
Label67.Caption = "0.000" 'Racunska vrednost broja zlebova
Label73.Caption = "0.000" 'Broj zlebova
Label33.Caption = "0.000" 'Obimna brzina
Label33.BackColor = &H8000000F 'Siva boja
Label42.Caption = "0.000" 'Ucestanost savijanja
Label42.BackColor = &H8000000F 'Siva boja
Label73.BackColor = &H8000000F 'Siva boja
Text6.Text = "" 'Osno rastojanje
Label75.Visible = False ' Sakrivanje Label-e za prikaz amin i amax
Command1.Visible = False 'Sakrivanje dugmeta za izbor jaceg profila

'-------------------------------------------
'Dodavanje tipova kaiseva na osnovu Pm i n1
'-------------------------------------------

If Combo4.Text = "Uski" Then '1
    Combo5.Clear 'Reset vrednosti
    Combo6.Clear 'Reset vrednosti
    Text5.Text = "" 'Reset vrednosti
    Label75.Visible = False ' Sakrivanje Label-e za prikaz amin i amax
    fsdoz = 100 'Dozvoljena ucestanost savijanja za uski profil
    Label45.Caption = fsdoz
    vmax = 42 'Maksimalna obimna brzina remenice uski profil
    Label36.Caption = vmax

    If Val(Pm) >= 2 And Val(Pm) <= 31.5 And _
        Val(n1) >= (-17.4751 + 44.5847 * Val(Pm) + 1.46107 * Val(Pm) ^ 2 - 0.0159782 * Val(Pm) ^ 3) _
        And Val(n1) >= (-3843.15 + 212.481 * Val(Pm)) And Val(n1) >= 200 And Val(n1) <= 2850 Then '1 SPZ
    
        Combo5.Clear
        Combo5.AddItem "SPZ - DIN 7753: 1988"
        Combo5.Text = "SPZ - DIN 7753: 1988"

        ElseIf Val(Pm) >= 4.35 And Val(Pm) <= 90 And Val(n1) < -17.4751 + 44.5847 * Val(Pm) + 1.46107 * Val(Pm) ^ 2 - 0.0159782 * Val(Pm) ^ 3 And Val(n1) >= -3843.15 + 212.481 * Val(Pm) Or _
        Val(n1) <= -3843.15 + 212.481 * Val(Pm) And Val(n1) >= -32.3734983808947 + 17.022198599342 * Val(Pm) + 0.160417921299349 * Val(Pm) ^ 2 - 5.9367958824142E-04 * Val(Pm) ^ 3 And _
        Val(n1) >= -3325.3846153846 + 68.6153846153845 * Val(Pm) And Val(n1) >= 200 And Val(n1) <= 2850 Then 'SPA
    
        Combo5.Clear
        Combo5.AddItem "SPA - DIN 7753: 1988"
        Combo5.Text = "SPA - DIN 7753: 1988"

        ElseIf Val(Pm) >= 12.3 And Val(Pm) <= 360 And Val(n1) < -32.3734983808947 + 17.0221985993427 * Val(Pm) + 0.160417921299349 * Val(Pm) ^ 2 - 5.936795882414E-04 * Val(Pm) ^ 3 And _
        Val(n1) >= -3325.3846153846 + 68.6153846153845 * Val(Pm) Or Val(n1) <= -3325.3846153846 + 68.6153846153845 * Val(Pm) And _
        Val(n1) >= -82.7628161884858 + 5.65586132849337 * Val(Pm) + 0.002359679550198 * Val(Pm) ^ 2 + 2.118496130732E-05 * Val(Pm) ^ 3 + (2.43715833364965 * Val(Pm) ^ 4) / 10 ^ 8 - (1.33639630116493 * Val(Pm) ^ 5) / 10 ^ 10 And _
        Val(n1) >= 200 And Val(n1) <= 2850 Then '3 SPB
    
        Combo5.Clear
        Combo5.AddItem "SPB - DIN 7753: 1988"
        Combo5.Text = "SPB - DIN 7753: 1988"

        ElseIf Val(Pm) >= 48 And Val(Pm) <= 400 And Val(n1) < -82.7628161884858 + 5.65586132849337 * Val(Pm) + 0.002359679550198 * Val(Pm) ^ 2 + 2.118496130732E-05 * Val(Pm) ^ 3 + (2.43715833364965 * Val(Pm) ^ 4) / 10 ^ 8 - (1.33639630116493 * Val(Pm) ^ 5) / 10 ^ 10 And _
        Val(n1) >= 200 And Val(n1) <= 2850 Then '4 SPC
    
        Combo5.Clear
        Combo5.AddItem "SPC - DIN 7753: 1988"
        Combo5.Text = "SPC - DIN 7753: 1988"

    End If '1 SPZ

Else:
    Combo5.Clear 'Reset vrednosti
    Combo6.Clear 'Reset vrednosti
    Text5.Text = "" 'Reset vrednosti
    Label75.Visible = False ' Sakrivanje Label-e za prikaz amin i amax
    fsdoz = 80 'Dozvoljena ucestanost savijanja za profil normalne sirine
    Label45.Caption = fsdoz
    vmax = 30 'Maksimalna obimna brzina remenice za profil normalne sirine
    Label36.Caption = vmax
    
    If Val(Pm) >= 2 And Val(Pm) <= 7.5 And Val(n1) >= (-150 + 400 * Val(Pm)) And Val(n1) >= 650 And Val(n1) <= 2850 Then '10/Z
    
    Combo5.Clear
    Combo5.AddItem "Z/10 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.Text = "Z/10 - ISO 4184: 1992 / DIN 2215: 1998"
        
    ElseIf Val(Pm) >= 2 And Val(Pm) <= 21.5 And Val(n1) < (-150 + 400 * Val(Pm)) And _
    Val(n1) >= (-71.7948717948716 + 135.897435897436 * Val(Pm)) And Val(n1) >= 200 And Val(n1) <= 2850 Then '13/A
    
    Combo5.Clear
    Combo5.AddItem "A/13 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.Text = "A/13 - ISO 4184: 1992 / DIN 2215: 1998"
    
    ElseIf Val(Pm) >= 2 And Val(Pm) <= 400 And Val(n1) < (-71.7948717948716 + 135.897435897436 * Val(Pm)) And _
    Val(n1) >= (-58.8372093023254 + 41.0852713178295 * Val(Pm)) And Val(n1) >= 200 And _
    Val(n1) <= 2850 Or (Val(Pm) >= 2 And Val(Pm) <= 400 And Val(n1) < (-71.7948717948716 + 135.897435897436 * Val(Pm)) And _
    Val(n1) >= (-58.8372093023254 + 41.0852713178295 * Val(Pm)) Or _
    Val(n1) <= (-58.8372093023254 + 41.0852713178295 * Val(Pm)) And Val(n1) >= 1790 And Val(n1) <= 2850) Then '17/B
    
    Combo5.Clear
    Combo5.AddItem "B/17 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998"
        
    ElseIf Val(Pm) >= 6.3 And Val(Pm) <= 400 And Val(n1) < (-58.8372093023254 + 41.0852713178295 * Val(Pm)) And _
    Val(n1) >= (-60.5714285714287 + 13.0285714285714 * Val(Pm)) And Val(n1) <= 1112 Or Val(n1) >= 1112 And Val(n1) < 1790 Then '22/C
    
    Combo5.Clear
    Combo5.AddItem "C/22 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998"
        
End If '10/Z
End If '1
End Sub

Private Sub Combo4_Click()
Call Combo4_Change
End Sub

Private Sub Combo5_Change()
'---------------------
'Resetovanje vrednosti
'---------------------
Label79.Caption = "0.000" 'Racunska vrednost dužine kaiša
Label89.Caption = "0.000" 'Standardna vrednost dužine kaiša
Label58.Caption = "0.000" 'Stvarna vrednost osnog rastojanja
Label67.Caption = "0.000" 'Racunska vrednost broja žlebova
Label73.Caption = "0.000" 'Broj žlebova
Label33.Caption = "0.000" 'Obimna brzina
Label33.BackColor = &H8000000F 'Siva boja
Label42.Caption = "0.000" 'Ucestanost savijanja:
Label42.BackColor = &H8000000F 'Siva boja
Label73.BackColor = &H8000000F 'Siva boja
Text5.Text = "" 'Precnik gonjene remenice:
Text6.Text = "" 'Osno rastojanje
Label75.Visible = False ' Sakrivanje Label-e za prikaz amin i amax

'---------------------------------------------------------------
'Dodavanje vrednosti precnika d1 na osnovu izabranog tipa kaisa
'---------------------------------------------------------------

If Combo4.Text = "Uski" Then '11

    If Combo5.Text = "SPZ - DIN 7753: 1988" Then '1 SPZ
    Combo6.Clear
    Combo6.AddItem "63"
    Combo6.AddItem "80"
    Combo6.AddItem "90"
    Combo6.AddItem "112"
    Combo6.AddItem "140"
    Combo6.AddItem "180"
    
    ElseIf Combo5.Text = "SPA - DIN 7753: 1988" Then '2 SPA
    Combo6.Clear
    Combo6.AddItem "90"
    Combo6.AddItem "112"
    Combo6.AddItem "125"
    Combo6.AddItem "160"
    Combo6.AddItem "200"
    Combo6.AddItem "250"
    
    ElseIf Combo5.Text = "SPB - DIN 7753: 1988" And Val(n1) <= 2570 Then '3 SPB
    Combo6.Clear
    Combo6.AddItem "140"
    Combo6.AddItem "180"
    Combo6.AddItem "200"
    Combo6.AddItem "250"
    Combo6.AddItem "315"
    
    ElseIf Combo5.Text = "SPB - DIN 7753: 1988" And Val(n1) <= 2030 Then '4 SPB
    Combo6.Clear
    Combo6.AddItem "140"
    Combo6.AddItem "180"
    Combo6.AddItem "200"
    Combo6.AddItem "250"
    Combo6.AddItem "315"
    Combo6.AddItem "400"
    
    ElseIf Combo5.Text = "SPB - DIN 7753: 1988" Then '5 SPB
    Combo6.Clear
    Combo6.AddItem "140"
    Combo6.AddItem "180"
    Combo6.AddItem "200"
    Combo6.AddItem "250"
    
    ElseIf Combo5.Text = "SPC - DIN 7753: 1988" And Val(n1) <= 1285 Then '6 SPC
    Combo6.Clear
    Combo6.AddItem "224"
    Combo6.AddItem "280"
    Combo6.AddItem "315"
    Combo6.AddItem "400"
    Combo6.AddItem "500"
    Combo6.AddItem "630"

    ElseIf Combo5.Text = "SPC - DIN 7753: 1988" And Val(n1) <= 1450 Then '7 SPC
    Combo6.Clear
    Combo6.AddItem "224"
    Combo6.AddItem "280"
    Combo6.AddItem "315"
    Combo6.AddItem "400"
    Combo6.AddItem "500"
    
    ElseIf Combo5.Text = "SPC - DIN 7753: 1988" And Val(n1) <= 1740 Then '8 SPC
    Combo6.Clear
    Combo6.AddItem "224"
    Combo6.AddItem "280"
    Combo6.AddItem "315"
    Combo6.AddItem "400"
    
    ElseIf Combo5.Text = "SPC - DIN 7753: 1988" And Val(n1) <= 2000 Then '9 SPC
    Combo6.Clear
    Combo6.AddItem "224"
    Combo6.AddItem "280"
    Combo6.AddItem "315"

    ElseIf Combo5.Text = "SPC - DIN 7753: 1988" And Val(n1) <= 2700 Then '10 SPC
    Combo6.Clear
    Combo6.AddItem "224"
    Combo6.AddItem "280"
    
    ElseIf Combo5.Text = "SPC - DIN 7753: 1988" And Val(n1) <= 2850 Then '11 SPC
    Combo6.Clear
    Combo6.AddItem "224"
    
    End If '1 SPZ

Else:

    If Combo5.Text = "Z/10 - ISO 4184: 1992 / DIN 2215: 1998" Then '1 10/Z
    Combo6.Clear
    Combo6.AddItem "50"
    Combo6.AddItem "63"
    Combo6.AddItem "80"
    Combo6.AddItem "100"
    Combo6.AddItem "125"
    
    ElseIf Combo5.Text = "A/13 - ISO 4184: 1992 / DIN 2215: 1998" Then '2 13/A
    Combo6.Clear
    Combo6.AddItem "80"
    Combo6.AddItem "100"
    Combo6.AddItem "125"
    Combo6.AddItem "160"
    Combo6.AddItem "200"
    
    ElseIf Combo5.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998" And Val(n1) <= 2000 Then '3 17/B
    Combo6.Clear
    Combo6.AddItem "125"
    Combo6.AddItem "160"
    Combo6.AddItem "200"
    Combo6.AddItem "250"
    Combo6.AddItem "315"
    
    ElseIf Combo5.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998" And Val(n1) <= 2360 Then '4 17/B
    Combo6.Clear
    Combo6.AddItem "125"
    Combo6.AddItem "160"
    Combo6.AddItem "200"
    Combo6.AddItem "250"
    
    ElseIf Combo5.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998" And Val(n1) <= 2850 Then '5 17/B
    Combo6.Clear
    Combo6.AddItem "125"
    Combo6.AddItem "160"
    Combo6.AddItem "200"
    
    ElseIf Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" And Val(n1) <= 1200 Then '6 22/C
    Combo6.Clear
    Combo6.AddItem "200"
    Combo6.AddItem "250"
    Combo6.AddItem "315"
    Combo6.AddItem "400"
    Combo6.AddItem "500"
    
    ElseIf Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" And Val(n1) <= 1630 Then '7 22/C
    Combo6.Clear
    Combo6.AddItem "200"
    Combo6.AddItem "250"
    Combo6.AddItem "315"
    Combo6.AddItem "400"
    
    ElseIf Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" And Val(n1) <= 2000 Then '8 22/C
    Combo6.Clear
    Combo6.AddItem "200"
    Combo6.AddItem "250"
    Combo6.AddItem "315"
    
    ElseIf Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" And Val(n1) <= 2330 Then '9 22/C
    Combo6.Clear
    Combo6.AddItem "200"
    Combo6.AddItem "250"
    
    ElseIf Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" And Val(n1) <= 2850 Then '10 22/C
    Combo6.Clear
    Combo6.AddItem "200"
    
End If '1 10/Z
End If '11
End Sub

Private Sub Combo5_Click()
Call Combo5_Change
End Sub

Private Sub Combo6_Change()

'---------------------
'Resetovanje vrednosti
'---------------------
Label79.Caption = "0.000" 'Racunska vrednost dužine kaiša
Label89.Caption = "0.000" 'Standardna vrednost dužine kaiša
Label58.Caption = "0.000" 'Stvarna vrednost osnog rastojanja
Label67.Caption = "0.000" 'Racunska vrednost broja žlebova
Label73.Caption = "0.000" 'Broj žlebova
Label33.Caption = "0.000" 'Obimna brzina
Label33.BackColor = &H8000000F 'Siva boja
Label42.Caption = "0.000" 'Ucestanost savijanja
Label42.BackColor = &H8000000F 'Siva boja
Label73.BackColor = &H8000000F 'Siva boja
Text6.Text = "" 'Osno rastojanje
Label75.Visible = True ' Sakrivanje Label-e za prikaz amin i amax
'---------------------

d1 = Val(Combo6.Text) 'Precnik pogonske remenice
d2r = d1 * i 'Racunska vrednost gonjene remenice
Label1.Caption = Round(d2r, 3) 'Racunska vrednost gonjene remenice

'------------------------------------------------
'Izbor najpribliznije vrednosti gonjene remenice
'------------------------------------------------

Open App.Path & "/data/Precnici_remenica.data" For Input As #11
Do
    Input #11, ds 'ds je standardni precnik iz baze (R20 red)
    If ds > d2r Then
        dmin = dmin
        dmax = ds
    ElseIf ds < d2r Then
        dmin = ds
        dmax = d2r
    ElseIf ds = d2r Then
        dmin = ds
        dmax = ds
        Exit Do
    End If
        If ds > d2r Then
        Exit Do
        End If
Loop Until EOF(11)
Close #11

'---------
'PRORACUN
'---------

ds1 = Abs(dmax - d2r) 'Prva razlika
ds2 = Abs(dmin - d2r) 'Druga razlika
If ds1 <= ds2 Then 'Izbor pribliznije vrednosti za d2
    d2 = dmax
Else: d2 = dmin
End If

'---------------------
'If d2 > 2500 Then
'MsgBox "Nepostoji standardna vrednost precnika gonjene remenice za unete vrednosti ", vbInformation, "Nepostojaca kombinacija"
'Text5.Text = ""
'End If
'---------------------

Text5.Text = d2
iss = d2 / d1 ' Stvarna vrednost prenosnog odnosa
Label52.Caption = Round(iss, 3)
amin = 0.7 * (d1 + d2) 'Minimalna vrednost osnog rastojanja
amax = 2 * (d1 + d2) 'Maksimalna vrednost osnog rastojanja
Label75.Caption = "amin= " & amin & ", amax= " & amax

End Sub

Private Sub Combo6_Click()
Call Combo6_Change
End Sub

Private Sub Combo7_Change()
If Combo7.Text = "Uski" Then
    Combo8.Clear
    Combo8.AddItem "SPZ - DIN 7753: 1988"
    Combo8.AddItem "SPA - DIN 7753: 1988"
    Combo8.AddItem "SPB - DIN 7753: 1988"
    Combo8.AddItem "SPC - DIN 7753: 1988"
Else:
    Combo8.Clear
    Combo8.AddItem "Z/10 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo8.AddItem "A/13 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo8.AddItem "B/17 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo8.AddItem "C/22 - ISO 4184: 1992 / DIN 2215: 1998"
End If
End Sub

Private Sub Combo7_Click()
Call Combo7_Change
End Sub

Private Sub Combo8_Change()

'---------------------
'Baza vrednosti za CAD
'---------------------

If Combo8.Text = "SPZ - DIN 7753: 1988" Then '1
    b1_CAD = 9.7
    c_CAD = 2
    E_CAD = 12
    f_CAD = 8
    t_CAD = 11
        If d1 <= 80 Then '1.1
        alfa_CAD = 34
        Else: alfa_CAD = 38
        End If '1.1
    b0_CAD = 9.7
    bd_CAD = 8.5
    h_CAD = 8
    hd_CAD = 2
        
    ElseIf Combo8.Text = "SPA - DIN 7753: 1988" Then '2
        b1_CAD = 12.7
        c_CAD = 2.8
        E_CAD = 15
        f_CAD = 10
        t_CAD = 14
            If d1 <= 118 Then '2.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '2.1
        b0_CAD = 12.7
        bd_CAD = 11
        h_CAD = 10
        hd_CAD = 2.8

    ElseIf Combo8.Text = "SPB - DIN 7753: 1988" Then '3
        b1_CAD = 16.3
        c_CAD = 3.5
        E_CAD = 19
        f_CAD = 12.5
        t_CAD = 18
            If d1 <= 190 Then '3.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '3.1
        b0_CAD = 16.3
        bd_CAD = 14
        h_CAD = 13
        hd_CAD = 3.5
        
    ElseIf Combo8.Text = "SPC - DIN 7753: 1988" Then '4
        b1_CAD = 22
        c_CAD = 4.6
        E_CAD = 25.5
        f_CAD = 17
        t_CAD = 24
            If d1 <= 315 Then '4.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '4.1
        b0_CAD = 22
        bd_CAD = 19
        h_CAD = 18
        hd_CAD = 4.8

    ElseIf Combo8.Text = "Z/10 - ISO 4184: 1992 / DIN 2215: 1998" Then '5
        b1_CAD = 9.7
        c_CAD = 2
        E_CAD = 12
        f_CAD = 8
        t_CAD = 11
            If d1 <= 80 Then '5.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '5.1
        b0_CAD = 10
        bd_CAD = 8.5
        h_CAD = 6
        hd_CAD = 2.5
        
    ElseIf Combo8.Text = "A/13 - ISO 4184: 1992 / DIN 2215: 1998" Then '6
        b1_CAD = 12.7
        c_CAD = 2.8
        E_CAD = 15
        f_CAD = 10
        t_CAD = 14
            If d1 <= 118 Then '6.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '6.1
        b0_CAD = 13
        bd_CAD = 11
        h_CAD = 8
        hd_CAD = 3.3
        
    ElseIf Combo8.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998" Then '7
        b1_CAD = 16.3
        c_CAD = 3.5
        E_CAD = 19
        f_CAD = 12.5
        t_CAD = 18
            If d1 <= 190 Then '7.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '7.1
        b0_CAD = 17
        bd_CAD = 14
        h_CAD = 11
        hd_CAD = 4.2
        
    ElseIf Combo8.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" Then '8
        b1_CAD = 22
        c_CAD = 4.6
        E_CAD = 25.5
        f_CAD = 17
        t_CAD = 24
            If d1 <= 315 Then '8.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '8.1
        b0_CAD = 22
        bd_CAD = 19
        h_CAD = 14
        hd_CAD = 5.7

End If '1

'STAMPANJE VREDNOSTI
Label43(18).Caption = b1_CAD
Label43(22).Caption = c_CAD
Label43(26).Caption = E_CAD
Label43(29).Caption = f_CAD
Label43(33).Caption = t_CAD
Label43(37).Caption = alfa_CAD

Label43(3).Caption = b0_CAD
Label43(7).Caption = bd_CAD
Label43(11).Caption = h_CAD
Label43(15).Caption = hd_CAD

End Sub

Private Sub Combo8_Click()
Call Combo8_Change
End Sub

Private Sub Command1_Click()
'---------------------
'Izbor jaceg profila
'---------------------

Combo6.Clear 'Reset vrednosti
Label1.Caption = "0.000" 'Racunska vrednost gonjene remenice
Label52.Caption = "0.000" 'Stvarna vrednost prenosnog odnosa
Text5.Text = "" 'Precnik gonjene remenice
Text6.Text = "" 'Osno rastojanje

If Combo4.Text = "Uski" Then '1

If Val(Pm) >= 2 And Val(Pm) <= 31.5 And _
    Val(n1) >= (-17.4751 + 44.5847 * Val(Pm) + 1.46107 * Val(Pm) ^ 2 - 0.0159782 * Val(Pm) ^ 3) _
    And Val(n1) >= (-3843.15 + 212.481 * Val(Pm)) And Val(n1) >= 200 And Val(n1) <= 2850 Then '1 SPZ
    
    Combo5.Clear
    Combo5.AddItem "SPZ - DIN 7753: 1988"
    Combo5.AddItem "SPA - DIN 7753: 1988"
    Combo5.Text = "SPA - DIN 7753: 1988"
    Combo5.AddItem "SPB - DIN 7753: 1988"
    Combo5.AddItem "SPC - DIN 7753: 1988"

ElseIf Val(Pm) >= 4.35 And Val(Pm) <= 90 And Val(n1) < -17.4751 + 44.5847 * Val(Pm) + 1.46107 * Val(Pm) ^ 2 - 0.0159782 * Val(Pm) ^ 3 And Val(n1) >= -3843.15 + 212.481 * Val(Pm) Or _
    Val(n1) <= -3843.15 + 212.481 * Val(Pm) And Val(n1) >= -32.3734983808947 + 17.022198599342 * Val(Pm) + 0.160417921299349 * Val(Pm) ^ 2 - 5.9367958824142E-04 * Val(Pm) ^ 3 And _
    Val(n1) >= -3325.3846153846 + 68.6153846153845 * Val(Pm) And Val(n1) >= 200 And Val(n1) <= 2850 Then 'SPA
    
    Combo5.Clear
    Combo5.AddItem "SPA - DIN 7753: 1988"
    Combo5.AddItem "SPB - DIN 7753: 1988"
    Combo5.Text = "SPB - DIN 7753: 1988"
    Combo5.AddItem "SPC - DIN 7753: 1988"

ElseIf Val(Pm) >= 12.3 And Val(Pm) <= 360 And Val(n1) < -32.3734983808947 + 17.0221985993427 * Val(Pm) + 0.160417921299349 * Val(Pm) ^ 2 - 5.936795882414E-04 * Val(Pm) ^ 3 And _
    Val(n1) >= -3325.3846153846 + 68.6153846153845 * Val(Pm) Or Val(n1) <= -3325.3846153846 + 68.6153846153845 * Val(Pm) And _
    Val(n1) >= -82.7628161884858 + 5.65586132849337 * Val(Pm) + 0.002359679550198 * Val(Pm) ^ 2 + 2.118496130732E-05 * Val(Pm) ^ 3 + (2.43715833364965 * Val(Pm) ^ 4) / 10 ^ 8 - (1.33639630116493 * Val(Pm) ^ 5) / 10 ^ 10 And _
    Val(n1) >= 200 And Val(n1) <= 2850 Then '3 SPB
    
    Combo5.Clear
    Combo5.AddItem "SPB - DIN 7753: 1988"
    Combo5.AddItem "SPC - DIN 7753: 1988"
    Combo5.Text = "SPC - DIN 7753: 1988"

End If '1 SPZ

Else: 'Combo4.Text="Normalni"

    If Val(Pm) >= 2 And Val(Pm) <= 7.5 And Val(n1) >= (-150 + 400 * Val(Pm)) And Val(n1) >= 650 And Val(n1) <= 2850 Then '10/Z
    
    Combo5.Clear
    Combo5.AddItem "Z/10 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.AddItem "A/13 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.Text = "A/13 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.AddItem "B/17 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.AddItem "C/22 - ISO 4184: 1992 / DIN 2215: 1998"
      
    ElseIf Val(Pm) >= 2 And Val(Pm) <= 21.5 And Val(n1) < (-150 + 400 * Val(Pm)) And _
    Val(n1) >= (-71.7948717948716 + 135.897435897436 * Val(Pm)) And Val(n1) >= 200 And Val(n1) <= 2850 Then '13/A
    
    Combo5.Clear
    Combo5.AddItem "A/13 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.AddItem "B/17 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.AddItem "C/22 - ISO 4184: 1992 / DIN 2215: 1998"

    
    ElseIf Val(Pm) >= 2 And Val(Pm) <= 400 And Val(n1) < (-71.7948717948716 + 135.897435897436 * Val(Pm)) And _
    Val(n1) >= (-58.8372093023254 + 41.0852713178295 * Val(Pm)) And Val(n1) >= 200 And _
    Val(n1) <= 2850 Or (Val(Pm) >= 2 And Val(Pm) <= 400 And Val(n1) < (-71.7948717948716 + 135.897435897436 * Val(Pm)) And _
    Val(n1) >= (-58.8372093023254 + 41.0852713178295 * Val(Pm)) Or _
    Val(n1) <= (-58.8372093023254 + 41.0852713178295 * Val(Pm)) And Val(n1) >= 1790 And Val(n1) <= 2850) Then '17/B
    
    Combo5.Clear
    Combo5.AddItem "B/17 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.AddItem "C/22 - ISO 4184: 1992 / DIN 2215: 1998"
    Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998"
     
End If '10/Z
End If '1
End Sub

Private Sub Command2_Click()
Call Command7_Click 'Napustanje programa
End Sub

Private Sub Command3_Click()
'-----------------
'Provera vrednosti
'-----------------
'Mnozenje sa 0.001 je prebacivanje u metre

If Option1.Value = True Then '1 Predhodno proracunata remenica
    
    If Option3.Value = True Then '2
        d_CAD = 0.001 * Text7.Text / 2 'Precnik remenice d1
        f_CAD = 0.001 * Label43(29).Caption 'Geom. velicina
        c_CAD = 0.001 * Label43(22).Caption 'Geom. velicina
        b1_CAD = 0.001 * Label43(18).Caption 'Geom. velicina
        alfa_CAD = Label43(37).Caption 'Geom. velicina
        t_CAD = 0.001 * Label43(33).Caption 'Geom. velicina
        t2_CAD = t_CAD * Tan(alfa_CAD / 2 * Pi / 180) 'Geom. velicina
        E_CAD = 0.001 * Label43(26).Caption 'Geom. velicina
        z = Label73.Caption 'Broj zlebova
        Form1.WindowState = vbMinimized 'Minimiziranje prozora
        Call CAD
        
    ElseIf Option4.Value = True Then
        d_CAD = 0.001 * Text8.Text / 2 'Precnik remenice d2
        f_CAD = 0.001 * Label43(29).Caption 'Geom. velicina
        c_CAD = 0.001 * Label43(22).Caption 'Geom. velicina
        b1_CAD = 0.001 * Label43(18).Caption 'Geom. velicina
        alfa_CAD = Label43(37).Caption 'Geom. velicina
        t_CAD = 0.001 * Label43(33).Caption 'Geom. velicina
        t2_CAD = t_CAD * Tan(alfa_CAD / 2 * Pi / 180) 'Geom. velicina
        E_CAD = 0.001 * Label43(26).Caption 'Geom. velicina
        z = Label73.Caption 'Broj zlebova
        Form1.WindowState = vbMinimized 'Minimiziranje prozora
        Call CAD
        
    End If
End If '1

If Option2.Value = True Then '1 Nezavisno modeliranje remenice
    
    If Combo7.Text = "" Then '2 'Profil kaisa
    MsgBox "Niste odabrali profil kaiša", vbInformation, "Neodgovarajuci unos"
    Combo7.SetFocus
    
    ElseIf Combo8.Text = "" Then 'Tip kaisa
    MsgBox "Niste odabrali tip kaiša", vbInformation, "Neodgovarajuci unos"
    Combo8.SetFocus
    
    ElseIf Text7.Text = "" Then 'Precnik remenice
    MsgBox "Niste uneli precnik remenice", vbInformation, "Neodgovarajuci unos"
    Text7.SetFocus
    
    ElseIf Text7.Text < 50 Then 'Precnik remenice d<50
    MsgBox "Precnik remenice mora biti veci od 50 mm", vbInformation, "Neodgovarajuci unos"
    Text7.Text = ""
    Text7.SetFocus
    
    ElseIf Text7.Text > 10000 Then 'Precnik remenice d>10000
    MsgBox "Precnik remenice mora biti manji od 10000 mm", vbInformation, "Neodgovarajuci unos"
    Text7.Text = ""
    Text7.SetFocus
    
    ElseIf Text9.Text = "" Or Text9.Text = "0" Then
    MsgBox "Niste uneli broj žlebova", vbInformation, "Neodgovarajuci unos"
    Text9.SetFocus
    Text9.Text = ""
    
    ElseIf Text9.Text > 12 Then
    MsgBox "Maksimalni broj žlebova je 12", vbInformation, "Neodgovarajuci unos"
    Text9.SetFocus
    Text9.Text = ""
    
    Else:
    
    d_CAD = 0.001 * Text7.Text / 2 'Precnik remenice d1
    f_CAD = 0.001 * Label43(29).Caption 'Geom. velicina
    c_CAD = 0.001 * Label43(22).Caption 'Geom. velicina
    b1_CAD = 0.001 * Label43(18).Caption 'Geom. velicina
    alfa_CAD = Label43(37).Caption 'Geom. velicina
    t_CAD = 0.001 * Label43(33).Caption 'Geom. velicina
    t2_CAD = t_CAD * Tan(alfa_CAD / 2 * Pi / 180) 'Geom. velicina
    E_CAD = 0.001 * Label43(26).Caption 'Geom. velicina
    z = Text9.Text 'Broj zlebova
    Form1.WindowState = vbMinimized 'Minimiziranje prozora
    Call CAD
    
    End If '2

End If '1 Nezavisno modeliranje remenice

End Sub

Public Sub CAD()
'----------------------------------------
'Modeliranje remenice u SolidWorks-u 2015
'----------------------------------------

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSkMgr As SldWorks.SketchManager
Dim swSketch As SldWorks.Sketch
Dim swSketchSegment As SldWorks.SketchSegment
Dim swSelMgr As SldWorks.SelectionMgr
Dim swSelData As SldWorks.SelectData
Dim myFeature As SldWorks.Feature
Dim bRet As Boolean
Dim boolstatus As Boolean
Dim longstatus As Long

Set swApp = New SldWorks.SldWorks
swApp.Visible = True
Set swModel = swApp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2015\templates\Part1.prtdot", 0, 0, 0) 'Lokacija sablona!!!
swApp.ActivateDoc2 "Part1", False, longstatus 'Otvaranje novog dokumenta
Set swModel = swApp.ActiveDoc

'Sistem jedinica: Millimeter, gram, second
swModel.Extension.SetUserPreferenceInteger swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swUnitSystem_e.swUnitSystem_MMGS

'-----------------------------------
'Podesavanje "snepovanja" - Snapping - nakon modeliranja
'-----------------------------------
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchInference, True 'Enable snapping
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchInferFromModel, False 'Snap to model geometry
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchAutomaticRelations, False 'Automatic relations
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsPoints, True 'End points and sketch points
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsCenterPoints, False ' Center points
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsMidPoints, False 'Mid-points
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, False 'Quadrant Points
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsIntersections, False 'Intersections
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsNearest, True 'Nearest
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsTangent, False 'Tanget
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsPerpendicular, False 'Perpendicular
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsParallel, False 'Parallel
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsHVLines, False 'Horizontal/vertical lines
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsHVPoints, False 'Horizontal/vertical points
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsLength, False 'Length
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsAngle, False 'Angle

'Selektovanje frontalne ravni
boolstatus = swModel.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, False, 0, Nothing, 0)

'Otvaranje nove skice
swModel.SketchManager.InsertSketch True

'----------------
'Dinamicki mirror
'----------------
Set swSelMgr = swModel.SelectionManager
Set swSelData = swSelMgr.CreateSelectData
Set swSketch = swModel.GetActiveSketch2
Set swSketchSegment = swModel.SketchManager.CreateCenterLine(0#, 0#, 0#, 0#, d_CAD + c_CAD - 2 * t_CAD, 0#) 'Linija oko koje se vrsi dinamicko preslikavanje u ogledalu (Dynamically Mirror Sketch Entities)
bRet = swSketchSegment.Select4(True, swSelData)
Set swSkMgr = swModel.SketchManager
swSkMgr.SetDynamicMirror (True)

'---------------
'Crtanje profila
'---------------
Select Case z

    Case 1
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD, 0, 0, -f_CAD, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.5 * (-f_CAD), d_CAD + c_CAD, 0, 1.5 * f_CAD, d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD, d_CAD + c_CAD, 0, -(b1_CAD / 2), d_CAD + c_CAD, 0) 'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-(b1_CAD / 2), d_CAD + c_CAD, 0, -(b1_CAD / 2 - t2_CAD), d_CAD + c_CAD - t_CAD, 0)  'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-(b1_CAD / 2 - t2_CAD), d_CAD + c_CAD - t_CAD, 0, 0, d_CAD + c_CAD - t_CAD, 0) 'Linija 6

    Case 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - E_CAD / 2, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - E_CAD / 2, 0, 0, -f_CAD - E_CAD / 2, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - E_CAD / 2), d_CAD + c_CAD, 0, 1.3 * (f_CAD + E_CAD / 2), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - E_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2 - E_CAD / 2, d_CAD + c_CAD, 0) 'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2 - E_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2 - E_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2 - E_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -t2_CAD - (E_CAD / 2 - b1_CAD / 2), d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-t2_CAD - (E_CAD / 2 - b1_CAD / 2), d_CAD + c_CAD - t_CAD, 0, -(E_CAD / 2 - b1_CAD / 2), d_CAD + c_CAD, 0) 'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, 0, d_CAD + c_CAD, 0) 'Linija 8
    
    Case 3
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - E_CAD, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - E_CAD, 0, 0, -f_CAD - E_CAD, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - E_CAD), d_CAD + c_CAD, 0, 1.3 * (f_CAD + E_CAD), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - E_CAD, d_CAD + c_CAD, 0, -E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, 0, d_CAD + c_CAD - t_CAD, 0) 'Linija 10
    
    Case 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - 15 * E_CAD / 10, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 15 * E_CAD / 10, 0, 0, -f_CAD - 15 * E_CAD / 10, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - 15 * E_CAD / 10), d_CAD + c_CAD, 0, 1.3 * (f_CAD + 15 * E_CAD / 10), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 15 * E_CAD / 10, d_CAD + c_CAD, 0, -15 * E_CAD / 10 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 - b1_CAD / 2, d_CAD + c_CAD, 0, -15 * E_CAD / 10 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -15 * E_CAD / 10 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -15 * E_CAD / 10 + b1_CAD / 2, d_CAD + c_CAD, 0)  'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 + b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 10
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 11
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, 0, d_CAD + c_CAD, 0) 'Linija 11
    
    Case 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - 2 * E_CAD, 0, 0)  'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 2 * E_CAD, 0, 0, -f_CAD - 2 * E_CAD, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - 2 * E_CAD), d_CAD + c_CAD, 0, 1.3 * (f_CAD + 2 * E_CAD), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 2 * E_CAD, d_CAD + c_CAD, 0, -2 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0)  'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -2 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -2 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -2 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 10
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 11
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 12
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 13
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, 0, d_CAD + c_CAD - t_CAD, 0) 'Linija 14
    
    Case 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - 5 * E_CAD / 2, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 5 * E_CAD / 2, 0, 0, -f_CAD - 5 * E_CAD / 2, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - 5 * E_CAD / 2), d_CAD + c_CAD, 0, 1.3 * (f_CAD + 5 * E_CAD / 2), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 5 * E_CAD / 2, d_CAD + c_CAD, 0, -5 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -5 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -5 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -5 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, -15 * E_CAD / 10 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 - b1_CAD / 2, d_CAD + c_CAD, 0, -15 * E_CAD / 10 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -15 * E_CAD / 10 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 10
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -15 * E_CAD / 10 + b1_CAD / 2, d_CAD + c_CAD, 0)  'Linija 11
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 + b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 12
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 13
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 14
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 15
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, 0, d_CAD + c_CAD, 0) 'Linija 16

    Case 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - 3 * E_CAD, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 3 * E_CAD, 0, 0, -f_CAD - 3 * E_CAD, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - 3 * E_CAD), d_CAD + c_CAD, 0, 1.3 * (f_CAD + 3 * E_CAD), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 3 * E_CAD, d_CAD + c_CAD, 0, -3 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -3 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -3 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -3 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -2 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -2 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -2 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 10
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -2 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 11
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 12
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 13
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 14
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 15
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 16
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 17
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, 0, d_CAD + c_CAD - t_CAD, 0) 'Linija 18

    Case 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - 7 * E_CAD / 2, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 7 * E_CAD / 2, 0, 0, -f_CAD - 7 * E_CAD / 2, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - 7 * E_CAD / 2), d_CAD + c_CAD, 0, 1.3 * (f_CAD + 7 * E_CAD / 2), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 7 * E_CAD / 2, d_CAD + c_CAD, 0, -7 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0)  'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -7 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -7 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -7 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, -5 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -5 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -5 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 10 (Isto kao Linija 6 u slucaju z=6)
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -5 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 11
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, -15 * E_CAD / 10 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 12
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 - b1_CAD / 2, d_CAD + c_CAD, 0, -15 * E_CAD / 10 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 13
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -15 * E_CAD / 10 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 14
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -15 * E_CAD / 10 + b1_CAD / 2, d_CAD + c_CAD, 0)  'Linija 15
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 + b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 16
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 17
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 18
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 19
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, 0, d_CAD + c_CAD, 0) 'Linija 20

    Case 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - 4 * E_CAD, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 4 * E_CAD, 0, 0, -f_CAD - 4 * E_CAD, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - 4 * E_CAD), d_CAD + c_CAD, 0, 1.3 * (f_CAD + 4 * E_CAD), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 4 * E_CAD, d_CAD + c_CAD, 0, -4 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-4 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -4 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-4 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -4 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-4 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -4 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-4 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -3 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -3 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -3 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 10 (isto kao linija 6 za z=7)
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -3 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 11
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -2 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 12
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -2 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 13
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -2 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 14
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -2 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 15
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 16
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 17
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 18
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 19
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 20
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 21
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, 0, d_CAD + c_CAD - t_CAD, 0) 'Linija 22
    
    Case 10
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - 9 * E_CAD / 2, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 9 * E_CAD / 2, 0, 0, -f_CAD - 9 * E_CAD / 2, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - 9 * E_CAD / 2), d_CAD + c_CAD, 0, 1.3 * (f_CAD + 9 * E_CAD / 2), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 9 * E_CAD / 2, d_CAD + c_CAD, 0, -9 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-9 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -9 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-9 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -9 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-9 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -9 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-9 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, -7 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -7 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -7 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 10 (Isto kao Linija 6 u slucaju z=8)
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -7 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 11
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, -5 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 12
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -5 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 13
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -5 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 14
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -5 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 15
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, -15 * E_CAD / 10 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 16
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 - b1_CAD / 2, d_CAD + c_CAD, 0, -15 * E_CAD / 10 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 17
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -15 * E_CAD / 10 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 18
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -15 * E_CAD / 10 + b1_CAD / 2, d_CAD + c_CAD, 0)  'Linija 19
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 + b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 20
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 21
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 22
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 23
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, 0, d_CAD + c_CAD, 0) 'Linija 24
    
    Case 11
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - 5 * E_CAD, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 5 * E_CAD, 0, 0, -f_CAD - 5 * E_CAD, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - 5 * E_CAD), d_CAD + c_CAD, 0, 1.3 * (f_CAD + 5 * E_CAD), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 5 * E_CAD, d_CAD + c_CAD, 0, -5 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -5 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -5 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -5 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0)  'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -4 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(-4 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -4 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-4 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -4 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 10 (isto kao linija 6 za z=9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-4 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -4 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 11
    Set swSketchSegment = swModel.SketchManager.CreateLine(-4 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -3 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 12
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -3 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 13
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -3 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 14
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -3 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 15
    Set swSketchSegment = swModel.SketchManager.CreateLine(-3 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -2 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 16
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -2 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 17
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -2 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 18
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -2 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 19
    Set swSketchSegment = swModel.SketchManager.CreateLine(-2 * E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 20
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD - b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 21
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 22
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 23
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD + b1_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 24
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2, d_CAD + c_CAD, 0, -b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 25
    Set swSketchSegment = swModel.SketchManager.CreateLine(-b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, 0, d_CAD + c_CAD - t_CAD, 0) 'Linija 26
    
    Case 12
    Set swSketchSegment = swModel.SketchManager.CreateLine(0, 0, 0, -f_CAD - 11 * E_CAD / 2, 0, 0) 'Linija 2
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 11 * E_CAD / 2, 0, 0, -f_CAD - 11 * E_CAD / 2, d_CAD + c_CAD, 0) 'Linija 3
    swModel.ViewZoomTo2 1.3 * (-f_CAD - 11 * E_CAD / 2), d_CAD + c_CAD, 0, 1.3 * (f_CAD + 11 * E_CAD / 2), d_CAD + c_CAD - t_CAD, 0 'Zoom na deo gde se crta profil zlebova
    Set swSketchSegment = swModel.SketchManager.CreateLine(-f_CAD - 11 * E_CAD / 2, d_CAD + c_CAD, 0, -11 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 4
    Set swSketchSegment = swModel.SketchManager.CreateLine(-11 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -11 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 5
    Set swSketchSegment = swModel.SketchManager.CreateLine(-11 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -11 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 6
    Set swSketchSegment = swModel.SketchManager.CreateLine(-11 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -11 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0)  'Linija 7
    Set swSketchSegment = swModel.SketchManager.CreateLine(-11 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, -9 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 8
    Set swSketchSegment = swModel.SketchManager.CreateLine(-9 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -9 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 9
    Set swSketchSegment = swModel.SketchManager.CreateLine(-9 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -9 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) ' Linija 10 (isto kao linija 6 za z=10)
    Set swSketchSegment = swModel.SketchManager.CreateLine(-9 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -9 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 11
    Set swSketchSegment = swModel.SketchManager.CreateLine(-9 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, -7 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 12
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -7 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 13
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -7 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 14
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -7 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 15
    Set swSketchSegment = swModel.SketchManager.CreateLine(-7 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, -5 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 16
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -5 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 17
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -5 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 18
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -5 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 19
    Set swSketchSegment = swModel.SketchManager.CreateLine(-5 * E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, -15 * E_CAD / 10 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 20
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 - b1_CAD / 2, d_CAD + c_CAD, 0, -15 * E_CAD / 10 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 21
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -15 * E_CAD / 10 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 22
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -15 * E_CAD / 10 + b1_CAD / 2, d_CAD + c_CAD, 0)  'Linija 23
    Set swSketchSegment = swModel.SketchManager.CreateLine(-15 * E_CAD / 10 + b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 24
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 - b1_CAD / 2, d_CAD + c_CAD, 0, -E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 25
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 - b1_CAD / 2 + t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0) 'Linija 26
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2 - t2_CAD, d_CAD + c_CAD - t_CAD, 0, -E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0) 'Linija 27
    Set swSketchSegment = swModel.SketchManager.CreateLine(-E_CAD / 2 + b1_CAD / 2, d_CAD + c_CAD, 0, 0, d_CAD + c_CAD, 0) 'Linija 28
    
End Select

'-----------------------------------
'Podesavanje "snepovanja" - Snapping - nakon modeliranja
'-----------------------------------

swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchInference, True 'Enable snapping
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchInferFromModel, True 'Snap to model geometry
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchAutomaticRelations, True 'Automatic relations
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsPoints, True 'End points and sketch points
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsCenterPoints, True ' Center points
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsMidPoints, True 'Mid-points
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsQuadrantPoints, True 'Quadrant Points
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsIntersections, True 'Intersections
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsNearest, True 'Nearest
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsTangent, True 'Tanget
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsPerpendicular, True 'Perpendicular
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsParallel, True 'Parallel
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsHVLines, True 'Horizontal/vertical lines
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsHVPoints, True 'Horizontal/vertical points
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsLength, True 'Length
swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchSnapsAngle, True 'Angle

'Selektovanje Linije 1 (osa rotacije)
boolstatus = swModel.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 16, Nothing, 0)

'Uradi Revolve
Set myFeature = swModel.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 6.2831853071796, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True)

'Prikaz izometrije
swModel.ShowNamedView2 "*Isometric", 7
swModel.ViewZoomtofit2 'Fit

'Obavestenje
MsgBox "Modeliranje je zavrseno.", vbMsgBoxSetForeground

End Sub

Private Sub Command4_Click()

'--------------------------------
'Provera unetih ulaznih podataka
'--------------------------------

If Text2.Text > 400 Then '1
    MsgBox "Uneta nominalna snaga mora biti manja od 400 kW", vbInformation, "Neodgovarajuci unos"
    Text2.Text = ""
    Text2.SetFocus
    
    ElseIf Text2.Text < 2 Then '2
    MsgBox "Uneta snaga mora biti veca od 2 kW", vbInformation, "Neodgovarajuci unos"
    Text2.Text = ""
    Text2.SetFocus
    
    ElseIf Text3.Text < 200 Then '3
    MsgBox "Uneti broj obrtaja mora biti veci od 200 min -1", vbInformation, "Neodgovarajuci unos"
    Text3.Text = ""
    Text3.SetFocus
    
    ElseIf Text3.Text > 2850 Then '4
    MsgBox "Uneti broj obrtaja mora biti manji od 2850 min -1", vbInformation, "Neodgovarajuci unos"
    Text3.Text = ""
    Text3.SetFocus
    
    ElseIf Text4.Text > 15 And Combo4.Text = "Normalni" Then '5
    MsgBox "Uneti prenosni odnos za normalni profil kaisa mora biti manji od 15", vbInformation, "Neodgovarajuci unos"
    Text4.Text = ""
    Text4.SetFocus
    
    ElseIf Text4.Text > 10 And Combo4.Text = "Uski" Then '5
    MsgBox "Uneti prenosni odnos za uski profil kaisa mora biti manji od 10", vbInformation, "Neodgovarajuci unos"
    Text4.Text = ""
    Text4.SetFocus
    
    ElseIf Text4.Text < 1 Then '6
    MsgBox "Uneti prenosni odnos mora biti veci ili jednak 1", vbInformation, "Neodgovarajuci unos"
    Text4.Text = ""
    Text4.SetFocus
    
    ElseIf Text1.Text = "" Then '7
    MsgBox "Unesite faktor radnih uslova", vbInformation, "Neodgovarajuci unos"
    Combo1.SetFocus
    
    ElseIf Val(Text1.Text) * Val(Text2.Text) > 400 Then '8
    MsgBox "Merodavna snaga Pm=" & Round(Val(Text1.Text) * Val(Text2.Text), 3) & " kW, mora biti manja od 400 kW", vbInformation, "Neodgovarajuci unos"
Else:

If Combo4.Text = "Uski" And Text4.Text > 10 Then
    MsgBox "Maksimalna vrednost prenosnog odnosa za uski profil kaisa je i=10", vbInformation, "Neodgovarajuci unos"
    Combo4.SetFocus

    ElseIf Combo4.Text = "" Then
    MsgBox "Niste odabrali profil kaiša", vbInformation, "Neodgovarajuci unos"
    Combo4.SetFocus

    ElseIf Combo5.Text = "" Then
    MsgBox "Niste odabrali tip kaiša", vbInformation, "Neodgovarajuci unos"
    Combo5.SetFocus

    ElseIf Combo6.Text = "" Then
    MsgBox "Niste odabrali precnik pogonske remenice", vbInformation, "Neodgovarajuci unos"
    Combo6.SetFocus

    ElseIf Text6.Text = "" Then
    MsgBox "Niste uneli osno rastojanje", vbInformation, "Neodgovarajuci unos"
    Text6.SetFocus
    
    ElseIf Text6.Text <> "" Then
    
    If Val(Text6.Text) < amin Or Val(Text6.Text) > amax Then '2
    MsgBox "Pogresan unos osnog rastojanja", vbInformation, "Neodgovarajuci unos"
    Text6.Text = ""
    Text6.SetFocus
    
    Else:
    az = Text6.Text 'Zeljeno osno rastojanje
    Call Proracun
    End If '2
    
End If '1
End If
End Sub

Public Sub Proracun()
'---------
'PRORACUN
'---------

X = (d2 - d1) / (2 * az) 'Potrebno za racunanje ugla nagiba ogranka
'Atn(X / Sqr(-X * X + 1)) * 180 / Pi  Opsti obrazac za ArcSin u stepenima
alfa = Atn(X / Sqr(-X * X + 1)) * 180 / Pi 'Ugao nagiba ogranka
beta = 180 - 2 * alfa 'Obvojni ugao
Ldr = 2 * az * (Cos(alfa * Pi / 180)) + ((Pi / 2) * (d1 + d2)) + ((alfa * Pi) / 180) * (d2 - d1) 'Racunska vrednost duzine kaisa
Label79.Caption = Round(Ldr, 1) 'Racunska vrednost duzine kaisa

If Combo4.Text = "Normalni" Then '1

    If Combo5.Text = "Z/10 - ISO 4184: 1992 / DIN 2215: 1998" Then '2
    deltaL = 22
    Ldr = Ldr - deltaL
    
    ElseIf Combo5.Text = "A/13 - ISO 4184: 1992 / DIN 2215: 1998" Then
    deltaL = 30
    Ldr = Ldr - deltaL
    
    ElseIf Combo5.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998" Then
    deltaL = 40
    Ldr = Ldr - deltaL
    
    ElseIf Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" Then
    deltaL = 58
    Ldr = Ldr - deltaL

End If '2
End If '1

'Izbor najpribliznije standardne duzine remena
Open App.Path & "/data/Duzine_remena.data" For Input As #100
Do
    Input #100, Ld 'Ld je standardna duzina remena iz baze (R40 red)
    If Ld > Ldr Then
        Lmin = Lmin
        Lmax = Ld
    ElseIf Ld < Ldr Then
        Lmin = Ld
        Lmax = Ldr
    ElseIf Ld = Ldr Then
        Lmin = Ld
        Lmax = Ld
        Exit Do
    End If
        If Ld > Ldr Then
        Exit Do
        End If
Loop Until EOF(100)
Close #100

dl1 = Abs(Lmax - Ldr) 'Prva razlika
dl2 = Abs(Lmin - Ldr) 'Druga razlika
If dl1 <= dl2 Then 'Izbor pribliznije vrednosti za Ld
    Ld = Lmax
Else: Ld = Lmin
End If

If Combo4.Text = "Normalni" Then '1 Petlja koja vrsi korekciju standardne duzine kaisa

    If Combo5.Text = "Z/10 - ISO 4184: 1992 / DIN 2215: 1998" Then '2
    deltaL = 22
    Ld = Ld + deltaL
    
    ElseIf Combo5.Text = "A/13 - ISO 4184: 1992 / DIN 2215: 1998" Then
    deltaL = 30
    Ld = Ld + deltaL
    
    ElseIf Combo5.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998" Then
    deltaL = 40
    Ld = Ld + deltaL
    
    ElseIf Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" Then
    deltaL = 58
    Ld = Ld + deltaL

End If '2
End If '1


Label89.Caption = Round(Ld, 1) 'Standardna vrednost duzine kaisa
a = (1 / 4) * (Ld - ((Pi / 2) * (d1 + d2)) + Sqr((Ld - (Pi / 2) * (d1 + d2)) ^ 2 - 2 * (d2 - d1) ^ 2)) 'Stvarna vrednost osnog rastojanja
Label58.Caption = Round(a, 3) 'Stvarna vrednost osnog rastojanja

'Faktor obvojnog ugla
Cbeta = 2.81601344341423E-03 + 1.06647577362363E-02 * beta - 3.90555160772226E-05 * beta ^ 2 + 5.83701470930601 * 10 ^ (-8) * beta ^ 3

'-----------------------------
'Odredjivanje CL, Pn i deltaP
'-----------------------------

If Combo4.Text = "Uski" Then '(*Pocetak petlje*)

If Combo5.Text = "SPZ - DIN 7753: 1988" Then 'Prva grana
    CL = 0.596033955032149 + 4.11292955686688E-04 * Ld - 1.20542738404947 * 10 ^ (-7) * Ld ^ 2 + 1.37695516501419 * 10 ^ (-11) * Ld ^ 3
    
        If Combo6.Text = "63" Then '1 'd1=63 mm (SPZ)
        Pn = 3.02917825650332E-02 + 7.05761908177262E-04 * n1 - 5.20099940985691 * 10 ^ (-8) * n1 ^ 2
    
        ElseIf Combo6.Text = "80" Then 'd1=80 mm (SPZ)
        Pn = 3.57044897840778E-02 + 1.29902915882218E-03 * n1 - 9.85584345117728 * 10 ^ (-8) * n1 ^ 2
    
        ElseIf Combo6.Text = "90" Then 'd1=90 mm (SPZ)
        Pn = 5.26558064969834E-02 + 1.62191431261831E-03 * n1 - 1.23874774852646 * 10 ^ (-7) * n1 ^ 2
    
        ElseIf Combo6.Text = "112" Then 'd1=112 mm (SPZ)
        Pn = 9.32261200861539E-02 + 2.31218193102963E-03 * n1 - 1.7055092067794 * 10 ^ (-7) * n1 ^ 2
    
        ElseIf Combo6.Text = "140" Then 'd1=140 mm (SPZ)
        Pn = 0.133226219329351 + 3.13380688850059E-03 * n1 - 2.42148307541912 * 10 ^ (-7) * n1 ^ 2
    
        ElseIf Combo6.Text = "180" Then 'd1=180 mm (SPZ)
        Pn = 0.118060880003054 + 4.42116860181043E-03 * n1 - 3.53841872023231 * 10 ^ (-7) * n1 ^ 2
        End If '1
        
            If iss <= 1.05 Then '5 iss=1.05 (SPZ)
            deltaP = 1.6140350877193E-05 * n1
        
            ElseIf iss = 1.2 Then 'iss=1.2 (SPZ)
            deltaP = 8.87719298245614E-05 * n1
            
            ElseIf iss > 1.05 And iss < 1.2 Then ' (SPZ)
            Y = 0.046 + ((iss - 1.05) * (0.253 - 0.046)) / (1.2 - 1.05)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.5 Then 'iss=1.5 (SPZ)
            deltaP = 1.2919649122807E-04 * n1
            
            ElseIf iss > 1.2 And iss < 1.5 Then ' (SPZ)
            Y = 0.253 + ((iss - 1.2) * (0.36821 - 0.253)) / (1.5 - 1.2)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss >= 3 Then 'iss=3 (SPZ)
            deltaP = 1.62268771929825E-04 * n1
            
            ElseIf iss > 1.5 And iss < 3 Then ' (SPZ)
            Y = 0.36821 + ((iss - 1.5) * (0.462466 - 0.36821)) / (3 - 1.5)
            deltaP = (Y / 2850) * n1
                              
            End If '5
            
    ElseIf Combo5.Text = "SPA - DIN 7753: 1988" Then '(SPA)
    CL = 0.612531793018454 + 2.56170363831215E-04 * Ld - 5.06059262351877 * 10 ^ (-8) * Ld ^ 2 + 3.86790494248202 * 10 ^ (-12) * Ld ^ 3
    
        If Combo6.Text = "90" Then '2 'd1=90 mm (SPA)
        Pn = 5.75014730133963E-02 + 1.92276654300578E-03 * n1 - 2.14491340183251 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "112" Then 'd1=112 mm (SPA)
        Pn = 0.116033969633235 + 3.0864506817201E-03 * n1 - 2.91316387434343 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "125" Then 'd1=125 mm (SPA)
        Pn = 0.137732511590396 + 3.80715108876425E-03 * n1 - 3.7702845319191 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "160" Then 'd1=160 mm (SPA)
        Pn = 0.10561928565909 + 5.68988117402795E-03 * n1 - 5.5378369505407 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "200" Then 'd1=200 mm (SPA)
        Pn = 0.131047373643233 + 7.85692893256546E-03 * n1 - 8.28066557978848 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "250" Then 'd1=250 mm (SPA)
        Pn = 0.236550063300297 + 1.04964071894825E-02 * n1 - 1.34945716910947 * 10 ^ (-6) * n1 ^ 2
        
        End If '2
        
            If iss <= 1.05 Then '6 iss=1.05 (SPA)
            deltaP = 4.6140350877193E-05 * n1
            
            ElseIf iss = 1.2 Then 'iss=1.2 (SPA)
            deltaP = 2.09566315789474E-04 * n1
            
            ElseIf iss > 1.05 And iss < 1.2 Then ' (SPA)
            Y = 0.1315 + ((iss - 1.05) * (0.597264 - 0.1315)) / (1.2 - 1.05)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.5 Then 'iss=1.5 (SPA)
            deltaP = 2.79740350877193E-04 * n1
            
            ElseIf iss > 1.2 And iss < 1.5 Then ' (SPA)
            Y = 0.597264 + ((iss - 1.2) * (0.79726 - 0.597264)) / (1.5 - 1.2)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss >= 3 Then 'iss=3 (SPA)
            deltaP = 3.79715789473684E-04 * n1
            
            ElseIf iss > 1.5 And iss < 3 Then ' (SPA)
            Y = 0.79726 + ((iss - 1.5) * (1.08219 - 0.79726)) / (3 - 1.5)
            deltaP = (Y / 2850) * n1
                              
            End If '6
    
    ElseIf Combo5.Text = "SPB - DIN 7753: 1988" Then '(SPB)
    CL = 0.675620363449437 + 1.45298151213704E-04 * Ld - 1.8164742494519 * 10 ^ (-8) * Ld ^ 2 + 8.82351335667328 * 10 ^ (-13) * Ld ^ 3
    
        If Combo6.Text = "140" Then '3 'd1=140 mm (SPB)
        Pn = 0.122032717815275 + 6.13880869044961E-03 * n1 - 8.72737874349275 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "180" Then 'd1=180 mm (SPB)
        Pn = 5.56516587791557E-02 + 9.57485020284786E-03 * n1 - 1.24606752652484 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "200" Then 'd1=200 mm (SPB)
        Pn = (-4.76849913319663E-02 + 0.012506503862904 * n1 - 1.85891899865425 * 10 ^ (-6) * n1 ^ 2)
        
        ElseIf Combo6.Text = "250" Then 'd1=250 mm (SPB)
        Pn = 6.96584502169044E-02 + 1.66166431965154E-02 * n1 - 2.57488970672599 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "315" Then 'd1=315 mm (SPB)
        Pn = 1.86402153145349E-02 + 2.27677704405289E-02 * n1 - 4.06497689433805 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "400" Then 'd1=400 mm (SPB)
        Pn = 6.76474869321939E-02 + 3.07389903731681E-02 * n1 - 6.23349934214868 * 10 ^ (-6) * n1 ^ 2 + 2.58059838636545 * 10 ^ (-11) * n1 ^ 3
        
        End If '3
        
            If iss <= 1.05 Then '7 iss=1.05 (SPB)
            deltaP = 7.49821052631579E-05 * n1
            
            ElseIf iss > 1.05 And iss < 1.2 Then ' (SPB)
            Y = 0.213699 + ((iss - 1.05) * (1.29315 - 0.213699)) / (1.2 - 1.05)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.2 Then 'iss=1.2 (SPB)
            deltaP = 4.53736842105263E-04 * n1
            
            ElseIf iss > 1.2 And iss < 1.5 Then ' (SPB)
            Y = 1.29315 + ((iss - 1.2) * (1.8411 - 1.29315)) / (1.5 - 1.2)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.5 Then 'iss=1.5 (SPB)
            deltaP = 0.000646 * n1
        
            ElseIf iss >= 3 Then 'iss=3 (SPB)
            deltaP = 8.01729824561404E-04 * n1
            
            ElseIf iss > 1.5 And iss < 3 Then ' (SPB)
            Y = 1.8411 + ((iss - 1.5) * (2.28493 - 1.8411)) / (3 - 1.5)
            deltaP = (Y / 2850) * n1
                              
            End If '7
    
    ElseIf Combo5.Text = "SPC - DIN 7753: 1988" Then '(SPC)
    CL = 0.696635987069896 + 9.14425009853484E-05 * Ld - 8.24481466124532 * 10 ^ (-9) * Ld ^ 2 + 2.91491210703685 * 10 ^ (-13) * Ld ^ 3
    
        If Combo6.Text = "224" Then '4 'd1=224 mm (SPC)
        Pn = 0.219948713106249 + 1.83644493859328E-02 * n1 - 3.32177019870555 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "280" Then 'd1=280 mm (SPC)
        Pn = 0.313138420590936 + 2.82218098291687E-02 * n1 - 5.44990795814957 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "315" Then 'd1=315 mm (SPC)
        Pn = 0.326569202746017 + 3.52309713003293E-02 * n1 - 7.43875404996951 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "400" Then 'd1=400 mm (SPC)
        Pn = 0.188355560258505 + 5.01330509025675E-02 * n1 - 1.17867613985069E-05 * n1 ^ 2
        
        ElseIf Combo6.Text = "500" Then 'd1=500 mm (SPC)
        Pn = 0.319496772206329 + 0.067830462329573 * n1 - 1.87939302933463E-05 * n1 ^ 2
        
        ElseIf Combo6.Text = "630" Then 'd1=630 mm (SPC)
        Pn = 9.50175828227011E-02 + 8.51388370096687E-02 * n1 + 4.13993108690236 * 10 ^ (-6) * n1 ^ 2 - 6.85026141524762 * 10 ^ (-8) * n1 ^ 3 + 5.35358669017674 * 10 ^ (-11) * n1 ^ 4 - 1.42573272392514 * 10 ^ (-14) * n1 ^ 5
        
        End If '4
        
            If iss <= 1.05 Then '8 iss=1.05 (SPC)
            deltaP = 2.36380350877193E-04 * n1
            
            ElseIf iss > 1.05 And iss < 1.2 Then ' (SPB)
            Y = 0.673684 + ((iss - 1.05) * (3.67597 - 0.673684)) / (1.2 - 1.05)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.2 Then 'iss=1.2 (SPC)
            deltaP = 1.28981403508772E-03 * n1
            
            ElseIf iss > 1.2 And iss < 1.5 Then ' (SPB)
            Y = 3.67597 + ((iss - 1.2) * (5.19908 - 3.67597)) / (1.5 - 1.2)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.5 Then 'iss=1.5 (SPC)
            deltaP = 1.82423859649123E-03 * n1
        
            ElseIf iss >= 3 Then 'iss=3 (SPC)
            deltaP = 2.24561403508772E-03 * n1
            
            ElseIf iss > 1.5 And iss < 3 Then ' (SPB)
            Y = 5.19908 + ((iss - 1.5) * (6.4 - 5.19908)) / (3 - 1.5)
            deltaP = (Y / 2850) * n1
                              
            End If '8
            
End If 'Prva grana

Else: 'Druga grana

    If Combo5.Text = "Z/10 - ISO 4184: 1992 / DIN 2215: 1998" Then '(10/Z)
    CL = 0.633357 + 0.000633894 * Ld - 2.7015 * 10 ^ (-7) * Ld ^ 2 + 4.71291 * 10 ^ (-11) * Ld ^ 3
    
        If Combo6.Text = "50" Then '1 'd1=50 mm (10/Z)
        Pn = 0.0173925 + 0.000287537 * n1 - 2.54991 * 10 ^ (-8) * n1 ^ 2
        
        ElseIf Combo6.Text = "63" Then 'd1=63 mm (10/Z)
        Pn = 0.0163234 + 0.000474647 * n1 - 4.204 * 10 ^ (-8) * n1 ^ 2
        
        ElseIf Combo6.Text = "80" Then 'd1=80 mm (10/Z)
        Pn = 0.0241626 + 0.000683722 * n1 - 5.36687 * 10 ^ (-8) * n1 ^ 2
        
        ElseIf Combo6.Text = "100" Then 'd1=100 mm (10/Z)
        Pn = 0.0293112 + 0.000925422 * n1 - 7.4718 * 10 ^ (-8) * n1 ^ 2
        
        ElseIf Combo6.Text = "125" Then 'd1=125 mm (10/Z)
        Pn = 0.0249536 + 0.00127287 * n1 - 1.17042 * 10 ^ (-7) * n1 ^ 2
        
        End If '1
        
            If iss <= 1.05 Then '2 iss=1.05 (10/Z)
            deltaP = 0.0000105263 * n1
            
            ElseIf iss > 1.05 And iss < 1.2 Then ' (10/Z)
            Y = 0.03 + ((iss - 1.05) * (0.08 - 0.03)) / (1.2 - 1.05)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.2 Then 'iss=1.2 (10/Z)
            deltaP = 0.0000280702 * n1
            
            ElseIf iss > 1.2 And iss < 1.5 Then ' (10/Z)
            Y = 0.08 + ((iss - 1.2) * (0.119 - 0.08)) / (1.5 - 1.2)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.5 Then 'iss=1.5 (10/Z)
            deltaP = 0.0000417544 * n1
        
            ElseIf iss >= 3 Then 'iss=3 (10/Z)
            deltaP = 0.0000526316 * n1
            
            ElseIf iss > 1.5 And iss < 3 Then '(10/Z)
            Y = 0.119 + ((iss - 1.5) * (0.15 - 0.119)) / (3 - 1.5)
            'deltaP = (y / 2850) * n1
                              
            End If '2
            
    ElseIf Combo5.Text = "A/13 - ISO 4184: 1992 / DIN 2215: 1998" Then '3 (13/A)
    CL = 0.52996 + 0.000496838 * Ld - 1.8152 * 10 ^ (-7) * Ld ^ 2 + 3.558417 * 10 ^ (-11) * Ld ^ 3 - 2.694825 * 10 ^ (-15) * Ld ^ 4
    
        If Combo6.Text = "80" Then '1 'd1=80 mm (13/A)
        Pn = 0.0140209 + 0.00129658 * n1 - 1.864064 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "100" Then 'd1=100 mm (13/A)
        Pn = 0.0671937 + 0.00194318 * n1 - 2.50408 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "125" Then 'd1=125 mm (13/A)
        Pn = 0.0732005 + 0.00283319 * n1 - 3.62944 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "160" Then 'd1=160 mm (13/A)
        Pn = 0.0931001 + 0.00393568 * n1 - 4.930972 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "200" Then 'd1=200 mm (13/A)
        Pn = 0.265721 + 0.00477755 * n1 - 5.180685 * 10 ^ (-7) * n1 ^ 2
        
        End If '3 (13/A)
        
            If iss <= 1.05 Then '4 iss=1.05 (13/A)
            deltaP = 0.000042807 * n1
            
            ElseIf iss > 1.05 And iss < 1.2 Then ' (13/A)
            Y = 0.122 + ((iss - 1.05) * (0.346 - 0.122)) / (1.2 - 1.05)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.2 Then 'iss=1.2 (13/A)
            deltaP = 0.000121404 * n1
            
            ElseIf iss > 1.2 And iss < 1.5 Then ' (13/A)
            Y = 0.346 + ((iss - 1.2) * (0.54 - 0.346)) / (1.5 - 1.2)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.5 Then 'iss=1.5 (13/A)
            deltaP = 0.000189474 * n1
        
            ElseIf iss >= 3 Then 'iss=3 (13/A)
            deltaP = 0.000230526 * n1
            
            ElseIf iss > 1.5 And iss < 3 Then '(13/A)
            Y = 0.54 + ((iss - 1.5) * (0.657 - 0.54)) / (3 - 1.5)
            deltaP = (Y / 2850) * n1
                              
            End If '4

    ElseIf Combo5.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998" Then '5 (17/B)
    CL = 0.51644 + 0.000391242 * Ld - 1.0874896 * 10 ^ (-7) * Ld ^ 2 + 1.547982 * 10 ^ (-11) * Ld ^ 3 - 8.275174 * 10 ^ (-16) * Ld ^ 4
    
        If Combo6.Text = "125" Then '1 'd1=125 mm (17/B)
        Pn = 0.118998 + 0.00329204 * n1 - 5.1499974 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "160" Then 'd1=160 mm (17/B)
        Pn = 0.26266 + 0.00527058 * n1 - 7.5895368 * 10 ^ (-7) * n1 ^ 2
        
        ElseIf Combo6.Text = "200" Then 'd1=200 mm (17/B)
        Pn = 0.15348015 + 0.00810806 * n1 - 1.40237887 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "250" Then 'd1=250 mm (17/B)
        Pn = 0.1706832 + 0.0108237 * n1 - 1.92281904 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "315" Then 'd1=315 mm (17/B)
        Pn = 0.130265 + 0.0149575 * n1 - 3.25028 * 10 ^ (-6) * n1 ^ 2
        
        End If '5 (17/B)
        
            If iss <= 1.05 Then '6 iss=1.05 (17/B)
            deltaP = 0.0000982456 * n1
            
            ElseIf iss > 1.05 And iss < 1.2 Then ' (17/B)
            Y = 0.28 + ((iss - 1.05) * (0.8 - 0.28)) / (1.2 - 1.05)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.2 Then 'iss=1.2 (17/B)
            deltaP = 0.000280702 * n1
            
            ElseIf iss > 1.2 And iss < 1.5 Then ' (17/B)
            Y = 0.8 + ((iss - 1.2) * (1.2 - 0.8)) / (1.5 - 1.2)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.5 Then 'iss=1.5 (17/B)
            deltaP = 0.000421053 * n1
        
            ElseIf iss >= 3 Then 'iss=3 (17/B)
            deltaP = 0.000526316 * n1
            
            ElseIf iss > 1.5 And iss < 3 Then '(17/B)
            Y = 1.2 + ((iss - 1.5) * (1.5 - 1.2)) / (3 - 1.5)
            deltaP = (Y / 2850) * n1
                              
            End If '6
            
    ElseIf Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" Then '7 (22/C)
    CL = 0.550081 + 0.000221326 * Ld - 3.867609 * 10 ^ (-8) * Ld ^ 2 + 3.629531 * 10 ^ (-12) * Ld ^ 3 - 1.3457647 * 10 ^ (-16) * Ld ^ 4
    
        If Combo6.Text = "200" Then '1 'd1=200 mm (22/C)
        Pn = 0.179525 + 0.0100296 * n1 - 1.9720618 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "250" Then 'd1=250 mm (22/C)
        Pn = 0.12323 + 0.0149723 * n1 - 3.166377 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "315" Then 'd1=315 mm (22/C)
        Pn = 0.118812 + 0.0209676 * n1 - 4.963986 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "400" Then 'd1=400 mm (22/C)
        Pn = 0.445658 + 0.0258082 * n1 - 6.005561 * 10 ^ (-6) * n1 ^ 2
        
        ElseIf Combo6.Text = "500" Then 'd1=500 mm (22/C)
        Pn = 0.493321 + 0.0363513 * n1 - 9.769366 * 10 ^ (-6) * n1 ^ 2
        
        End If '7 (22/C)
        
            If iss <= 1.05 Then '8 iss=1.05 (22/C)
            deltaP = 0.000214035 * n1
            
            ElseIf iss > 1.05 And iss < 1.2 Then ' (22/C)
            Y = 0.61 + ((iss - 1.05) * (1.83 - 0.61)) / (1.2 - 1.05)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.2 Then 'iss=1.2 (22/C)
            deltaP = 0.000642105 * n1
            
            ElseIf iss > 1.2 And iss < 1.5 Then '(22/C)
            Y = 1.83 + ((iss - 1.2) * (2.8 - 1.83)) / (1.5 - 1.2)
            deltaP = (Y / 2850) * n1
        
            ElseIf iss = 1.5 Then 'iss=1.5 (22/C)
            deltaP = 0.000982456 * n1
        
            ElseIf iss >= 3 Then 'iss=3 (22/C)
            deltaP = 0.00124561 * n1
            
            ElseIf iss > 1.5 And iss < 3 Then '(22/C)
            Y = 2.8 + ((iss - 1.5) * (3.55 - 2.8)) / (3 - 1.5)
            deltaP = (Y / 2850) * n1
                              
            End If '8

End If 'Druga grana

End If '(*Kraj petlje*)

v = (d1 * Pi * n1) / 60000 'Obimna brzina u m/s
Label33.Caption = Round(v, 3) 'Obimna brzina u m/s

Hi = 2 'Broj remenica
fs = (Hi * v) / (Ld / 1000) 'Ucestanost savijanja u s^(-1)
Label42.Caption = Round(fs, 3) 'Ucestanost savijanja u s^(-1)

zr = Pm / ((Pn + deltaP) * Cbeta * CL) 'Racunska vrednost broja zlebova
Label67.Caption = Round(zr, 3) 'Racunska vrednost broja zlebova

'Usvajanje prve vece celobrojne vrednosti broja zlebova:
z = Round(zr, 0)
If zr < z Then
z = Round(zr, 0)
Else: z = z + 1
End If
Label73.Caption = z

'Kontrola rezultata
If v < vmax Then 'Provera obimne brzine
Label33.BackColor = &H80FF80 'zelena
Else: Label33.BackColor = &H8080FF 'crvena
End If

If fs < fsdoz Then 'Provera ucestanosti
Label42.BackColor = &H80FF80 'zelena
Else: Label42.BackColor = &H8080FF 'crvena
End If

If z <= 5 Then 'Provera broja zlebova
Label73.BackColor = &H80FF80 'zelena
ElseIf z > 5 And z <= 12 Then
Label73.BackColor = &H80FFFF 'zuta
ElseIf z > 12 Then
Label73.BackColor = &H8080FF 'crvena
End If

'Prikaz dugmeta "Izbor jaceg profila"
If Combo5.Text = "SPC - DIN 7753: 1988" Or Combo5.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" Then
    Command1.Visible = False
Else: Command1.Visible = True
End If

'Prepisvanje vrednosti sa taba "Proracun" u tab "CAD"
Combo7.Enabled = False 'Profil kaisa
Combo8.Enabled = False 'Tip Kaisa
Text7.Enabled = False 'Precnik d1
Combo7.Text = Combo4.Text 'Profil kaisa
Combo8.Text = Combo5.Text 'Tip Kaisa
Text7.Text = Combo6.Text 'Precnik d1
Text8.Text = Text5.Text 'Precnik d2
Option1.Value = True 'Predhodno proracunata remenica
Option3.Value = True 'd1 aktivno

'PROVERA
'MsgBox "alfa= " & Round(alfa, 3) & "  beta= " & Round(beta, 3) & "  Ldr= " & Round(Ldr, 3) & "  Cbeta= " & Round(Cbeta, 3) & " CL= " & Round(CL, 3) & "  Pn=" & Round(Pn, 3) & "  deltaP=" & Round(deltaP, 3)
End Sub

Private Sub Command5_Click()
'----------------------------
'Provera kontrolnih podataka
'----------------------------

If Label73.BackColor = &H8000000F Or Label33.BackColor = &H8000000F Or Label42.BackColor = &H8000000F Then 'Siva boja svih kontrolnih labela
    MsgBox "Proracun nije izvršen.", vbInformation, "Neodgovarajuci unos"
    
    ElseIf Label73.BackColor = &H8080FF Then 'Crvena boja za broj zlebova
    MsgBox "Dozvoljeni broj žlebova je z=12. Odaberite jaci tip kaiša.", vbInformation, "Neodgovarajuci unos"
    Combo5.SetFocus
    
    ElseIf Label33.BackColor = &H8080FF Then 'Crvena boja za obimnu brzinu
    MsgBox "Maksimalna obimna brzina je v=" & Label36.Caption & " m/s. Odaberite manji precnik pogonske remenice.", vbInformation, "Neodgovarajuci unos"
    Combo6.SetFocus
    
    ElseIf Label42.BackColor = &H8080FF Then 'Crvena boja za ucestanost savijanja
    MsgBox "Dozvolena ucestanost savijanja je fs=" & Label45.Caption & " s-1.", vbInformation, "Neodgovarajuci unos"
    Combo6.SetFocus
    
Else:
    ime_prezime = InputBox("Unesite Vase ime i prezime:", "Ime i prezime")
    Call Izvestaj
End If

End Sub

Public Sub Izvestaj()
'-------------------
'Izvestaj u WORD-u
'-------------------
Dim objWord As Word.Application
Dim objDoc As Word.Document

'Pokretanje novog dokumenta u Word-u
Set objWord = CreateObject("Word.Application")
objWord.Visible = True
objWord.Application.WindowState = wdWindowStateMaximize
Set objDoc = objWord.Documents.Add

'Formatiranje stranice
With objDoc.PageSetup
    .LineNumbering.Active = False
    .Orientation = wdOrientPortrait
    .TopMargin = objWord.MillimetersToPoints(10)
    .BottomMargin = objWord.MillimetersToPoints(10)
    .LeftMargin = objWord.MillimetersToPoints(20)
    .RightMargin = objWord.MillimetersToPoints(10)
    .Gutter = objWord.MillimetersToPoints(0)
    .HeaderDistance = objWord.MillimetersToPoints(12.7)
    .FooterDistance = objWord.MillimetersToPoints(12.7)
    .PageWidth = objWord.MillimetersToPoints(210)
    .PageHeight = objWord.MillimetersToPoints(297)
    .FirstPageTray = wdPrinterDefaultBin
    .OtherPagesTray = wdPrinterDefaultBin
    .SectionStart = wdSectionNewPage
    .OddAndEvenPagesHeaderFooter = False
    .DifferentFirstPageHeaderFooter = False
    .VerticalAlignment = wdAlignVerticalTop
    .SuppressEndnotes = False
    .MirrorMargins = False
    .TwoPagesOnOne = False
    .BookFoldPrinting = False
    .BookFoldRevPrinting = False
    .BookFoldPrintingSheets = 1
    .GutterPos = wdGutterPosLeft
End With
   
'Pisanje uvodnog dela
    objWord.Selection.Font.Name = "Courier New"
    objWord.Selection.Font.Size = 12
    objWord.Selection.ParagraphFormat.LineSpacing = objWord.LinesToPoints(1.15) 'LineSpacing: 1.15
    objWord.Selection.TypeText Text:="Mašinski fakultet Univerziteta u Nišu"
    objWord.Selection.TypeParagraph
    objWord.Selection.TypeText Text:="Katedra za Mašinske konstrukcije, razvoj i inženjering"
    objWord.Selection.TypeParagraph
    objWord.Selection.TypeText Text:="Obradio: " & ime_prezime
    objWord.Selection.TypeParagraph
    objWord.Selection.TypeText Text:="Datum: "
    objWord.Application.WindowState = wdWindowStateMaximize
    objWord.Selection.InsertDateTime DateTimeFormat:="d.M.yyyy", InsertAsField:=False, _
    DateLanguage:=9242, CalendarType:=wdCalendarWestern, InsertAsFullWidth:=False
    objWord.Selection.TypeParagraph
    
'Naslov dokumenta
    objWord.Selection.Font.Size = 14
    objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    objWord.Selection.Font.Bold = wdToggle 'Upali Bold
    objWord.Selection.TypeText Text:="Prora" & ChrW(269) & "un remenih prenosnika" 'Proracun remenih prenosnika
    objWord.Selection.Font.Bold = wdToggle 'Ugasi Bold
    objWord.Selection.TypeParagraph
    
'Podnaslov "Ulazni podaci:"
    objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    objWord.Selection.Font.Size = 12
    objWord.Selection.TypeParagraph
    objWord.Selection.Font.Underline = wdUnderlineSingle 'Upali underline
    objWord.Selection.TypeText Text:="Ulazni podaci:" 'Ulazni podaci
    objWord.Selection.Font.Underline = wdUnderlineNone 'Ugasi underline
    objWord.Selection.TypeParagraph
    
'Podesavanje TabStop-a
    objWord.ActiveDocument.DefaultTabStop = objWord.MillimetersToPoints(12.7)
    objWord.Selection.ParagraphFormat.TabStops.Add Position:=objWord.MillimetersToPoints(90), _
    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces 'Tabulator je na 90 mm
    
'Stampanje ulaznih podataka
    objWord.Selection.Font.Size = 11
    objWord.Selection.TypeText Text:="Nominalna snaga na ulazu" & vbTab & "P" 'Nominalna snaga na ulazu P1
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="1"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Form1.Text2.Text & " kW"
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Broj obrtaja na ulazu" & vbTab & "n" 'Broj obrtaja na ulazu n1
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="1"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Form1.Text3.Text & " min"
    objWord.Selection.Font.Superscript = wdToggle
    objWord.Selection.TypeText Text:="-1"
    objWord.Selection.Font.Superscript = wdToggle
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Prenosni odnos" & vbTab & "i=" + Form1.Text4.Text 'Prenosni odnos i
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Faktor radnih uslova" & vbTab & "K" 'Faktor radnih uslova Ka
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="A"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Form1.Text1.Text
    objWord.Selection.TypeParagraph

'Podnaslov "Proracun:"
    objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    objWord.Selection.Font.Size = 12
    objWord.Selection.TypeParagraph
    objWord.Selection.Font.Underline = wdUnderlineSingle 'Upali underline
    objWord.Selection.TypeText Text:="Prora" & ChrW(269) & "un:" 'Proracun
    objWord.Selection.Font.Underline = wdUnderlineNone 'Ugasi underline
    objWord.Selection.TypeParagraph
    
'Stampanje proracuna
    objWord.Selection.Font.Size = 11
    
    objWord.Selection.TypeText Text:="Pre" & ChrW(269) & "nik pogonske remenice" & vbTab & "d" 'Precnik pogonske remenice d1
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="1"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Combo6.Text & " mm"
    objWord.Selection.TypeParagraph

    objWord.Selection.TypeText Text:="Pre" & ChrW(269) & "nik pogonske remenice" & vbTab & "d" 'Precnik gonjene remenice d2
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="2"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Text5.Text & " mm"
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Ra" & ChrW(269) & "unska vrednost osnog rastojanja" & vbTab & "a=" & Form1.Text6.Text & " mm" 'Racunska vrednost osnog rastojanja a
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Ugao nagiba ogranka" & vbTab & ChrW(945) 'Ugao nagiba ogranka "alfa"
    objWord.Selection.TypeText Text:="=" & Round(alfa, 3)
    objWord.Selection.Font.Superscript = wdToggle
    objWord.Selection.TypeText Text:="O"
    objWord.Selection.Font.Superscript = wdToggle
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Obvojni ugao" & vbTab & ChrW(946) 'Obvojni ugao "beta"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="1"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Round(beta, 3)
    objWord.Selection.Font.Superscript = wdToggle
    objWord.Selection.TypeText Text:="O"
    objWord.Selection.Font.Superscript = wdToggle
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Ra" & ChrW(269) & "unska vrednost dužine kaiša" & vbTab & "L" 'Racunska vrednost duzine remen Ldr
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="dr"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Round(Ldr, 3) & " mm"
    objWord.Selection.TypeParagraph
    
If Form1.Combo4.Text = "Normalni" Then
    objWord.Selection.TypeText Text:="Korekcija dužine kaiša" & vbTab & ChrW(916) & "L=" & deltaL & " mm" 'Korekcija dužine kaiša deltaL
    objWord.Selection.TypeParagraph
End If
    
    objWord.Selection.TypeText Text:="Stvarno osno rastojanje" & vbTab & "a" 'Stvarno osno rastojanje a
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="s"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Round(a, 3) & " mm"
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Nominalna specifi" & ChrW(269) & "na snaga" & vbTab & "P" 'Nominalna specificna snaga Pn
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="N"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Round(Pn, 3) & " kW"
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Dodatna snaga po remenu" & vbTab & ChrW(916) & "P" 'Nominalna specificna snaga Pn
    objWord.Selection.TypeText Text:="=" & Round(deltaP, 3) & " kW"
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Faktor dužine remena" & vbTab & "C" 'Faktor duzine remena CL
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="L"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Round(CL, 3)
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Faktor obvojnog ugla" & vbTab & "C" 'Faktor obvojnog ugla Cbeta
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:=ChrW(946)
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Round(Cbeta, 3)
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Ra" & ChrW(269) & "unska vrednost broja žlebova" & vbTab & "z" 'Racunska vrednost broja žlebova zr
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="r"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Round(zr, 3)
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Obimna brzina" & vbTab & ChrW(957) 'Obimna brzina v
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="1"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Round(v, 3) & " m/s"
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Maksimalna obimna brzina" & vbTab & ChrW(957) 'Maksimalna obimna brzina vmax
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="max"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Form1.Label36.Caption & " m/s"
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="U" & ChrW(269) & "estnost savijanja" & vbTab & "f" 'Ucestanost savijanja fs
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="s"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Form1.Label42.Caption & " s"
    objWord.Selection.Font.Superscript = wdToggle
    objWord.Selection.TypeText Text:="-1"
    objWord.Selection.Font.Superscript = wdToggle
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Dozvoljena u" & ChrW(269) & "estnost savijanja" & vbTab & "f" 'Dozvoljena ucestanost savijanja fsdoz
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="sdoz"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Form1.Label45.Caption & " s"
    objWord.Selection.Font.Superscript = wdToggle
    objWord.Selection.TypeText Text:="-1"
    objWord.Selection.Font.Superscript = wdToggle
    objWord.Selection.TypeParagraph
    
'Podnaslov "Rezultati:"
    objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    objWord.Selection.Font.Size = 12
    objWord.Selection.TypeParagraph
    objWord.Selection.Font.Underline = wdUnderlineSingle 'Upali underline
    objWord.Selection.TypeText Text:="Rezultati:" 'Rezultati
    objWord.Selection.Font.Underline = wdUnderlineNone 'Ugasi underline
    objWord.Selection.TypeParagraph
    
'Stampanje rezultata
    objWord.Selection.Font.Size = 11
    
    objWord.Selection.TypeText Text:="Profil kaiša" & vbTab & Form1.Combo4.Text  'Profil kaisa: Normalni/Uski
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Tip kaiša" & vbTab & Form1.Combo5.Text  'Tip kaisa: Z/10, A/13, B/17...
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Standardna vrednost dužine kaiša" & vbTab & "L" 'Standardna vrednost dužine kaiša Ld
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="d"
    objWord.Selection.Font.Subscript = wdToggle
    objWord.Selection.TypeText Text:="=" & Ld & " mm"
    objWord.Selection.TypeParagraph
    
    objWord.Selection.TypeText Text:="Broj žlebova" & vbTab & "z=" & z  'Broj žlebova z
    
'------
'ASCII:
'------

'delta: ChrW(916)
'alfa: ChrW(945)
'beta: ChrW(946)
'è: ChrW(269)
'Grcko v: ChrW(957)
End Sub

Private Sub Command7_Click()
'Zapisivanje podataka u temp fajlu
Open App.Path & "/data/temp.data" For Output As #1
Write #1, Text2.Text, Text3.Text, Text4.Text
Close #1
'Napustanje programa
End
End Sub

Private Sub Form_Load()

 SSTab1.Tab = 0 'Pokretanje taba "Proracun" pri ucitavanju forme
 Option2.Value = True 'Aktivno "Nezavisno modeliranje remenice" pri pokretanju forme
 
'----------------------------------
'Ucitavanje podataka iz temp fajla
'----------------------------------
Open App.Path & "/data/temp.data" For Input As #1
Input #1, P1, n1, i
Close #1

Text2.Text = Val(P1)
Text3.Text = Val(n1)
Text4.Text = Val(i)

Label75.Visible = False 'Sakrivanje Label-e za prikaz amin i amax
Command1.Visible = False 'Sakrivanje dugmeta za "Izbor jaceg profila"

'---------------------------------------------------------------------------
'Definisanje faktora radnih uslova Ka (Dodavanje vrednosti u ComboBox-ovima)
'---------------------------------------------------------------------------

'Pogonska masina
Combo1.AddItem "AC/DC elektromotor sa normalnim polaznim momentom"  'Grupa A
Combo1.AddItem "SUS motor i turbine sa n>600 min-1"                 'Grupa A
Combo1.AddItem "AC/DC elektromotor sa velikim polaznim momentom"    'Grupa B
Combo1.AddItem "SUS motor i turbine sa n<600 min-1"                 'Grupa B

'Dnevni rad
Combo2.AddItem "Do 10 h"
Combo2.AddItem "Iznad 10 h a manje od 16 h"
Combo2.AddItem "Iznad 16 h"

'Radna masina
'Laki spektar opterecenja
Combo3.AddItem "Turbopumpe"
Combo3.AddItem "Turbokompresori"
Combo3.AddItem "Trakasti kompresori za lake materijale"
Combo3.AddItem "Ventilatori i pumpe do 7.4 kW"
Combo3.AddItem "Ravnomerno optereceni strugovi"
Combo3.AddItem "Ravnomerno opterecene brusilice"

'Srednji spektar opterecenja
Combo3.AddItem "Makaze za lim"
Combo3.AddItem "Prese"
Combo3.AddItem "Lancani transporteri za teske materijale"
Combo3.AddItem "Trakasti transporteri za teske materijale"
Combo3.AddItem "Elektro generatori"
Combo3.AddItem "Alatne masine"
Combo3.AddItem "Stamparske masine"
Combo3.AddItem "Ventilatori i pumpe preko 7.4 kW"

'Tezak spektar opterecenja
Combo3.AddItem "Mlinovi"
Combo3.AddItem "Klipni kompresori"
Combo3.AddItem "Puzni kompresori"
Combo3.AddItem "Tekstilne masine"
Combo3.AddItem "Masine za hartiju"
Combo3.AddItem "Testere"
Combo3.AddItem "Gateri"
Combo3.AddItem "Prese za briket"

'Veoma tezak spektar opterecenja
Combo3.AddItem "Jako optereceni mlinovi"
Combo3.AddItem "Drobilice"
Combo3.AddItem "Mesalice"
Combo3.AddItem "Kalenderi"
Combo3.AddItem "Vitla"
Combo3.AddItem "Kranovi"
Combo3.AddItem "Bageri"
Combo3.AddItem "Valjaonicki stanovi"

'Ostale masine
Combo3.AddItem "Ostale masine"
Text1.Enabled = False  'Iskljucivanje mogucnosti unosa faktora radnih uslova Ka

Pi = 4 * Atn(1) ' Pi=3.1415...

'---------------------------------
'Definisanje redosleda TabIndex-a
'---------------------------------
Text2.TabIndex = 0
Text3.TabIndex = 1
Text4.TabIndex = 2
Combo4.TabIndex = 3
Combo1.TabIndex = 4
Combo2.TabIndex = 5
Combo3.TabIndex = 6
Combo5.TabIndex = 7
Combo6.TabIndex = 8
Text6.TabIndex = 9
Command4.TabIndex = 10
Command5.TabIndex = 11
Command7.TabIndex = 12
Command1.TabIndex = 13

'----------------------------
'Definisanje profila kaiseva
'----------------------------
Combo4.AddItem "Normalni"
Combo4.AddItem "Uski"

'--------------------------------------
'Pocetne vrednosti za prikaz u labelama
'--------------------------------------
Label1.Caption = "0.000"
Label52.Caption = "0.000"
Label79.Caption = "0.000"
Label89.Caption = "0.000"
Label58.Caption = "0.000"
Label67.Caption = "0.000"
Label73.Caption = "0.000"
Label33.Caption = "0.000"
Label36.Caption = "0.000"
Label42.Caption = "0.000"
Label45.Caption = "0.000"

'------------------------------------
'Definisanje najveceg broja karaktera
'------------------------------------
Text1.MaxLength = 7 'Unos Ka
Text2.MaxLength = 7 'Unos snage P1
Text3.MaxLength = 7 'Unos broja obrtaja n1
Text4.MaxLength = 7 'Unos prenosnog odnosa i
Text6.MaxLength = 7 'Unos osnog rastojanja a
Text7.MaxLength = 7 'Unos precnika d1
Text8.MaxLength = 7 'Unos precnika d2
Text9.MaxLength = 2 'Unos broja zlebova

'--------------
'ToolTip poruke
'--------------
Option1.ToolTipText = "Ova opcija se odnosi na modeliranje vec proracunate remenice"
Option2.ToolTipText = "Ova opcija se odnosi na nezavisno modeliranje remenice"
Text5.ToolTipText = "Standardne vrednosti precnika gonjene remenice bira se iz reda standardnih brojeva R20"
Text6.ToolTipText = "Vrednost osnog rastojanja mora biti u granicama amin<a<amax"
Label89.ToolTipText = "Standardna vrednost duzine kaisa bira se iz reda standardnih brojeva R40"

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Zapisivanje podataka u temp fajlu
Open App.Path & "/data/temp.data" For Output As #1
Write #1, Text2.Text, Text3.Text, Text4.Text
Close #1
End Sub

Private Sub Image4_Click()

End Sub

Private Sub Option1_Click()

'Predhodno proracunata remenica
'-----------------------------------------
'Ako je proracun izvrsen prepisi vrednosti
'-----------------------------------------
If Label73.BackColor = &H80000004 Then
    MsgBox "Proracun nije izvrsen", vbInformation, "Neodgovarajuci izbor"
    Option2.Value = True
Else:
    Combo7.Enabled = False
    Combo8.Enabled = False
    Text7.Enabled = False
    Text8.Enabled = False
    Label71(1).Visible = False
    Text9.Visible = False
    Combo7.Text = Combo4.Text
    Combo8.Text = Combo5.Text
    Text7.Text = Combo6.Text
    Text8.Text = Text5.Text
    
    Option4.Visible = True 'Precnik gonjene remenice
    Option3.Visible = True 'Precnik gonjene remenice
    Text8.Visible = True 'Precnik gonjene remenice
    Label31.Visible = True 'Precnik gonjene remenice
    Label29.Visible = True 'Precnik gonjene remenice
    Label30.Visible = True 'Precnik gonjene remenice
    Label34.Visible = True 'Precnik gonjene remenice
    
End If

End Sub

Private Sub Option2_Click()

'Nezavisno modeliranje remenice

Combo7.Enabled = True 'Profil kaisa
Combo8.Enabled = True 'Tip kaisa
Text7.Enabled = True 'Precnik pogonske remenice

Option4.Visible = False 'Precnik gonjene remenice
Option3.Visible = False 'Precnik gonjene remenice
Text8.Visible = False 'Precnik gonjene remenice
Label31.Visible = False 'Precnik gonjene remenice
Label29.Visible = False 'Precnik gonjene remenice
Label30.Visible = False 'Precnik gonjene remenice
Label34.Visible = False 'Precnik gonjene remenice

Label71(1).Visible = True 'Broj zlebova
Text9.Visible = True 'Broj zlebova

Combo7.Clear 'Profil kaisa - ocisti vrednosti ako ih ima
Combo7.AddItem "Normalni"
Combo7.AddItem "Uski"
Combo8.Clear
Text7.Text = "0.000"
Text9.Text = "0"

End Sub

Private Sub Option3_Click()

'Aktiviranje polja za unos d1
Text8.BackColor = &H8000000F
Text8.Enabled = False
Text8.Text = ""
Text7.BackColor = &HFFFFFF
Text7.Enabled = True
Text7.Text = ""

If Option1.Value = True Then
Text7.Enabled = False
Text7.Text = Combo6.Text
End If

End Sub

Private Sub Option4_Click()

'Aktiviranje polja za unos d2
Text7.BackColor = &H8000000F
Text7.Enabled = False
Text7.Text = ""
Text8.BackColor = &HFFFFFF
Text8.Enabled = True
Text8.Text = ""

If Option1.Value = True Then
Text8.Enabled = False
Text8.Text = Text6.Text
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

'------------------------------------------------------------
'Select Case struktura koja sluzi za unos numerickih podataka
'------------------------------------------------------------

Select Case KeyAscii
Case 49 To 57 'Unos brojeva (1,2,3,4,5,6,7,8,9)
Exit Sub

Case 48 'Unos nule
If Text1.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos nule na prvom mestu
Exit Sub

Case 8 'Dozvoljen unos "Backspace"
Exit Sub

Case 46 'Dozvoljen unos tacke
If Text1.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos na prvom mestu
If InStr(1, Text1.Text, ".") > 0 Then KeyAscii = 0 'Dozvoljen unos samo jednom
Exit Sub

Case Else 'Nije dozvoljen unos u bilo kojem drugom slucaju
KeyAscii = 0
End Select

End Sub

Private Sub Text2_Click()
Clipboard.Clear

'-----------------------------------
'Reset unetih vrednosti i proracuna
'-----------------------------------
Text2.Text = ""
Combo4.ListIndex = -1 'Profil kaisa
Combo5.Clear 'Tip kaisa
Combo6.Clear 'Precnik pogonske remenice d1
Text6.Text = "" 'Osno rastojanje
Option2.Value = True

Label1.Caption = "0.000" 'Racunska vrednost gonjene remenice
Label52.Caption = "0.000" 'Stvarna vrednost prenosnog odnosa
Label36.Caption = "0.000" 'Maksimalna obimna brzina
Label79.Caption = "0.000" 'Racunska vrednost duzine kaisa
Label89.Caption = "0.000" 'Standardna vrednost duzine kaisa
Label58.Caption = "0.000" 'Stvarna vrednost osnog rastojanja
Label67.Caption = "0.000" 'Racunska vrednost broja zlebova
Label73.Caption = "0.000" 'Broj zlebova
Label33.Caption = "0.000" 'Obimna brzina
Label33.BackColor = &H8000000F 'Siva boja
Label42.Caption = "0.000" 'Ucestanost savijanja
Label45.Caption = "0.000" 'Dozvoljena ucestanost savijanja
Label42.BackColor = &H8000000F 'Siva boja
Label73.BackColor = &H8000000F 'Siva boja
Label75.Visible = False ' Sakrivanje Label-e za prikaz amin i amax
Command1.Visible = False 'Sakrivanje dugmeta za izbor jaceg profila

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

'------------------------------------------------------------
'Select Case struktura koja sluzi za unos numerickih podataka
'------------------------------------------------------------

Select Case KeyAscii
Case 48 To 57 'Unos brojeva (0,1,2,3,4,5,6,7,8,9)
Exit Sub

Case 8 'Dozvoljen unos "Backspace"
Exit Sub

Case 46 'Dozvoljen unos tacke
If Text2.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos na prvom mestu
If InStr(1, Text2.Text, ".") > 0 Then KeyAscii = 0 'Dozvoljen unos samo jednom
Exit Sub

Case Else 'Nije dozvoljen unos u bilo kojem drugom slucaju
KeyAscii = 0
End Select

End Sub

Private Sub Text3_Click()
Clipboard.Clear

'-----------------------------------
'Reset unetih vrednosti i proracuna
'-----------------------------------
Text3.Text = ""
Combo4.ListIndex = -1 'Profil kaisa
Combo5.Clear 'Tip kaisa
Combo6.Clear 'Precnik pogonske remenice d1
Text6.Text = "" 'Osno rastojanje

Label1.Caption = "0.000" 'Racunska vrednost gonjene remenice
Label52.Caption = "0.000" 'Stvarna vrednost prenosnog odnosa
Label36.Caption = "0.000" 'Maksimalna obimna brzina
Label79.Caption = "0.000" 'Racunska vrednost duzine kaisa
Label89.Caption = "0.000" 'Standardna vrednost duzine kaisa
Label58.Caption = "0.000" 'Stvarna vrednost osnog rastojanja
Label67.Caption = "0.000" 'Racunska vrednost broja zlebova
Label73.Caption = "0.000" 'Broj zlebova
Label33.Caption = "0.000" 'Obimna brzina
Label33.BackColor = &H8000000F 'Siva boja
Label42.Caption = "0.000" 'Ucestanost savijanja
Label45.Caption = "0.000" 'Dozvoljena ucestanost savijanja
Label42.BackColor = &H8000000F 'Siva boja
Label73.BackColor = &H8000000F 'Siva boja
Label75.Visible = False ' Sakrivanje Label-e za prikaz amin i amax
Command1.Visible = False 'Sakrivanje dugmeta za izbor jaceg profila
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

'------------------------------------------------------------
'Select Case struktura koja sluzi za unos numerickih podataka
'------------------------------------------------------------

Select Case KeyAscii
Case 48 To 57 'Unos brojeva (1,2,3,4,5,6,7,8,9)
Exit Sub

Case 8 'Dozvoljen unos "Backspace"
Exit Sub

Case 46 'Dozvoljen unos tacke
If Text3.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos na prvom mestu
If InStr(1, Text3.Text, ".") > 0 Then KeyAscii = 0 'Dozvoljen unos samo jednom
Exit Sub

Case Else 'Nije dozvoljen unos u bilo kojem drugom slucaju
KeyAscii = 0
End Select

End Sub

Private Sub Text4_Click()
Clipboard.Clear

'-----------------------------------
'Reset unetih vrednosti i proracuna
'-----------------------------------
Text4.Text = ""
Combo4.ListIndex = -1 'Profil kaisa
Combo5.Clear 'Tip kaisa
Combo6.Clear 'Precnik pogonske remenice d1
Text6.Text = "" 'Osno rastojanje

Label1.Caption = "0.000" 'Racunska vrednost gonjene remenice
Label52.Caption = "0.000" 'Stvarna vrednost prenosnog odnosa
Label36.Caption = "0.000" 'Maksimalna obimna brzina
Label79.Caption = "0.000" 'Racunska vrednost duzine kaisa
Label89.Caption = "0.000" 'Standardna vrednost duzine kaisa
Label58.Caption = "0.000" 'Stvarna vrednost osnog rastojanja
Label67.Caption = "0.000" 'Racunska vrednost broja zlebova
Label73.Caption = "0.000" 'Broj zlebova
Label33.Caption = "0.000" 'Obimna brzina
Label33.BackColor = &H8000000F 'Siva boja
Label42.Caption = "0.000" 'Ucestanost savijanja
Label45.Caption = "0.000" 'Dozvoljena ucestanost savijanja
Label42.BackColor = &H8000000F 'Siva boja
Label73.BackColor = &H8000000F 'Siva boja
Label75.Visible = False ' Sakrivanje Label-e za prikaz amin i amax
Command1.Visible = False 'Sakrivanje dugmeta za izbor jaceg profila
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

'------------------------------------------------------------
'Select Case struktura koja sluzi za unos numerickih podataka
'------------------------------------------------------------

Select Case KeyAscii
Case 49 To 57 'Unos brojeva (1,2,3,4,5,6,7,8,9)
Exit Sub

Case 48 'Unos nule
If Text4.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos nule na prvom mestu
Exit Sub

Case 8 'Dozvoljen unos "Backspace"
Exit Sub

Case 46 'Dozvoljen unos tacke
If Text4.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos na prvom mestu
If InStr(1, Text4.Text, ".") > 0 Then KeyAscii = 0 'Dozvoljen unos samo jednom
Exit Sub

Case Else 'Nije dozvoljen unos u bilo kojem drugom slucaju
KeyAscii = 0
End Select

End Sub

Private Sub Text1_Click()
Clipboard.Clear
End Sub
Private Sub Combo1_Change()
'--------------------------------------------
'Definisanje faktora radnih uslova Ka
'POCETAK PETLJE
'--------------------------------------------
If Combo3.Text = "Ostale masine" Then
    Text1.Enabled = True
Else
    Text1.Enabled = False
    Text1.Text = ""
End If

If Combo1.Text = "AC/DC elektromotor sa normalnim polaznim momentom" Or Combo1.Text = "SUS motor i turbine sa n>600 min-1" Then '1
    If Combo2.Text = "Do 10 h" Then '2
        If Combo3.Text = "Turbopumpe" Or Combo3.Text = "Turbokompresori" Or _
        Combo3.Text = "Trakasti kompresori za lake materijale" Or Combo3.Text = "Ventilatori i pumpe do 7.4 kW" Or _
        Combo3.Text = "Ravnomerno optereceni strugovi" Or Combo3.Text = "Ravnomerno opterecene brusilice" Then '3
            Text1.Text = "1.0"
            Ka = 1
        ElseIf Combo3.Text = "Makaze za lim" Or Combo3.Text = "Prese" Or _
        Combo3.Text = "Lancani transporteri za teske materijale" Or Combo3.Text = "Trakasti transporteri za teske materijale" Or _
        Combo3.Text = "Elektro generatori" Or Combo3.Text = "Alatne masine" Or _
        Combo3.Text = "Stamparske masine" Or Combo3.Text = "Ventilatori i pumpe preko 7.4 kW" Then '3
            Text1.Text = "1.1"
            Ka = 1.1
        ElseIf Combo3.Text = "Mlinovi" Or Combo3.Text = "Klipni kompresori" Or _
        Combo3.Text = "Puzni kompresori" Or Combo3.Text = "Tekstilne masine" Or _
        Combo3.Text = "Masine za hartiju" Or Combo3.Text = "Testere" Or _
        Combo3.Text = "Gateri" Or Combo3.Text = "Prese za briket" Then '3
            Text1.Text = "1.2"
            Ka = 1.2
        ElseIf Combo3.Text = "Jako optereceni mlinovi" Or Combo3.Text = "Drobilice" Or _
        Combo3.Text = "Mesalice" Or Combo3.Text = "Kalenderi" Or _
        Combo3.Text = "Vitla" Or Combo3.Text = "Kranovi" Or _
        Combo3.Text = "Bageri" Or Combo3.Text = "Valjaonicki stanovi" Then '3
            Text1.Text = "1.3"
            Ka = 1.3
        End If '3
    
    ElseIf Combo2.Text = "Iznad 10 h a manje od 16 h" Then '2
        If Combo3.Text = "Turbopumpe" Or Combo3.Text = "Turbokompresori" Or _
        Combo3.Text = "Trakasti kompresori za lake materijale" Or Combo3.Text = "Ventilatori i pumpe do 7.4 kW" Or _
        Combo3.Text = "Ravnomerno optereceni strugovi" Or Combo3.Text = "Ravnomerno opterecene brusilice" Then '3
            Text1.Text = "1.1"
            Ka = 1.1
        ElseIf Combo3.Text = "Makaze za lim" Or Combo3.Text = "Prese" Or _
        Combo3.Text = "Lancani transporteri za teske materijale" Or Combo3.Text = "Trakasti transporteri za teske materijale" Or _
        Combo3.Text = "Elektro generatori" Or Combo3.Text = "Alatne masine" Or _
        Combo3.Text = "Stamparske masine" Or Combo3.Text = "Ventilatori i pumpe preko 7.4 kW" Then '3
            Text1.Text = "1.2"
            Ka = 1.2
        ElseIf Combo3.Text = "Mlinovi" Or Combo3.Text = "Klipni kompresori" Or _
        Combo3.Text = "Puzni kompresori" Or Combo3.Text = "Tekstilne masine" Or _
        Combo3.Text = "Masine za hartiju" Or Combo3.Text = "Testere" Or _
        Combo3.Text = "Gateri" Or Combo3.Text = "Prese za briket" Then '3
            Text1.Text = "1.3"
            Ka = 1.3
        ElseIf Combo3.Text = "Jako optereceni mlinovi" Or Combo3.Text = "Drobilice" Or _
        Combo3.Text = "Mesalice" Or Combo3.Text = "Kalenderi" Or _
        Combo3.Text = "Vitla" Or Combo3.Text = "Kranovi" Or _
        Combo3.Text = "Bageri" Or Combo3.Text = "Valjaonicki stanovi" Then '3
            Text1.Text = "1.4"
            Ka = 1.4
        End If '3
    
    ElseIf Combo2.Text = "Iznad 16 h" Then '2
        If Combo3.Text = "Turbopumpe" Or Combo3.Text = "Turbokompresori" Or _
        Combo3.Text = "Trakasti kompresori za lake materijale" Or Combo3.Text = "Ventilatori i pumpe do 7.4 kW" Or _
        Combo3.Text = "Ravnomerno optereceni strugovi" Or Combo3.Text = "Ravnomerno opterecene brusilice" Then '3
            Text1.Text = "1.2"
            Ka = 1.2
        ElseIf Combo3.Text = "Makaze za lim" Or Combo3.Text = "Prese" Or _
        Combo3.Text = "Lancani transporteri za teske materijale" Or Combo3.Text = "Trakasti transporteri za teske materijale" Or _
        Combo3.Text = "Elektro generatori" Or Combo3.Text = "Alatne masine" Or _
        Combo3.Text = "Stamparske masine" Or Combo3.Text = "Ventilatori i pumpe preko 7.4 kW" Then '3
            Text1.Text = "1.3"
            Ka = 1.3
        ElseIf Combo3.Text = "Mlinovi" Or Combo3.Text = "Klipni kompresori" Or _
        Combo3.Text = "Puzni kompresori" Or Combo3.Text = "Tekstilne masine" Or _
        Combo3.Text = "Masine za hartiju" Or Combo3.Text = "Testere" Or _
        Combo3.Text = "Gateri" Or Combo3.Text = "Prese za briket" Then '3
            Text1.Text = "1.4"
            Ka = 1.4
        ElseIf Combo3.Text = "Jako optereceni mlinovi" Or Combo3.Text = "Drobilice" Or _
        Combo3.Text = "Mesalice" Or Combo3.Text = "Kalenderi" Or _
        Combo3.Text = "Vitla" Or Combo3.Text = "Kranovi" Or _
        Combo3.Text = "Bageri" Or Combo3.Text = "Valjaonicki stanovi" Then '3
            Text1.Text = "1.5"
            Ka = 1.5
        End If '3
    End If '2

Else '1
    If Combo2.Text = "Do 10 h" Then '2
        If Combo3.Text = "Turbopumpe" Or Combo3.Text = "Turbokompresori" Or _
        Combo3.Text = "Trakasti kompresori za lake materijale" Or Combo3.Text = "Ventilatori i pumpe do 7.4 kW" Or _
        Combo3.Text = "Ravnomerno optereceni strugovi" Or Combo3.Text = "Ravnomerno opterecene brusilice" Then '3
            Text1.Text = "1.1"
            Ka = 1.1
        ElseIf Combo3.Text = "Makaze za lim" Or Combo3.Text = "Prese" Or _
        Combo3.Text = "Lancani transporteri za teske materijale" Or Combo3.Text = "Trakasti transporteri za teske materijale" Or _
        Combo3.Text = "Elektro generatori" Or Combo3.Text = "Alatne masine" Or _
        Combo3.Text = "Stamparske masine" Or Combo3.Text = "Ventilatori i pumpe preko 7.4 kW" Then '3
            Text1.Text = "1.2"
            Ka = 1.2
        ElseIf Combo3.Text = "Mlinovi" Or Combo3.Text = "Klipni kompresori" Or _
        Combo3.Text = "Puzni kompresori" Or Combo3.Text = "Tekstilne masine" Or _
        Combo3.Text = "Masine za hartiju" Or Combo3.Text = "Testere" Or _
        Combo3.Text = "Gateri" Or Combo3.Text = "Prese za briket" Then '3
            Text1.Text = "1.4"
            Ka = 1.4
        ElseIf Combo3.Text = "Jako optereceni mlinovi" Or Combo3.Text = "Drobilice" Or _
        Combo3.Text = "Mesalice" Or Combo3.Text = "Kalenderi" Or _
        Combo3.Text = "Vitla" Or Combo3.Text = "Kranovi" Or _
        Combo3.Text = "Bageri" Or Combo3.Text = "Valjaonicki stanovi" Then '3
            Text1.Text = "1.5"
            Ka = 1.5
        End If '3
    
    ElseIf Combo2.Text = "Iznad 10 h a manje od 16 h" Then '2
        If Combo3.Text = "Turbopumpe" Or Combo3.Text = "Turbokompresori" Or _
        Combo3.Text = "Trakasti kompresori za lake materijale" Or Combo3.Text = "Ventilatori i pumpe do 7.4 kW" Or _
        Combo3.Text = "Ravnomerno optereceni strugovi" Or Combo3.Text = "Ravnomerno opterecene brusilice" Then '3
            Text1.Text = "1.2"
            Ka = 1.2
        ElseIf Combo3.Text = "Makaze za lim" Or Combo3.Text = "Prese" Or _
        Combo3.Text = "Lancani transporteri za teske materijale" Or Combo3.Text = "Trakasti transporteri za teske materijale" Or _
        Combo3.Text = "Elektro generatori" Or Combo3.Text = "Alatne masine" Or _
        Combo3.Text = "Stamparske masine" Or Combo3.Text = "Ventilatori i pumpe preko 7.4 kW" Then '3
            Text1.Text = "1.3"
            Ka = 1.3
        ElseIf Combo3.Text = "Mlinovi" Or Combo3.Text = "Klipni kompresori" Or _
        Combo3.Text = "Puzni kompresori" Or Combo3.Text = "Tekstilne masine" Or _
        Combo3.Text = "Masine za hartiju" Or Combo3.Text = "Testere" Or _
        Combo3.Text = "Gateri" Or Combo3.Text = "Prese za briket" Then '3
            Text1.Text = "1.5"
            Ka = 1.5
        ElseIf Combo3.Text = "Jako optereceni mlinovi" Or Combo3.Text = "Drobilice" Or _
        Combo3.Text = "Mesalice" Or Combo3.Text = "Kalenderi" Or _
        Combo3.Text = "Vitla" Or Combo3.Text = "Kranovi" Or _
        Combo3.Text = "Bageri" Or Combo3.Text = "Valjaonicki stanovi" Then '3
            Text1.Text = "1.6"
            Ka = 1.6
        End If '3
    
    ElseIf Combo2.Text = "Iznad 16 h" Then '2
        If Combo3.Text = "Turbopumpe" Or Combo3.Text = "Turbokompresori" Or _
        Combo3.Text = "Trakasti kompresori za lake materijale" Or Combo3.Text = "Ventilatori i pumpe do 7.4 kW" Or _
        Combo3.Text = "Ravnomerno optereceni strugovi" Or Combo3.Text = "Ravnomerno opterecene brusilice" Then '3
            Text1.Text = "1.3"
            Ka = 1.3
        ElseIf Combo3.Text = "Makaze za lim" Or Combo3.Text = "Prese" Or _
        Combo3.Text = "Lancani transporteri za teske materijale" Or Combo3.Text = "Trakasti transporteri za teske materijale" Or _
        Combo3.Text = "Elektro generatori" Or Combo3.Text = "Alatne masine" Or _
        Combo3.Text = "Stamparske masine" Or Combo3.Text = "Ventilatori i pumpe preko 7.4 kW" Then '3
            Text1.Text = "1.4"
            Ka = 1.4
        ElseIf Combo3.Text = "Mlinovi" Or Combo3.Text = "Klipni kompresori" Or _
        Combo3.Text = "Puzni kompresori" Or Combo3.Text = "Tekstilne masine" Or _
        Combo3.Text = "Masine za hartiju" Or Combo3.Text = "Testere" Or _
        Combo3.Text = "Gateri" Or Combo3.Text = "Prese za briket" Then '3
            Text1.Text = "1.6"
            Ka = 1.6
        ElseIf Combo3.Text = "Jako optereceni mlinovi" Or Combo3.Text = "Drobilice" Or _
        Combo3.Text = "Mesalice" Or Combo3.Text = "Kalenderi" Or _
        Combo3.Text = "Vitla" Or Combo3.Text = "Kranovi" Or _
        Combo3.Text = "Bageri" Or Combo3.Text = "Valjaonicki stanovi" Then '3
            Text1.Text = "1.8"
            Ka = 1.8
        End If '3
    End If '2
End If
'-------------
'KRAJ PETLJE
'-------------
End Sub

Private Sub Combo1_Click()
Call Combo1_Change 'Pozivanje funkcije
End Sub

Private Sub Combo2_Change()
Call Combo1_Change 'Pozivanje funkcije
End Sub

Private Sub Combo2_Click()
Call Combo1_Change 'Pozivanje funkcije
End Sub

Private Sub Combo3_Change()
Call Combo1_Change 'Pozivanje funkcije
End Sub

Private Sub Combo3_Click()
Call Combo1_Change 'Pozivanje funkcije
End Sub

Private Sub Text6_Click()
Text6.Text = ""
Clipboard.Clear
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

'------------------------------------------------------------
'Select Case struktura koja sluzi za unos numerickih podataka
'------------------------------------------------------------
Clipboard.Clear

Select Case KeyAscii
Case 49 To 57 'Unos brojeva (1,2,3,4,5,6,7,8,9)
Exit Sub

Case 48 'Unos nule
If Text6.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos nule na prvom mestu
Exit Sub

Case 8 'Dozvoljen unos "Backspace"
Exit Sub

Case 46 'Dozvoljen unos tacke
If Text6.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos na prvom mestu
If InStr(1, Text6.Text, ".") > 0 Then KeyAscii = 0 'Dozvoljen unos samo jednom
Exit Sub

Case Else 'Nije dozvoljen unos u bilo kojem drugom slucaju
KeyAscii = 0
End Select

End Sub

Private Sub Text7_Change()
Call Text7_Click
End Sub

Private Sub Text7_Click()
Clipboard.Clear

'--------
'alfa_CAD
'--------

If Combo8.Text = "SPZ - DIN 7753: 1988" Then '1

        If Val(Text7.Text) <= 80 Then '1.1
        alfa_CAD = 34
        Else: alfa_CAD = 38
        End If '1.1
        
    ElseIf Combo8.Text = "SPA - DIN 7753: 1988" Then '2
    
            If Val(Text7.Text) <= 118 Then '2.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '2.1

    ElseIf Combo8.Text = "SPB - DIN 7753: 1988" Then '3

            If Val(Text7.Text) <= 190 Then '3.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '3.1
        
    ElseIf Combo8.Text = "SPC - DIN 7753: 1988" Then '4

            If Val(Text7.Text) <= 315 Then '4.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '4.1

    ElseIf Combo8.Text = "Z/10 - ISO 4184: 1992 / DIN 2215: 1998" Then '5

            If Val(Text7.Text) <= 80 Then '5.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '5.1
        
    ElseIf Combo8.Text = "A/13 - ISO 4184: 1992 / DIN 2215: 1998" Then '6

            If Val(Text7.Text) <= 118 Then '6.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '6.1
        
    ElseIf Combo8.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998" Then '7

            If Val(Text7.Text) <= 190 Then '7.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '7.1
        
    ElseIf Combo8.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" Then '8

            If Val(Text7.Text) <= 315 Then '8.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '8.1
End If '1

Label43(37).Caption = alfa_CAD

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)

'------------------------------------------------------------
'Select Case struktura koja sluzi za unos numerickih podataka
'------------------------------------------------------------

Select Case KeyAscii
Case 49 To 57 'Unos brojeva (1,2,3,4,5,6,7,8,9)
Exit Sub

Case 48 'Unos nule
If Text7.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos nule na prvom mestu
Exit Sub

Case 8 'Dozvoljen unos "Backspace"
Exit Sub

Case 46 'Dozvoljen unos tacke
If Text7.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos na prvom mestu
If InStr(1, Text7.Text, ".") > 0 Then KeyAscii = 0 'Dozvoljen unos samo jednom
Exit Sub

Case Else 'Nije dozvoljen unos u bilo kojem drugom slucaju
KeyAscii = 0
End Select

Call Text7_Click

End Sub

Private Sub Text8_Change()
Call Text8_Click
End Sub

Private Sub Text8_Click()
Clipboard.Clear

'--------
'alfa_CAD
'--------

If Combo8.Text = "SPZ - DIN 7753: 1988" Then '1

        If Val(Text8.Text) <= 80 Then '1.1
        alfa_CAD = 34
        Else: alfa_CAD = 38
        End If '1.1
        
    ElseIf Combo8.Text = "SPA - DIN 7753: 1988" Then '2
    
            If Val(Text8.Text) <= 118 Then '2.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '2.1

    ElseIf Combo8.Text = "SPB - DIN 7753: 1988" Then '3

            If Val(Text8.Text) <= 190 Then '3.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '3.1
        
    ElseIf Combo8.Text = "SPC - DIN 7753: 1988" Then '4

            If Val(Text8.Text) <= 315 Then '4.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '4.1

    ElseIf Combo8.Text = "Z/10 - ISO 4184: 1992 / DIN 2215: 1998" Then '5

            If Val(Text8.Text) <= 80 Then '5.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '5.1
        
    ElseIf Combo8.Text = "A/13 - ISO 4184: 1992 / DIN 2215: 1998" Then '6

            If Val(Text8.Text) <= 118 Then '6.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '6.1
        
    ElseIf Combo8.Text = "B/17 - ISO 4184: 1992 / DIN 2215: 1998" Then '7

            If Val(Text8.Text) <= 190 Then '7.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '7.1
        
    ElseIf Combo8.Text = "C/22 - ISO 4184: 1992 / DIN 2215: 1998" Then '8

            If Val(Text8.Text) <= 315 Then '8.1
            alfa_CAD = 34
            Else: alfa_CAD = 38
            End If '8.1
End If '1

Label43(37).Caption = alfa_CAD

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)

'------------------------------------------------------------
'Select Case struktura koja sluzi za unos numerickih podataka
'------------------------------------------------------------

Select Case KeyAscii
Case 49 To 57 'Unos brojeva (1,2,3,4,5,6,7,8,9)
Exit Sub

Case 48 'Unos nule
If Text8.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos nule na prvom mestu
Exit Sub

Case 8 'Dozvoljen unos "Backspace"
Exit Sub

Case 46 'Dozvoljen unos tacke
If Text8.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos na prvom mestu
If InStr(1, Text8.Text, ".") > 0 Then KeyAscii = 0 'Dozvoljen unos samo jednom
Exit Sub

Case Else 'Nije dozvoljen unos u bilo kojem drugom slucaju
KeyAscii = 0
End Select

Call Text8_Click

End Sub

Private Sub Text9_Click()
Text9.Text = ""
Clipboard.Clear
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)

'------------------------------------------------------------
'Select Case struktura koja sluzi za unos numerickih podataka
'------------------------------------------------------------

Select Case KeyAscii
Case 49 To 57 'Unos brojeva (1,2,3,4,5,6,7,8,9)
Exit Sub

Case 48 'Unos nule
If Text9.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos nule na prvom mestu
Exit Sub

Case 8 'Dozvoljen unos "Backspace"
Exit Sub

Case 46 'Dozvoljen unos tacke
If Text9.SelStart = 0 Then KeyAscii = 0 'Nije dozvoljen unos na prvom mestu
If InStr(1, Text9.Text, ".") > 0 Then KeyAscii = 0 'Dozvoljen unos samo jednom
Exit Sub

Case Else 'Nije dozvoljen unos u bilo kojem drugom slucaju
KeyAscii = 0
End Select

End Sub
