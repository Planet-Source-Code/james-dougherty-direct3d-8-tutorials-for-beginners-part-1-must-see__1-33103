VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11490
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   498
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6030
      Index           =   8
      Left            =   -9120
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   402
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   124
      Top             =   8280
      Width           =   10815
      Begin VB.Label Label109 
         BackStyle       =   0  'Transparent
         Caption         =   "Thanks   -James-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5040
         TabIndex        =   128
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label108 
         BackStyle       =   0  'Transparent
         Caption         =   "I hope this tutorial helped you understand              Direct3D better and its basic uses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3120
         TabIndex        =   127
         Top             =   3120
         Width           =   5055
      End
      Begin VB.Label Label107 
         BackStyle       =   0  'Transparent
         Caption         =   "In Part 2 we will discuss Textures, Materials,          Lighting, and create our first object"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3120
         TabIndex        =   126
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Label Label118 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is the end of Part 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2400
         TabIndex        =   125
         Top             =   600
         Width           =   6285
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6030
      Index           =   7
      Left            =   -8760
      Picture         =   "frmMain.frx":0D48
      ScaleHeight     =   402
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   105
      Top             =   8520
      Width           =   10815
      Begin VB.Label lblDemo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Demo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9930
         TabIndex        =   129
         Top             =   5580
         Width           =   480
      End
      Begin VB.Image imgDemo 
         Height          =   360
         Left            =   9570
         Picture         =   "frmMain.frx":1786
         Stretch         =   -1  'True
         Top             =   5520
         Width           =   1170
      End
      Begin VB.Label Label106 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   123
         Top             =   3720
         Width           =   3705
      End
      Begin VB.Label Label96 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DDevice.EndScene"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   122
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label95 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "'The objects would be rendered right here"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   1080
         TabIndex        =   121
         Top             =   3240
         Width           =   3525
      End
      Begin VB.Label Label94 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DDevice.BeginScene"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   120
         Top             =   3000
         Width           =   1860
      End
      Begin VB.Label Label93 
         BackStyle       =   0  'Transparent
         Caption         =   "EndLoop = True and the loop will end, Cleanup DX, and unload the form."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   119
         Top             =   5040
         Width           =   5895
      End
      Begin VB.Label Label92 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: EndLoop is defined as public and under Form_Click or whatever you would simply put"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   118
         Top             =   4800
         Width           =   7575
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2640
         TabIndex        =   117
         Top             =   3960
         Width           =   390
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loop Until"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   840
         TabIndex        =   116
         Top             =   3960
         Width           =   825
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   840
         TabIndex        =   115
         Top             =   2280
         Width           =   210
      End
      Begin VB.Label Label105 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Render(ClearColor as long)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1620
         TabIndex        =   114
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label104 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Public Sub"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   720
         TabIndex        =   113
         Top             =   2040
         Width           =   855
      End
      Begin VB.Shape Shape7 
         Height          =   2655
         Left            =   480
         Top             =   1920
         Width           =   7815
      End
      Begin VB.Label Label103 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Making A Basic Rendering Loop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   112
         Top             =   120
         Width           =   3675
      End
      Begin VB.Label Label102 
         BackStyle       =   0  'Transparent
         Caption         =   "Now that know how to initialize Direct3D and clean it up lets make a rendering loop as our final step."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   111
         Top             =   720
         Width           =   9735
      End
      Begin VB.Label Label101 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":1D34
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   110
         Top             =   1200
         Width           =   9735
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DoEvents "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   109
         Top             =   2520
         Width           =   810
      End
      Begin VB.Label Label99 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, ClearColor, 1, 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   108
         Top             =   2760
         Width           =   7035
      End
      Begin VB.Label Label98 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " EndLoop = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1680
         TabIndex        =   107
         Top             =   3960
         Width           =   945
      End
      Begin VB.Label Label97 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Sub"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   720
         TabIndex        =   106
         Top             =   4200
         Width           =   660
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6030
      Index           =   6
      Left            =   -8880
      Picture         =   "frmMain.frx":1DFD
      ScaleHeight     =   402
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   86
      Top             =   8280
      Width           =   10815
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nothing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1590
         TabIndex        =   104
         Top             =   3000
         Width           =   630
      End
      Begin VB.Label Label87 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nothing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1680
         TabIndex        =   103
         Top             =   2760
         Width           =   630
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nothing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2280
         TabIndex        =   102
         Top             =   2520
         Width           =   630
      End
      Begin VB.Label Label85 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nothing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1800
         TabIndex        =   101
         Top             =   2280
         Width           =   630
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   840
         TabIndex        =   100
         Top             =   2520
         Width           =   270
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   840
         TabIndex        =   99
         Top             =   2760
         Width           =   270
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   840
         TabIndex        =   98
         Top             =   3000
         Width           =   270
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   840
         TabIndex        =   97
         Top             =   2280
         Width           =   270
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Sub"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   720
         TabIndex        =   96
         Top             =   3240
         Width           =   660
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DX = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1200
         TabIndex        =   95
         Top             =   3000
         Width           =   390
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3D = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1200
         TabIndex        =   94
         Top             =   2760
         Width           =   480
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DDevice = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1200
         TabIndex        =   93
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "All we need to do is release all the objects we created from memory. Heres how the code will look."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   92
         Top             =   1560
         Width           =   9735
      End
      Begin VB.Label Label79 
         BackStyle       =   0  'Transparent
         Caption         =   "Now that know how to initialize Direct3D we have to know how to clean it up also."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   91
         Top             =   720
         Width           =   9735
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cleaning Up Direct3D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   90
         Top             =   120
         Width           =   2460
      End
      Begin VB.Shape Shape6 
         Height          =   1695
         Left            =   480
         Top             =   1920
         Width           =   6735
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Public Sub"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   720
         TabIndex        =   89
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cleanup_DX()"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1620
         TabIndex        =   88
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DX = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1200
         TabIndex        =   87
         Top             =   2280
         Width           =   585
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6030
      Index           =   5
      Left            =   -9360
      Picture         =   "frmMain.frx":283B
      ScaleHeight     =   402
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   66
      Top             =   8040
      Width           =   10815
      Begin VB.Label lblCode1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4020
         Index           =   1
         Left            =   10800
         TabIndex        =   67
         Top             =   1365
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":3279
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   720
         TabIndex        =   85
         Top             =   5400
         Width           =   6345
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DPP.hDeviceWindow = hWnd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   720
         TabIndex        =   84
         Top             =   5160
         Width           =   2505
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DPP.AutoDepthStencilFormat = D3DFMT_D16"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   720
         TabIndex        =   83
         Top             =   4920
         Width           =   3765
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DPP.EnableAutoDepthStencil = 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   720
         TabIndex        =   82
         Top             =   4680
         Width           =   2790
      End
      Begin VB.Shape Shape5 
         Height          =   1335
         Left            =   480
         Top             =   4560
         Width           =   6735
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Finally we want DX to handle the depth stencil at 16 bit. Then we set the handle to our winow and create our device."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   81
         Top             =   4320
         Width           =   9855
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DPP.BackBufferCount = 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   720
         TabIndex        =   80
         Top             =   3810
         Width           =   2220
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":3315
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   79
         Top             =   2700
         Width           =   9855
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   9675
         TabIndex        =   68
         Top             =   5460
         Width           =   945
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Initializing Direct3D  2 of 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   78
         Top             =   120
         Width           =   2940
      End
      Begin VB.Label Label76 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":33B4
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   77
         Top             =   600
         Width           =   9735
      End
      Begin VB.Label Label75 
         BackStyle       =   0  'Transparent
         Caption         =   "First we create the D3DPRESENT_PARAMETERS and then tell DirectX we want a window application."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   76
         Top             =   1560
         Width           =   9855
      End
      Begin VB.Shape Shape4 
         Height          =   735
         Left            =   480
         Top             =   1800
         Width           =   4455
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DPP.Windowed = 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   720
         TabIndex        =   75
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dim"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   720
         TabIndex        =   74
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " D3DPP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1065
         TabIndex        =   73
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "As"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1680
         TabIndex        =   72
         Top             =   1920
         Width           =   225
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DPRESENT_PARAMETERS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1995
         TabIndex        =   71
         Top             =   1920
         Width           =   2190
      End
      Begin VB.Shape Shape3 
         Height          =   975
         Left            =   480
         Top             =   3180
         Width           =   5535
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   720
         TabIndex        =   70
         Top             =   3330
         Width           =   3810
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DPP.BackBufferFormat = Mode.Format"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   720
         TabIndex        =   69
         Top             =   3570
         Width           =   3315
      End
      Begin VB.Image imgCode 
         Height          =   360
         Index           =   1
         Left            =   9555
         Picture         =   "frmMain.frx":355A
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   1170
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6030
      Index           =   4
      Left            =   -9600
      Picture         =   "frmMain.frx":3B08
      ScaleHeight     =   402
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   38
      Top             =   7800
      Width           =   10815
      Begin VB.Label lblCode1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Index           =   0
         Left            =   5805
         TabIndex        =   65
         Top             =   3525
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   9675
         TabIndex        =   64
         Top             =   5460
         Width           =   945
      End
      Begin VB.Image imgCode 
         Height          =   360
         Index           =   0
         Left            =   9555
         Picture         =   "frmMain.frx":4546
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   1170
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   840
         TabIndex        =   63
         Top             =   5400
         Width           =   4800
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dim"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   840
         TabIndex        =   62
         Top             =   5160
         Width           =   315
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1200
         TabIndex        =   61
         Top             =   5160
         Width           =   465
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "As"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1740
         TabIndex        =   60
         Top             =   5160
         Width           =   225
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DDISPLAYMODE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2040
         TabIndex        =   59
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   480
         Top             =   5040
         Width           =   5535
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":4AF4
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   58
         Top             =   4200
         Width           =   9855
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Note: These are called in your InitializeDX sub"
         Height          =   195
         Left            =   510
         TabIndex        =   57
         Top             =   3720
         Width           =   3780
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3D = DX.Direct3DCreate()"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1170
         TabIndex        =   56
         Top             =   3240
         Width           =   2085
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   840
         TabIndex        =   55
         Top             =   3240
         Width           =   270
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DX8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4110
         TabIndex        =   54
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3615
         TabIndex        =   53
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DX = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3120
         TabIndex        =   52
         Top             =   3000
         Width           =   585
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Is Nothing Then Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1440
         TabIndex        =   51
         Top             =   3000
         Width           =   1650
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " D3DX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   50
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   840
         TabIndex        =   49
         Top             =   3000
         Width           =   105
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DirectX8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3840
         TabIndex        =   48
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3360
         TabIndex        =   47
         Top             =   2760
         Width           =   405
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DX = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3000
         TabIndex        =   46
         Top             =   2760
         Width           =   390
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Is Nothing Then Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1320
         TabIndex        =   45
         Top             =   2760
         Width           =   1605
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " DX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   44
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   840
         TabIndex        =   43
         Top             =   2760
         Width           =   105
      End
      Begin VB.Shape Shape1 
         Height          =   1575
         Left            =   480
         Top             =   2400
         Width           =   4455
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Now that we have decided to make a window application, we create our DX, D3DX, and D3D objects. Heres how its done."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   2040
         Width           =   9855
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":4C2D
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   41
         Top             =   1200
         Width           =   9735
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Initializing Direct3D  1 of 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   2940
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Now that we have a basic understanding about objects in Direct3D, lets learn how to initialize Direct3D. "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   720
         Width           =   9735
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6030
      Index           =   3
      Left            =   -10080
      Picture         =   "frmMain.frx":4D29
      ScaleHeight     =   402
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   28
      Top             =   7800
      Width           =   10815
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertex 3"
         Height          =   195
         Left            =   1080
         TabIndex        =   37
         Top             =   5520
         Width           =   585
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertex 1"
         Height          =   195
         Left            =   600
         TabIndex        =   36
         Top             =   2760
         Width           =   585
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertex 2"
         Height          =   195
         Left            =   3840
         TabIndex        =   35
         Top             =   5520
         Width           =   585
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertex 6"
         Height          =   195
         Left            =   4320
         TabIndex        =   34
         Top             =   5280
         Width           =   585
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertex 5"
         Height          =   195
         Left            =   3840
         TabIndex        =   33
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertex 4"
         Height          =   195
         Left            =   840
         TabIndex        =   32
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":5767
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   5160
         TabIndex        =   31
         Top             =   2760
         Width           =   4965
      End
      Begin VB.Image Image1 
         Height          =   3135
         Left            =   1080
         Picture         =   "frmMain.frx":58E7
         Top             =   2520
         Width           =   3345
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":75EE
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   480
         TabIndex        =   30
         Top             =   840
         Width           =   9735
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direct3D's Use Of 3D Objects"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   3315
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6030
      Index           =   2
      Left            =   -10440
      Picture         =   "frmMain.frx":78F2
      ScaleHeight     =   402
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   14
      Top             =   7680
      Width           =   10815
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2835
         Left            =   600
         Picture         =   "frmMain.frx":8330
         ScaleHeight     =   189
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   194
         TabIndex        =   15
         Top             =   1200
         Width           =   2910
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Understanding What A Mesh, A Face, And Vertex Are And Thier Purpose Before Jumping In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   10350
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mesh - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         TabIndex        =   24
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A mesh is a collection of faces formed to make a 3D object in space."
         Height          =   195
         Left            =   4200
         TabIndex        =   23
         Top             =   1560
         Width           =   4845
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Face - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         TabIndex        =   22
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "A face is simply a polygon defined by a collection of points, called vertices(Plural for vertex)."
         Height          =   495
         Left            =   4200
         TabIndex        =   21
         Top             =   2280
         Width           =   5880
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertex(Vertices) - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         TabIndex        =   20
         Top             =   2880
         Width           =   1530
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "A vertex is a point in 3D Space that is used to give an objects face(s) their position, their scale, and also their angle."
         Height          =   495
         Left            =   4200
         TabIndex        =   19
         Top             =   3120
         Width           =   5880
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":B204
         Height          =   975
         Left            =   840
         TabIndex        =   18
         Top             =   4080
         Width           =   2640
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You may be asking, ""So what does this have to do with D3D?"""
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5400
         TabIndex        =   17
         Top             =   5280
         Width           =   5040
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Let's find out, come on I'll explain it to you..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5400
         TabIndex        =   16
         Top             =   5520
         Width           =   3600
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6030
      Index           =   1
      Left            =   -10200
      Picture         =   "frmMain.frx":B2A8
      ScaleHeight     =   402
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   6
      Top             =   7800
      Width           =   10815
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3DX8 Object - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2040
         TabIndex        =   27
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":BCE6
         Height          =   615
         Left            =   2280
         TabIndex        =   26
         Top             =   4080
         Width           =   6720
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":BDEC
         Height          =   855
         Left            =   2280
         TabIndex        =   13
         Top             =   2880
         Width           =   6600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direct3DDevice8 Object - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2040
         TabIndex        =   12
         Top             =   2640
         Width           =   2025
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":BF37
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   2040
         Width           =   6600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direct3D8 Object - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2040
         TabIndex        =   10
         Top             =   1800
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is the core object for DirectX. It gives you access to the DirectX API's and other objects."
         Height          =   195
         Left            =   2280
         TabIndex        =   9
         Top             =   1320
         Width           =   6555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DirectX8 Object- "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2040
         TabIndex        =   8
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Understanding The Direct3D 8 Objects"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4380
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4470
      Index           =   0
      Left            =   2640
      Picture         =   "frmMain.frx":BFEE
      ScaleHeight     =   298
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   436
      TabIndex        =   5
      Top             =   1200
      Width           =   6540
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   11025
      TabIndex        =   4
      Top             =   255
      Width           =   105
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   10950
      Picture         =   "frmMain.frx":F692
      Top             =   240
      Width           =   225
   End
   Begin VB.Label lblPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   630
      TabIndex        =   3
      Top             =   6660
      Width           =   675
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   10845
      TabIndex        =   0
      Top             =   6675
      Width           =   180
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¥"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   10455
      TabIndex        =   1
      Top             =   6675
      Width           =   90
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   9990
      TabIndex        =   2
      Top             =   6675
      Width           =   180
   End
   Begin VB.Image imgPreview 
      Enabled         =   0   'False
      Height          =   360
      Left            =   390
      Picture         =   "frmMain.frx":F976
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1170
   End
   Begin VB.Image imgStep 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   9900
      Picture         =   "frmMain.frx":FF24
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgStep 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   10320
      Picture         =   "frmMain.frx":10390
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgStep 
      Height          =   375
      Index           =   2
      Left            =   10740
      Picture         =   "frmMain.frx":107FC
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgStepDown 
      Height          =   375
      Left            =   1200
      Picture         =   "frmMain.frx":10C68
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image imgStepOver 
      Height          =   375
      Left            =   1200
      Picture         =   "frmMain.frx":110E4
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image imgStepNormal 
      Height          =   375
      Left            =   1200
      Picture         =   "frmMain.frx":1151D
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image imgPreviewNormal 
      Height          =   360
      Left            =   0
      Picture         =   "frmMain.frx":11989
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1170
   End
   Begin VB.Image imgPreviewDown 
      Height          =   360
      Left            =   0
      Picture         =   "frmMain.frx":11F37
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1170
   End
   Begin VB.Image imgPreviewOver 
      Height          =   360
      Left            =   0
      Picture         =   "frmMain.frx":12615
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1170
   End
   Begin VB.Image imgBack 
      Height          =   7470
      Left            =   0
      Picture         =   "frmMain.frx":12C44
      Top             =   0
      Width           =   11490
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private HoldIndex As Integer

Private Sub Form_Load()
Dim i  As Integer

HoldIndex = 0

For i = 1 To 8
 picFrame(i).Visible = False
 picFrame(i).Move 24, 32, 721, 402
Next
picFrame(0).Move 176, 80, 436, 298
picFrame(0).Visible = True
SetupCodeBox
End Sub

Private Sub imgStep_Click(Index As Integer)
On Local Error Resume Next
Dim i As Integer

Select Case Index
 Case 0
  HoldIndex = HoldIndex - 1
  If HoldIndex = 5 Then
   lblPreview.Enabled = True
   imgPreview.Enabled = True
  Else
   lblPreview.Enabled = False
   imgPreview.Enabled = False
  End If
  If imgStep(2).Enabled = False Then imgStep(2).Enabled = True
  If lblStep(2).Enabled = False Then lblStep(2).Enabled = True
  If HoldIndex < 0 Then
   imgStep(0).Enabled = False
   lblStep(0).Enabled = False
   imgStep(1).Enabled = False
   lblStep(1).Enabled = False
   If imgPreview.Enabled Then imgPreview.Enabled = False
   If lblPreview.Enabled Then lblPreview.Enabled = False
   Exit Sub
  End If
  picFrame(HoldIndex + 1).Visible = False
  picFrame(HoldIndex).Visible = True
 Case 1
  imgStep(0).Enabled = False
  lblStep(0).Enabled = False
  imgStep(1).Enabled = False
  lblStep(1).Enabled = False
  imgStep(2).Enabled = True
  lblStep(2).Enabled = True
  If imgPreview.Enabled Then imgPreview.Enabled = False
  If lblPreview.Enabled Then lblPreview.Enabled = False
  For i = 0 To 8
   picFrame(i).Visible = False
  Next
  picFrame(0).Visible = True
  HoldIndex = 0
 Case 2
  HoldIndex = HoldIndex + 1
  If HoldIndex = 5 Then
   lblPreview.Enabled = True
   imgPreview.Enabled = True
  Else
   lblPreview.Enabled = False
   imgPreview.Enabled = False
  End If
  If imgStep(0).Enabled = False Then imgStep(0).Enabled = True
  If lblStep(0).Enabled = False Then lblStep(0).Enabled = True
  If imgStep(1).Enabled = False Then imgStep(1).Enabled = True
  If lblStep(1).Enabled = False Then lblStep(1).Enabled = True
  If HoldIndex > 8 Then
   imgStep(2).Enabled = False
   lblStep(2).Enabled = False
   Exit Sub
  End If
  picFrame(HoldIndex - 1).Visible = False
  picFrame(HoldIndex).Visible = True
End Select
End Sub

Private Sub imgPreview_Click()
If HoldIndex = 5 Then
 SelectedIndex = 0
 frmDX.Show
Else
 Exit Sub
End If
End Sub

'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|
'|¶¶    The code below is all for the skinning,  not    ¶¶|
'|¶¶    very important for this app.                    ¶¶|
'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|

Private Sub imgCode_Click(Index As Integer)
Select Case Index
 Case 0
  lblCode1(0).Move 387, 235
 Case 1
  lblCode1(1).Move 331, 91
End Select
lblCode1(Index).Visible = True
End Sub

Private Sub imgCode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCode(Index).Picture = imgPreviewDown.Picture
End Sub

Private Sub imgCode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCode(Index).Picture = imgPreviewOver.Picture
End Sub

Private Sub imgCode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCode(Index).Picture = imgPreviewNormal.Picture
End Sub

Private Sub lblCode_Click(Index As Integer)
imgCode_Click Index
End Sub

Private Sub lblCode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCode(Index).Picture = imgPreviewDown.Picture
End Sub

Private Sub lblCode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCode(Index).Picture = imgPreviewOver.Picture
End Sub

Private Sub lblCode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCode(Index).Picture = imgPreviewNormal.Picture
End Sub

Private Sub imgDemo_Click()
On Local Error Resume Next
SelectedIndex = 1
frmDX.Show
End Sub

Private Sub imgDemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewDown.Picture
End Sub

Private Sub imgDemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewOver.Picture
End Sub

Private Sub imgDemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewNormal.Picture
End Sub

Private Sub lblDemo_Click()
imgDemo_Click
End Sub

Private Sub lblDemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewDown.Picture
End Sub

Private Sub lblDemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewOver.Picture
End Sub

Private Sub lblDemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewNormal.Picture
End Sub

Private Sub imgPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewDown.Picture
End Sub

Private Sub imgPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewOver.Picture
End Sub

Private Sub imgPreview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewNormal.Picture
End Sub

Private Sub imgStep_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepDown.Picture
End Sub

Private Sub imgStep_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepOver.Picture
End Sub

Private Sub imgStep_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepNormal.Picture
End Sub

Private Sub lblClose_Click()
imgClose_Click
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = &HFF&
End Sub

Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = &H0
End Sub

Private Sub lblPreview_Click()
imgPreview_Click
End Sub

Private Sub lblPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewDown.Picture
End Sub

Private Sub lblPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewOver.Picture
End Sub

Private Sub lblPreview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewNormal.Picture
End Sub

Private Sub lblStep_Click(Index As Integer)
imgStep_Click Index
End Sub

Private Sub lblStep_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepDown.Picture
End Sub

Private Sub lblStep_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepOver.Picture
End Sub

Private Sub lblStep_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepNormal.Picture
End Sub

Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = &H0
imgStep(0).Picture = imgStepNormal.Picture
imgStep(1).Picture = imgStepNormal.Picture
imgStep(2).Picture = imgStepNormal.Picture
imgPreview.Picture = imgPreviewNormal.Picture
imgCode(0).Picture = imgPreviewNormal.Picture
imgCode(1).Picture = imgPreviewNormal.Picture
imgDemo.Picture = imgPreviewNormal.Picture
lblCode1(0).Visible = False
lblCode1(1).Visible = False
End Sub

Private Sub picFrame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = &H0
imgStep(0).Picture = imgStepNormal.Picture
imgStep(1).Picture = imgStepNormal.Picture
imgStep(2).Picture = imgStepNormal.Picture
imgPreview.Picture = imgPreviewNormal.Picture
imgCode(0).Picture = imgPreviewNormal.Picture
imgCode(1).Picture = imgPreviewNormal.Picture
imgDemo.Picture = imgPreviewNormal.Picture
lblCode1(0).Visible = False
lblCode1(1).Visible = False
End Sub

Public Sub SetupCodeBox()
lblCode1(0).Caption = "Public Function Initialize_Win(hWnd As Long) As Boolean" & vbCrLf & _
                      vbCrLf & _
                      "  Dim Mode As D3DDISPLAYMODE" & vbCrLf & _
                      vbCrLf & _
                      "  If DX Is Nothing Then Set DX = New DirectX8" & vbCrLf & _
                      "  If D3DX Is Nothing Then Set D3DX = New D3DX8" & vbCrLf & _
                      "  Set D3D = DX.Direct3DCreate" & vbCrLf & _
                      vbCrLf & _
                      "  D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode" & vbCrLf & _
                      vbCrLf & _
                      "End Function"
                      
lblCode1(1).Caption = "Public Function Initialize_Win(hWnd As Long) As Boolean" & vbCrLf & _
                      vbCrLf & _
                      "  Dim D3DPP As D3DPRESENT_PARAMETERS" & vbCrLf & _
                      "  Dim Mode As D3DDISPLAYMODE" & vbCrLf & _
                      vbCrLf & _
                      "  If DX Is Nothing Then Set DX = New DirectX8" & vbCrLf & _
                      "  If D3DX Is Nothing Then Set D3DX = New D3DX8" & vbCrLf & _
                      "  Set D3D = DX.Direct3DCreate" & vbCrLf & _
                      vbCrLf & _
                      "  D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode" & vbCrLf & _
                      vbCrLf & _
                      "  D3DPP.Windowed = 1" & vbCrLf & _
                      "  D3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD" & vbCrLf & _
                      "  D3DPP.BackBufferFormat = Mode.Format" & vbCrLf & _
                      "  D3DPP.BackBufferCount = 1" & vbCrLf & _
                      "  D3DPP.EnableAutoDepthStencil = 1" & vbCrLf & _
                      "  D3DPP.AutoDepthStencilFormat = D3DFMT_D16" & vbCrLf & _
                      "  D3DPP.hDeviceWindow = hWnd" & vbCrLf & _
                      vbCrLf & _
                      "  Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, _ " & _
                      "  D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)" & vbCrLf & _
                      vbCrLf & _
                      "End Function"
End Sub
