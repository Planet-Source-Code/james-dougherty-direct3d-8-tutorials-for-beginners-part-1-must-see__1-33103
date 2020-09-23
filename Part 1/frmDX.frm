VERSION 5.00
Begin VB.Form frmDX 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Direct3D 8 Tutorial"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   464
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optColor 
      Caption         =   "Blue"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   5
      Top             =   6600
      Width           =   975
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Red"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   6600
      Width           =   975
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Black"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   6600
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox Canvas 
      BackColor       =   &H00000000&
      Height          =   6255
      Left            =   0
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   485
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Label lblInit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D3D Initialized!!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   1755
         TabIndex        =   1
         Top             =   2820
         Width           =   3825
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clear Color - "
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
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   1080
   End
End
Attribute VB_Name = "frmDX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Canvas_KeyDown(KeyCode As Integer, Shift As Integer)
If SelectedIndex = 0 Then
 Unload frmDX
ElseIf SelectedIndex = 1 Then
 EndLoop = True
End If
End Sub

Private Sub Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SelectedIndex = 0 Then
 Unload frmDX
ElseIf SelectedIndex = 1 Then
 EndLoop = True
End If
End Sub

Private Sub Form_Load()
Dim OK As Boolean

If SelectedIndex = 0 Then
 Me.Height = 6585
 OK = Initialize_Win(Canvas.hWnd)
 If Not (OK) Then Cleanup_DX: Unload frmDX
ElseIf SelectedIndex = 1 Then
 Me.Height = 7275
 OK = Initialize_Win(Canvas.hWnd)
 If Not (OK) Then Cleanup_DX: Unload frmDX
 Render
Else
 Cleanup_DX
 Unload frmDX
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Local Error Resume Next
EndLoop = True
Cleanup_DX
Unload frmDX
End Sub

Private Sub optColor_Click(Index As Integer)
On Local Error Resume Next

If SelectedIndex = 1 Then
 Select Case Index
  Case 0
   EndLoop = True
   EndLoop = False
   Render vbBlack
  Case 1
   EndLoop = True
   EndLoop = False
   Render vbBlue 'No this isn't a mistake dx reverses it to red?
  Case 2
   EndLoop = True
   EndLoop = False
   Render vbRed 'No this isn't a mistake dx reverses it to blue?
 End Select
Else
 Exit Sub
End If

End Sub
