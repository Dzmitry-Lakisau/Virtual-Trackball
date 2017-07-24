VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Trackball 
   BackColor       =   &H80000007&
   Caption         =   "Trackball"
   ClientHeight    =   7500
   ClientLeft      =   180
   ClientTop       =   810
   ClientWidth     =   7500
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
      Begin VB.Label Label1 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu SpeedMenu 
      Caption         =   "Скорость трекбола"
      Begin VB.Menu SpeedOne 
         Caption         =   "1/8x"
      End
      Begin VB.Menu SpeedTwo 
         Caption         =   "1/4x"
      End
      Begin VB.Menu SpeedFour 
         Caption         =   "1/2x"
      End
      Begin VB.Menu SpeedEight 
         Caption         =   "1x"
      End
      Begin VB.Menu Speed16 
         Caption         =   "2x"
      End
   End
End
Attribute VB_Name = "Trackball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Initialize()
    InitOpenGL
    TrackballSetCursor
End Sub
Private Sub Form_Paint()
    Draw
End Sub
Private Sub Form_Resize()
    Static W&, H&
    Sizing ScaleWidth, ScaleHeight
    If ScaleWidth <= W And ScaleHeight <= H Then Draw
    W = ScaleWidth: H = ScaleHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload PicField
    UnloadGL
End Sub
Public Sub Form_KeyDown(Keycode%, Shift%) 'реагирование на нажатие кнопок
    TrackballEscape (Keycode)
End Sub
Public Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TrackballMouseDown Button, Shift, X, Y
End Sub
Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        Exit Sub
    End If
    If Button = 1 Then
        If Abs(Y - ScaleHeight / 2) < 1 / 6 * ScaleHeight And Abs(X - ScaleWidth / 2) < 1 / 6 * ScaleWidth Then
           TrackballMouseMove X, Y
       End If
    End If
End Sub
Private Sub Speed16_Click()
    UnCheckMenu
    Speed16.Checked = True
End Sub
Private Sub SpeedEight_Click()
    UnCheckMenu
    SpeedEight.Checked = True
End Sub
Private Sub SpeedFour_Click()
    UnCheckMenu
    SpeedFour.Checked = True
End Sub
Private Sub SpeedOne_Click()
    UnCheckMenu
    SpeedOne.Checked = True
End Sub
Private Sub SpeedTwo_Click()
    UnCheckMenu
    SpeedTwo.Checked = True
End Sub
