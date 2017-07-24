VERSION 5.00
Begin VB.Form PicField 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Picture"
   ClientHeight    =   5400
   ClientLeft      =   7290
   ClientTop       =   570
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   960
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "PicField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_KeyDown(Keycode%, Shift%) 'реагирование на нажатие кнопок
    PicFieldEscape (Keycode)
End Sub
Private Sub Form_Load()
    PicFieldSetCursor
End Sub
Private Sub Form_Unload(Cancel As Integer)
    PicField.Hide
End Sub
