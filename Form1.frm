VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   5040
   ClientTop       =   3690
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Quick Ugly interface as an example... Click and drag here to drag the form"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Line Line4 
      X1              =   5880
      X2              =   5880
      Y1              =   2280
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   -240
      X2              =   5880
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      X1              =   720
      X2              =   600
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   2280
      Y2              =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNSnap Me, Button, X, Y
End Sub
