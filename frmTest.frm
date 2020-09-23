VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Control Resize"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Tag             =   "msiresizeYYYNY"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Tag             =   "MSIRESIZEYYYYN"
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmResize As New ControlResizer

Private Sub Form_Load()
  
  frmResize.KeepRatio = True
  frmResize.FontResize = True
  Call frmResize.InitializeResizer(Me)
    
End Sub
Private Sub Form_Resize()

  Call frmResize.FormResized(Me)
    
End Sub
