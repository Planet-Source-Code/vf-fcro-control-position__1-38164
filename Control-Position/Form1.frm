VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Solid Resizer By Vanja Fuckar,EMAIL:INGA@VIP.HR"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   2415
      Left            =   3960
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Text1"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   3840
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Text1"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3840
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub Form_Load()
ReadFormControlCollenction Me
MinHeight = 3000
MinWidth = 5000
'MaxHeight = 8000
'MaxWidth = 9000
End Sub

Private Sub Form_Resize()


FormResize Me

ResizeCtrl "Command3", Me, , , , , , "Text2"
ResizeCtrl "Command2", Me, , , , , , "Text1"
ResizeCtrl "Command1", Me, , , , , , "Text1"
ResizeCtrl "Text1", Me, , , , , "Text2", "Label1"
ResizeCtrl "Text2", Me, , , , , , "Label2"

ResizeCtrl "Label1", Me, , , , False
ResizeCtrl "Label2", Me, , , , False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set ControlColl = Nothing
End Sub
