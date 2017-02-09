VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3840
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MSComm1.Output = "AT" & Chr(13)
i = 0
Dim buffer As String
Do
buffer = buffer & MSComm1.Input
i = i + 1
Loop Until (i = 200)
Text1.Text = buffer
End Sub

Private Sub Form_Load()
MSComm1.PortOpen = True

End Sub
