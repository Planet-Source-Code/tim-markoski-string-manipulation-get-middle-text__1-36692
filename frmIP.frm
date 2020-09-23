VERSION 5.00
Begin VB.Form frmIP 
   Caption         =   "String Manipulation"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtIP 
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CommandButton cmdGetText 
      Caption         =   "Get Middle Text"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Text            =   "ftp://usfhj:pghjghjass@10.225.12.238/dir/d ir/"
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label lblIP 
      AutoSize        =   -1  'True
      Caption         =   "Returned String Data"
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label lblIP 
      AutoSize        =   -1  'True
      Caption         =   "Original String Data"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   1365
   End
End
Attribute VB_Name = "frmIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGetText_Click()
    
    On Error GoTo PROC_ERR

    txtIP(1).Text = GetIPfromWeb("@", txtIP(0).Text, "/")

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    txtIP(1).Text = Err.Description
    MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
    
End Sub
