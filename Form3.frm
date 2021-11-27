VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   8355
   ClientLeft      =   9825
   ClientTop       =   1590
   ClientWidth     =   9105
   LinkTopic       =   "Form3"
   ScaleHeight     =   8355
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "VOLVER"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contactos"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   600
      Picture         =   "Form3.frx":0000
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Unload Form3
    Form1.Enabled = True
    Form1.Show
End Sub

Private Sub Image1_Click()
    Form3.Enabled = False
    Form2.Show
    Form2.Text1.SetFocus
End Sub
