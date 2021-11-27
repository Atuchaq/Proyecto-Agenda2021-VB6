VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   10935
   ClientLeft      =   9270
   ClientTop       =   1740
   ClientWidth     =   10455
   DrawMode        =   4  'Mask Not Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   10455
   Begin VB.CommandButton botonIngreso 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ingresar"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   2500
   End
   Begin VB.CommandButton botonSalir 
      BackColor       =   &H0000C0C0&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   2500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      Width           =   11295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End
End Sub

Private Sub botonIngreso_Click()
    Form1.Enabled = False
    Form1.Visible = False
    Form3.Show
    End Sub

Private Sub botonSalir_Click()
End
End Sub

Private Sub Form_Load()
    Label1.Caption = "Agenda de:" & Chr(10) & Chr(13) & "Jose Luis Figueroa"
    
End Sub

