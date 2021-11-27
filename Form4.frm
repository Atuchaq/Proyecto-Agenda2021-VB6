VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   9975
   ClientLeft      =   8235
   ClientTop       =   2055
   ClientWidth     =   18030
   LinkTopic       =   "Form4"
   ScaleHeight     =   9975
   ScaleWidth      =   18030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   16215
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         MaxLength       =   7
         TabIndex        =   22
         Text            =   "numero"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         MaxLength       =   3
         TabIndex        =   21
         Text            =   "cod"
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Form4.frx":0000
         Left            =   3960
         List            =   "Form4.frx":000D
         TabIndex        =   20
         Text            =   "Genero"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   19
         Top             =   5040
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12600
         TabIndex        =   18
         Top             =   5040
         Width           =   3135
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Form4.frx":002C
         Left            =   4920
         List            =   "Form4.frx":0036
         TabIndex        =   17
         Text            =   "Estado"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         TabIndex        =   16
         Text            =   "correo"
         Top             =   3480
         Width           =   4095
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         MaxLength       =   4
         TabIndex        =   15
         Text            =   "año"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "dia"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Form4.frx":0054
         Left            =   360
         List            =   "Form4.frx":007F
         TabIndex        =   13
         Text            =   "Mes"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9000
         MaxLength       =   20
         TabIndex        =   12
         Text            =   "localidad"
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         MaxLength       =   17
         TabIndex        =   11
         Text            =   "barrio"
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "numero"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   17
         TabIndex        =   9
         Text            =   "calle"
         Top             =   1920
         Width           =   3615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Form4.frx":00E8
         Left            =   12480
         List            =   "Form4.frx":0143
         TabIndex        =   8
         Text            =   "Provincia"
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   15
         TabIndex        =   7
         Text            =   "apodo"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   43
         TabIndex        =   6
         Text            =   "Apellido y nombre"
         Top             =   480
         Width           =   8775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   4455
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C000&
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0000C000&
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000C000&
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.CommandButton volver 
      Caption         =   "Volver"
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
      Left            =   480
      TabIndex        =   0
      Top             =   8760
      Width           =   3015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Option3_Click()
    
    Text1.Text = Trim(RegAg.Datos.AyN)
    Text2.Text = Trim(RegAg.Datos.apodo)
    Combo1.Text = Genero(RegAg.Datos.sexo)
    Call f_Dir(Trim(RegAg.Datos.dir), 1)
    Text3.Text = Dire(1)
    Text4.Text = Dire(2)
    Text5.Text = Dire(3)
    Text6.Text = Dire(4)
    Combo2.Text = Dire(5)
    Fechas (RegAg.Datos.F_nac)
    Text9.Text = Trim(RegAg.Contac.corr)
    Combo4.Text = estadoNum(RegAg.Contac.estado)
    Form2.Text1.Enabled = True
    Form2.Text2.Enabled = True
    Text10.Text = Form2.Text1.Text
    Text11.Text = Form2.Text2.Text
    
    
    
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8) Then
                                        KeyAscii = 0
                End If
End Sub
Private Sub Text1_LostFocus()
    If Not (Len(Text1.Text) >= 1 And Len(Text1.Text) <= 43) Then
                    z = MsgBox("Error: deben ser entre 1 y 43 caracteres", , "")
                    Text1.SetFocus
            End If
End Sub



Private Sub Text10_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
                    KeyAscii = 0
                    End If
End Sub

Private Sub Text10_LostFocus()
    If Not (Len(Text10.Text) = 3) Then
                    z = MsgBox("Error: deben ser 3 digitos", , "")
                    Text10.SetFocus
            End If
End Sub



Private Sub Text11_KeyPress(KeyAscii As Integer)
       If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
                    KeyAscii = 0
                    End If
End Sub

Private Sub Text11_LostFocus()
     If Not (Len(Text11.Text) = 7) Then
                    z = MsgBox("Error: deben ser 7 digitos", , "")
                    Text11.SetFocus
            End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8) Then
                                        KeyAscii = 0
                End If
End Sub

Private Sub Text2_LostFocus()
    If Not (Len(Text2.Text) >= 1) Then
                    z = MsgBox("Error: debe escribir en el campo", , "")
                    Text2.SetFocus
            End If
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8) Then
                                        KeyAscii = 0
                End If
End Sub

Private Sub Text3_LostFocus()
     If Not (Len(Text3.Text) >= 1) Then
                    z = MsgBox("Error: debe escribir en el campo", , "")
                    Text3.SetFocus
        End If
End Sub



Private Sub Text4_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
                    KeyAscii = 0
                    End If
End Sub

Private Sub Text4_LostFocus()
     If Not (Len(Text4.Text) >= 1) Then
                    z = MsgBox("Error: debe escribir el nuemero", , "")
                    Text4.SetFocus
            End If
End Sub



Private Sub Text5_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8) Then
                                        KeyAscii = 0
                End If
End Sub

Private Sub Text5_LostFocus()
    If Not (Len(Text5.Text) >= 1) Then
                    z = MsgBox("Error: debe escribir en el campo", , "")
                    Text5.SetFocus
            End If
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8) Then
                                        KeyAscii = 0
                End If
End Sub

Private Sub Text6_LostFocus()
     If Not (Len(Text6.Text) >= 1) Then
                    z = MsgBox("Error: debe escribir en el campo", , "")
                    Text6.SetFocus
            End If
End Sub



Private Sub Text7_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
                    KeyAscii = 0
                    End If
End Sub

Private Sub Text7_LostFocus()

    
    If (Combo3.Text = "Febrero") Then
        If Not (Val(Text7.Text) >= 1 And Val(Text7.Text) <= 29) Then
                    z = MsgBox("Error: dia incorrecto", , "")
                    Text7.SetFocus
        Else
            If Not (Val(Text7.Text) >= 1 And Val(Text7.Text) <= 31) Then
                    z = MsgBox("Error: dia incorrecto", , "")
                    Text7.SetFocus
            End If
        End If
     End If
            
            
            
End Sub



Private Sub Text8_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
                    KeyAscii = 0
                    End If
End Sub

Private Sub Text8_LostFocus()
     Key = Val(Year(Date)) - 10
    If Not (Val(Text8.Text) >= 1930 And Val(Text8.Text) <= Key) Then
                 z = MsgBox("Error: Año incorrecto", , "")
    End If
End Sub

Private Sub volver_Click()
    Form2.Enabled = True
    Unload Form4
    Form2.Show
    Close #1
    
End Sub

Public Sub Fechas(Fecha As String)
    Dim i As Byte
    Dim Longitud As Byte
    Dim h As Byte
    Dim Last As Byte
    Dim b As Boolean
    b = True
    Longitud = Len(Fecha)
    
    For i = 1 To Longitud Step 1
            If (Mid(Fecha, i, 1) = "/") Then                  ''enero/22/2010
                        h = h + 1
                        If (b) Then
                            b = False
                            Last = i + 1
                        End If
                        Select Case h
                        Case 1: Combo3.Text = mesy(Mid(Fecha, 1, i - 1))
                        Case 2: Text7.Text = Mid(Fecha, Last, i - Last)
                        Case Else
                         
                        End Select
            End If
    Next i
    Text8.Text = Right(Fecha, 4)
    
End Sub
Public Function mesy(mes As String) As String
    
    Select Case mes
    Case "01": mesy = "Enero"
    Case "02": mesy = "Febrero"
    Case "03": mesy = "Marzo"
    Case "04": mesy = "Abril"
    Case "05": mesy = "Mayo"
    Case "06": mesy = "Junio"
    Case "07": mesy = "Julio"
    Case "08": mesy = "Agosto"
    Case "09": mesy = "Septiembre"
    Case "10": mesy = "Octubre"
    Case "11": mesy = "Nobiembre"
    Case "12": mesy = "Diciembre"
    Case Else
    End Select
    
End Function

