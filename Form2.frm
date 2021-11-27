VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Busqueda"
   ClientHeight    =   10650
   ClientLeft      =   7545
   ClientTop       =   1590
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H0000C0C0&
      Caption         =   "Contacto"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   1920
      TabIndex        =   27
      Top             =   6120
      Visible         =   0   'False
      Width           =   11535
      Begin VB.CommandButton bt_guardar 
         Enabled         =   0   'False
         Height          =   1215
         Left            =   9720
         Picture         =   "Form2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
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
         ItemData        =   "Form2.frx":06F5
         Left            =   7800
         List            =   "Form2.frx":06FF
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
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
         Height          =   495
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   15
         Top             =   720
         Width           =   4215
      End
      Begin VB.Image Image2 
         Enabled         =   0   'False
         Height          =   1080
         Left            =   960
         Picture         =   "Form2.frx":071D
         Top             =   1680
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
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
         Left            =   480
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
   End
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
      Height          =   3615
      Left            =   1920
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   11535
      Begin VB.ComboBox Combo4 
         Enabled         =   0   'False
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
         ItemData        =   "Form2.frx":1871
         Left            =   7440
         List            =   "Form2.frx":18CC
         TabIndex        =   11
         Top             =   2400
         Width           =   3135
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
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
         ItemData        =   "Form2.frx":19D6
         Left            =   5880
         List            =   "Form2.frx":19E3
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         MaxLength       =   17
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
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
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "Dia"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
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
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   14
         Text            =   "Año"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
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
         ItemData        =   "Form2.frx":1A02
         Left            =   2760
         List            =   "Form2.frx":1A2D
         TabIndex        =   12
         Text            =   "Mes"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         MaxLength       =   17
         TabIndex        =   7
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         MaxLength       =   43
         TabIndex        =   4
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
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
         Left            =   5880
         TabIndex        =   30
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Genero"
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
         Left            =   4920
         TabIndex        =   26
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad"
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "barrio"
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
         Left            =   6720
         TabIndex        =   24
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nac."
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Apodo"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "N°"
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   21
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "calle"
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido y Nombre"
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9600
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   11535
      Begin VB.CommandButton bus 
         Enabled         =   0   'False
         Height          =   735
         Left            =   8640
         Picture         =   "Form2.frx":1A96
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         MaxLength       =   7
         TabIndex        =   2
         ToolTipText     =   "numero"
         Top             =   480
         Width           =   4335
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
         Height          =   495
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   1
         ToolTipText     =   "caracteristica"
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bt_guardar_Click()
    tot = FileLen(App.Path + "/AGENDA.dat") / Len(RegAg)
    RegAg.Contac.cel = "(" & Text1.Text & ")" & Text2.Text
    RegAg.Datos.AyN = Text3.Text
    RegAg.Datos.apodo = Text5.Text
    RegAg.Datos.sexo = Left(Combo2.Text, 1)
    RegAg.Datos.dir = Text4.Text & "/" & Text9.Text & "/" & Text10.Text & "/" & Text6.Text & "/" & Combo4.Text
    RegAg.Datos.F_nac = mesx(Combo1.Text) & "/" & Text8.Text & "/" & Text7.Text
    RegAg.Contac.corr = Text11.Text
    RegAg.Contac.estado = estado(Combo3.Text)
    RegAg.Contac.f_cre = Date
    Put #1, tot + 1, RegAg
    ''Call borro
    Close #1
    Unload Form2
    Form3.Enabled = True
    Form3.Show
    
       
End Sub

Public Sub borro()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text5.Text = ""
    Combo2.ListIndex = -1
    Text4.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text6.Text = ""
    Combo4.ListIndex = -1
    Combo1.ListIndex = -1
    Text8.Text = ""
    Text7.Text = ""
    Text11.Text = ""
    Combo3.ListIndex = -1
    Frame3.Visible = False: Frame3.Enabled = False
    Frame2.Visible = False: Frame2.Enabled = False
    bt_guardar.Enabled = False
    bt_guardar.Visible = False
    Text1.SetFocus
    
    
    
    
    
    
    
    
End Sub

Private Sub bus_Click()
    celbus = "(" & Text1.Text & ")" & Text2.Text
    Open App.Path + "/agenda.dat" For Random As #1 Len = Len(RegAg)
    tot = FileLen(App.Path + "/agenda.dat") / Len(RegAg)
    Dim b As Byte
    Dim C As Byte
    
    b = 0
    For C = 1 To tot Step 1
        Get #1, C, RegAg
        If (RegAg.Contac.cel = celbus) Then
                                        b = 1
                                        pos = C
                                        C = tot
            End If
    Next C
    If (b = 0) Then
                    Frame2.Visible = True: Frame2.Enabled = True
                    Text3.Enabled = True
                    Text3.SetFocus
                    bus.Enabled = False
            Else
                    z = MsgBox("Contacto Existente", , "")
                    Form4.Show
                    Form2.Enabled = False
                    Form4.volver.SetFocus
           End If
    
End Sub







Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
          If (Combo1.ListIndex <> -1) Then
                Combo1.Locked = True
                Text8.Enabled = True
                Text8.SetFocus
           End If
           
        Else
            KeyAscii = 0
    End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
          If (Combo2.ListIndex <> -1) Then
                Combo2.Locked = True
                Text4.Enabled = True
                Text4.SetFocus
           End If
           
        Else
            KeyAscii = 0
    End If
End Sub





Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
          If (Combo3.ListIndex <> -1) Then
                Combo3.Locked = True
                bt_guardar.Enabled = True
                bt_guardar.Visible = True
                bt_guardar.SetFocus
           End If
           
        Else
            KeyAscii = 0
    End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
     If (KeyAscii = 13) Then
          If (Combo4.ListIndex <> -1) Then
                Combo4.Locked = True
                Combo1.Enabled = True
                Combo1.SetFocus
           End If
           
        Else
            KeyAscii = 0
    End If
End Sub



Private Sub Command2_Click()
Unload Form2
Unload Form3
Form1.Enabled = True
Form1.Visible = True
Form1.Show
Close #1
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
                            If (Len(Text1.Text) = 3) Then
                                                    Text1.Enabled = False
                                                    Text2.Enabled = True
                                                    Text2.SetFocus
                                    Else
                                        z = MsgBox("Error en el Codigo de area", , "")
                            End If
                Else
                        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
                                                        KeyAscii = 0
                         End If
        End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
     If (KeyAscii = 13) Then
                    If (Text10.Text <> "") Then
                                Text10.Locked = True
                                Text6.Enabled = True
                                Text6.SetFocus
                        Else
                               z = MsgBox("Campo vacio", , "")
                               Text10.SetFocus
                    End If
        Else
               key = Asc(UCase(Chr(KeyAscii)))
               If Not (key >= 65 And key <= 90 Or key = 8 Or key = 32 Or key >= 48 And key <= 57) Then
                            KeyAscii = 0
            End If
    End If
End Sub



Private Sub Text11_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
                    If (Text11.Text <> "") Then
                                    Text11.Locked = True
                                    Combo3.Enabled = True
                                    Combo3.SetFocus
                            Else
                                z = MsgBox("Ingreso no valido", , "")
                                Text11.SetFocus
                    End If
                Else
                    key = Asc(LCase(Chr(KeyAscii)))
                    If Not (key >= 33 And key <= 57 Or key = 64 Or key = 8 Or key >= 91 And key <= 122) Then
                                        KeyAscii = 0
                    End If
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
                            If (Len(Text2.Text) = 7) Then
                                                    bus.Enabled = True
                                                    bus.SetFocus
                                                    Text2.Enabled = False
                                    Else
                                        z = MsgBox("Error en el Numero", , "")
                            End If
                Else
                        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
                                                        KeyAscii = 0
                         End If
        End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
                    If (Text3.Text <> "") Then
                                Text3.Locked = True
                                Text5.Enabled = True
                                Text5.SetFocus
                        Else
                               z = MsgBox("Campo vacio", , "")
                               Text3.SetFocus
                    End If
        Else
               key = Asc(UCase(Chr(KeyAscii)))
               If Not (key >= 65 And key <= 90 Or key = 8 Or key = 32) Then
                            KeyAscii = 0
            End If
    End If
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
                If (Text4.Text <> "") Then
                            Text4.Locked = True
                            Text9.Enabled = True
                            Text9.SetFocus
                        Else
                            z = MsgBox("Campo vacio", , "")
                            Text4.SetFocus
                End If
                
            Else
                    key = Asc(UCase(Chr(KeyAscii)))
                    If Not (key >= 65 And key <= 90 Or key = 8 Or key = 32) Then
                                                KeyAscii = 0
                     End If
     End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)

If (KeyAscii = 13) Then
    If (Text5.Text <> "") Then
                Text5.Locked = True
                Combo2.Enabled = True
                Combo2.SetFocus
        Else
                z = MsgBox("Campo vacio", , "")
                Text5.SetFocus
    End If
End If
                
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
                If (Text6.Text <> "") Then
                            Text6.Locked = True
                            Combo4.Enabled = True
                            Combo4.SetFocus
                        Else
                            z = MsgBox("Campo vacio", , "")
                            Text6.SetFocus
                End If
                
            Else
                    key = Asc(UCase(Chr(KeyAscii)))
                    If Not (key >= 65 And key <= 90 Or key = 8 Or key = 32) Then
                                                KeyAscii = 0
                     End If
     End If
End Sub

Private Sub Text7_GotFocus()
Text7.Text = ""
End Sub



Private Sub Text7_KeyPress(KeyAscii As Integer)
     If (KeyAscii = 13) Then
                        key = Val(Year(Date)) - 10
                        If (Val(Text7.Text) >= 1930 And Val(Text7.Text) <= key) Then
                                        Text7.Locked = True
                                        Frame3.Enabled = True: Frame3.Visible = True
                                        Text11.Enabled = True
                                        Text11.SetFocus
                                        
                                    Else
                                        z = MsgBox("Ingreso no valido", , "")
                                        Text7.Text = ""
                                        Text7.SetFocus
                                        
                                        
                        End If
                    Else
                        
                        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
                                            KeyAscii = 0
                        End If
    End If
End Sub

Private Sub Text7_LostFocus()
If (Text7.Text = "") Then
                    Text7.Text = "Año"
                    End If
                    
End Sub



Private Sub Text8_GotFocus()
Text8.Text = ""
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
                        
                        If (Val(Text8.Text) >= 1 And 31 >= Val(Text8.Text)) Then
                                        Text8.Locked = True
                                        Text7.Enabled = True: Text7.SetFocus
                                    Else
                                        z = MsgBox("Ingreso no valido", , "")
                                        Text8.Text = ""
                                        Text8.SetFocus
                        End If
                    Else
                        
                        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
                                            KeyAscii = 0
                        End If
    End If
End Sub

Private Sub Text8_LostFocus()
If (Text8.Text = "") Then
                    Text8.Text = "Dia"
                    End If
End Sub


Private Sub Text9_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
                If (Val(Text9.Text) >= 0 And Val(Text9.Text) <= 9999) Then
                                Text9.Locked = True
                                Text10.Enabled = True: Text10.SetFocus
                        Else
                                z = MsgBox("Fuera de rango numerico", , "")
                                Text10.SetFocus
                                
                End If
            Else
                key = Asc(UCase(Chr(KeyAscii)))
                If Not (key >= 48 And key <= 57 Or key = 8) Then
                                KeyAscii = 0
                End If
    End If
End Sub
