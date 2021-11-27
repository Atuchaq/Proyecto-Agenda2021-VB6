Attribute VB_Name = "Module1"
Option Explicit
Public z As Byte
Public Key As Integer
Public Type A
    cel As String * 12
    corr As String * 20
    f_cre As String * 10
    relac As String * 15
    estado As Byte
    End Type
    
Public Type b
    AyN As String * 43
    F_nac As String * 10
    apodo As String * 15
    dir As String * 80
    sexo As String * 1
    End Type
    
Public Type G
    Id_C As Byte
    Contac As A
    Datos As b
    End Type
    
Public RegAg As G
Public celbus As String * 12
Public tot As Byte
Public pos As Byte
Public Dire(1 To 5) As String

Public Function estado(E As String) As Byte
    UCase (E)
    If (E = "HABILITADO") Then
                estado = 1
            Else
                estado = 0
    End If
    
End Function
Public Function estadoNum(E As Byte) As String
    UCase (E)
    If (E = 1) Then
                estadoNum = "Habilitado"
            Else
                estadoNum = "Inhabilitado"
    End If
    
End Function

Public Function Genero(G As String) As String
    UCase (G)
    If (G = "M") Then
                Genero = "Masculino"
            Else
                    If (G = "F") Then
                             Genero = "Femenino"
                        Else
                             Genero = "Otro"
                    End If
    End If
    
End Function

Public Function f_Dir(Cad As String, i As Byte) As String
    Dim h As Byte
    Dim b As Byte
    Dim Longitud As Byte
    Longitud = Len(Cad)
    b = 0
    h = 1
    
    If (i <> 6) Then
        If Not (i = 5) Then
                While (Longitud >= h And b = 0)
                                If (Mid(Cad, h, 1) = "/") Then
                                                    b = 1
                                                    Dire(i) = Trim(Left(Cad, h - 1))
                                End If
                                h = h + 1
                 Wend
                 Call f_Dir(Mid(Cad, h, Longitud), i + 1)
                 Else
                    Dire(i) = Trim(Cad)
        End If
                 
    End If
    
End Function

Public Function mesx(mes As String) As String
    
    Select Case mes
    Case "Enero": mesx = "01"
    Case "Febrero": mesx = "02"
    Case "Marzo": mesx = "03"
    Case "Abril": mesx = "04"
    Case "Mayo": mesx = "05"
    Case "Junio": mesx = "06"
    Case "Julio": mesx = "07"
    Case "Agosto": mesx = "08"
    Case "Septiembre": mesx = "09"
    Case "Octubre": mesx = "10"
    Case "Nobiembre": mesx = "11"
    Case "Diciembre": mesx = "12"
    Case Else
    End Select
    
End Function




    
