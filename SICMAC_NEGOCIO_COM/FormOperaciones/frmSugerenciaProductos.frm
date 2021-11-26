VERSION 5.00
Begin VB.Form frmSugerenciaProductos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sugerencia de venta de  Productos "
   ClientHeight    =   3495
   ClientLeft      =   11580
   ClientTop       =   5400
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   240
      Picture         =   "frmSugerenciaProductos.frx":0000
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "El Cliente no dispone de los siguientes Productos : "
      Height          =   1215
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmSugerenciaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Public Sub Inicio(ByVal R As ADODB.Recordset)
    Dim ClsPersonas As COMDPersona.DCOMPersonas
    Dim cPersCod As String
    Dim cPersona As String
    Dim nCajaSueldo As Integer
    Dim nPlazoFijo As Integer
    Dim nCTS As Integer
    Dim count As Integer
    Dim SaldoDis As ADODB.Recordset
    Dim Campo As ADODB.Field
    Dim Saldo As Double
    Saldo = 0
    count = R.RecordCount
    
    If count = 1 Then
        
        cPersCod = R!cPersCod
        nCajaSueldo = R!nCajaSueldo
        nPlazoFijo = R!nPlazoFijo
        nCTS = R!nCTS
        
        If nCajaSueldo = 0 And nPlazoFijo = 0 And nCTS = 0 Then
        Label2.Caption = "Caja Sueldo, Plazo Fijo y CTS"
        ElseIf nCajaSueldo > 0 And nPlazoFijo = 0 And nCTS = 0 Then
            Label2.Caption = "Plazo Fijo y CTS"
        ElseIf nCajaSueldo = 0 And nPlazoFijo > 0 And nCTS = 0 Then
            Set ClsPersonas = New COMDPersona.DCOMPersonas
            Set SaldoDis = ClsPersonas.SaldoDisponibleDPF(cPersCod)
            If SaldoDis Is Nothing Then
                Label2.Caption = "Caja Sueldo y CTS"
               
            Else
                Saldo = SaldoDis!nSaldoDisp
                    If Saldo >= 350 Then
                    Label2.Caption = "Caja Sueldo , CTS y Crédito RapiFlash"
                   
                    Else
                    Label2.Caption = "Caja Sueldo y CTS"
                   
                    End If
            End If
        ElseIf nCajaSueldo = 0 And nPlazoFijo = 0 And nCTS > 0 Then
            Label2.Caption = "Caja Sueldo y Plazo Fijo"
           
        ElseIf nCajaSueldo = 0 And nPlazoFijo > 0 And nCTS > 0 Then
            Set ClsPersonas = New COMDPersona.DCOMPersonas
            Set SaldoDis = ClsPersonas.SaldoDisponibleDPF(cPersCod)
            If SaldoDis Is Nothing Then
                Label2.Caption = "Caja Sueldo"
                
            Else
                Saldo = SaldoDis!nSaldoDisp
                    If Saldo >= 350 Then
                    Label2.Caption = "Caja Sueldo y Crédito RapiFlash"
                    
                    Else
                    Label2.Caption = "Caja Sueldo"
                    End If
            End If
        ElseIf nCajaSueldo > 0 And nPlazoFijo = 0 And nCTS > 0 Then
            Label2.Caption = "Plazo Fijo"
            
        ElseIf nCajaSueldo > 0 And nPlazoFijo > 0 And nCTS = 0 Then
            Set ClsPersonas = New COMDPersona.DCOMPersonas
            Set SaldoDis = ClsPersonas.SaldoDisponibleDPF(cPersCod)
            If SaldoDis Is Nothing Then
                Label2.Caption = "CTS"
                
            Else
                Saldo = SaldoDis!nSaldoDisp
                If Saldo >= 350 Then
                Label2.Caption = "CTS  y Crédito RapiFlash"
                
                Else
                Label2.Caption = "CTS"
                
                End If
            End If
        Else
            Set ClsPersonas = New COMDPersona.DCOMPersonas
            Set SaldoDis = ClsPersonas.SaldoDisponibleDPF(cPersCod)
            If SaldoDis Is Nothing Then
                Label2.Caption = ""
            Else
                Saldo = SaldoDis!nSaldoDisp
                If Saldo >= 350 Then
                Label2.Caption = "Crédito RapiFlash"
                
                Else
                Label2.Caption = ""
                End If
            End If
        End If
        
        Me.Show 1
    Else
         Do While Not R.EOF
            cPersCod = R!cPersCod
            cPersona = R!cPersNombre
            nCajaSueldo = R!nCajaSueldo
            nPlazoFijo = R!nPlazoFijo
            nCTS = R!nCTS
            
            If nCajaSueldo = 0 And nPlazoFijo = 0 And nCTS = 0 Then
            Label1.Caption = "El Cliente " & cPersona & " no dispone de los siguientes Productos : "
            Label2.Caption = "Caja Sueldo, Plazo Fijo y CTS"
            
            ElseIf nCajaSueldo > 0 And nPlazoFijo = 0 And nCTS = 0 Then
                Label1.Caption = "El Cliente " & cPersona & " no dispone de los siguientes Productos : "
                Label2.Caption = "Plazo Fijo y CTS"
                
            ElseIf nCajaSueldo = 0 And nPlazoFijo > 0 And nCTS = 0 Then
                Label1.Caption = "El Cliente " & cPersona & " no dispone de los siguientes Productos : "
                Set ClsPersonas = New COMDPersona.DCOMPersonas
                Set SaldoDis = ClsPersonas.SaldoDisponibleDPF(cPersCod)
                If SaldoDis Is Nothing Then
                Label2.Caption = "Caja Sueldo y CTS"
                
                Else
                Saldo = SaldoDis!nSaldoDisp
                    If Saldo >= 350 Then
                    Label2.Caption = "Caja Sueldo , CTS y Crédito RapiFlash"
                    
                    Else
                    Label2.Caption = "Caja Sueldo y CTS"
                    
                    End If
                End If
            ElseIf nCajaSueldo = 0 And nPlazoFijo = 0 And nCTS > 0 Then
                Label1.Caption = "El Cliente " & cPersona & " no dispone de los siguientes Productos : "
                Label2.Caption = "Caja Sueldo y Plazo Fijo"
               
            ElseIf nCajaSueldo = 0 And nPlazoFijo > 0 And nCTS > 0 Then
                Label1.Caption = "El Cliente " & cPersona & " no dispone de los siguientes Productos : "
                Set ClsPersonas = New COMDPersona.DCOMPersonas
                Set SaldoDis = ClsPersonas.SaldoDisponibleDPF(cPersCod)
                If SaldoDis Is Nothing Then
                Label2.Caption = "Caja Sueldo"
 
                Else
                Saldo = SaldoDis!nSaldoDisp
                    If Saldo >= 350 Then
                    Label2.Caption = "Caja Sueldo y Crédito RapiFlash"
                    
                    Else
                    Label2.Caption = "Caja Sueldo"
                   
                    End If
                End If
            ElseIf nCajaSueldo > 0 And nPlazoFijo = 0 And nCTS > 0 Then
                Label1.Caption = "El Cliente " & cPersona & " no dispone de los siguientes Productos : "
                Label2.Caption = "Plazo Fijo"
                
            ElseIf nCajaSueldo > 0 And nPlazoFijo > 0 And nCTS = 0 Then
                Set ClsPersonas = New COMDPersona.DCOMPersonas
                Set SaldoDis = ClsPersonas.SaldoDisponibleDPF(cPersCod)
                If SaldoDis Is Nothing Then
                Label2.Caption = "CTS"
                
                Else
                Saldo = SaldoDis!nSaldoDisp
                    If Saldo >= 350 Then
                    Label2.Caption = "CTS  y Crédito RapiFlash"
                    
                    Else
                    Label2.Caption = "CTS"
                    
                    End If
                End If
            Else
                Set ClsPersonas = New COMDPersona.DCOMPersonas
                Set SaldoDis = ClsPersonas.SaldoDisponibleDPF(cPersCod)
                If SaldoDis Is Nothing Then
                Label2.Caption = ""
                Else
                Saldo = SaldoDis!nSaldoDisp
                    If Saldo >= 350 Then
                    Label2.Caption = "Crédito RapiFlash"
                    
                    Else
                    Label2.Caption = ""
                    End If
                End If
            End If
            
            Me.Show 1
         R.MoveNext
        Loop
        R.Close
        Set R = Nothing

    End If

End Sub

