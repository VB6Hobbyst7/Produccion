VERSION 5.00
Begin VB.Form frmStockIngreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Tarjetas"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmStockIngreso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   765
      Left            =   0
      TabIndex        =   17
      Top             =   2880
      Width           =   6990
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   5850
         TabIndex        =   7
         Top             =   225
         Width           =   1050
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo Ingreso"
         Height          =   375
         Left            =   105
         TabIndex        =   8
         Top             =   225
         Width           =   1380
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   2760
      Left            =   45
      TabIndex        =   9
      Top             =   105
      Width           =   7005
      Begin VB.Frame Frame4 
         Caption         =   "Observación :"
         Height          =   1305
         Left            =   90
         TabIndex        =   16
         Top             =   1380
         Width           =   6840
         Begin VB.TextBox txtObservacion 
            Height          =   945
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   210
            Width           =   6675
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Detalle de Lote"
         Height          =   765
         Left            =   105
         TabIndex        =   12
         Top             =   570
         Width           =   6825
         Begin VB.TextBox txtAl 
            Height          =   315
            Left            =   5025
            Locked          =   -1  'True
            MaxLength       =   16
            TabIndex        =   4
            Top             =   300
            Width           =   1710
         End
         Begin VB.TextBox txtDel 
            Height          =   315
            Left            =   2595
            Locked          =   -1  'True
            MaxLength       =   16
            TabIndex        =   3
            Top             =   285
            Width           =   1755
         End
         Begin VB.TextBox txtCantidad 
            Height          =   315
            Left            =   930
            TabIndex        =   2
            Text            =   "0"
            Top             =   315
            Width           =   840
         End
         Begin VB.Label Label6 
            Caption         =   "AL :"
            Height          =   300
            Left            =   4515
            TabIndex        =   15
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label5 
            Caption         =   "DEL :"
            Height          =   300
            Left            =   2055
            TabIndex        =   14
            Top             =   315
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Cantidad :"
            Height          =   300
            Left            =   105
            TabIndex        =   13
            Top             =   345
            Width           =   735
         End
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   720
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "dd/mm/aaaa"
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox txtLote 
         Height          =   315
         Left            =   2850
         TabIndex        =   1
         Text            =   "0"
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha :"
         Height          =   300
         Left            =   150
         TabIndex        =   11
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   "Lote : "
         Height          =   300
         Left            =   2355
         TabIndex        =   10
         Top             =   225
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmStockIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CmdAceptar_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim sResp As String
Dim Inicial As Long
Dim Final As Long
Dim cant As Long
    
If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "Monto Invalido", vbInformation, "Aviso"
        Exit Sub
    End If

If Trim(Me.txtDel.Text) = "" Or Trim(Me.txtAl.Text) = "" Then
        MsgBox "Rango de Tarjetas, Incorrecto", vbInformation, "Aviso"
        Exit Sub
End If



If Not IsDate(Me.txtFecha.Text) Then
        MsgBox "Fecha Invalida"
        Exit Sub
    End If

    Inicial = (Left(Mid(txtDel.Text, 9), 7))
    Final = (Left(Mid(txtAl.Text, 9), 7))
    
    
    'MsgBox ("El valor a restar es: " & (Right(txtDel.Text, 9)))
'    If Final < Inicial Or Final = Inicial Then
'        MsgBox ("Las cantidades no son validas")
'        Exit Sub
'    Else
        'cant = Final - Inicial
        'txtCantidad.Text = (cant)
        cant = CInt(txtCantidad.Text)
        If MsgBox("Cantidad de Tarjetas a Ingresar : " & Str(cant) & "  Unidades, Desea Continuar ?", vbInformation + vbYesNo, "Ingreso de Tarjetas") = vbNo Then
            Exit Sub
        End If
        
        
        If txtCantidad.Text > 0 Then
            Set Cmd = New ADODB.Command
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@dFecha", adDate, adParamInput, 16, txtFecha.Text)
            Cmd.Parameters.Append Prm
            
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@nLote", adInteger, adParamInput, , CInt(txtLote.Text))
            Cmd.Parameters.Append Prm
            
            
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@nCantidad", adBigInt, adParamInput, , cant)
            Cmd.Parameters.Append Prm
                        
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@nDel", adVarChar, adParamInput, 50, txtDel.Text)
            Cmd.Parameters.Append Prm
            
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@nAl", adVarChar, adParamInput, 50, txtAl.Text)
            Cmd.Parameters.Append Prm
            
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@cObservacion", adVarChar, adParamInput, 100, txtObservacion.Text)
            Cmd.Parameters.Append Prm
            
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@cUserActiv", adChar, adParamInput, 4, gsCodUser)
            Cmd.Parameters.Append Prm
            
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@nCodAge", adChar, adParamInput, 4, gsCodAge)
            Cmd.Parameters.Append Prm
            
            oConec.AbreConexion
            Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
            Cmd.CommandType = adCmdStoredProc
            Cmd.CommandText = "ATM_RegistrarStockIngreso"
            Cmd.Execute
            
            oConec.CierraConexion
            cmdAceptar.Enabled = False
            MsgBox "Tarjetas Ingresadas con Exito", vbInformation, "Ingreso de Tarjetas"
        Else
            MsgBox "Las cantidades deben ser distintas de 0"
            Exit Sub
        End If
        limpiaForm
'    End If
    cmdAceptar.Enabled = False

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdNuevo_Click()
    Call limpiaForm
End Sub

Private Sub limpiaForm()
    Frame2.Enabled = True
    txtFecha.Text = gdFecSis
    txtLote.Text = "0"
    txtCantidad.Text = "0"
    txtDel.Text = "8901000000000000"
    txtAl.Text = "8901000000000000"
    txtObservacion.Text = " "
    cmdAceptar.Enabled = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    txtFecha.Text = gdFecSis
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
Dim nCant As Integer
Dim sIni As String
Dim sFin As String

    If KeyAscii = 13 Then
        If Not IsNumeric(Me.txtCantidad.Text) Then
            MsgBox "Cantidad Incorrecta", vbInformation, "Aviso"
            Me.txtCantidad.SetFocus
            Exit Sub
        End If
    
        Call RecuperaRangosDETarjetasEmitidas(CInt(Me.txtCantidad.Text), sIni, sFin, nCant)
        
        Me.txtDel.Text = sIni
        Me.txtAl.Text = sFin
        
        If CInt(Me.txtCantidad.Text) <> nCant Then
            MsgBox "Solo Existen " & nCant & " Tarjetas Emitidas", vbInformation, "Aviso"
            Me.txtCantidad.Text = Trim(Str(nCant))
            Exit Sub
        End If
        
        Me.cmdAceptar.SetFocus
    End If
    

End Sub

Private Sub txtCantidad_LostFocus()
Dim nCant As Integer
Dim sIni As String
Dim sFin As String

    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "Cantidad Incorrecta", vbInformation, "Aviso"
        Me.txtCantidad.SetFocus
        Exit Sub
    End If

        Call RecuperaRangosDETarjetasEmitidas(CInt(Me.txtCantidad.Text), sIni, sFin, nCant)
        
        Me.txtDel.Text = sIni
        Me.txtAl.Text = sFin
        
      If CInt(Me.txtCantidad.Text) <> nCant Then
            MsgBox "Solo Existen " & nCant & " Tarjetas Emitidas", vbInformation, "Aviso"
            Me.txtCantidad.Text = Trim(Str(nCant))
            Exit Sub
        End If

End Sub

Private Sub txtDel_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48) Or (KeyAscii > 57) Then
        KeyAscii = 0
        txtDel.SetFocus
    End If
End Sub
Private Sub txtAl_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48) Or (KeyAscii > 57) Then
        KeyAscii = 0
        txtAl.SetFocus
    End If
End Sub
Private Sub txtLote_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48) Or (KeyAscii > 57) Then
        KeyAscii = 0
        txtLote.SetFocus
    End If
End Sub

