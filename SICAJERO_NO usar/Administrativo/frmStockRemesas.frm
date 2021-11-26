VERSION 5.00
Begin VB.Form frmStockRemesas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remesas de Tarjetas"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "frmStockRemesas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   765
      Left            =   0
      TabIndex        =   14
      Top             =   3480
      Width           =   7455
      Begin VB.CommandButton cmdAcepta 
         Caption         =   "Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nueva Remesa"
         Height          =   375
         Left            =   105
         TabIndex        =   16
         Top             =   225
         Width           =   1560
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6315
         TabIndex        =   15
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3390
      Left            =   45
      TabIndex        =   5
      Top             =   45
      Width           =   7470
      Begin VB.OptionButton rdbDevolucion 
         Caption         =   "Devolucion"
         Height          =   285
         Left            =   4230
         TabIndex        =   21
         Top             =   315
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.OptionButton rdbRemesa 
         Caption         =   "Remesa"
         Height          =   285
         Left            =   2730
         TabIndex        =   20
         Top             =   315
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   780
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "dd/mm/aaaa"
         Top             =   240
         Width           =   1245
      End
      Begin VB.Frame Frame4 
         Caption         =   "Detalle de Remesa"
         Height          =   765
         Left            =   75
         TabIndex        =   10
         Top             =   2565
         Width           =   7275
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
            Left            =   2580
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
            TabIndex        =   13
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label5 
            Caption         =   "DEL :"
            Height          =   300
            Left            =   2055
            TabIndex        =   12
            Top             =   315
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Cantidad :"
            Height          =   300
            Left            =   105
            TabIndex        =   11
            Top             =   345
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Destino "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   60
         TabIndex        =   8
         Top             =   1590
         Width           =   7305
         Begin VB.ComboBox CboDestino 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   330
            Width           =   5700
         End
         Begin VB.Label Label2 
            Caption         =   "Area / Agencia :"
            Height          =   285
            Left            =   75
            TabIndex        =   9
            Top             =   360
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   75
         TabIndex        =   6
         Top             =   660
         Width           =   7290
         Begin VB.ComboBox CboOrigen 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   300
            Width           =   5700
         End
         Begin VB.Label Label1 
            Caption         =   "Area / Agencia :"
            Height          =   285
            Left            =   135
            TabIndex        =   7
            Top             =   330
            Width           =   1185
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha :"
         Height          =   300
         Left            =   105
         TabIndex        =   17
         Top             =   270
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmStockRemesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta


Private Sub cmdAcepta_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim sResp As String
Dim Inicial As Long
Dim Final As Long
Dim cant As Long
    

    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "Cantidad Incorrecta", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If CInt(Me.txtCantidad.Text) <= 0 Then
        MsgBox "Cantidad Incorrecta", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Not IsDate(Me.txtFecha.Text) Then
        MsgBox "Fecha Invalida"
        Exit Sub
    End If
    
    If txtFecha.Text < gdFecSis Then
        MsgBox ("La Fecha no puede ser menor a la fecha del sistema"), vbInformation, "Salida de Tarjetas"
        Exit Sub
    End If
    
    If (Mid(CboOrigen.Text, 1, InStr(CboOrigen.Text, "-") - 1)) = (Mid(CboDestino.Text, 1, InStr(CboDestino.Text, "-") - 1)) Then
        MsgBox "Los valores de Origen y Destino deben ser diferentes"
        Exit Sub
    End If
    
    Inicial = (Left(Mid(txtDel.Text, 9), 7))
    Final = (Left(Mid(txtAl.Text, 9), 7))
    
'    If Final < Inicial Or Final = Inicial Then
'        MsgBox ("Los Rangos no son validos")
'        Exit Sub
'    End If
'    Else
        'cant = Final - Inicial
        cant = CInt(txtCantidad.Text)
        If Mid(CboOrigen.Text, 1, InStr(CboOrigen.Text, "-") - 1) = 0 Or Mid(CboDestino.Text, 1, InStr(CboOrigen.Text, "-") - 1) = 0 Then
            Set Cmd = New ADODB.Command
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@dFecha", adDate, adParamInput, 16, txtFecha.Text)
            Cmd.Parameters.Append Prm
            
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@nOrigen", adInteger, adParamInput, 8, Mid(CboOrigen.Text, 1, InStr(CboOrigen.Text, "-") - 1))
            Cmd.Parameters.Append Prm
            
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@nDestino", adInteger, adParamInput, 8, Mid(CboDestino.Text, 1, InStr(CboDestino.Text, "-") - 1))
            Cmd.Parameters.Append Prm
            
                    
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@nCantidad", adInteger, adParamInput, , cant)
            Cmd.Parameters.Append Prm
            
            
            
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@nDel", adVarChar, adParamInput, 50, txtDel.Text)
            Cmd.Parameters.Append Prm
            
            Set Prm = New ADODB.Parameter
            Set Prm = Cmd.CreateParameter("@nAl", adVarChar, adParamInput, 50, txtAl.Text)
            Cmd.Parameters.Append Prm
            
            oConec.AbreConexion
            Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
            Cmd.CommandType = adCmdStoredProc
            Cmd.CommandText = "ATM_RegistrarRemesa"
            Cmd.Execute
            
            cmdAcepta.Enabled = False
            Frame1.Enabled = False
        
            MsgBox "Las Remesas se registraron con exito", vbInformation, "Remesa de Tarjetas"
            'CerrarConexion
            oConec.CierraConexion
            Set Cmd = Nothing
            Set Prm = Nothing
        Else
            MsgBox "Debe escoger un Origen y/o Destino validos", vbInformation, "Remesa de Tarjetas"
                
        End If
'
'    End If
    
        
End Sub

Private Sub cmdNuevo_Click()
    'cmdAcepta.Enabled = True
    Frame1.Enabled = True
    txtFecha.Text = gdFecSis
    'CboOrigen.ListIndex = -1
    'CboDestino.ListIndex = -1
    txtCantidad.Text = "0"
    txtDel.Text = ""
    txtAl.Text = ""
    Me.cmdAcepta.Enabled = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    txtFecha.Text = gdFecSis
    'Call VerificaAgencia
    Call CargaDatos
    
    Frame1.Enabled = True
    CboDestino.Enabled = True
    CboDestino.Text = CboDestino.List(1)
    Call consultaAreaBvda
    CboOrigen.Enabled = False
    CboOrigen.Enabled = False
    cmdAcepta.Enabled = True
    
End Sub

Private Sub VerificaAgencia()
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodAge", adInteger, adParamInput, , gsCodAge)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@estado", adInteger, adParamInput, , 0)
    Cmd.Parameters.Append Prm

    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaAgenciaXCodigo"
    
    Set R = Cmd.Execute
    
    If Cmd.Parameters(1).Value = 0 Then
        rdbRemesa.Visible = False
    End If
'    'CboCtas.Clear
'    Do While Not R.EOF
'         CboOrigen.AddItem R!Codigo & "-" & R!Agencia
'
'       R.MoveNext
'    Loop
    'R.Close
    
    'cboOrigen.Text = cboOrigen.List(0)
    'cboDestino.Text = cboDestino.List(1)
    Set R = Nothing
    oConec.CierraConexion
    
End Sub

Private Sub CargaDatos()

    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    Set Prm = New ADODB.Parameter
    'Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 50, LblNumTarj.Caption)
    'Cmd.Parameters.Append Prm

    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaAgencias"
    
    Set R = Cmd.Execute
    'CboCtas.Clear
    Do While Not R.EOF
         CboOrigen.AddItem R!Codigo & "-" & R!Agencia
        CboDestino.AddItem R!Codigo & "-" & R!Agencia
       R.MoveNext
    Loop
    R.Close
    'cboOrigen.Text = cboOrigen.List(0)
    'cboDestino.Text = cboDestino.List(1)
    Set R = Nothing
    oConec.CierraConexion
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub rdbDevolucion_Click()
    Frame1.Enabled = True
    CboDestino.Enabled = False
    CboOrigen.Enabled = False
    CboDestino.Text = CboDestino.List(0)
    CboOrigen.List(1) = gsCodAge & "-" & gsNomAge
    CboOrigen.Text = CboOrigen.List(1)
    
    CboOrigen.Enabled = False
    cmdAcepta.Enabled = True
End Sub

Private Sub rdbRemesa_Click()
    Frame1.Enabled = True
    CboDestino.Enabled = True
    CboDestino.Text = CboDestino.List(1)
    Call consultaAreaBvda
    CboOrigen.Enabled = False
    CboOrigen.Enabled = False
    cmdAcepta.Enabled = True
End Sub
Private Sub consultaAreaBvda()
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    Set Prm = New ADODB.Parameter
    'Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 50, LblNumTarj.Caption)
    'Cmd.Parameters.Append Prm

    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaBVDA"
    
    Set R = Cmd.Execute
    'CboCtas.Clear
    Do While Not R.EOF
         CboOrigen.AddItem R!Codigo & "-" & R!Agencia
        
       R.MoveNext
    Loop
    R.Close
    
    CboOrigen.Text = CboOrigen.List(0)
    'cboDestino.Text = cboDestino.List(1)
    Set R = Nothing
    oConec.CierraConexion
    
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
    
        Call RecuperaRangosDETarjetasIngresadas(CInt(Me.txtCantidad.Text), sIni, sFin, nCant)
        
        Me.txtDel.Text = sIni
        Me.txtAl.Text = sFin
        
        If CInt(Me.txtCantidad.Text) <> nCant Then
            MsgBox "Solo Existen " & nCant & " Tarjetas Emitidas", vbInformation, "Aviso"
            Me.txtCantidad.Text = Trim(Str(nCant))
            Exit Sub
        End If
            
        Me.cmdAcepta.SetFocus
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

        Call RecuperaRangosDETarjetasIngresadas(CInt(Me.txtCantidad.Text), sIni, sFin, nCant)
        
        Me.txtDel.Text = sIni
        Me.txtAl.Text = sFin
        
      If CInt(Me.txtCantidad.Text) <> nCant Then
            MsgBox "Solo Existen " & nCant & " Tarjetas Ingresadas", vbInformation, "Aviso"
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

