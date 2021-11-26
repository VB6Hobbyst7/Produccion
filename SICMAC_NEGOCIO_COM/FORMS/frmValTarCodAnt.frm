VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmValTarCodAnt 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmValTarCodAnt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1335
      Width           =   1035
   End
   Begin VB.Frame fraDato 
      Caption         =   "Dato"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1155
      Left            =   1860
      TabIndex        =   7
      Top             =   120
      Width           =   3195
      Begin MSMask.MaskEdBox txtCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-###-###-#######A"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTarjeta 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-####-####-#### "
         PromptChar      =   "_"
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1380
      TabIndex        =   4
      Top             =   1335
      Width           =   1035
   End
   Begin VB.Frame fraTipoBusq 
      Caption         =   "Buscar por :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1155
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   1755
      Begin VB.OptionButton optTipoBusq 
         Caption         =   "Tarjeta"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   720
         Width           =   1035
      End
      Begin VB.OptionButton optTipoBusq 
         Caption         =   "Código Antiguo"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmValTarCodAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nProducto As Producto
Dim sCuenta As String
Dim pbAbono As Boolean

Public Function Inicia(ByVal nProd As Producto, Optional bTarjeta As Boolean = False, Optional bAbono As Boolean = False) As String
If bTarjeta Then
    optTipoBusq(1).Visible = True
    txtTarjeta.Visible = True
    Me.Caption = "Relación Cuenta Antigua - Nuevo y Tarjeta Magnética"
Else
    optTipoBusq(1).Visible = False
    txtTarjeta.Visible = False
    Me.Caption = "Relación Cuenta Antigua - Nuevo"
End If
nProducto = nProd
pbAbono = bAbono
Select Case nProducto
'ARCV 17-02-2007
'    Case gCapAhorros
'        txtCuenta.Mask = "###-###-###-#########"
'        txtCuenta.Text = "___-___-___-_________"
'    Case gCapCTS, gCapPlazoFijo
'        txtCuenta.Mask = "###-###-###-#########"
'        txtCuenta.Text = "___-___-___-_________"
'    Case gColConsuPrendario
'        txtCuenta.Mask = "###-###-###-#########"
'        txtCuenta.Text = "___-___-___-_________"
'    Case Else
'        txtCuenta.Mask = "###-###-###-#########"
'        txtCuenta.Text = "___-___-___-_________"
    Case gCapAhorros
        txtCuenta.Mask = "###-###-##-#########A"
        txtCuenta.Text = "___-___-__-__________"
    Case gCapCTS, gCapPlazoFijo
        txtCuenta.Mask = "###-###-##-#########A"
        txtCuenta.Text = "___-___-__-__________"
    Case gColConsuPrendario
        txtCuenta.Mask = "###-###-##-#########A"
        txtCuenta.Text = "___-___-__-__________"
    Case Else
        txtCuenta.Mask = "###-###-##-#########A"
        txtCuenta.Text = "___-___-__-__________"
'--------
End Select
sCuenta = ""
Me.Show 1
Inicia = sCuenta
End Function

Private Sub CmdAceptar_Click()
Dim nEstado As Integer
If optTipoBusq(0).value = True Then
    Dim clsGen As COMDConstSistema.DCOMGeneral   'DGeneral
    Dim sCuentaAnt As String
    sCuentaAnt = Trim(Replace(txtCuenta.Text, "-", "", 1, , vbTextCompare))
    sCuentaAnt = Trim(Replace(sCuentaAnt, "_", "", 1, , vbTextCompare))
    If sCuentaAnt <> "" Then
        Set clsGen = New COMDConstSistema.DCOMGeneral 'DGeneral
        sCuenta = clsGen.GetCuentaNueva(sCuentaAnt)
        Set clsGen = Nothing
    Else
        MsgBox "Cuenta Incorrecta, por favor digite una cuenta de 18 dígitos", vbInformation, "Aviso"
        txtCuenta.SetFocus
        Set clsGen = Nothing
        Exit Sub
    End If
ElseIf optTipoBusq(1).value = True Then
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales  'NCapMantenimiento
    Dim rsTarj As ADODB.Recordset
    Dim sTarjeta As String, sPersona As String
    sTarjeta = Trim(Replace(txtTarjeta, "-", "", 1, , vbTextCompare))
    sTarjeta = Trim(Replace(sTarjeta, "_", "", 1, , vbTextCompare))
    
    If Len(sTarjeta) <> 16 Then
        MsgBox "Tarjeta Incorrecta, por favor digite una tarjeta de 16 dígitos", vbInformation, "Aviso"
        txtTarjeta.SetFocus
        Exit Sub
    End If
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento

    'Dim rsTarj As ADODB.Recordset
    Set rsTarj = New ADODB.Recordset
    Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
    Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
    Set rsTarj = ObjTarj.Get_Datos_Tarj(sTarjeta)

    If rsTarj.EOF And rsTarj.BOF Then
        MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
        Set ObjTarj = Nothing
        'Me.Caption = sCaption
        Exit Sub
    Else
        nEstado = rsTarj("nEstado")
        If nEstado = gCapTarjEstBloqueada Or nEstado = gCapTarjEstCancelada Then
            If nEstado = gCapTarjEstBloqueada Then
                MsgBox "Número de Tarjeta Bloqueada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
            ElseIf nEstado = gCapTarjEstCancelada Then
                MsgBox "Número de Tarjeta Cancelada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
            End If
            'Me.Caption = sCaption
            Set ObjTarj = Nothing
            Exit Sub
        End If
        

    End If

    Dim rsPers As New ADODB.Recordset
    Dim sCta As String, sProducto As String, sMoneda As String
    Dim clsCuenta As UCapCuenta
                            
    'Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    'Dim rsTarj As New ADODB.Recordset

    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    'Set rsTarj = clsMant.GetTarjetaCuentas(sNumTar, nProducto)
    
    Set rsPers = clsMant.GetCuentasPersona(rsTarj("cPersCod"), nProducto)
    'Set rsPers = clsMant.GetTarjetaCuentas(sNumTar, nProducto)
    'Set clsMant = Nothing

    
    
'    Set rsTarj = clsMant.GetTarjetaCuentas(sTarjeta, nProducto)
'    'Set rsTarj = clsMant.GetPersonaTarj(sTarjeta, nProducto)
'    If rsTarj.EOF And rsTarj.BOF Then
'        MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
'        Set clsMant = Nothing
'        Exit Sub
'    Else
        'Dim rsPers As ADODB.Recordset
        'Dim sCta As String, sRelac As String, sEstado As String
        Dim sRelac As String, sEstado As String
        'Dim clsCuenta As UCapCuenta 'UCapCuentas
                
        sPersona = rsTarj("cPersCod")
        If pbAbono = False Then
            Set rsPers = clsMant.GetCuentasPersona(sPersona, nProducto, True, True)
        Else
            Set rsPers = clsMant.GetCuentasPersona(sPersona, nProducto, True)
        End If
        
        Set clsMant = Nothing
        If Not (rsPers.EOF And rsPers.EOF) Then
            Do While Not rsPers.EOF
                sCta = rsPers("cCtaCod")
                sRelac = rsPers("cRelacion")
                sEstado = Trim(rsPers("cEstado"))
                frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
                rsPers.MoveNext
            Loop
            Set clsCuenta = New UCapCuenta 'UCapCuentas
            Set clsCuenta = frmCapMantenimientoCtas.Inicia
            sCuenta = clsCuenta.sCtaCod
            Set clsCuenta = Nothing
        Else
            MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
        End If
        rsPers.Close
        Set rsPers = Nothing
    End If
    Set rsTarj = Nothing
    'Set clsMant = Nothing
'End If
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub optTipoBusq_Click(Index As Integer)
If Index = 0 Then
    txtTarjeta.Visible = False
    txtCuenta.Visible = True
    txtCuenta.SetFocus
ElseIf Index = 1 Then
    txtCuenta.Visible = False
    txtTarjeta.Visible = True
    txtTarjeta.SetFocus
End If
End Sub

Private Sub txtCuenta_Change()
'txtCuenta.Text = UCase(txtCuenta.Text)
End Sub

Private Sub txtCuenta_GotFocus()
With txtCuenta
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii > 57 Then
    KeyAscii = Letras(KeyAscii)
End If
End Sub

Private Sub txtTarjeta_GotFocus()
With txtTarjeta
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub
