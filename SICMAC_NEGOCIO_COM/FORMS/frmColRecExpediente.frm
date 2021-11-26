VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmColRecExpediente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperaciones - Expediente "
   ClientHeight    =   6975
   ClientLeft      =   1470
   ClientTop       =   1230
   ClientWidth     =   8685
   Icon            =   "frmColRecExpediente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8685
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   57
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6300
      TabIndex        =   56
      Top             =   6495
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7470
      TabIndex        =   55
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estudio Juridico "
      Height          =   1095
      Left            =   180
      TabIndex        =   46
      Top             =   2430
      Width           =   8295
      Begin VB.CommandButton cmdCambiarComision 
         Caption         =   "&Comisión"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7020
         TabIndex        =   52
         Top             =   630
         Width           =   1095
      End
      Begin VB.CommandButton cmdComi 
         Caption         =   "..."
         Height          =   285
         Left            =   2340
         TabIndex        =   50
         Top             =   630
         Width           =   285
      End
      Begin VB.TextBox txtComision 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   630
         Width           =   915
      End
      Begin VB.TextBox txtAbogado 
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   270
         Width           =   4875
      End
      Begin Sicmact.TxtBuscar AxCodAbogado 
         Height          =   285
         Left            =   1440
         TabIndex        =   47
         Top             =   270
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Estudio Juridico"
         Height          =   195
         Left            =   180
         TabIndex        =   53
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label17 
         Caption         =   "Comisión:"
         Height          =   195
         Left            =   180
         TabIndex        =   51
         Top             =   675
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   90
      TabIndex        =   35
      Top             =   5670
      Width           =   8475
      Begin VB.CheckBox chkComplementos 
         Caption         =   "Complementos"
         Height          =   285
         Left            =   180
         TabIndex        =   45
         Top             =   450
         Width           =   285
      End
      Begin VB.CheckBox chkMedProb 
         Caption         =   "Medios Prob"
         Height          =   285
         Left            =   5400
         TabIndex        =   44
         Top             =   180
         Width           =   195
      End
      Begin VB.CheckBox chkFundJuridico 
         Caption         =   "Fund. Juridico"
         Height          =   285
         Left            =   3600
         TabIndex        =   43
         Top             =   180
         Width           =   285
      End
      Begin VB.CheckBox chkFundHecho 
         Caption         =   "Fund. Hecho"
         Height          =   285
         Left            =   1890
         TabIndex        =   42
         Top             =   180
         Width           =   285
      End
      Begin VB.CheckBox chkPetitorio 
         Caption         =   "Petitorio"
         Height          =   285
         Left            =   180
         TabIndex        =   41
         Top             =   180
         Width           =   285
      End
      Begin VB.CommandButton cmdPetitorio 
         Caption         =   "&Petitorio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   450
         TabIndex        =   40
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdHecho 
         Caption         =   "&Fund.Hecho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   39
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdJuridico 
         Caption         =   "Fund.&Juridico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3870
         TabIndex        =   38
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdProbatorios 
         Caption         =   "&Med Probatorio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5670
         TabIndex        =   37
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdComplementos 
         Caption         =   "Com&plementos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   450
         TabIndex        =   36
         Top             =   450
         Width           =   1365
      End
   End
   Begin Sicmact.ActXCodCta AXCodCta 
      Height          =   465
      Left            =   180
      TabIndex        =   29
      Top             =   90
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   820
      Texto           =   "Crédito"
   End
   Begin VB.ComboBox cbDemanda 
      Height          =   315
      ItemData        =   "frmColRecExpediente.frx":030A
      Left            =   4860
      List            =   "frmColRecExpediente.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   1710
      Width           =   1635
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar ..."
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7110
      TabIndex        =   0
      Top             =   90
      Width           =   1005
   End
   Begin VB.ComboBox cbxTipoCob 
      Height          =   315
      ItemData        =   "frmColRecExpediente.frx":0354
      Left            =   1710
      List            =   "frmColRecExpediente.frx":035E
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1710
      Width           =   1815
   End
   Begin VB.Frame frmCliente 
      Height          =   1005
      Left            =   90
      TabIndex        =   8
      Top             =   540
      Width           =   8025
      Begin VB.TextBox txtCodPers 
         Height          =   285
         Left            =   1005
         TabIndex        =   28
         Tag             =   "TXTCODIGO"
         Top             =   255
         Width           =   1380
      End
      Begin VB.TextBox txtSaldoCapital 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   630
         Width           =   1185
      End
      Begin VB.TextBox txtMontoPrestamo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox txtTitular 
         Height          =   285
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   255
         Width           =   5520
      End
      Begin MSMask.MaskEdBox txtFechaCJ 
         Height          =   285
         Left            =   6570
         TabIndex        =   31
         Top             =   630
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Ingreso Cob.Jud"
         Height          =   240
         Left            =   5220
         TabIndex        =   32
         Top             =   630
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Capital"
         Height          =   195
         Left            =   2640
         TabIndex        =   12
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Prestamo"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.Frame frmJudi 
      Height          =   3930
      Left            =   90
      TabIndex        =   7
      Top             =   2160
      Width           =   8475
      Begin Sicmact.TxtBuscar AXCodJuzgado 
         Height          =   285
         Left            =   990
         TabIndex        =   33
         Top             =   1710
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtMontoPet 
         Height          =   285
         Left            =   7020
         TabIndex        =   30
         Top             =   2970
         Width           =   1275
      End
      Begin VB.TextBox txtFechaT 
         Height          =   285
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2520
         Width           =   1275
      End
      Begin VB.ComboBox cbxViaP 
         Height          =   315
         ItemData        =   "frmColRecExpediente.frx":037B
         Left            =   1170
         List            =   "frmColRecExpediente.frx":037D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2970
         Width           =   1635
      End
      Begin VB.TextBox txtSecretario 
         Height          =   285
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2580
         Width           =   2805
      End
      Begin VB.TextBox txtJuez 
         Height          =   285
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2145
         Width           =   2805
      End
      Begin VB.TextBox txtJuzgado 
         Height          =   285
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1710
         Width           =   2805
      End
      Begin VB.TextBox txtNroExp 
         Height          =   285
         Left            =   7020
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1710
         Width           =   1275
      End
      Begin MSMask.MaskEdBox txtFechaI 
         Height          =   285
         Left            =   7020
         TabIndex        =   6
         Top             =   2160
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.TxtBuscar AxCodJuez 
         Height          =   285
         Left            =   990
         TabIndex        =   34
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Sicmact.TxtBuscar AxCodSecre 
         Height          =   285
         Left            =   990
         TabIndex        =   54
         Top             =   2610
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label19 
         Caption         =   "Via Procesal"
         Height          =   285
         Left            =   180
         TabIndex        =   24
         Top             =   3060
         Width           =   915
      End
      Begin VB.Label Label13 
         Caption         =   "NroExpediente"
         Height          =   195
         Left            =   5850
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Juzgado"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Monto Petitorio"
         Height          =   195
         Left            =   5850
         TabIndex        =   17
         Top             =   2970
         Width           =   1185
      End
      Begin VB.Label Label15 
         Caption         =   "Fecha Termino"
         Height          =   195
         Left            =   5850
         TabIndex        =   16
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   5850
         TabIndex        =   15
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Secret."
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   2610
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Juez"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   2145
         Width           =   825
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Demanda"
      Height          =   195
      Left            =   3960
      TabIndex        =   26
      Top             =   1710
      Width           =   825
   End
   Begin VB.Label Label18 
      Caption         =   "Tipo de Cobranza"
      Height          =   195
      Left            =   180
      TabIndex        =   23
      Top             =   1710
      Width           =   1455
   End
End
Attribute VB_Name = "frmColREcExpediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* EXPEDIENTE DE RECUPERACIONES
'Archivo:  frmColRecExpediente.frm
'LAYG   :  01/08/2001.
'Resumen:  Nos permite hacer el mantenimiento del expediente de Recuperaciones
Option Explicit

Public Petitorio As String
Public Hecho As String
Public Juridico As String
Public Probatorios As String
Public Complementos As String

Public CodJuzgado As String
Public CodJuez As String
Public CodSecretario As String
Public AboResp As String
Public AboCont As String
Public CodAbo As String
Public Identidad As Long

Dim cJuz As String
Dim cJuez As String
Dim cSec As String
Dim cAboR As String
Dim cAboC As String
Dim cAbo As String
Dim fdoce As Boolean
Dim pConexionJud As New ADODB.Connection
Dim lbConexion As Boolean
Dim lbCambiaComision As Boolean
Dim vCodAboAntes As String
Dim vComiAntes As Long
Dim vNroMovDetExpedAntes As Long
Dim vFecIniAntes  As String


Sub LimpiaMemos()
    Petitorio = ""
    Hecho = ""
    Juridico = ""
    Probatorios = ""
    Complementos = ""
End Sub

Function FueIngresado(vCuenta As String, pConex As ADODB.Connection) As Boolean
Dim sql As String
Dim rs As New ADODB.Recordset
sql = " SELECT * FROM ExpedJud where cCodCta = '" & vCuenta & "'"
rs.Open sql, pConex, adOpenForwardOnly, adLockReadOnly, adCmdText

If RSVacio(rs) Then
   FueIngresado = False
Else
   FueIngresado = True
End If
End Function
Function DameNombre(vCodigo As String) As String
Dim sql As String
Dim rsNombre As New ADODB.Recordset
sql = " SELECT cNomPers FROM " & gcCentralPers & "persona WHERE cCodPers = '" & vCodigo & "'"
rsNombre.Open sql, pConexionJud, adOpenForwardOnly, adLockOptimistic, adCmdText
If RSVacio(rsNombre) Then
  DameNombre = ""
Else
  DameNombre = rsNombre!cNomPers
End If
CloseR rsNombre
End Function
Function ValidaDatos() As Boolean
ValidaDatos = True
If txtNroExp = "" Then
    ValidaDatos = False
    MsgBox "Falta Ingresar el Expediente"
    txtNroExp.SetFocus
    Exit Function
End If
End Function
Function FechasOK() As Boolean
Dim Mensaje As String
FechasOK = True
If Me.txtFechaI <> "__/__/____" Then
Mensaje = ValidaFecha(txtFechaI.Text)
    If Mensaje <> "" Then
        MsgBox Mensaje, vbInformation, "Aviso"
        Me.txtFechaI.SetFocus
        FechasOK = False
        Exit Function
    End If
Else
    MsgBox "Ingrese Fecha", vbInformation, "Aviso"
    Me.txtFechaI.SetFocus
    FechasOK = False
    Exit Function
End If

End Function

Function TipoProceso(Cuenta As String) As String
Dim sql As String
Dim rs As New ADODB.Recordset
sql = "SELECT cTipCJ FROM CredCJudi WHERE cCodCta = '" & Cuenta & "'"
rs.Open sql, pConexionJud, adOpenStatic, adLockOptimistic, adCmdText
If RSVacio(rs) Or IsNull(rs!cTipcj) Then
   TipoProceso = "J"
Else
   TipoProceso = rs!cTipcj
End If
CloseR rs
End Function

Function DemandaJud(Cuenta As String) As String
Dim sql As String
Dim rs As New ADODB.Recordset
sql = "SELECT cDemanda FROM CredCJudi WHERE cCodCta = '" & Cuenta & "'"
rs.Open sql, pConexionJud, adOpenStatic, adLockOptimistic, adCmdText
If RSVacio(rs) Or IsNull(rs!cDemanda) Then
   DemandaJud = "N"
Else
   DemandaJud = rs!cDemanda
End If
CloseR rs
End Function
Private Sub ActxCtaCred1_keypressEnter()
Dim sql As String
Dim sqlJud As String
Dim rs As New ADODB.Recordset
Dim rsJud As New ADODB.Recordset
Dim Cuenta As String
Dim Moneda As String
Dim Memo As Variant
Dim ValComi As String
Dim TProceso As String
Dim TDemanda As String
On Error GoTo ERROR
Cuenta = ActxCtaCred1.Text
Moneda = Mid(Cuenta, 6, 1)
Limpiar False
LimpiaMemos
sql = " SELECT PersCJudi.cRelaCta, Persona.cNomPers, CredCJudi.dFecJud, " _
    & " CredCJudi.nSaldCap,PersCJudi.cCodCta, PersCJudi.cCodPers FROM " & gcCentralPers & "Persona Persona INNER JOIN PersCJudi ON" _
    & " Persona.cCodPers = PersCJudi.cCodPers INNER JOIN CredCJudi ON PersCJudi.cCodCta" _
    & " = CredCJudi.cCodCta WHERE (PersCJudi.cRelaCta = 'TI') AND" _
    & " (PersCJudi.cCodCta = '" & Cuenta & "') "
    
rs.Open sql, pConexionJud, adOpenStatic, adLockOptimistic, adCmdText
If RSVacio(rs) Then
   MsgBox "La Cuenta no Existe o Todavia no Puede pasar a Judicial "
Else
   TProceso = TipoProceso(Cuenta)
   If TProceso = "J" Then
     Me.cbxTipoCob.ListIndex = 0
   Else
     Me.cbxTipoCob.ListIndex = 1
   End If
   TDemanda = DemandaJud(Cuenta)
   If TDemanda = "S" Then
     cbDemanda.ListIndex = 0
   Else
     cbDemanda.ListIndex = 1
   End If
   
   txtCodPers.Text = rs!cCodPers
   txtTitular.Text = rs!cNomPers
   txtFechaCJ.Text = Format(rs!dFecJud, "dd/mm/yyyy")
   txtMontoPrestamo.Text = DameDesembolso(Cuenta)
   txtSaldoCapital.Text = Format(rs!nSaldCap, "#,##0.00")
   
   cbxTipoCob.SetFocus
   cmdGrabar.Enabled = True
   CloseR rs
   'Datos del Expediente
   sqlJud = "SELECT * FROM ExpedJud WHERE cCodCta = '" & Cuenta & "'"
   rsJud.Open sqlJud, pConexionJud, adOpenKeyset, adLockOptimistic, adCmdText
   If RSVacio(rsJud) Then
      txtComision.Enabled = True  ' Habilita Comision
      lbCambiaComision = True
   Else  ' Si ya esta registrado
       'Identidad
        txtAbogado = DameNombre(rsJud!cCodAbog)
        CodAbo = rsJud!cCodAbog
        vCodAboAntes = rsJud!cCodAbog  ' Para Cambio de Comision
        txtAbogResp = DameNombre(rsJud!cAboResp)
        AboResp = rsJud!cAboResp
        txtJuez = DameNombre(rsJud!cCodJuez)
        CodJuez = rsJud!cCodJuez
        txtJuzgado = DameNombre(rsJud!cCodJuz)
        CodJuzgado = rsJud!cCodJuz
        txtSecretario = DameNombre(rsJud!cCodSecre)
        CodSecretario = rsJud!cCodSecre
        txtDomProcesal = rsJud!cdomproc
        txtFechaI.Text = Format(rsJud!dFecIni, "dd/mm/yyyy")
        txtFechaT.Text = IIf(IsNull(rsJud!dFecFin), "__/__/____", Format(rsJud!dFecFin, "dd/mm/yyyy"))
        txtMontoPet.Value = rsJud!nMonPetit
        txtNroExp = rsJud!cNumExp
        Select Case rsJud!cViaProce
            Case "A"
                    Me.cbxViaP.ListIndex = 0
            Case "C"
                    Me.cbxViaP.ListIndex = 1
            Case "E"
                    Me.cbxViaP.ListIndex = 2
            Case "S"
                    Me.cbxViaP.ListIndex = 3
        End Select
                    
        If Trim(rsJud!mPetitor) <> "" Then
           chkPetitorio.Value = 1
           Petitorio = rsJud!mPetitor
        Else
           chkPetitorio.Value = 0
        End If
        If Trim(rsJud!mHechos) <> "" Then
           chkFundHecho.Value = 1
           Hecho = rsJud!mHechos
        Else
           chkFundHecho.Value = 0
        End If
        If Trim(rsJud!mDatComp) <> "" Then
           chkComplementos.Value = 1
           Complementos = rsJud!mDatComp
        Else
           chkComplementos.Value = 0
        End If
        
        If Trim(rsJud!mFundJur) <> "" Then
           chkFundJuridico.Value = 1
           Juridico = rsJud!mFundJur
        Else
           chkFundJuridico.Value = 0
        End If
        If Trim(rsJud!mMedProb) <> "" Then
           chkMedProb.Value = 1
           Probatorios = rsJud!mMedProb
        Else
           chkMedProb.Value = 0
        End If
    End If
    CloseR rsJud ' Cierra
    sql = "SELECT nCodComi FROM CredCjudi WHERE cCodCta = '" & Cuenta & "'"
    rs.Open sql, pConexionJud, adOpenStatic, adLockOptimistic, adCmdText
    If RSVacio(rs) Or IsNull(rs(0)) Then
       txtComision.Text = ""
    Else
       Identidad = rs(0)
       ValComi = Comision(rs(0), pConexionJud)
       If Mid(ValComi, 1, 1) = "M" Then
          ValComi = "S/. " + Mid(ValComi, 2)
       Else
          ValComi = Mid(ValComi, 2) + "%"
       End If
       txtComision.Text = ValComi
       vComiAntes = Identidad
       ' Ubica Fecha Inicial y Nro Mov Abog
       sql = "SELECT dFecIni, nNroMovCta FROM ExpedJudDet " _
           & " WHERE cCodCta='" & Cuenta & "' AND cCodAbog='" _
           & vCodAboAntes & "' ORDER BY nNroMovCta DESC "
       rsJud.Open sql, pConexionJud, adOpenStatic, adLockOptimistic, adCmdText
       If RSVacio(rsJud) Then
          MsgBox "ERROR : No existe Det Exp Judicial "
       Else
          vNroMovDetExpedAntes = rsJud!nNroMovCta
          vFecIniAntes = rsJud!dFecIni
       End If
       CloseR rsJud ' Cierra
    End If
    
    cmdComi.Enabled = False  ' Deshabilita Comision
    cmdCambiarComision.Enabled = True
    CloseR rs
    If Moneda = "1" Then
        lblSoles.Caption = "Soles"
    Else
        lblSoles.Caption = "Dolares"
    End If
        
End If
Exit Sub

ERROR:
     MsgBox Err.Description
End Sub

Private Sub ActxCtaCred1_Validate(Cancel As Boolean)
If Len(Me.ActxCtaCred1.Text) < 12 Then
  Cancel = True
End If
End Sub



Private Sub AXCodCta_KeyPress(KeyAscii As Integer)

End Sub

Private Sub cbDemanda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtJuzgado.SetFocus
End If
End Sub

Private Sub cbxCuentas_Click()
    Me.ActxCtaCred1.Text = cbxCuentas.Text
End Sub
Private Sub cbxCuentas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Me.ActxCtaCred1.Enfoque 3
End If
End Sub
Private Sub cbxTipoCob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cbDemanda.SetFocus
End If
End Sub

Private Sub cbxViaP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdComi.SetFocus
End If
End Sub

Private Sub cmdBuscar_Click()

Dim loPers As UPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As DColRecCredito
Dim lrCreditos As ADODB.Recordset
Dim loCuentas As UProdPersona

On Error GoTo ControlError

Set loPers = New UPersona
    Set loPers = frmBuscaPersona.Inicio
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast

If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New DColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If

Set loCuentas = New UProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCambiarComision_Click()
 lbCambiaComision = True
 ' Capturo Datos antes de Cambio
 ' Llama a Comision
    frmJudComi.cmdImprimir.Caption = "&Aceptar"
    frmJudComi.txtCodPers = Me.CodAbo
    frmJudComi.txtNomPers = Me.txtAbogado
    frmJudComi.Show 1
    ActivaConexionJud
 
End Sub

Private Sub cmdcancelar_Click()
    Limpiar True
    cmdGrabar.Enabled = False
    Me.cmdCambiarComision.Enabled = False
    cbxTipoCob.ListIndex = -1
End Sub
Private Sub cmdComi_Click()
  lbCambiaComision = True
   frmJudComi.cmdImprimir.Caption = "&Aceptar"
   frmJudComi.txtCodPers = Me.CodAbo
   frmJudComi.txtNomPers = Me.txtAbogado
   frmJudComi.Show 1
End Sub
Private Sub cmdComplementos_Click()
 frmJudMemos.Cuenta = Me.ActxCtaCred1.Text
 frmJudMemos.Tipo = "DATOS COMPLEMENTARIOS"
 frmJudMemos.Caption = "Edición de Complementos"
 frmJudMemos.rtfMemo = Me.Complementos
 frmJudMemos.Com = True
 frmJudMemos.Show 1
End Sub

Private Sub cmdDetExpedJud_Click()
Dim SSQL As String
Dim rs As New ADODB.Recordset
SSQL = "SELECT E.cCodCta,E.cCodAbog,E.dfecIni,E.CCODABOG,C.NCODCOMI FROM ExpedJud E INNER JOIN CredCJudi C ON E.CCODCTA = C.CCODCTA "
rs.Open SSQL, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
Do While Not rs.EOF
   SSQL = "INSERT INTO ExpedJudDet VALUES ('" & rs!cCodCta & "',1,'" _
   & rs!cCodAbog & "','" & Format(rs!dFecIni, "mm/dd/yyyy hh:mm:ss") & "',null," & rs!nCodComi _
   & ",'" & Format(gdFecSis, "mm/dd/yyyy hh:mm:ss") & "','" & gsCodUser & "')"
   dbCmact.Execute SSQL
   rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub

Private Sub cmdExaminar_Click()

End Sub
Private Sub CmdGrabar_Click()
Dim Cuenta As String
Dim lsHoy As String
Dim sql As String
Dim SQL1 As String
Dim SQL2 As String
Dim SQL3 As String
Dim Moneda As String
Dim ViaP As String
Dim MontoPet As Currency
Dim lsDemanda As String * 1
Dim lsTipCJ As String * 1
On Error GoTo ERROR
If MsgBox("Desea Grabar el Expediente", vbQuestion + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If
Cuenta = ActxCtaCred1.Text

lsHoy = Format(gdFecSis, "mm/dd/yyyy") + " " + Format(Time(), "hh:mm:ss")
Moneda = Mid(Cuenta, 6, 1)
ViaP = Mid(Me.cbxViaP.Text, 1, 1)
If ValidaDatos = False Then
   Exit Sub
End If
If FechasOK = False Then
   Exit Sub
End If

MontoPet = CCur(Format(txtMontoPet.Text, "#0"))

If FueIngresado(Cuenta, pConexionJud) Then
    'Identidad
   sql = "UPDATE ExpedJud SET cCodJuz = '" & CodJuzgado & "',cNumExp = '" & Trim(txtNroExp) & "'," _
    & " cCodJuez = '" & CodJuez & "',cCodSecre = '" & CodSecretario & "'," _
    & " dFecIni = '" & Format(txtFechaI.Text, "mm/dd/yyyy") & "'," _
    & " cCodAbog =  '" & CodAbo & "',cDomProc = '" & txtDomProcesal & "'," _
    & " cAboResp = '" & AboResp & "'," _
    & " cAboCont = '" & AboCont & "',cDomCont = '',nMonPetit = " & MontoPet & "," _
    & " cMoneda  = '" & Moneda & "', mPetitor = '" & Petitorio & "'," _
    & " mHechos = '" & Hecho & "', mFundJur = '" & Juridico & "'," _
    & " mMedProb = '" & Probatorios & "'," _
    & " mDatComp = '" & Complementos & "',cViaProce = '" & ViaP & "'," _
    & " dFecMod = '" & Format(lsHoy, "mm/dd/yyyy") & "'" _
    & " WHERE cCodCta = '" & Cuenta & "'"
Else
   sql = "INSERT Expedjud (cCodCta,cCodJuz,cNumExp,cCodJuez,cCodSecre,dFecIni,cCodAbog," _
    & " cDomProc,cAboResp,cAboCont,cDomCont,nMonPetit,cMoneda,mPetitor,mHechos,mFundJur," _
    & " mMedProb,mDatComp,cViaProce,dFecFin,cForFin,cEstProc,dFecMod,cCodUsu)" _
    & " VALUES('" & Cuenta & "','" & CodJuzgado & "','" & Trim(txtNroExp) & "','" & CodJuez & "','" _
    & CodSecretario & " ','" & Format(txtFechaI, "mm/dd/yyyy") & "'," _
    & " '" & CodAbo & "','" & txtDomProcesal & "','" & AboResp & "','" & AboCont & "',''," _
    & " " & MontoPet & ",'" & Moneda & "','" & Petitorio & "','" & Hecho & "'," _
    & " '" & Juridico & "','" & Probatorios & "','" & Complementos & "','" & ViaP & "'," _
    & " null,null,null,'" & Format(lsHoy, "mm/dd/yyyy") & "','" & gsCodUser & "')"
End If
lsDemanda = IIf(cbDemanda.ListIndex = 0, "S", "N")
lsTipCJ = IIf(cbxTipoCob.ListIndex = 0, "J", "E")
SQL1 = "UPDATE CredCJudi SET nCodComi = " & Identidad & " , cTipCJ = '" & lsTipCJ & "'," _
     & " cDemanda = '" & lsDemanda & "' WHERE cCodCta = '" & Cuenta & "'"
     
If lbCambiaComision = True Then
   SQL2 = "UPDATE ExpedJudDet SET dFecFin = '" & Format(lsHoy, "mm/dd/yyyy hh:mm:ss") & "'," _
     & " dFecMod = '" & Format(lsHoy, "mm/dd/yyyy hh:mm:ss") & "'," _
     & " cCodUsu = '" & gsCodUser & "' " _
     & " WHERE cCodCta='" & Cuenta & "' AND cCodAbog='" & vCodAboAntes & "'" _
     & " AND nNroMovCta = " & vNroMovDetExpedAntes
End If

SQL3 = "INSERT INTO ExpedJudDet (cCodCta,nNroMovCta,cCodAbog,dFecIni,dFecFin, " _
     & " nCodComi, dFecMod,cCodUsu) VALUES ('" & Cuenta & "'," _
     & vNroMovDetExpedAntes + 1 & ",'" & CodAbo & "','" & Format(lsHoy, "mm/dd/yyyy hh:mm:ss") _
     & "',null," & Identidad & ",'" & Format(lsHoy, "mm/dd/yyyy hh:mm:ss") _
     & "','" & gsCodUser & "')"
pConexionJud.BeginTrans ' Inicia Transaccion
pConexionJud.Execute sql
pConexionJud.Execute SQL1
If lbCambiaComision = True Then
   pConexionJud.Execute SQL2
End If
pConexionJud.Execute SQL3
pConexionJud.CommitTrans ' Termina Transacc

cmdGrabar.Enabled = False
ActxCtaCred1.SetFocus
Limpiar True
ActxCtaCred1.SetFocus
cbxTipoCob.ListIndex = -1

lbCambiaComision = False
cmdCambiarComision = Enabled = False
Exit Sub
ERROR:
    MsgBox Err.Description
    pConexionJud.RollbackTrans
End Sub
Private Sub cmdHecho_Click()
 frmJudMemos.Cuenta = Me.ActxCtaCred1.Text
 frmJudMemos.Caption = "Edición de Fundamentos de Hecho"
 frmJudMemos.rtfMemo = Me.Hecho
 frmJudMemos.Hec = True
 frmJudMemos.Tipo = "FUNDAMENTOS DE HECHO "
 frmJudMemos.Show 1
 ActivaConexionJud
End Sub
Private Sub cmdJuridico_Click()
 frmJudMemos.Caption = "Edición de Fundamento Juridico"
 frmJudMemos.rtfMemo = Me.Juridico
 frmJudMemos.Cuenta = Me.ActxCtaCred1.Text
 frmJudMemos.Jur = True
 frmJudMemos.Tipo = "FUNDAMENTOS JURIDICOS "
 frmJudMemos.Show 1
 ActivaConexionJud
End Sub

Private Sub cmdPetitorio_Click()
 frmJudMemos.Cuenta = Me.ActxCtaCred1.Text
 frmJudMemos.Caption = "Edición de Petitorio"
 frmJudMemos.rtfMemo = Me.Petitorio
 frmJudMemos.Pet = True
 frmJudMemos.Tipo = "PETITORIO "
 frmJudMemos.Show 1
 ActivaConexionJud
End Sub
Private Sub cmdProbatorios_Click()
 frmJudMemos.Cuenta = Me.ActxCtaCred1.Text
 frmJudMemos.Caption = "Edición de Medios Probatorios"
 frmJudMemos.rtfMemo = Me.Probatorios
 frmJudMemos.Pro = True
 frmJudMemos.Tipo = "MEDIOS PROBATORIOS"
 frmJudMemos.Show 1
 ActivaConexionJud
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Limpiar(vAct As Boolean)

'lblSoles = ""
txtAbogado = ""
'txtDomProcesal = ""
txtFechaCJ.Mask = ""
txtFechaCJ.Text = ""
txtFechaCJ.Mask = "##/##/####"
txtFechaI.Mask = ""
txtFechaI.Text = ""
txtFechaT.Text = ""
txtJuez = ""
txtJuzgado = ""
'txtMontoPet.Value = 0
txtMontoPrestamo.Text = ""
txtNroExp = ""
txtSaldoCapital = ""
'txtDomProcesal = ""
'txtAbogResp = ""
txtSecretario = ""
txtTitular = ""
txtComision = ""
cbxViaP.ListIndex = -1
cbDemanda.ListIndex = -1
chkComplementos.Value = 0
chkFundHecho.Value = 0
chkFundJuridico.Value = 0
chkMedProb.Value = 0
chkPetitorio.Value = 0
If vAct = True Then
   AXCodCta.NroCuenta = Mid(gsCodAge, 4, 2) + ""
End If
AXCodCta.SetFocusCuenta
End Sub
Private Sub Form_Activate()
'If fdoce = True Then
' Me.ActxCtaCred1.Text = gsNueCod
'End If
fdoce = False
'frmJudMemos.Pet = False
'frmJudMemos.Com = False
'frmJudMemos.Pro = False
'frmJudMemos.Jur = False
'frmJudMemos.Hec = False
If Me.Petitorio = "" Then
   Me.chkPetitorio.Value = 0
Else
   Me.chkPetitorio.Value = 1
End If
If Me.Hecho = "" Then
   Me.chkFundHecho.Value = 0
Else
   Me.chkFundHecho.Value = 1
End If
If Me.Juridico = "" Then
   Me.chkFundJuridico.Value = 0
Else
   Me.chkFundJuridico.Value = 1
End If
If Me.Probatorios = "" Then
   Me.chkMedProb.Value = 0
Else
   Me.chkMedProb.Value = 1
End If
If Me.Complementos = "" Then
   Me.chkComplementos.Value = 0
Else
   Me.chkComplementos.Value = 1
End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 123 Then
'  Limpiar True
'  cbxTipoCob.ListIndex = -1
'  cmdGrabar.Enabled = False
'  frmRelaCredJudi.Show 1
'  If Len(Me.ActxCtaCred1.Text) > 0 Then
'    Me.ActxCtaCred1.completo
'    Me.ActxCtaCred1.Enfoque 3
'  End If
'  fdoce = True
'End If
End Sub

Private Sub Form_Load()

    'CargaForma Me.cbxViaP, "80", pConexionJud
    Petitorio = ""
    Hecho = ""
    Juridico = ""
    Probatorios = ""
    Complementos = ""
    'gsNueCod = ""
End Sub




Private Sub txtAbogado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then
  frmBusIdentJudicial.Pant = "EXPEDIENTES"
  frmBusIdentJudicial.ROL = "001"
  frmBusIdentJudicial.Abo = 1
  frmBusIdentJudicial.Caption = "Lista de Abogados"
  frmBusIdentJudicial.Show 1
  ActivaConexionJud
End If
If KeyCode = 46 Then
    CodAbo = ""
    Me.txtAbogado.Text = ""
End If
End Sub

Private Sub txtAbogado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtDomProcesal.SetFocus
End Sub


Private Sub txtAbogResp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then
  frmBusIdentJudicial.Pant = "EXPEDIENTES"
  frmBusIdentJudicial.ROL = "001"
  frmBusIdentJudicial.Abo = 2
  frmBusIdentJudicial.Caption = "Lista de Abogados"
  frmBusIdentJudicial.Show 1
  ActivaConexionJud
End If
If KeyCode = 46 Then
AboResp = ""
Me.txtAbogResp.Text = ""
End If
End Sub

Private Sub txtAbogResp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Me.cbxViaP.SetFocus
End If
End Sub

Private Sub txtComision_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Me.txtNroExp.SetFocus
End If
End Sub

Private Sub txtDomProcesal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtAbogResp.SetFocus
End Sub

Private Sub txtFechaCJ_GotFocus()
fEnfoque txtFechaCJ
End Sub

Private Sub txtFechaCJ_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.cbxTipoCob.SetFocus
End If
End Sub

Private Sub txtFechaI_GotFocus()
fEnfoque txtFechaI
End Sub

Private Sub txtFechaI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtMontoPet.SetFocus
End Sub

Private Sub txtFechaT_GotFocus()
fEnfoque txtFechaT
End Sub

Private Sub txtJuez_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then
  frmBusIdentJudicial.Pant = "EXPEDIENTES"
  frmBusIdentJudicial.ROL = "002"
  frmBusIdentJudicial.Caption = "Lista de Jueces"
  frmBusIdentJudicial.Show 1
  ActivaConexionJud
End If
If KeyCode = 46 Then
CodJuez = ""
Me.txtJuez.Text = ""
End If

End Sub

Private Sub txtJuez_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtSecretario.SetFocus
End Sub

Private Sub txtJuzgado_GotFocus()
lblMensaje.Visible = True
End Sub

Private Sub txtJuzgado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then
  frmBusIdentJudicial.Pant = "EXPEDIENTES"
  frmBusIdentJudicial.ROL = "004"
  frmBusIdentJudicial.Caption = "Lista de Juzgados"
  frmBusIdentJudicial.Show 1
  ActivaConexionJud
End If
If KeyCode = 46 Then
  CodJuzgado = ""
  Me.txtJuzgado.Text = ""
End If

End Sub

Private Sub txtJuzgado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtJuez.SetFocus
End Sub

Private Sub txtMontoPet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.cmdPetitorio.SetFocus
End Sub

Private Sub txtNroExp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtFechaI.SetFocus
End Sub

Private Sub txtSecretario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then
  frmBusIdentJudicial.Pant = "EXPEDIENTES"
  frmBusIdentJudicial.ROL = "003"
  frmBusIdentJudicial.Caption = "Lista de Secretarios"
  frmBusIdentJudicial.Show 1
  ActivaConexionJud
End If
If KeyCode = 46 Then
  CodSecretario = ""
  Me.txtSecretario.Text = ""
End If

End Sub

Private Sub txtSecretario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAbogado.SetFocus
End Sub

