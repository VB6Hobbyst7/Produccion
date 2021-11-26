VERSION 5.00
Begin VB.Form frmColRecExped 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperaciones - Expediente "
   ClientHeight    =   5985
   ClientLeft      =   1470
   ClientTop       =   1230
   ClientWidth     =   8505
   Icon            =   "frmColRecExped.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8505
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdExpRelacion 
      Caption         =   "&Relación con Otros Expedientes"
      Enabled         =   0   'False
      Height          =   400
      Left            =   2960
      TabIndex        =   23
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraInforme 
      Height          =   2055
      Left            =   6480
      TabIndex        =   50
      Top             =   3120
      Width           =   1935
      Begin VB.CheckBox chkComplementos 
         Caption         =   "Complementos"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Top             =   1560
         Width           =   195
      End
      Begin VB.CheckBox chkMedProb 
         Caption         =   "Medios Prob"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   54
         Top             =   1230
         Width           =   195
      End
      Begin VB.CheckBox chkFundJuridico 
         Caption         =   "Fund. Juridico"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   53
         Top             =   900
         Width           =   195
      End
      Begin VB.CheckBox chkFundHecho 
         Caption         =   "Fund. Hecho"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   52
         Top             =   570
         Width           =   195
      End
      Begin VB.CheckBox chkPetitorio 
         Caption         =   "Petitorio"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   195
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
         Left            =   420
         TabIndex        =   18
         Top             =   240
         Width           =   1395
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
         Left            =   420
         TabIndex        =   19
         Top             =   570
         Width           =   1395
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
         Left            =   420
         TabIndex        =   20
         Top             =   915
         Width           =   1395
      End
      Begin VB.CommandButton cmdProbatorios 
         Caption         =   "&M Probatorio"
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
         Left            =   420
         TabIndex        =   21
         Top             =   1245
         Width           =   1395
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
         Left            =   420
         TabIndex        =   22
         Top             =   1575
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1440
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
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
      Height          =   400
      Left            =   4920
      TabIndex        =   2
      Top             =   5520
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
      Height          =   400
      Left            =   6120
      TabIndex        =   3
      Top             =   5520
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
      Height          =   400
      Left            =   7320
      TabIndex        =   4
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame fraEstudioJur 
      Caption         =   "Estudio Juridico "
      Enabled         =   0   'False
      Height          =   1005
      Left            =   180
      TabIndex        =   36
      Top             =   2100
      Width           =   8175
      Begin VB.CommandButton cmdComi 
         Caption         =   "..."
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
         Left            =   2820
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin SICMACT.TxtBuscar AxCodAbogado 
         Height          =   285
         Left            =   1425
         TabIndex        =   10
         Top             =   240
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
         TipoBusqueda    =   3
      End
      Begin VB.Label lblComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   56
         Top             =   600
         Width           =   1395
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNomAbogado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3240
         TabIndex        =   45
         Top             =   240
         Width           =   4575
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Caption         =   "Estudio Juridico"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label17 
         Caption         =   "Comisión:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   735
      End
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   465
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   820
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar ..."
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   90
      Width           =   1005
   End
   Begin VB.Frame fraCliente 
      Height          =   1005
      Left            =   120
      TabIndex        =   25
      Top             =   540
      Width           =   8355
      Begin VB.Label lblFecIngRecup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6480
         TabIndex        =   46
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMontoPrestamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   44
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2400
         TabIndex        =   43
         Top             =   240
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCodPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   42
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblSaldoCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3720
         TabIndex        =   41
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Ingreso Recup."
         Height          =   240
         Left            =   5220
         TabIndex        =   35
         Top             =   630
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Capital"
         Height          =   195
         Left            =   2640
         TabIndex        =   28
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Prestamo"
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   90
         TabIndex        =   26
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.Frame fraDatos 
      Enabled         =   0   'False
      Height          =   3720
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   8355
      Begin VB.ComboBox cboTipMedCau 
         Height          =   315
         ItemData        =   "frmColRecExped.frx":030A
         Left            =   3840
         List            =   "frmColRecExped.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   3120
         Width           =   2475
      End
      Begin VB.ComboBox CboProcesos 
         Height          =   315
         ItemData        =   "frmColRecExped.frx":030E
         Left            =   6480
         List            =   "frmColRecExped.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   180
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cbxTipoCob 
         Height          =   315
         ItemData        =   "frmColRecExped.frx":0312
         Left            =   1380
         List            =   "frmColRecExped.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   180
         Width           =   1815
      End
      Begin VB.ComboBox cbDemanda 
         Height          =   315
         ItemData        =   "frmColRecExped.frx":0316
         Left            =   3960
         List            =   "frmColRecExped.frx":0318
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   180
         Width           =   1635
      End
      Begin SICMACT.TxtBuscar AXCodJuzgado 
         Height          =   285
         Left            =   900
         TabIndex        =   12
         Top             =   1710
         Width           =   1695
         _ExtentX        =   2990
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
         TipoBusqueda    =   3
      End
      Begin VB.TextBox txtMontoPet 
         Height          =   285
         Left            =   900
         TabIndex        =   17
         Top             =   3200
         Width           =   1695
      End
      Begin VB.ComboBox cbxViaP 
         Height          =   315
         ItemData        =   "frmColRecExped.frx":031A
         Left            =   900
         List            =   "frmColRecExped.frx":0321
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtNroExp 
         Height          =   285
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   16
         Top             =   2760
         Width           =   2400
      End
      Begin SICMACT.TxtBuscar AxCodJuez 
         Height          =   285
         Left            =   900
         TabIndex        =   13
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
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
         TipoBusqueda    =   3
      End
      Begin SICMACT.TxtBuscar AxCodSecre 
         Height          =   285
         Left            =   900
         TabIndex        =   14
         Top             =   2400
         Width           =   1695
         _ExtentX        =   2990
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
         TipoBusqueda    =   3
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Medida Cautelar :"
         Height          =   435
         Left            =   2640
         TabIndex        =   59
         Top             =   3060
         Width           =   1185
      End
      Begin VB.Label LblProceso 
         Caption         =   "Procesos"
         Height          =   195
         Left            =   5680
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblNomSecre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2640
         TabIndex        =   49
         Top             =   2400
         Width           =   3615
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNomJuez 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2640
         TabIndex        =   48
         Top             =   2040
         Width           =   3615
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNomJuzgado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2640
         TabIndex        =   47
         Top             =   1710
         Width           =   3615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         Caption         =   "Tipo de Cobranza"
         Height          =   195
         Left            =   60
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Demanda"
         Height          =   195
         Left            =   3240
         TabIndex        =   39
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Demanda"
         Height          =   405
         Left            =   120
         TabIndex        =   34
         Top             =   2700
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "NroExpediente"
         Height          =   195
         Left            =   2640
         TabIndex        =   33
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Juzgado"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto Petitorio"
         Height          =   435
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Secretario"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Juez"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmColRecExped"
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

Dim fsEstudioJuridicoCod As String
Dim fsJuzgadoCod As String
Dim fsJuezCod As String
Dim fsSecretarioCod As String

Dim fbCambiaComision As Boolean
Dim fnComisionCodAnt As Integer
Dim fnComisionCodNue As Integer

Dim fsPetitorio As String
Dim fsHecho As String
Dim fsJuridico As String
Dim fsMProbatorios As String
Dim fsComplementos As String

Dim frsPers As ADODB.Recordset
Dim fsNuevoEdita As String
Dim fbExisteExpediente As Boolean
Dim objPista As COMManejador.Pista  ''*** PEAC 20090126

Function ValidaDatos() As Boolean
ValidaDatos = True
If txtNroExp = "" Then
    ValidaDatos = False
    MsgBox "Falta Ingresar el Expediente"
    txtNroExp.SetFocus
    Exit Function
End If
End Function

Private Sub AxCodAbogado_EmiteDatos()
    lblNomAbogado.Caption = AxCodAbogado.psDescripcion
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaDatos(AXCodCta.NroCuenta)
End Sub

Private Sub BuscaDatos(ByVal psNroCredito As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValCredito As COMNColocRec.NColRecValida
Dim lsmensaje As String
Dim i As Integer
Dim lnTipoProceso As String
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValCredito = New COMNColocRec.NColRecValida
        Set lrValida = loValCredito.nValidaExpediente(psNroCredito, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             CmdExpRelacion.Enabled = False
             Exit Sub
        End If

        If lrValida Is Nothing Then
            EstadoBotones False, False, True
        Else
            If lrValida Is Nothing Then ' Hubo un Error
                Limpiar
                Set lrValida = Nothing
                Exit Sub
            End If
            
            fbExisteExpediente = IIf(lrValida!nControl = 9, False, True)
            
            If lrValida!nTipCJ = gColRecTipCobJudicial Then
                Me.cbxTipoCob.ListIndex = 0
            Else
                Me.cbxTipoCob.ListIndex = 1
            End If
            
            
            If lrValida!nDemanda = gColRecDemandaNo Then
                Me.cbDemanda.ListIndex = 1
            Else
                Me.cbDemanda.ListIndex = 0
            End If
            If gsProyectoActual = "H" Then
                lnTipoProceso = lrValida!TipoProceso
                Call UbicaCombo(Me.CboProcesos, lnTipoProceso, True)
            End If
                
            lblCodPers.Caption = lrValida!cCodClie
            lblNomPers.Caption = lrValida!cNomClie
            lblMontoPrestamo = Format(lrValida!nMontoCol, "#,##0.00")
            lblSaldoCapital = Format(lrValida!nSaldo, "#,##0.00")
            lblFecIngRecup = Format(lrValida!dIngRecup, "dd/mm/yyyy")
            
            AxCodAbogado.Text = lrValida!cCodAbog
            lblNomAbogado.Caption = lrValida!cNomAbog
            
            If Len(lrValida!nTipComis) > 0 Then
                If lrValida!nTipComis = 1 Then
                    lblComision.Caption = Format(lrValida!nComisionValor, "#0.00") & " Mon"
                Else
                    lblComision.Caption = Format(lrValida!nComisionValor, "#0.00") & " % "
                End If
                fnComisionCodAnt = lrValida!nComisionCod
                fnComisionCodNue = lrValida!nComisionCod
            Else
                lblComision.Caption = ""
                fnComisionCodAnt = 0
                fnComisionCodNue = 0
            End If
                
            Me.AXCodJuzgado.Text = lrValida!cCodJuzg
            Me.lblNomJuzgado.Caption = lrValida!cNomJuzg
            Me.AxCodJuez.Text = lrValida!cCodJuez
            Me.lblNomJuez.Caption = lrValida!cNomJuez
            Me.AxCodSecre.Text = lrValida!cCodSecre
            Me.lblNomSecre.Caption = lrValida!cNomSecre
            Me.txtMontoPet.Text = Format(lrValida!nMonPetit, "#0.00")
            Me.txtNroExp.Text = lrValida!cNumExp
            
            Call UbicaCombo(Me.cbxViaP, lrValida!nViaProce, True)
            
            Call UbicaCombo(cboTipMedCau, lrValida!nTipMedCau, True) 'DAOR 20070417
            
            fsPetitorio = IIf(Trim(lrValida!mPetit) <> "", lrValida!mPetit, "")
            fsHecho = IIf(Trim(lrValida!mHechos) <> "", lrValida!mHechos, "")
            fsJuridico = IIf(Trim(lrValida!mFundJur) <> "", lrValida!mFundJur, "")
            fsMProbatorios = IIf(Trim(lrValida!mMedProb) <> "", lrValida!mMedProb, "")
            fsComplementos = IIf(Trim(lrValida!mDatComp) <> "", lrValida!mDatComp, "")
            
            EstadoBotones False, True, False
            AXCodCta.Enabled = False
            Me.fraInforme.Enabled = True
            CmdExpRelacion.Enabled = True
        End If
        Set lrValida = Nothing
    Set loValCredito = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AxCodJuez_EmiteDatos()
    lblNomJuez.Caption = AxCodJuez.psDescripcion
    If Len(lblNomJuez.Caption) > 0 Then
        AxCodSecre.SetFocus
    End If
End Sub

Private Sub AXCodJuzgado_EmiteDatos()
    lblNomJuzgado.Caption = AXCodJuzgado.psDescripcion
    If Len(lblNomJuzgado.Caption) > 0 Then
        AxCodJuez.SetFocus
    End If
End Sub

Private Sub AxCodSecre_EmiteDatos()
    lblNomSecre.Caption = AxCodSecre.psDescripcion
    If Len(lblNomSecre.Caption) > 0 Then
        cbxViaP.SetFocus
    End If
    
End Sub

Private Sub cbDemanda_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    txtJuzgado.SetFocus
'End If
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

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito 'DColRecCredito
Dim lrCreditos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast

Limpiar ' True
EstadoBotones True, False, False
cbxTipoCob.ListIndex = -1
Me.fraDatos.Enabled = False
Me.fraEstudioJur.Enabled = False
AXCodCta.Enabled = True
    
If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If


Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub


Private Sub cmdCancelar_Click()
    Limpiar ' True
    EstadoBotones True, False, False
    EstadoControles False, False, False
    AXCodCta.Enabled = True
End Sub

Private Sub cmdComi_Click()
Dim loComision As UColRecComisionSelecciona
'lbCambiaComision = True
Set loComision = New UColRecComisionSelecciona
    Set loComision = frmColRecComisionSelecciona.Inicio(Me.AxCodAbogado.Text, Me.lblNomAbogado)
    If loComision.nComisionCod > 0 Then
        fnComisionCodNue = loComision.nComisionCod
        lblComision.Caption = loComision.nComisionValor
        If loComision.nComisionTipo = 1 Then 'Moneda
            lblComision.Caption = lblComision.Caption & " Mon"
        Else
            lblComision.Caption = lblComision.Caption & " % "
        End If
        AXCodJuzgado.SetFocus
    End If
Set loComision = Nothing
End Sub

Private Sub cmdComplementos_Click()
Dim loMemo  As UColRecMemos
Set loMemo = New UColRecMemos
    Set loMemo = frmColRecMemos.Inicio("C", fsComplementos, AXCodCta.NroCuenta, Me.lblNomPers)
        fsComplementos = loMemo.sMemo
Set loMemo = Nothing
End Sub

Private Sub cmdEditar_Click()

EstadoBotones False, False, True
EstadoControles True, True, True
CmdExpRelacion.Enabled = False
fsNuevoEdita = "E"

End Sub

Private Sub CmdExpRelacion_Click()
    If Len(Me.AXCodCta.NroCuenta) = 18 Then
        frmColRecExpedRelacionados.Inicia Trim(Me.AXCodCta.NroCuenta)
    End If
End Sub

Private Sub cmdGrabar_Click()

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito
Dim lbCambiaComision As Boolean

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
'On Error GoTo ControlError
Dim lnTipCobJud As Integer
Dim lnDemanda As Integer
Dim lnTipProceso As Integer
Dim lnViaProce As Integer
Dim lnTipMedCau As Integer 'DAOR 20070417
'validar numero de cuenta
If Len(AXCodCta.NroCuenta) <> 18 Then
    MsgBox "Ingrese un Nro. Cuenta Valido", vbInformation, "Aviso"
    Exit Sub
End If
'Valida Datos a Grabar
If Me.cbDemanda.ListIndex = 0 Then
    lnTipCobJud = gColRecTipCobJudicial
    lnDemanda = gColRecDemandaSi
Else
    lnTipCobJud = gColRecTipCobExtraJudi
    lnDemanda = gColRecDemandaNo
End If

'Via Procesal
If Len(Right(cbxViaP.Text, 2)) = 0 Then
    MsgBox "Seleccione un Tipo de Demanda", vbInformation, "Aviso"
    'cbxViaP.SetFocus
    Exit Sub
End If

' tipos de Proceso
If gsProyectoActual = "H" Then
    If Len(Right(Trim(CboProcesos.Text), 2)) = 0 Then
        MsgBox "Seleccione Tipo Proceso", vbInformation, "Aviso"
        Exit Sub
    End If
    lnTipProceso = Right(Trim(CboProcesos.Text), 2)
End If

lnViaProce = Right(cbxViaP.Text, 2)

If cboTipMedCau.ListIndex > 0 Then
    lnTipMedCau = Right(cboTipMedCau.Text, 2) 'DAOR 20070417
End If

CargaPersonasRelacionadas
lbCambiaComision = IIf(fnComisionCodNue <> fnComisionCodAnt, True, False)
If fValidaData = False Then
    Exit Sub
End If
If MsgBox(" Grabar Registro de Expediente ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        
    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        
    Set loGrabar = New COMNColocRec.NCOMColRecCredito
    If gsProyectoActual = "H" Then
        Call loGrabar.nRegistraExpedienteRecup(AXCodCta.NroCuenta, lsFechaHoraGrab, lsMovNro, _
               fnComisionCodNue, lnTipCobJud, Me.txtNroExp.Text, CCur(val(Me.txtMontoPet.Text)), Mid(AXCodCta.NroCuenta, 9, 1), _
               frsPers, fsPetitorio, fsHecho, fsComplementos, fsMProbatorios, fsComplementos, lnViaProce, _
               0, lbCambiaComision, fbExisteExpediente, False, lnTipProceso, gsProyectoActual)
    Else
       Call loGrabar.nRegistraExpedienteRecup(AXCodCta.NroCuenta, lsFechaHoraGrab, lsMovNro, _
               fnComisionCodNue, lnTipCobJud, Me.txtNroExp.Text, CCur(val(Me.txtMontoPet.Text)), Mid(AXCodCta.NroCuenta, 9, 1), _
               frsPers, fsPetitorio, fsHecho, fsComplementos, fsMProbatorios, fsComplementos, lnViaProce, _
               0, lbCambiaComision, fbExisteExpediente, False, , gsProyectoActual, lnDemanda, lnTipMedCau)   'ARCV 25-01-2007 , 'DAOR 20070417 lnTipMedCau
            
            ''*** PEAC 20090126
            objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , AXCodCta.NroCuenta, gCodigoCuenta
                                  
    End If
    Set loGrabar = Nothing
    
    EstadoBotones True, True, False
    EstadoControles False, False, False
    
    CmdExpRelacion.Enabled = True
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdHecho_Click()
Dim loMemo  As UColRecMemos
Set loMemo = New UColRecMemos
    Set loMemo = frmColRecMemos.Inicio("H", fsHecho, AXCodCta.NroCuenta, Me.lblNomPers)
        fsHecho = loMemo.sMemo
Set loMemo = Nothing
End Sub

Private Sub cmdJuridico_Click()
Dim loMemo  As UColRecMemos
Set loMemo = New UColRecMemos
    Set loMemo = frmColRecMemos.Inicio("J", fsJuridico, AXCodCta.NroCuenta, Me.lblNomPers)
        fsJuridico = loMemo.sMemo
Set loMemo = Nothing
End Sub

Private Sub cmdNuevo_Click()
Limpiar
fsNuevoEdita = "N"
EstadoBotones False, False, True
CmdExpRelacion.Enabled = False
AXCodCta.Enabled = True
AXCodCta.SetFocusAge
End Sub

Private Sub cmdPetitorio_Click()
Dim loMemo  As UColRecMemos
Set loMemo = New UColRecMemos
    Set loMemo = frmColRecMemos.Inicio("P", fsPetitorio, AXCodCta.NroCuenta, Me.lblNomPers)
        fsPetitorio = loMemo.sMemo
Set loMemo = Nothing
End Sub

Private Sub cmdProbatorios_Click()
Dim loMemo  As UColRecMemos
Set loMemo = New UColRecMemos
    Set loMemo = frmColRecMemos.Inicio("M", fsMProbatorios, AXCodCta.NroCuenta, Me.lblNomPers)
        fsMProbatorios = loMemo.sMemo
Set loMemo = Nothing
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Limpiar()
lblCodPers = ""
lblNomPers = ""
lblFecIngRecup = ""
lblMontoPrestamo = ""
txtNroExp = ""
lblSaldoCapital = ""
AxCodAbogado.Text = "": lblNomAbogado = ""
AXCodJuzgado.Text = "": lblNomJuzgado = ""
AxCodJuez.Text = "": lblNomJuez = ""
AxCodSecre.Text = "": lblNomSecre = ""
fsPetitorio = ""
fsHecho = ""
fsJuridico = ""
fsMProbatorios = ""
fsComplementos = ""

lblComision.Caption = ""
cbxViaP.ListIndex = -1
'cbDemanda.ListIndex = -1
cboTipMedCau.ListIndex = -1 'DAOR 20070417
chkComplementos.value = 0
chkFundHecho.value = 0
chkFundJuridico.value = 0
chkMedProb.value = 0
chkPetitorio.value = 0
AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
AXCodCta.SetFocusCuenta
txtMontoPet.Text = ""
'fsNuevoEdita = ""  ' Indicador si se realiza Nuevo / Edita
fbExisteExpediente = False
End Sub

Private Sub Form_Activate()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Me.chkPetitorio.value = IIf(Trim(fsPetitorio) = "", 0, 1)
    Me.chkFundHecho.value = IIf(Trim(fsHecho) = "", 0, 1)
    Me.chkFundJuridico.value = IIf(Trim(fsJuridico) = "", 0, 1)
    Me.chkMedProb.value = IIf(Trim(fsMProbatorios) = "", 0, 1)
    Me.chkComplementos.value = IIf(Trim(fsComplementos) = "", 0, 1)
    
End Sub

Private Sub Form_Load()
   Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
   'CargaForma Me.cbxViaP, "80", pConexionJud
   CargaViaProcesal
   CargaComboDatos cbxTipoCob, gColocRecTipoCobranza
   CargaComboDatos cbDemanda, gColocRecDemandado
   CargaComboDatos cboTipMedCau, gColocRecTipoMedCautelar
   'Carga procesos y deshabilita botones Relacion otros expediente
   If gsProyectoActual = "H" Then
      LblProceso.Visible = True
      CargaComboDatos CboProcesos, gColocRecTipoProceso
      CboProcesos.Visible = True
      CmdExpRelacion.Visible = True
      CmdExpRelacion.Enabled = False
   End If
   
   Set objPista = New COMManejador.Pista
   gsOpeCod = gRecRegistrarExpediente
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub txtMontoPet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If CDbl(txtMontoPet.Text) > 0 Then
        Me.cmdPetitorio.SetFocus
      Else
        MsgBox "El Monto debe ser Mayor a 0", vbInformation, "Aviso"
        Exit Sub
      End If
    End If
End Sub

Private Sub txtNroExp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtMontoPet.SetFocus
End Sub

Private Sub CargaPersonasRelacionadas()
Set frsPers = New ADODB.Recordset
    frsPers.Fields.Append "cPersCod", adVarChar, 13
    frsPers.Fields.Append "cPersNombre", adVarChar, 250 'RECO20150421
    frsPers.Fields.Append "cPersRelac", adVarChar, 5
    
    frsPers.Open
    If Len(Me.AxCodAbogado.Text) > 0 Then
        frsPers.AddNew
        frsPers.Fields("cPersCod") = Me.AxCodAbogado.Text
        frsPers.Fields("cPersNombre") = Trim(Me.lblNomAbogado.Caption)
        frsPers.Fields("cPersRelac") = gColRelPersEstudioJuridico
        frsPers.Update
    End If
    If Len(Me.AXCodJuzgado.Text) > 0 Then
        frsPers.AddNew
        frsPers.Fields("cPersCod") = Me.AXCodJuzgado.Text
        frsPers.Fields("cPersNombre") = Trim(Me.lblNomJuzgado.Caption)
        frsPers.Fields("cPersRelac") = gColRelPersJuzgado
        frsPers.Update
    End If
    If Len(Me.AxCodJuez.Text) > 0 Then
        frsPers.AddNew
        frsPers.Fields("cPersCod") = Me.AxCodJuez.Text
        frsPers.Fields("cPersNombre") = Trim(Me.lblNomJuez.Caption)
        frsPers.Fields("cPersRelac") = gColRelPersJuez
        frsPers.Update
    End If
    If Len(Me.AxCodSecre.Text) > 0 Then
        frsPers.AddNew
        frsPers.Fields("cPersCod") = Me.AxCodSecre.Text
        frsPers.Fields("cPersNombre") = Trim(Me.lblNomSecre.Caption)
        frsPers.Fields("cPersRelac") = gColRelPersSecretario
        frsPers.Update
    End If

End Sub

Private Sub CargaViaProcesal()
Dim oDatos As COMDConstSistema.DCOMGeneral
Dim rs As New ADODB.Recordset
cbxViaP.Clear
Set oDatos = New COMDConstSistema.DCOMGeneral
Set rs = oDatos.GetConstante(gColocRecViaProcesal)
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            cbxViaP.AddItem rs!cDescripcion + Space(70) + CStr(rs!nConsValor)
            rs.MoveNext
        Loop
    End If
Set rs = Nothing
Set oDatos = Nothing
End Sub

Private Function fValidaData() As Boolean
Dim lbOk As Boolean
lbOk = True

If Len(Me.AxCodAbogado.Text) > 0 Then
    If Me.AxCodAbogado.Text = Me.AXCodJuzgado Or _
       Me.AxCodAbogado.Text = Me.AxCodJuez Or _
       Me.AxCodAbogado.Text = Me.AxCodSecre Then
        MsgBox "Persona no puede ejercer varios roles al mismo tiempo", vbInformation, "Aviso"
        lbOk = False
        Exit Function
    End If
End If
If Len(Me.AXCodJuzgado.Text) > 0 Then
    If Me.AXCodJuzgado.Text = Me.AxCodJuez Or _
       Me.AXCodJuzgado.Text = Me.AxCodSecre Then
        MsgBox "Persona no puede ejercer varios roles al mismo tiempo", vbInformation, "Aviso"
        lbOk = False
        Exit Function
    End If
End If
If Len(Me.AxCodJuez.Text) > 0 Then
    If Me.AxCodJuez.Text = Me.AxCodSecre Then
        MsgBox "Persona no puede ejercer varios roles al mismo tiempo", vbInformation, "Aviso"
        lbOk = False
        Exit Function
    End If
End If
If frsPers.RecordCount = 0 Then
    MsgBox "Ingrese un Abogado, Juez, Juzgado", vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If
fValidaData = lbOk
End Function

Private Sub CargaComboDatos(ByVal combo As ComboBox, ByVal pnValor As Integer)
    Dim oConst As COMDConstantes.DCOMConstantes
    Dim rs As New ADODB.Recordset
    Set oConst = New COMDConstantes.DCOMConstantes
        Set rs = oConst.ObtenerVarRecuperaciones(pnValor)
        combo.Clear
        If Not (rs.EOF And rs.BOF) Then
            Do Until rs.EOF
                combo.AddItem rs(0)
                rs.MoveNext
            Loop
        End If
    Set oConst = Nothing
    Set rs = Nothing
End Sub

Public Sub EstadoBotones(ByVal bnuevo As Boolean, ByVal beditar As Boolean, ByVal bgrabar As Boolean)
    cmdNuevo.Enabled = bnuevo
    cmdEditar.Enabled = beditar
    cmdGrabar.Enabled = bgrabar
End Sub

Public Sub EstadoControles(ByVal bDatos As Boolean, ByVal bEstudio As Boolean, ByVal bInforme As Boolean)
    Me.fraDatos.Enabled = bDatos
    Me.fraEstudioJur.Enabled = bEstudio
    Me.fraInforme.Enabled = bInforme
End Sub
