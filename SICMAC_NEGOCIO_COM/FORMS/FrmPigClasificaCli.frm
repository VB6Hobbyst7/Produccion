VERSION 5.00
Begin VB.Form FrmPigClasificaCli 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clasificacion de Cliente - CMCPL"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   FillColor       =   &H8000000B&
   Icon            =   "FrmPigClasificaCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   5715
      TabIndex        =   20
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   390
      Left            =   7410
      Picture         =   "FrmPigClasificaCli.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Buscar ..."
      Top             =   180
      Width           =   435
   End
   Begin VB.Frame frmCambiar 
      Caption         =   "CMCPL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2505
      TabIndex        =   11
      Top             =   3315
      Width           =   5520
      Begin VB.TextBox txtClasiant 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1635
         TabIndex        =   16
         Top             =   150
         Width           =   1155
      End
      Begin VB.ComboBox cbClasInt 
         Height          =   315
         Left            =   3855
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   165
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "Clasificacion Interna"
         Height          =   180
         Left            =   2880
         TabIndex        =   13
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "Clasificacion Anterior "
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   255
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "SBS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   30
      TabIndex        =   9
      Top             =   3315
      Width           =   2460
      Begin VB.TextBox txtclasiSbs 
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
         Height          =   330
         Left            =   1065
         TabIndex        =   14
         Top             =   150
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Clasificacion"
         Height          =   195
         Left            =   75
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Creditos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2685
      Left            =   60
      TabIndex        =   7
      Top             =   645
      Width           =   7995
      Begin SICMACT.FlexEdit feContratos 
         Height          =   2340
         Left            =   60
         TabIndex        =   8
         Top             =   225
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   4128
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Credito-Piezas-Situacion-Dias Atraso-Mto.Colocado-Fecha Vcto."
         EncabezadosAnchos=   "380-1800-550-1800-880-1100-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         Enabled         =   0   'False
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   375
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6885
      TabIndex        =   6
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6900
      TabIndex        =   5
      Top             =   4455
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   7245
      Begin VB.TextBox TxtNombre 
         Height          =   300
         Left            =   1905
         TabIndex        =   3
         Top             =   195
         Width           =   5265
      End
      Begin VB.TextBox txtcodper 
         Height          =   300
         Left            =   735
         TabIndex        =   2
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   225
         Left            =   45
         TabIndex        =   1
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5715
      TabIndex        =   4
      Top             =   4455
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame frmEvaluar 
      Height          =   520
      Left            =   2520
      TabIndex        =   18
      Top             =   3788
      Width           =   5505
      Begin VB.CheckBox chkEvaluar 
         Caption         =   "Entra al Proceso de Evaluación Pignoraticia Mensual"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   120
         Width           =   4215
      End
   End
End
Attribute VB_Name = "FrmPigClasificaCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdBuscar_Click()
Dim loPers As UPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As dColPContrato
Dim loPersCredito As DPigContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As UProdPersona
Dim i As Integer
Dim liEvalCli As Integer
On Error GoTo ControlError

Set loPers = New UPersona
    Set loPers = frmBuscaPersona.Inicio
    If Not loPers Is Nothing Then
        lsPersCod = loPers.sPersCod
        lsPersNombre = loPers.sPersNombre
    End If
Set loPers = Nothing

txtcodper.Text = lsPersCod
txtNombre.Text = lsPersNombre

If Trim(lsPersCod) <> "" Then

    Set loPersCredito = New DPigContrato
        Set lrContratos = loPersCredito.dObtieneCreditosPignoPersona(lsPersCod)
    Set loPersCredito = Nothing

    i = 1
    If Not lrContratos.EOF And Not lrContratos.BOF Then
        CmdEditar.Enabled = True
        Do While Not lrContratos.EOF
           feContratos.AdicionaFila
           feContratos.TextMatrix(i, 0) = i
           feContratos.TextMatrix(i, 1) = lrContratos!cCtaCod ' codigo de la cuenta
           feContratos.TextMatrix(i, 2) = lrContratos!npiezas 'numero de piezas
           feContratos.TextMatrix(i, 3) = lrContratos!descri 'nEstadoCont 'estado de cont
           feContratos.TextMatrix(i, 4) = lrContratos!nDiasAtraso 'numero dias de atraso
           feContratos.TextMatrix(i, 5) = lrContratos!nMontoCol 'monto colocado
           feContratos.TextMatrix(i, 6) = lrContratos!dVenc 'fecha de vencimiento
           lrContratos.MoveNext
           i = i + 1
        Loop
    Else
        MsgBox "Cliente no posee contratos", vbInformation, "Aviso"
        Exit Sub
    End If
End If

If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New DPigContrato
        Set lrContratos = loPersCredito.dEvaluacionCli(lsPersCod)
        Set loPersCredito = Nothing
     If Not (lrContratos.EOF) Then
        liEvalCli = lrContratos!Califi
        txtclasiSbs.Text = lrContratos!SBS
        txtClasiant.Text = IIf(IsNull(lrContratos!EvalAnt), "", lrContratos!EvalAnt)
     End If
End If

If Trim(lsPersCod) <> "" Then
    cbClasInt.ListIndex = liEvalCli - 1
    cbClasInt.Tag = liEvalCli
End If
If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New DPigContrato
        Set lrContratos = loPersCredito.dEvaluacionCliManual(lsPersCod)
    Set loPersCredito = Nothing
    If Not lrContratos.EOF Then
        If lrContratos!iCalifica = "N" Then
           chkEvaluar.value = 0
        Else
           chkEvaluar.value = 1
        End If
    Else
        chkEvaluar.value = 1
    End If
End If
cbClasInt.Enabled = True
feContratos.Enabled = True
'cbClasInt.SetFocus
Exit Sub

ControlError:
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
     Limpiar
End Sub
Private Function Limpiar()
    txtcodper.Text = ""
    txtNombre.Text = ""
    txtclasiSbs.Text = ""
    txtClasiant.Text = ""
    cbClasInt.ListIndex = -1
    feContratos.Clear
    feContratos.Rows = 2
    feContratos.FormaCabecera
    cbClasInt.Enabled = False
    CmdEditar.Visible = True
    CmdEditar.Enabled = True
    cmdSalir.Visible = True
    cmdGrabar.Visible = False
    CmdCancelar.Visible = False
    frmEvaluar.Enabled = False
    frmCambiar.Enabled = False
    chkEvaluar.value = 0
    CmdBuscar.SetFocus
End Function
Private Sub CmdEditar_Click()
If Trim(txtcodper.Text) <> "" Then
     CmdEditar.Visible = False 'esconde el boton editar
     cmdSalir.Visible = False
     cmdGrabar.Visible = True
     CmdCancelar.Visible = True
     CmdBuscar.Enabled = True
     CmdBuscar.SetFocus
     frmCambiar.Enabled = True
     frmEvaluar.Enabled = True
End If
End Sub

Private Sub cmdGrabar_Click()
Dim loGrabar As DPigActualizaBD
Dim lsCodPer As String
Dim lsEvalCli As String
Dim lsEvalCliAnt As String

If MsgBox(" Grabar Nueva Evaluacion del Cliente ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'        cmEditar.Enabled = False
        lsCodPer = txtcodper.Text
        lsEvalCli = Mid(cbClasInt.Text, 1, 1)
        lsEvalCliAnt = cbClasInt.Tag
        If chkEvaluar.value = 1 Then
            liCalifica = "S"
        Else
            liCalifica = "N"
        End If
            
        Set loGrabar = New DPigActualizaBD
               Call loGrabar.dUpdateEval(lsCodPer, lsEvalCli, False)
               Call loGrabar.dInsertEvalPigManual(lsCodPer, lsEvalCli, lsEvalCliAnt, gsCodUser, gdFecSis, liCalifica, False)
         
         Set loGrabar = Nothing
         Limpiar
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If

    Limpiar
    cmdSalir.Visible = True

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdsalir_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
    If CmdBuscar.Visible = True Then
        CmdBuscar.SetFocus
    End If
End Sub

Private Sub Form_Load()
 Dim loPersContrato As dColPContrato
 Dim loPersCredito As DPigContrato
 Dim lrContratos As ADODB.Recordset
    
    frmEvaluar.Enabled = False
    frmCambiar.Enabled = False
    txtcodper.Enabled = False
    txtNombre.Enabled = False
    cbClasInt.Enabled = False
    CmdBuscar.Enabled = True
    
     
    Set loPersCredito = New DPigContrato
       Set lrContratos = loPersCredito.dEvaluacion(gColocPigCalifCte)
    Set loPersCredito = Nothing
    
    Do While Not lrContratos.EOF
           cbClasInt.AddItem lrContratos!Nom 'cConsDescripcion
            lrContratos.MoveNext
    Loop
     
    Set lrContratos = Nothing
    'Limpiar
    Me.Icon = LoadPicture(App.path & "\bmps\cm.ico")
    
End Sub
