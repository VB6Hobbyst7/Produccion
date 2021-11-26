VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColRecIngresoOtrasEnt 
   Caption         =   "Recuperaciones - Ingreso de creditos de otras entidades"
   ClientHeight    =   5745
   ClientLeft      =   1155
   ClientTop       =   1290
   ClientWidth     =   8955
   Icon            =   "frmColRecIngresoOtrasEnt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   1980
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   8640
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
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
         Left            =   7425
         TabIndex        =   2
         Top             =   1365
         Width           =   945
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7425
         TabIndex        =   1
         Top             =   930
         Width           =   960
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7410
         TabIndex        =   0
         Top             =   510
         Width           =   960
      End
      Begin MSComctlLib.ListView lstMiembros 
         Height          =   1050
         Left            =   195
         TabIndex        =   32
         Top             =   765
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   1852
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Relacion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "CodRelacion"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.ComboBox cbxRelacion 
         Height          =   315
         Left            =   5820
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   435
         Width           =   1470
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1770
         TabIndex        =   28
         Tag             =   "TXTNOMBRE"
         Top             =   450
         Width           =   4050
      End
      Begin VB.TextBox txtCodPers 
         Height          =   285
         Left            =   195
         TabIndex        =   13
         Tag             =   "TXTCODIGO"
         Top             =   435
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Relación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5850
         TabIndex        =   31
         Top             =   165
         Width           =   840
      End
      Begin VB.Label Label14 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   210
         Width           =   840
      End
      Begin VB.Label Label13 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   195
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   5940
      TabIndex        =   10
      Top             =   5310
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
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
      Left            =   7380
      TabIndex        =   11
      Top             =   5295
      Width           =   1335
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "E&liminar"
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
      Left            =   4440
      TabIndex        =   9
      Top             =   5295
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   5310
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   2430
      Left            =   3795
      TabIndex        =   15
      Top             =   2760
      Width           =   4935
      Begin VB.TextBox txtLineaCredito 
         Height          =   285
         Left            =   1530
         TabIndex        =   38
         Top             =   1440
         Width           =   1185
      End
      Begin VB.ComboBox cbxCondicion 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   2100
      End
      Begin VB.TextBox txtInteres 
         Height          =   315
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin MSMask.MaskEdBox mskFecUlPag 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   645
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   330
         Left            =   1560
         TabIndex        =   3
         Top             =   255
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNomAnalista 
         Height          =   285
         Left            =   1545
         TabIndex        =   6
         Top             =   1050
         Width           =   1170
      End
      Begin VB.Label lblCodAnalista 
         Height          =   225
         Left            =   3390
         TabIndex        =   26
         Top             =   1095
         Width           =   1260
      End
      Begin VB.Label Label10 
         Caption         =   "Analista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   25
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Condición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   24
         Top             =   1965
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Linea Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   23
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "T.Int Comp."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2985
         TabIndex        =   22
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label6 
         Caption         =   "Fec.UltPago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   701
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Fec.Ing. Recup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2430
      Left            =   120
      TabIndex        =   14
      Top             =   2745
      Width           =   3615
      Begin SICMACT.EditMoney txtSaldoCap 
         Height          =   285
         Left            =   1710
         TabIndex        =   34
         Top             =   360
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtSaldoIntCom 
         Height          =   285
         Left            =   1710
         TabIndex        =   35
         Top             =   810
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtSaldoIntMorat 
         Height          =   285
         Left            =   1710
         TabIndex        =   36
         Top             =   1260
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtSaldoGasto 
         Height          =   285
         Left            =   1710
         TabIndex        =   37
         Top             =   1710
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Gastos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   19
         Top             =   1860
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Saldo Int. Mora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   75
         TabIndex        =   18
         Top             =   1360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo Int. Comp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   17
         Top             =   875
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo Capital"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   16
         Top             =   375
         Width           =   1335
      End
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   375
      Left            =   180
      TabIndex        =   33
      Top             =   180
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmColRecIngresoOtrasEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REGISTRO DE CREDITO EN RECUPERACIONES
'Archivo:  frmColRecIngresoOtrasEnt.frm
'LAYG   :  01/08/2001.
'Resumen:  Nos permite registrar un credito en recuperaciones

Option Explicit
Dim rsAux As ADODB.Recordset
Function FueIngresado(Cuenta As String) As Boolean
'Dim sql As String
'Dim RS As New ADODB.Recordset
'sql = " SELECT cCodCta FROM CREDCJUDI WHERE CCODCTA = '" & Cuenta & "'"
'RS.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'FueIngresado = IIf(RSVacio(RS), False, True)
'RClose RS
End Function
Private Sub Limpiar()
    
    AXCodCta.Cuenta = ""
    lblCodAnalista.Caption = ""
    txtNomAnalista = ""
    txtInteres.Text = ""
    txtNombre.Text = ""
    'txtSaldCap.Value = 0
    'txtSaldGast.Value = 0
    'txtSaldIntComp.Value = 0
    'txtSaldIntMor.Value = 0
    mskFecha.Mask = ""
    mskFecha.Text = ""
    mskFecha.Mask = "##/##/####"
    mskFecUlPag.Mask = ""
    mskFecUlPag.Text = ""
    mskFecUlPag.Mask = "##/##/####"
    'cbxLineaCred.ListIndex = -1
    cbxCondicion.ListIndex = -1
    'cbxEstado.ListIndex = -1
    cbxRelacion.ListIndex = -1
    lstMiembros.ListItems.Clear
    
    'txtMetLiq.Text = ""
    
End Sub







Private Sub cbxRelacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.cmdAgregar.SetFocus
End If
End Sub



Private Sub CmdAgregar_Click()
Dim itm As ListItem
Dim i As Byte
If Me.txtCodPers.Text = "" Then
 MsgBox "Ingrese un Cliente", vbInformation, "Aviso"
 Exit Sub
End If
For i = 1 To lstMiembros.ListItems.Count Step 1
    If lstMiembros.ListItems(i).Text = txtCodPers.Text Then
    MsgBox "Cliente ya Fue Ingresado", vbInformation, "Aviso"
    Exit Sub
    End If
Next
Set itm = lstMiembros.ListItems.Add(, , Me.txtCodPers.Text)
itm.SubItems(1) = txtNombre.Text
itm.SubItems(2) = cbxRelacion.Text
txtCodPers.Text = ""
txtNombre.Text = ""
cbxRelacion.ListIndex = -1
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As UPersona
Dim lsPersCod As String
Dim liFil As Integer
Dim ls As String
On Error GoTo ControlError

Set loPers = New UPersona
Set loPers = frmBuscaPersona.Inicio

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod
    '** Verifica que no este en lista
    For liFil = 1 To Me.lstMiembros.ListItems.Count
        If lsPersCod = Me.lstMiembros.ListItems.Item(liFil).Text Then
           MsgBox " Cliente Duplicado ", vbInformation, "Aviso"
           Exit Sub
        End If
    Next liFil
    Me.txtCodPers.Text = loPers.sPersCod
    Me.txtNombre.Text = loPers.sPersNombre
End If

Set loPers = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

Me.cbxRelacion.SetFocus
End Sub
Private Sub cmdCancelar_Click()
Limpiar
End Sub

Private Sub Valida()
Dim lnFila As Integer

Set rsAux = New ADODB.Recordset
    rsAux.Fields.Append "cPersCod", adVarChar, 13
    rsAux.Fields.Append "cPersNombre", adVarChar, 50
    rsAux.Fields.Append "cPersRelac", adVarChar, 5
    
    rsAux.Open
    For lnFila = 1 To Me.lstMiembros.ListItems.Count
        rsAux.AddNew
        rsAux.Fields("cPersCod") = lstMiembros.ListItems.Item(lnFila).Text
        rsAux.Fields("cPersNombre") = lstMiembros.ListItems.Item(lnFila).ListSubItems.Item(1)
        rsAux.Fields("cPersRelac") = lstMiembros.ListItems.Item(lnFila).ListSubItems.Item(9)
        
        rsAux.Update
    Next
    rsAux.MoveFirst
    'Set fgGetCodigoPersonaListaRsNew (   = rsAux
    
End Sub

Private Sub cmdGrabar_Click()

Dim loContFunct As NContFunciones
Dim loGrabar As NColRecCredito

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

'On Error GoTo ControlError

If MsgBox(" Grabar Registo de Credito en Recuperaciones ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        
    'Genera el Mov Nro
    Set loContFunct = New NContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        
    Set loGrabar = New NColRecCredito
        Call loGrabar.nRegistraCredEnRecup(AXCodCta.NroCuenta, Format(Me.txtInteres.Text, "#0.00"), CCur(Format(Me.txtSaldoCap.Text, "#0.0")), _
             0, Trim(Me.txtLineaCredito.Text), rsAux, Format(Me.mskFecha.Text, "mm/dd/yyyy"), CCur(Format(Me.txtSaldoIntCom.Text, "#0.00")), _
             CCur(Format(Me.txtSaldoIntMorat.Text, "#0.00")), CCur(Format(Me.txtSaldoGasto.Text, "#0.00")), 0, lsFechaHoraGrab, _
             lsMovNro, False)
    Set loGrabar = Nothing
    Limpiar
    Me.AXCodCta.SetFocus
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "


'Dim SQLCredCJudi As String
'Dim SQLCJCredito As String
'Dim SQLRelCtaMigJudi As String
'Dim SQLKardexJudi As String
'Dim SQLPersCJudi As String
'Dim lsCuenta As String
'Dim lsHoy As String
'Dim lsCad As String
'lsCuenta = Me.acxCuenta.Text
'If FueIngresado(lsCuenta) = True Then
'  MsgBox "El Crédito ya se encuentra en Cobranza Judicial", vbInformation, "Aviso"
'  Exit Sub
'End If
'lsHoy = FechaHora(gdFecSis)
'lsCad = ValidaFecha(Me.mskFecha.Text)
'If lsCad <> "" Then
'    MsgBox lsCad, vbInformation, "Aviso"
'    mskFecha.SetFocus
'    Exit Sub
'End If
'lsCad = ValidaFecha(Me.mskFecUlPag)
'If lsCad <> "" Then
'    MsgBox lsCad, vbInformation, "Aviso"
'    mskFecha.SetFocus
'    Exit Sub
'End If
'
'On Error GoTo ERROR
'If MsgBox("Desea Guardar los Cambios?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
' dbCmact.BeginTrans
'   SQLCredCJudi = "INSERT CREDCJUDI(cCodCta,dFecJud,nSaldCap,nSaldIntCom,nSaldIntMor," _
'        & " nSaldGast,nTasaInt,cMetLiq,nCapPag,nIntComPag,nIntMorPag,nGastoPag,cEstado," _
'        & " cCondicion,cCodUsu,dFecMod,cTipCJ,nCodComi,dFecCast,nTasaIntMor,cPagare," _
'        & " dFecUltPago,nIntComGen,nNumTranCta,cCodLinCred) VALUES('" & lsCuenta & "'," _
'        & " '" & Format(mskFecha.Text, "mm/dd/yyyy") & "'," & Format(Me.txtSaldCap.Value, "#0.00") & "," & Format(Me.txtSaldIntComp.Value, "#0.00") & "," _
'        & " " & Format(txtSaldIntMor.Value, "#0.00") & " , " & Format(txtSaldGast.Value, "#0.00") & "," & Format(txtInteres.Text, "#0.00") & "," _
'        & " '" & Me.txtMetLiq.Text & "',0,0,0,0,'" & Mid(cbxEstado.Text, 1, 1) & "','" & Mid(cbxCondicion.Text, 1, 1) & "'," _
'        & " '" & gsCodUser & "', '" & lsHoy & "',null,null,null,0,'" & txtPagare.Text & "','" & Format(mskFecUlPag.Text, "mm/dd/yyyy") & "'," _
'        & " null,0,'" & Mid(cbxLineaCred.Text, 1, 8) & "')"
'
'        dbCmact.Execute SQLCredCJudi
'       ' dbCmact.RollbackTrans
'   SQLCJCredito = " INSERT CJCREDITO (cCodCta,cEstado,dAsignacion,nMontoSol,nCuotasSol," _
'        & " nPlazoSol,cNumFuente,cCondCre,cDestCre,cCodLinCred,nMontoSug,nCuotasSug,nPlazoSug," _
'        & " nGraciaSug,nCuotaSug,dFecApr,nTasaInt,nMontoApr,nIntApr,nCuotasApr,nPlazoApr," _
'        & " nGraciaApr,nCuotaApr,ctipCuota,cperiodo,cCodAnalista,cApoderado,cCauRech,dFecVig," _
'        & " cTipoDesemb,nMontoDesemb,nNroProxDesemb,dFecUltDesemb,nSaldoCap,nCapPag,nIntComPag," _
'        & " nIntMorPag,nGastoPag,nDiasAtraso,nIntMorCal,nNroProxCuota,dFecUltPago,cFlagProtesto," _
'        & " cLibAmort,dCancelado,dJudicial,cNota1,cNota2,nDiAtrAcu,cCalifica,cCodInst,cCodModular," _
'        & " nNroTransac,cCodUsu,dFecModif,cComenta,nDiaFijo,nNroRepro,cDescarAuto,cMetLiquid," _
'        & " cRefinan ) VALUES('" & lsCuenta & "','H','" & lsHoy & "'," & Format(Me.txtSaldCap.Value, "#0.00") & ",0," _
'        & " 0,'','','','" & Mid(cbxLineaCred.Text, 1, 8) & "',0,0,0," _
'        & " 0,0,'" & lsHoy & "'," & Me.txtInteres.Text & ",0,0,0,0," _
'        & " 0,0,null,null,'" & Me.lblCodAnalista.Caption & "',null,null,'" & lsHoy & "'," _
'        & " null,0,0,'" & lsHoy & "'," & Format(Me.txtSaldCap.Text, "#0.00") & ",0,0," _
'        & " 0,0,0,0,0,null,null," _
'        & " null,null,'" & lsHoy & "',null,null,0,null,null,null," _
'        & " 1,'" & gsCodUser & "','" & lsHoy & "',null,0,0,null,'" & Me.txtMetLiq.Text & "',NULL)"
'
'        dbCmact.Execute SQLCJCredito
'
'        SQLRelCtaMigJudi = "INSERT RELCTAMIGJUDI(cCodCta,cCodAnt)" _
'                 & " VALUES('" & lsCuenta & "','" & txtPagare.Text & "')"
'        dbCmact.Execute SQLRelCtaMigJudi
'
'        SQLKardexJudi = " INSERT KARDEXJUDI(cCodCta,dFecTran,nNumTranCta,nMonTran,nCapital," _
'        & " nIntComp,nIntMorat,nGastos,cCodOpe,cNumDoc,cCodAge,cCodUsu,cEstado,cFlag," _
'        & " cMetLiquid,cTipTrans) VALUES('" & lsCuenta & "','" & Format(mskFecha.Text, "mm/dd/yyyy") & "'," _
'        & " " & 1 & ",0," & Format(txtSaldCap.Value, "#0.00") & "," & Format(txtSaldIntComp.Value, "#0.00") & "," & Format(txtSaldIntMor.Value, "#0.00") & "," _
'        & " " & Format(txtSaldGast.Value, "#0.00") & ",'020100',null,'" & gsCodAge & "','" & gsCodUser & "',null,null," _
'        & " '" & Me.txtMetLiq.Text & "','0')"
'
'        dbCmact.Execute SQLKardexJudi
'
'        For i = 1 To lstMiembros.ListItems.Count
'            SQLPersCJudi = "INSERT PERSCJUDI(cCodPers,cCodCta,cRelaCta)" _
'            & " VALUES('" & lstMiembros.ListItems(i).Text & "','" & lsCuenta & "'," _
'            & " '" & lstMiembros.ListItems.Item(i).ListSubItems(2) & "')"
'            dbCmact.Execute SQLPersCJudi
'        Next
'   dbCmact.CommitTrans
'   Limpiar
'   Me.acxCuenta.SetFocus
'End If
'Exit Sub
'ERROR:
'   MsgBox Err.Description
'   dbCmact.RollbackTrans
End Sub

Private Sub CmdQuitar_Click()
If lstMiembros.ListItems.Count <= 0 Then Exit Sub
lstMiembros.ListItems.Remove (lstMiembros.SelectedItem.Index)
'lstMiembros.SelectedItem.ListSubItems.Clear
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub
Private Sub CargaLineas()
'Dim sql As String
'Dim RS As New ADODB.Recordset
' sql = " select cCodLinCred,cDesLinCred from " & gcCentralCom & "lineacredito"
' RS.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
' If RSVacio(RS) Then
'
' Else
'  Do While Not RS.EOF
'   Me.cbxLineaCred.AddItem (RS!cCodLinCred & "   " & RS!cDesLinCred)
'   RS.MoveNext
'  Loop
' End If
'RClose RS
'cbxLineaCred.ListIndex = -1
End Sub
Private Sub Form_Load()
'CargaLineas
'CargaForma Me.cbxEstado, "83", dbCmact
'CargaForma Me.cbxCondicion, "84", dbCmact
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
cbxRelacion.AddItem ("TI")
cbxRelacion.AddItem ("RE")
cbxRelacion.AddItem ("GA")
cbxRelacion.AddItem ("CO")
cbxRelacion.AddItem ("CY")
'lstMiembros.ColumnHeaders.Add , , "Código", 1550
'lstMiembros.ColumnHeaders.Add , , "Nombre", 4050
'lstMiembros.ColumnHeaders.Add , , "Relacíon", 1300
'lstMiembros.View = lvwReport
'lstMiembros.FullRowSelect = True
'lstMiembros.HideSelection = True
Me.txtCodPers.Locked = True
Me.txtNombre.Locked = True
End Sub



Private Sub mskFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.txtInteres.SetFocus
End If
End Sub

Private Sub mskFecUlPag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'Me.txtMetLiq.SetFocus
End If
End Sub



Private Sub txtinteres_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.mskFecUlPag.SetFocus
End If
End Sub

Private Sub txtinteres_LostFocus()
txtInteres.Text = Format(txtInteres.Text, "#0.00")
End Sub

Private Sub txtNomAnalista_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 119 Then
'      frmBusIdentJudicial.Pant = "EXPEDIENTES"
'      frmBusIdentJudicial.ROL = "ANA"
'      frmBusIdentJudicial.Abo = 1
'      frmBusIdentJudicial.Caption = "Lista de Analistas"
'      frmBusIdentJudicial.Show 1
'   End If
End Sub
Private Sub txtNomAnalista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'Me.cbxEstado.SetFocus
End If
End Sub



Private Sub LlenaPantalla()
'Dim sql As String
'Dim RS As New ADODB.Recordset
'Dim i As Byte
'Dim itm As ListItem
'cmdGrabar.Enabled = True
'sql = "SELECT * FROM CREDCJUDI WHERE cCodCta = '" & Me.acxCuenta.Text & "' "
'RS.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'If RSVacio(RS) Then
'Else
'  MsgBox "Crédito Ya se Encuentra en cobranza Judicial", vbInformation, "Aviso"
'
'  Me.cmdGrabar.Enabled = True
'  For i = 1 To cbxLineaCred.ListCount
'    If Trim(RS!cCodLinCred) = Mid(cbxLineaCred.List(i), 1, 8) Then
'       cbxLineaCred.ListIndex = i
'       Exit For
'    End If
'  Next
'  Me.txtInteres.Text = RS!nTasaInt
'  Me.txtMetLiq.Text = RS!cMetLiq
'  Me.txtPagare.Text = RS!cPagare
'  Me.txtSaldCap.Value = RS!nSaldCap
'  Me.txtSaldGast.Value = RS!nsaldgast
'  Me.txtSaldIntComp.Value = RS!nSaldIntcom
'  Me.txtSaldIntMor.Value = RS!nSaldIntMor
'  Me.mskFecha.Text = Format(RS!dFecJud, "dd/mm/yyyy")
'  Me.mskFecUlPag.Text = Format(RS!dFecUltPago, "dd/mm/yyyy")
'  For i = 0 To cbxCondicion.ListCount
'      If Trim(RS!cCondicion) = Mid(cbxCondicion.List(i), 1, 1) Then
'         Me.cbxCondicion.ListIndex = i
'         Exit For
'      End If
'  Next
'  For i = 0 To cbxEstado.ListCount
'      If Trim(RS!cEstado) = Mid(cbxEstado.List(i), 1, 1) Then
'              Me.cbxEstado.ListIndex = i
'      Exit For
'      End If
'  Next
'End If
'RClose RS
'sql = "SELECT CJ.cCodAnalista,U.cNomUsu FROM CJCredito CJ INNER JOIN " & gcCentralCom & "USUARIO U " _
'    & " ON CJ.cCodAnalista = U.cCodUsu WHERE cCodCta = '" & Me.acxCuenta.Text & "'"
'RS.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'If RSVacio(RS) Then
'   Me.txtNomAnalista.Text = ""
'   Me.lblCodAnalista.Caption = ""
'Else
'   Me.txtNomAnalista.Text = RS!cNomUsu
'   Me.lblCodAnalista.Caption = RS!cCodAnalista
'End If
'RClose RS
'sql = " SELECT PJ.cCodPers,PJ.cRelaCta,P.cNomPers FROM PERSCJUDI PJ" _
'    & " INNER JOIN " & gcCentralPers & "PERSONA P ON PJ.cCodPers = P.cCodPers" _
'    & " WHERE PJ.cCodCta = '" & Me.acxCuenta.Text & "'"
'RS.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'If RSVacio(RS) Then
'Else
' lstMiembros.ListItems.Clear
' Do While Not RS.EOF
'        Set itm = lstMiembros.ListItems.Add(, , RS!cCodPers)
'         itm.SubItems(1) = RS!cNomPers
'         itm.SubItems(2) = RS!cRelaCta
'         RS.MoveNext
'    Loop
'cmdGrabar.Enabled = False
'End If
'RClose RS
End Sub

