VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenLista 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5190
   ClientLeft      =   855
   ClientTop       =   2445
   ClientWidth     =   10605
   Icon            =   "frmCajaGenLista.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
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
      Left            =   9120
      TabIndex        =   6
      Top             =   705
      Width           =   1335
   End
   Begin VB.Frame FraFechas 
      Height          =   600
      Left            =   6795
      TabIndex        =   15
      Top             =   45
      Width           =   3660
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   315
         Left            =   795
         TabIndex        =   16
         Top             =   188
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8388608
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txthasta 
         Height          =   315
         Left            =   2490
         TabIndex        =   17
         Top             =   195
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8388608
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   135
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1935
         TabIndex        =   18
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.CheckBox chkTodo 
      Caption         =   "Se&leccionar"
      Height          =   195
      Left            =   255
      TabIndex        =   0
      Top             =   30
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Height          =   4005
      Left            =   120
      TabIndex        =   7
      Top             =   1065
      Width           =   10350
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   8700
         TabIndex        =   11
         Top             =   3405
         Width           =   1470
      End
      Begin VB.TextBox txtMovDesc 
         Height          =   660
         Left            =   105
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3210
         Width           =   6270
      End
      Begin VB.CommandButton cmdConfHabCG 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   7230
         TabIndex        =   8
         Top             =   3405
         Width           =   1470
      End
      Begin Sicmact.FlexEdit fgLista 
         Height          =   2940
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   10080
         _extentx        =   17780
         _extenty        =   5186
         cols0           =   11
         highlight       =   1
         allowuserresizing=   3
         encabezadosnombres=   "N°-Fecha-Destino-Monto Hab.-Tipo Transporte-Empresa-Monto Serv-cMovDesc-cMovNroHab-cAreaCodDest-cAgeCodDest"
         encabezadosanchos=   "350-1200-3500-1200-2000-2000-1200-0-0-0-0"
         font            =   "frmCajaGenLista.frx":030A
         font            =   "frmCajaGenLista.frx":0332
         font            =   "frmCajaGenLista.frx":035A
         font            =   "frmCajaGenLista.frx":0382
         font            =   "frmCajaGenLista.frx":03AA
         fontfixed       =   "frmCajaGenLista.frx":03D2
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-R-L-L-R-L-C-C-C"
         formatosedit    =   "0-0-0-2-0-0-2-0-0-0-0"
         textarray0      =   "N°"
         lbformatocol    =   -1  'True
         lbpuntero       =   -1  'True
         lbordenacol     =   -1  'True
         colwidth0       =   345
         rowheight0      =   285
         forecolorfixed  =   -2147483630
      End
      Begin VB.CommandButton cmdExtornaHabCG 
         Caption         =   "&Extornar"
         Height          =   375
         Left            =   7230
         TabIndex        =   10
         Top             =   3405
         Width           =   1470
      End
   End
   Begin VB.Frame fraConfHab 
      Height          =   990
      Left            =   120
      TabIndex        =   1
      Top             =   45
      Width           =   6645
      Begin Sicmact.TxtBuscar TxtBuscarOrig 
         Height          =   345
         Left            =   840
         TabIndex        =   2
         Top             =   180
         Width           =   1065
         _extentx        =   1879
         _extenty        =   609
         appearance      =   1
         appearance      =   1
         font            =   "frmCajaGenLista.frx":03F8
         appearance      =   1
         stitulo         =   ""
         forecolor       =   16512
      End
      Begin Sicmact.TxtBuscar TxtBuscarDest 
         Height          =   345
         Left            =   840
         TabIndex        =   3
         Top             =   555
         Width           =   1050
         _extentx        =   1852
         _extenty        =   609
         appearance      =   1
         appearance      =   1
         font            =   "frmCajaGenLista.frx":041C
         appearance      =   1
         stitulo         =   ""
         forecolor       =   16512
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   75
         TabIndex        =   14
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Origen :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   75
         TabIndex        =   13
         Top             =   225
         Width           =   660
      End
      Begin VB.Label lblDestDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   1905
         TabIndex        =   5
         Top             =   570
         Width           =   4620
      End
      Begin VB.Label lblOrigDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   1905
         TabIndex        =   4
         Top             =   180
         Width           =   4620
      End
   End
End
Attribute VB_Name = "frmCajaGenLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnOpeBovedaCaja As OpeBovedaCajaGeneral
Dim lnOpeBovAge As OpeBovedaAgencia
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim oCaja As nCajaGeneral
Dim lbSalir As Boolean
Private Sub chkTodo_Click()
Me.fraConfHab.Enabled = chkTodo.value
If chkTodo.value = 1 Then
    If fraConfHab.Visible Then
        If TxtBuscarDest.Enabled Then
            TxtBuscarDest.SetFocus
        ElseIf TxtBuscarOrig.Enabled Then
            TxtBuscarOrig.SetFocus
       End If
    End If
Else
    If TxtBuscarDest.Enabled Then
        TxtBuscarDest = ""
        lblDestDesc = ""
    End If
    If TxtBuscarOrig.Enabled Then
        TxtBuscarOrig = ""
        lblOrigDesc = ""
    End If
   
End If
End Sub

Private Sub cmdConfHabCG_Click()
Dim lsMovNro As String
Dim lsMovNroHab As String
Dim lnImporteHab As Currency
Dim oCon As NContFunciones
Dim ldFechaHab  As Date
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset
Dim lsMensaje  As String
Dim lsObjOrig As String
Dim lsObjDest As String
Dim lnMovNroHab As Long
Set rsBill = New ADODB.Recordset
Set rsMon = New ADODB.Recordset

If fgLista.TextMatrix(1, 0) = "" Then Exit Sub
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Ingrese Descripción de Movimiento", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If
lsMovNroHab = fgLista.TextMatrix(fgLista.Row, 10)
lnMovNroHab = fgLista.TextMatrix(fgLista.Row, 11)
lnImporteHab = CCur(fgLista.TextMatrix(fgLista.Row, 3))
ldFechaHab = CDate(fgLista.TextMatrix(fgLista.Row, 1))
If lnImporteHab = 0 Then
    MsgBox "Monto de Habilitación no Válida", vbInformation, "Aviso"
    Exit Sub
End If
frmCajaGenEfectivo.Muestra lsMovNroHab
If frmCajaGenEfectivo.lbOk Then
    Set rsBill = frmCajaGenEfectivo.rsBilletes
    Set rsMon = frmCajaGenEfectivo.rsMonedas
Else
    Unload frmCajaGenEfectivo
    Set frmCajaGenEfectivo = Nothing
    Exit Sub
End If
Unload frmCajaGenEfectivo
Set frmCajaGenEfectivo = Nothing

Set oCon = New NContFunciones
Select Case gsOpeCod
    Case gOpeBoveAgeConfHabCGMN, gOpeBoveAgeConfHabCGME
        lsObjOrig = IIf(Mid(TxtBuscarOrig, 1, 3) = Mid(TxtBuscarDest, 1, 3), TxtBuscarOrig, TxtBuscarDest)
        lsObjDest = TxtBuscarDest
        lsMensaje = "Desea Confirmar la Habilitación de Caja General Y/O Agencias??"
    Case gOpeBoveCGConfHabAgeBoveMN, gOpeBoveCGConfHabAgeBoveME
        lsMensaje = "Desea Confirmar la Habilitación realizada por la Agencia??"
        lsObjOrig = TxtBuscarOrig
        lsObjDest = TxtBuscarDest
End Select
If MsgBox(lsMensaje, vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    

    If oCaja.GrabaConfHabEfectivo(lsMovNro, rsBill, rsMon, gsOpeCod, txtMovDesc, _
                            lsCtaDebe, lsCtaHaber, lnImporteHab, lsObjOrig, lsObjDest, _
                            lnMovNroHab) = 0 Then
'
'    If oCaja.GrabaConfHabEfectivoNew(lsMovNro, rsBill, rsMon, gsOpeCod, txtMovDesc, _
'                            lsCtaDebe, lsCtaHaber, lnImporteHab, lsObjOrig, lsObjDest, _
'                            lnMovNroHab) = 0 Then
        
        Select Case gsOpeCod
            Case gOpeBoveAgeConfHabCGMN, gOpeBoveAgeConfHabCGME
                Dim oContImp As NContImprimir
                Dim lsTexto As String
                Set oContImp = New NContImprimir
            
                lsTexto = oContImp.ImprimeDocConfHabEfectivo(gnColPage, gdFecSis, gsOpeCod, lsMovNro, gsNomCmac)
                
                EnviaPrevio lsTexto, Me.Caption, gnLinPage, False
                Set oContImp = Nothing
            Case gOpeBoveCGConfHabAgeBoveMN, gOpeBoveCGConfHabAgeBoveME
                ImprimeAsientoContable lsMovNro, , , , True, True, txtMovDesc
        End Select
        
        
        Set oCon = Nothing
        If MsgBox("Desea Realizar otra confirmación de Habilitación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            fgLista.EliminaFila fgLista.Row
            txtMovDesc = ""
        Else
            Unload Me
        End If
    End If
End If

End Sub

Private Sub cmdExtornaHabCG_Click()
Dim lsMovNro As String
Dim lsMovNroHab As String
Dim lnImporteHab As Currency
Dim oCon As NContFunciones
Dim ldFechaHab  As Date
Dim lnMovNroHab As Long
Dim lbEliminaMov As Boolean
Set oCon = New NContFunciones

If fgLista.TextMatrix(1, 0) = "" Then Exit Sub
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Ingrese Descripción de Movimiento", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If
gsMovNro = fgLista.TextMatrix(fgLista.Row, 10)
gnMovNro = fgLista.TextMatrix(fgLista.Row, 11)
lnImporteHab = CCur(fgLista.TextMatrix(fgLista.Row, 3))
ldFechaHab = CDate(fgLista.TextMatrix(fgLista.Row, 1))
lbEliminaMov = True
If lnImporteHab = 0 Then
    MsgBox "Monto de Habilitación no Válida", vbInformation, "Aviso"
    Exit Sub
End If
If MsgBox("Desea Extornar la habilitación seleccionada?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    If gdFecSis <> ldFechaHab Then
        If MsgBox("Se va a Realizar el Extorno de Movimientos de dias anteriores" & vbCrLf & " Desea Proseguir??", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        lbEliminaMov = False
    End If
    lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    If gbBitCentral Then
        oCaja.ExtornaHabEfectivo gdFecSis, ldFechaHab, lsMovNro, gnMovNro, gsOpeCod, _
                    Trim(txtMovDesc), lnImporteHab
    Else
        Dim oMov As New DMov
        oMov.ExtornaMovimiento lsMovNro, gnMovNro, gsOpeCod, txtMovDesc, lbEliminaMov, lsMovNro
        Set oMov = Nothing
    End If
    If gdFecSis <> ldFechaHab Then
        ImprimeAsientoContable lsMovNro, , , , True, True, txtMovDesc, , lnImporteHab
    End If
    Set oCon = Nothing
    If MsgBox("Desea Realizar otro Extorno de Habilitacion", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        txtMovDesc = ""
        fgLista.EliminaFila fgLista.Row
    Else
        Unload Me
    End If
End If


End Sub

Private Sub cmdProcesar_Click()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If fraConfHab.Visible Then
    If fraConfHab.Enabled And TxtBuscarOrig = "" Then
        MsgBox "Ingrese Area de Origen a Buscar", vbInformation, "Aviso"
        If TxtBuscarOrig.Enabled Then TxtBuscarOrig.SetFocus
        Exit Sub
    End If
    If fraConfHab.Enabled And TxtBuscarDest = "" Then
        MsgBox "Ingrese Area de Destino a Buscar", vbInformation, "Aviso"
        If TxtBuscarDest.Enabled Then TxtBuscarDest.SetFocus
        Exit Sub
    End If
End If
If FraFechas.Visible Then
    If ValFecha(Me.txtDesde) = False Then Exit Sub
    If ValFecha(txthasta) = False Then Exit Sub
    
    If CDate(txtDesde) > CDate(txthasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
End If
Me.MousePointer = 11
fgLista.Clear
fgLista.FormaCabecera
fgLista.Rows = 2
Select Case gsOpeCod
    Case gOpeBoveCGConfHabAgeBoveMN, gOpeBoveCGConfHabAgeBoveME
            Set rs = oCaja.GetHabCajaGen(lsCtaHaber, TxtBuscarOrig, CDate(txtDesde), CDate(txthasta))
    Case gOpeBoveCGExtHabAgeMN, gOpeBoveCGExtHabAgeME
            Set rs = oCaja.GetHabCajaGen(lsCtaHaber, TxtBuscarOrig, CDate(txtDesde), CDate(txthasta), Mid(TxtBuscarDest, 1, 3), Mid(TxtBuscarDest, 4, 2), False)
    Case gOpeBoveAgeExtHabAgeACGMN, gOpeBoveAgeExtHabAgeACGME, _
            gOpeBoveAgeExtHabEntreAgeMN, gOpeBoveAgeExtHabEntreAgeMe
            Set rs = oCaja.GetHabCajaGen(lsCtaHaber, TxtBuscarOrig, CDate(txtDesde), CDate(txthasta), Mid(TxtBuscarDest, 1, 3), Mid(TxtBuscarDest, 4, 2))
    Case gOpeBoveCGExtConfHabAgeBovMN, gOpeBoveCGExtConfHabAgeBovME
        If TxtBuscarOrig = TxtBuscarDest Then
            Me.MousePointer = 0
            MsgBox "Agencia de Origen no puede ser igual que la de Destino", vbInformation, "Aviso"
            If TxtBuscarOrig.Enabled Then TxtBuscarOrig.SetFocus
            Exit Sub
        End If
        Dim oOpe As New DOperacion
        Set rs = oCaja.GetDatosConfHabilitacion(oOpe.GetOperacionRefencia(gsOpeCod), TxtBuscarOrig, CDate(txtDesde), CDate(txthasta))
        Set oOpe = Nothing
    Case gOpeBoveAgeConfHabCGMN, gOpeBoveAgeConfHabCGME
        If TxtBuscarOrig = TxtBuscarDest Then
            Me.MousePointer = 0
            MsgBox "Agencia de Origen no puede ser igual que la de Destino", vbInformation, "Aviso"
            If Me.TxtBuscarOrig.Enabled Then TxtBuscarOrig.SetFocus
            Exit Sub
        End If
        Set rs = oCaja.GetHabCajaGen(lsCtaHaber, TxtBuscarOrig, CDate(txtDesde), CDate(txthasta), Mid(TxtBuscarDest, 1, 3), Mid(TxtBuscarDest, 4, 2))
End Select

If Not rs.EOF And Not rs.BOF Then
    Set fgLista.Recordset = rs
    fgLista.SetFocus
Else
    Me.MousePointer = 0
    MsgBox "Datos no Encontrados ", vbInformation, "Aviso"
End If
rs.Close
Set rs = Nothing
Me.MousePointer = 0
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgLista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub

Private Sub fgLista_RowColChange()
If fgLista.TextMatrix(1, 0) = "" Then
    txtMovDesc = ""
    Exit Sub
End If
txtMovDesc = fgLista.TextMatrix(fgLista.Row, 7)

End Sub

Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim oOpe As DOperacion
Set oOpe = New DOperacion
Set oCaja = New nCajaGeneral

CentraForm Me
lbSalir = False
txtDesde = gdFecSis
txthasta = gdFecSis
Me.Caption = gsOpeDesc
cmdExtornaHabCG.Visible = False
cmdConfHabCG.Visible = False


fraConfHab.Visible = False

Select Case gsOpeCod
    Case gOpeBoveCGExtHabAgeMN, gOpeBoveCGExtHabAgeME
        fraConfHab.Visible = True
        fraConfHab.Enabled = False
        lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", , , , , False)
        TxtBuscarDest.rs = GetObjetosOpeCta(gsOpeCod, "0", lsCtaDebe, "")
        TxtBuscarOrig.rs = GetObjetosOpeCta(gsOpeCod, "1", lsCtaDebe, "")
        TxtBuscarOrig.Enabled = False
        cmdExtornaHabCG.Visible = True
        If lsCtaHaber = "" Then
            MsgBox "Faltan Definir Cuentas Contables de Operación", vbExclamation, "Aviso"
            lbSalir = True
            Exit Sub
        End If
  
  Case gOpeBoveCGConfHabAgeBoveMN, gOpeBoveCGConfHabAgeBoveME, gOpeBoveCGExtConfHabAgeBovMN, gOpeBoveCGExtConfHabAgeBovME
        fgLista.EncabezadosNombres = "N°-Fecha-Origen-Monto Hab.-Tipo Transporte-Empresa-Monto Serv-cMovDesc-cMovNroHab-cAreaCodDest-cAgeCodDest"
        If gsOpeCod = gOpeBoveCGExtConfHabAgeBovMN Or gsOpeCod = gOpeBoveCGExtConfHabAgeBovME Then
            cmdExtornaHabCG.Visible = True
        Else
            cmdConfHabCG.Visible = True
        End If
        fraConfHab.Visible = True
        chkTodo.Visible = False
        lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", , , , , False)
        lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , , , , False)
        If lsCtaDebe = "" Or lsCtaHaber = "" Then
            MsgBox "Faltan Definir Cuentas Contables de Operación", vbExclamation, "Aviso"
            lbSalir = True
            Exit Sub
        End If
        TxtBuscarDest.rs = GetObjetosOpeCta(gsOpeCod, "0", lsCtaDebe, "")
        TxtBuscarOrig.rs = GetObjetosOpeCta(gsOpeCod, "1", lsCtaDebe, "")
    Case gOpeBoveAgeExtHabAgeACGMN, gOpeBoveAgeExtHabAgeACGME, gOpeBoveAgeExtHabEntreAgeMN, gOpeBoveAgeExtHabEntreAgeMe
        fraConfHab.Visible = True
        chkTodo.Visible = False
        lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
        TxtBuscarDest.rs = GetObjetosOpeCta(gsOpeCod, "0", lsCtaDebe, "")
        TxtBuscarOrig.rs = GetObjetosOpeCta(gsOpeCod, "1", lsCtaDebe, "", , gsCodAge)
        TxtBuscarOrig.Enabled = False
        cmdExtornaHabCG.Visible = True
        If lsCtaHaber = "" Then
            MsgBox "Faltan Definir Cuentas Contables de Operación", vbExclamation, "Aviso"
            lbSalir = True
            Exit Sub
        End If
        
End Select

Set oOpe = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set oCaja = Nothing
End Sub

Private Sub txtBuscarDest_EmiteDatos()
lblDestDesc = TxtBuscarDest.psDescripcion
If cmdProcesar.Visible Then cmdProcesar.SetFocus
End Sub

Private Sub TxtBuscarOrig_EmiteDatos()
lblOrigDesc = TxtBuscarOrig.psDescripcion
If Me.TxtBuscarDest.Enabled Then
    If TxtBuscarDest.Visible Then TxtBuscarDest.SetFocus
ElseIf Me.cmdProcesar.Enabled Then
    If Me.cmdProcesar.Visible Then Me.cmdProcesar.SetFocus
    End If
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtDesde) = False Then Exit Sub
    If CDate(txtDesde) > CDate(txthasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    txthasta.SetFocus
End If
End Sub


Private Sub txtHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txthasta) = False Then Exit Sub
    If CDate(txtDesde) > CDate(txthasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    cmdProcesar.SetFocus
End If
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdExtornaHabCG.Visible Then cmdExtornaHabCG.SetFocus
    If Me.cmdConfHabCG.Visible Then cmdConfHabCG.SetFocus
End If
End Sub
