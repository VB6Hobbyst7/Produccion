VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSegTarjetaDepositoPorActivacion 
   Caption         =   "Depósito por Activacón Seguro"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   Icon            =   "frmSegTarjetaDepositoPorActivacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   14
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   13
      Top             =   6240
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   4410
      TabCaption(0)   =   "Datos del depósito"
      TabPicture(0)   =   "frmSegTarjetaDepositoPorActivacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Cuenta Institución Financiera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   9855
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   8520
            TabIndex        =   9
            Top             =   720
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   300
            Left            =   8520
            TabIndex        =   10
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Sicmact.TxtBuscar txtObjOrig 
            Height          =   330
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   582
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
         End
         Begin VB.Label lblDescCtabanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   6255
         End
         Begin VB.Label lblDescbanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   290
            Left            =   2400
            TabIndex        =   18
            Top             =   370
            Width           =   4095
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Deposito:"
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
            Left            =   7080
            TabIndex        =   12
            Top             =   375
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Monto:   S/. "
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
            Left            =   7440
            TabIndex        =   11
            Top             =   750
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operaciones con Depósitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3735
         Left            =   240
         TabIndex        =   1
         Top             =   2040
         Width           =   9855
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   16
            Top             =   3000
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2400
            TabIndex        =   15
            Top             =   3000
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   5
            Top             =   3000
            Width           =   975
         End
         Begin VB.CommandButton cmdEditar 
            Caption         =   "Editar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4560
            TabIndex        =   4
            Top             =   3000
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1320
            TabIndex        =   3
            Top             =   3000
            Width           =   975
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   3120
            Width           =   1215
         End
         Begin Sicmact.FlexEdit feSolicitudes 
            Height          =   2415
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   4260
            Cols0           =   12
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Nº Solicitud-Valor-Titular-Cuenta-Monto-Glosa-SolicitudDos-AreaCod-AgeCod-cObjetoCod-CuentaBackup"
            EncabezadosAnchos=   "300-1200-0-3000-1700-900-2200-0-0-0-0-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-X-X-5-6-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-R-R-L-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-2-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label3 
            Caption         =   "Total:"
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
            Left            =   7800
            TabIndex        =   7
            Top             =   3120
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmSegTarjetaDepositoPorActivacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oOpe As New DOperacion
Dim oCtaIf As New NCajaCtaIF
Dim lsCtaContBanco As String
Dim lbDeposito As Boolean
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Private Type Operaciones
    sNSolicitud As String
    sCtaAhorro As String
    nMonto As Currency
    sGlosa As String
End Type
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub feSolicitudes_OnCellChange(pnRow As Long, pnCol As Long)
    Dim fila As Integer, col As Integer
    fila = pnRow
    col = pnCol
    If col = 5 Then
        If feSolicitudes.TextMatrix(fila, 1) <> "" Then
            Dim i As Integer
            Dim Total As Currency
            Total = 0#
            For i = 1 To feSolicitudes.Rows - 1
                If feSolicitudes.TextMatrix(i, 1) <> "" Then
                    Total = Total + CCur(feSolicitudes.TextMatrix(i, 5))
                End If
            Next
            txtTotal.Text = Format(Total, "#,##0.00")
            feSolicitudes.TextMatrix(fila, 5) = Format(feSolicitudes.TextMatrix(fila, 5), "#,##0.00")
            feSolicitudes.row = fila
            feSolicitudes.col = 5
        End If
    End If
    If col = 6 Then
        If feSolicitudes.TextMatrix(fila, 1) <> "" Then
            cmdAgregar.Enabled = True
            cmdAgregar.SetFocus
        End If
    End If
End Sub
Private Sub feSolicitudes_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim lcNumSol As String, lcPersNombre As String
    Dim ldFechaSol As Date
    Dim fila As Integer, i As Integer
    Dim lnMonto As Currency, nTotal As Currency
    Dim oNSeg As New NSeguros
    Dim rsSeg As New ADODB.Recordset
    Dim clsCuentas As New UCapCuentas
    Dim lcAreaCod As String, lcAgeCod As String, lcObjetoCod As String
    
    If feSolicitudes.TextMatrix(pnRow, 4) <> "" Then
        MsgBox "Solo puede modificar el Monto y la Glosa. Debera quitar y agregar otra solicitud", vbInformation, "Aviso"
        feSolicitudes.TextMatrix(pnRow, 1) = feSolicitudes.TextMatrix(pnRow, 7)
        Exit Sub
    End If
    feSolicitudes.TextMatrix(pnRow, pnCol) = ""
    nTotal = 0#
    lcNumSol = ""
    For i = 1 To feSolicitudes.Rows - 1
        If feSolicitudes.TextMatrix(1, 1) = "" Then
            Set rsSeg = oNSeg.ObtenerSegTarjetaSolicitudesPendientes(1)
            Exit For
        Else
            If feSolicitudes.TextMatrix(i, 1) <> "" Then
                lcNumSol = lcNumSol & feSolicitudes.TextMatrix(i, 1) & ","
            End If
        End If
    Next
    If lcNumSol <> "" Then
        lcNumSol = Mid(lcNumSol, 1, Len(lcNumSol) - 1)
        Set rsSeg = oNSeg.ObtenerSegTarjetaSolicitudesPendientes(2, lcNumSol)
    End If
    lcNumSol = ""
    fila = feSolicitudes.row
    If Not (rsSeg.BOF And rsSeg.EOF) Then
        Do While Not rsSeg.EOF
            lcNumSol = rsSeg!cNumSolicitud
            lcPersNombre = rsSeg!cPersNombre
            ldFechaSol = rsSeg!dFechaSolicitud
            lnMonto = rsSeg!nMontoSolicitud
            lcAreaCod = rsSeg!AreaCod
            lcAgeCod = rsSeg!AgeCod
            lcObjetoCod = rsSeg!cObjetoCod
            frmSegTarjetaSolicitudesPendientes.lstSolicitudes.AddItem lcNumSol & Space(2) & lcPersNombre & Space(2) & Format(CStr(lnMonto), "#,##0.00")
            rsSeg.MoveNext
        Loop
        Set clsCuentas = frmSegTarjetaSolicitudesPendientes.Inicio
        If clsCuentas Is Nothing Then
        Else
            If clsCuentas.sCuenta <> "" Then 'Si tiene cuenta de ahorro o no encuentra ninguna solicitud entra
                feSolicitudes.TextMatrix(fila, 1) = clsCuentas.sNumSolictud
                feSolicitudes.TextMatrix(fila, 3) = clsCuentas.sPersNombre
                feSolicitudes.TextMatrix(fila, 4) = clsCuentas.sCuenta
                feSolicitudes.TextMatrix(fila, 5) = Format(clsCuentas.sMonto, "#,##0.00")
                feSolicitudes.TextMatrix(fila, 7) = clsCuentas.sNumSolictud
                feSolicitudes.TextMatrix(fila, 8) = lcAreaCod
                feSolicitudes.TextMatrix(fila, 9) = lcAgeCod
                feSolicitudes.TextMatrix(fila, 10) = lcObjetoCod
                feSolicitudes.TextMatrix(fila, 11) = clsCuentas.sCuenta 'Cuenta Backup
                txtTotal.Text = Format(CCur(txtTotal.Text) + CCur(feSolicitudes.TextMatrix(fila, 5)), "#,##0.00") 'Total
                feSolicitudes.row = fila
                feSolicitudes.col = 4
            Else
                'Entra aca si la persona no tiene ninguna cuenta de ahorro o la solicitud no existe
                MsgBox "Esta persona no tienen ninguna cuenta de ahorro.", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        Set clsCuentas = Nothing
    Else
        MsgBox "No hay ninguna solicitud pendiente", vbInformation, "Aviso"
        cmdAgregar.Enabled = True
    End If
    Set oNSeg = Nothing
    Set rsSeg = Nothing
End Sub
Private Sub feSolicitudes_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'    Dim valor As String
'    valor = feSolicitudes.TextMatrix(pnRow, pnCol)
    Dim scolumnas() As String
    scolumnas = Split(feSolicitudes.ColumnasAEditar, "-")
    If scolumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    txtObjOrig.rs = oOpe.GetRsOpeObj(gsOpeCod, "0")
    txtFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
    txtTotal.Text = 0#
    txtMonto.Text = 0#
    lbDeposito = True
End Sub
'Botones
Private Sub cmdAgregar_Click()
    Dim fila As Integer
    feSolicitudes.AdicionaFila
    feSolicitudes.SetFocus
    SendKeys "{Enter}"
    cmdAgregar.Enabled = False
End Sub
Private Sub cmdQuitar_Click()
    Dim fila As Integer
    Dim nTotal As Currency
    nTotal = 0#
    fila = feSolicitudes.row
    If MsgBox("¿Desea eliminar la fila seleccionada?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        feSolicitudes.EliminaFila fila
    End If
    For i = 1 To feSolicitudes.Rows - 1
        If feSolicitudes.TextMatrix(1, 1) = "" Then
            'Set rsSeg = oNSeg.ObtenerSegTarjetaSolicitudesPendientes(1)
            Exit For
        Else
            If feSolicitudes.TextMatrix(i, 1) <> "" Then
                nTotal = nTotal + CCur(feSolicitudes.TextMatrix(i, 5))
                'lcNumSol = lcNumSol & feSolicitudes.TextMatrix(i, 1) & ","
            End If
        End If
    Next
    txtTotal.Text = Format(nTotal, "#,##0.00")
End Sub
Private Sub cmdAceptar_Click()
    Dim i As Integer
    Dim Total As Currency
    Total = 0#
    For i = 1 To feSolicitudes.Rows - 1
        If feSolicitudes.TextMatrix(i, 1) <> "" Then
            Total = Total + CCur(feSolicitudes.TextMatrix(i, 5))
        End If
    Next
    txtTotal.Text = Format(Total, "#,##0.00")
End Sub
'Fin Botones
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtMonto.SetFocus
End Sub
Private Sub txtFecha_LostFocus()
    If Not IsDate(txtFecha.Text) Then
        MsgBox "Ingrese una fecha valida"
        txtFecha.SetFocus
        Exit Sub
    ElseIf CDate(txtFecha.Text) > gdFecSis Then
        MsgBox "La fecha del deposito, no puede ser posterior a la fecha actual"
        txtFecha.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii)
    If KeyAscii = 13 Then
        'cmdAgregar.SetFocus
        txtMonto.Text = Format(txtMonto.Text, "#,##0.00")
    End If
End Sub
Private Sub txtMonto_LostFocus()
    txtMonto.Text = Format(txtMonto.Text, "#,##0.00")
End Sub
Private Sub txtObjOrig_EmiteDatos()
    If Len(txtObjOrig) > 15 Then
        lblDescbanco = oCtaIf.NombreIF(Mid(txtObjOrig, 4, 13))
        lblDescCtabanco = oCtaIf.EmiteTipoCuentaIF(Mid(txtObjOrig, 18, 10)) + " " + txtObjOrig.psDescripcion
        lsCtaContBanco = oOpe.EmiteOpeCta(gsOpeCod, IIf(lbDeposito, "D", "H"), , txtObjOrig, CtaOBjFiltroIF)
        If lsCtaContBanco = "" Then
            MsgBox "Institución Financiera no tiene definida Cuenta Contable", vbInformation, "Aviso"
        End If
        txtFecha.SetFocus
    End If
End Sub
Private Sub cmdGuardar_Click()
    Dim oNSeg As New NSeguros
    Dim lsCtaInsti As String
    Dim ldFechaDeposito As Date
    Dim lnMontoDeposito As Currency, lnTotal As Currency
    Dim lnTotalOpe As Integer, fila As Integer
    Dim lsPersCodIf As String, lsIFTpo As String, lsCtaIFCod As String, lsMovNro As String
    Dim lbPendiente As Boolean
    Dim lbValidar As Boolean 'FRHU 20140812
    
    Dim depositos() As String
    Dim MatPendiente As Variant
    
    If Not ValidaDatos Then Exit Sub
    
    ldFechaDeposito = CDate(txtFecha.Text)
    lnMontoDeposito = CCur(txtMonto.Text)
    lsPersCodIf = Mid(txtObjOrig, 4, 13)
    lsIFTpo = Left(txtObjOrig, 2)
    lsCtaIFCod = Right(txtObjOrig, Len(txtObjOrig) - 17)
    
    If ldFechaDeposito = gdFecSis Then
        lbPendiente = False
    Else
        MatPendiente = frmSegTarjetaPendientes.Inicio(lsPersCodIf, lsIFTpo, lsCtaIFCod, lnMontoDeposito, ldFechaDeposito)
        If frmSegTarjetaPendientes.nLogico = 0 Then
            MsgBox "No hay ninguna pendiente disponible", vbInformation, "Aviso"
            Exit Sub
        ElseIf frmSegTarjetaPendientes.nLogico = 1 Then
            MsgBox "Debe elegir una pendiente para continuar con el proceso", vbInformation, "Aviso"
            Exit Sub
        End If
        lbPendiente = True
    End If
    
    lnTotalOpe = feSolicitudes.Rows - 1
    ReDim depositos(1 To lnTotalOpe, 1 To 7) As String
    ldFechaDeposito = CDate(txtFecha.Text)
    lnTotal = CCur(txtTotal.Text)
    
    For fila = 1 To lnTotalOpe
        depositos(fila, 1) = feSolicitudes.TextMatrix(fila, 1) 'nSolicitud
        depositos(fila, 2) = feSolicitudes.TextMatrix(fila, 11) 'CuentaAhorro
        depositos(fila, 3) = Format(feSolicitudes.TextMatrix(fila, 5), "#,##0.00") 'Monto
        depositos(fila, 4) = feSolicitudes.TextMatrix(fila, 6) 'Glosa
        depositos(fila, 5) = feSolicitudes.TextMatrix(fila, 8) 'AreaCod
        depositos(fila, 6) = feSolicitudes.TextMatrix(fila, 9) 'AgeCod
        depositos(fila, 7) = feSolicitudes.TextMatrix(fila, 10) 'cObjetoCod
    Next
    '*** FRHU 20140812: Cuando es una fecha anterior no se realizara asientos contables en esta opcion
    'If ldFechaDeposito = gdFecSis Then
    '    lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , txtObjOrig, ObjEntidadesFinancieras)
    'Else
    '    lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", 16, txtObjOrig, ObjEntidadesFinancieras)
    'End If
    'lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", 15, txtObjOrig, ObjEntidadesFinancieras)
    '
    'lsMovNro = oNSeg.GrabarDepositoXActivacionSegTarjeta(lsPersCodIf, ldFechaDeposito, lnMontoDeposito, depositos, gdFecSis, lnTotalOpe, gsCodAge, gsCodUser, _
    '                                               gsOpeCod, lsCtaDebe, txtObjOrig, lsCtaHaber, MatPendiente, lbPendiente)
    'ImprimeAsientoContable lsMovNro
    If ldFechaDeposito = gdFecSis Then
        lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , txtObjOrig, ObjEntidadesFinancieras)
        lsCtaHaber = "29180703" 'Caja General
        lsMovNro = oNSeg.GrabarDepositoXActivacionSegTarjeta(lsPersCodIf, ldFechaDeposito, lnMontoDeposito, depositos, gdFecSis, lnTotalOpe, gsCodAge, gsCodUser, _
                                                   gsOpeCod, lsCtaDebe, txtObjOrig, lsCtaHaber, MatPendiente, lbPendiente)
        ImprimeAsientoContable lsMovNro
    Else
        'lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", 15, txtObjOrig, ObjEntidadesFinancieras)
        lbValidar = oNSeg.GrabarDepositoXActivacionSegTarjetaFechaAnterior(depositos, gdFecSis, lnTotalOpe, gsCodAge, gsCodUser, gsOpeCod, MatPendiente)
    End If
    '*** FIN FRHU 20140812
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Grabo Operación "
        Set objPista = Nothing
        '****
    Call Limpiar
End Sub
Private Sub Limpiar()
    txtFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
    txtTotal.Text = 0#
    lblDescbanco.Caption = ""
    lblDescCtabanco.Caption = ""
    txtObjOrig.Text = ""
    txtTotal.Text = 0#
    txtMonto.Text = 0#
    Call FormateaFlex(feSolicitudes)
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Function ValidaDatos() As Boolean
    Dim fila As Integer, totalfilas As Integer
    Dim Total As Currency
    
    ValidaDatos = True
    
    If txtMonto.Text = "" Then
        MsgBox "Se requiere ingresar el monto del deposito"
        txtMonto.SetFocus
        ValidaDatos = False
        Exit Function
    ElseIf CCur(txtMonto.Text) = 0# Then
        MsgBox "Se requiere ingresar el monto del deposito"
        txtMonto.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    If feSolicitudes.TextMatrix(1, 0) = "" Then
        MsgBox "Debe ingresar por lo menos una solicitud para continuar con la operacion", vbInformation
        ValidaDatos = False
        Exit Function
    End If
    
    totalfilas = feSolicitudes.Rows - 1
    For fila = 1 To totalfilas
        If feSolicitudes.TextMatrix(fila, 6) = "" Then
            MsgBox "Debe ingresar una glosa para cada solicitud", vbInformation
            ValidaDatos = False
            Exit Function
        End If
        Total = Total + IIf(feSolicitudes.TextMatrix(fila, 5) = "", 0, feSolicitudes.TextMatrix(fila, 5))
    Next
    If CCur(Total) <> CCur(txtMonto.Text) Then
        ValidaDatos = False
        MsgBox "El monto del deposito debe ser igual a la suma de los montos de cada solicitud seleccionada", vbInformation, "Aviso"
        Exit Function
    End If
End Function
