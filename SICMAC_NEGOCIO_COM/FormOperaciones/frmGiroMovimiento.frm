VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGiroMovimiento 
   Caption         =   "Giros - Movimientos"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   Icon            =   "frmGiroMovimiento.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6810
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkComision 
      Caption         =   "Cobro Comisión"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   6360
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Movimientos MN"
      TabPicture(0)   =   "frmGiroMovimiento.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FEMovMN"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Movimientos ME"
      TabPicture(1)   =   "frmGiroMovimiento.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FEMovME"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin SICMACT.FlexEdit FEMovMN 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha/Hora-Operación-Nº Giro-Estado-Abono-Cargo-Agencia-Usuario-Receptor"
         EncabezadosAnchos=   "400-2000-2500-2000-1200-1200-1200-2500-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-C-C-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit FEMovME 
         Height          =   3615
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha/Hora-Operación-Nº Giro-Estado-Abono-Cargo-Agencia-Usuario-Receptor"
         EncabezadosAnchos=   "400-2000-2500-2000-1200-1200-1200-1200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame4 
      Height          =   510
      Left            =   2640
      TabIndex        =   18
      Top             =   6240
      Width           =   5295
      Begin VB.Label lblTotalComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3840
         TabIndex        =   22
         Top             =   135
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Total Comisiones:"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   165
         Width           =   1335
      End
      Begin VB.Label lblTotalGirado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1200
         TabIndex        =   20
         Top             =   135
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Total Girado:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   165
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdImpimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   8040
      TabIndex        =   17
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9120
      TabIndex        =   16
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10200
      TabIndex        =   15
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Búsqueda"
      Height          =   1695
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Width           =   3975
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   3735
         Begin VB.TextBox txtNum 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1200
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   375
            Left            =   2160
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtFecFin 
            Height          =   330
            Left            =   480
            TabIndex        =   13
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFecIni 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   480
            TabIndex        =   12
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblNum 
            Caption         =   "Nº"
            Height          =   255
            Left            =   720
            TabIndex        =   24
            Top             =   435
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblAl 
            Caption         =   "Al:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblDel 
            Caption         =   "Del:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.OptionButton optMov 
         Caption         =   "Últimos Movimientos"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optFecha 
         Caption         =   "Rango Fechas"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente:"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin SICMACT.TxtBuscar TxtBuscar 
         Height          =   285
         Left            =   720
         TabIndex        =   30
         Top             =   360
         Width           =   1860
         _extentx        =   3281
         _extenty        =   503
         appearance      =   0
         appearance      =   0
         font            =   "frmGiroMovimiento.frx":0342
         appearance      =   0
         tipobusqueda    =   3
         stitulo         =   ""
      End
      Begin VB.Label lblDOI 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "DOI:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label lblMontoCom 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   1680
      TabIndex        =   29
      Top             =   6360
      Visible         =   0   'False
      Width           =   800
   End
End
Attribute VB_Name = "frmGiroMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nmoneda As Integer
Private Sub chkComision_Click()
    If chkComision.value = 0 Then
        lblMontoCom.Visible = False
    Else
        lblMontoCom.Visible = True
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Dim rsDatos As ADODB.Recordset
    Dim rsDatosME As ADODB.Recordset
    Dim rsTotalesMN As ADODB.Recordset
    Dim rsTotalesME As ADODB.Recordset
    Dim lnTpoOpe As Integer
    
    Dim lnMonTotMN As Double
    Dim lnMonTotME As Double
    Dim lnMonTotComMN As Double
    Dim lnMonTotComME As Double
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set rsDatos = New ADODB.Recordset
    Set rsDatosME = New ADODB.Recordset
    Set rsTotalesMN = New ADODB.Recordset
    Set rsTotalesME = New ADODB.Recordset
    If (Me.optFecha.value = True) Then
        lnTpoOpe = 1
    Else
        lnTpoOpe = 2
    End If
    
    If lnTpoOpe = 1 Then
        If Me.txtFecIni.Text = "__/__/____" Or Me.txtFecFin.Text = "__/__/____" Then
            MsgBox "Los datos no pueden ser vacios", vbInformation, "Alerta"
            Exit Sub
        End If
    Else
        If Me.txtNum.Text = "" Then
            MsgBox "Los datos no pueden ser vacios", vbInformation, "Alerta"
            Exit Sub
        End If
    End If
    Set rsDatos = oServ.HistorialMovimientoGiros(Me.TxtBuscar.Text, lnTpoOpe, Format(Me.txtFecIni.Text, "yyyyMMdd"), Format(Me.txtFecFin.Text, "yyyyMMdd"), 1)
    Set rsDatosME = oServ.HistorialMovimientoGiros(Me.TxtBuscar.Text, lnTpoOpe, Format(Me.txtFecIni.Text, "yyyyMMdd"), Format(Me.txtFecFin.Text, "yyyyMMdd"), 2)
    Set rsTotalesMN = oServ.ObtieneTotalesMovimientoGiro(Me.TxtBuscar.Text, lnTpoOpe, Format(Me.txtFecIni.Text, "yyyyMMdd"), Format(Me.txtFecFin.Text, "yyyyMMdd"), 1)
    Set rsTotalesME = oServ.ObtieneTotalesMovimientoGiro(Me.TxtBuscar.Text, lnTpoOpe, Format(Me.txtFecIni.Text, "yyyyMMdd"), Format(Me.txtFecFin.Text, "yyyyMMdd"), 2)
    If Not (rsDatos.BOF And rsDatos.EOF) Then
        Dim nIndex As Integer
        Dim nCanReg As Integer
        Dim nMaxReg As Integer
        
        If lnTpoOpe = 1 Then
            nCanReg = rsDatos.RecordCount
        Else
            nMaxReg = rsDatos.RecordCount
            nCanReg = Val(Me.txtNum.Text)
            If nCanReg > nMaxReg Then
                nCanReg = nMaxReg
            End If
        End If
         Me.FEMovMN.Clear
         FormateaFlex Me.FEMovMN
        For nIndex = 1 To nCanReg
            Me.FEMovMN.AdicionaFila
            Me.FEMovMN.TextMatrix(nIndex, 1) = rsDatos!dFecha
            Me.FEMovMN.TextMatrix(nIndex, 2) = rsDatos!cOpedesc
            Me.FEMovMN.TextMatrix(nIndex, 3) = rsDatos!cNumRecibo
            Me.FEMovMN.TextMatrix(nIndex, 4) = rsDatos!cGiroEstado
            Me.FEMovMN.TextMatrix(nIndex, 5) = Format(rsDatos!nAbono, gsFormatoNumeroView)
            Me.FEMovMN.TextMatrix(nIndex, 6) = Format(rsDatos!nCargo, gsFormatoNumeroView)
            Me.FEMovMN.TextMatrix(nIndex, 7) = rsDatos!cAgeDescripcion
            Me.FEMovMN.TextMatrix(nIndex, 8) = rsDatos!cUser
            Me.FEMovMN.TextMatrix(nIndex, 9) = rsDatos!cNumDoc
            rsDatos.MoveNext
        Next
    End If
    If Not (rsDatosME.BOF And rsDatosME.EOF) Then
        Dim nIndx As Integer
        Dim nNumReg As Integer
        Dim nCanMaxReg As Integer
        
        If lnTpoOpe = 1 Then
            nCanReg = rsDatosME.RecordCount
        Else
            nCanMaxReg = rsDatosME.RecordCount
            nNumReg = Val(Me.txtNum.Text)
            If nNumReg > nCanMaxReg Then
                nNumReg = nCanMaxReg
            End If
        End If
         Me.FEMovME.Clear
         FormateaFlex Me.FEMovME
        For nIndx = 1 To nNumReg
            Me.FEMovME.AdicionaFila
            Me.FEMovME.TextMatrix(nIndx, 1) = rsDatosME!dFecha
            Me.FEMovME.TextMatrix(nIndx, 2) = rsDatosME!cOpedesc
            Me.FEMovME.TextMatrix(nIndx, 3) = rsDatosME!cNumRecibo
            Me.FEMovME.TextMatrix(nIndx, 4) = rsDatosME!cGiroEstado
            Me.FEMovME.TextMatrix(nIndx, 5) = Format(rsDatosME!nAbono, gsFormatoNumeroView)
            Me.FEMovME.TextMatrix(nIndx, 6) = Format(rsDatosME!nCargo, gsFormatoNumeroView)
            Me.FEMovME.TextMatrix(nIndx, 7) = rsDatosME!cAgeDescripcion
            Me.FEMovME.TextMatrix(nIndx, 8) = rsDatosME!cUser
            Me.FEMovME.TextMatrix(nIndx, 9) = rsDatosME!cNumDoc
            rsDatosME.MoveNext
        Next
    End If
    If Not (rsTotalesMN.BOF And rsTotalesMN.EOF) Then
        lnMonTotMN = Format(rsTotalesMN!nMontoTotGiro, gsFormatoNumeroView)
        lnMonTotComMN = Format(rsTotalesMN!nMontoTotComi, gsFormatoNumeroView)
    End If
    If Not (rsTotalesME.BOF And rsTotalesME.EOF) Then
        lnMonTotME = Format(rsTotalesME!nMontoTotGiro, gsFormatoNumeroView)
        lnMonTotComME = Format(rsTotalesME!nMontoTotComi, gsFormatoNumeroView)
    End If
    Me.lblTotalComision = Format(lnMonTotComMN + lnMonTotComME, gsFormatoNumeroView)
    Me.lblTotalGirado = Format(lnMonTotMN + lnMonTotME, gsFormatoNumeroView)
    Set rsDatos = Nothing
    Set rsDatosME = Nothing
    'Set rsTotales = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarFormulario
End Sub

Private Sub cmdImpimir_Click()
    Dim sPerscod As String
    Dim sMoneda As String
    Dim i As Integer, J As Integer
    Dim sTotAbono As String * 13
    Dim sTotCargo As String * 13
    Dim sTitRp1 As String, sTitRp2 As String, sUser As String
    Dim rsO As New ADODB.Recordset
    Dim rsOP As New ADODB.Recordset
    Dim rsC As New ADODB.Recordset
    Dim clsPrev As previo.clsprevio
    Dim oCapImp As COMNCaptaServicios.NCOMCaptaServicios
    Dim lsCadImp As String
    Dim nCntPag As Integer
    Dim objPista As COMManejador.Pista
    Dim sComentario As String
    Dim bCobrComi As Boolean
    Dim clsGiro As COMNCaptaServicios.NCOMCaptaServicios
    Set clsGiro = New COMNCaptaServicios.NCOMCaptaServicios
    
    Dim lsBoleta As String, nMontoComis As String
    If FEMovME.Rows > 2 Or FEMovMN.Rows > 2 Then
    Else
        MsgBox "No se encontraron datos para imprimir", vbInformation, "Alerta"
        Exit Sub
    End If
    nMontoComis = Val(lblMontoCom.Caption)
    
    sPerscod = TxtBuscar.Text
    sMoneda = IIf(nmoneda = gMonedaNacional, "SOLES", "DOLARES")
    sTitRp1 = Space(55) & "EXTRACTO DE CUENTA"
    If Me.optMov.value Then
        sTitRp2 = Space(55) & "ÚLTIMOS " & Trim(txtNum) & " MOVIMIENTOS "
    Else
        sTitRp2 = Space(55) & "( DEL " & txtFecIni & " AL " & txtFecFin & " )"
    End If
    
    'RSet sTotAbono = lblTotAbono
    'RSet sTotCargo = lblTotCargo

            With rsO
                .Fields.Append "dFecha", adDate
                .Fields.Append "sOperacion", adVarChar, 200
                .Fields.Append "sReceptor", adVarChar, 50
                .Fields.Append "sCtaCod", adVarChar, 50
                .Fields.Append "sEstado", adVarChar, 50
                .Fields.Append "nAbono", adVarChar, 50
                .Fields.Append "nCargo", adVarChar, 200
                .Fields.Append "sAgencia", adVarChar, 50
                .Fields.Append "sUsuario", adVarChar, 6
                .Open
            
                For i = 1 To Me.FEMovMN.Rows - 1
                    .AddNew
                    .Fields("dFecha") = Trim(FEMovMN.TextMatrix(i, 1))
                    
                    .Fields("sOperacion") = FEMovMN.TextMatrix(i, 2)
                    .Fields("sReceptor") = FEMovMN.TextMatrix(i, 9)
                    .Fields("sCtaCod") = FEMovMN.TextMatrix(i, 3)
                    .Fields("sEstado") = FEMovMN.TextMatrix(i, 4)
                    .Fields("nAbono") = FEMovMN.TextMatrix(i, 5)
                    .Fields("nCargo") = FEMovMN.TextMatrix(i, 6)
                    .Fields("sAgencia") = FEMovMN.TextMatrix(i, 7)
                    .Fields("sUsuario") = FEMovMN.TextMatrix(i, 8)
                Next i
                
                If FEMovME.row - 1 > 0 Then
                    For i = 1 To FEMovME.Rows - 1
                        .AddNew
                        .Fields("dFecha") = Trim(FEMovME.TextMatrix(i, 1))
                        .Fields("sOperacion") = FEMovME.TextMatrix(i, 2)
                        .Fields("sReceptor") = FEMovME.TextMatrix(i, 9)
                        .Fields("sCtaCod") = FEMovME.TextMatrix(i, 3)
                        .Fields("sEstado") = FEMovME.TextMatrix(i, 4)
                        .Fields("nAbono") = FEMovME.TextMatrix(i, 5)
                        .Fields("nCargo") = FEMovME.TextMatrix(i, 6)
                        .Fields("sAgencia") = FEMovME.TextMatrix(i, 7)
                        .Fields("sUsuario") = FEMovME.TextMatrix(i, 8)
                 Next i
                End If
                
                rsO.Sort = "dFecha,sOperacion"
                 
            End With
            
            With rsC
                'Crear RecordSet
                .Fields.Append "sNombre", adVarChar, 200
                .Fields.Append "sDOI", adVarChar, 50
                .Open
                'Llenar Recordset
                .AddNew
                .Fields("sNombre") = Trim(Me.lblCliente.Caption)
                .Fields("sDOI") = Trim(Me.lblDOI.Caption)
            End With
    'ServGiroEstadoCuenta
    Set oCapImp = New COMNCaptaServicios.NCOMCaptaServicios
    oCapImp.IniciaImpresora gImpresora
       
    If Me.chkComision.value = 1 Then
        If MsgBox("Se realizara cobro por comisión de movimientos. ¿Desea continuar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            bCobrComi = False
        Else
            bCobrComi = True
        End If
    End If
    On Error GoTo ErrGraba
    
    Dim oCredMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Set oCredMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsMov = New COMNContabilidad.NCOMContFunciones
        
    Dim sMovNro As String
    Dim nMovNro As Long
    Dim fsOpeCod As String, fsGlosa As String
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    
    fsOpeCod = gServGiroComiMov
    fsGlosa = "comisión por emisión de movimientos - GIRO"
    If bCobrComi = True Then
        nMovNro = oCredMov.OtrasOperaciones(sMovNro, gServGiroComiMov, CDbl(lblMontoCom.Caption), "", fsGlosa & " cliente: " & lblCliente.Caption, gMonedaNacional, TxtBuscar.Text, , , , , , , gnMovNro)
    End If
    lsCadImp = oCapImp.ImprimeConsultaMovimiento(rsO, rsC, "Comisión cobrada :" & Format(CDbl(lblMontoCom.Caption), gcFormView) & Chr(10) & "  " & gsCodUser, gsNomCmac, gsNomAge, sMoneda, gdFecSis, sTitRp1, sTitRp2, sPerscod, "", _
                "", 0, 0, 0, 0, 0, "", sTotAbono, sTotCargo, nCntPag)
                
    Set oCapImp = Nothing
    sComentario = "Consulta de Movimientos - GIROS"
    
    Set clsPrev = New previo.clsprevio
    clsPrev.Show lsCadImp, "Extracto de Cuenta", True, , gImpresora
    Set clsPrev = Nothing
    
    If Trim(lsBoleta) <> "" Then
        Dim lbOk As Boolean
        lbOk = True
        Do While lbOk
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
            Print #nFicSal, lsBoleta
            Print #nFicSal, ""
            Close #nFicSal
            If MsgBox("Desea Reimprimir Boleta ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbOk = False
            End If
        Loop
    End If
    Call LimpiarFormulario
    Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsGiroCom As ADODB.Recordset
    Dim clsGiro As COMNCaptaServicios.NCOMCaptaServicios
    nmoneda = 1
    Call LimpiarFormulario
    Set clsGiro = New COMNCaptaServicios.NCOMCaptaServicios
    Set rsGiroCom = New ADODB.Recordset
    Set rsGiroCom = clsGiro.RecuperaValorComisionTarGiro(3)
    
    If Not (rsGiroCom.EOF And rsGiroCom.BOF) Then
        lblMontoCom = Format(IIf(nmoneda = 1, rsGiroCom!nMontoMN, rsGiroCom!nMontoME), "#,##0.00")
    Else
        MsgBox "No se encontró valor de comisión. Comuníquese con el departamento de TI", vbInformation, "SICMACM - Aviso"
    End If
    Set rsGiroCom = Nothing
End Sub

Private Sub optFecha_Click()
    If optFecha.value = True Then
        txtFecIni.Visible = True
        txtFecFin.Visible = True
        lblDel.Visible = True
        lblAl.Visible = True
        lblNum.Visible = False
        txtNum.Visible = False
    End If
End Sub

Private Sub optMov_Click()
    If optMov.value = True Then
        txtFecIni.Visible = False
        txtFecFin.Visible = False
        lblDel.Visible = False
        lblAl.Visible = False
        lblNum.Visible = True
        txtNum.Visible = True
    End If
End Sub




Private Sub TxtBuscar_EmiteDatos()
    lblCliente.Caption = TxtBuscar.psDescripcion
    lblDOI.Caption = TxtBuscar.sPersNroDoc
    Me.optFecha.SetFocus
End Sub

'Private Sub TxtBuscar_EmiteDatos()
'    Dim oPersona As COMDPersona.UCOMPersona
'    Dim sPerscod As String
'
'    'Set oPersona = frmBuscaPersona.Inicio
'
'    'If Not oPersona Is Nothing Then
'    '    sPerscod = oPersona.sPerscod
'    '    TxtBuscar.Text = sPerscod
'    '    lblCliente.Caption = oPersona.sPersNombre
'    '    lblDOI.Caption = Trim(oPersona.sPersIdnroDNI)
'    'End If
'    'Me.optFecha.SetFocus
'End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case Else
            KeyAscii = 0
    End Select
End Sub

Public Sub LimpiarFormulario()
    lblMontoCom.Visible = False
    TxtBuscar.Text = ""
    lblCliente.Caption = ""
    lblDOI.Caption = ""
    txtFecIni.Text = "__/__/____"
    txtFecFin.Text = "__/__/____"
    txtNum.Text = ""
    FEMovMN.Clear
    FormateaFlex FEMovMN
    FEMovME.Clear
    FormateaFlex FEMovME
    'lblMontoCom.Caption = ""
    lblTotalGirado.Caption = ""
    lblTotalComision.Caption = ""
    chkComision.value = 0
End Sub


