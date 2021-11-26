VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogSubastaIni 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "frmLogSubastaIni.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   360
      Left            =   6240
      TabIndex        =   11
      Top             =   5790
      Width           =   1155
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   7485
      TabIndex        =   10
      Top             =   5790
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   8715
      TabIndex        =   9
      Top             =   5790
      Width           =   1155
   End
   Begin VB.Frame fraSubasta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Subasta"
      ForeColor       =   &H00800000&
      Height          =   1080
      Left            =   105
      TabIndex        =   1
      Top             =   15
      Width           =   9795
      Begin MSMask.MaskEdBox mskIni 
         Height          =   300
         Left            =   1620
         TabIndex        =   13
         Top             =   690
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtSubasta 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1575
         TabIndex        =   12
         Top             =   270
         Width           =   8100
      End
      Begin Sicmact.TxtBuscar txtSubastaCod 
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   255
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   300
         Left            =   8280
         TabIndex        =   16
         Top             =   690
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblEstado 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   3300
         TabIndex        =   22
         Top             =   735
         Width           =   3465
      End
      Begin VB.Label lblFin 
         Caption         =   "Fin :"
         Height          =   180
         Left            =   7020
         TabIndex        =   15
         Top             =   750
         Width           =   915
      End
      Begin VB.Label lblIni 
         Caption         =   "Inicio :"
         Height          =   240
         Left            =   150
         TabIndex        =   14
         Top             =   720
         Width           =   780
      End
   End
   Begin TabDlg.SSTab sTab 
      Height          =   4530
      Left            =   90
      TabIndex        =   0
      Top             =   1185
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   7990
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Miembros"
      TabPicture(0)   =   "frmLogSubastaIni.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraMiembros"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bienes"
      TabPicture(1)   =   "frmLogSubastaIni.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraBB"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Resultados"
      TabPicture(2)   =   "frmLogSubastaIni.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   4020
         Left            =   -74925
         TabIndex        =   18
         Top             =   375
         Width           =   9570
         Begin Sicmact.FlexEdit flexResult 
            Height          =   3435
            Left            =   75
            TabIndex        =   19
            Top             =   195
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   6059
            Cols0           =   7
            EncabezadosNombres=   "#-NroIng-Codigo-Bien-Cantidad Ven-Monto Ven-Estado"
            EncabezadosAnchos=   "400-1200-1200-3000-1200-1200-700"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-6"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-4"
            EncabezadosAlineacion=   "C-R-L-L-R-R-L"
            FormatosEdit    =   "0-0-0-0-2-2-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
         End
         Begin VB.Label lblTotalG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   7080
            TabIndex        =   21
            Top             =   3675
            Width           =   1215
         End
         Begin VB.Label lblTotal 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total"
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
            Height          =   255
            Left            =   5910
            TabIndex        =   20
            Top             =   3675
            Width           =   2400
         End
      End
      Begin VB.Frame fraBB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   4020
         Left            =   75
         TabIndex        =   4
         Top             =   375
         Width           =   9570
         Begin Sicmact.FlexEdit flexBB 
            Height          =   3735
            Left            =   75
            TabIndex        =   8
            Top             =   195
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   6588
            Cols0           =   15
            EncabezadosNombres=   $"frmLogSubastaIni.frx":035E
            EncabezadosAnchos=   "400-1200-1200-4000-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-R-L-L-L-R-R-R-R-R-R-R-R-R-L"
            FormatosEdit    =   "0-0-0-0-0-0-2-2-2-2-2-2-2-3-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
         End
      End
      Begin VB.Frame fraMiembros 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   4035
         Left            =   -74925
         TabIndex        =   3
         Top             =   375
         Width           =   9570
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   345
            Left            =   1155
            TabIndex        =   7
            Top             =   3570
            Width           =   1005
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            Height          =   345
            Left            =   90
            TabIndex        =   6
            Top             =   3570
            Width           =   1005
         End
         Begin Sicmact.FlexEdit flexMiembros 
            Height          =   3285
            Left            =   75
            TabIndex        =   5
            Top             =   195
            Width           =   9420
            _ExtentX        =   16616
            _ExtentY        =   5794
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Relacion"
            EncabezadosAnchos=   "300-1200-4200-3000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-3"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-0-3"
            EncabezadosAlineacion=   "C-L-L-L"
            FormatosEdit    =   "0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   7485
      TabIndex        =   17
      Top             =   5790
      Width           =   1155
   End
End
Attribute VB_Name = "frmLogSubastaIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String
Dim lsOpeCod As String
Dim lbImprime As Boolean
Dim lnMovNroG As Long
Dim lbCierre As Boolean

Public Sub Ini(psOpeCod As String, psCaption As String, pbImprime As Boolean, Optional pbCierre As Boolean = False)
    lsCaption = psCaption
    lsOpeCod = psOpeCod
    lbImprime = pbImprime
    lbCierre = pbCierre
    Me.Show 1
End Sub

Public Function Valida() As Boolean
    Dim i As Integer
    Dim lbPresidente As Boolean
    Dim lbMartllero As Boolean
    Dim lbObservador As Boolean
    
    If Me.txtSubasta.Text = "" Then
        MsgBox "Debe Ingresar un comentario.", vbInformation, "Aviso"
        Me.txtSubasta.SetFocus
        Valida = False
        Exit Function
    ElseIf Not IsDate(Me.mskIni.Text) Then
        MsgBox "Debe Ingresar una fecha valida.", vbInformation, "Aviso"
        Me.mskIni.SetFocus
        Valida = False
        Exit Function
    ElseIf Not IsDate(Me.mskFin.Text) Then
        MsgBox "Debe Ingresar una fecha valida.", vbInformation, "Aviso"
        Me.mskFin.SetFocus
        Valida = False
        Exit Function
    End If
    
    lbPresidente = True
    lbMartllero = True
    lbObservador = True
    
    For i = 1 To Me.flexMiembros.Rows - 1
        If Trim(Right(flexMiembros.TextMatrix(i, 3), 5)) = "1" Then
            lbPresidente = False
        ElseIf Trim(Right(flexMiembros.TextMatrix(i, 3), 5)) = "5" Then
            lbMartllero = False
        ElseIf Trim(Right(flexMiembros.TextMatrix(i, 3), 5)) = "4" Then
            lbObservador = False
        ElseIf Trim(Right(flexMiembros.TextMatrix(i, 3), 5)) = "" Then
            MsgBox "Debe Ingresar un cargo en el comite de venta.", vbInformation
            Valida = False
            flexMiembros.Col = 3
            flexMiembros.Row = i
            Me.flexMiembros.SetFocus
            Exit Function
        End If
    Next i
    
    If lbPresidente Then
        MsgBox "Debe Ingresar al presidente de la comision de venta.", vbInformation
        Valida = False
        Me.flexMiembros.SetFocus
        Exit Function
    ElseIf lbObservador Then
        MsgBox "Debe Ingresar al observador de la comision de venta.", vbInformation
        Valida = False
        Me.flexMiembros.SetFocus
        Exit Function
    ElseIf lbMartllero Then
        MsgBox "Debe Ingresar al martillero.", vbInformation
        Valida = False
        Me.flexMiembros.SetFocus
        Exit Function
    End If
    
    If Me.flexBB.TextMatrix(1, 1) = "" Then
        MsgBox "No puede crear una subasta si no hay bienes a subastar.", vbInformation
        Valida = False
        Me.cmdSalir.SetFocus
        Exit Function
    End If
    
    Valida = True
End Function

Private Sub cmdAgregar_Click()
    flexMiembros.AdicionaFila
    Me.flexMiembros.SetFocus
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea Eliminar el registro " & Me.flexMiembros.Row, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    flexMiembros.EliminaFila flexMiembros.Row
End Sub

Private Sub CmdGrabar_Click()
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim oMov As DMov
    Set oMov = New DMov
    Dim oSubasta As DSubasta
    Set oSubasta = New DSubasta
    
    If lbCierre Then
        If MsgBox("Desae cerrar el proceso de Subasta ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        oSubasta.CierreRemate lnMovNroG
        lsMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        oMov.InsertaMov lsMovNro, lsOpeCod, Me.Caption & " - " & Me.txtSubasta.Text, gMovEstContabMovContable, gMovFlagVigente
        
    Else
        If Not Valida Then Exit Sub
        
        If MsgBox("Desae Iniciar el proceso de Subasta, si elige si no podra eliminar ese proceso de remate ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        lsMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        
        oMov.InsertaMov lsMovNro, lsOpeCod, Me.txtSubasta.Text, gMovEstContabMovContable, gMovFlagVigente
        
        lnMovNro = oMov.GetnMovNro(lsMovNro)
            
        
        oSubasta.InicioRemate lnMovNro, Me.txtSubastaCod.Text, CDate(Me.mskIni.Text), CDate(Me.mskFin.Text), Me.flexMiembros.GetRsNew, Me.flexBB.GetRsNew
    End If
    
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    Dim lsCadena As String
    Dim lnI As Integer
    Dim lnCorr As Integer
    
    
    Dim lsCorr As String * 5
    Dim lsNro As String * 12
    Dim lsCodigo As String * 15
    Dim lsDescripcion As String * 50
    Dim lsValor As String * 15
    Dim lsCantidad As String * 10
    
    Dim lnPag As Long
    Dim lnItem  As Long
    lsCadena = ""
    lnPag = 0
    lnItem = 0
    lsCadena = lsCadena & CabeceraPagina("LISTADOS DE SUBASTA", lnPag, lnItem, gsNomAge, gsEmpresa, gdFecSis)
    lsCadena = lsCadena & Encabezado("Item;6; ;2;Lote;6; ;2;Codigo;15; ;5;Descripcion;20; ;10;Cant;35; ;2;Valor;15; ;1;", lnItem)
    
    If flexBB.TextMatrix(1, 1) = "" Then Exit Sub
    
    For lnI = 1 To Me.flexBB.Rows - 1
        lnItem = lnItem + 1
        lsCorr = Format(lnI, "00000")
        lsNro = flexBB.TextMatrix(lnI, 1)
        lsCodigo = flexBB.TextMatrix(lnI, 2)
        lsDescripcion = flexBB.TextMatrix(lnI, 3)
        RSet lsValor = Format(flexBB.TextMatrix(lnI, 10), "#,##0.00")
        RSet lsCantidad = Format(flexBB.TextMatrix(lnI, 6) - flexBB.TextMatrix(lnI, 11), "#,##0.00")
        
        lsCadena = lsCadena & "  " & lsCorr & "  " & lsNro & "  " & lsCodigo & "  " & lsDescripcion & "  " & lsCantidad & "  " & lsValor & oImpresora.gPrnSaltoLinea
        
        If lnItem > 54 Then
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & CabeceraPagina("LISTADOS DE SUBASTA", lnPag, lnItem, gsNomAge, gsEmpresa, gdFecSis)
            lsCadena = lsCadena & Encabezado("Lote;4; ;2;Codigo;7; ;5;Descripcion;12; ;10;Cant;5; ;2;Valor;7; ;3;", lnItem)
        End If
    Next lnI
    
    oPrevio.Show lsCadena, "LISTADOS DE SUBASTA", True, 66, gImpresora
    
End Sub

Private Sub cmdNuevo_Click()
    Dim oSubasta As DSubasta
    Set oSubasta = New DSubasta
    
    Me.txtSubastaCod.Text = oSubasta.GetCodigo(Format(gdFecSis, "yyyy"))
    Me.txtSubastaCod.Enabled = False
    Me.cmdNuevo.Enabled = False
    
    Me.flexBB.rsFlex = oSubasta.GetNuevaSubasta
    
    Me.txtSubasta.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim oSubasta As DSubasta
    Set oSubasta = New DSubasta
    
    flexMiembros.CargaCombo oCon.GetConstante(5009, , , , , , True)
    
    Caption = lsCaption
    sTab.Tab = 0
    
    If lbImprime And Not lbCierre Then
        Me.cmdAgregar.Visible = False
        Me.cmdEliminar.Visible = False
        Me.cmdNuevo.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdImprimir.Visible = True
        Me.txtSubastaCod.rs = oSubasta.GetSubasta(False, False)
    Else
        Me.cmdImprimir.Visible = False
    End If

    Me.flexResult.lbEditarFlex = False
    
    If lbCierre Then
        Me.cmdAgregar.Visible = False
        Me.cmdEliminar.Visible = False
        Me.cmdNuevo.Visible = False
        Me.cmdGrabar.Visible = True
        Me.cmdImprimir.Visible = True
        Me.txtSubastaCod.rs = oSubasta.GetSubasta(True, lbCierre)
        Me.flexResult.lbEditarFlex = True
    End If
End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAgregar.SetFocus
    End If
End Sub

Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 50
End Sub

Private Sub mskFin_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 50
End Sub

Private Sub mskIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskFin.SetFocus
    End If
End Sub

Private Sub txtSubasta_GotFocus()
    txtSubasta.SelStart = 0
    txtSubasta.SelLength = 300
End Sub

Private Sub txtSubasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskIni.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtSubastaCod_EmiteDatos()
    Dim oSubasta As DSubasta
    Set oSubasta = New DSubasta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lnI As Integer
    Dim lnResult As Currency
    
    If Me.txtSubastaCod.Text <> "" Then
        txtSubasta.Text = txtSubastaCod.psDescripcion
        lnMovNroG = CLng(Mid(txtSubastaCod.psDescripcion, InStr(1, txtSubastaCod.psDescripcion, "[") + 1, InStr(1, txtSubastaCod.psDescripcion, "]") - InStr(1, txtSubastaCod.psDescripcion, "[") - 1))
        Me.flexMiembros.rsFlex = oSubasta.GetSubastaPers(lnMovNroG)
        Me.flexBB.rsFlex = oSubasta.GetSubastaDetalle(lnMovNroG)
        Me.flexResult.rsFlex = oSubasta.GetSubastaDetalleResul(lnMovNroG)
        
        Set rs = oSubasta.GetSubastaDet(lnMovNroG)
        
        Me.mskIni = Format(rs!dInicio, gsFormatoFechaView)
        Me.mskFin = Format(rs!dFin, gsFormatoFechaView)
        
        If rs!bCerrada Then
            Me.lblEstado = "Cerrado"
            Me.flexResult.lbEditarFlex = False
            Me.cmdGrabar.Enabled = False
        Else
            Me.lblEstado = "Abierto"
            Me.flexResult.lbEditarFlex = True
            Me.cmdGrabar.Enabled = True
        End If
        
        lnResult = 0
        
        For lnI = 1 To Me.flexResult.Rows - 1
            lnResult = lnResult + CCur(flexResult.TextMatrix(lnI, 4))
        Next lnI
                
        Me.lblTotalG.Caption = Format(lnResult, "#,##0.00")
    End If
End Sub
