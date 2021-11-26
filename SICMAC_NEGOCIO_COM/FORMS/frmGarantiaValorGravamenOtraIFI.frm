VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantiaValorGravamenOtraIFI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gravamen a favor de Otra(s) IFI(s)"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "frmGarantiaValorGravamenOtraIFI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      ToolTipText     =   "Aceptar"
      Top             =   5370
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2970
      TabIndex        =   8
      ToolTipText     =   "Cancelar"
      Top             =   5370
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5220
      Left            =   75
      TabIndex        =   9
      Top             =   75
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   9208
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Valorización Directa"
      TabPicture(0)   =   "frmGarantiaValorGravamenOtraIFI.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraDatos 
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
         Height          =   4725
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   5595
         Begin VB.Frame frOp 
            Caption         =   "Registros Públicos :"
            Height          =   600
            Left            =   120
            TabIndex        =   18
            Top             =   1185
            Width           =   2595
            Begin VB.OptionButton opIfi_PerNJ 
               Caption         =   "IFI"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   250
               Width           =   615
            End
            Begin VB.OptionButton opIfi_PerNJ 
               Caption         =   "P. Natural/Juridica"
               Height          =   255
               Index           =   2
               Left            =   870
               TabIndex        =   19
               Top             =   250
               Width           =   1695
            End
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   285
            Left            =   120
            MaxLength       =   300
            TabIndex        =   0
            Tag             =   "txtPrincipal"
            Top             =   360
            Width           =   5370
         End
         Begin VB.TextBox txtVRMDisponible 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4200
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   4320
            Width           =   1260
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   555
            TabIndex        =   4
            ToolTipText     =   "Quitar"
            Top             =   3720
            Width           =   375
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Agregar"
            Top             =   3720
            Width           =   375
         End
         Begin VB.ComboBox cmbMoneda 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   1515
         End
         Begin VB.TextBox txtTotalVRM 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4200
            MaxLength       =   15
            TabIndex        =   2
            Text            =   "0.00"
            Top             =   720
            Width           =   1260
         End
         Begin VB.TextBox txtTotalGravado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4200
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   3840
            Width           =   1260
         End
         Begin SICMACT.FlexEdit feGravamen 
            Height          =   1890
            Left            =   120
            TabIndex        =   11
            Top             =   1800
            Width           =   5340
            _ExtentX        =   9419
            _ExtentY        =   3334
            Cols0           =   5
            HighLight       =   2
            EncabezadosNombres=   "N°-Código-Institución Financiera-Gravado-Aux"
            EncabezadosAnchos=   "400-1400-2200-1200-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-3-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-1-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R-C"
            FormatosEdit    =   "0-0-0-2-0"
            CantEntero      =   12
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descripción de Garantía :"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   150
            Width           =   1830
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   120
            X2              =   5450
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Registros Públicos :"
            Height          =   195
            Left            =   4080
            TabIndex        =   16
            Top             =   1560
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            X1              =   120
            X2              =   5450
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Monto disponible :"
            Height          =   195
            Left            =   2715
            TabIndex        =   15
            Top             =   4365
            Width           =   1290
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Valor de Garantía :"
            Height          =   195
            Left            =   2760
            TabIndex        =   13
            Top             =   750
            Width           =   1350
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Gravado :"
            Height          =   195
            Left            =   2895
            TabIndex        =   12
            Top             =   3885
            Width           =   1110
         End
      End
   End
End
Attribute VB_Name = "frmGarantiaValorGravamenOtraIFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************
'** Nombre : frmGarantiaValorGravamenOtraIFI
'** Descripción : Para registro de los gravamenes que se tiene en otra IFIs creado segun TI-ERS002-2016
'** Creación : EJVG, 20160215 06:48:15 PM
'******************************************************************************************************
Option Explicit

Dim fbRegistrar As Boolean
Dim fbEditar As Boolean
Dim fbConsultar As Boolean

Dim fbOk As Boolean
Dim fbPrimero As Boolean
Dim fnMoneda As Moneda
Dim fvValorGravamenFavorOtraIFi As tValorGravamenFavorOtraIFi
Dim fvValorGravamenFavorOtraIFi_ULT As tValorGravamenFavorOtraIFi
Dim fbFocoGrilla As Boolean
Dim fnTipoRegPubli As Integer 'JOEP20180524 ERS009-2018

Private Sub cmbMoneda_Click()
    Dim lnMoneda As Moneda
    Dim lnColor As Long
    
    lnMoneda = val(Trim(Right(cmbMoneda.Text, 3)))
    lnColor = &H80000005
    
    If lnMoneda = gMonedaExtranjera Then
        lnColor = &HC0FFC0
    End If
    
    txtTotalVRM.BackColor = lnColor
    txtTotalGravado.BackColor = lnColor
    txtVRMDisponible.BackColor = lnColor
End Sub
Private Sub cmdQuitar_Click()
    If feGravamen.TextMatrix(1, 0) = "" Then Exit Sub
    
    If MsgBox("Se va a eliminar el item N° " & feGravamen.row & Chr(13) & "¿Está seguro de continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    feGravamen.EliminaFila feGravamen.row
    SumarizarVRM
End Sub

Private Sub feGravamen_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 3 Then
        SumarizarVRM
    End If
End Sub

Private Sub feGravamen_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(feGravamen.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
    If pnCol = 3 Then
        If Not IsNumeric(feGravamen.TextMatrix(pnRow, pnCol)) Then
            MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
            feGravamen.row = pnRow
            feGravamen.TopRow = pnRow
            feGravamen.Col = pnCol
            Cancel = False
            Exit Sub
        End If
    End If
End Sub
Private Sub Form_Load()
    Dim lvIFi() As tValorGravamenFavorOtraIFiDet
    Dim index As Integer
    
    fnTipoRegPubli = 0 'JOEP20180524 ERS009-2018
    
    fbOk = False
    Screen.MousePointer = 11
    
    CargarControles
    LimpiarControles
    
    If fbEditar Or fbConsultar Then
        frOp.Enabled = False 'JOEP20180524 ERS009-2018
        
        txtDescripcion.Text = UCase(fvValorGravamenFavorOtraIFi.sDescripcion)
        txtTotalVRM.Text = Format(fvValorGravamenFavorOtraIFi.nValorComercial, "#,##0.00")
          
         opIfi_PerNJ(fvValorGravamenFavorOtraIFi.nOpTpRegPub).value = True  'JOEP20180524 ERS009-2018
          
        lvIFi = fvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet
        For index = 1 To UBound(lvIFi)
            feGravamen.AdicionaFila
                        
            If fnTipoRegPubli = 2 Then 'JOEP20180524 ERS009-2018
                feGravamen.TextMatrix(feGravamen.row, 1) = lvIFi(index).sIFICod 'JOEP20180524 ERS009-2018
            Else
                feGravamen.TextMatrix(feGravamen.row, 1) = lvIFi(index).sIFTpo & "." & lvIFi(index).sIFICod
            End If
            
            feGravamen.TextMatrix(feGravamen.row, 2) = lvIFi(index).sIFINombre
            feGravamen.TextMatrix(feGravamen.row, 3) = Format(lvIFi(index).nGravado, "#,##0.00")
        Next
        
        If fbConsultar Then
            fraDatos.Enabled = False
            cmdAceptar.Enabled = False
        End If
    End If
    
    If fbPrimero Then
       fnMoneda = gMonedaNacional
    End If
    cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
    cmbMoneda.Enabled = False
    
    If fbRegistrar Then
        txtDescripcion.Text = UCase(fvValorGravamenFavorOtraIFi_ULT.sDescripcion)
        txtTotalVRM.Text = Format(fvValorGravamenFavorOtraIFi_ULT.nValorComercial, "#,##0.00")
         
        lvIFi = fvValorGravamenFavorOtraIFi_ULT.vValorGravamenFavorOtraIFiDet
        If IsArray(lvIFi) Then
            For index = 1 To UBound(lvIFi)
                feGravamen.AdicionaFila
                feGravamen.TextMatrix(feGravamen.row, 1) = lvIFi(index).sIFICod
                feGravamen.TextMatrix(feGravamen.row, 1) = lvIFi(index).sIFINombre
                feGravamen.TextMatrix(feGravamen.row, 1) = Format(lvIFi(index).nGravado, "#,##0.00")
            Next
        End If
    End If
    
    feGravamen.row = 1
    feGravamen.TopRow = 1
    
    SumarizarVRM
    
    If fbRegistrar Then
        Caption = "GRAVAMEN A FAVOR DE OTRA(s) IFI(s) [ NUEVO ]"
    End If
    If fbConsultar Then
        Caption = "GRAVAMEN A FAVOR DE OTRA(s) IFI(s) [ CONSULTAR ]"
    End If
    If fbEditar Then
        Caption = "GRAVAMEN A FAVOR DE OTRA(s) IFI(s) [ EDITAR ]"
    End If
    
    Screen.MousePointer = 0
End Sub
Public Function Registrar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef pvValorGravamenFavorOtraIFi As tValorGravamenFavorOtraIFi, ByRef pvValorGravamenFavorOtraIFi_ULT As tValorGravamenFavorOtraIFi) As Boolean
    fbRegistrar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fvValorGravamenFavorOtraIFi = pvValorGravamenFavorOtraIFi
    fvValorGravamenFavorOtraIFi_ULT = pvValorGravamenFavorOtraIFi_ULT
    Show 1
    pnMoneda = fnMoneda
    pvValorGravamenFavorOtraIFi = fvValorGravamenFavorOtraIFi
    
    Registrar = fbOk
End Function
Public Function Editar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef pvValorGravamenFavorOtraIFi As tValorGravamenFavorOtraIFi, ByRef pvValorGravamenFavorOtraIFi_ULT As tValorGravamenFavorOtraIFi) As Boolean
    fbEditar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fvValorGravamenFavorOtraIFi = pvValorGravamenFavorOtraIFi
    fvValorGravamenFavorOtraIFi_ULT = pvValorGravamenFavorOtraIFi_ULT
    Show 1
    pnMoneda = fnMoneda
    pvValorGravamenFavorOtraIFi = fvValorGravamenFavorOtraIFi
    
    Editar = fbOk
End Function
Public Sub Consultar(ByVal pnMoneda As Moneda, ByRef pvValorGravamenFavorOtraIFi As tValorGravamenFavorOtraIFi)
    fbConsultar = True
    fnMoneda = pnMoneda
    fvValorGravamenFavorOtraIFi = pvValorGravamenFavorOtraIFi
    Show 1
End Sub
Private Sub cmdAceptar_Click()
    Dim vDet() As tValorGravamenFavorOtraIFiDet
    Dim i As Integer
    On Error GoTo ErrAceptar
    
    SumarizarVRM
    
    If Len(Trim(txtDescripcion.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la descripción de la Garantía", vbInformation, "Aviso"
        EnfocaControl txtDescripcion
        Exit Sub
    End If
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Moneda", vbInformation, "Aviso"
        EnfocaControl cmbMoneda
        Exit Sub
    End If
    If Not IsNumeric(txtTotalVRM.Text) Then
        MsgBox "Ud. debe ingresar el Total del VRM", vbInformation, "Aviso"
        EnfocaControl txtTotalVRM
        Exit Sub
    Else
        If CCur(txtTotalVRM.Text) <= 0 Then
            MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtTotalVRM
            Exit Sub
        End If
    End If
    
    If Not validarDetalle Then Exit Sub
    
    If Not IsNumeric(txtVRMDisponible.Text) Then
        MsgBox "El monto del VRM disponible debe ser mayor a cero", vbInformation, "Aviso"
        EnfocaControl txtVRMDisponible
        Exit Sub
    Else
        If CCur(txtVRMDisponible.Text) <= 0 Then
            MsgBox "El monto del VRM disponible debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtVRMDisponible
            Exit Sub
        End If
    End If
    
    fnMoneda = CInt(Trim(Right(cmbMoneda.Text, 3)))
    fvValorGravamenFavorOtraIFi.sDescripcion = Trim(txtDescripcion.Text)
    fvValorGravamenFavorOtraIFi.nValorComercial = CCur(txtTotalVRM.Text)
    fvValorGravamenFavorOtraIFi.nOpTpRegPub = fnTipoRegPubli 'JOEP20180524 ERS009-2018
    
    For i = 1 To feGravamen.rows - 1
        ReDim Preserve vDet(i)
        vDet(i).sIFTpo = Left(feGravamen.TextMatrix(i, 1), 2)
        
        If fnTipoRegPubli = 2 Then 'JOEP20180524 ERS009-2018
            vDet(i).sIFICod = feGravamen.TextMatrix(i, 1) 'JOEP20180524 ERS009-2018
        Else
            vDet(i).sIFICod = Mid(feGravamen.TextMatrix(i, 1), 4, 13)
        End If
        
        vDet(i).sIFINombre = feGravamen.TextMatrix(i, 2)
        vDet(i).nGravado = CCur(feGravamen.TextMatrix(i, 3))
    Next
    
    fvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet = vDet
    
    fbOk = True
    Unload Me
    Exit Sub
ErrAceptar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdAgregar_Click()
    If val(txtTotalVRM.Text) <= 0 Then
        MsgBox "Ud. debe ingresar el Total del VRM", vbInformation, "Aviso"
        EnfocaControl txtTotalVRM
        Exit Sub
    End If
        
'JOEP20180524 ERS009-2018
    If fnTipoRegPubli = 0 Then
        MsgBox "Ud. debe seleccionar la Opción de Registros Públicos", vbInformation, "Aviso"
        Exit Sub
    End If
'JOEP20180524 ERS009-2018
        
    If feGravamen.TextMatrix(1, 0) <> "" Then
        If Not validarDetalle Then Exit Sub
    End If
    
    feGravamen.AdicionaFila
    
    EnfocaControl feGravamen

    SendKeys "{Enter}"
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

'JOEP20180524 ERS009-2018
Private Sub opIfi_PerNJ_Click(index As Integer)
fnTipoRegPubli = index
    If fnTipoRegPubli = 1 Then
        feGravamen.TipoBusqueda = BuscaArbol
        feGravamen.EncabezadosNombres = "N°-Código-Institución Financiera-Gravado-Aux"
        LimpiaFlex feGravamen
    ElseIf fnTipoRegPubli = 2 Then
        feGravamen.TipoBusqueda = BuscaPersona
        feGravamen.EncabezadosNombres = "N°-Código-Persona/Juridica-Gravado-Aux"
        LimpiaFlex feGravamen
    End If
End Sub
'JOEP20180524 ERS009-2018

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmbMoneda.Enabled And cmbMoneda.Visible Then
            EnfocaControl cmbMoneda
        Else
            EnfocaControl txtTotalVRM
        End If
    End If
End Sub
Private Sub txtDescripcion_LostFocus()
    txtDescripcion.Text = Trim(UCase(txtDescripcion.Text))
End Sub
Private Sub txtTotalVRM_Change()
    SumarizarVRM
End Sub
Private Sub txtTotalVRM_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTotalVRM, KeyAscii, 15)
    
    If KeyAscii = 13 Then
        EnfocaControl cmdAgregar
    End If
End Sub
Private Sub txtTotalVRM_LostFocus()
    txtTotalVRM.Text = Format(txtTotalVRM, "#,##0.00")
End Sub
Private Sub CargarControles()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim oDGar As New COMDCredito.DCOMGarantia
    Dim rs As New ADODB.Recordset
    
    Set rs = oCons.RecuperaConstantes(1011)
    Call Llenar_Combo_con_Recordset(rs, cmbMoneda)
    
    feGravamen.rsTextBuscar = oDGar.RecuperaIFixGravamenOtraIFi()
    
    RSClose rs
    Set oCons = Nothing
    Set oDGar = Nothing
End Sub
Private Sub LimpiarControles()
    cmbMoneda.ListIndex = -1
    txtTotalVRM.Text = "0.00"
    txtTotalGravado.Text = "0.00"
    txtVRMDisponible.Text = "0.00"
    FormateaFlex feGravamen
End Sub
Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtTotalVRM
    End If
End Sub
Private Sub txtVRMDisponible_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVRMDisponible, KeyAscii, 15)
    
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub txtVRMDisponible_LostFocus()
    txtVRMDisponible.Text = Format(txtVRMDisponible, "#,##0.00")
End Sub
Private Sub txtTotalGravado_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTotalGravado, KeyAscii, 15)
    
    If KeyAscii = 13 Then
        EnfocaControl txtVRMDisponible
    End If
End Sub
Private Sub txtTotalGravado_LostFocus()
    txtTotalGravado.Text = Format(txtTotalGravado, "#,##0.00")
End Sub
Public Function validarDetalle() As Boolean
    Dim i As Integer, j As Integer
    
    If feGravamen.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe ingresar el Valor Gravado de las IFIs", vbInformation, "Aviso"
        EnfocaControl cmdAgregar
        Exit Function
    End If
    For i = 1 To feGravamen.rows - 1
        For j = 1 To feGravamen.cols - 1
            If feGravamen.ColWidth(j) > 0 Then
                If Len(Trim(feGravamen.TextMatrix(i, j))) = 0 Then
                    MsgBox "El campo " & UCase(feGravamen.TextMatrix(0, j)) & " está vacio, verifique..", vbInformation, "Aviso"
                    EnfocaControl feGravamen
                    feGravamen.TopRow = i
                    feGravamen.row = i
                    feGravamen.Col = j
                    Exit Function
                End If
            End If
        Next
    Next
    For i = 1 To feGravamen.rows - 1
        'Valida se hayan ingresado monto
        If val(feGravamen.TextMatrix(i, 3)) = 0 Then
            MsgBox "Ud. debe ingresar un valor mayor a cero", vbInformation, "Aviso"
            EnfocaControl feGravamen
            feGravamen.TopRow = i
            feGravamen.row = i
            feGravamen.Col = 3
            Exit Function
        End If
    Next
    validarDetalle = True
End Function
Private Sub SumarizarVRM()
    Dim lnTotalGravado As Currency
    Dim lnTotalVRM As Currency
    
    lnTotalGravado = CCur(feGravamen.SumaRow(3))
    txtTotalGravado.Text = Format(lnTotalGravado, "#,##0.00")
    
    txtVRMDisponible.Text = "0.00"
    If IsNumeric(txtTotalVRM.Text) Then
        If (Len(txtTotalVRM.Text) >= 14) Then
            txtTotalVRM.Text = Left(txtTotalVRM.Text, 13)
        End If
        lnTotalVRM = CCur(txtTotalVRM.Text)
    End If
    
    txtVRMDisponible.Text = Format(lnTotalVRM - lnTotalGravado, "#,##0.00")
End Sub
