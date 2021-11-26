VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInversionesExtorno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   1965
   ClientTop       =   2925
   ClientWidth     =   11835
   Icon            =   "frmInversionesExtorno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   11835
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   11580
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   345
         Left            =   10260
         TabIndex        =   15
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   8760
         TabIndex        =   14
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame FraConcepto 
      Height          =   1395
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   11580
      Begin VB.TextBox txtMovDesc 
         Height          =   750
         Left            =   870
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   195
         Width           =   10635
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11580
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
         Left            =   10005
         TabIndex        =   2
         Top             =   203
         Width           =   1335
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   345
         Left            =   4095
         TabIndex        =   3
         Top             =   225
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   4210816
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
         Height          =   345
         Left            =   5805
         TabIndex        =   4
         Top             =   225
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   4210816
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
      Begin VB.Label lbltitulo 
         AutoSize        =   -1  'True
         Caption         =   "EXTORNO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   270
         Left            =   90
         TabIndex        =   8
         Top             =   255
         Width           =   1140
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   5205
         TabIndex        =   7
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   3435
         TabIndex        =   6
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
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
         Left            =   6960
         TabIndex        =   5
         Top             =   270
         Width           =   495
      End
   End
   Begin Sicmact.FlexEdit fgIF 
      Height          =   2625
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   4630
      Cols0           =   19
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmInversionesExtorno.frx":030A
      EncabezadosAnchos=   "350-0-900-0-800-1800-2500-1200-800-900-600-900-1200-900-0-0-0-0-0"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   11861226
      BackColorControl=   11861226
      BackColorControl=   11861226
      EncabezadosAlineacion=   "C-L-L-R-L-L-L-R-R-L-L-L-R-C-C-C-L-L-L"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-2-0-0-0-2-0-0-0-0-0-0"
      TextArray0      =   "N°"
      SelectionMode   =   1
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmInversionesExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnColAdd As Integer
Dim lsOpeCodExt As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdAceptar_Click()
    If Not Valida Then Exit Sub
    Dim oCaja As New nCajaGeneral
    Dim nMovNroExt As Long
    Dim sMovNroExt As String
    Dim nCapital As Currency
    Dim nCalculado As Currency
    Dim nInteres As Currency
    'Dim dFechaMov As Date
    Dim sObjetoCod As String
    Dim sMovNro As String
    
    If MsgBox("Esta Seguro de Continuar con el EXTORNO", vbYesNo, "!Aviso¡") = vbYes Then
        With Me.fgIF
            nMovNroExt = .TextMatrix(.row, 1)
            sMovNroExt = .TextMatrix(.row, 18)
            'nCapital = CCur(.TextMatrix(.Row, 7))
            nCalculado = CCur(.TextMatrix(.row, 3))
            nInteres = CCur(.TextMatrix(.row, 12))
            sObjetoCod = .TextMatrix(.row, 15) + "." + .TextMatrix(.row, 14) + "." + .TextMatrix(.row, 16)
            
            'Verificar Mes Cerrado
            Dim oFun As New NContFunciones
            Dim oCon As New NContFunciones
            Dim lbEliminaMov As Boolean
            
            lbEliminaMov = oFun.PermiteModificarAsiento(sMovNroExt, False)
            If Not lbEliminaMov Then
               If MsgBox("Fecha de Extorno corresponde a un mes ya Cerrado, ¿ Desea Extornar Operación ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then Exit Sub
            End If
            Set oFun = Nothing
            
            sMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            
            oCaja.GrabaExtornoInversiones sMovNro, nMovNroExt, sMovNroExt, gsOpeCod, Me.txtMovDesc.Text, nCalculado, True, sObjetoCod, nInteres, lsOpeCodExt, Trim(Right(Me.fgIF.TextMatrix(fgIF.row, 4), 2))
                
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
                Set objPista = Nothing
                '****
                                           
                If MsgBox("Desea Realizar otro Extorno ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                        .EliminaFila .row
                        Me.txtMovDesc.Text = ""
                                                
                Else
                    Unload Me
                End If
            
        End With
    End If
End Sub
Private Function Valida() As Boolean
    Valida = True
    If Me.fgIF.TextMatrix(Me.fgIF.row, 1) = "" Then
        MsgBox "El Dato seleccionado de la Lista no es Valida", vbInformation, "Aviso!"
        Valida = False
        Exit Function
    End If

    If Me.txtMovDesc.Text = "" Then
        MsgBox "Ingrese un Descripcion o Glosa de la Operacion", vbInformation, "Aviso!"
        Valida = False
        txtMovDesc.SetFocus
        Exit Function
    End If
    
End Function
Private Sub cmdProcesar_Click()
    Dim oCaja As nCajaGeneral
    Dim rs As Recordset
    Dim sOpeCod As String
    Dim sOpeCodVig As String
    
     If ValidaFecha(Me.txtDesde.Text) <> "" Then
        MsgBox "Fecha de Inicio no Valida", vbInformation, "AVISO"
        Me.txtDesde.SetFocus
        Exit Sub
     ElseIf ValidaFecha(Me.txthasta.Text) <> "" Then
        MsgBox "Fecha Fin no Valida", vbInformation, "AVISO"
        Me.txthasta.SetFocus
        Exit Sub
     End If
     
     obtenerOperaciones gsOpeCod, sOpeCod, sOpeCodVig
     Set oCaja = New nCajaGeneral
     Set rs = New Recordset
     Set rs = oCaja.obtenerInversionesListaExtornar(sOpeCod, sOpeCodVig, Me.txtDesde.Text, Me.txthasta.Text, Trim(Right(Me.cmbTipo.Text, 2)))
     
    fgIF.Clear
    fgIF.FormaCabecera
    fgIF.Rows = 2
    
    Me.txtMovDesc.Text = ""
            
        If Not rs.EOF And Not rs.BOF Then
            Set fgIF.Recordset = rs
            fgIF.SetFocus
        Else
            MsgBox "Datos no encontrados para proceso seleccionado", vbInformation, "Aviso"
        End If
    
    
    RSClose rs
      
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgIF_Click()
    With Me.fgIF
            Me.txtMovDesc.Text = .TextMatrix(.row, 17)
    End With
End Sub
Private Sub Form_Load()
    txtDesde = DateAdd("d", (Day(gdFecSis) - 1) * -1, gdFecSis)
    txthasta = gdFecSis
    lnColAdd = 0
    cargarCabecera
    mostrarTitulo
    cargarTipoInversion
End Sub
Private Sub cargarTipoInversion()
    Dim rsTpoInversion As ADODB.Recordset
    Dim oCons As DConstante
    
    Set rsTpoInversion = New ADODB.Recordset
    Set oCons = New DConstante
    
    Set rsTpoInversion = oCons.CargaConstante(9990)
    If Not (rsTpoInversion.EOF And rsTpoInversion.BOF) Then
        cmbTipo.Clear
        Do While Not rsTpoInversion.EOF
            cmbTipo.AddItem Trim(rsTpoInversion(2)) & Space(100) & Trim(rsTpoInversion(1))
            rsTpoInversion.MoveNext
        Loop
        cmbTipo.AddItem "Todos los Tipos" + Space(70) + "%", 0
        cmbTipo.ListIndex = 0
    End If
    
End Sub
Private Sub obtenerOperaciones(ByVal pgsOpeCod As String, ByRef psOpeCod As String, ByRef psOpeCodVig As String)

    Select Case pgsOpeCod
    Case "421306", "422306" 'Codigo Apertura
        psOpeCod = "42" + Mid(pgsOpeCod, 3, 1) + "301"
        psOpeCodVig = ""
       
    Case "421307", "422307" 'Codigo Confirmacion
        psOpeCod = "42" + Mid(pgsOpeCod, 3, 1) + "302"
        psOpeCodVig = ""
    Case "421308", "422308" 'Codigo Cancelacion
        psOpeCod = "42" + Mid(pgsOpeCod, 3, 1) + "303"
        psOpeCodVig = ""
    Case "421309", "422309" 'Codigo Provision
        psOpeCod = "42" + Mid(pgsOpeCod, 3, 1) + "304"
        psOpeCodVig = "42" + Mid(pgsOpeCod, 3, 1) + "302"
    End Select
    lsOpeCodExt = psOpeCod
End Sub
Private Sub mostrarTitulo()
     Select Case gsOpeCod
    Case "421306", "422306"
        Me.Caption = "EXTORNO APERTURA INVERSIONES " + IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME")
        Me.lbltitulo = "EXTORNO APERTURA"
    Case "421307", "422307"
        Me.Caption = "EXTORNO CONFIRMACION INVERSIONES " + IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME")
        Me.lbltitulo = "EXTORNO CONFIRMACION"
     Case "421308", "422308"
        Me.Caption = "EXTORNO CANCELACION INVERSIONES " + IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME")
        Me.lbltitulo = "EXTORNO CANCELACION"
    Case "421309", "422309"
        Me.Caption = "EXTORNO PROVISION INVERSIONES " + IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME")
        Me.lbltitulo = "EXTORNO PROVISION"
    End Select
End Sub
Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txthasta.SetFocus
    End If
End Sub
Private Sub txthasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub
Private Sub cargarCabecera()
    
    Select Case gsOpeCod
        Case "421306", "422306", "421307", "422307"
            Me.fgIF.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
            Me.fgIF.EncabezadosAlineacion = "C-L-L-R-L-L-L-R-R-L-L-L-R-C-C-C-L-L-L"
            Me.fgIF.EncabezadosAnchos = "350-0-900-0-800-1800-2500-1200-800-900-600-900-0-0-0-0-0-0-0"
            Me.fgIF.EncabezadosNombres = "N°-nMovNro-Fecha-Calculado-Tipo-Nro_Cuenta-Inst_Financ-Capital-Tasa-Fec_Ape-Plazo-Fec_Venc-Interes-Fec_Int-cPersCodAdm-cIFTpoAdm-cCtaIFCodAdm-cGlosa-cMovNro"
            Me.fgIF.FormatosEdit = "0-0-0-0-0-0-0-2-2-0-0-0-2-0-0-0-0-0-0"
            Me.fgIF.ListaControles = "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
            
        Case "421308", "422308", "421309", "422309"
            Me.fgIF.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
            Me.fgIF.EncabezadosAlineacion = "C-L-L-R-L-L-L-R-R-L-L-L-R-C-C-C-L-L-L"
            Me.fgIF.EncabezadosAnchos = "350-0-900-900-800-1800-2500-1200-800-900-600-900-0-900-0-0-0-0-0"
            Me.fgIF.EncabezadosNombres = "N°-nMovNro-Fecha-Calculado-Tipo-Nro_Cuenta-Inst_Financ-Capital-Tasa-Fec_Ape-Plazo-Fec_Venc-Interes-Fec_Int-cPersCodAdm-cIFTpoAdm-cCtaIFCodAdm-cGlosa-cMovNro"
            Me.fgIF.FormatosEdit = "0-0-0-0-0-0-0-2-2-0-0-0-2-0-0-0-0-0-0"
            Me.fgIF.ListaControles = "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
            lnColAdd = 1
    End Select
   
End Sub
