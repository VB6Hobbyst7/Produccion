VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInversionesMantenimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   2730
   ClientTop       =   3675
   ClientWidth     =   10605
   Icon            =   "frmInversionesMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   10605
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   10215
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   7200
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   8640
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10290
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
         Left            =   8805
         TabIndex        =   2
         Top             =   203
         Width           =   1335
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   345
         Left            =   2895
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
         Left            =   4605
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
         Caption         =   "Mantenimiento"
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
         Left            =   210
         TabIndex        =   8
         Top             =   255
         Width           =   1575
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
         Left            =   4005
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
         Left            =   2235
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
         Left            =   5880
         TabIndex        =   5
         Top             =   270
         Width           =   495
      End
   End
   Begin Sicmact.FlexEdit fgInversiones 
      Height          =   3000
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   5292
      Cols0           =   16
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "N°-Fecha-Estado-Tipo-Institucion-Cuenta-Plazo-Apertura-Vencimiento-Tasa-Importe-cCtaIfCod-cIfTpo-cPersCod-nMovNro-cOpeCod"
      EncabezadosAnchos=   "350-900-1000-800-3000-1350-600-900-1000-600-1200-0-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-L-C-C-C-C-R-R-L-L-L-L-L"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-2-2-0-0-0-0-0"
      TextArray0      =   "N°"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmInversionesMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdModificar_Click()
    With Me.fgInversiones
        If .TextMatrix(.row, 1) <> "" Then
            If frmInversiones.Inicio(.TextMatrix(.row, 14), .TextMatrix(.row, 15)) Then
                cmdProcesar_Click
            End If
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
                Set objPista = Nothing
                '****
        End If
    End With
End Sub

Private Sub cmdProcesar_Click()
    Dim oCaja As nCajaGeneral
     Dim rs As Recordset
     
     If ValidaFecha(Me.txtDesde.Text) <> "" Then
        MsgBox "Fecha de Inicio no Valida", vbInformation, "AVISO"
        Me.txtDesde.SetFocus
        Exit Sub
     ElseIf ValidaFecha(Me.txthasta.Text) <> "" Then
        MsgBox "Fecha Fin no Valida", vbInformation, "AVISO"
        Me.txthasta.SetFocus
        Exit Sub
     End If
     
     Set oCaja = New nCajaGeneral
     Set rs = New Recordset
     Set rs = oCaja.getInversionesListadoMantenimiento(Mid(gsOpeCod, 3, 1), Me.txtDesde.Text, Me.txthasta.Text, Trim(Right(Me.cmbTipo.Text, 2)))
     
    fgInversiones.Clear
    fgInversiones.FormaCabecera
    fgInversiones.Rows = 2
                
        If Not rs.EOF And Not rs.BOF Then
            Set fgInversiones.Recordset = rs
            fgInversiones.SetFocus
        Else
            MsgBox "Datos no encontrados para proceso seleccionado", vbInformation, "Aviso"
        End If
    
    
    RSClose rs
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgInversiones_DblClick()
 With Me.fgInversiones
        If .TextMatrix(.row, 1) <> "" Then
            If frmInversiones.Inicio(.TextMatrix(.row, 14), .TextMatrix(.row, 15)) Then
                cmdProcesar_Click
            End If
        
        End If
    End With
End Sub
Private Sub Form_Load()
    If gsOpeCod = "421310" Then
        Me.Caption = "Lista de Mantenimiento de Inversiones MN"
    ElseIf gsOpeCod = "422310" Then
        Me.Caption = "Lista de Mantenimiento de Inversiones ME"
    End If
    txtDesde = DateAdd("d", (Day(gdFecSis) - 1) * -1, gdFecSis)
    txthasta = gdFecSis
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
