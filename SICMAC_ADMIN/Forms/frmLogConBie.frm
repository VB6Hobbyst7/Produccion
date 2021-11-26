VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogConBie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contratación"
   ClientHeight    =   5820
   ClientLeft      =   690
   ClientTop       =   2055
   ClientWidth     =   8985
   Icon            =   "frmLogConBie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtConNro 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1815
      TabIndex        =   11
      Top             =   120
      Width           =   2505
   End
   Begin VB.CommandButton cmdCon 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   5715
      TabIndex        =   9
      Top             =   5280
      Width           =   1260
   End
   Begin VB.CommandButton cmdCon 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   3870
      TabIndex        =   8
      Top             =   5280
      Width           =   1260
   End
   Begin VB.CommandButton cmdCon 
      Caption         =   "&Nuevo"
      Height          =   390
      Index           =   0
      Left            =   2085
      TabIndex        =   7
      Top             =   5280
      Width           =   1260
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   6810
      TabIndex        =   1
      Top             =   90
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   23461889
      CurrentDate     =   37099
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7470
      TabIndex        =   0
      Top             =   5280
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   0
      Top             =   5370
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin TabDlg.SSTab sstContrata 
      Height          =   4260
      Left            =   150
      TabIndex        =   3
      Top             =   900
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   7514
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Detalle"
      TabPicture(0)   =   "frmLogConBie.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEtiqueta(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEtiqueta(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProNom"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProDir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLugar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblEtiqueta(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraMoneda"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtLugar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtLugEnt"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Bienes/Servicios"
      TabPicture(1)   =   "frmLogConBie.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgeEvaEco"
      Tab(1).Control(1)=   "fgeEvaEcoTot"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtLugEnt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   70
         TabIndex        =   28
         Top             =   1455
         Width           =   5700
      End
      Begin VB.Frame Frame1 
         Caption         =   "Parámetros"
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
         Height          =   2190
         Left            =   195
         TabIndex        =   23
         Top             =   1830
         Width           =   8175
         Begin Sicmact.FlexEdit fgeEvaParTec 
            Height          =   1590
            Left            =   180
            TabIndex        =   24
            Top             =   450
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   2805
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Item-Codigo-Descripción-Valor-Tipo"
            EncabezadosAnchos=   "400-0-2200-550-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-L-R-L"
            FormatosEdit    =   "0-0-0-3-0"
            CantDecimales   =   0
            AvanceCeldas    =   1
            TextArray0      =   "Item"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
         End
         Begin Sicmact.FlexEdit fgeEvaParEco 
            Height          =   1590
            Left            =   4410
            TabIndex        =   25
            Top             =   450
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   2805
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Item-Codigo-Descripción-Valor-Tipo"
            EncabezadosAnchos=   "400-0-2200-550-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-L-R-L"
            FormatosEdit    =   "0-0-0-3-0"
            CantDecimales   =   0
            AvanceCeldas    =   1
            TextArray0      =   "Item"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Técnicos"
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
            Height          =   210
            Index           =   10
            Left            =   315
            TabIndex        =   27
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Económicos"
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
            Height          =   210
            Index           =   11
            Left            =   4575
            TabIndex        =   26
            Top             =   240
            Width           =   1140
         End
      End
      Begin Sicmact.TxtBuscar txtLugar 
         Height          =   315
         Left            =   225
         TabIndex        =   20
         Top             =   1140
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
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
         Enabled         =   0   'False
         Enabled         =   0   'False
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.Frame fraMoneda 
         Caption         =   "Moneda "
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
         ForeColor       =   &H8000000D&
         Height          =   900
         Left            =   6990
         TabIndex        =   4
         Top             =   510
         Width           =   1185
         Begin VB.OptionButton optMoneda 
            Caption         =   "Soles"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   6
            Top             =   240
            Width           =   750
         End
         Begin VB.OptionButton optMoneda 
            Caption         =   "Dólares"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   5
            Top             =   495
            Width           =   900
         End
      End
      Begin Sicmact.FlexEdit fgeEvaEco 
         Height          =   2715
         Left            =   -74685
         TabIndex        =   12
         Top             =   600
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   4789
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-Precio-SubTotal"
         EncabezadosAnchos=   "400-0-3500-900-900-900-900"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2"
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeEvaEcoTot 
         Height          =   915
         Left            =   -74685
         TabIndex        =   13
         Top             =   3000
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   1614
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "------Total"
         EncabezadosAnchos=   "400-0-3500-900-900-900-900"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColor       =   -2147483624
         EncabezadosAlineacion=   "C-L-L-L-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2"
         AvanceCeldas    =   1
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         CellBackColor   =   -2147483624
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Dirección"
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
         Height          =   210
         Index           =   5
         Left            =   3585
         TabIndex        =   22
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label lblLugar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Top             =   1140
         Width           =   3195
      End
      Begin VB.Label lblProDir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3525
         TabIndex        =   19
         Top             =   615
         Width           =   3195
      End
      Begin VB.Label lblProNom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   195
         TabIndex        =   18
         Top             =   615
         Width           =   3285
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Lugar de entrega"
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
         Height          =   210
         Index           =   4
         Left            =   270
         TabIndex        =   17
         Top             =   930
         Width           =   1560
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Proveedor"
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
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   14
         Top             =   405
         Width           =   1095
      End
   End
   Begin Sicmact.TxtBuscar txtSelNro 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
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
      Enabled         =   0   'False
      Enabled         =   0   'False
      TipoBusqueda    =   2
      sTitulo         =   ""
      EnabledText     =   0   'False
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Nro Selección :"
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
      Height          =   210
      Index           =   2
      Left            =   450
      TabIndex        =   16
      Top             =   540
      Width           =   1425
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Contratación :"
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
      Height          =   210
      Index           =   0
      Left            =   450
      TabIndex        =   10
      Top             =   150
      Width           =   1335
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Fecha :"
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
      Height          =   210
      Index           =   3
      Left            =   5940
      TabIndex        =   2
      Top             =   150
      Width           =   750
   End
End
Attribute VB_Name = "frmLogConBie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCon_Click(Index As Integer)
    Dim clsDGnral As DLogGeneral
    Dim clsDMov As DLogMov
    Dim sConNro As String
    Dim nConNro As Long, nSelNro As Long
    Dim sActualiza As String
    Dim sBSCod As String
    Dim nCantid As Currency, nPrecio As Currency
    Dim nCont As Integer, nResult As Integer
    
    If Index = 0 Then
        'NUEVO
        Call Limpiar
        txtConNro.Text = ""
        txtSelNro.Text = ""
        txtLugar.Enabled = True
        txtLugEnt.Enabled = True
        
        Set clsDGnral = New DLogGeneral
        txtConNro.Text = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        Set clsDGnral = Nothing
        
        Call CargaProcesos
    ElseIf Index = 1 Then
        'CANCELAR
        If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Call Limpiar
            txtConNro.Text = ""
            txtSelNro.Text = ""
            txtSelNro.Enabled = False
            txtLugEnt.Enabled = False
            txtLugar.Enabled = False
            cmdCon(0).Enabled = True
            cmdCon(1).Enabled = False
            cmdCon(2).Enabled = False
        End If
    ElseIf Index = 2 Then
        'GRABAR
        txtLugEnt.Text = Trim(txtLugEnt.Text)
        If txtSelNro.Text = "" Then
            MsgBox "Tiene que determinar un proceso de Selección", vbInformation, " Aviso"
            Exit Sub
        End If
        If txtLugar.Text = "" Then
            MsgBox "Tiene que determinar la agencia", vbInformation, " Aviso"
            Exit Sub
        End If
        If txtLugEnt.Text = "" Then
            MsgBox "Falta determinar en que dirección se hará la entrega", vbInformation, " Aviso"
            Exit Sub
        End If
        
        If MsgBox("¿ Estás seguro de realizar la Contratación " & txtConNro.Text & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Set clsDGnral = New DLogGeneral
            Set clsDMov = New DLogMov
            
            sConNro = txtConNro.Text
            nSelNro = clsDGnral.GetnMovNro(txtSelNro.Text)
            
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)

            'Grabación de MOV -MOVREF
            clsDMov.InsertaMov sConNro, Trim(Str(gLogOpeConRegistro)), "", gLogConEstadoInicio
            nConNro = clsDMov.GetnMovNro(sConNro)
            'clsDMov.InsertaMovRef nConNro, nConNro
            
            'Actualiza LogSelección
            clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoContratacion, sActualiza
            
            'Inserta LogContratación
            clsDMov.InsertaContratacion nConNro, nSelNro, dtpFecha.value, _
                txtLugar.Text, txtLugEnt.Text, sActualiza
            
            For nCont = 1 To fgeEvaEco.Rows - 1
                sBSCod = Trim(fgeEvaEco.TextMatrix(fgeEvaEco.Row, 1))
                nCantid = CCur(fgeEvaEco.TextMatrix(fgeEvaEco.Row, 4))
                nPrecio = CCur(fgeEvaEco.TextMatrix(fgeEvaEco.Row, 5))
                
                'Inserta LogConDetalle
                clsDMov.InsertaConDetalle nConNro, sBSCod, nCantid, nPrecio, sActualiza
            Next
            
            'Ejecuta todos los querys en una transacción
            'nResult = clsDMov.EjecutaBatch
            Set clsDMov = Nothing
            Set clsDGnral = Nothing
            
            If nResult = 0 Then
                txtConNro.Enabled = False
                txtSelNro.Enabled = False
                txtLugar.Enabled = False
                txtLugEnt.Enabled = False
                cmdCon(0).Enabled = True
                cmdCon(1).Enabled = False
                cmdCon(2).Enabled = False
            Else
                MsgBox "Error al grabar la información", vbInformation, " Aviso "
            End If
        End If
    Else
        MsgBox "Opción no reconocida", vbInformation, " Aviso"
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim clsDAdq As DLogAdquisi
    
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        txtSelNro.Enabled = False
        'cmdReq(0).Enabled = False
        'sstReq.Enabled = False
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    
    
End Sub

Private Sub txtLugar_EmiteDatos()
    Dim clsDGnral As DLogGeneral
    Dim rs As ADODB.Recordset
    
    If Not txtLugar.Ok Then
        Exit Sub
    End If
    
    Set clsDGnral = New DLogGeneral
    Set rs = New ADODB.Recordset
    Set rs = clsDGnral.CargaAgencia(AgeUnRegistro, txtLugar.Text)
    If rs.RecordCount > 0 Then
        lblLugar.Caption = rs!cAgeDescripcion
        txtLugEnt.Text = rs!cAgeDireccion
    End If
    
End Sub

Private Sub txtSelNro_EmiteDatos()
    Dim rsCot As ADODB.Recordset, rs As ADODB.Recordset
    Dim clsDGnral As DLogGeneral
    Dim clsDAdq As DLogAdquisi
    Dim nSelNro As Long, nSelCotNro As Long
        
    Set rs = New ADODB.Recordset
    Set rsCot = New ADODB.Recordset
    Set clsDGnral = New DLogGeneral
    Set clsDAdq = New DLogAdquisi
    
    If Not txtSelNro.Ok Then
        Exit Sub
    End If
    Call Limpiar
    
    nSelNro = clsDGnral.GetnMovNro(txtSelNro.Text)
    Set rsCot = clsDAdq.CargaSeleccion(SelUnRegistro, nSelNro)
    If rsCot.RecordCount > 0 Then
        nSelCotNro = rsCot!nLogSelCotNro
        
        If rsCot!nLogSelMoneda = gMonedaNacional Then
            optMoneda(0).value = True
        ElseIf rsCot!nLogSelMoneda = gMonedaExtranjera Then
            optMoneda(1).value = True
        End If
        
        Set rs = clsDAdq.CargaSelCotDetalle(SelCotDetUnRegistro, nSelNro, nSelCotNro)
        If rs.RecordCount > 0 Then
            Set fgeEvaEco.Recordset = rs
            fgeEvaEcoTot.BackColorRow &HC0FFFF
            fgeEvaEcoTot.TextMatrix(1, 0) = "="
            fgeEvaEcoTot.TextMatrix(1, 2) = "T O T A L "
            fgeEvaEcoTot.TextMatrix(1, 6) = Format(fgeEvaEco.SumaRow(6), "#,##0.00")
        End If
        
        Set rs = clsDAdq.CargaSelCotPar(SelCotParRegistro, nSelNro, 1, nSelCotNro)
        If rs.RecordCount > 0 Then Set fgeEvaParTec.Recordset = rs
        
        Set rs = clsDAdq.CargaSelCotPar(SelCotParRegistro, nSelNro, 2, nSelCotNro)
        If rs.RecordCount > 0 Then Set fgeEvaParEco.Recordset = rs
    
        Set rs = clsDAdq.CargaSelCotiza(SelCotPersona, nSelNro, nSelCotNro)
        If rs.RecordCount > 0 Then
            lblProNom.Caption = rs!cPersNombre
            lblProDir.Caption = rs!cPersDireccDomicilio
        End If
        
        Set rs = clsDGnral.CargaAgencia(AgeTotal)
        If rs.RecordCount > 0 Then txtLugar.rs = rs
    End If
    
    Set clsDGnral = Nothing
    Set clsDAdq = Nothing
End Sub

Private Sub Limpiar()
    dtpFecha.value = gdFecSis
    optMoneda(0).value = False
    optMoneda(1).value = False
    
    lblProNom.Caption = ""
    lblProDir.Caption = ""
    txtLugar.Text = ""
    lblLugar.Caption = ""
    txtLugEnt.Text = ""
    
    fgeEvaEco.Clear
    fgeEvaEco.FormaCabecera
    fgeEvaEco.Rows = 2
    fgeEvaEcoTot.Clear
    fgeEvaEcoTot.FormaCabecera
    fgeEvaEcoTot.Rows = 2
    fgeEvaParTec.Clear
    fgeEvaParTec.FormaCabecera
    fgeEvaParTec.Rows = 2
    fgeEvaParEco.Clear
    fgeEvaParEco.FormaCabecera
    fgeEvaParEco.Rows = 2
End Sub

Private Sub CargaProcesos()
    Dim rs As ADODB.Recordset
    Dim clsDAdq As DLogAdquisi
    
    Set rs = New ADODB.Recordset
    Set clsDAdq = New DLogAdquisi
    Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, , gLogSelEstadoConsentimiento)
    If rs.RecordCount > 0 Then
        txtSelNro.EditFlex = True
        txtSelNro.rs = rs
        txtSelNro.EditFlex = False
        txtSelNro.Enabled = True
        
        cmdCon(0).Enabled = False
        cmdCon(1).Enabled = True
        cmdCon(2).Enabled = True
    Else
        txtSelNro.Enabled = False
        txtSelNro.Text = ""
    End If
    
    Set clsDAdq = Nothing
    Set rs = Nothing
End Sub
